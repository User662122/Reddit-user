#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AutoKeyPress - Excel-Driven Keyboard Automation
================================================
Reads rows from an Excel file and performs key sequences
when the trigger key (e.g. Esc) is pressed.

REQUIRED COLUMNS (any position in Excel, any letter case):
  Trigger   -> key to wait for before running that row (e.g. Esc, F5, Enter)
              Special values:
                cont  -> no trigger needed, row executes instantly (continuous)
                PB    -> Pause/Break key
                IT    -> Pause/Break key (alias for PB)
                2211  -> numeric key sequence (type those digits to trigger)
  KeyMap    -> the keystroke sequence to perform for that row

ALL OTHER COLUMNS are data columns. Add or remove freely - their names
are used directly in KeyMap via v"COLNAME". Missing columns are skipped
(with a one-time warning instead of silence, so typos get noticed).

TRIGGER SPECIAL VALUES:
  cont      -> Row runs instantly without waiting for any key press.
  PB or IT  -> Wait for the Pause/Break key on the keyboard.
  2211      -> Wait for that digit sequence to be typed.
  Esc, F5   -> Standard named keys.

KEYMAP SYNTAX:
  v"COLNAME"       -> type the value from that column (skip if column absent)
                      Special: a column named "Date" is formatted DDMMYY (6 digits)
                        Excel value  01-01-2025  ->  types  010125
                        Excel value   5-3-2025   ->  types  050325
                        Excel value  2025-01-01  ->  types  010125   (ISO works too)
  t"any text"      -> type the literal text inside the quotes (not a column)
                        t"Hello World" -> types Hello World
                      Both " and ' quotes work, and a '+' inside quotes is NOT
                      treated as a separator: t"a+b" types a+b
  Enter            -> press Enter key
  Tab              -> press Tab key
  Esc              -> press Escape key
  Space            -> press Space key
  Up/Down/Left/Right -> arrow keys
  F1..F12          -> function keys
  Ctrl+c           -> hold Ctrl and press c  (any combo, e.g. Ctrl+Shift+s)
  Alt+F4           -> hold Alt and press F4
  Shift+Tab        -> hold Shift and press Tab
  2*Enter          -> repeat Enter 2 times  (N*KEY)
                      Works with combos and columns too: 3*Ctrl+c, 2*v"AMOUNT"
  +                -> separator between actions

EXAMPLE KeyMap:
  Enter + v"Date" + 2*Enter + v"NARRATION" + Tab + v"AMOUNT" + Enter
  Enter + t"PREFIX" + v"AMOUNT" + Tab + v"Date" + Enter

EMERGENCY STOP - MOVE THE MOUSE
  From the moment a row starts typing, a background watcher tracks the mouse.
  Move it by more than --mouse-threshold pixels (default 5) and everything
  stops immediately:
    * no further keys are sent for this row or any later row
    * any Ctrl/Alt/Shift/Win still held down is released, so the keyboard is
      never left with a stuck modifier
    * the script reports how many rows completed and exits
  The watcher is armed only while keys are being sent - during the trigger
  wait and the countdown you can freely move the mouse / click into the
  target window.
  Hard abort backup: slam the mouse into any SCREEN CORNER (pyautogui
  failsafe - current pyautogui treats all four corners as abort points).
  Disable the watcher with --no-mouse-stop.

COMMAND LINE:
  python E.py "C:\\path\\to\\file.xlsx" [--sheet Sheet2] [--delay 3]
              [--mouse-threshold 5] [--no-mouse-stop] [--ask-resume]
              [--typing-interval 0.03] [--paste-columns Narration]
              [--date-columns Date] [--dry-run]
"""

from __future__ import annotations

import argparse
import datetime
import os
import re
import sys
import threading
import time

VERSION = "2.0.0"


# ── Optional third-party imports (friendly errors instead of tracebacks) ─────
def _try_import(name):
    try:
        return __import__(name), None
    except Exception as exc:  # pragma: no cover - depends on environment
        return None, exc


openpyxl, OPENPYXL_ERROR = _try_import("openpyxl")
pyautogui, PYAUTOGUI_ERROR = _try_import("pyautogui")
keyboard, KEYBOARD_ERROR = _try_import("keyboard")
pyperclip, PYPERCLIP_ERROR = _try_import("pyperclip")

HAS_KEYBOARD = keyboard is not None
HAS_PYPERCLIP = pyperclip is not None
HAS_PYAUTOGUI = pyautogui is not None

if HAS_PYAUTOGUI:
    pyautogui.FAILSAFE = True  # mouse into a screen corner = hard abort
    pyautogui.PAUSE = 0.05     # 50ms between actions for reliability


# ── Errors ───────────────────────────────────────────────────────────────────
class ScriptError(Exception):
    """Something in the Excel file / CLI arguments cannot be used."""


class AutomationStopped(Exception):
    """Raised internally to unwind the run as soon as a stop is requested."""


# ── Runtime configuration (overwritten from the command line) ────────────────
DRY_RUN = False
TYPING_INTERVAL = 0.03
ACTION_GAP = 0.05
PASTE_SETTLE = 0.15
MAX_REPEAT = 500
TRIGGER_SETTLE = 0.25
PASTE_COLUMNS = {"narration"}   # columns pasted from the clipboard
DATE_COLUMNS = {"date"}         # columns reformatted to DDMMYY

# Record of everything sent, in --dry-run mode (also handy for tests).
DRY_RUN_LOG: list = []

# Modifier keys currently held down, so a stop can always release them.
_HELD_MODIFIERS: list = []

# Messages already reported once (keeps the console readable).
_WARNED: set = set()


def warn_once(message):
    if message not in _WARNED:
        _WARNED.add(message)
        print(f"    ! {message}", flush=True)


# ── Emergency stop state ─────────────────────────────────────────────────────
STOP_EVENT = threading.Event()
STOP_REASON = ""


def request_stop(reason):
    """Ask the automation to stop (called by the mouse watcher / signal)."""
    global STOP_REASON
    STOP_REASON = reason
    STOP_EVENT.set()


def reset_stop():
    global STOP_REASON
    STOP_REASON = ""
    STOP_EVENT.clear()


def check_stop():
    """Raise AutomationStopped if a stop was requested. Call often."""
    if STOP_EVENT.is_set():
        release_modifiers()
        raise AutomationStopped(STOP_REASON or "stop requested")


def release_modifiers():
    """Release any Ctrl/Alt/Shift/Win the script is still holding down."""
    if not _HELD_MODIFIERS or not HAS_PYAUTOGUI:
        _HELD_MODIFIERS.clear()
        return
    while _HELD_MODIFIERS:
        key = _HELD_MODIFIERS.pop()
        try:
            pyautogui.keyUp(key)
        except Exception:
            pass


def mouse_position():
    """Current (x, y) of the mouse, or None when it cannot be read."""
    if not HAS_PYAUTOGUI:
        return None
    try:
        pos = pyautogui.position()
        return (pos[0], pos[1])
    except Exception:
        return None


class MouseStopWatcher:
    """
    Watches the mouse while a row is typing.

    Any movement bigger than `threshold` pixels away from the position the
    mouse had when the watcher was armed raises the global stop flag, which
    the executor notices at the next key/character and unwinds cleanly.

    The watcher keeps one background thread for the whole run and is armed /
    disarmed per row, so moving the mouse between rows is harmless.
    """

    def __init__(self, threshold=5.0, poll=0.02):
        self.threshold = float(threshold)
        self.poll = max(0.005, float(poll))
        self._armed = threading.Event()     # set while a row is typing
        self._idle = threading.Event()      # set when the loop is not polling
        self._shutdown = threading.Event()
        self._thread = None
        self._idle.set()
        self.origin = None
        self.armed = False

    # -- internals ----------------------------------------------------------
    def _loop(self):
        while not self._shutdown.is_set():
            if not self._armed.wait(self.poll):
                self._idle.set()
                continue
            self._idle.clear()
            origin = mouse_position()
            if origin is None:               # cannot read the mouse: stand down
                self.armed = False
                self._armed.clear()
                self._idle.set()
                continue
            self.origin = origin
            self.armed = True
            while self._armed.is_set() and not self._shutdown.is_set():
                pos = mouse_position()
                if pos is None:
                    time.sleep(self.poll)
                    continue
                dx = pos[0] - origin[0]
                dy = pos[1] - origin[1]
                distance = (dx * dx + dy * dy) ** 0.5
                if distance > self.threshold:
                    # Re-check the arm flag so a movement that lands in the
                    # same instant the row finished cannot stop the next row.
                    if self._armed.is_set():
                        request_stop(
                            f"mouse moved {distance:.0f}px "
                            f"(threshold {self.threshold:g}px)"
                        )
                    break
                time.sleep(self.poll)
            self.armed = False
            self._armed.clear()
            self._idle.set()

    def _ensure_thread(self):
        if self._thread is None or not self._thread.is_alive():
            self._shutdown.clear()
            self._thread = threading.Thread(
                target=self._loop, daemon=True, name="mouse-stop-watcher"
            )
            self._thread.start()

    # -- public API ---------------------------------------------------------
    def watch(self):
        """Arm the watcher (records the current mouse position as the origin)."""
        if self.threshold <= 0:
            return
        self._ensure_thread()
        self._armed.set()

    def cancel(self):
        """Disarm the watcher and wait until it really stopped polling."""
        self._armed.clear()
        self.armed = False
        if self._thread is not None and self._thread.is_alive():
            self._idle.wait(0.25)

    def shutdown(self):
        """Stop the background thread for good."""
        self.cancel()
        self._shutdown.set()
        self._armed.clear()
        self.armed = False

    def is_alive(self):
        return self._thread is not None and self._thread.is_alive()


# ── Key name map ─────────────────────────────────────────────────────────────
KEY_MAP = {
    "enter":     "enter",
    "return":    "enter",
    "tab":       "tab",
    "esc":       "escape",
    "escape":    "escape",
    "space":     "space",
    "backspace": "backspace",
    "delete":    "delete",
    "del":       "delete",
    "up":        "up",
    "down":      "down",
    "left":      "left",
    "right":     "right",
    "home":      "home",
    "end":       "end",
    "pageup":    "pageup",
    "pgup":      "pageup",
    "pagedown":  "pagedown",
    "pgdn":      "pagedown",
    "insert":    "insert",
    "printscreen": "printscreen",
    "capslock":  "capslock",
    "numlock":   "numlock",
    "f1":  "f1",  "f2":  "f2",  "f3":  "f3",  "f4":  "f4",
    "f5":  "f5",  "f6":  "f6",  "f7":  "f7",  "f8":  "f8",
    "f9":  "f9",  "f10": "f10", "f11": "f11", "f12": "f12",
    "f13": "f13", "f14": "f14", "f15": "f15", "f16": "f16",
    "f17": "f17", "f18": "f18", "f19": "f19", "f20": "f20",
    "f21": "f21", "f22": "f22", "f23": "f23", "f24": "f24",
    "ctrl":  "ctrl",  "alt": "alt",  "shift": "shift",
    "win":   "win",   "cmd": "command",
    # Pause/Break key support - write PB or IT in the Trigger column
    "pb":           "pause",
    "it":           "pause",
    "pause":        "pause",
    "pausebreak":   "pause",
    "pause/break":  "pause",
    "break":        "pause",
}

MODIFIER_MAP = {
    "ctrl":  "ctrl",
    "control": "ctrl",
    "alt":   "alt",
    "shift": "shift",
    "win":   "win",
    "windows": "win",
    "cmd":   "command",
    "command": "command",
    "option": "option",
}

# Extra names accepted when pyautogui is not importable (e.g. --dry-run here).
_FALLBACK_KEY_NAMES = {
    "accept", "add", "alt", "altleft", "altright", "apps", "backspace",
    "capslock", "clear", "convert", "ctrl", "ctrlleft", "ctrlright", "decimal",
    "del", "delete", "divide", "down", "end", "enter", "esc", "escape",
    "execute", "final", "fn", "hanguel", "hangul", "hanja", "help", "home",
    "insert", "junja", "kana", "kanji", "launchapp1", "launchapp2",
    "launchmail", "launchmediaselect", "left", "modechange", "multiply",
    "nexttrack", "nonconvert", "numlock", "pagedown", "pageup", "pause",
    "pgdn", "pgup", "playpause", "prevtrack", "print", "printscreen",
    "prntscrn", "prtsc", "prtscr", "return", "right", "scrolllock", "select",
    "separator", "shift", "shiftleft", "shiftright", "sleep", "space", "stop",
    "subtract", "tab", "up", "volumedown", "volumemute", "volumeup", "win",
    "winleft", "winright", "yen", "command", "option", "optionleft",
    "optionright", "browserback", "browserfavorites", "browserforward",
    "browserhome", "browserrefresh", "browsersearch", "browserstop",
} | {f"f{i}" for i in range(1, 25)} | {f"num{i}" for i in range(10)}


_KNOWN_KEY_NAMES = None


def known_key_names():
    """Every key name pyautogui will actually accept on this machine."""
    global _KNOWN_KEY_NAMES
    if _KNOWN_KEY_NAMES is None:
        names = set(_FALLBACK_KEY_NAMES)
        names |= {v for v in KEY_MAP.values()}
        names |= set(getattr(pyautogui, "KEY_NAMES", ()) or ())
        names |= set(getattr(pyautogui, "KEYBOARD_KEYS", ()) or ())
        names |= set("abcdefghijklmnopqrstuvwxyz0123456789")
        _KNOWN_KEY_NAMES = {n.lower() for n in names if isinstance(n, str)}
    return _KNOWN_KEY_NAMES


# pyautogui's "mouse in the corner" abort, when it is available.
FAILSAFE_ERRORS = ()
if HAS_PYAUTOGUI:
    _failsafe = getattr(pyautogui, "FailSafeException", None)
    if _failsafe is not None:
        FAILSAFE_ERRORS = (_failsafe,)


# Names the `keyboard` module may report for a trigger key.
TRIGGER_ALIASES = {
    "escape": {"esc", "escape"},
    "enter": {"enter", "return"},
    "pause": {"pause", "break"},
    "delete": {"delete", "del"},
    "printscreen": {"printscreen", "print screen", "prtsc", "prtscr",
                    "prntscrn", "print"},
    "space": {"space", " "},
    "pageup": {"pageup", "pgup", "page up"},
    "pagedown": {"pagedown", "pgdn", "page down"},
    "backspace": {"backspace", "back space"},
    "capslock": {"capslock", "caps lock"},
    "win": {"win", "windows", "left windows", "right windows", "super",
            "winleft", "winright", "left meta", "right meta"},
    "ctrl": {"ctrl", "left ctrl", "right ctrl", "ctrlleft", "ctrlright"},
    "alt": {"alt", "left alt", "right alt", "altleft", "altright"},
    "shift": {"shift", "left shift", "right shift", "shiftleft", "shiftright"},
}


def accepted_trigger_names(key_name):
    """All spellings of `key_name` that count as a trigger press."""
    key = str(key_name).strip().lower()
    names = {key}
    names |= TRIGGER_ALIASES.get(key, set())
    # every KeyMap spelling that resolves to this pyautogui key
    for source, target in KEY_MAP.items():
        if target == key:
            names.add(source)
    return {n.lower() for n in names}


# ── Excel value -> text ──────────────────────────────────────────────────────
def value_to_text(value):
    """
    Turn a raw openpyxl cell value into the text that should be typed.

    Fixes the classic "1500.0" problem: Excel numbers arrive as floats, so
    `str(1500.0)` would type 1500.0 into an amount field. Integers are typed
    without the decimal part, floats keep their decimals, dates are written
    DD-MM-YYYY and booleans become TRUE/FALSE like Excel shows them.
    """
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, datetime.datetime):
        if value.hour == value.minute == value.second == 0:
            return value.strftime("%d-%m-%Y")
        return value.strftime("%d-%m-%Y %H:%M:%S")
    if isinstance(value, datetime.date):
        return value.strftime("%d-%m-%Y")
    if isinstance(value, datetime.time):
        if value.hour == 0 and value.minute == 0 and value.second == 0:
            return "00:00"
        return value.strftime("%H:%M:%S")
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return f"{value:.10f}".rstrip("0").rstrip(".")
    if isinstance(value, int):
        return str(value)
    try:  # Decimal / Fraction and friends
        import decimal
        if isinstance(value, decimal.Decimal):
            return format(value.normalize(), "f")
    except Exception:
        pass
    return str(value).strip()


MONTH_NAMES = {
    "jan": 1, "january": 1, "feb": 2, "february": 2, "mar": 3, "march": 3,
    "apr": 4, "april": 4, "may": 5, "jun": 6, "june": 6, "jul": 7, "july": 7,
    "aug": 8, "august": 8, "sep": 9, "sept": 9, "september": 9, "oct": 10,
    "october": 10, "nov": 11, "november": 11, "dec": 12, "december": 12,
}


def format_date_ddmmyy(value):
    """
    Convert a date value from Excel into a 6-digit DDMMYY string.

    Handles:
      - Python date/datetime objects (real date cells read by openpyxl)
      - '01-01-2025', '1-1-2025', '01/01/2025', '1.1.25'
      - ISO order '2025-01-01' (year first)
      - '01-Jan-2025', 'Jan 1 2025'
      - '20250101' and already-formatted '010125'
      - '2025-01-01 00:00:00' (stringified datetime)

    Anything unrecognisable is returned unchanged instead of raising, so one
    bad cell can never crash a whole run.

    Examples:
      01-01-2025  ->  010125
       5-3-2025   ->  050325
      2025-01-01  ->  010125
      01/01/25    ->  010125
    """
    if value is None:
        return ""
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, (datetime.date, datetime.datetime)):
        return value.strftime("%d%m%y")

    s = value_to_text(value).strip() if isinstance(value, (int, float)) \
        else str(value).strip()
    if not s:
        return ""

    # Strip a trailing clock part only: '2025-01-01 00:00:00' -> '2025-01-01'
    # (a plain split on whitespace would turn 'Jan 5, 2025' into 'Jan').
    s = re.sub(r"\s+\d{1,2}:\d{2}(:\d{2})?(\.\d+)?$", "", s)
    s = re.sub(r"[Tt]\d{1,2}:\d{2}.*$", "", s)

    parsed = _parse_date_parts(s)
    if parsed is not None:
        day, month, year = parsed
        return f"{day:02d}{month:02d}{year % 100:02d}"

    if re.fullmatch(r"\d+", s):
        if len(s) == 8 and s[:2] in ("19", "20"):      # 20250101 -> 010125
            year, month, day = s[:4], int(s[4:6]), int(s[6:8])
            if 1 <= month <= 12 and 1 <= day <= 31:
                return f"{day:02d}{month:02d}{year[2:]}"
        if len(s) == 6:                                # 010125 - formatted
            month, day = int(s[2:4]), int(s[:2])
            if 1 <= month <= 12 and 1 <= day <= 31:
                return s
    return s


def _parse_date_parts(s):
    """Return (day, month, year) from a string, or None when unsure."""
    def valid(day, month, year):
        if not (1 <= month <= 12 and 1 <= day <= 31 and 0 <= year <= 9999):
            return None
        return day, month, year

    # 01-Jan-2025
    m = re.match(r"^(\d{1,2})[-/. ]([A-Za-z]{3,9})[-/. ](\d{2,4})$", s)
    if m:
        month = MONTH_NAMES.get(m.group(2).lower())
        if month:
            return valid(int(m.group(1)), month, int(m.group(3)))

    # Jan 1 2025 / Jan 1, 2025
    m = re.match(r"^([A-Za-z]{3,9})[-/. ](\d{1,2})[-/., ]+(\d{2,4})$", s)
    if m:
        month = MONTH_NAMES.get(m.group(1).lower())
        if month:
            return valid(int(m.group(2)), month, int(m.group(3)))

    # 01-01-2025  or  2025-01-01
    m = re.match(r"^(\d{1,4})[-/.](\d{1,2})[-/.](\d{1,4})$", s)
    if m:
        a, b, c = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if a > 31:                       # ISO: YYYY-MM-DD
            return valid(c, b, a)
        return valid(a, b, c)

    return None


# ── KeyMap parser ────────────────────────────────────────────────────────────
def split_actions(script_str):
    """
    Split a KeyMap on '+' separators, ignoring '+' inside quotes.

    "Ctrl+c" still splits into Ctrl / c - the modifier grouping below turns
    that back into one combo, exactly like the old behaviour.
    """
    parts, buf, quote = [], [], None
    for ch in script_str:
        if quote:
            buf.append(ch)
            if ch == quote:
                quote = None
        elif ch in "\"'":
            quote = ch
            buf.append(ch)
        elif ch == "+":
            parts.append("".join(buf))
            buf = []
        else:
            buf.append(ch)
    parts.append("".join(buf))
    return [p.strip() for p in parts]


def _unknown_key_error(token):
    return ScriptError(
        f"Unknown key '{token}' in KeyMap. Use a key name "
        "(Enter, Tab, Esc, Space, Up/Down/Left/Right, F1-F12, ...), a single "
        "character, a combo (Ctrl+c) or t\"text\" to type literal text."
    )


def parse_simple_token(token):
    """Parse one non-combo token into an action tuple."""
    text = token.strip()
    if not text:
        return None
    low = text.lower()
    if low in KEY_MAP:
        return ("key", KEY_MAP[low])
    if len(text) == 1:
        return ("key", text)
    if low in known_key_names():
        return ("key", low)
    raise _unknown_key_error(text)


def _parse_group(tokens, i):
    """
    Parse the action that starts at tokens[i].

    Returns (action_or_None, index_of_next_token). A modifier token pulls in
    every following modifier plus the next key token to build a combo.
    """
    token = tokens[i].strip()
    if not token:
        return None, i + 1

    # v"COLNAME" / v'COLNAME'
    m = re.fullmatch(r'[vV](["\'])(.*?)\1', token)
    if m:
        return ("type_col", m.group(2).strip()), i + 1

    # t"literal text" / t'literal text'
    m = re.fullmatch(r'[tT](["\'])(.*)\1', token, re.S)
    if m:
        return ("type_literal", m.group(2)), i + 1

    # modifier(s) followed by a key -> combo
    if token.lower() in MODIFIER_MAP:
        mods = []
        j = i
        while j < len(tokens) and tokens[j].strip().lower() in MODIFIER_MAP:
            mods.append(MODIFIER_MAP[tokens[j].strip().lower()])
            j += 1
        if j >= len(tokens):
            # Trailing modifier with no key: press it as a plain key and let
            # the next loop iteration deal with any further modifiers.
            return ("key", MODIFIER_MAP[token.lower()]), i + 1
        key_action = _parse_group(tokens, j)[0]
        if key_action is None or key_action[0] != "key":
            raise ScriptError(
                f"Modifier combo '{'+'.join(mods)}' must be followed by a key "
                f"(got '{tokens[j]}')."
            )
        return ("combo", mods, key_action[1]), j + 1

    return parse_simple_token(token), i + 1


def parse_script(script_str):
    """
    Parse the script string into a list of action tuples:
      ('key',   key_name)
      ('combo', [mod1, mod2, ...], key_name)
      ('type_col', 'COLNAME')
      ('type_literal', 'text')
      ('repeat', n, sub_action)
    """
    tokens = split_actions(str(script_str))
    actions = []
    i = 0
    while i < len(tokens):
        token = tokens[i].strip()
        if not token:
            i += 1
            continue

        repeat = 1
        original = token
        m = re.match(r"^(\d+)\s*\*\s*(.+)$", token, re.S)
        if m:
            repeat = int(m.group(1))
            token = m.group(2).strip()
            tokens[i] = token
            if not token:
                i += 1
                continue

        if repeat > MAX_REPEAT:
            raise ScriptError(
                f"Repeat count {repeat} in '{original}' is larger than the "
                f"safety limit ({MAX_REPEAT})."
            )

        action, i = _parse_group(tokens, i)
        if action is None:
            continue
        if repeat == 0:
            continue
        if repeat == 1:
            actions.append(action)
        else:
            actions.append(("repeat", repeat, action))
    return actions


# ── Low level key sending (all stop-aware) ───────────────────────────────────
def _press(key):
    if DRY_RUN:
        DRY_RUN_LOG.append(("key", key))
        print(f"      [dry-run] press {key}")
        return
    pyautogui.press(key)


def _press_char(ch):
    """
    Press one character of a longer text without pyautogui's extra PAUSE,
    so typing speed matches the old typewrite(interval=...) behaviour.
    """
    try:
        pyautogui.press(ch, _pause=False)
    except TypeError:                      # older/newer pyautogui signature
        pyautogui.press(ch)


def _combo(mods, key):
    if DRY_RUN:
        DRY_RUN_LOG.append(("combo", list(mods), key))
        print(f"      [dry-run] hotkey {'+'.join(mods)}+{key}")
        return
    for mod in mods:
        pyautogui.keyDown(mod)
        _HELD_MODIFIERS.append(mod)
    try:
        pyautogui.press(key)
    finally:
        for mod in reversed(mods):
            try:
                pyautogui.keyUp(mod)
            finally:
                if mod in _HELD_MODIFIERS:
                    _HELD_MODIFIERS.remove(mod)


def _type_text(text):
    """Type text one character at a time so a mouse move can interrupt it."""
    if not text:
        return
    if DRY_RUN:
        DRY_RUN_LOG.append(("type", text))
        print(f"      [dry-run] type {text!r}")
        for _ in text:
            check_stop()
        return
    for ch in text:
        check_stop()
        _press_char(ch)
        if TYPING_INTERVAL > 0:
            time.sleep(TYPING_INTERVAL)


def _needs_clipboard(text):
    """True when the text contains characters pyautogui cannot type."""
    return any(ord(ch) > 126 or ch in "\t\r\n" for ch in text)


def _paste_text(text):
    """Paste via the clipboard (fast + handles any character), then restore."""
    if DRY_RUN:
        DRY_RUN_LOG.append(("paste", text))
        print(f"      [dry-run] paste {text!r}")
        return
    if not HAS_PYPERCLIP:
        warn_once(
            "pyperclip is not installed - falling back to slow typing "
            "(pip install pyperclip)."
        )
        _type_text(text)
        return

    previous = None
    try:
        previous = pyperclip.paste()
    except Exception:
        previous = None

    try:
        pyperclip.copy(text)
    except Exception as exc:
        warn_once(f"Clipboard unavailable ({exc}) - typing the text instead.")
        _type_text(text)
        return

    try:
        _combo(["ctrl"], "v")
        time.sleep(PASTE_SETTLE)
    finally:
        if previous is not None:
            try:
                pyperclip.copy(previous)
            except Exception:
                pass


# ── Action executor ──────────────────────────────────────────────────────────
def execute_actions(actions, row_data, sheet=None):
    """Run parsed actions for one row. Raises AutomationStopped on a stop."""
    if sheet is None:
        sheet = SheetData(list(row_data.keys()), [row_data])

    for action in actions:
        check_stop()
        kind = action[0]

        if kind == "type_col":
            col_name = action[1]
            if not sheet.has(col_name):
                warn_once(
                    f"Column '{col_name}' is not in the sheet - skipping "
                    f"(available: {', '.join(sheet.headers) or 'none'})."
                )
                continue
            raw_value = sheet.get(row_data, col_name)

            if col_name.strip().lower() in DATE_COLUMNS:
                text = format_date_ddmmyy(raw_value)
            else:
                text = value_to_text(raw_value)

            if not text:
                continue

            if col_name.strip().lower() in PASTE_COLUMNS or (
                HAS_PYPERCLIP and _needs_clipboard(text)
            ):
                _paste_text(text)
            else:
                _type_text(text)

        elif kind == "type_literal":
            text = action[1]
            if not text:
                continue
            if HAS_PYPERCLIP and _needs_clipboard(text):
                _paste_text(text)
            else:
                _type_text(text)

        elif kind == "key":
            _press(action[1])

        elif kind == "combo":
            _combo(action[1], action[2])

        elif kind == "repeat":
            count, sub = action[1], action[2]
            for _ in range(count):
                check_stop()
                execute_actions([sub], row_data, sheet)

        if ACTION_GAP > 0:
            time.sleep(ACTION_GAP)


# ── Trigger key resolver ─────────────────────────────────────────────────────
def normalize_trigger(trigger_str):
    """
    Turn the Trigger cell into ('cont', None) | ('digits', '2211') |
    ('key', pyautogui_key_name).
    """
    t = str(trigger_str).strip().lower()
    if not t:
        return ("none", None)
    if t == "cont":
        return ("cont", None)
    if re.fullmatch(r"\d+", t):
        return ("digits", t)          # typed digit sequence, e.g. 2211
    if t in KEY_MAP:
        return ("key", KEY_MAP[t])
    if len(t) == 1:
        return ("key", t)
    if t in known_key_names():
        return ("key", t)
    warn_once(
        f"Trigger '{trigger_str}' is not a known key name - it will never fire. "
        "Use e.g. Esc, F5, Enter, PB/IT, cont or a digit sequence."
    )
    return ("key", t)


def _read_keyboard_event():
    """Blocking read of the next keyboard event, with a friendly failure path."""
    failures = 0
    while True:
        try:
            return keyboard.read_event(suppress=False)
        except Exception as exc:
            failures += 1
            if failures == 1:
                print(f"    ! keyboard module error ({exc}) - retrying...")
            if failures >= 5:
                raise ScriptError(
                    f"keyboard module keeps failing ({exc}). On Windows run the "
                    "terminal as Administrator; on Linux the keyboard module "
                    "needs root. Or uninstall it to use the ENTER fallback."
                )
            time.sleep(0.25)


def _event_key_name(event):
    return str(getattr(event, "name", "") or "").strip().lower()


def _is_key_down(event):
    return getattr(event, "event_type", None) == getattr(keyboard, "KEY_DOWN", "down")


def wait_for_trigger_key(key_name):
    """
    Block until the user presses `key_name`.

    Accepts every spelling the keyboard module may report for that key, so
    'Esc' matches the 'esc' event the module produces (comparing the
    pyautogui name 'escape' against 'esc' would otherwise wait forever).
    """
    wanted = accepted_trigger_names(key_name)
    print(f"  -> Waiting for [{key_name.upper()}] key... (Ctrl+C quits)", flush=True)
    while True:
        check_stop()
        event = _read_keyboard_event()
        if not _is_key_down(event):
            continue
        if _event_key_name(event) in wanted:
            time.sleep(TRIGGER_SETTLE)
            return


def digit_buffer_step(buffer, key_name, sequence):
    """
    Feed one key name into the digit-sequence buffer.

    Returns (new_buffer, matched). Non-digit keys reset the buffer so a stray
    keystroke can never complete the sequence.
    """
    m = re.fullmatch(r"(?:num\s*)?(\d)", key_name.strip().lower())
    if not m:
        return "", False
    buffer = (buffer + m.group(1))[-max(20, 4 * len(sequence)):]
    if buffer.endswith(sequence):
        return "", True
    return buffer, False


def wait_for_trigger_digits(sequence):
    """Block until the user types the digit sequence (e.g. 2211)."""
    print(f"  -> Waiting for digit sequence [{sequence}]... (Ctrl+C quits)",
          flush=True)
    buffer = ""
    while True:
        check_stop()
        event = _read_keyboard_event()
        if not _is_key_down(event):
            continue
        buffer, matched = digit_buffer_step(buffer, _event_key_name(event), sequence)
        if matched:
            time.sleep(TRIGGER_SETTLE)
            return


def wait_for_trigger(spec):
    kind, value = spec
    if kind == "cont":
        return
    if not HAS_KEYBOARD:
        wait_for_trigger_input(value)
        return
    if kind == "digits":
        wait_for_trigger_digits(value)
    else:
        wait_for_trigger_key(value)


def wait_for_trigger_input(trigger_key):
    """Fallback when the `keyboard` module is missing: press ENTER here."""
    try:
        input(f"  -> Press ENTER in this terminal to trigger "
              f"(configured key: {str(trigger_key).upper()})... ")
    except EOFError:
        raise ScriptError(
            "No trigger input available (stdin is closed). Install the "
            "'keyboard' module (pip install keyboard) to use real key triggers."
        ) from None
    except Exception as exc:
        raise ScriptError(
            f"Could not read the trigger from the terminal: {exc}"
        ) from exc


# ── Excel loader ─────────────────────────────────────────────────────────────
class SheetData:
    """Rows plus a case-insensitive header lookup."""

    def __init__(self, headers, rows):
        self.headers = list(headers)
        self.rows = rows
        self._index = {h.strip().lower(): h for h in self.headers if h.strip()}

    def has(self, name):
        return str(name).strip().lower() in self._index

    def get(self, row, name, default=""):
        header = self._index.get(str(name).strip().lower())
        if header is None:
            return default
        value = row.get(header, default)
        return default if value is None else value

    def preview(self, row, skip=()):
        skip_lower = {str(s).strip().lower() for s in skip}
        return {
            header: row.get(header)
            for header in self.headers
            if header.strip().lower() not in skip_lower
            and not header.startswith("__")
        }


def load_excel(path, sheet_name=None):
    if openpyxl is None:
        raise ScriptError(
            "openpyxl could not be imported "
            f"({OPENPYXL_ERROR}). Install it with:  pip install openpyxl"
        )
    if not os.path.isfile(path):
        raise ScriptError(f"File not found: {path}")

    try:
        # data_only=True -> formula cells give their cached RESULT, not "=A1+B2"
        wb = openpyxl.load_workbook(path, data_only=True)
    except Exception as exc:
        raise ScriptError(f"Could not open '{path}': {exc}") from exc

    try:
        if sheet_name:
            if sheet_name not in wb.sheetnames:
                raise ScriptError(
                    f"Sheet '{sheet_name}' not found. Available sheets: "
                    f"{', '.join(wb.sheetnames)}"
                )
            ws = wb[sheet_name]
        else:
            ws = wb.active
        sheet_label = ws.title

        seen, headers = set(), []
        for position, cell in enumerate(ws[1], start=1):
            raw = cell.value
            name = str(raw).strip() if raw is not None else ""
            if not name:
                name = f"Column{position}"      # blank headers stay usable
            unique, suffix = name, 2
            while unique.lower() in seen:
                unique = f"{name}_{suffix}"
                suffix += 1
            seen.add(unique.lower())
            headers.append(unique)

        rows = []
        for number, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            if row is None:
                continue
            if all(v is None or (isinstance(v, str) and not v.strip()) for v in row):
                continue
            row_dict = {"__excel_row__": number}
            for i, header in enumerate(headers):
                row_dict[header] = row[i] if i < len(row) and row[i] is not None else ""
            rows.append(row_dict)

        return SheetData(headers, rows), sheet_label
    finally:
        wb.close()


# ── Default Excel path ───────────────────────────────────────────────────────
DEFAULT_EXCEL = r"C:\Users\admin\Downloads\win\WPy64-31312\WPy64-3.13.12.0\b.xlsx"


# ── CLI ──────────────────────────────────────────────────────────────────────
def build_arg_parser():
    parser = argparse.ArgumentParser(
        prog=os.path.basename(sys.argv[0] or "E.py"),
        description="AutoKeyPress - Excel-driven keyboard automation",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "EMERGENCY STOP: move the mouse while a row is typing and the run "
            "stops immediately.\n"
            "Hard abort backup: slam the mouse into any SCREEN CORNER."
        ),
    )
    parser.add_argument(
        "excel_file", nargs="?", default=DEFAULT_EXCEL,
        help=f"path to your .xlsx file (default: {DEFAULT_EXCEL})",
    )
    parser.add_argument(
        "--delay", type=float, default=3.0,
        help="seconds to wait after the trigger before typing (default: 3)",
    )
    parser.add_argument(
        "--sheet", type=str, default=None,
        help="sheet name to read (default: first/active sheet)",
    )
    parser.add_argument(
        "--mouse-threshold", type=float, default=5.0, metavar="PX",
        help="mouse movement in pixels that stops the automation (default: 5)",
    )
    parser.add_argument(
        "--no-mouse-stop", action="store_true",
        help="disable the move-the-mouse emergency stop",
    )
    parser.add_argument(
        "--ask-resume", action="store_true",
        help="after a mouse stop, ask whether to continue with the next row",
    )
    parser.add_argument(
        "--typing-interval", type=float, default=0.03, metavar="SEC",
        help="delay between typed characters (default: 0.03)",
    )
    parser.add_argument(
        "--paste-columns", type=str, default="Narration", metavar="LIST",
        help="comma separated columns pasted via the clipboard "
             "(default: Narration)",
    )
    parser.add_argument(
        "--date-columns", type=str, default="Date", metavar="LIST",
        help="comma separated columns reformatted to DDMMYY (default: Date)",
    )
    parser.add_argument(
        "--dry-run", action="store_true",
        help="print what would be typed without sending any keys",
    )
    parser.add_argument(
        "--version", action="version", version=f"AutoKeyPress {VERSION}",
    )
    return parser


def apply_config(args):
    global DRY_RUN, TYPING_INTERVAL, PASTE_COLUMNS, DATE_COLUMNS
    DRY_RUN = bool(args.dry_run)
    TYPING_INTERVAL = max(0.0, args.typing_interval)
    PASTE_COLUMNS = {c.strip().lower() for c in args.paste_columns.split(",") if c.strip()}
    DATE_COLUMNS = {c.strip().lower() for c in args.date_columns.split(",") if c.strip()}


# ── Main ─────────────────────────────────────────────────────────────────────
def _countdown(seconds):
    """
    Wait `seconds` before typing so the user can switch to the target window.

    Fractional delays are honoured too (the old int() cast silently turned a
    0.5s delay into no delay at all).
    """
    remaining = float(seconds)
    if remaining <= 0:
        return
    print(f"  v Triggered! Starting in {remaining:g}s - switch to your target "
          f"window!", flush=True)
    while remaining > 0:
        step = min(1.0, remaining)
        print(f"    {remaining:g}...", end="\r", flush=True)
        time.sleep(step)
        remaining -= step
    print(" " * 30, end="\r", flush=True)


def run_rows(args, sheet, prepared, start_index=0, completed=0, mouse_stop=True):
    """
    Run rows from `start_index` onwards.

    Returns (exit_code, completed, stopped_at) where `stopped_at` is the index
    of the row that was interrupted by a stop (None when the run finished).
    """
    total = len(sheet.rows)
    watcher = MouseStopWatcher(threshold=args.mouse_threshold)
    idx = start_index
    try:
        while idx < total:
            entry = prepared[idx]
            if entry is None:
                idx += 1
                continue
            row = sheet.rows[idx]
            number = idx + 1
            spec, actions, trigger_raw, script_str = entry

            print(f"  [Row {number}/{total}] Ready")
            print(f"    Trigger : {trigger_raw}")
            print(f"    KeyMap  : {script_str}")
            print(f"    Data    : {sheet.preview(row, skip=('Trigger', 'KeyMap'))}")

            if spec[0] == "cont":
                print("  * CONT row - executing instantly (no trigger needed)")
            elif spec[0] == "none":
                print(f"  [Row {number}/{total}] Skipping - no trigger defined.")
                idx += 1
                continue
            else:
                wait_for_trigger(spec)
                _countdown(args.delay)
                print("  > Typing now...")

            if mouse_stop:
                print("  (mouse stop armed - keep the mouse still)")
                watcher.watch()
            try:
                execute_actions(actions, row, sheet)
            finally:
                watcher.cancel()

            completed += 1
            print(f"  v Row {number} done.\n")
            idx += 1
            time.sleep(0.2)

        return 0, completed, None

    except AutomationStopped as stop:
        watcher.cancel()
        release_modifiers()
        print(f"\n  STOPPED: {stop}")
        print(f"  Completed {completed}/{total} row(s). No keys are held down.")
        return 2, completed, idx

    except FAILSAFE_ERRORS:
        watcher.cancel()
        release_modifiers()
        print("\n  STOPPED: mouse hit a screen corner (pyautogui failsafe).")
        print(f"  Completed {completed}/{total} row(s). No keys are held down.")
        return 2, completed, idx

    except KeyboardInterrupt:
        watcher.cancel()
        release_modifiers()
        print(f"\n  Interrupted (Ctrl+C). Completed {completed}/{total} row(s).")
        return 130, completed, None

    except ScriptError as exc:
        watcher.cancel()
        release_modifiers()
        print(f"\n  ERROR: {exc}")
        return 1, completed, None

    finally:
        watcher.shutdown()


def main(argv=None):
    args = build_arg_parser().parse_args(argv)
    apply_config(args)

    script_name = os.path.basename(sys.argv[0] or "E.py")

    if not HAS_PYAUTOGUI and not args.dry_run:
        print(f"\n  ERROR: pyautogui could not be imported ({PYAUTOGUI_ERROR}).")
        print("  Install it with:  pip install pyautogui")
        print("  (On Linux also:  sudo apt install python3-xlib scrot)")
        print("  (Use --dry-run to preview a sheet without sending keys.)")
        return 1

    mouse_stop = (
        not args.no_mouse_stop
        and not args.dry_run
        and args.mouse_threshold > 0
        and HAS_PYAUTOGUI
    )

    print("\n" + "=" * 60)
    print(f"  AutoKeyPress - Excel Keyboard Automation  v{VERSION}")
    print("=" * 60)
    print(f"  File  : {args.excel_file}")
    print(f"  Delay : {args.delay:g}s after trigger key")
    if args.dry_run:
        print("  Mode  : DRY RUN - no keys will be sent")
    if mouse_stop:
        print(f"  STOP  : move the mouse > {args.mouse_threshold:g}px while typing")
        print("          hard abort: slam the mouse into any screen corner")
    elif not args.dry_run:
        print("  STOP  : mouse stop disabled - slam the mouse into any screen")
        print("          corner to abort")
    if not HAS_KEYBOARD:
        print("  NOTE  : 'keyboard' module not installed - triggers must be")
        print("          confirmed with ENTER in this terminal.")
    if not HAS_PYPERCLIP:
        print("  NOTE  : 'pyperclip' not installed - long/unicode text is typed")
        print("          character by character.")
    print("=" * 60 + "\n")

    try:
        sheet, sheet_label = load_excel(args.excel_file, args.sheet)
    except ScriptError as exc:
        print(f"\n  ERROR: {exc}")
        print("\n  Either:")
        print(f"    1. Save your Excel file as:  {DEFAULT_EXCEL}")
        print("    2. Or run with a custom path:")
        print(f"       python {script_name} \"C:\\path\\to\\file.xlsx\"")
        return 1

    print(f"  Sheet : {sheet_label}")
    if not sheet.has("Trigger"):
        print("  ERROR: No 'Trigger' column found in the Excel file!")
        print(f"  Found columns: {sheet.headers}")
        return 1
    if not sheet.has("KeyMap"):
        print("  ERROR: No 'KeyMap' column found in the Excel file!")
        print(f"  Found columns: {sheet.headers}")
        return 1

    rows = sheet.rows
    total = len(rows)
    print(f"  Loaded {total} row(s)\n")

    # ── Parse every KeyMap up front: a typo must never surface mid-run ──────
    prepared = []
    for idx, row in enumerate(rows, 1):
        # value_to_text so a numeric trigger cell (2211) stays "2211"
        trigger_raw = value_to_text(sheet.get(row, "Trigger", "")).strip()
        script_str = str(sheet.get(row, "KeyMap", "")).strip()
        if not trigger_raw or not script_str:
            print(f"  [Row {idx}/{total}] Skipping - no Trigger or KeyMap defined.")
            prepared.append(None)
            continue
        try:
            actions = parse_script(script_str)
        except ScriptError as exc:
            excel_row = row.get("__excel_row__", idx + 1)
            print(f"\n  ERROR in the KeyMap of row {idx} (Excel row {excel_row}):")
            print(f"    {script_str}")
            print(f"    {exc}\n")
            return 1
        if not actions:
            print(f"  [Row {idx}/{total}] Skipping - KeyMap produced no actions.")
            prepared.append(None)
            continue
        prepared.append(
            (normalize_trigger(trigger_raw), actions, trigger_raw, script_str)
        )

    code, completed, stopped_at = run_rows(
        args, sheet, prepared, 0, 0, mouse_stop
    )

    # --ask-resume: offer to continue with the row after the one that stopped
    if code == 2 and args.ask_resume and stopped_at is not None and stopped_at + 1 < total:
        try:
            answer = input(f"\n  Continue with row {stopped_at + 2}/{total}? [y/N] ")
        except (EOFError, KeyboardInterrupt):
            answer = "n"
        if answer.strip().lower().startswith("y"):
            reset_stop()
            print()
            code, completed, _ = run_rows(
                args, sheet, prepared, stopped_at + 1, completed, mouse_stop
            )

    if code == 0:
        print("=" * 60)
        print(f"  All done - {completed}/{total} row(s) processed.")
        print("=" * 60 + "\n")
    return code


if __name__ == "__main__":
    try:
        sys.exit(main())
    except SystemExit:
        raise
    except KeyboardInterrupt:
        release_modifiers()
        print("\nInterrupted.")
        sys.exit(130)
    except Exception as exc:  # never leave the keyboard with stuck modifiers
        release_modifiers()
        print(f"\nUnexpected error: {type(exc).__name__}: {exc}")
        sys.exit(1)
