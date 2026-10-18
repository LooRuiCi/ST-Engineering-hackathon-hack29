"""
Office GUI Automation Demo  (fully automatic, human pace)
Windows + Apache OpenOffice 4.1.x + Chrome

Just press Run (or double-click the .exe). Nothing is asked.
A full run takes at least MIN_RUNTIME_MINUTES (default 16).

FLOW
----
1.  Show desktop, double-click the OpenOffice icon (found automatically)
2.  Start Center -> Spreadsheet, maximize with the mouse
3.  Fill a small budget table, save as e.g. "Budget 23", click X
4.  Same for Writer: ~200-word meeting report, save as "Meeting 67", click X
5.  Show desktop, double-click the Chrome icon
6.  Website 1 in the first tab; click "+" for each next website
7.  Click the address bar, type the address, read and scroll with the
    wheel, click a link, click Back, close a tab with its X
8.  Click Chrome's X

Anything that can't be found on screen falls back to the keyboard
or to starting the program directly, and that is written in the log.

STOP
----
Press Esc at any time: the program is killed immediately.

LOG
---
office_demo_log.csv next to the script / .exe:
time, elapsed, program (.exe that had the window), action, detail
"""

import csv
import ctypes
import datetime
import glob
import math
import os
import random
import subprocess
import sys
import threading
import time
from ctypes import wintypes

if os.name != "nt":
    raise SystemExit("This script only runs on Windows.")

import winreg

# When built with "pyinstaller --noconsole" there is no console;
# make print() safe in that case.
if sys.stdout is None:
    sys.stdout = open(os.devnull, "w")
if sys.stderr is None:
    sys.stderr = open(os.devnull, "w")


# ============================================================
# SETTINGS  (nothing is asked at runtime)
# ============================================================

SLOWDOWN = 3.0               # 3x slower than a fast run
TYPING_SLOWDOWN = 1.5        # typing: twice as fast as SLOWDOWN
MIN_RUNTIME_MINUTES = 16     # extra website reading fills any gap
START_DELAY = 3              # seconds before starting
SITES_PER_RUN = 3
LOOPS = 1

HOME = os.path.expanduser("~")
DOCUMENTS = os.path.join(HOME, "Documents")

if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

LOG_FILE = os.path.join(BASE_DIR, "office_demo_log.csv")

SITES = [
    "https://www.coinmarketcap.com",
    "https://www.coursera.org",
    "https://www.coinbase.com",
    "https://coingecko.com",
    "https://www.samsung.com",
    "https://aws.amazon.com",
    "https://www.amazon.com",
    "https://www.apple.com",
    "https://apnews.com",
    "https://www.dell.com",
]

ITEMS = [
    "Laptop", "Monitor", "Keyboard", "Mouse",
    "Headset", "Cable", "Router", "Chair",
]

CALC_NAMES = ["Budget {n}", "Office Budget {n}", "Equipment List {n}"]
WRITER_NAMES = ["Meeting {n}", "Meeting Notes {n}", "Weekly Meeting {n}"]

REPORT_TITLE = "Weekly Project Meeting Report"

MEETING_PARAGRAPHS = [
    (
        "The team reviewed the current project timeline and "
        "discussed progress made since the previous meeting. "
        "Most planned activities are progressing according to "
        "schedule. Several items still require additional "
        "information from team members before they can be "
        "completed."
    ),
    (
        "The preliminary analysis is progressing well and the "
        "available information has been organized for review. "
        "The team agreed that the upcoming presentation should "
        "focus on the most important findings and provide a "
        "clear summary of the current project status."
    ),
    (
        "Communication remains important because several tasks "
        "depend on updates from different members of the team. "
        "Everyone was encouraged to provide updates as soon as "
        "meaningful progress is made. This should help identify "
        "potential delays and resolve issues before they affect "
        "the overall schedule."
    ),
    (
        "The next meeting will focus on reviewing the completed "
        "analysis and preparing the first version of the final "
        "presentation. Supporting documents should also be "
        "prepared before the meeting. Overall, the project "
        "remains on schedule and progress will be reviewed "
        "again during the next team discussion."
    ),
]

ACTION_ITEMS = [
    "Complete the preliminary analysis.",
    "Collect outstanding team updates.",
    "Review the presentation structure.",
    "Prepare supporting documents.",
    "Confirm responsibilities for the next phase.",
]


# ============================================================
# WIN32 SETUP
# ============================================================

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

user32 = ctypes.WinDLL("user32", use_last_error=True)
kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
ole32 = ctypes.WinDLL("ole32")
oleacc = ctypes.WinDLL("oleacc")
oleaut32 = ctypes.WinDLL("oleaut32")

HWND, DWORD, BOOL, UINT = wintypes.HWND, wintypes.DWORD, wintypes.BOOL, wintypes.UINT
HANDLE, LPCWSTR, LPWSTR = wintypes.HANDLE, wintypes.LPCWSTR, wintypes.LPWSTR
POINT, RECT = wintypes.POINT, wintypes.RECT

WNDENUMPROC = ctypes.WINFUNCTYPE(BOOL, HWND, wintypes.LPARAM)

user32.EnumWindows.argtypes = [WNDENUMPROC, wintypes.LPARAM]
user32.GetForegroundWindow.restype = HWND
user32.GetWindowTextLengthW.argtypes = [HWND]
user32.GetWindowTextW.argtypes = [HWND, LPWSTR, ctypes.c_int]
user32.IsWindow.argtypes = [HWND]
user32.IsWindowVisible.argtypes = [HWND]
user32.IsZoomed.argtypes = [HWND]
user32.IsIconic.argtypes = [HWND]
user32.ShowWindow.argtypes = [HWND, ctypes.c_int]
user32.SetForegroundWindow.argtypes = [HWND]
user32.GetWindowRect.argtypes = [HWND, ctypes.POINTER(RECT)]
user32.GetAsyncKeyState.restype = ctypes.c_short
user32.FindWindowW.argtypes = [LPCWSTR, LPCWSTR]
user32.FindWindowW.restype = HWND
user32.FindWindowExW.argtypes = [HWND, HWND, LPCWSTR, LPCWSTR]
user32.FindWindowExW.restype = HWND
user32.GetWindowThreadProcessId.argtypes = [HWND, ctypes.POINTER(DWORD)]
user32.SendMessageW.argtypes = [HWND, UINT, wintypes.WPARAM, wintypes.LPARAM]
user32.SendMessageW.restype = wintypes.LPARAM
user32.ClientToScreen.argtypes = [HWND, ctypes.POINTER(POINT)]

kernel32.GetConsoleWindow.restype = HWND
kernel32.OpenProcess.argtypes = [DWORD, BOOL, DWORD]
kernel32.OpenProcess.restype = HANDLE
kernel32.CloseHandle.argtypes = [HANDLE]
kernel32.QueryFullProcessImageNameW.argtypes = [
    HANDLE, DWORD, LPWSTR, ctypes.POINTER(DWORD)
]
kernel32.VirtualAllocEx.argtypes = [
    HANDLE, ctypes.c_void_p, ctypes.c_size_t, DWORD, DWORD
]
kernel32.VirtualAllocEx.restype = ctypes.c_void_p
kernel32.VirtualFreeEx.argtypes = [HANDLE, ctypes.c_void_p, ctypes.c_size_t, DWORD]
kernel32.ReadProcessMemory.argtypes = [
    HANDLE, ctypes.c_void_p, ctypes.c_void_p, ctypes.c_size_t,
    ctypes.POINTER(ctypes.c_size_t),
]
kernel32.WriteProcessMemory.argtypes = [
    HANDLE, ctypes.c_void_p, ctypes.c_void_p, ctypes.c_size_t,
    ctypes.POINTER(ctypes.c_size_t),
]

SW_MINIMIZE, SW_RESTORE, SW_MAXIMIZE = 6, 9, 3
VK_ESCAPE = 0x1B


class _KBD(ctypes.Structure):
    _fields_ = [
        ("wVk", ctypes.c_uint16), ("wScan", ctypes.c_uint16),
        ("dwFlags", ctypes.c_uint32), ("time", ctypes.c_uint32),
        ("dwExtraInfo", ctypes.c_size_t),
    ]


class _MOU(ctypes.Structure):
    _fields_ = [
        ("dx", ctypes.c_int32), ("dy", ctypes.c_int32),
        ("mouseData", ctypes.c_uint32), ("dwFlags", ctypes.c_uint32),
        ("time", ctypes.c_uint32), ("dwExtraInfo", ctypes.c_size_t),
    ]


class _UNI(ctypes.Union):
    _fields_ = [("ki", _KBD), ("mi", _MOU)]


class _INPUT(ctypes.Structure):
    _fields_ = [("type", ctypes.c_uint32), ("u", _UNI)]


user32.SendInput.argtypes = [UINT, ctypes.POINTER(_INPUT), ctypes.c_int]


# ============================================================
# INPUT CONTROLLER
# ============================================================

class InputController:

    VK = {
        "enter": 0x0D, "tab": 0x09, "backspace": 0x08,
        "win": 0x5B, "ctrl": 0x11, "alt": 0x12, "shift": 0x10,
        "space": 0x20,
        "pgup": 0x21, "pgdn": 0x22, "end": 0x23, "home": 0x24,
        "left": 0x25, "up": 0x26, "right": 0x27, "down": 0x28,
        "delete": 0x2E, "f12": 0x7B,
    }
    EXTENDED = {0x21, 0x22, 0x23, 0x24, 0x25, 0x26, 0x27, 0x28, 0x2E, 0x5B}

    @staticmethod
    def _kb(vk=0, scan=0, flags=0):
        item = _INPUT(type=1)
        item.u.ki = _KBD(vk, scan, flags, 0, 0)
        return item

    @staticmethod
    def _ms(flags, data=0):
        item = _INPUT(type=0)
        item.u.mi = _MOU(0, 0, data & 0xFFFFFFFF, flags, 0, 0)
        return item

    def _send(self, *items):
        arr = (_INPUT * len(items))(*items)
        user32.SendInput(len(items), arr, ctypes.sizeof(_INPUT))

    def position(self):
        p = POINT()
        user32.GetCursorPos(ctypes.byref(p))
        return (p.x, p.y)

    def size(self):
        return (user32.GetSystemMetrics(0), user32.GetSystemMetrics(1))

    def move(self, x, y):
        user32.SetCursorPos(int(x), int(y))

    def click(self):
        self._send(self._ms(0x0002))
        time.sleep(random.uniform(0.05, 0.11))
        self._send(self._ms(0x0004))

    def double_click(self):
        self.click()
        time.sleep(random.uniform(0.08, 0.14))
        self.click()

    def scroll(self, amount):
        self._send(self._ms(0x0800, int(amount)))

    def _vk(self, name):
        name = name.lower()
        return self.VK[name] if name in self.VK else ord(name.upper())

    def _key(self, vk, up=False):
        flags = 0x0001 if vk in self.EXTENDED else 0
        if up:
            flags |= 0x0002
        self._send(self._kb(vk, user32.MapVirtualKeyW(vk, 0), flags))

    def hotkey(self, *names):
        vks = [self._vk(n) for n in names]
        for vk in vks:
            self._key(vk)
            time.sleep(random.uniform(0.04, 0.09))
        time.sleep(random.uniform(0.05, 0.10))
        for vk in reversed(vks):
            self._key(vk, up=True)
            time.sleep(random.uniform(0.03, 0.07))

    def press(self, name):
        self.hotkey(name)

    def write(self, text):
        for ch in text:
            c = ord(ch)
            self._send(self._kb(0, c, 0x0004))
            time.sleep(random.uniform(0.02, 0.06))
            self._send(self._kb(0, c, 0x0004 | 0x0002))

    def release_all(self):
        try:
            for vk in (0x11, 0x12, 0x10, 0x5B):
                self._key(vk, up=True)
            self._send(self._ms(0x0004))
        except Exception:
            pass


ui = InputController()


# ============================================================
# LOGGING
# ============================================================

RUN_START = time.time()


def window_exe(hwnd):
    """Name of the .exe that owns a window, e.g. chrome.exe."""
    if not hwnd:
        return ""
    pid = DWORD()
    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    h = kernel32.OpenProcess(0x1000, False, pid.value)
    if not h:
        return ""
    try:
        buf = ctypes.create_unicode_buffer(520)
        size = DWORD(520)
        if kernel32.QueryFullProcessImageNameW(h, 0, buf, ctypes.byref(size)):
            return os.path.basename(buf.value)
        return ""
    finally:
        kernel32.CloseHandle(h)


def log(action, detail="", hwnd=None):
    global LOG_FILE
    try:
        app = window_exe(hwnd or user32.GetForegroundWindow())
    except Exception:
        app = ""
    row = [
        datetime.datetime.now().isoformat(timespec="milliseconds"),
        f"{time.time() - RUN_START:.1f}",
        app, action, detail,
    ]
    for path in (LOG_FILE, os.path.join(HOME, "office_demo_log.csv")):
        try:
            new_file = not os.path.exists(path)
            with open(path, "a", newline="", encoding="utf-8") as f:
                w = csv.writer(f)
                if new_file:
                    w.writerow(["time", "elapsed_s", "program",
                                "action", "detail"])
                w.writerow(row)
            LOG_FILE = path
            return
        except OSError:
            continue


# ============================================================
# ESC = KILL
# ============================================================

def _esc_watcher():
    while True:
        if user32.GetAsyncKeyState(VK_ESCAPE) & 0x8000:
            ui.release_all()
            print("\n\nESC pressed - program killed.")
            log("killed", "Esc pressed")
            os._exit(0)
        time.sleep(0.03)


def start_esc_watcher():
    threading.Thread(target=_esc_watcher, daemon=True).start()


# ============================================================
# TIMING
# ============================================================

def pause(lo, hi):
    """Human pause, multiplied by SLOWDOWN."""
    middle = (lo + hi) / 2
    d = random.lognormvariate(math.log(middle), 0.25)
    time.sleep(min(max(d, lo * 0.75), hi * 1.5) * SLOWDOWN)


def wait_until(condition, timeout, interval=0.25):
    end = time.perf_counter() + timeout
    while time.perf_counter() < end:
        result = condition()
        if result:
            return result
        time.sleep(interval)
    return condition() or None


def banner(text):
    print()
    print("==============================")
    print(text)
    print("==============================")


# ============================================================
# WINDOW HELPERS
# ============================================================

def window_title(hwnd):
    if not hwnd:
        return ""
    n = user32.GetWindowTextLengthW(hwnd)
    if n <= 0:
        return ""
    buf = ctypes.create_unicode_buffer(n + 1)
    user32.GetWindowTextW(hwnd, buf, n + 1)
    return buf.value


def top_windows():
    found = []

    @WNDENUMPROC
    def callback(hwnd, lparam):
        if user32.IsWindowVisible(hwnd):
            t = window_title(hwnd)
            if t:
                found.append((hwnd, t))
        return True

    user32.EnumWindows(callback, 0)
    return found


def find_window(predicate):
    for hwnd, title in top_windows():
        if predicate(hwnd, title):
            return hwnd
    return None


def foreground_title():
    return window_title(user32.GetForegroundWindow())


def window_alive(hwnd):
    return bool(hwnd and user32.IsWindow(hwnd) and user32.IsWindowVisible(hwnd))


def visible_rect(hwnd):
    rect = RECT()
    try:
        ctypes.windll.dwmapi.DwmGetWindowAttribute(
            HWND(hwnd), ctypes.c_uint(9), ctypes.byref(rect), ctypes.sizeof(rect)
        )
    except Exception:
        pass
    if rect.right <= rect.left:
        user32.GetWindowRect(hwnd, ctypes.byref(rect))
    w, h = ui.size()
    return (max(rect.left, 0), max(rect.top, 0),
            min(rect.right, w), min(rect.bottom, h))


def dpi_scale(hwnd):
    try:
        dpi = user32.GetDpiForWindow(HWND(hwnd))
        if dpi:
            return dpi / 96.0
    except Exception:
        pass
    return 1.0


def point_inside(hwnd, fx=(0.30, 0.70), fy=(0.35, 0.70)):
    left, top, right, bottom = visible_rect(hwnd)
    w, h = right - left, bottom - top
    return (left + random.randint(int(w * fx[0]), int(w * fx[1])),
            top + random.randint(int(h * fy[0]), int(h * fy[1])))


def focus_window(hwnd):
    if user32.IsIconic(hwnd):
        user32.ShowWindow(hwnd, SW_RESTORE)
        time.sleep(0.5)
    user32.SetForegroundWindow(hwnd)
    time.sleep(0.4)
    if user32.GetForegroundWindow() != hwnd:
        left, top, right, _ = visible_rect(hwnd)
        click_at((left + right) // 2 + random.randint(-60, 60),
                 top + int(10 * dpi_scale(hwnd)))
        pause(0.3, 0.5)


def maximize_window(hwnd, label):
    """Click the window's own maximize button."""
    focus_window(hwnd)
    if user32.IsZoomed(hwnd):
        return
    print(f"Maximizing {label}...")
    s = dpi_scale(hwnd)
    left, top, right, _ = visible_rect(hwnd)
    click_at(right - int(69 * s) + random.randint(-4, 4),
             top + int(11 * s) + random.randint(-2, 2))
    if wait_until(lambda: user32.IsZoomed(hwnd), timeout=3):
        log("maximize", f"{label}: clicked maximize button", hwnd)
    else:
        ui.hotkey("win", "up")
        if not wait_until(lambda: user32.IsZoomed(hwnd), timeout=3):
            user32.ShowWindow(hwnd, SW_MAXIMIZE)
        log("maximize", f"{label}: fallback (button not hit)", hwnd)
    pause(0.6, 1.0)


def minimize_console():
    hwnd = kernel32.GetConsoleWindow()
    if hwnd:
        user32.ShowWindow(hwnd, SW_MINIMIZE)


def show_desktop():
    print("Showing desktop...")
    ui.hotkey("win", "d")
    pause(0.8, 1.3)


# ============================================================
# ACCESSIBILITY  (lets the mouse find buttons by their name,
# e.g. Chrome's "New Tab", "Back", "Close", the address bar)
# ============================================================

class VARIANT(ctypes.Structure):
    _fields_ = [
        ("vt", ctypes.c_ushort), ("r1", ctypes.c_ushort),
        ("r2", ctypes.c_ushort), ("r3", ctypes.c_ushort),
        ("data", ctypes.c_void_p * 2),
    ]


oleacc.AccessibleObjectFromPoint.argtypes = [
    POINT, ctypes.POINTER(ctypes.c_void_p), ctypes.POINTER(VARIANT)
]
oleacc.AccessibleObjectFromPoint.restype = ctypes.c_long
oleaut32.SysFreeString.argtypes = [ctypes.c_void_p]

ROLE_LINK = 0x1E
ROLE_PAGETAB = 0x25
ROLE_TEXT = 0x2A
ROLE_BUTTON = 0x2B
STATE_UNAVAILABLE = 0x1


def _vfunc(ptr, index, *argtypes):
    vtbl = ctypes.cast(ptr, ctypes.POINTER(ctypes.POINTER(ctypes.c_void_p))).contents
    proto = ctypes.WINFUNCTYPE(ctypes.c_long, ctypes.c_void_p, *argtypes)
    return proto(vtbl[index])


def _variant_int(v):
    if v.vt == 3:  # VT_I4
        return ctypes.cast(ctypes.addressof(v) + 8,
                           ctypes.POINTER(ctypes.c_long)).contents.value
    return None


def accessible_at(x, y):
    """(name, role, state) of whatever is on screen at x, y."""
    acc = ctypes.c_void_p()
    child = VARIANT()
    try:
        hr = oleacc.AccessibleObjectFromPoint(
            POINT(int(x), int(y)), ctypes.byref(acc), ctypes.byref(child)
        )
        if hr != 0 or not acc.value:
            return None

        bstr = ctypes.c_void_p()
        name = ""
        if _vfunc(acc, 10, VARIANT, ctypes.POINTER(ctypes.c_void_p))(
                acc, child, ctypes.byref(bstr)) == 0 and bstr.value:
            name = ctypes.wstring_at(bstr.value)
            oleaut32.SysFreeString(bstr)

        role_v = VARIANT()
        _vfunc(acc, 13, VARIANT, ctypes.POINTER(VARIANT))(
            acc, child, ctypes.byref(role_v))

        state_v = VARIANT()
        _vfunc(acc, 14, VARIANT, ctypes.POINTER(VARIANT))(
            acc, child, ctypes.byref(state_v))

        _vfunc(acc, 2)(acc)  # Release
        return (name or "", _variant_int(role_v), _variant_int(state_v) or 0)
    except Exception:
        return None


def find_on_row(ys, x_from, x_to, step, match):
    """
    Slide along one or more horizontal rows; return the centre of
    every run of points whose accessible object matches.
    """
    for y in ys:
        runs, start, last = [], None, None
        for x in range(int(x_from), int(x_to), max(2, int(step))):
            info = accessible_at(x, y)
            ok = bool(info) and match(*info)
            if ok and start is None:
                start = x
            if ok:
                last = x
            if not ok and start is not None:
                runs.append(((start + last) // 2, y))
                start = None
        if start is not None:
            runs.append(((start + last) // 2, y))
        if runs:
            return runs
    return []


# ============================================================
# DESKTOP ICONS  (found automatically, no fixed coordinates)
# ============================================================

class LVITEMW(ctypes.Structure):
    _fields_ = [
        ("mask", ctypes.c_uint), ("iItem", ctypes.c_int),
        ("iSubItem", ctypes.c_int), ("state", ctypes.c_uint),
        ("stateMask", ctypes.c_uint), ("pszText", ctypes.c_void_p),
        ("cchTextMax", ctypes.c_int), ("iImage", ctypes.c_int),
        ("lParam", ctypes.c_void_p), ("iIndent", ctypes.c_int),
        ("iGroupId", ctypes.c_int), ("cColumns", ctypes.c_uint),
        ("puColumns", ctypes.c_void_p), ("piColFmt", ctypes.c_void_p),
        ("iGroup", ctypes.c_int),
    ]


def _desktop_listview():
    progman = user32.FindWindowW("Progman", None)
    defview = user32.FindWindowExW(progman, None, "SHELLDLL_DefView", None)
    if not defview:
        w = None
        while True:
            w = user32.FindWindowExW(None, w, "WorkerW", None)
            if not w:
                break
            defview = user32.FindWindowExW(w, None, "SHELLDLL_DefView", None)
            if defview:
                break
    if not defview:
        return None
    return user32.FindWindowExW(defview, None, "SysListView32", None)


ICON_EXCLUDE = ("install", "setup", "uninstall", "files", "readme", "folder")


def find_desktop_icon(keywords):
    """
    Screen position of the desktop icon for a program.

    Several icons can contain the keyword (e.g. "OpenOffice 4.1.16"
    and the folder "OpenOffice 4.1.16 (en-US) Installation Files").
    Icons with words like "install"/"files" are skipped, and the
    shortest remaining name wins.
    """
    matches = _all_desktop_icons(keywords)
    good = [m for m in matches
            if not any(bad in m[0].lower() for bad in ICON_EXCLUDE)]
    if matches:
        log("icon_candidates",
            "; ".join(f"{n} @ {p}" for n, p in matches)
            + f" -> chose {min(good, key=lambda m: len(m[0]))[0] if good else 'none'}")
    if not good:
        return None
    return min(good, key=lambda m: len(m[0]))


def _all_desktop_icons(keywords):
    """All (name, (x, y)) desktop icons whose name contains a keyword."""
    results = []
    # 32-bit Python on 64-bit Windows cannot read the desktop's icon list
    if ctypes.sizeof(ctypes.c_void_p) == 4 and os.environ.get(
            "PROCESSOR_ARCHITEW6432"):
        log("icon_search_skipped", "32-bit Python on 64-bit Windows")
        return results
    try:
        lv = _desktop_listview()
        if not lv:
            return results
        pid = DWORD()
        user32.GetWindowThreadProcessId(lv, ctypes.byref(pid))
        h = kernel32.OpenProcess(0x0008 | 0x0010 | 0x0020 | 0x0400,
                                 False, pid.value)
        if not h:
            return results
        remote = kernel32.VirtualAllocEx(h, None, 4096, 0x3000, 0x04)
        if not remote:
            kernel32.CloseHandle(h)
            return results
        try:
            text_addr, rect_addr = remote + 512, remote + 2048
            count = user32.SendMessageW(lv, 0x1004, 0, 0)  # GETITEMCOUNT
            for i in range(count):
                item = LVITEMW(mask=0x1, iItem=i, pszText=text_addr,
                               cchTextMax=260)
                kernel32.WriteProcessMemory(h, remote, ctypes.byref(item),
                                            ctypes.sizeof(item), None)
                user32.SendMessageW(lv, 0x1073, i, remote)  # GETITEMTEXTW
                buf = ctypes.create_unicode_buffer(260)
                kernel32.ReadProcessMemory(h, text_addr, buf, 520, None)
                name = buf.value
                if not any(k in name.lower() for k in keywords):
                    continue
                r = RECT(1, 0, 0, 0)  # LVIR_ICON
                kernel32.WriteProcessMemory(h, rect_addr, ctypes.byref(r),
                                            ctypes.sizeof(r), None)
                if not user32.SendMessageW(lv, 0x100E, i, rect_addr):
                    continue
                kernel32.ReadProcessMemory(h, rect_addr, ctypes.byref(r),
                                           ctypes.sizeof(r), None)
                pt = POINT((r.left + r.right) // 2, (r.top + r.bottom) // 2)
                user32.ClientToScreen(lv, ctypes.byref(pt))
                results.append((name, (pt.x, pt.y)))
        finally:
            kernel32.VirtualFreeEx(h, remote, 0, 0x8000)
            kernel32.CloseHandle(h)
    except Exception as error:
        log("icon_search_error", repr(error))
    return results


# ============================================================
# FINDING PROGRAMS ON DISK  (only a fallback)
# ============================================================

def _app_path_from_registry(exe_name):
    keys = [
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths" + "\\" + exe_name,
        r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths"
        + "\\" + exe_name,
    ]
    for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        for key in keys:
            try:
                with winreg.OpenKey(root, key) as k:
                    value = winreg.QueryValueEx(k, "")[0].strip().strip('"')
                    if os.path.isfile(value):
                        return value
            except OSError:
                pass
    return None


def _program_dirs():
    dirs = []
    for var in ("ProgramFiles", "ProgramFiles(x86)", "ProgramW6432",
                "LOCALAPPDATA"):
        d = os.environ.get(var)
        if d and d not in dirs:
            dirs.append(d)
    for d in (r"C:\Program Files", r"C:\Program Files (x86)"):
        if d not in dirs:
            dirs.append(d)
    return dirs


def find_soffice():
    path = _app_path_from_registry("soffice.exe")
    if path:
        return path
    for base in _program_dirs():
        for p in glob.glob(os.path.join(base, "*OpenOffice*", "program",
                                        "soffice.exe")):
            return p
    return None


def find_chrome():
    path = _app_path_from_registry("chrome.exe")
    if path:
        return path
    for base in _program_dirs():
        p = os.path.join(base, "Google", "Chrome", "Application", "chrome.exe")
        if os.path.isfile(p):
            return p
    return None


# ============================================================
# TYPING  (all delays x SLOWDOWN)
# ============================================================

KEYBOARD_ROWS = ["qwertyuiop", "asdfghjkl", "zxcvbnm"]


def neighbor(ch):
    for row in KEYBOARD_ROWS:
        i = row.find(ch.lower())
        if i != -1:
            opts = [row[j] for j in (i - 1, i + 1) if 0 <= j < len(row)]
            wrong = random.choice(opts)
            return wrong.upper() if ch.isupper() else wrong
    return ch


def human_type(text, wpm=(58, 82), typo_rate=0.0, variable_pace=False):
    current_wpm = random.uniform(*wpm)
    change_after = random.randint(12, 35)
    since_change = 0

    for ch in text:
        if variable_pace and since_change >= change_after:
            r = random.random()
            if r < 0.25:
                current_wpm = random.uniform(48, 60)
            elif r < 0.75:
                current_wpm = random.uniform(62, 78)
            else:
                current_wpm = random.uniform(80, 95)
            since_change = 0
            change_after = random.randint(12, 40)
        since_change += 1

        if ch == "\n":
            ui.press("enter")
            pause(0.18, 0.45)
            continue

        if typo_rate and ch.isalpha() and random.random() < typo_rate:
            wrong = neighbor(ch)
            if wrong != ch:
                ui.write(wrong)
                if random.random() < 0.7:
                    pause(0.06, 0.16)
                else:
                    pause(0.18, 0.38)
                ui.press("backspace")
                pause(0.035, 0.10)

        ui.write(ch)

        base = 60.0 / (current_wpm * 5.0)
        delay = random.lognormvariate(math.log(base * 0.72), 0.20)
        if ch == " ":
            delay *= random.uniform(0.65, 1.05)
        elif ch in ",;:":
            delay += random.uniform(0.035, 0.11)
        elif ch in ".!?":
            delay += random.uniform(0.08, 0.22)
            if variable_pace and random.random() < 0.18:
                delay += random.uniform(0.20, 0.55)
        time.sleep(min(delay, 0.65) * TYPING_SLOWDOWN)


def type_calc_text(text):
    human_type(text, wpm=(40, 52), typo_rate=0.008)


def type_calc_number(text):
    human_type(str(text), wpm=(42, 55), typo_rate=0)


def type_writer_text(text, typo_rate=0.010):
    human_type(text, wpm=(58, 82), typo_rate=typo_rate, variable_pace=True)


def type_name(text):
    """File names: normal human typing, occasional fixed typo."""
    human_type(text, wpm=(45, 60), typo_rate=0.02)


def type_url(text):
    human_type(text, wpm=(50, 65), typo_rate=0)


# ============================================================
# MOUSE
# ============================================================

def move_to(x, y):
    x0, y0 = ui.position()
    dist = math.hypot(x - x0, y - y0)
    steps = max(15, int(dist / random.uniform(14, 22)))
    total = (random.uniform(0.25, 0.45) + dist / 3000) * min(SLOWDOWN, 2.0)

    bend = random.uniform(-0.12, 0.12) * dist
    nx, ny = -(y - y0), (x - x0)
    norm = math.hypot(nx, ny) or 1.0
    nx, ny = nx / norm, ny / norm

    for i in range(1, steps + 1):
        t = i / steps
        s = t * t * (3 - 2 * t)
        arc = math.sin(math.pi * t) * bend
        ui.move(x0 + (x - x0) * s + nx * arc, y0 + (y - y0) * s + ny * arc)
        time.sleep(total / steps)
    ui.move(x, y)


def click_at(x, y):
    move_to(x, y)
    pause(0.12, 0.28)
    ui.click()


def double_click_at(x, y):
    move_to(x, y)
    pause(0.12, 0.28)
    ui.double_click()


def double_click_desktop_icon(keyword, label):
    """
    Find the desktop icon, move the cursor onto it, CHECK that the
    thing under the cursor really is that icon, then double-click.
    Returns True only if the verified double-click happened.
    """
    show_desktop()
    icon = find_desktop_icon([keyword])
    w, h = ui.size()

    if not icon:
        log("icon_not_found", f"{label}; screen={w}x{h}")
        return False

    name, (x, y) = icon
    if not (0 <= x < w and 0 <= y < h):
        log("icon_off_screen", f"{label} '{name}' at ({x},{y}); screen={w}x{h}")
        return False

    move_to(x, y)
    pause(0.3, 0.5)
    under = accessible_at(x, y)
    under_name = under[0] if under else ""
    log("icon_check",
        f"{label}: icon '{name}' at ({x},{y}); under cursor '{under_name}'; "
        f"screen={w}x{h}")

    if keyword not in under_name.lower():
        print(f"Icon check failed for {label} - not clicking.")
        return False

    print(f"Double-clicking desktop icon '{name}'")
    pause(0.12, 0.28)
    ui.double_click()
    log("double_click_icon", name)
    return True


def close_with_cursor(hwnd, label):
    """Move the cursor to the window's own X button and click it."""
    print(f"Closing {label} with cursor...")
    if not window_alive(hwnd):
        return
    focus_window(hwnd)
    s = dpi_scale(hwnd)
    left, top, right, _ = visible_rect(hwnd)
    x = right - int(23 * s) + random.randint(-4, 4)
    y = top + int(11 * s) + random.randint(-2, 2)
    exe = window_exe(hwnd)
    click_at(x, y)
    if wait_until(lambda: not window_alive(hwnd), timeout=10):
        log("close_window", f"{label}: clicked X")
        pause(1.0, 2.0)
    else:
        log("close_failed", f"{label}: still open", hwnd)
        print(f"WARNING: {label} did not close ({exe})")


# ============================================================
# OPENOFFICE
# ============================================================

def is_start_center(hwnd, title):
    t = title.strip()
    return t in ("OpenOffice", "Apache OpenOffice") or t.endswith("Start Center")


def open_openoffice(kind):
    """kind = 'calc' or 'writer'. Returns the new document window."""
    if kind == "calc":
        title_key, switch, button_words = "OpenOffice Calc", "-calc", ("spreadsheet",)
    else:
        title_key, switch, button_words = "OpenOffice Writer", "-writer", ("text document",)

    print(f"Opening {title_key}...")
    existing = {h for h, _ in top_windows()}

    def is_new_doc(hwnd, title):
        return title_key in title and hwnd not in existing

    hwnd = None

    # 1) Mouse: desktop icon -> Start Center -> button (only if verified)
    if double_click_desktop_icon("openoffice", "OpenOffice"):
        found = wait_until(
            lambda: find_window(is_new_doc) or find_window(is_start_center),
            timeout=60,
        )
        if found and is_new_doc(found, window_title(found)):
            hwnd = found
        elif found:
            pause(1.0, 1.5)
            hwnd = click_start_center_button(found, button_words, is_new_doc)
            if not hwnd:
                # couldn't use the Start Center: close it cleanly first
                close_with_cursor(found, "OpenOffice Start Center")
        else:
            log("no_window_after_icon", "OpenOffice")

    # 2) Fallback: start the program directly (the method that works)
    if not hwnd:
        exe = find_soffice()
        if not exe:
            raise RuntimeError("OpenOffice not found")
        print("Fallback: starting", exe, switch)
        subprocess.Popen([exe, switch, "-nofirststartwizard", "-norestore"])
        hwnd = wait_until(lambda: find_window(is_new_doc), timeout=60)
        log("fallback_launch", f"{os.path.basename(exe)} {switch}", hwnd)

    if not hwnd:
        raise RuntimeError(f"{title_key} did not open")

    log("window_open", window_title(hwnd), hwnd)
    pause(1.5, 2.5)
    maximize_window(hwnd, title_key)
    return hwnd


gdi32 = ctypes.WinDLL("gdi32")
user32.GetDC.argtypes = [HWND]
user32.GetDC.restype = wintypes.HDC
user32.ReleaseDC.argtypes = [HWND, wintypes.HDC]
gdi32.GetPixel.argtypes = [wintypes.HDC, ctypes.c_int, ctypes.c_int]
gdi32.GetPixel.restype = DWORD

# Button label positions inside the white Start Center panel,
# measured from a real Start Center, as fractions of the panel WIDTH
# (so they stay correct at any screen size / scaling):
#   (x from, x to, y centre)  measured from the panel's top-left corner
START_CENTER_LABELS = {
    "text document": (0.23, 0.33, 0.218),
    "spreadsheet":   (0.23, 0.31, 0.290),
}


def _brightness(hdc, x, y):
    c = gdi32.GetPixel(hdc, int(x), int(y))
    if c == 0xFFFFFFFF:
        return None
    return ((c & 0xFF) + ((c >> 8) & 0xFF) + ((c >> 16) & 0xFF)) / 3


def find_start_center_panel(start_center):
    """
    The Start Center is a white panel centred on a grey background.
    Returns (panel_left, panel_top, panel_width) by reading pixels.
    """
    left, top, right, bottom = visible_rect(start_center)
    s = dpi_scale(start_center)
    hdc = user32.GetDC(None)
    try:
        # top of panel: walking down the grey left margin, the light
        # menu bar ends and the grey begins
        xg = left + int(12 * s)
        panel_top = None
        for y in range(top + int(30 * s), min(bottom, top + int(200 * s))):
            b = _brightness(hdc, xg, y)
            if b is not None and b < 215:
                panel_top = y
                break
        if panel_top is None:
            return None

        row = panel_top + int(40 * s)
        mid = (left + right) // 2

        panel_left = None
        for x in range(left + int(4 * s), mid):
            b = _brightness(hdc, x, row)
            if b is not None and b > 240:
                panel_left = x
                break

        panel_right = None
        for x in range(right - int(4 * s), mid, -1):
            b = _brightness(hdc, x, row)
            if b is not None and b > 240:
                panel_right = x
                break

        if panel_left is None or panel_right is None:
            return None
        width = panel_right - panel_left
        if width < (right - left) * 0.4:
            return None
        return panel_left, panel_top, width
    finally:
        user32.ReleaseDC(None, hdc)


def click_start_center_button(start_center, words, is_new_doc):
    """Click 'Spreadsheet' or 'Text Document' in the Start Center."""
    focus_window(start_center)
    pause(0.5, 0.8)

    panel = find_start_center_panel(start_center)
    if not panel:
        log("start_center_button", "panel not found on screen", start_center)
        return None

    pl, pt, pw = panel
    x_from, x_to, y_frac = START_CENTER_LABELS[words[0]]
    x = pl + int(pw * random.uniform(x_from, x_to))
    y = pt + int(pw * y_frac) + random.randint(-4, 4)

    print(f"Clicking '{words[0]}' in the Start Center")
    click_at(x, y)
    log("click_button",
        f"Start Center '{words[0]}' at ({x},{y}); panel=({pl},{pt}) w={pw}",
        start_center)
    hwnd = wait_until(lambda: find_window(is_new_doc), timeout=30)
    if not hwnd:
        log("start_center_button", "clicked but no document opened")
    return hwnd


def close_start_center_if_open():
    hwnd = find_window(is_start_center)
    if hwnd:
        close_with_cursor(hwnd, "OpenOffice Start Center")


def pick_file_name(templates, ext):
    for _ in range(50):
        name = random.choice(templates).format(n=random.randint(1, 99))
        if not os.path.exists(os.path.join(DOCUMENTS, name + ext)):
            return name
    return random.choice(templates).format(n=random.randint(100, 999))


def save_document(hwnd, templates, ext):
    name = pick_file_name(templates, ext)
    print("Saving as:", name)
    focus_window(hwnd)

    dialog = None
    for _ in range(3):
        ui.hotkey("ctrl", "s")
        dialog = wait_until(lambda: "Save" in foreground_title(), timeout=8)
        if dialog:
            break
        focus_window(hwnd)
    if not dialog:
        raise RuntimeError("Save dialog never appeared")

    pause(0.8, 1.4)
    type_name(name)              # the name box is already selected
    pause(0.4, 0.8)
    ui.press("enter")

    def saved():
        t = foreground_title()
        if "Confirm" in t or "already exists" in t:
            ui.hotkey("alt", "y")
        return name in window_title(hwnd)

    if not wait_until(saved, timeout=20, interval=0.5):
        log("save_failed", name, hwnd)
        raise RuntimeError(f"Could not confirm save of '{name}'")

    log("save", f"{name}{ext}", hwnd)
    print("Saved:", window_title(hwnd))
    pause(1.0, 1.8)


# ============================================================
# CALC
# ============================================================

def calc_session():
    banner("CALC SESSION")
    log("session_start", "calc")

    hwnd = open_openoffice("calc")

    ui.hotkey("ctrl", "home")
    pause(0.5, 0.8)

    headers = ("Item", "Qty", "Unit Price", "Total")
    for i, h in enumerate(headers):
        type_calc_text(h)
        ui.press("tab" if i < len(headers) - 1 else "enter")
        pause(0.20, 0.45)
    ui.press("home")
    pause(0.4, 0.7)

    items = random.sample(ITEMS, 6)
    first, last = 2, 1 + len(items)

    for row, item in enumerate(items, start=first):
        type_calc_text(item)
        ui.press("tab")
        pause(0.20, 0.45)

        type_calc_number(random.randint(1, 20))
        ui.press("tab")
        pause(0.20, 0.45)

        type_calc_number(f"{random.uniform(20, 450):.2f}")
        ui.press("tab")
        pause(0.20, 0.45)

        type_calc_number(f"=B{row}*C{row}")
        ui.press("enter")
        ui.press("home")
        pause(0.45, 0.85)

    total_row = last + 1
    type_calc_text("Total")
    for _ in range(3):
        ui.press("tab")
        pause(0.1, 0.2)
    type_calc_number(f"=SUM(D{first}:D{last})")
    ui.press("enter")
    ui.press("home")
    pause(0.6, 1.0)

    # bold header
    ui.hotkey("ctrl", "home")
    pause(0.3, 0.5)
    ui.hotkey("ctrl", "shift", "right")
    pause(0.2, 0.4)
    ui.hotkey("ctrl", "b")
    pause(0.5, 0.8)

    # 2 decimal places on C2:D<total>
    ui.hotkey("ctrl", "home")
    ui.press("down")
    ui.press("right")
    ui.press("right")
    pause(0.3, 0.5)
    for _ in range(total_row - first):
        ui.hotkey("shift", "down")
        pause(0.05, 0.12)
    ui.hotkey("shift", "right")
    pause(0.2, 0.4)
    ui.hotkey("ctrl", "shift", "1")
    pause(0.5, 0.8)

    # bold total row
    ui.hotkey("ctrl", "home")
    for _ in range(total_row - 1):
        ui.press("down")
        pause(0.04, 0.09)
    for _ in range(3):
        ui.hotkey("shift", "right")
        pause(0.05, 0.12)
    ui.hotkey("ctrl", "b")
    pause(0.6, 1.0)

    # look over the sheet with the mouse wheel
    move_to(*point_inside(hwnd))
    ui.scroll(-120)
    pause(0.8, 1.2)
    ui.scroll(120)
    pause(0.6, 1.0)
    ui.hotkey("ctrl", "home")
    pause(0.7, 1.1)

    save_document(hwnd, CALC_NAMES, ".ods")
    close_with_cursor(hwnd, "Calc")
    close_start_center_if_open()
    log("session_end", "calc")
    print("Calc completed.")


# ============================================================
# WRITER
# ============================================================

def write_heading(text, level):
    ui.hotkey("ctrl", str(level))       # Heading 1 / Heading 2
    pause(0.2, 0.4)
    if level == 1:
        ui.hotkey("ctrl", "e")
        pause(0.2, 0.4)
    type_writer_text(text, typo_rate=0.006)
    ui.press("enter")                   # next paragraph = Text Body
    pause(0.3, 0.5)
    if level == 1:
        ui.hotkey("ctrl", "l")
        pause(0.2, 0.4)


def writer_session():
    banner("WRITER SESSION")
    log("session_start", "writer")

    hwnd = open_openoffice("writer")

    write_heading(REPORT_TITLE, 1)
    type_writer_text("Date: " + datetime.date.today().strftime("%d %B %Y"),
                     typo_rate=0)
    ui.press("enter")
    pause(0.4, 0.7)

    write_heading("Discussion Summary", 2)
    for paragraph in MEETING_PARAGRAPHS:
        type_writer_text(paragraph)
        ui.press("enter")
        pause(0.55, 1.20)

    write_heading("Action Items", 2)
    ui.hotkey("shift", "f12")           # bullets on
    pause(0.25, 0.45)
    for i, item in enumerate(ACTION_ITEMS):
        type_writer_text(item, typo_rate=0.008)
        if i < len(ACTION_ITEMS) - 1:
            ui.press("enter")
            pause(0.25, 0.60)
    pause(0.7, 1.2)

    # read back through it with the mouse wheel
    move_to(*point_inside(hwnd))
    for _ in range(random.randint(3, 5)):
        ui.scroll(random.choice((240, 360)))
        pause(0.6, 1.1)
    for _ in range(random.randint(3, 5)):
        ui.scroll(-random.choice((120, 240)))
        pause(0.8, 1.4)

    save_document(hwnd, WRITER_NAMES, ".odt")
    close_with_cursor(hwnd, "Writer")
    close_start_center_if_open()
    log("session_end", "writer")
    print("Writer completed.")


# ============================================================
# BROWSER
# ============================================================

def is_browser(hwnd, title):
    return "Google Chrome" in title


def page_title(window_text):
    return window_text.rsplit(" - ", 1)[0].strip()


def wait_page_settled(hwnd, old_title, timeout=20):
    end = time.perf_counter() + timeout
    last, since = None, None
    while time.perf_counter() < end:
        title = window_title(hwnd)
        if title != old_title and page_title(title) not in ("", "New Tab"):
            now = time.perf_counter()
            if title != last:
                last, since = title, now
            elif now - since >= 1.5:
                return True
        time.sleep(0.25)
    return False


def _rows(hwnd, offsets):
    _, top, _, _ = visible_rect(hwnd)
    s = dpi_scale(hwnd)
    return [top + int(o * s) for o in offsets]


def tab_strip_rows(hwnd):
    return _rows(hwnd, (20, 16, 24, 12, 28))


def toolbar_rows(hwnd):
    return _rows(hwnd, (58, 54, 62, 50, 66))


def chrome_find(hwnd, rows, match, x_to_fraction=1.0):
    left, _, right, _ = visible_rect(hwnd)
    s = dpi_scale(hwnd)
    x_to = left + (right - left) * x_to_fraction - int(150 * s)
    return find_on_row(rows, left + int(4 * s), x_to, int(6 * s), match)


def click_address_bar(hwnd):
    runs = chrome_find(
        hwnd, toolbar_rows(hwnd),
        lambda n, r, st: r == ROLE_TEXT or "address" in n.lower(),
    )
    if runs:
        x, y = runs[0]
        click_at(x + random.randint(-40, 40), y)
        pause(0.2, 0.4)
        ui.hotkey("ctrl", "a")   # select whatever is in the box
        log("click_address_bar", "mouse", hwnd)
    else:
        ui.hotkey("ctrl", "l")
        log("click_address_bar", "fallback Ctrl+L", hwnd)
    pause(0.3, 0.6)


def click_new_tab(hwnd):
    runs = chrome_find(
        hwnd, tab_strip_rows(hwnd),
        lambda n, r, st: r == ROLE_BUTTON and n.strip().lower() == "new tab",
    )
    if runs:
        click_at(*runs[-1])
        if wait_until(lambda: page_title(window_title(hwnd)) == "New Tab",
                      timeout=5):
            log("new_tab", "clicked + button", hwnd)
            pause(0.5, 1.0)
            return
    ui.hotkey("ctrl", "t")
    wait_until(lambda: page_title(window_title(hwnd)) == "New Tab", timeout=5)
    log("new_tab", "fallback Ctrl+T", hwnd)
    pause(0.5, 1.0)


def close_current_tab(hwnd):
    """Click the X on the right-most (newest, active) tab."""
    before = window_title(hwnd)
    runs = chrome_find(
        hwnd, tab_strip_rows(hwnd),
        lambda n, r, st: r == ROLE_BUTTON and n.strip().lower() == "close",
    )
    if runs:
        click_at(*runs[-1])
        if wait_until(lambda: window_title(hwnd) != before, timeout=5):
            log("close_tab", "clicked tab X", hwnd)
            pause(0.8, 1.4)
            return
    ui.hotkey("ctrl", "w")
    log("close_tab", "fallback Ctrl+W", hwnd)
    pause(0.8, 1.4)


def back_button(hwnd):
    """Returns (x, y, enabled) of Chrome's Back button, or None."""
    runs = chrome_find(
        hwnd, toolbar_rows(hwnd),
        lambda n, r, st: r == ROLE_BUTTON and n.strip().lower() == "back",
        x_to_fraction=0.3,
    )
    if not runs:
        return None
    x, y = runs[0]
    info = accessible_at(x, y)
    enabled = bool(info) and not (info[2] & STATE_UNAVAILABLE)
    return x, y, enabled


def go_back(hwnd):
    before = window_title(hwnd)
    btn = back_button(hwnd)
    if btn and btn[2]:
        click_at(btn[0], btn[1])
        log("back", "clicked Back button", hwnd)
    else:
        ui.hotkey("alt", "left")
        log("back", "fallback Alt+Left", hwnd)
    wait_page_settled(hwnd, before, timeout=10)
    pause(1.5, 2.5)


def handle_cert_warning(hwnd):
    title = window_title(hwnd)
    if not page_title(title).startswith("Privacy error"):
        return True
    print("Certificate warning page - continuing past it.")
    log("cert_warning", title, hwnd)
    click_at(*point_inside(hwnd, fx=(0.05, 0.15), fy=(0.75, 0.85)))
    pause(0.4, 0.8)
    type_url("thisisunsafe")
    wait_page_settled(hwnd, title, timeout=20)
    return not page_title(window_title(hwnd)).startswith("Privacy error")


def open_chrome():
    print("Opening Chrome...")
    existing = {h for h, _ in top_windows()}

    def is_new(hwnd, title):
        return is_browser(hwnd, title) and hwnd not in existing

    hwnd = None
    if double_click_desktop_icon("chrome", "Chrome"):
        hwnd = wait_until(lambda: find_window(is_new), timeout=40)
        if not hwnd:
            log("no_window_after_icon", "Chrome")

    if not hwnd:
        exe = find_chrome()
        if not exe:
            raise RuntimeError("Chrome not found")
        print("Fallback: starting", exe)
        subprocess.Popen([exe, "--new-window", "--no-first-run",
                          "--no-default-browser-check",
                          "--hide-crash-restore-bubble"])
        hwnd = wait_until(lambda: find_window(is_new), timeout=40)
        log("fallback_launch", os.path.basename(exe), hwnd)

    if not hwnd:
        raise RuntimeError("Chrome did not open")

    log("window_open", window_title(hwnd), hwnd)
    pause(1.5, 2.5)
    maximize_window(hwnd, "Chrome")
    return hwnd


def read_page(hwnd, seconds):
    """Slow human reading: wheel scrolling with long pauses."""
    end = time.time() + seconds
    move_to(*point_inside(hwnd))
    while time.time() < end:
        r = random.random()
        if r < 0.70:
            ui.scroll(-random.choice((120, 120, 240, 360)))
        elif r < 0.85:
            ui.scroll(random.choice((120, 240)))
        else:
            move_to(*point_inside(hwnd, fx=(0.2, 0.8), fy=(0.25, 0.85)))
        time.sleep(min(random.uniform(2.5, 8.0), max(0, end - time.time())))


def click_a_link(hwnd):
    """Hover around the page until the cursor is over a link, click it."""
    for _ in range(40):
        x, y = point_inside(hwnd, fx=(0.15, 0.85), fy=(0.30, 0.85))
        info = accessible_at(x, y)
        if info and info[1] == ROLE_LINK:
            click_at(x, y)
            return info[0]
    x, y = point_inside(hwnd)
    click_at(x, y)
    return None


def browse_site(hwnd, url, number, total, deadline):
    print()
    print(f"Website {number}/{total}: {url}")
    log("website", url, hwnd)
    focus_window(hwnd)

    old_title = window_title(hwnd)
    click_address_bar(hwnd)
    type_url(url)
    pause(0.20, 0.40)
    ui.press("enter")

    if not wait_page_settled(hwnd, old_title, timeout=30):
        log("load_timeout", url, hwnd)

    if not handle_cert_warning(hwnd):
        log("skipped", f"{url}: certificate page", hwnd)
        return

    print("Loaded:", page_title(window_title(hwnd)))
    log("page_loaded", page_title(window_title(hwnd)), hwnd)
    pause(1.5, 3.0)

    # normal reading, plus extra time if the run is ahead of schedule
    sites_left = total - number + 1
    reserve = 60 + 45 * sites_left
    extra = max(0, (deadline - time.time() - reserve) / sites_left)
    read_time = random.uniform(25, 45) + min(extra, 8 * 60)
    print(f"Reading for about {read_time / 60:.1f} min...")
    read_page(hwnd, read_time)

    before = window_title(hwnd)
    print("Clicking a link...")
    link = click_a_link(hwnd)
    log("page_click", link or "page content", hwnd)

    if wait_page_settled(hwnd, before, timeout=8):
        log("page_loaded", page_title(window_title(hwnd)), hwnd)
        read_page(hwnd, random.uniform(10, 25))
        btn = back_button(hwnd)
        if btn and not btn[2]:
            # link opened a new tab (nothing to go back to) -> close it
            close_current_tab(hwnd)
        else:
            go_back(hwnd)
    else:
        pause(2.0, 4.0)


def browser_session(deadline):
    banner("BROWSER SESSION")
    log("session_start", "browser")

    selected = random.sample(SITES, min(SITES_PER_RUN, len(SITES)))
    print("Selected websites:")
    for s in selected:
        print(" -", s)

    hwnd = open_chrome()

    for n, url in enumerate(selected, start=1):
        try:
            if not window_alive(hwnd):
                hwnd = open_chrome()
            if n > 1:
                click_new_tab(hwnd)
            browse_site(hwnd, url, n, len(selected), deadline)
            if 1 < n < len(selected):
                close_current_tab(hwnd)
        except Exception as error:
            print("Website error:", repr(error))
            log("website_error", f"{url}: {error!r}")
            ui.release_all()
        pause(0.7, 1.2)

    # if still early, keep reading the last page
    remaining = deadline - time.time() - 20
    if remaining > 0 and window_alive(hwnd):
        log("extra_reading", f"{remaining:.0f}s", hwnd)
        read_page(hwnd, remaining)

    close_with_cursor(hwnd, "Chrome")
    log("session_end", "browser")
    print("Browser completed.")


# ============================================================
# MAIN  (the only entry point - just press Run)
# ============================================================

def main():
    global RUN_START

    start_esc_watcher()
    ole32.CoInitialize(None)

    print()
    print("======================================")
    print(" OFFICE GUI AUTOMATION")
    print("======================================")
    print("Calc -> Writer -> websites")
    print(f"Minimum run time: {MIN_RUNTIME_MINUTES} min")
    print(f"Log: {LOG_FILE}")
    print()
    print("Press ESC at any time to kill the program.")
    print(f"Starting in {START_DELAY} seconds...")
    time.sleep(START_DELAY)

    minimize_console()

    RUN_START = time.time()
    w, h = ui.size()
    log("run_start", f"exe={os.path.basename(sys.executable)} "
                     f"frozen={bool(getattr(sys, 'frozen', False))} "
                     f"python={ctypes.sizeof(ctypes.c_void_p) * 8}bit "
                     f"screen={w}x{h} soffice={find_soffice()}")
    failures = 0

    try:
        for loop in range(1, LOOPS + 1):
            if LOOPS > 1:
                banner(f"LOOP {loop}/{LOOPS}")
            loop_start = time.time()
            deadline = loop_start + MIN_RUNTIME_MINUTES * 60

            for session in (calc_session, writer_session,
                            lambda: browser_session(deadline)):
                label = getattr(session, "__name__", "browser_session")
                if label == "<lambda>":
                    label = "browser_session"
                try:
                    session()
                except Exception as error:
                    failures += 1
                    print(f"\nERROR in {label}: {error!r}")
                    log("error", f"{label}: {error!r}")
                    ui.release_all()
                pause(1.5, 2.5)

            log("loop_end", f"loop {loop}: "
                            f"{(time.time() - loop_start) / 60:.1f} min")
    finally:
        ui.release_all()

    elapsed = time.time() - RUN_START
    log("run_end", f"{elapsed / 60:.1f} min, failures={failures}")

    print()
    print("======================================")
    print("Workflow completed."
          + (f"  ({failures} part(s) failed - see log)" if failures else ""))
    print(f"Runtime: {elapsed / 60:.1f} minutes")
    print("======================================")


if __name__ == "__main__":
    main()
