"""
APM - AdsPower Window Manager v2.1
Rebuilt from AutoIt v5.5 source to Python.
Manages AdsPower (SunBrowser) browser windows.
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import json
import time
import os
import sys
import math
import configparser
import re
import struct
import ctypes
import ctypes.wintypes
from datetime import datetime
from io import BytesIO
from urllib.request import Request, urlopen
from urllib.parse import urlencode
import subprocess

try:
    import requests
except ImportError:
    requests = None

try:
    import websocket as _ws_mod
    HAS_WS = True
except ImportError:
    HAS_WS = False

# Windows API with proper 64-bit type declarations
try:
    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32
    psapi = ctypes.windll.psapi

    HWND = ctypes.wintypes.HWND      # pointer-sized (8 bytes on x64)
    DWORD = ctypes.wintypes.DWORD
    BOOL = ctypes.wintypes.BOOL
    UINT = ctypes.wintypes.UINT
    LPARAM = ctypes.wintypes.LPARAM   # pointer-sized
    HANDLE = ctypes.wintypes.HANDLE

    # Set argtypes for all user32 functions we use
    user32.EnumWindows.argtypes = [ctypes.c_void_p, LPARAM]
    user32.EnumWindows.restype = BOOL
    user32.IsWindowVisible.argtypes = [HWND]
    user32.IsWindowVisible.restype = BOOL
    user32.IsIconic.argtypes = [HWND]
    user32.IsIconic.restype = BOOL
    user32.GetClassNameW.argtypes = [HWND, ctypes.c_wchar_p, ctypes.c_int]
    user32.GetClassNameW.restype = ctypes.c_int
    user32.GetWindowTextW.argtypes = [HWND, ctypes.c_wchar_p, ctypes.c_int]
    user32.GetWindowTextW.restype = ctypes.c_int
    user32.GetWindowThreadProcessId.argtypes = [HWND, ctypes.POINTER(DWORD)]
    user32.GetWindowThreadProcessId.restype = DWORD
    user32.ShowWindow.argtypes = [HWND, ctypes.c_int]
    user32.ShowWindow.restype = BOOL
    user32.SetForegroundWindow.argtypes = [HWND]
    user32.SetForegroundWindow.restype = BOOL
    user32.SetWindowPos.argtypes = [HWND, HWND, ctypes.c_int, ctypes.c_int,
                                     ctypes.c_int, ctypes.c_int, UINT]
    user32.SetWindowPos.restype = BOOL
    user32.PostMessageW.argtypes = [HWND, UINT, ctypes.wintypes.WPARAM, LPARAM]
    user32.PostMessageW.restype = BOOL
    user32.GetWindowRect.argtypes = [HWND, ctypes.POINTER(ctypes.wintypes.RECT)]
    user32.GetWindowRect.restype = BOOL
    user32.OpenClipboard.argtypes = [HWND]
    user32.OpenClipboard.restype = BOOL
    user32.EmptyClipboard.restype = BOOL
    user32.SetClipboardData.argtypes = [UINT, HANDLE]
    user32.CloseClipboard.restype = BOOL
    user32.GetSystemMetrics.argtypes = [ctypes.c_int]
    user32.GetSystemMetrics.restype = ctypes.c_int

    kernel32.OpenProcess.argtypes = [DWORD, BOOL, DWORD]
    kernel32.OpenProcess.restype = HANDLE
    kernel32.CloseHandle.argtypes = [HANDLE]
    kernel32.CloseHandle.restype = BOOL
    kernel32.GlobalAlloc.argtypes = [UINT, ctypes.c_size_t]
    kernel32.GlobalAlloc.restype = HANDLE
    kernel32.GlobalLock.argtypes = [HANDLE]
    kernel32.GlobalLock.restype = ctypes.c_void_p
    kernel32.GlobalUnlock.argtypes = [HANDLE]
    kernel32.GlobalUnlock.restype = BOOL
    kernel32.GetCurrentThreadId.argtypes = []
    kernel32.GetCurrentThreadId.restype = DWORD

    user32.AttachThreadInput.argtypes = [DWORD, DWORD, BOOL]
    user32.AttachThreadInput.restype = BOOL
    user32.BringWindowToTop.argtypes = [HWND]
    user32.BringWindowToTop.restype = BOOL

    HAS_WIN32 = True
except Exception:
    HAS_WIN32 = False

try:
    from PIL import Image, ImageDraw, ImageFont, ImageGrab
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

# ─── Constants ────────────────────────────────────────────────────────────────

VERSION = "4.9"
WINDOW_TITLE = f"AdsPower Window Manager v{VERSION} - Dev ChingChing"
CHROME_CLASS = "Chrome_WidgetWin_1"

# Win32 constants
SW_SHOWNORMAL = 1
SW_MINIMIZE = 6
SW_RESTORE = 9
SW_MAXIMIZE = 3
SW_SHOW = 5
SWP_SHOWWINDOW = 0x0040
SWP_FRAMECHANGED = 0x0020
SWP_NOACTIVATE = 0x0010
HWND_TOP = 0
GW_OWNER = 4
WM_CLOSE = 0x0010
PROCESS_QUERY_INFORMATION = 0x0400
PROCESS_VM_READ = 0x0010

# ─── Helpers ──────────────────────────────────────────────────────────────────

def get_base_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

BASE_DIR = get_base_dir()
DATA_DIR = os.path.join(BASE_DIR, 'APManagerData')
CONFIG_PATH = os.path.join(DATA_DIR, 'config.ini')

def ensure_dirs():
    os.makedirs(DATA_DIR, exist_ok=True)

def load_config():
    cfg = configparser.ConfigParser()
    if os.path.exists(CONFIG_PATH):
        cfg.read(CONFIG_PATH, encoding='utf-8')
    for s in ['MAIN', 'HOTKEYS', 'HOTKEYS2', 'DISCORD', 'DISTRIBTE', 'POSITIONER', 'ADSPOWER']:
        if not cfg.has_section(s):
            cfg.add_section(s)
    defaults = {
        'MAIN': {
            'GUIW': '374', 'GUIH': '680', 'GUIX': '920', 'GUIY': '180',
            'SortColumn': '0', 'AutoSorting': '0', 'InjectControls': '1',
            'AlwaysOnTop': '0', 'AllHotkeysON': '1',
            'HotkeysToggleExtra': 'CTRL+SHIFT+H',
            'MainURL': 'https://www.ticketmaster.com',
            'CustomNavSize': '1', 'NavWidth': '480', 'NavHeight': '540',
            'MinimizeOthers': '0',
        },
        'HOTKEYS': {
            'FORWARD': 'CTRL+SHIFT+RIGHT', 'BACKWARD': 'CTRL+SHIFT+LEFT',
            'TOP': 'CTRL+SHIFT+UP', 'SORTTAB': 'CTRL+SHIFT+T',
            'SORTPROFILE': 'CTRL+SHIFT+P', 'GROUPNEXT': '[', 'GROUPBACK': ']',
        },
        'DISCORD': {
            'QueWebhook': 'https://discord.com/api/webhooks/1464267517139877930/Ae0LDeglr3CEYK_vjsTd1htYoevub_ajXCcb4CAWVSGkg-s2XweTo9MIqNiZNNAH_iOQ',
            'ProdWebhook': 'https://discord.com/api/webhooks/1464267286918594652/tz1Go3i_cGsHz0f08bpAAR_C1wRhw6eU629CrajA4uDxf4kd5L-0ZKbxh6vLduCLibPo',
            'VfWebhook': 'https://discord.com/api/webhooks/1483812119135912059/hAPUWxiqRVNAw43gjl8L1rLt9M6emWWCpyPjIwS5K0JN6WhyiysyZdky5wGWt6iLpBBr',
            'ProfileName': '', 'ScreenshotFolder': os.path.join(BASE_DIR, 'Screenshots'),
            'SheetUrl': '',
        },
        'ADSPOWER': {'ApiKey': ''},
        'DISTRIBTE': {'Email': '', 'Password': ''},
        'POSITIONER': {
            'Cols': '4', 'Rows': '2', 'Width': '480', 'Height': '540',
            'GapX': '0', 'GapY': '0', 'URL': 'https://www.ticketmaster.com',
        },
    }
    for section, vals in defaults.items():
        for k, v in vals.items():
            if not cfg.has_option(section, k):
                cfg.set(section, k, v)
    return cfg

def save_config(cfg):
    ensure_dirs()
    with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
        cfg.write(f)


# ─── Win32 Browser Management ────────────────────────────────────────────────

def enum_windows():
    """Get all visible windows with Chrome_WidgetWin_* class.
    SunBrowser uses Chrome_WidgetWin_1 for main windows but
    Chrome_WidgetWin_2/3/4/etc for dragged-out tab windows."""
    if not HAS_WIN32:
        return []
    results = []
    WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.wintypes.BOOL, ctypes.wintypes.HWND, ctypes.wintypes.LPARAM)
    def callback(hwnd, _):
        try:
            if not user32.IsWindowVisible(hwnd):
                return True
            class_name = ctypes.create_unicode_buffer(256)
            user32.GetClassNameW(hwnd, class_name, 256)
            if class_name.value.startswith('Chrome_WidgetWin_'):
                title = ctypes.create_unicode_buffer(512)
                user32.GetWindowTextW(hwnd, title, 512)
                if title.value:
                    results.append((hwnd, title.value))
        except Exception:
            pass
        return True
    cb = WNDENUMPROC(callback)
    user32.EnumWindows(cb, 0)
    return results


def enum_all_windows_for_pids(target_pids):
    """Find ALL visible windows belonging to specific PIDs, any class.
    Used for diagnostics to detect windows with non-standard class names."""
    if not HAS_WIN32:
        return []
    results = []
    WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.wintypes.BOOL, ctypes.wintypes.HWND, ctypes.wintypes.LPARAM)
    def callback(hwnd, _):
        try:
            if not user32.IsWindowVisible(hwnd):
                return True
            pid = ctypes.wintypes.DWORD()
            user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            if pid.value in target_pids:
                class_name = ctypes.create_unicode_buffer(256)
                user32.GetClassNameW(hwnd, class_name, 256)
                title = ctypes.create_unicode_buffer(512)
                user32.GetWindowTextW(hwnd, title, 512)
                results.append((hwnd, title.value, class_name.value, pid.value))
        except Exception:
            pass
        return True
    cb = WNDENUMPROC(callback)
    user32.EnumWindows(cb, 0)
    return results


def get_window_pid(hwnd):
    """Get the process ID of a window."""
    pid = ctypes.wintypes.DWORD()
    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    return pid.value


def get_process_exe(pid):
    """Get the executable path for a process."""
    if not HAS_WIN32:
        return ''
    h = None
    try:
        h = kernel32.OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, False, pid)
        if not h:
            # Try with just PROCESS_QUERY_LIMITED_INFORMATION (0x1000)
            h = kernel32.OpenProcess(0x1000, False, pid)
        if not h:
            return ''
        buf = ctypes.create_unicode_buffer(1024)
        size = ctypes.wintypes.DWORD(1024)
        ret = kernel32.QueryFullProcessImageNameW(h, 0, buf, ctypes.byref(size))
        if ret:
            return buf.value
    except Exception:
        pass
    finally:
        if h:
            kernel32.CloseHandle(h)
    return ''


def get_process_cmdline(pid):
    """Get the command line of a process via WMI (wmic or PowerShell fallback)."""
    # Try wmic first
    try:
        result = subprocess.run(
            ['wmic', 'process', 'where', f'ProcessId={pid}', 'get', 'CommandLine', '/format:list'],
            capture_output=True, text=True, timeout=5, creationflags=0x08000000  # CREATE_NO_WINDOW
        )
        for line in result.stdout.splitlines():
            if line.startswith('CommandLine='):
                return line[12:]
    except Exception:
        pass
    # PowerShell fallback (wmic deprecated on newer Windows)
    try:
        ps_cmd = f'(Get-CimInstance Win32_Process -Filter "ProcessId={pid}").CommandLine'
        result = subprocess.run(
            ['powershell', '-NoProfile', '-Command', ps_cmd],
            capture_output=True, text=True, timeout=5, creationflags=0x08000000
        )
        cmdline = result.stdout.strip()
        if cmdline:
            return cmdline
    except Exception:
        pass
    return ''


def is_sunbrowser(pid):
    """Check if a PID is a SunBrowser (AdsPower browser) process."""
    exe = get_process_exe(pid)
    return 'SunBrowser' in exe or 'sun_browser' in exe.lower()


def get_adspower_userid_from_cmdline(cmdline):
    """Extract user_id from --user-data-dir path in command line.
    Matches AutoIt regex: user-data-dir="?[^"]*[/\\]([a-z0-9]+)[/\\]?"?
    The user_id is the last path segment (e.g. k1blafdl_h26ptei)."""
    # AutoIt pattern: last folder segment of user-data-dir path
    # IMPORTANT: [a-z0-9] without underscore - matches AutoIt original exactly
    m = re.search(r'user-data-dir="?[^"]*[/\\]([a-z0-9]+)[/\\]?"?', cmdline, re.IGNORECASE)
    if m:
        return m.group(1)
    # Also try --session_name
    m = re.search(r'--session[_-]name="?([^"\s]+)"?', cmdline)
    if m:
        return m.group(1).strip()
    return ''


def force_foreground(hwnd):
    """Force a window to foreground - mimics AutoIt WinActivate.
    Uses multiple tricks to bypass Windows foreground lock."""
    if not HAS_WIN32:
        return
    try:
        # Trick 1: Simulate Alt key release - this allows SetForegroundWindow to work
        # because Windows allows foreground changes during keyboard input
        user32.keybd_event(0x12, 0, 2, 0)  # Alt key up (KEYEVENTF_KEYUP=2)

        # Trick 2: Attach our thread to the FOREGROUND window's thread
        # This gives our thread foreground privileges
        fore = user32.GetForegroundWindow()
        if fore:
            fore_tid = user32.GetWindowThreadProcessId(fore, None)
            cur_tid = kernel32.GetCurrentThreadId()
            attached = False
            if fore_tid and fore_tid != cur_tid:
                attached = bool(user32.AttachThreadInput(cur_tid, fore_tid, True))

            user32.BringWindowToTop(hwnd)
            user32.SetForegroundWindow(hwnd)

            if attached:
                user32.AttachThreadInput(cur_tid, fore_tid, False)
        else:
            user32.BringWindowToTop(hwnd)
            user32.SetForegroundWindow(hwnd)
    except Exception:
        try:
            user32.SetForegroundWindow(hwnd)
        except Exception:
            pass


def show_window(hwnd):
    """Show and maximize a window (used for profile select/navigate)."""
    if HAS_WIN32:
        user32.ShowWindow(hwnd, SW_MAXIMIZE)
        force_foreground(hwnd)


def activate_window(hwnd):
    """Activate/bring to front without maximizing (used for URL sending, groups)."""
    if HAS_WIN32:
        force_foreground(hwnd)


def minimize_window(hwnd):
    if HAS_WIN32:
        user32.ShowWindow(hwnd, SW_MINIMIZE)


def close_window(hwnd):
    if HAS_WIN32:
        user32.PostMessageW(hwnd, WM_CLOSE, 0, 0)


def restore_and_resize(hwnd, w, h):
    """Restore window (un-maximize) and resize to w x h, positioned at left side (0,0)."""
    if HAS_WIN32:
        user32.ShowWindow(hwnd, SW_RESTORE)
        user32.SetWindowPos(hwnd, None, 0, 0, w, h, SWP_SHOWWINDOW | SWP_FRAMECHANGED)
        force_foreground(hwnd)


def set_window_pos(hwnd, x, y, w, h):
    if HAS_WIN32:
        user32.ShowWindow(hwnd, SW_RESTORE)
        user32.SetWindowPos(hwnd, HWND_TOP, x, y, w, h, SWP_SHOWWINDOW | SWP_FRAMECHANGED)
        user32.SetForegroundWindow(hwnd)


def get_window_title(hwnd):
    if not HAS_WIN32:
        return ''
    buf = ctypes.create_unicode_buffer(512)
    user32.GetWindowTextW(hwnd, buf, 512)
    return buf.value


def is_window_visible(hwnd):
    if not HAS_WIN32:
        return False
    return bool(user32.IsWindowVisible(hwnd))


def send_keys_to_window(hwnd, keys):
    """Activate window and send keystrokes (without maximizing)."""
    if not HAS_WIN32:
        return
    activate_window(hwnd)
    time.sleep(0.15)
    try:
        import keyboard as kb
        for key in keys:
            if key == '{F5}':
                kb.send('f5')
            elif key == '!l':
                kb.send('alt+l')
            elif key == '^t':
                kb.send('ctrl+t')
            elif key == '^l':
                kb.send('ctrl+l')
            elif key == '^v':
                kb.send('ctrl+v')
            elif key == '{ENTER}':
                kb.send('enter')
            else:
                kb.send(key)
            time.sleep(0.05)
    except ImportError:
        pass


def set_clipboard(text):
    """Set clipboard text using ctypes."""
    if not HAS_WIN32:
        return
    user32.OpenClipboard(0)
    user32.EmptyClipboard()
    # CF_UNICODETEXT = 13
    data = text.encode('utf-16-le') + b'\x00\x00'
    h = kernel32.GlobalAlloc(0x0042, len(data))
    p = kernel32.GlobalLock(h)
    ctypes.memmove(p, data, len(data))
    kernel32.GlobalUnlock(h)
    user32.SetClipboardData(13, h)
    user32.CloseClipboard()


def get_screen_size():
    if HAS_WIN32:
        return (user32.GetSystemMetrics(0), user32.GetSystemMetrics(1))
    return (1920, 1080)


# ─── AdsPower API ─────────────────────────────────────────────────────────────

class AdsPowerAPI:
    def __init__(self, port=50325, api_key=''):
        self.bases = [f'http://127.0.0.1:{port}', f'http://local.adspower.net:{port}']
        self.api_key = api_key
        self._cache = None
        self._cache_time = 0
        self._cache_ttl = 30  # seconds

    def _api_path(self, path):
        """Append api_key to path if configured."""
        if self.api_key:
            sep = '&' if '?' in path else '?'
            return f'{path}{sep}api_key={self.api_key}'
        return path

    def get_user_list(self, force=False, max_pages=20):
        """Get profiles from AdsPower user/list API.
        Limited to max_pages to avoid stalling on large accounts (70K+ profiles).
        Individual profiles are resolved via get_user_by_id as fallback."""
        now = time.time()
        if not force and self._cache and (now - self._cache_time) < self._cache_ttl:
            return self._cache

        all_users = []
        page = 1
        while page <= max_pages:
            data = self._get(self._api_path(f'/api/v1/user/list?page={page}&page_size=100'))
            if not data:
                break
            raw = data.get('data', {})
            if isinstance(raw, dict):
                users = raw.get('list', [])
            elif isinstance(raw, list):
                users = raw
            else:
                break
            if not users:
                break
            all_users.extend(users)
            page += 1
            if len(users) < 100:
                break

        # Don't wipe cache if API returned nothing (API might be temporarily down)
        if all_users:
            self._cache = all_users
            self._cache_time = now
        elif not self._cache:
            self._cache = []
            self._cache_time = now
        return self._cache if self._cache else []

    def get_user_by_id(self, user_id):
        """Get a single profile by user_id."""
        data = self._get(self._api_path(f'/api/v1/user/list?user_id={user_id}'))
        if not data:
            return None
        if data.get('code') != 0:
            return None
        raw = data.get('data', {})
        if isinstance(raw, dict):
            lst = raw.get('list', [])
            return lst[0] if lst else None
        elif isinstance(raw, list) and raw:
            return raw[0]
        return None

    def _get(self, path, timeout=4):
        for base in self.bases:
            # Try requests first
            if requests:
                try:
                    r = requests.get(base + path, timeout=timeout)
                    if r.ok:
                        return r.json()
                except Exception:
                    pass
            # Fallback to urllib
            else:
                try:
                    req = Request(base + path)
                    resp = urlopen(req, timeout=timeout)
                    return json.loads(resp.read().decode('utf-8'))
                except Exception:
                    pass
        return None

    def resolve_profile_name(self, user_id):
        """Get the display name for a profile.
        Priority: serial_number > name > remark > user_id
        Matches AutoIt _SEARCHCACHEFORSERIAL logic."""
        def _extract_name(u):
            if not u:
                return ''
            # Try fields in priority order (matching AutoIt)
            for field in ['serial_number', 'serialnumber']:
                val = str(u.get(field, '') or '').strip()
                if val and val != '0' and val.lower() != 'null':
                    return val
            for field in ['name', 'remark']:
                val = str(u.get(field, '') or '').strip()
                if val and val != '0' and val.lower() != 'null':
                    return val
            return ''

        # First check cache (limited pages for performance)
        users = self.get_user_list()
        if users:
            for u in users:
                uid = str(u.get('user_id') or u.get('userId') or u.get('id') or '')
                if uid == user_id:
                    name = _extract_name(u)
                    return name if name else user_id
        # Direct fetch by user_id (fast, single API call)
        u = self.get_user_by_id(user_id)
        if u:
            name = _extract_name(u)
            return name if name else user_id
        return user_id


# ─── Distribte Extension ─────────────────────────────────────────────────────

DISTRIBTE_SEARCH_PATHS = [
    r'C:\.ADSPOWER_GLOBAL', r'D:\.ADSPOWER_GLOBAL', r'E:\.ADSPOWER_GLOBAL',
    r'C:\ADSPOWER_GLOBAL', r'D:\ADSPOWER_GLOBAL', r'E:\ADSPOWER_GLOBAL',
]

def find_distribte_configs():
    """Find all autologin-config.js files in AdsPower extension dirs."""
    configs = []
    search_dirs = list(DISTRIBTE_SEARCH_PATHS)
    # Add user dirs
    local = os.environ.get('LOCALAPPDATA', '')
    appdata = os.environ.get('APPDATA', '')
    home = os.environ.get('USERPROFILE', '')
    if local:
        search_dirs.append(os.path.join(local, 'AdsPower Global'))
    if appdata:
        search_dirs.append(os.path.join(appdata, 'AdsPower Global'))
    if home:
        search_dirs.append(os.path.join(home, '.adspower_global'))

    for base_dir in search_dirs:
        if not os.path.isdir(base_dir):
            continue
        for root, dirs, files in os.walk(base_dir):
            for f in files:
                if f.lower() == 'autologin-config.js':
                    configs.append(os.path.join(root, f))
            # Limit depth
            depth = root[len(base_dir):].count(os.sep)
            if depth > 8:
                dirs.clear()
    return configs


def write_distribte_config(filepath, email, password, enabled=True):
    """Write autologin-config.js for the Distribte extension."""
    content = f'const AUTOLOGIN_CONFIG = {{ email: "{email}", password: "{password}", enabled: {str(enabled).lower()} }};\n'
    try:
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(content)
        return True
    except Exception:
        return False


# ─── Discord Integration ─────────────────────────────────────────────────────

def discord_webhook_send_text(webhook_url, content, username='APM'):
    if not webhook_url or not requests:
        return False
    try:
        r = requests.post(webhook_url, json={'content': content, 'username': username}, timeout=10)
        return r.ok
    except Exception:
        return False


def discord_webhook_upload_image(webhook_url, image_bytes, filename='profiles_1.png',
                                  content='', username='APM'):
    """Upload an image to Discord via webhook. Tries requests first, then http.client."""
    if not webhook_url:
        return None

    log_path = os.path.join(os.path.dirname(sys.argv[0]) or '.', 'discord_debug.log')
    debug = [f'=== v{VERSION} upload at {time.strftime("%Y-%m-%d %H:%M:%S")} ===',
             f'Image: {len(image_bytes)} bytes, user: {username}, content: {content[:50]}']

    url = webhook_url
    if '?' not in url:
        url += '?wait=true'
    elif 'wait=' not in url:
        url += '&wait=true'

    # Approach 1: requests (uses system proxy, same lib as working text send)
    if requests:
        try:
            debug.append('A1: requests with BytesIO...')
            buf = BytesIO(image_bytes)
            r = requests.post(url, data={'username': username, 'content': content},
                            files={'file0': (filename, buf, 'image/png')},
                            timeout=30, verify=False)
            debug.append(f'A1 status: {r.status_code}')
            debug.append(f'A1 body: {r.text[:400]}')
            if r.status_code in [200, 204]:
                resp = r.json()
                atts = resp.get('attachments', [])
                if atts:
                    debug.append('A1 SUCCESS')
                    _write_debug_log(log_path, debug)
                    return atts[0].get('url', '')
                if resp.get('id'):
                    debug.append('A1 SUCCESS (id only)')
                    _write_debug_log(log_path, debug)
                    return 'sent'
                debug.append('A1 WARN: 200 but no attachments/id')
            else:
                debug.append(f'A1 FAILED: HTTP {r.status_code}')
        except Exception as e:
            debug.append(f'A1 error: {type(e).__name__}: {e}')

    # Approach 2: http.client (pure stdlib, manual multipart like AutoIt WinHttp)
    try:
        import random
        import http.client
        import ssl
        from urllib.parse import urlparse

        boundary = f'----APMBoundary{random.randint(100000, 999999)}'
        body_parts = []
        body_parts.append(f'--{boundary}\r\nContent-Disposition: form-data; name="username"\r\n\r\n{username}\r\n'.encode('utf-8'))
        body_parts.append(f'--{boundary}\r\nContent-Disposition: form-data; name="content"\r\n\r\n{content}\r\n'.encode('utf-8'))
        body_parts.append(f'--{boundary}\r\nContent-Disposition: form-data; name="file0"; filename="{filename}"\r\nContent-Type: image/png\r\n\r\n'.encode('utf-8'))
        body_parts.append(image_bytes)
        body_parts.append(f'\r\n--{boundary}--\r\n'.encode('utf-8'))
        body = b''.join(body_parts)

        parsed = urlparse(url)
        ctx = ssl.create_default_context()
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE

        debug.append(f'A2: http.client to {parsed.hostname}, body={len(body)}b...')
        conn = http.client.HTTPSConnection(parsed.hostname, context=ctx, timeout=30)
        path = parsed.path
        if parsed.query:
            path += '?' + parsed.query
        conn.request('POST', path, body=body, headers={
            'Content-Type': f'multipart/form-data; boundary={boundary}',
            'Content-Length': str(len(body)),
        })
        resp = conn.getresponse()
        resp_body = resp.read().decode('utf-8')
        debug.append(f'A2 status: {resp.status}')
        debug.append(f'A2 body: {resp_body[:400]}')
        conn.close()

        if resp.status in [200, 204]:
            resp_json = json.loads(resp_body)
            atts = resp_json.get('attachments', [])
            if atts:
                debug.append('A2 SUCCESS')
                _write_debug_log(log_path, debug)
                return atts[0].get('url', '')
            if resp_json.get('id'):
                debug.append('A2 SUCCESS (id)')
                _write_debug_log(log_path, debug)
                return 'sent'
            debug.append('A2 WARN: 200 but no attachments')
        else:
            debug.append(f'A2 FAILED: HTTP {resp.status}')
    except Exception as e:
        debug.append(f'A2 error: {type(e).__name__}: {e}')

    debug.append('ALL FAILED')
    _write_debug_log(log_path, debug)
    return None


def _write_debug_log(path, lines):
    try:
        with open(path, 'a') as f:
            f.write('\n'.join(lines) + '\n\n')
    except Exception:
        pass


def log_to_google_sheets(sheet_url, sheet_name, rows):
    """POST data to Google Apps Script webhook."""
    if not sheet_url or not requests:
        return
    try:
        requests.post(sheet_url, json={'sheet': sheet_name, 'rows': rows}, timeout=10)
    except Exception:
        pass


# ─── Profile Image Generator ─────────────────────────────────────────────────

def generate_profile_image(browsers):
    """Generate a PNG table image of profiles matching AutoIt GDI+ specs.
    512px wide, Consolas 9pt, alternating rows, header with count."""
    if not HAS_PIL:
        return None

    col_profile_w = 220
    col_tab_w = 280
    padding = 12
    total_w = col_profile_w + col_tab_w + padding
    row_h = 20
    header_h = 24
    total_h = header_h + len(browsers) * row_h + padding

    img = Image.new('RGB', (total_w, max(total_h, 44)), 'white')
    draw = ImageDraw.Draw(img)

    try:
        font = ImageFont.truetype('consola.ttf', 12)
        font_bold = ImageFont.truetype('consolab.ttf', 12)
    except Exception:
        try:
            font = ImageFont.truetype('C:/Windows/Fonts/consola.ttf', 12)
            font_bold = ImageFont.truetype('C:/Windows/Fonts/consolab.ttf', 12)
        except Exception:
            try:
                font = ImageFont.truetype('cour.ttf', 12)
                font_bold = font
            except Exception:
                font = ImageFont.load_default()
                font_bold = font

    # Header - light gray like original
    draw.rectangle([0, 0, total_w, header_h], fill='#D6D6D6')
    draw.text((6, 4), f'Profile ({len(browsers)})', fill='black', font=font_bold)
    draw.text((col_profile_w + 6, 4), 'Tab', fill='black', font=font_bold)

    # Column separator
    draw.line([(col_profile_w, 0), (col_profile_w, total_h)], fill='#999999', width=1)
    # Header separator
    draw.line([(0, header_h), (total_w, header_h)], fill='#999999', width=1)

    # Rows with alternating colors
    for i, (hwnd, title, profile, tab) in enumerate(browsers):
        y = header_h + i * row_h
        bg = '#FFFFFF' if i % 2 == 0 else '#F0F0F0'
        draw.rectangle([0, y, total_w, y + row_h], fill=bg)
        draw.text((6, y + 2), str(profile)[:30], fill='black', font=font)
        draw.text((col_profile_w + 6, y + 2), str(tab)[:38], fill='black', font=font)

    # Border rectangle
    draw.rectangle([0, 0, total_w - 1, total_h - 1], outline='#999999', width=1)

    buf = BytesIO()
    img.save(buf, format='PNG')
    return buf.getvalue()


# ─── Main Application ────────────────────────────────────────────────────────

class APMApp:
    BG = '#f0f0f0'  # System-like light theme (matching AutoIt default)

    def __init__(self):
        ensure_dirs()
        self.cfg = load_config()
        api_key = self.cfg.get('ADSPOWER', 'ApiKey', fallback='')
        self.api = AdsPowerAPI(50325, api_key=api_key)
        self.browsers = []  # [(hwnd, title, profile_name, tab_title), ...]
        self.current_pos = 0
        self.running = True
        self.stop_url_loop = False
        self.browser_move_in_progress = False
        self.active_group = -1
        self.sort_by = 0  # 0=profile, 1=tab
        self.sort_reverse = True  # True=descending (highest first), False=ascending
        self.user_sorted = False  # only sort when user clicks header
        self.hotkeys_on = True
        self.extra_hotkeys_on = False
        self.pid_profile_cache = {}  # pid -> profile_name
        self.uid_map = {}  # user_id -> custom_number (permanent, like AutoIt $GUIDMAP)
        self.sunpid_cache = {}  # pid -> bool
        self.cmdline_cache = {}  # pid -> cmdline
        self.cmdline_cache_time = 0
        self.dist_queue = []  # [(hwnd, countdown), ...]
        self.dist_running = False
        self.debug_log = []

        self._build_gui()
        self._load_all_settings()
        self._register_hotkeys()
        self._start_polling()

    def _log(self, msg):
        ts = datetime.now().strftime('%H:%M:%S')
        self.debug_log.append(f'[{ts}] {msg}')
        if len(self.debug_log) > 500:
            self.debug_log = self.debug_log[-300:]

    # ══════════════════════════════════════════════════════════════════════════
    # GUI CONSTRUCTION - matches AutoIt layout
    # ══════════════════════════════════════════════════════════════════════════

    def _build_gui(self):
        self.root = tk.Tk()
        self.root.title(WINDOW_TITLE)

        w = int(self.cfg.get('MAIN', 'GUIW'))
        h = int(self.cfg.get('MAIN', 'GUIH'))
        x = int(self.cfg.get('MAIN', 'GUIX'))
        y = int(self.cfg.get('MAIN', 'GUIY'))
        self.root.geometry(f'{w}x{h}+{x}+{y}')
        self.root.minsize(374, 500)
        self.root.maxsize(374, 2000)
        self.root.resizable(True, True)

        always_on_top = self.cfg.get('MAIN', 'AlwaysOnTop') == '1'
        self.root.attributes('-topmost', always_on_top)

        # ── Tab Control ──
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=4, pady=(4, 0))

        # Create tab frames
        self.tab_main = ttk.Frame(self.notebook)
        self.tab_settings = ttk.Frame(self.notebook)
        self.tab_discord = ttk.Frame(self.notebook)
        self.tab_distribte = ttk.Frame(self.notebook)
        self.tab_pos = ttk.Frame(self.notebook)

        self.notebook.add(self.tab_main, text='Main')
        self.notebook.add(self.tab_settings, text='Settings')
        self.notebook.add(self.tab_discord, text='Discord')
        self.notebook.add(self.tab_distribte, text='Distribte')
        self.notebook.add(self.tab_pos, text='Pos')

        # Hotkeys + On top vars (checkboxes placed in bottom bar)
        self.hotkeys_var = tk.BooleanVar(value=self.cfg.get('MAIN', 'AllHotkeysON') == '1')
        self.ontop_var = tk.BooleanVar(value=always_on_top)

        self._build_main_tab()
        self._build_settings_tab()
        self._build_discord_tab()
        self._build_distribte_tab()
        self._build_pos_tab()
        self._build_bottom_bar()

        self.root.protocol('WM_DELETE_WINDOW', self._on_close)

    # ── MAIN TAB ──────────────────────────────────────────────────────────────

    def _build_main_tab(self):
        f = self.tab_main

        # Navigation: <<<, TOP, >>>
        nav = tk.Frame(f)
        nav.pack(fill='x', padx=4, pady=(4, 2))

        tk.Button(nav, text='<<<', width=8, command=self._move_back).pack(side='left', padx=2)
        tk.Button(nav, text='TOP', width=8, command=self._move_top).pack(side='left', padx=2)
        tk.Button(nav, text='>>>', width=8, command=self._move_fwd).pack(side='left', padx=2)

        # Main content: ListView + side buttons
        content = tk.Frame(f)
        content.pack(fill='both', expand=True, padx=4, pady=2)

        # Treeview (Profile | Tab | Handle)
        tree_frame = tk.Frame(content)
        tree_frame.pack(side='left', fill='both', expand=True)

        # Use 'clam' theme for column separator lines in treeview
        style = ttk.Style()
        style.theme_use('clam')

        cols = ('profile', 'tab', 'handle')
        self.tree = ttk.Treeview(tree_frame, columns=cols, show='headings', selectmode='extended')
        self.tree.heading('profile', text='Profile', command=lambda: self._sort_tree(0))
        self.tree.heading('tab', text='Tab', command=lambda: self._sort_tree(1))
        self.tree.heading('handle', text='Handle')
        self.tree.column('profile', width=120, minwidth=60)
        self.tree.column('tab', width=120, minwidth=60)
        self.tree.column('handle', width=0, stretch=False)  # Hidden

        scrollbar = ttk.Scrollbar(tree_frame, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        self.tree.bind('<<TreeviewSelect>>', self._on_select)
        self.tree.bind('<Double-1>', lambda e: self._browser_action('show'))

        # Side buttons
        side = tk.Frame(content, width=65)
        side.pack(side='right', fill='y', padx=(4, 0))
        side.pack_propagate(False)

        buttons = [
            ('Show', lambda: self._browser_action('show')),
            ('Minimize', lambda: self._browser_action('minimize')),
            ('RefreshAll', lambda: self._browser_action('refresh', all_=True)),
            ('Close All', lambda: self._browser_action('close', all_=True)),
            ('Close Sel', lambda: self._browser_action('close')),
            ('Show All', lambda: self._browser_action('show', all_=True)),
            ('MinimizeAll', lambda: self._browser_action('minimize', all_=True)),
        ]
        for text, cmd in buttons:
            tk.Button(side, text=text, font=('', 7), width=9, command=cmd).pack(pady=1)

        # Groups A-Z
        grp_label = tk.Label(side, text='Grp:', font=('', 7, 'bold'))
        grp_label.pack(pady=(6, 1))

        self.grp_btns = {}
        for row_start in range(0, 26, 3):
            row_frame = tk.Frame(side)
            row_frame.pack()
            for i in range(3):
                idx = row_start + i
                if idx >= 26:
                    break
                letter = chr(65 + idx)
                lbl = tk.Label(row_frame, text=letter, font=('', 6, 'bold'),
                               width=2, padx=1, pady=0, relief='raised', bd=1,
                               bg='#E0E0E0', fg='black', cursor='hand2')
                lbl.bind('<Button-1>', lambda e, l=idx: self._switch_group(l))
                lbl.pack(side='left', padx=1)
                self.grp_btns[idx] = lbl

        # W/H resize
        size_frame = tk.Frame(side)
        size_frame.pack(pady=(4, 0))

        wh = tk.Frame(size_frame)
        wh.pack()
        tk.Label(wh, text='W:', font=('', 7)).pack(side='left')
        self.main_w_entry = tk.Entry(wh, width=5, font=('', 8))
        self.main_w_entry.insert(0, self.cfg.get('MAIN', 'NavWidth'))
        self.main_w_entry.pack(side='left')

        wh2 = tk.Frame(size_frame)
        wh2.pack()
        tk.Label(wh2, text='H:', font=('', 7)).pack(side='left')
        self.main_h_entry = tk.Entry(wh2, width=5, font=('', 8))
        self.main_h_entry.insert(0, self.cfg.get('MAIN', 'NavHeight'))
        self.main_h_entry.pack(side='left')

        tk.Button(size_frame, text='Apply', font=('', 7), width=7,
                  command=self._apply_resize).pack(pady=1)
        tk.Button(size_frame, text='Fix', font=('', 7), width=7,
                  command=self._fix_all_sizes).pack(pady=1)

        # (TM Lite removed per client request)

    # ── SETTINGS TAB ──────────────────────────────────────────────────────────

    def _build_settings_tab(self):
        f = self.tab_settings
        canvas = tk.Canvas(f)
        scrollbar = tk.Scrollbar(f, orient='vertical', command=canvas.yview)
        inner = tk.Frame(canvas)
        inner.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.create_window((0, 0), window=inner, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        row = 0
        tk.Label(inner, text='Primary Hotkeys', font=('', 9, 'bold')).grid(
            row=row, column=0, columnspan=2, sticky='w', pady=(6, 2), padx=8)
        row += 1

        self.hk_entries = {}
        hk_fields = [
            ('FORWARD', 'Forward'), ('BACKWARD', 'Backward'), ('TOP', 'Top'),
            ('SORTTAB', 'Sort by Tab'), ('SORTPROFILE', 'Sort by Profile'),
            ('GROUPNEXT', 'Group Next'), ('GROUPBACK', 'Group Back'),
        ]
        for key, label in hk_fields:
            tk.Label(inner, text=label, font=('', 8)).grid(row=row, column=0, sticky='w', padx=8)
            entry = tk.Entry(inner, font=('', 8), width=20)
            entry.insert(0, self.cfg.get('HOTKEYS', key, fallback=''))
            entry.grid(row=row, column=1, padx=4, pady=1)
            self.hk_entries[key] = entry
            row += 1

        # Extra hotkeys
        tk.Label(inner, text='Extra Hotkeys (9 slots)', font=('', 9, 'bold')).grid(
            row=row, column=0, columnspan=2, sticky='w', pady=(10, 2), padx=8)
        row += 1

        tk.Label(inner, text='Desc | Action | Key', font=('', 7)).grid(
            row=row, column=0, columnspan=2, sticky='w', padx=8)
        row += 1

        self.ehk_entries = []
        for i in range(9):
            fr = tk.Frame(inner)
            fr.grid(row=row, column=0, columnspan=2, sticky='w', padx=8, pady=1)
            e1 = tk.Entry(fr, font=('', 7), width=10)
            e1.insert(0, self.cfg.get('HOTKEYS2', f'EHK{i+1}-0', fallback=''))
            e1.pack(side='left', padx=1)
            e2 = tk.Entry(fr, font=('', 7), width=10)
            e2.insert(0, self.cfg.get('HOTKEYS2', f'EHK{i+1}-1', fallback=''))
            e2.pack(side='left', padx=1)
            e3 = tk.Entry(fr, font=('', 7), width=10)
            e3.insert(0, self.cfg.get('HOTKEYS2', f'EHK{i+1}-2', fallback=''))
            e3.pack(side='left', padx=1)
            self.ehk_entries.append((e1, e2, e3))
            row += 1

        # Toggle hotkey
        fr = tk.Frame(inner)
        fr.grid(row=row, column=0, columnspan=2, sticky='w', padx=8, pady=1)
        tk.Label(fr, text='Toggle Extra HK:', font=('', 7)).pack(side='left')
        self.ehk_toggle_entry = tk.Entry(fr, font=('', 7), width=15)
        self.ehk_toggle_entry.insert(0, self.cfg.get('MAIN', 'HotkeysToggleExtra'))
        self.ehk_toggle_entry.pack(side='left', padx=2)
        row += 1

        # Options
        tk.Label(inner, text='Options', font=('', 9, 'bold')).grid(
            row=row, column=0, columnspan=2, sticky='w', pady=(10, 2), padx=8)
        row += 1

        self.opt_autosorting = tk.BooleanVar(value=self.cfg.get('MAIN', 'AutoSorting') == '1')
        tk.Checkbutton(inner, text='Auto column sorting', variable=self.opt_autosorting).grid(
            row=row, column=0, columnspan=2, sticky='w', padx=12)
        row += 1

        self.opt_inject = tk.BooleanVar(value=self.cfg.get('MAIN', 'InjectControls') == '1')
        tk.Checkbutton(inner, text='Inject controls inside each browser',
                        variable=self.opt_inject).grid(
            row=row, column=0, columnspan=2, sticky='w', padx=12)
        row += 1

        self.opt_minimize_others = tk.BooleanVar(value=self.cfg.get('MAIN', 'MinimizeOthers') == '1')
        tk.Checkbutton(inner, text='Minimize others on profile select',
                        variable=self.opt_minimize_others).grid(
            row=row, column=0, columnspan=2, sticky='w', padx=12)
        row += 1

        self.opt_profile_saver = tk.BooleanVar(value=self.cfg.get('MAIN', 'AutoProfileSaver', fallback='0') == '1')
        tk.Checkbutton(inner, text='Auto-save Chrome profile (click "Continue as" popup)',
                        variable=self.opt_profile_saver).grid(
            row=row, column=0, columnspan=2, sticky='w', padx=12)
        row += 1

        self.opt_custom_nav = tk.BooleanVar(value=self.cfg.get('MAIN', 'CustomNavSize') == '1')
        cb_frame = tk.Frame(inner)
        cb_frame.grid(row=row, column=0, columnspan=2, sticky='w', padx=12)
        tk.Checkbutton(cb_frame, text='Use custom Click/Nav size:', variable=self.opt_custom_nav).pack(side='left')
        row += 1

        sz_frame = tk.Frame(inner)
        sz_frame.grid(row=row, column=0, columnspan=2, sticky='w', padx=24)
        tk.Label(sz_frame, text='W:').pack(side='left')
        self.set_navw = tk.Entry(sz_frame, width=5)
        self.set_navw.insert(0, self.cfg.get('MAIN', 'NavWidth'))
        self.set_navw.pack(side='left', padx=2)
        tk.Label(sz_frame, text='H:').pack(side='left')
        self.set_navh = tk.Entry(sz_frame, width=5)
        self.set_navh.insert(0, self.cfg.get('MAIN', 'NavHeight'))
        self.set_navh.pack(side='left', padx=2)
        row += 1

        # AdsPower API Key
        tk.Label(inner, text='AdsPower', font=('', 9, 'bold')).grid(
            row=row, column=0, columnspan=2, sticky='w', pady=(10, 2), padx=8)
        row += 1
        ak_frame = tk.Frame(inner)
        ak_frame.grid(row=row, column=0, columnspan=2, sticky='w', padx=12)
        tk.Label(ak_frame, text='API Key:', font=('', 8)).pack(side='left')
        self.set_apikey = tk.Entry(ak_frame, width=30, font=('', 8))
        self.set_apikey.insert(0, self.cfg.get('ADSPOWER', 'ApiKey', fallback=''))
        self.set_apikey.pack(side='left', padx=4)
        row += 1

        tk.Button(inner, text='Save Settings', command=self._save_settings).grid(
            row=row, column=0, columnspan=2, pady=10)

    # ── DISCORD TAB ───────────────────────────────────────────────────────────

    def _build_discord_tab(self):
        f = self.tab_discord
        row = 0

        tk.Label(f, text='Profile Name:').grid(row=row, column=0, sticky='w', padx=8, pady=2)
        self.dc_profile = tk.Entry(f, width=30)
        self.dc_profile.grid(row=row, column=1, padx=4, pady=2)
        row += 1

        tk.Label(f, text='Message:').grid(row=row, column=0, sticky='w', padx=8, pady=2)
        self.dc_message = tk.Text(f, height=3, width=30)
        self.dc_message.grid(row=row, column=1, padx=4, pady=2)
        row += 1

        btn_frame2 = tk.Frame(f)
        btn_frame2.grid(row=row, column=0, columnspan=2, pady=4)
        tk.Button(btn_frame2, text='QUE Screenshot', command=lambda: self._discord_screenshot('que')).pack(side='left', padx=2)
        tk.Button(btn_frame2, text='PROD Screenshot', command=lambda: self._discord_screenshot('prod')).pack(side='left', padx=2)
        tk.Button(btn_frame2, text='VF Screenshot', command=lambda: self._discord_screenshot('vf')).pack(side='left', padx=2)
        tk.Button(btn_frame2, text='Save Only', command=self._save_screenshot).pack(side='left', padx=2)
        row += 1

        self.dc_status = tk.Label(f, text='', fg='green')
        self.dc_status.grid(row=row, column=0, columnspan=2, pady=2)
        row += 1

        # Webhook URLs
        for key, label in [('QueWebhook', 'QUE Webhook:'), ('ProdWebhook', 'PROD Webhook:'),
                            ('VfWebhook', 'VF Webhook:')]:
            tk.Label(f, text=label, font=('', 7)).grid(row=row, column=0, sticky='w', padx=8)
            entry = tk.Entry(f, width=35, font=('', 7))
            entry.insert(0, self.cfg.get('DISCORD', key, fallback=''))
            entry.grid(row=row, column=1, padx=4, pady=1)
            setattr(self, f'dc_{key.lower()}', entry)
            row += 1

        tk.Label(f, text='Sheet URL:', font=('', 7)).grid(row=row, column=0, sticky='w', padx=8)
        self.dc_sheeturl = tk.Entry(f, width=35, font=('', 7))
        self.dc_sheeturl.insert(0, self.cfg.get('DISCORD', 'SheetUrl', fallback=''))
        self.dc_sheeturl.grid(row=row, column=1, padx=4, pady=1)
        row += 1

        tk.Label(f, text='Save Folder:', font=('', 7)).grid(row=row, column=0, sticky='w', padx=8)
        folder_frame = tk.Frame(f)
        folder_frame.grid(row=row, column=1, sticky='w', padx=4)
        self.dc_folder = tk.Entry(folder_frame, width=25, font=('', 7))
        self.dc_folder.insert(0, self.cfg.get('DISCORD', 'ScreenshotFolder', fallback=''))
        self.dc_folder.pack(side='left')
        tk.Button(folder_frame, text='...', font=('', 7), width=3,
                  command=self._browse_screenshot_folder).pack(side='left')
        row += 1

        tk.Button(f, text='Save Discord Settings', command=self._save_discord).grid(
            row=row, column=0, columnspan=2, pady=8)

    # ── DISTRIBTE TAB ─────────────────────────────────────────────────────────

    def _build_distribte_tab(self):
        f = self.tab_distribte

        tk.Label(f, text='Distribte Auto-Login', font=('', 10, 'bold')).pack(
            anchor='w', padx=10, pady=(10, 4))
        tk.Label(f, text='Saves email/password to Distribte extension configs\n'
                          'in all AdsPower profile directories.', font=('', 8)).pack(
            anchor='w', padx=12, pady=4)

        form = tk.Frame(f)
        form.pack(padx=12, pady=4)
        tk.Label(form, text='Email:').grid(row=0, column=0, sticky='w')
        self.dist_email = tk.Entry(form, width=30)
        self.dist_email.insert(0, self.cfg.get('DISTRIBTE', 'Email', fallback=''))
        self.dist_email.grid(row=0, column=1, padx=4, pady=2)

        tk.Label(form, text='Password:').grid(row=1, column=0, sticky='w')
        self.dist_password = tk.Entry(form, width=30, show='*')
        self.dist_password.insert(0, self.cfg.get('DISTRIBTE', 'Password', fallback=''))
        self.dist_password.grid(row=1, column=1, padx=4, pady=2)

        btn_frame = tk.Frame(f)
        btn_frame.pack(pady=8)
        tk.Button(btn_frame, text='Save & Apply', bg='#4CA070', fg='white',
                  command=self._dist_save).pack(side='left', padx=4)
        tk.Button(btn_frame, text='Clear', bg='#e94560', fg='white',
                  command=self._dist_clear).pack(side='left', padx=4)

        self.dist_status = tk.Label(f, text='', font=('', 8))
        self.dist_status.pack(pady=4)

    # ── POSITIONER TAB ────────────────────────────────────────────────────────

    def _build_pos_tab(self):
        f = self.tab_pos

        grid_frame = tk.Frame(f)
        grid_frame.pack(padx=10, pady=8)

        self.pos_entries = {}
        for i, (key, label, default) in enumerate([
            ('Cols', 'Cols:', '4'), ('Rows', 'Rows:', '2'),
            ('Width', 'Width:', '480'), ('Height', 'Height:', '540'),
            ('GapX', 'Gap X:', '0'), ('GapY', 'Gap Y:', '0'),
        ]):
            row, col = divmod(i, 2)
            tk.Label(grid_frame, text=label, font=('', 8)).grid(row=row, column=col*2, sticky='w', padx=4)
            entry = tk.Entry(grid_frame, width=6, font=('', 8))
            entry.insert(0, self.cfg.get('POSITIONER', key, fallback=default))
            entry.grid(row=row, column=col*2+1, padx=2, pady=2)
            self.pos_entries[key] = entry

        tk.Button(f, text='Position Windows', command=self._position_windows).pack(pady=4)

        # URL bar
        url_frame = tk.Frame(f)
        url_frame.pack(fill='x', padx=10, pady=4)
        self.pos_url = tk.Entry(url_frame, font=('', 8))
        self.pos_url.insert(0, self.cfg.get('POSITIONER', 'URL', fallback='https://www.ticketmaster.com'))
        self.pos_url.pack(side='left', fill='x', expand=True)
        tk.Button(url_frame, text='Open URL', font=('', 8),
                  command=self._pos_open_url).pack(side='left', padx=2)
        tk.Button(url_frame, text='STOP', font=('', 8), fg='white', bg='red',
                  command=self._stop_url).pack(side='left', padx=2)

        tk.Button(f, text='Save Positioner Settings', command=self._save_pos).pack(pady=4)

        # Group buttons A-Z (larger, 9 columns)
        tk.Label(f, text='Groups:', font=('', 8, 'bold')).pack(anchor='w', padx=10, pady=(8, 2))
        self.pos_grp_btns = {}
        for row_start in range(0, 26, 9):
            row_frame = tk.Frame(f)
            row_frame.pack(padx=10)
            for i in range(9):
                idx = row_start + i
                if idx >= 26:
                    break
                letter = chr(65 + idx)
                lbl = tk.Label(row_frame, text=letter, font=('', 8, 'bold'),
                               width=3, padx=2, pady=1, relief='raised', bd=1,
                               bg='#E0E0E0', fg='black', cursor='hand2')
                lbl.bind('<Button-1>', lambda e, l=idx: self._switch_group(l))
                lbl.pack(side='left', padx=1, pady=1)
                self.pos_grp_btns[idx] = lbl

    # ── BOTTOM BAR ────────────────────────────────────────────────────────────

    def _build_bottom_bar(self):
        # Top row: Hotkeys, On top, Debug Log
        top_bar = tk.Frame(self.root)
        top_bar.pack(fill='x', side='bottom', padx=4, pady=(0, 1))

        tk.Checkbutton(top_bar, text='Hotkeys', variable=self.hotkeys_var,
                        font=('', 7), command=self._toggle_hotkeys).pack(side='left', padx=2)
        tk.Checkbutton(top_bar, text='On top', variable=self.ontop_var,
                        font=('', 7), command=self._toggle_ontop).pack(side='left', padx=2)
        tk.Button(top_bar, text='Debug Log', font=('', 6),
                  command=self._show_debug_log).pack(side='right', padx=2)

        # Bottom row: Version, URL, Open, Stop
        bottom = tk.Frame(self.root)
        bottom.pack(fill='x', side='bottom', padx=4, pady=(0, 4))

        tk.Label(bottom, text=f'APM v{VERSION}', font=('', 7), fg='gray').pack(side='left')

        self.main_url = tk.Entry(bottom, font=('', 8), width=18)
        self.main_url.insert(0, self.cfg.get('MAIN', 'MainURL'))
        self.main_url.pack(side='left', padx=4)

        tk.Button(bottom, text='Open URL', font=('', 7),
                  command=self._open_url_all).pack(side='left', padx=2)
        tk.Button(bottom, text='STOP', font=('', 7, 'bold'), fg='white', bg='red',
                  command=self._stop_url).pack(side='left', padx=2)

    # ══════════════════════════════════════════════════════════════════════════
    # CORE LOGIC
    # ══════════════════════════════════════════════════════════════════════════

    def _get_browsers(self):
        """Scan for SunBrowser windows - equivalent to GetBrowsers() in AutoIt."""
        if not HAS_WIN32:
            self._log('No Win32 API available')
            return

        try:
            chrome_windows = enum_windows()
        except Exception as e:
            self._log(f'EnumWindows error: {e}')
            return

        new_browsers = []

        # Log scan results periodically
        if not hasattr(self, '_scan_log_count'):
            self._scan_log_count = 0
        self._scan_log_count += 1
        verbose = self._scan_log_count <= 5 or self._scan_log_count % 20 == 0
        if verbose:
            self._log(f'Scan: {len(chrome_windows)} Chrome_WidgetWin_1 windows')
            for hw, t in chrome_windows:
                pid = get_window_pid(hw)
                sun = self.sunpid_cache.get(pid, '?')
                self._log(f'  hwnd={hw} pid={pid} sun={sun} title={t[:50]}')
            # PID-based diagnostic: find ALL windows for known SunBrowser PIDs
            sun_pids = {p for p, v in self.sunpid_cache.items() if v}
            if sun_pids:
                all_wins = enum_all_windows_for_pids(sun_pids)
                non_chrome = [w for w in all_wins if w[2] != CHROME_CLASS]
                if non_chrome:
                    self._log(f'  Extra SunBrowser windows (non-Chrome class):')
                    for hw, t, cls, pid in non_chrome:
                        self._log(f'    hwnd={hw} pid={pid} class={cls} title={t[:40]}')
                chrome_count = len([w for w in all_wins if w[2] == CHROME_CLASS])
                self._log(f'  SunBrowser PIDs: {len(sun_pids)}, Chrome wins: {chrome_count}, Other wins: {len(non_chrome)}')

        # Refresh cmdline cache every 30s (profile cache stays permanent like AutoIt)
        now = time.time()
        if now - self.cmdline_cache_time > 30:
            self.cmdline_cache.clear()
            self.cmdline_cache_time = now

        # Budget for API calls per cycle (matches AutoIt $GRESOLVEBUDGET)
        resolve_budget = 3
        new_budget = 10  # Max new profiles per cycle

        for hwnd, title in chrome_windows:
            pid = get_window_pid(hwnd)
            if not pid:
                continue

            # Check SunBrowser (cache result)
            if pid not in self.sunpid_cache:
                is_sun = is_sunbrowser(pid)
                self.sunpid_cache[pid] = is_sun
                if self._scan_log_count <= 5:
                    exe = get_process_exe(pid)
                    self._log(f'PID {pid}: exe={exe[-40:]}, sun={is_sun}')
            if not self.sunpid_cache[pid]:
                continue

            # Resolve profile name
            if pid in self.pid_profile_cache:
                profile_name = self.pid_profile_cache[pid]

                # RETRY: if profile name looks like a raw user_id (6-12 lowercase alnum),
                # retry resolution on each cycle with budget (matches AutoIt retry logic)
                if (re.match(r'^[a-z0-9]{6,12}$', profile_name)
                        and resolve_budget > 0):
                    # Check uid_map first (instant, no budget needed)
                    if profile_name in self.uid_map:
                        old = profile_name
                        profile_name = self.uid_map[profile_name]
                        self.pid_profile_cache[pid] = profile_name
                        self._log(f'Retry resolved (map): {old} -> {profile_name}')
                    else:
                        resolve_budget -= 1
                        new_name = self.api.resolve_profile_name(profile_name)
                        if new_name and new_name != profile_name:
                            self.uid_map[profile_name] = new_name
                            self.pid_profile_cache[pid] = new_name
                            profile_name = new_name
                            self._log(f'Retry resolved (API): {profile_name}')
            else:
                if new_budget <= 0:
                    continue
                new_budget -= 1

                if pid not in self.cmdline_cache:
                    self.cmdline_cache[pid] = get_process_cmdline(pid)
                cmdline = self.cmdline_cache[pid]
                user_id = get_adspower_userid_from_cmdline(cmdline)
                if user_id:
                    # Check uid_map first (instant lookup like AutoIt $GUIDMAP)
                    if user_id in self.uid_map:
                        profile_name = self.uid_map[user_id]
                        self._log(f'Profile from map: uid={user_id} -> {profile_name}')
                    else:
                        profile_name = self.api.resolve_profile_name(user_id)
                        if profile_name != user_id:
                            self.uid_map[user_id] = profile_name
                            self._log(f'Profile resolved: uid={user_id} -> {profile_name}')
                        else:
                            # Show user_id for now, will retry on next cycles
                            self._log(f'Profile pending: uid={user_id} (will retry)')
                else:
                    # Fallback: session name from cmdline
                    m = re.search(r'--session[_-]name="?([^"\s]+)"?', cmdline)
                    profile_name = m.group(1) if m else f'PID:{pid}'
                    if self._scan_log_count <= 5:
                        self._log(f'PID {pid}: no user_id, using {profile_name}')
                self.pid_profile_cache[pid] = profile_name

            # Tab title
            tab_title = title if title else 'New Tab'

            new_browsers.append((hwnd, title, profile_name, tab_title))

        self.browsers = new_browsers
        self.root.after(0, self._refresh_tree)

    def _refresh_tree(self):
        """Incremental tree update - matches AutoIt GETBROWSERS behavior.
        Only adds new windows, removes closed ones, updates changed titles.
        Never rebuilds from scratch (prevents navigation jumping)."""
        self._refreshing_tree = True
        try:
            # Build lookup of current scan results by HWND
            scan_hwnds = {}  # hwnd_str -> (profile, tab)
            for hwnd, title, profile, tab in self.browsers:
                scan_hwnds[str(hwnd)] = (profile, tab)

            needs_sort = False

            # 1. Update existing items / remove closed windows
            for item in list(self.tree.get_children()):
                vals = self.tree.item(item, 'values')
                if not vals or len(vals) < 3:
                    self.tree.delete(item)
                    continue
                hwnd_str = vals[2]
                if hwnd_str not in scan_hwnds:
                    # Window closed - remove from tree
                    self.tree.delete(item)
                    needs_sort = True
                else:
                    # Window still open - update profile/tab if changed
                    new_profile, new_tab = scan_hwnds[hwnd_str]
                    old_profile, old_tab = vals[0], vals[1]
                    if new_profile != old_profile or new_tab != old_tab:
                        self.tree.item(item, values=(new_profile, new_tab, hwnd_str))
                        needs_sort = True
                    # Mark as processed
                    del scan_hwnds[hwnd_str]

            # 2. Add new windows (those not already in tree)
            for hwnd_str, (profile, tab) in scan_hwnds.items():
                self.tree.insert('', 'end', values=(profile, tab, hwnd_str))
                needs_sort = True

            # 3. Update count in heading
            count = len(self.tree.get_children())
            self.tree.heading('profile', text=f'Profile ({count})')

            # 4. No auto-sort - sorting only happens when user clicks column header
            # (matches MLM behavior)

            # 5. Update current_pos to match selected item's position
            sel = self.tree.selection()
            if sel:
                items = self.tree.get_children()
                for i, item in enumerate(items):
                    if item == sel[0]:
                        self.current_pos = i
                        break
        finally:
            self._refreshing_tree = False

    def _sort_tree(self, col, toggle=True):
        # Debounce: ignore rapid sort clicks (within 300ms)
        now = time.time()
        if hasattr(self, '_last_sort_time') and (now - self._last_sort_time) < 0.3:
            return
        self._last_sort_time = now
        self._sorting_in_progress = True
        try:
            if toggle:
                if col == self.sort_by:
                    self.sort_reverse = not self.sort_reverse
                else:
                    self.sort_reverse = True
                self.user_sorted = True
            self.sort_by = col
            items = [(self.tree.set(k, ('profile', 'tab')[col]), k) for k in self.tree.get_children()]

            def sort_key(x):
                val = x[0].strip()
                try:
                    return (0, float(val))
                except (ValueError, TypeError):
                    return (1, val.lower())

            items.sort(key=sort_key, reverse=self.sort_reverse)
            for i, (_, k) in enumerate(items):
                self.tree.move(k, '', i)

            col_name = ('profile', 'tab')[col]
            arrow = ' v' if self.sort_reverse else ' ^'
            count = len(self.tree.get_children())
            if col == 0:
                self.tree.heading('profile', text=f'Profile ({count}){arrow}')
                self.tree.heading('tab', text='Tab')
            else:
                self.tree.heading('profile', text=f'Profile ({count})')
                self.tree.heading('tab', text=f'Tab{arrow}')
        finally:
            self._sorting_in_progress = False

    def _on_select(self, event):
        """Show selected browser window on user click.
        Guards: skip during refresh, sort, or browser_move (those are programmatic)."""
        if self.browser_move_in_progress:
            return
        if getattr(self, '_refreshing_tree', False):
            return
        if getattr(self, '_sorting_in_progress', False):
            return
        sel = self.tree.selection()
        if sel:
            items = self.tree.get_children()
            for i, item in enumerate(items):
                if item == sel[0]:
                    self.current_pos = i
                    break
            # Show the selected window (matches AutoIt click behavior)
            vals = self.tree.item(sel[0], 'values')
            if vals and len(vals) > 2:
                try:
                    hwnd = int(vals[2])
                    if self.opt_minimize_others.get():
                        for other_item in items:
                            if other_item != sel[0]:
                                other_vals = self.tree.item(other_item, 'values')
                                if other_vals and len(other_vals) > 2:
                                    try:
                                        minimize_window(int(other_vals[2]))
                                    except Exception:
                                        pass
                    if self.opt_custom_nav.get():
                        # On click: keep current position, just resize (like AutoIt)
                        self._show_custom_size_keep_pos(hwnd)
                    else:
                        show_window(hwnd)
                except (ValueError, TypeError):
                    pass

    def _get_selected_hwnds(self):
        """Get HWNDs of selected items."""
        hwnds = []
        for item in self.tree.selection():
            vals = self.tree.item(item, 'values')
            if vals and len(vals) > 2:
                try:
                    hwnds.append(int(vals[2]))
                except (ValueError, TypeError):
                    pass
        return hwnds

    def _get_all_hwnds(self):
        """Get all HWNDs from treeview."""
        hwnds = []
        for item in self.tree.get_children():
            vals = self.tree.item(item, 'values')
            if vals and len(vals) > 2:
                try:
                    hwnds.append(int(vals[2]))
                except (ValueError, TypeError):
                    pass
        return hwnds

    def _show_custom_size(self, hwnd):
        """Restore + resize to custom nav dimensions at x=0 (for forward/backward)."""
        try:
            w = int(self.set_navw.get())
        except (ValueError, AttributeError):
            w = 480
        try:
            h = int(self.set_navh.get())
        except (ValueError, AttributeError):
            h = 540
        if w < 50:
            w = 480
        if h < 50:
            h = 540
        restore_and_resize(hwnd, w, h)

    def _show_custom_size_keep_pos(self, hwnd):
        """Restore + resize but keep window at its current position (for click activation).
        Matches AutoIt: on click, window stays where it is, just resized."""
        try:
            w = int(self.set_navw.get())
        except (ValueError, AttributeError):
            w = 480
        try:
            h = int(self.set_navh.get())
        except (ValueError, AttributeError):
            h = 540
        if w < 50:
            w = 480
        if h < 50:
            h = 540
        if HAS_WIN32:
            user32.ShowWindow(hwnd, SW_RESTORE)
            # Get current position, keep it
            rect = ctypes.wintypes.RECT()
            user32.GetWindowRect(hwnd, ctypes.byref(rect))
            x, y = rect.left, rect.top
            if x < 0 or y < 0:
                x, y = 100, 100
            user32.SetWindowPos(hwnd, None, x, y, w, h, SWP_SHOWWINDOW | SWP_FRAMECHANGED)
            force_foreground(hwnd)

    # ── Navigation ────────────────────────────────────────────────────────────

    def _move_fwd(self):
        self._browser_move('fwd')

    def _move_back(self):
        self._browser_move('bck')

    def _move_top(self):
        self._browser_move('top')

    def _browser_move(self, direction):
        items = self.tree.get_children()
        if not items:
            return
        self.browser_move_in_progress = True
        try:
            count = len(items)
            # Read current selection first (matches AutoIt BROWSERMOVE)
            sel = self.tree.selection()
            if sel:
                for i, it in enumerate(items):
                    if it == sel[0]:
                        self.current_pos = i
                        break

            if direction == 'fwd':
                self.current_pos = (self.current_pos + 1) % count
            elif direction == 'bck':
                self.current_pos = (self.current_pos - 1) % count
            elif direction == 'top':
                self.current_pos = 0

            item = items[self.current_pos]
            self.tree.selection_set(item)
            self.tree.see(item)

            vals = self.tree.item(item, 'values')
            if vals and len(vals) > 2:
                hwnd = int(vals[2])

                # Minimize others if enabled
                if self.opt_minimize_others.get():
                    for other_item in items:
                        if other_item != item:
                            other_vals = self.tree.item(other_item, 'values')
                            if other_vals and len(other_vals) > 2:
                                try:
                                    minimize_window(int(other_vals[2]))
                                except Exception:
                                    pass

                # Custom nav size: restore + resize (keep position)
                # No custom nav: maximize (SW_MAXIMIZE)
                if self.opt_custom_nav.get():
                    self._show_custom_size(hwnd)
                else:
                    show_window(hwnd)
        finally:
            self.browser_move_in_progress = False

    # ── Browser Actions ───────────────────────────────────────────────────────

    def _browser_action(self, action, all_=False):
        if all_:
            hwnds = self._get_all_hwnds()
        else:
            hwnds = self._get_selected_hwnds()

        if not hwnds:
            return

        if action == 'close' and len(hwnds) > 1:
            if not messagebox.askyesno('Confirm', f'Close {len(hwnds)} windows?'):
                return

        def do_action():
            for hwnd in hwnds:
                try:
                    if not is_window_visible(hwnd):
                        continue
                    if action == 'show':
                        if self.opt_custom_nav.get():
                            self._show_custom_size(hwnd)
                        else:
                            show_window(hwnd)
                    elif action == 'minimize':
                        minimize_window(hwnd)
                    elif action == 'close':
                        close_window(hwnd)
                        time.sleep(0.5)
                        if is_window_visible(hwnd):
                            user32.DestroyWindow(hwnd)
                    elif action == 'refresh':
                        # AutoIt uses SW_RESTORE for refresh, not maximize
                        if HAS_WIN32:
                            user32.ShowWindow(hwnd, SW_RESTORE)
                            user32.SetForegroundWindow(hwnd)
                        time.sleep(0.3)
                        send_keys_to_window(hwnd, ['{F5}'])
                        time.sleep(1.0)
                except Exception:
                    pass

        threading.Thread(target=do_action, daemon=True).start()

    # ── Groups ────────────────────────────────────────────────────────────────

    def _switch_group(self, group_index):
        """Switch to group (virtual partition based on Cols*Rows).
        Matches AutoIt SWITCHGROUP logic exactly."""
        if self.active_group == group_index:
            # Toggle off
            self.active_group = -1
            all_btns = list(self.grp_btns.items()) + list(self.pos_grp_btns.items())
            for idx, btn in all_btns:
                letter = chr(65 + idx)
                btn.configure(text=letter, bg='#E0E0E0', fg='black',
                              relief='raised', bd=1)
            return

        # Read positioner settings
        try:
            cols = max(1, int(self.pos_entries.get('Cols', tk.Entry()).get() or '4'))
            rows = max(1, int(self.pos_entries.get('Rows', tk.Entry()).get() or '2'))
            width = int(self.pos_entries.get('Width', tk.Entry()).get() or '480')
            height = int(self.pos_entries.get('Height', tk.Entry()).get() or '540')
            gap_x = int(self.pos_entries.get('GapX', tk.Entry()).get() or '0')
            gap_y = int(self.pos_entries.get('GapY', tk.Entry()).get() or '0')
        except (ValueError, TypeError):
            cols, rows, width, height, gap_x, gap_y = 4, 2, 480, 540, 0, 0

        group_size = cols * rows
        hwnds = self._get_all_hwnds()
        count = len(hwnds)
        if count == 0:
            return

        total_groups = math.ceil(count / group_size)
        if group_index >= total_groups:
            return

        start = group_index * group_size
        end = min(start + group_size, count)

        self.active_group = group_index
        # Highlight active button, reset others (iterate both sets separately)
        all_btns = list(self.grp_btns.items()) + list(self.pos_grp_btns.items())
        for idx, btn in all_btns:
            letter = chr(65 + idx)
            if idx == group_index:
                btn.configure(text=f'[{letter}]', bg='#00AA00', fg='white',
                              relief='sunken', bd=2)
            else:
                btn.configure(text=letter, bg='#E0E0E0', fg='black',
                              relief='raised', bd=1)

        def do_switch():
            # Step 1: Minimize all windows NOT in this group
            for i, hwnd in enumerate(hwnds):
                if i < start or i >= end:
                    minimize_window(hwnd)
            time.sleep(0.1)

            # Step 2: Position group windows in grid using SW_SHOWNORMAL
            idx = 0
            for i in range(start, end):
                hwnd = hwnds[i]
                col = idx % cols
                row = idx // cols
                x = col * (width + gap_x)
                y = row * (height + gap_y)
                # SW_SHOWNORMAL (1) - show without maximizing
                if HAS_WIN32:
                    user32.ShowWindow(hwnd, SW_SHOWNORMAL)
                    time.sleep(0.05)
                    user32.SetWindowPos(hwnd, HWND_TOP, x, y, width, height, SWP_SHOWWINDOW | SWP_FRAMECHANGED)
                idx += 1

        threading.Thread(target=do_switch, daemon=True).start()

    def _group_next(self):
        """Switch to next group (like AutoIt GROUPNEXT hotkey)."""
        hwnds = self._get_all_hwnds()
        if not hwnds:
            return
        try:
            cols = max(1, int(self.pos_entries.get('Cols', tk.Entry()).get() or '4'))
            rows = max(1, int(self.pos_entries.get('Rows', tk.Entry()).get() or '2'))
        except (ValueError, TypeError):
            cols, rows = 4, 2
        group_size = cols * rows
        total_groups = math.ceil(len(hwnds) / group_size)
        if total_groups == 0:
            return
        next_group = self.active_group + 1
        if next_group >= total_groups:
            next_group = 0
        self._switch_group(next_group)

    def _group_back(self):
        """Switch to previous group (like AutoIt GROUPBACK hotkey)."""
        hwnds = self._get_all_hwnds()
        if not hwnds:
            return
        try:
            cols = max(1, int(self.pos_entries.get('Cols', tk.Entry()).get() or '4'))
            rows = max(1, int(self.pos_entries.get('Rows', tk.Entry()).get() or '2'))
        except (ValueError, TypeError):
            cols, rows = 4, 2
        group_size = cols * rows
        total_groups = math.ceil(len(hwnds) / group_size)
        if total_groups == 0:
            return
        prev_group = self.active_group - 1
        if prev_group < 0:
            prev_group = total_groups - 1
        self._switch_group(prev_group)

    def _show_all_browsers(self):
        """Show all browsers positioned in grid using Pos tab settings."""
        hwnds = self._get_all_hwnds()
        if not hwnds:
            return

        try:
            cols = max(1, int(self.pos_entries.get('Cols', tk.Entry()).get() or '4'))
            rows = max(1, int(self.pos_entries.get('Rows', tk.Entry()).get() or '2'))
            width = int(self.pos_entries.get('Width', tk.Entry()).get() or '480')
            height = int(self.pos_entries.get('Height', tk.Entry()).get() or '540')
            gap_x = int(self.pos_entries.get('GapX', tk.Entry()).get() or '0')
            gap_y = int(self.pos_entries.get('GapY', tk.Entry()).get() or '0')
        except (ValueError, TypeError):
            cols, rows, width, height, gap_x, gap_y = 4, 2, 480, 540, 0, 0

        def do_show():
            for idx, hwnd in enumerate(hwnds):
                col = idx % cols
                row = (idx // cols) % rows
                x = col * (width + gap_x)
                y = row * (height + gap_y)
                if HAS_WIN32:
                    user32.ShowWindow(hwnd, SW_SHOWNORMAL)
                    time.sleep(0.05)
                    user32.SetWindowPos(hwnd, HWND_TOP, x, y, width, height, SWP_SHOWWINDOW | SWP_FRAMECHANGED)
                else:
                    show_window(hwnd)
                    time.sleep(0.05)

        # Reset active group
        self.active_group = -1
        for idx, btn in {**self.grp_btns, **self.pos_grp_btns}.items():
            letter = chr(65 + idx)
            btn.configure(text=letter, bg='#E0E0E0', fg='black',
                          activebackground='#D0D0D0', activeforeground='black',
                          relief='raised', highlightthickness=0)

        threading.Thread(target=do_show, daemon=True).start()

    # ── Resize ────────────────────────────────────────────────────────────────

    def _apply_resize(self):
        """Resize selected window to W/H."""
        hwnds = self._get_selected_hwnds()
        if not hwnds:
            return
        try:
            w = int(self.main_w_entry.get())
            h = int(self.main_h_entry.get())
        except ValueError:
            return
        self.cfg.set('MAIN', 'NavWidth', str(w))
        self.cfg.set('MAIN', 'NavHeight', str(h))
        save_config(self.cfg)

        def do_resize():
            for hwnd in hwnds:
                rect = ctypes.wintypes.RECT()
                user32.GetWindowRect(hwnd, ctypes.byref(rect))
                set_window_pos(hwnd, rect.left, rect.top, w, h)
        threading.Thread(target=do_resize, daemon=True).start()

    def _fix_all_sizes(self):
        """Fix all browser windows to W/H."""
        try:
            w = int(self.main_w_entry.get())
            h = int(self.main_h_entry.get())
        except ValueError:
            return
        def do_fix():
            for hwnd in self._get_all_hwnds():
                rect = ctypes.wintypes.RECT()
                user32.GetWindowRect(hwnd, ctypes.byref(rect))
                set_window_pos(hwnd, rect.left, rect.top, w, h)
        threading.Thread(target=do_fix, daemon=True).start()

    # ── TM Lite ───────────────────────────────────────────────────────────────

    def _tm_lite_all(self):
        """Send Alt+L to all browser windows (TM Lite toggle)."""
        def do_it():
            try:
                import keyboard as kb
            except ImportError:
                return
            for hwnd in self._get_all_hwnds():
                activate_window(hwnd)
                time.sleep(0.2)
                kb.send('alt+l')
                time.sleep(0.3)
        threading.Thread(target=do_it, daemon=True).start()

    # ── URL Opening ───────────────────────────────────────────────────────────

    def _open_url_all(self):
        """Open URL in all browsers using keyboard automation.
        Matches AutoIt: ClipPut, Activate, Ctrl+T, Ctrl+L, Ctrl+V, Enter."""
        url = self.main_url.get().strip()
        if not url:
            return
        self.stop_url_loop = False

        def do_open():
            set_clipboard(url)
            try:
                import keyboard as kb
            except ImportError:
                return
            hwnds = self._get_all_hwnds()
            for i, hwnd in enumerate(hwnds):
                if self.stop_url_loop:
                    break
                self._log(f'Opening URL: {i+1}/{len(hwnds)}')
                # Restore minimized windows first, then activate
                if HAS_WIN32:
                    user32.ShowWindow(hwnd, SW_SHOWNORMAL)
                    time.sleep(0.05)
                activate_window(hwnd)
                time.sleep(0.15)
                kb.send('ctrl+t')       # New tab
                time.sleep(0.3)
                kb.send('ctrl+l')       # Focus address bar
                time.sleep(0.2)
                kb.send('ctrl+v')       # Paste URL
                time.sleep(0.15)
                kb.send('enter')        # Navigate
                time.sleep(0.2)
            self.root.after(0, lambda: self._log('URL open complete'))

        threading.Thread(target=do_open, daemon=True).start()

    def _pos_open_url(self):
        url = self.pos_url.get().strip()
        if not url:
            return
        self.stop_url_loop = False
        old_url = self.main_url.get()
        self.main_url.delete(0, 'end')
        self.main_url.insert(0, url)
        self._open_url_all()
        self.main_url.delete(0, 'end')
        self.main_url.insert(0, old_url)

    def _stop_url(self):
        self.stop_url_loop = True

    # ── Position Windows ──────────────────────────────────────────────────────

    def _position_windows(self):
        try:
            cols = int(self.pos_entries['Cols'].get())
            rows = int(self.pos_entries['Rows'].get())
            width = int(self.pos_entries['Width'].get())
            height = int(self.pos_entries['Height'].get())
            gap_x = int(self.pos_entries['GapX'].get())
            gap_y = int(self.pos_entries['GapY'].get())
        except ValueError:
            return

        def do_pos():
            hwnds = self._get_all_hwnds()
            for i, hwnd in enumerate(hwnds):
                col = i % cols
                row = i // cols
                x = col * (width + gap_x)
                y = row * (height + gap_y)
                set_window_pos(hwnd, x, y, width, height)
        threading.Thread(target=do_pos, daemon=True).start()

    def _save_pos(self):
        for key, entry in self.pos_entries.items():
            self.cfg.set('POSITIONER', key, entry.get())
        self.cfg.set('POSITIONER', 'URL', self.pos_url.get())
        save_config(self.cfg)
        self._log('Positioner settings saved')

    # ── Distribte ─────────────────────────────────────────────────────────────

    def _dist_save(self):
        email = self.dist_email.get().strip()
        password = self.dist_password.get().strip()
        if not email or not password:
            self.dist_status.configure(text='Email and password required', fg='red')
            return
        self.cfg.set('DISTRIBTE', 'Email', email)
        self.cfg.set('DISTRIBTE', 'Password', password)
        save_config(self.cfg)

        def do_save():
            configs = find_distribte_configs()
            count = 0
            for path in configs:
                if write_distribte_config(path, email, password):
                    count += 1
            self.root.after(0, lambda: self.dist_status.configure(
                text=f'Saved to {count} config files', fg='green'))

        threading.Thread(target=do_save, daemon=True).start()
        self.dist_status.configure(text='Searching configs...', fg='blue')

    def _dist_clear(self):
        self.cfg.set('DISTRIBTE', 'Email', '')
        self.cfg.set('DISTRIBTE', 'Password', '')
        save_config(self.cfg)
        self.dist_email.delete(0, 'end')
        self.dist_password.delete(0, 'end')

        def do_clear():
            configs = find_distribte_configs()
            count = 0
            for path in configs:
                if write_distribte_config(path, '', '', enabled=False):
                    count += 1
            self.root.after(0, lambda: self.dist_status.configure(
                text=f'Cleared {count} config files', fg='orange'))

        threading.Thread(target=do_clear, daemon=True).start()

    # ── Discord ───────────────────────────────────────────────────────────────

    def _discord_send(self, channel):
        webhook = {
            'que': self.dc_quewebhook.get(),
            'prod': self.dc_prodwebhook.get(),
            'vf': self.dc_vfwebhook.get(),
        }.get(channel, '')

        profile = self.dc_profile.get().strip() or 'APM'
        message = self.dc_message.get('1.0', 'end').strip()
        content = f'**{profile}**: {message}' if message else f'**{profile}** queue update'

        def do_send():
            ok = discord_webhook_send_text(webhook, content, username=profile)
            self.root.after(0, lambda: self.dc_status.configure(
                text='Sent!' if ok else 'Failed', fg='green' if ok else 'red'))

        threading.Thread(target=do_send, daemon=True).start()

    def _discord_screenshot(self, channel):
        webhook = {
            'prod': self.dc_prodwebhook.get(),
            'que': self.dc_quewebhook.get(),
            'vf': getattr(self, 'dc_vfwebhook', None) and self.dc_vfwebhook.get() or '',
        }.get(channel, '')

        if not webhook:
            self.dc_status.configure(text=f'Error: No webhook URL set for {channel}', fg='red')
            return

        profile = self.dc_profile.get().strip() or 'APM'
        message = self.dc_message.get('1.0', 'end').strip()

        if not self.browsers:
            self.dc_status.configure(text='Error: No profiles in list', fg='red')
            return

        self.dc_status.configure(text=f'Preparing {len(self.browsers)} profiles...', fg='blue')

        def do_send():
            try:
                img_bytes = generate_profile_image(self.browsers)
            except Exception as e:
                self.root.after(0, lambda: self.dc_status.configure(
                    text=f'Image error: {e}', fg='red'))
                return
            if not img_bytes:
                self.root.after(0, lambda: self.dc_status.configure(
                    text='Failed to generate image (PIL missing?)', fg='red'))
                return
            self.root.after(0, lambda: self.dc_status.configure(
                text=f'Uploading {len(img_bytes)} bytes to Discord {channel}...', fg='blue'))
            try:
                url = discord_webhook_upload_image(webhook, img_bytes, content=message, username=profile)
            except Exception as e:
                self.root.after(0, lambda: self.dc_status.configure(
                    text=f'Upload error: {e}', fg='red'))
                return
            if not url:
                # Read debug log for error details
                err_detail = ''
                try:
                    log_p = os.path.join(os.path.dirname(sys.argv[0]) or '.', 'discord_debug.log')
                    if os.path.exists(log_p):
                        with open(log_p, 'r') as lf:
                            lines = lf.read().strip().split('\n')
                            err_detail = lines[-1] if lines else ''
                except Exception:
                    pass
                self.root.after(0, lambda: self.dc_status.configure(
                    text=f'v{VERSION} FAILED: {err_detail[:60]}', fg='red'))
                return

            # Log to sheets
            sheet_url = self.dc_sheeturl.get().strip()
            if sheet_url and url:
                sheet_name = 'PRODUCTION' if channel == 'prod' else 'QUEUE'
                rows = []
                for _, _, prof, tab in self.browsers:
                    rows.append({
                        'datetime': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                        'vaname': profile, 'profileid': prof,
                        'tab': tab, 'message': '', 'screenshot': url,
                    })
                log_to_google_sheets(sheet_url, sheet_name, rows)

            # Save locally
            folder = self.dc_folder.get().strip()
            if folder and img_bytes:
                date_folder = os.path.join(folder, datetime.now().strftime('%Y-%m-%d'), channel)
                os.makedirs(date_folder, exist_ok=True)
                ts = datetime.now().strftime('%H%M%S')
                with open(os.path.join(date_folder, f'{ts}.png'), 'wb') as f:
                    f.write(img_bytes)

            self.root.after(0, lambda: self.dc_status.configure(text=f'v{VERSION} Screenshot sent!', fg='green'))

        threading.Thread(target=do_send, daemon=True).start()

    def _save_screenshot(self):
        img_bytes = generate_profile_image(self.browsers)
        if not img_bytes:
            return
        folder = self.dc_folder.get().strip()
        if not folder:
            return
        os.makedirs(folder, exist_ok=True)
        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        with open(os.path.join(folder, f'screenshot_{ts}.png'), 'wb') as f:
            f.write(img_bytes)
        self.dc_status.configure(text='Screenshot saved!', fg='green')

    def _browse_screenshot_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.dc_folder.delete(0, 'end')
            self.dc_folder.insert(0, folder)

    def _save_discord(self):
        self.cfg.set('DISCORD', 'QueWebhook', self.dc_quewebhook.get())
        self.cfg.set('DISCORD', 'ProdWebhook', self.dc_prodwebhook.get())
        self.cfg.set('DISCORD', 'VfWebhook', self.dc_vfwebhook.get())
        self.cfg.set('DISCORD', 'ProfileName', self.dc_profile.get())
        self.cfg.set('DISCORD', 'ScreenshotFolder', self.dc_folder.get())
        self.cfg.set('DISCORD', 'SheetUrl', self.dc_sheeturl.get())
        save_config(self.cfg)
        self.dc_status.configure(text='Settings saved', fg='green')

    # ── Settings ──────────────────────────────────────────────────────────────

    def _save_settings(self):
        # Hotkeys
        for key, entry in self.hk_entries.items():
            self.cfg.set('HOTKEYS', key, entry.get())
        # Extra hotkeys
        if not self.cfg.has_section('HOTKEYS2'):
            self.cfg.add_section('HOTKEYS2')
        for i, (e1, e2, e3) in enumerate(self.ehk_entries):
            self.cfg.set('HOTKEYS2', f'EHK{i+1}-0', e1.get())
            self.cfg.set('HOTKEYS2', f'EHK{i+1}-1', e2.get())
            self.cfg.set('HOTKEYS2', f'EHK{i+1}-2', e3.get())
        self.cfg.set('MAIN', 'HotkeysToggleExtra', self.ehk_toggle_entry.get())
        # Options
        self.cfg.set('MAIN', 'AutoSorting', '1' if self.opt_autosorting.get() else '0')
        self.cfg.set('MAIN', 'InjectControls', '1' if self.opt_inject.get() else '0')
        self.cfg.set('MAIN', 'MinimizeOthers', '1' if self.opt_minimize_others.get() else '0')
        self.cfg.set('MAIN', 'AutoProfileSaver', '1' if self.opt_profile_saver.get() else '0')
        self.cfg.set('MAIN', 'CustomNavSize', '1' if self.opt_custom_nav.get() else '0')
        self.cfg.set('MAIN', 'NavWidth', self.set_navw.get())
        self.cfg.set('MAIN', 'NavHeight', self.set_navh.get())
        # AdsPower API key
        new_key = self.set_apikey.get().strip()
        self.cfg.set('ADSPOWER', 'ApiKey', new_key)
        self.api.api_key = new_key
        save_config(self.cfg)
        # Re-register hotkeys
        self._unregister_hotkeys()
        self._register_hotkeys()
        self._log('Settings saved')

    def _load_all_settings(self):
        """Load Discord entries from config."""
        self.dc_profile.delete(0, 'end')
        self.dc_profile.insert(0, self.cfg.get('DISCORD', 'ProfileName', fallback=''))

    # ── Toggles ───────────────────────────────────────────────────────────────

    def _toggle_hotkeys(self):
        self.hotkeys_on = self.hotkeys_var.get()
        self.cfg.set('MAIN', 'AllHotkeysON', '1' if self.hotkeys_on else '0')
        save_config(self.cfg)

    def _toggle_ontop(self):
        on = self.ontop_var.get()
        self.root.attributes('-topmost', on)
        self.cfg.set('MAIN', 'AlwaysOnTop', '1' if on else '0')
        save_config(self.cfg)

    # ── Hotkeys ───────────────────────────────────────────────────────────────

    def _register_hotkeys(self):
        try:
            import keyboard as kb
        except ImportError:
            return

        hotkey_map = {
            'FORWARD': lambda: self.root.after(0, self._hk_fwd),
            'BACKWARD': lambda: self.root.after(0, self._hk_bck),
            'TOP': lambda: self.root.after(0, self._move_top),
            'SORTTAB': lambda: self.root.after(0, lambda: self._sort_tree(1)),
            'SORTPROFILE': lambda: self.root.after(0, lambda: self._sort_tree(0)),
            'GROUPNEXT': lambda: self.root.after(0, self._group_next),
            'GROUPBACK': lambda: self.root.after(0, self._group_back),
        }

        for key_name, callback in hotkey_map.items():
            try:
                combo = self.cfg.get('HOTKEYS', key_name, fallback='').lower().replace('none', '').strip()
                if combo:
                    kb.add_hotkey(combo, callback, suppress=False)
                    self._log(f'Hotkey registered: {key_name}={combo}')
            except Exception as e:
                self._log(f'Hotkey {key_name} error: {e}')

    def _unregister_hotkeys(self):
        try:
            import keyboard as kb
            kb.unhook_all_hotkeys()
        except Exception:
            pass

    def _hk_fwd(self):
        if self.hotkeys_on:
            self._move_fwd()

    def _hk_bck(self):
        if self.hotkeys_on:
            self._move_back()

    # ── Polling ───────────────────────────────────────────────────────────────

    def _start_polling(self):
        """Start background polling for browser windows."""
        def poll_loop():
            while self.running:
                self._get_browsers()
                time.sleep(1.5)  # Smoother than AutoIt's 1000ms

        threading.Thread(target=poll_loop, daemon=True).start()

        # Refresh AdsPower cache periodically
        def cache_loop():
            while self.running:
                self.api.get_user_list(force=True)
                time.sleep(30)
        threading.Thread(target=cache_loop, daemon=True).start()

        if HAS_WS:
            threading.Thread(target=self._chrome_profile_monitor, daemon=True).start()
            if self.cfg.get('MAIN', 'AutoProfileSaver', fallback='0') == '1':
                self._log('Chrome profile auto-saver is ON')
        else:
            self._log('websocket-client not installed, Chrome profile auto-saver unavailable')

    def _chrome_profile_monitor(self):
        """Auto-click 'Continue as' button via Chrome DevTools Protocol."""
        clicked_targets = set()
        port_cache = {}  # pid -> debug_port
        while self.running:
            if not self.opt_profile_saver.get():
                time.sleep(3)
                continue
            try:
                sun_pids = {p for p, v in self.sunpid_cache.items() if v}
                for pid in sun_pids:
                    if pid in port_cache:
                        port = port_cache[pid]
                    else:
                        cmdline = self.cmdline_cache.get(pid, '')
                        if not cmdline:
                            cmdline = get_process_cmdline(pid)
                        m = re.search(r'--remote-debugging-port=(\d+)', cmdline)
                        if not m:
                            continue
                        port = int(m.group(1))
                        port_cache[pid] = port
                    profile = self.pid_profile_cache.get(pid, f'PID:{pid}')
                    self._try_click_signin(port, profile, clicked_targets)
            except Exception:
                pass
            if len(clicked_targets) > 1000:
                clicked_targets.clear()
            if len(port_cache) > 500:
                port_cache.clear()
            time.sleep(5)

    def _try_click_signin(self, debug_port, serial, clicked_targets):
        """Check browser for signin popup and click accept via CDP."""
        label = f'#{serial}' if serial else f'port {debug_port}'
        try:
            # Method 1: Direct target list - check /json for signin page
            url = f'http://127.0.0.1:{debug_port}/json'
            req = Request(url)
            with urlopen(req, timeout=2) as r:
                targets = json.loads(r.read().decode())
            for target in targets:
                t_url = target.get('url', '')
                tid = target.get('id', '')
                if 'signin' in t_url.lower() and tid not in clicked_targets:
                    ws_url = target.get('webSocketDebuggerUrl', '')
                    if ws_url and self._cdp_click_accept(ws_url):
                        clicked_targets.add(tid)
                        self._log(f'Chrome profile saved for {label}')
                        return

            # Method 2: Browser-level CDP to find top-chrome WebUI targets
            url2 = f'http://127.0.0.1:{debug_port}/json/version'
            req2 = Request(url2)
            with urlopen(req2, timeout=2) as r2:
                version = json.loads(r2.read().decode())
            browser_ws = version.get('webSocketDebuggerUrl', '')
            if not browser_ws:
                return
            ws = _ws_mod.create_connection(browser_ws, timeout=5)
            ws.send(json.dumps({'id': 1, 'method': 'Target.getTargets'}))
            result = self._cdp_read_response(ws, 1)
            if not result:
                ws.close()
                return
            target_infos = result.get('result', {}).get('targetInfos', [])
            for t in target_infos:
                t_url = t.get('url', '')
                tid = t.get('targetId', '')
                if 'signin' in t_url.lower() and tid not in clicked_targets:
                    # Attach to target and click
                    ws.send(json.dumps({
                        'id': 2,
                        'method': 'Target.attachToTarget',
                        'params': {'targetId': tid, 'flatten': True}
                    }))
                    attach = self._cdp_read_response(ws, 2)
                    if not attach:
                        continue
                    sid = attach.get('result', {}).get('sessionId', '')
                    if not sid:
                        continue
                    ws.send(json.dumps({
                        'id': 3,
                        'sessionId': sid,
                        'method': 'Runtime.evaluate',
                        'params': {
                            'expression': '(function(){var b=document.querySelector("#accept-button");if(b){b.click();return"ok"}return"nf"})()'
                        }
                    }))
                    ev = self._cdp_read_response(ws, 3)
                    if ev:
                        val = ev.get('result', {}).get('result', {}).get('value', '')
                        if val == 'ok':
                            clicked_targets.add(tid)
                            self._log(f'Chrome profile saved for {label}')
                    break
            ws.close()
        except Exception:
            pass

    def _cdp_click_accept(self, ws_url):
        """Connect to a target WebSocket and click #accept-button."""
        try:
            ws = _ws_mod.create_connection(ws_url, timeout=3)
            ws.send(json.dumps({
                'id': 1,
                'method': 'Runtime.evaluate',
                'params': {
                    'expression': '(function(){var b=document.querySelector("#accept-button");if(b){b.click();return"ok"}return"nf"})()'
                }
            }))
            result = json.loads(ws.recv())
            ws.close()
            return result.get('result', {}).get('result', {}).get('value', '') == 'ok'
        except Exception:
            return False

    def _cdp_read_response(self, ws, msg_id):
        """Read WebSocket messages until we get the response matching msg_id."""
        for _ in range(10):
            try:
                msg = json.loads(ws.recv())
                if msg.get('id') == msg_id:
                    return msg
            except Exception:
                return None
        return None

    # ── Debug Log ─────────────────────────────────────────────────────────────

    def _show_debug_log(self):
        win = tk.Toplevel(self.root)
        win.title('Debug Log')
        win.geometry('500x300')
        text = tk.Text(win, font=('Consolas', 8), bg='black', fg='#00ff00')
        text.pack(fill='both', expand=True)
        text.insert('end', '\n'.join(self.debug_log[-200:]))
        text.see('end')

    # ── Close ─────────────────────────────────────────────────────────────────

    def _on_close(self):
        try:
            geo = self.root.geometry()
            m = re.match(r'(\d+)x(\d+)\+(-?\d+)\+(-?\d+)', geo)
            if m:
                self.cfg.set('MAIN', 'GUIW', m.group(1))
                self.cfg.set('MAIN', 'GUIH', m.group(2))
                self.cfg.set('MAIN', 'GUIX', m.group(3))
                self.cfg.set('MAIN', 'GUIY', m.group(4))
                save_config(self.cfg)
        except Exception:
            pass
        self.running = False
        self._unregister_hotkeys()
        self.root.destroy()

    def run(self):
        self.root.mainloop()


# ─── Entry ────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    if not requests:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror('APM', 'Missing "requests" package.\npip install requests')
        sys.exit(1)

    app = APMApp()
    app.run()
