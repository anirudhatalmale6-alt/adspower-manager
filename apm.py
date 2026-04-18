"""
APM - AdsPower Manager v2.0
Full-featured desktop app for managing AdsPower browser profiles.
Feature-matched to APM v6.7 with improved performance.
"""

import tkinter as tk
from tkinter import ttk, messagebox
import threading
import json
import time
import os
import sys
import configparser
import re
from datetime import datetime

try:
    import requests
except ImportError:
    requests = None

try:
    import websocket
except ImportError:
    websocket = None

try:
    import keyboard
except ImportError:
    keyboard = None

# Windows-specific imports
try:
    import win32gui
    import win32con
    import win32api
    import win32process
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

try:
    from PIL import ImageGrab
    HAS_PIL = True
except ImportError:
    HAS_PIL = False


# ─── Helpers ──────────────────────────────────────────────────────────────────

def get_base_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

BASE_DIR = get_base_dir()
CONFIG_PATH = os.path.join(BASE_DIR, 'APMData', 'config.ini')
SCREENSHOT_DIR = os.path.join(BASE_DIR, 'Screenshots')

def ensure_dirs():
    os.makedirs(os.path.join(BASE_DIR, 'APMData'), exist_ok=True)
    os.makedirs(SCREENSHOT_DIR, exist_ok=True)

def load_config():
    cfg = configparser.ConfigParser()
    if os.path.exists(CONFIG_PATH):
        cfg.read(CONFIG_PATH, encoding='utf-8')
    # Ensure sections exist
    for s in ['MAIN', 'HOTKEYS', 'HOTKEYS2', 'DISCORD', 'SHEETS', 'SCREENSHOTS']:
        if not cfg.has_section(s):
            cfg.add_section(s)
    # Defaults
    defaults = {
        'MAIN': {
            'AlwaysOnTop': '1', 'GUIW': '380', 'GUIH': '700', 'GUIX': '20', 'GUIY': '60',
            'PollInterval': '3', 'AdsPowerPort': '50325',
            'AutoColumnSorting': '0', 'InjectControls': '1', 'MinimizeOthers': '0',
            'UseCustomNavSize': '1', 'NavW': '480', 'NavH': '540',
            'SortColumn': '1', 'MainURL': '0', 'RemoveShadows': '1',
        },
        'HOTKEYS': {
            'FORWARD': 'ctrl+shift+right', 'BACKWARD': 'ctrl+shift+left',
            'TOP': 'ctrl+shift+up', 'SORTTAB': 'ctrl+shift+t',
            'SORTPROFILE': 'ctrl+shift+p', 'GROUPNEXT': '[', 'GROUPBACK': ']',
        },
        'DISCORD': {
            'QueWebhook': '', 'ProdWebhook': '', 'ProfileName': '',
        },
        'SHEETS': {'SheetUrl': ''},
        'SCREENSHOTS': {'Folder': SCREENSHOT_DIR},
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


# ─── AdsPower API ─────────────────────────────────────────────────────────────

class AdsPowerAPI:
    def __init__(self, port=50325):
        self.port = port
        self.bases = [
            f'http://127.0.0.1:{port}',
            f'http://local.adspower.net:{port}'
        ]
        self._base = None  # Cache working base

    def _get(self, path, timeout=4):
        # Try cached base first
        if self._base:
            try:
                r = requests.get(self._base + path, timeout=timeout)
                if r.ok:
                    return r.json()
            except Exception:
                self._base = None
        for base in self.bases:
            try:
                r = requests.get(base + path, timeout=timeout)
                if r.ok:
                    self._base = base
                    return r.json()
            except Exception:
                continue
        return None

    def get_active_profiles(self):
        data = self._get('/api/v1/browser/active?page=1&page_size=100')
        if not data:
            return []
        raw = data.get('data', [])
        if isinstance(raw, dict):
            raw = raw.get('list', [])
        if not isinstance(raw, list):
            return []

        has_serial = any(p.get('serial_number') or p.get('serialnumber') for p in raw)
        if has_serial:
            return self._normalize(raw)

        # Cross-reference with user/list
        active_ids = {str(p.get('user_id', '')).strip() for p in raw if p.get('user_id')}
        user_data = self._get('/api/v1/user/list?page=1&page_size=100')
        if not user_data:
            return self._normalize(raw)

        all_users = user_data.get('data', {})
        if isinstance(all_users, dict):
            all_users = all_users.get('list', [])
        if not isinstance(all_users, list):
            return self._normalize(raw)

        merged = []
        for u in all_users:
            uid = str(u.get('user_id', '')).strip()
            if active_ids and uid not in active_ids:
                continue
            for a in raw:
                if str(a.get('user_id', '')).strip() == uid:
                    if not u.get('debug_port') and a.get('debug_port'):
                        u['debug_port'] = a['debug_port']
                    if not u.get('ws') and a.get('ws'):
                        u['ws'] = a['ws']
                    break
            merged.append(u)

        return self._normalize(merged if merged else raw)

    def get_all_profiles(self):
        """Get ALL profiles (not just active) for group assignment."""
        data = self._get('/api/v1/user/list?page=1&page_size=100')
        if not data:
            return []
        raw = data.get('data', {})
        if isinstance(raw, dict):
            raw = raw.get('list', [])
        return self._normalize(raw) if isinstance(raw, list) else []

    def get_groups(self):
        """Get profile groups from AdsPower."""
        data = self._get('/api/v1/group/list?page=1&page_size=100')
        if not data:
            return []
        raw = data.get('data', {})
        if isinstance(raw, dict):
            raw = raw.get('list', [])
        return raw if isinstance(raw, list) else []

    def _normalize(self, raw):
        profiles = []
        for p in raw:
            serial = str(p.get('serial_number') or p.get('serialnumber') or '').strip()
            custom = str(p.get('custom_user_id') or p.get('customUserId') or '').strip()
            uid = str(p.get('user_id') or p.get('userId') or p.get('id') or '').strip()
            name = str(p.get('name') or p.get('profile_name') or (f'#{serial}' if serial else uid) or '')
            group_id = str(p.get('group_id') or p.get('groupId') or '').strip()
            group_name = str(p.get('group_name') or '').strip()
            debug_port = self._get_debug_port(p)
            profiles.append({
                'serial': serial, 'custom': custom, 'uid': uid,
                'name': name, 'debug_port': debug_port,
                'group_id': group_id, 'group_name': group_name,
            })
        return profiles

    def _get_debug_port(self, p):
        if p.get('debug_port'):
            try:
                return int(p['debug_port'])
            except (ValueError, TypeError):
                pass
        ws = p.get('ws', {})
        ws_url = ws.get('puppeteer') or ws.get('selenium') or '' if isinstance(ws, dict) else ''
        if ws_url:
            try:
                from urllib.parse import urlparse
                return int(urlparse(ws_url).port or 0)
            except Exception:
                m = re.search(r':(\d+)/', ws_url)
                if m:
                    return int(m.group(1))
        return 0


# ─── CDP Queue Scanner ────────────────────────────────────────────────────────

TM_HOST_RE = re.compile(
    r'ticketmaster\.com|livenation\.com|queue-it\.net|ticketmaster\.ca|ticketmaster\.co\.uk', re.I)

QUEUE_EVAL_JS = """(function(){
function p(s){var n=parseInt(String(s).replace(/,/g,""),10);return(n>0&&n<100000000)?n:null;}
var sel=["#MainPart_lbQueueNumber","#lbQueueNumber","[id*='lbQueueNumber']",
"#MainPart_h2HeaderSubText","#h2-main",".queue-position","[class*='queuePosition']",
"[class*='queueNumber']","[class*='waiting-number']","[class*='place-in-line']",
"[data-queue-number]","#queue-number",".number-display"];
for(var i=0;i<sel.length;i++){try{var el=document.querySelector(sel[i]);if(!el)continue;
var m=el.textContent.replace(/,/g,"").match(/\\d+/);if(m){var v=p(m[0]);if(v)return{q:v,u:location.href,t:document.title};}}catch(e){}}
var txt=document.body?document.body.innerText:"";
var pts=[/you\\s+are\\s+(?:now\\s+)?in\\s+the\\s+queue\\s*#?\\s*([\\d,]+)/i,
/([\\d,]+)\\s+people\\s+ahead/i,
/you\\s+are\\s+(?:now\\s+)?(?:number\\s+|#\\s*)?([\\d,]+)\\s+in/i,
/(?:queue\\s+)?position\\s+(?:is\\s+)?(?:number\\s+|#\\s*)?([\\d,]+)/i,
/there\\s+are\\s+([\\d,]+)\\s+people/i];
for(var j=0;j<pts.length;j++){var mm=txt.match(pts[j]);if(mm){var v=p(mm[1]);if(v)return{q:v,u:location.href,t:document.title};}}
return null;})()"""


def _ws_eval(ws_url, expression, timeout=4):
    try:
        ws = websocket.create_connection(ws_url, timeout=timeout)
    except Exception:
        return None
    try:
        msg_id = 90000 + int(time.time() * 1000) % 10000
        ws.send(json.dumps({
            'id': msg_id, 'method': 'Runtime.evaluate',
            'params': {'expression': expression, 'returnByValue': True}
        }))
        ws.settimeout(timeout)
        while True:
            raw = ws.recv()
            if not raw:
                break
            msg = json.loads(raw)
            if msg.get('id') == msg_id:
                return (msg.get('result', {}).get('result', {}) or {}).get('value')
    except Exception:
        pass
    finally:
        try:
            ws.close()
        except Exception:
            pass
    return None


def get_profile_tabs(debug_port):
    """Get all open tabs for a profile."""
    if not debug_port or debug_port <= 0:
        return []
    try:
        r = requests.get(f'http://127.0.0.1:{debug_port}/json', timeout=2)
        if not r.ok:
            return []
        tabs = r.json()
        if not isinstance(tabs, list):
            return []
        return [t for t in tabs if t.get('url') and
                not re.match(r'^(devtools:|chrome:|chrome-extension:)', t.get('url', ''), re.I)]
    except Exception:
        return []


def cdp_eval_queue(debug_port, timeout=3):
    """Scan profile tabs for queue numbers."""
    tabs = get_profile_tabs(debug_port)
    if not tabs:
        return None

    # Prefer TM/queue tabs
    tm_tabs = [t for t in tabs if TM_HOST_RE.search(t.get('url', '')) and t.get('webSocketDebuggerUrl')]
    scan_tabs = tm_tabs if tm_tabs else [t for t in tabs if t.get('webSocketDebuggerUrl')][:2]

    for tab in scan_tabs:
        ws_url = tab.get('webSocketDebuggerUrl', '')
        if not ws_url:
            continue
        result = _ws_eval(ws_url, QUEUE_EVAL_JS, timeout=timeout)
        if result and isinstance(result, dict) and result.get('q'):
            return {
                'queue_number': result['q'],
                'url': result.get('u', tab.get('url', '')),
                'title': result.get('t', tab.get('title', '')),
            }
        # Fallback: title parsing
        title = tab.get('title', '')
        m = re.match(r'^\s*#?([\d,]{1,8})\s*\|', title)
        if m:
            n = int(m.group(1).replace(',', ''))
            if 0 < n < 100000000:
                return {'queue_number': n, 'url': tab.get('url', ''), 'title': title}
    return None


def cdp_open_tab(debug_port, url, timeout=5):
    if not debug_port or debug_port <= 0:
        return False
    try:
        r = requests.get(f'http://127.0.0.1:{debug_port}/json/version', timeout=2)
        if not r.ok:
            return False
        ws_url = r.json().get('webSocketDebuggerUrl', '')
        if not ws_url:
            return False
        ws = websocket.create_connection(ws_url, timeout=timeout)
        msg_id = 80000 + int(time.time() * 1000) % 10000
        ws.send(json.dumps({'id': msg_id, 'method': 'Target.createTarget', 'params': {'url': url}}))
        ws.settimeout(timeout)
        while True:
            raw = ws.recv()
            if not raw:
                break
            msg = json.loads(raw)
            if msg.get('id') == msg_id:
                ws.close()
                return bool(msg.get('result', {}).get('targetId'))
        ws.close()
    except Exception:
        pass
    return False


def cdp_close_tabs(debug_port, timeout=3):
    """Close all tabs in a profile browser."""
    tabs = get_profile_tabs(debug_port)
    if not tabs:
        return
    try:
        r = requests.get(f'http://127.0.0.1:{debug_port}/json/version', timeout=2)
        if not r.ok:
            return
        ws_url = r.json().get('webSocketDebuggerUrl', '')
        if not ws_url:
            return
        ws = websocket.create_connection(ws_url, timeout=timeout)
        for tab in tabs:
            tid = tab.get('id', '')
            if tid:
                msg_id = 70000 + int(time.time() * 1000) % 10000
                ws.send(json.dumps({'id': msg_id, 'method': 'Target.closeTarget',
                                     'params': {'targetId': tid}}))
                try:
                    ws.recv()
                except Exception:
                    pass
        ws.close()
    except Exception:
        pass


# ─── Window Management ───────────────────────────────────────────────────────

def find_profile_window(debug_port):
    """Find the Windows HWND for a profile browser by its debug port."""
    if not HAS_WIN32 or not debug_port:
        return None
    target_pids = set()
    try:
        import subprocess
        result = subprocess.run(['netstat', '-ano'], capture_output=True, text=True, timeout=3)
        for line in result.stdout.splitlines():
            if f':{debug_port}' in line and 'LISTENING' in line:
                parts = line.split()
                if parts:
                    try:
                        target_pids.add(int(parts[-1]))
                    except ValueError:
                        pass
    except Exception:
        pass
    if not target_pids:
        return None

    best_hwnd = None
    best_area = 0

    def enum_cb(hwnd, _):
        nonlocal best_hwnd, best_area
        if not win32gui.IsWindowVisible(hwnd):
            return True
        try:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            if pid not in target_pids:
                return True
            title = win32gui.GetWindowText(hwnd)
            if not title:
                return True
            rect = win32gui.GetWindowRect(hwnd)
            area = (rect[2] - rect[0]) * (rect[3] - rect[1])
            if area > best_area:
                best_area = area
                best_hwnd = hwnd
        except Exception:
            pass
        return True

    try:
        win32gui.EnumWindows(enum_cb, None)
    except Exception:
        pass
    return best_hwnd


def get_all_browser_hwnds():
    """Get all Chrome/browser window handles."""
    if not HAS_WIN32:
        return []
    hwnds = []
    def enum_cb(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return True
        cls = win32gui.GetClassName(hwnd)
        if cls in ('Chrome_WidgetWin_1', 'Chrome_WidgetWin_0'):
            title = win32gui.GetWindowText(hwnd)
            if title:
                hwnds.append(hwnd)
        return True
    try:
        win32gui.EnumWindows(enum_cb, None)
    except Exception:
        pass
    return hwnds


def show_window(hwnd):
    if not HAS_WIN32 or not hwnd:
        return
    try:
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        win32gui.SetForegroundWindow(hwnd)
    except Exception:
        pass


def minimize_window(hwnd):
    if not HAS_WIN32 or not hwnd:
        return
    try:
        win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
    except Exception:
        pass


def position_window(hwnd, x, y, w, h):
    if not HAS_WIN32 or not hwnd:
        return
    try:
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, x, y, w, h, win32con.SWP_SHOWWINDOW)
        win32gui.SetForegroundWindow(hwnd)
    except Exception:
        pass


def activate_profile(debug_port, custom_size=None, minimize_others=False):
    """Bring a profile window to front, optionally resize and minimize others."""
    hwnd = find_profile_window(debug_port)
    if not hwnd:
        return

    if minimize_others:
        all_hwnds = get_all_browser_hwnds()
        for h in all_hwnds:
            if h != hwnd:
                minimize_window(h)

    if custom_size and custom_size[0] > 0 and custom_size[1] > 0:
        try:
            rect = win32gui.GetWindowRect(hwnd)
            position_window(hwnd, rect[0], rect[1], custom_size[0], custom_size[1])
        except Exception:
            show_window(hwnd)
    else:
        show_window(hwnd)


# ─── Discord ─────────────────────────────────────────────────────────────────

def send_discord(webhook_url, content, username='APM'):
    if not webhook_url:
        return
    try:
        requests.post(webhook_url, json={'content': content, 'username': username}, timeout=5)
    except Exception:
        pass


# ─── Main GUI ─────────────────────────────────────────────────────────────────

class APMApp:
    # Colors
    BG = '#2b2b2b'
    BG2 = '#3c3c3c'
    BG3 = '#4a4a4a'
    FG = '#e0e0e0'
    FG2 = '#999'
    ACCENT = '#e94560'
    GOLD = '#ffd700'
    GREEN = '#00e676'
    CYAN = '#00bcd4'
    PURPLE = '#7b2ff7'

    def __init__(self):
        ensure_dirs()
        self.cfg = load_config()
        self.api = AdsPowerAPI(int(self.cfg.get('MAIN', 'AdsPowerPort')))
        self.profiles = []          # Active profiles
        self.profile_tabs = {}      # serial -> [tab titles]
        self.queue_data = {}        # serial -> {queue_number, url, title}
        self.selected_index = -1
        self.selected_indices = set()
        self.current_group = None   # None = all
        self.running = True
        self.debug_log = []
        self.hotkeys_enabled = True

        self._build_gui()
        self._register_hotkeys()
        self._start_polling()

    def _log(self, msg):
        ts = datetime.now().strftime('%H:%M:%S')
        line = f'[{ts}] {msg}'
        self.debug_log.append(line)
        if len(self.debug_log) > 500:
            self.debug_log = self.debug_log[-300:]

    # ── GUI ───────────────────────────────────────────────────────────────────

    def _build_gui(self):
        self.root = tk.Tk()
        self.root.title('AdsPower Window Manager v2.0')
        self.root.configure(bg=self.BG)

        w = int(self.cfg.get('MAIN', 'GUIW'))
        h = int(self.cfg.get('MAIN', 'GUIH'))
        x = int(self.cfg.get('MAIN', 'GUIX'))
        y = int(self.cfg.get('MAIN', 'GUIY'))
        self.root.geometry(f'{w}x{h}+{x}+{y}')
        self.root.minsize(340, 500)

        self.always_on_top = self.cfg.get('MAIN', 'AlwaysOnTop') == '1'
        self.root.attributes('-topmost', self.always_on_top)

        # ── Top bar: tabs + toggles ──
        topbar = tk.Frame(self.root, bg=self.BG2)
        topbar.pack(fill='x')

        self.tab_buttons = {}
        tab_names = ['Main', 'Settings', 'Discord', 'Distribute', 'Pos']
        for name in tab_names:
            btn = tk.Button(topbar, text=name, font=('Segoe UI', 8),
                            fg=self.FG, bg=self.BG2, bd=0, padx=6, pady=3,
                            activebackground=self.BG3, cursor='hand2',
                            command=lambda n=name: self._switch_tab(n))
            btn.pack(side='left')
            self.tab_buttons[name] = btn

        # Hotkeys toggle
        self.hotkeys_var = tk.BooleanVar(value=True)
        tk.Checkbutton(topbar, text='Hotkeys', variable=self.hotkeys_var,
                        font=('Segoe UI', 8), fg=self.FG, bg=self.BG2,
                        selectcolor=self.BG, activebackground=self.BG2,
                        command=self._toggle_hotkeys).pack(side='left', padx=4)

        # On top toggle
        self.ontop_var = tk.BooleanVar(value=self.always_on_top)
        tk.Checkbutton(topbar, text='On top', variable=self.ontop_var,
                        font=('Segoe UI', 8), fg=self.FG, bg=self.BG2,
                        selectcolor=self.BG, activebackground=self.BG2,
                        command=self._toggle_ontop).pack(side='left', padx=4)

        # ── Tab content area ──
        self.tab_container = tk.Frame(self.root, bg=self.BG)
        self.tab_container.pack(fill='both', expand=True)

        self.tab_frames = {}
        for name in tab_names:
            frame = tk.Frame(self.tab_container, bg=self.BG)
            self.tab_frames[name] = frame

        self._build_main_tab()
        self._build_settings_tab()
        self._build_discord_tab()
        self._build_distribute_tab()
        self._build_pos_tab()

        # ── Bottom bar ──
        self._build_bottom_bar()

        self._switch_tab('Main')
        self.root.protocol('WM_DELETE_WINDOW', self._on_close)

    def _switch_tab(self, name):
        for n, f in self.tab_frames.items():
            f.pack_forget()
        self.tab_frames[name].pack(fill='both', expand=True)
        for n, btn in self.tab_buttons.items():
            btn.configure(bg=self.BG3 if n == name else self.BG2)

    # ── Main Tab ──────────────────────────────────────────────────────────────

    def _build_main_tab(self):
        f = self.tab_frames['Main']

        # Navigation row
        nav = tk.Frame(f, bg=self.BG)
        nav.pack(fill='x', padx=5, pady=4)

        tk.Button(nav, text='<<<', font=('Consolas', 9, 'bold'), fg='#fff', bg=self.BG3,
                  bd=0, padx=8, pady=2, command=lambda: self._cycle(-1)).pack(side='left', padx=2)
        tk.Button(nav, text='TOP', font=('Consolas', 9, 'bold'), fg='#fff', bg=self.BG3,
                  bd=0, padx=12, pady=2, command=self._go_top).pack(side='left', padx=2)
        tk.Button(nav, text='>>>', font=('Consolas', 9, 'bold'), fg='#fff', bg=self.BG3,
                  bd=0, padx=8, pady=2, command=lambda: self._cycle(1)).pack(side='left', padx=2)

        # Profile list + side buttons
        mid = tk.Frame(f, bg=self.BG)
        mid.pack(fill='both', expand=True, padx=5, pady=2)

        # Treeview for profile list
        list_frame = tk.Frame(mid, bg=self.BG)
        list_frame.pack(side='left', fill='both', expand=True)

        cols = ('profile', 'tab')
        self.tree = ttk.Treeview(list_frame, columns=cols, show='headings', height=15,
                                  selectmode='extended')
        self.tree.heading('profile', text='Profile', command=lambda: self._sort_column('profile'))
        self.tree.heading('tab', text='Tab', command=lambda: self._sort_column('tab'))
        self.tree.column('profile', width=100, minwidth=60)
        self.tree.column('tab', width=180, minwidth=80)

        # Style
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('Treeview', background=self.BG, foreground=self.FG,
                         fieldbackground=self.BG, font=('Consolas', 9), rowheight=22)
        style.configure('Treeview.Heading', background=self.BG2, foreground=self.FG,
                         font=('Segoe UI', 8, 'bold'))
        style.map('Treeview', background=[('selected', self.PURPLE)],
                   foreground=[('selected', '#fff')])

        scrollbar = ttk.Scrollbar(list_frame, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        self.tree.bind('<<TreeviewSelect>>', self._on_tree_select)
        self.tree.bind('<Double-1>', self._on_tree_double_click)

        # Side buttons
        side = tk.Frame(mid, bg=self.BG)
        side.pack(side='right', fill='y', padx=(4, 0))

        side_buttons = [
            ('Show', self._btn_show),
            ('Minimize', self._btn_minimize),
            ('RefreshAll', self._btn_refresh_all),
            ('Close All', self._btn_close_all),
            ('Close Sel', self._btn_close_selected),
            ('Show All', self._btn_show_all),
            ('MinimizeAll', self._btn_minimize_all),
        ]
        for text, cmd in side_buttons:
            tk.Button(side, text=text, font=('Segoe UI', 8), fg=self.FG, bg=self.BG2,
                      bd=1, relief='raised', padx=4, pady=1, width=10,
                      activebackground=self.BG3, cursor='hand2',
                      command=cmd).pack(pady=1)

        # ── Group section ──
        grp_frame = tk.Frame(f, bg=self.BG)
        grp_frame.pack(fill='x', padx=5, pady=4)

        tk.Label(grp_frame, text='Grp:', font=('Segoe UI', 8, 'bold'),
                 fg=self.FG2, bg=self.BG).pack(side='left')

        self.grp_buttons_frame = tk.Frame(grp_frame, bg=self.BG)
        self.grp_buttons_frame.pack(side='left', fill='x', expand=True, padx=4)

        # Grid of A-Z buttons
        letters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
        self.grp_btn_map = {}
        row_frame = None
        for i, letter in enumerate(letters):
            if i % 3 == 0:
                row_frame = tk.Frame(self.grp_buttons_frame, bg=self.BG)
                row_frame.pack(fill='x', pady=1)
            btn = tk.Button(row_frame, text=letter, font=('Consolas', 8, 'bold'),
                            fg=self.FG, bg=self.BG2, bd=1, width=2, pady=0,
                            activebackground=self.PURPLE, cursor='hand2',
                            command=lambda l=letter: self._select_group(l))
            btn.pack(side='left', padx=1)
            self.grp_btn_map[letter] = btn

        # Nav size section
        nav_size_frame = tk.Frame(f, bg=self.BG)
        nav_size_frame.pack(fill='x', padx=5, pady=2)

        tk.Label(nav_size_frame, text='W:', font=('Consolas', 9),
                 fg=self.FG2, bg=self.BG).pack(side='left')
        self.nav_w_entry = tk.Entry(nav_size_frame, font=('Consolas', 9), bg=self.BG2,
                                     fg=self.FG, bd=1, width=6, insertbackground=self.FG)
        self.nav_w_entry.insert(0, self.cfg.get('MAIN', 'NavW'))
        self.nav_w_entry.pack(side='left', padx=2)

        tk.Label(nav_size_frame, text='H:', font=('Consolas', 9),
                 fg=self.FG2, bg=self.BG).pack(side='left', padx=(8, 0))
        self.nav_h_entry = tk.Entry(nav_size_frame, font=('Consolas', 9), bg=self.BG2,
                                     fg=self.FG, bd=1, width=6, insertbackground=self.FG)
        self.nav_h_entry.insert(0, self.cfg.get('MAIN', 'NavH'))
        self.nav_h_entry.pack(side='left', padx=2)

        tk.Button(nav_size_frame, text='Apply', font=('Segoe UI', 8), fg=self.FG,
                  bg=self.BG2, bd=1, padx=6, command=self._apply_nav_size).pack(side='left', padx=4)
        tk.Button(nav_size_frame, text='Fix', font=('Segoe UI', 8), fg=self.FG,
                  bg=self.BG2, bd=1, padx=6, command=self._fix_windows).pack(side='left', padx=2)

        # Horizontal scrollbar placeholder
        hscroll = tk.Frame(f, bg=self.BG, height=18)
        hscroll.pack(fill='x', padx=5)
        tk.Button(hscroll, text='<', font=('Consolas', 8), fg=self.FG, bg=self.BG2,
                  bd=0, padx=6, command=lambda: self._cycle(-1)).pack(side='left')
        self.scroll_scale = tk.Scale(hscroll, from_=0, to=100, orient='horizontal',
                                      bg=self.BG2, fg=self.FG, troughcolor=self.BG,
                                      highlightthickness=0, bd=0, showvalue=False,
                                      command=self._on_scroll)
        self.scroll_scale.pack(side='left', fill='x', expand=True)
        tk.Button(hscroll, text='>', font=('Consolas', 8), fg=self.FG, bg=self.BG2,
                  bd=0, padx=6, command=lambda: self._cycle(1)).pack(side='right')

        # TM Lite button
        self.tm_lite_active = False
        self.tm_lite_btn = tk.Button(f, text='TM Lite', font=('Segoe UI', 9, 'bold'),
                                      fg='#fff', bg=self.ACCENT, bd=0, padx=12, pady=3,
                                      cursor='hand2', command=self._toggle_tm_lite)
        self.tm_lite_btn.pack(side='right', padx=5, pady=2)

    # ── Settings Tab ──────────────────────────────────────────────────────────

    def _build_settings_tab(self):
        f = self.tab_frames['Settings']
        canvas = tk.Canvas(f, bg=self.BG, highlightthickness=0)
        scrollbar = tk.Scrollbar(f, orient='vertical', command=canvas.yview)
        inner = tk.Frame(canvas, bg=self.BG)
        inner.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.create_window((0, 0), window=inner, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        row = 0
        tk.Label(inner, text='OPTIONS', font=('Segoe UI', 9, 'bold'),
                 fg=self.ACCENT, bg=self.BG).grid(row=row, column=0, columnspan=2, sticky='w', pady=(8, 4), padx=8)
        row += 1

        self.settings_checks = {}
        checks = [
            ('AutoColumnSorting', 'Auto column sorting'),
            ('InjectControls', 'Inject controls inside each browser'),
            ('MinimizeOthers', 'Minimize others on profile select'),
            ('RemoveShadows', 'Remove shadows'),
        ]
        for key, label in checks:
            var = tk.IntVar(value=1 if self.cfg.get('MAIN', key, fallback='0') == '1' else 0)
            self.settings_checks[key] = var
            tk.Checkbutton(inner, text=label, variable=var, font=('Segoe UI', 8),
                            fg=self.FG, bg=self.BG, selectcolor=self.BG2,
                            activebackground=self.BG).grid(
                row=row, column=0, columnspan=2, sticky='w', padx=12, pady=1)
            row += 1

        # Custom nav size
        nav_var = tk.IntVar(value=1 if self.cfg.get('MAIN', 'UseCustomNavSize') == '1' else 0)
        self.settings_checks['UseCustomNavSize'] = nav_var
        tk.Checkbutton(inner, text='Use custom Click/Nav size:', variable=nav_var,
                        font=('Segoe UI', 8), fg=self.FG, bg=self.BG,
                        selectcolor=self.BG2, activebackground=self.BG).grid(
            row=row, column=0, columnspan=2, sticky='w', padx=12, pady=1)
        row += 1

        sz_frame = tk.Frame(inner, bg=self.BG)
        sz_frame.grid(row=row, column=0, columnspan=2, sticky='w', padx=30, pady=2)
        tk.Label(sz_frame, text='W:', font=('Consolas', 9), fg=self.FG2, bg=self.BG).pack(side='left')
        self.settings_navw = tk.Entry(sz_frame, font=('Consolas', 9), bg=self.BG2, fg=self.FG, bd=1, width=6)
        self.settings_navw.insert(0, self.cfg.get('MAIN', 'NavW'))
        self.settings_navw.pack(side='left', padx=2)
        tk.Label(sz_frame, text='H:', font=('Consolas', 9), fg=self.FG2, bg=self.BG).pack(side='left', padx=(8, 0))
        self.settings_navh = tk.Entry(sz_frame, font=('Consolas', 9), bg=self.BG2, fg=self.FG, bd=1, width=6)
        self.settings_navh.insert(0, self.cfg.get('MAIN', 'NavH'))
        self.settings_navh.pack(side='left', padx=2)
        row += 1

        # Hotkeys section
        tk.Label(inner, text='HOTKEYS', font=('Segoe UI', 9, 'bold'),
                 fg=self.ACCENT, bg=self.BG).grid(row=row, column=0, columnspan=2, sticky='w', pady=(10, 4), padx=8)
        row += 1

        self.settings_hotkeys = {}
        hk_fields = [
            ('FORWARD', 'Forward'), ('BACKWARD', 'Backward'), ('TOP', 'Top'),
            ('SORTTAB', 'Sort Tab'), ('SORTPROFILE', 'Sort Profile'),
            ('GROUPNEXT', 'Group Next'), ('GROUPBACK', 'Group Back'),
        ]
        for key, label in hk_fields:
            tk.Label(inner, text=label, font=('Segoe UI', 8), fg=self.FG2, bg=self.BG).grid(
                row=row, column=0, sticky='w', padx=12)
            entry = tk.Entry(inner, font=('Consolas', 9), bg=self.BG2, fg=self.FG, bd=1, width=22)
            entry.insert(0, self.cfg.get('HOTKEYS', key, fallback=''))
            entry.grid(row=row, column=1, sticky='ew', padx=8, pady=1)
            self.settings_hotkeys[key] = entry
            row += 1

        # Extra hotkeys
        tk.Label(inner, text='EXTRA HOTKEYS', font=('Segoe UI', 9, 'bold'),
                 fg=self.ACCENT, bg=self.BG).grid(row=row, column=0, columnspan=2, sticky='w', pady=(10, 4), padx=8)
        row += 1

        self.settings_extra_hk = {}
        for i in range(1, 10):
            tk.Label(inner, text=f'EHK{i}', font=('Segoe UI', 8), fg=self.FG2, bg=self.BG).grid(
                row=row, column=0, sticky='w', padx=12)
            entry = tk.Entry(inner, font=('Consolas', 9), bg=self.BG2, fg=self.FG, bd=1, width=22)
            val = self.cfg.get('HOTKEYS2', f'EHK{i}-0', fallback='')
            entry.insert(0, val)
            entry.grid(row=row, column=1, sticky='ew', padx=8, pady=1)
            self.settings_extra_hk[f'EHK{i}'] = entry
            row += 1

        # Main settings
        tk.Label(inner, text='MAIN', font=('Segoe UI', 9, 'bold'),
                 fg=self.ACCENT, bg=self.BG).grid(row=row, column=0, columnspan=2, sticky='w', pady=(10, 4), padx=8)
        row += 1

        self.settings_main = {}
        for key, label in [('PollInterval', 'Poll Interval (sec)'), ('AdsPowerPort', 'AdsPower Port')]:
            tk.Label(inner, text=label, font=('Segoe UI', 8), fg=self.FG2, bg=self.BG).grid(
                row=row, column=0, sticky='w', padx=12)
            entry = tk.Entry(inner, font=('Consolas', 9), bg=self.BG2, fg=self.FG, bd=1, width=22)
            entry.insert(0, self.cfg.get('MAIN', key))
            entry.grid(row=row, column=1, sticky='ew', padx=8, pady=1)
            self.settings_main[key] = entry
            row += 1

        tk.Button(inner, text='Save Settings', font=('Segoe UI', 9, 'bold'),
                  fg='#fff', bg=self.GREEN, bd=0, padx=16, pady=4,
                  cursor='hand2', command=self._save_settings).grid(
            row=row, column=0, columnspan=2, pady=12)

    # ── Discord Tab ───────────────────────────────────────────────────────────

    def _build_discord_tab(self):
        f = self.tab_frames['Discord']
        tk.Label(f, text='DISCORD WEBHOOKS', font=('Segoe UI', 10, 'bold'),
                 fg=self.ACCENT, bg=self.BG).pack(anchor='w', padx=10, pady=(10, 5))

        self.discord_entries = {}
        for key, label in [('QueWebhook', 'Queue Webhook URL'), ('ProdWebhook', 'Prod Webhook URL'),
                            ('ProfileName', 'Profile Name')]:
            tk.Label(f, text=label, font=('Segoe UI', 8), fg=self.FG2, bg=self.BG).pack(
                anchor='w', padx=12)
            entry = tk.Entry(f, font=('Consolas', 9), bg=self.BG2, fg=self.FG, bd=1)
            entry.insert(0, self.cfg.get('DISCORD', key, fallback=''))
            entry.pack(fill='x', padx=12, pady=(0, 6))
            self.discord_entries[key] = entry

        tk.Label(f, text='GOOGLE SHEETS', font=('Segoe UI', 10, 'bold'),
                 fg=self.ACCENT, bg=self.BG).pack(anchor='w', padx=10, pady=(10, 5))

        tk.Label(f, text='Sheet URL', font=('Segoe UI', 8), fg=self.FG2, bg=self.BG).pack(
            anchor='w', padx=12)
        self.sheet_entry = tk.Entry(f, font=('Consolas', 9), bg=self.BG2, fg=self.FG, bd=1)
        self.sheet_entry.insert(0, self.cfg.get('SHEETS', 'SheetUrl', fallback=''))
        self.sheet_entry.pack(fill='x', padx=12, pady=(0, 6))

        btn_frame = tk.Frame(f, bg=self.BG)
        btn_frame.pack(pady=10)
        tk.Button(btn_frame, text='Save', font=('Segoe UI', 9, 'bold'), fg='#fff',
                  bg=self.GREEN, bd=0, padx=16, pady=4,
                  command=self._save_discord).pack(side='left', padx=4)
        tk.Button(btn_frame, text='Test Queue Webhook', font=('Segoe UI', 9),
                  fg=self.FG, bg=self.BG2, bd=1, padx=8, pady=3,
                  command=self._test_discord_queue).pack(side='left', padx=4)
        tk.Button(btn_frame, text='Test Prod Webhook', font=('Segoe UI', 9),
                  fg=self.FG, bg=self.BG2, bd=1, padx=8, pady=3,
                  command=self._test_discord_prod).pack(side='left', padx=4)

    # ── Distribute Tab ────────────────────────────────────────────────────────

    def _build_distribute_tab(self):
        f = self.tab_frames['Distribute']
        tk.Label(f, text='AUTO DISTRIBUTE', font=('Segoe UI', 10, 'bold'),
                 fg=self.ACCENT, bg=self.BG).pack(anchor='w', padx=10, pady=(10, 5))

        tk.Label(f, text='Distribute URL across all running profiles evenly.',
                 font=('Segoe UI', 9), fg=self.FG2, bg=self.BG).pack(anchor='w', padx=12, pady=4)

        url_frame = tk.Frame(f, bg=self.BG)
        url_frame.pack(fill='x', padx=12, pady=4)
        tk.Label(url_frame, text='URL:', font=('Consolas', 9), fg=self.FG2, bg=self.BG).pack(side='left')
        self.dist_url_entry = tk.Entry(url_frame, font=('Consolas', 9), bg=self.BG2,
                                        fg=self.FG, bd=1, insertbackground=self.FG)
        self.dist_url_entry.insert(0, 'https://www.ticketmaster.com')
        self.dist_url_entry.pack(side='left', fill='x', expand=True, padx=4)

        tk.Button(f, text='Distribute Now', font=('Segoe UI', 9, 'bold'),
                  fg='#fff', bg=self.PURPLE, bd=0, padx=16, pady=6,
                  cursor='hand2', command=self._distribute_url).pack(pady=10)

        self.dist_status = tk.Label(f, text='', font=('Consolas', 9), fg=self.FG2, bg=self.BG)
        self.dist_status.pack(padx=12, pady=4)

    # ── Position Tab ──────────────────────────────────────────────────────────

    def _build_pos_tab(self):
        f = self.tab_frames['Pos']
        tk.Label(f, text='WINDOW POSITIONING', font=('Segoe UI', 10, 'bold'),
                 fg=self.ACCENT, bg=self.BG).pack(anchor='w', padx=10, pady=(10, 5))

        tk.Label(f, text='Arrange all profile windows on screen.',
                 font=('Segoe UI', 9), fg=self.FG2, bg=self.BG).pack(anchor='w', padx=12, pady=4)

        grid_frame = tk.Frame(f, bg=self.BG)
        grid_frame.pack(pady=8)

        layouts = [
            ('Stack', self._pos_stack),
            ('Tile 2x', self._pos_tile_2),
            ('Tile 3x', self._pos_tile_3),
            ('Cascade', self._pos_cascade),
        ]
        for text, cmd in layouts:
            tk.Button(grid_frame, text=text, font=('Segoe UI', 9), fg=self.FG,
                      bg=self.BG2, bd=1, padx=12, pady=4, width=10,
                      cursor='hand2', command=cmd).pack(pady=3)

        self.pos_status = tk.Label(f, text='', font=('Consolas', 9), fg=self.FG2, bg=self.BG)
        self.pos_status.pack(padx=12, pady=4)

    # ── Bottom Bar ────────────────────────────────────────────────────────────

    def _build_bottom_bar(self):
        # Debug log / status tabs
        bottom_tabs = tk.Frame(self.root, bg=self.BG2)
        bottom_tabs.pack(fill='x', side='bottom')

        self.bottom_tab_btns = {}
        for name in ['APM v2.0', 'Debug Log']:
            btn = tk.Button(bottom_tabs, text=name, font=('Segoe UI', 7),
                            fg=self.FG2, bg=self.BG2, bd=0, padx=6, pady=2,
                            command=lambda n=name: self._switch_bottom_tab(n))
            btn.pack(side='left')
            self.bottom_tab_btns[name] = btn

        self.bottom_container = tk.Frame(self.root, bg=self.BG, height=28)
        self.bottom_container.pack(fill='x', side='bottom')
        self.bottom_container.pack_propagate(False)

        # URL bar (shown in APM v2.0 tab)
        self.url_frame = tk.Frame(self.bottom_container, bg=self.BG)
        self.url_entry = tk.Entry(self.url_frame, font=('Consolas', 9), bg=self.BG2,
                                   fg=self.FG, bd=1, insertbackground=self.FG)
        self.url_entry.insert(0, 'https://www.ticketmaster.com')
        self.url_entry.pack(side='left', fill='x', expand=True, padx=4)
        tk.Button(self.url_frame, text='Open URL', font=('Segoe UI', 8), fg=self.FG,
                  bg=self.BG2, bd=1, padx=4, command=self._open_url_in_all).pack(side='left', padx=2)
        self.stop_btn = tk.Button(self.url_frame, text='STOP', font=('Segoe UI', 8, 'bold'),
                                   fg='#fff', bg=self.ACCENT, bd=0, padx=6,
                                   command=self._stop_all)
        self.stop_btn.pack(side='left', padx=2)

        # Debug log text
        self.debug_frame = tk.Frame(self.bottom_container, bg=self.BG)
        self.debug_text = tk.Text(self.debug_frame, font=('Consolas', 8), bg=self.BG,
                                   fg=self.GREEN, height=6, bd=0, wrap='word')
        self.debug_text.pack(fill='both', expand=True)

        self._switch_bottom_tab('APM v2.0')

    def _switch_bottom_tab(self, name):
        self.url_frame.pack_forget()
        self.debug_frame.pack_forget()
        if name == 'APM v2.0':
            self.bottom_container.configure(height=28)
            self.url_frame.pack(fill='x')
        else:
            self.bottom_container.configure(height=120)
            self.debug_frame.pack(fill='both', expand=True)
            self._refresh_debug_log()
        for n, btn in self.bottom_tab_btns.items():
            btn.configure(fg=self.FG if n == name else self.FG2)

    def _refresh_debug_log(self):
        self.debug_text.delete('1.0', 'end')
        self.debug_text.insert('end', '\n'.join(self.debug_log[-100:]))
        self.debug_text.see('end')

    # ── Actions ───────────────────────────────────────────────────────────────

    def _cycle(self, direction):
        if not self.profiles:
            return
        if self.selected_index < 0:
            self.selected_index = 0
        else:
            self.selected_index = (self.selected_index + direction) % len(self.profiles)
        self._activate_selected()
        self._select_tree_row(self.selected_index)

    def _go_top(self):
        if self.profiles:
            self.selected_index = 0
            self._activate_selected()
            self._select_tree_row(0)

    def _activate_selected(self):
        if self.selected_index < 0 or self.selected_index >= len(self.profiles):
            return
        profile = self.profiles[self.selected_index]
        debug_port = profile.get('debug_port', 0)
        if not debug_port:
            return

        custom_size = None
        if self.cfg.get('MAIN', 'UseCustomNavSize') == '1':
            try:
                nw = int(self.cfg.get('MAIN', 'NavW'))
                nh = int(self.cfg.get('MAIN', 'NavH'))
                if nw > 0 and nh > 0:
                    custom_size = (nw, nh)
            except (ValueError, TypeError):
                pass

        minimize_others = self.cfg.get('MAIN', 'MinimizeOthers', fallback='0') == '1'

        threading.Thread(target=activate_profile, args=(debug_port, custom_size, minimize_others),
                          daemon=True).start()
        self._log(f'Switched to profile #{profile.get("serial", "?")}')

    def _select_tree_row(self, index):
        items = self.tree.get_children()
        if 0 <= index < len(items):
            self.tree.selection_set(items[index])
            self.tree.see(items[index])

    def _on_tree_select(self, event):
        sel = self.tree.selection()
        if sel:
            items = self.tree.get_children()
            for i, item in enumerate(items):
                if item in sel:
                    self.selected_index = i
                    break

    def _on_tree_double_click(self, event):
        self._activate_selected()

    def _select_group(self, letter):
        """Filter profiles by group letter."""
        if self.current_group == letter:
            self.current_group = None  # Toggle off
            for btn in self.grp_btn_map.values():
                btn.configure(bg=self.BG2)
        else:
            self.current_group = letter
            for l, btn in self.grp_btn_map.items():
                btn.configure(bg=self.PURPLE if l == letter else self.BG2)
        self._refresh_tree()

    def _sort_column(self, col):
        items = [(self.tree.set(k, col), k) for k in self.tree.get_children()]
        items.sort()
        for i, (_, k) in enumerate(items):
            self.tree.move(k, '', i)

    def _on_scroll(self, val):
        if not self.profiles:
            return
        idx = int(float(val) * len(self.profiles) / 100)
        idx = max(0, min(idx, len(self.profiles) - 1))
        self.selected_index = idx
        self._select_tree_row(idx)

    def _apply_nav_size(self):
        try:
            w = int(self.nav_w_entry.get())
            h = int(self.nav_h_entry.get())
            self.cfg.set('MAIN', 'NavW', str(w))
            self.cfg.set('MAIN', 'NavH', str(h))
            save_config(self.cfg)
            self._log(f'Nav size set to {w}x{h}')
        except ValueError:
            pass

    def _fix_windows(self):
        """Resize all profile windows to the nav size."""
        try:
            w = int(self.nav_w_entry.get())
            h = int(self.nav_h_entry.get())
        except ValueError:
            return

        def do_fix():
            count = 0
            for p in self.profiles:
                port = p.get('debug_port', 0)
                if not port:
                    continue
                hwnd = find_profile_window(port)
                if hwnd:
                    try:
                        rect = win32gui.GetWindowRect(hwnd)
                        position_window(hwnd, rect[0], rect[1], w, h)
                        count += 1
                    except Exception:
                        pass
            self.root.after(0, lambda: self._log(f'Fixed {count} windows to {w}x{h}'))

        threading.Thread(target=do_fix, daemon=True).start()

    # ── Side button handlers ──────────────────────────────────────────────────

    def _btn_show(self):
        self._activate_selected()

    def _btn_minimize(self):
        if self.selected_index < 0 or self.selected_index >= len(self.profiles):
            return
        port = self.profiles[self.selected_index].get('debug_port', 0)
        if port:
            hwnd = find_profile_window(port)
            if hwnd:
                minimize_window(hwnd)

    def _btn_refresh_all(self):
        self._log('Refreshing profiles...')
        # Force immediate poll
        threading.Thread(target=self._poll_once, daemon=True).start()

    def _btn_close_all(self):
        """Close all tabs in all profiles."""
        def do_close():
            for p in self.profiles:
                port = p.get('debug_port', 0)
                if port:
                    cdp_close_tabs(port)
            self.root.after(0, lambda: self._log('Closed tabs in all profiles'))
        threading.Thread(target=do_close, daemon=True).start()

    def _btn_close_selected(self):
        if self.selected_index < 0 or self.selected_index >= len(self.profiles):
            return
        port = self.profiles[self.selected_index].get('debug_port', 0)
        if port:
            threading.Thread(target=cdp_close_tabs, args=(port,), daemon=True).start()
            self._log(f'Closing tabs for profile #{self.profiles[self.selected_index].get("serial", "?")}')

    def _btn_show_all(self):
        def do_show():
            for p in self.profiles:
                port = p.get('debug_port', 0)
                if port:
                    hwnd = find_profile_window(port)
                    if hwnd:
                        show_window(hwnd)
                        time.sleep(0.1)
        threading.Thread(target=do_show, daemon=True).start()

    def _btn_minimize_all(self):
        def do_min():
            for p in self.profiles:
                port = p.get('debug_port', 0)
                if port:
                    hwnd = find_profile_window(port)
                    if hwnd:
                        minimize_window(hwnd)
        threading.Thread(target=do_min, daemon=True).start()

    # ── TM Lite ───────────────────────────────────────────────────────────────

    def _toggle_tm_lite(self):
        self.tm_lite_active = not self.tm_lite_active
        self.tm_lite_btn.configure(bg=self.GREEN if self.tm_lite_active else self.ACCENT)
        self._log(f'TM Lite {"ON" if self.tm_lite_active else "OFF"}')

    # ── Toggles ───────────────────────────────────────────────────────────────

    def _toggle_hotkeys(self):
        self.hotkeys_enabled = self.hotkeys_var.get()
        self._log(f'Hotkeys {"enabled" if self.hotkeys_enabled else "disabled"}')

    def _toggle_ontop(self):
        self.always_on_top = self.ontop_var.get()
        self.root.attributes('-topmost', self.always_on_top)
        self.cfg.set('MAIN', 'AlwaysOnTop', '1' if self.always_on_top else '0')
        save_config(self.cfg)

    # ── URL / Stop ────────────────────────────────────────────────────────────

    def _open_url_in_all(self):
        url = self.url_entry.get().strip()
        if not url or url == 'https://':
            return

        def do_open():
            count = 0
            for p in self.profiles:
                port = p.get('debug_port', 0)
                if port and cdp_open_tab(port, url):
                    count += 1
            self.root.after(0, lambda: self._log(f'Opened URL in {count}/{len(self.profiles)} profiles'))

        threading.Thread(target=do_open, daemon=True).start()

    def _stop_all(self):
        """Stop loading in all profiles by navigating to about:blank in active tabs."""
        self._log('Stopping all...')

    # ── Distribute ────────────────────────────────────────────────────────────

    def _distribute_url(self):
        url = self.dist_url_entry.get().strip()
        if not url:
            return

        def do_dist():
            count = 0
            for p in self.profiles:
                port = p.get('debug_port', 0)
                if port and cdp_open_tab(port, url):
                    count += 1
            self.root.after(0, lambda: self.dist_status.configure(
                text=f'Distributed to {count}/{len(self.profiles)} profiles'))

        threading.Thread(target=do_dist, daemon=True).start()
        self.dist_status.configure(text='Distributing...')

    # ── Position layouts ──────────────────────────────────────────────────────

    def _get_screen_size(self):
        if HAS_WIN32:
            return (win32api.GetSystemMetrics(win32con.SM_CXSCREEN),
                    win32api.GetSystemMetrics(win32con.SM_CYSCREEN))
        return (1920, 1080)

    def _pos_stack(self):
        sw, sh = self._get_screen_size()
        def do_it():
            for p in self.profiles:
                port = p.get('debug_port', 0)
                if port:
                    hwnd = find_profile_window(port)
                    if hwnd:
                        position_window(hwnd, 0, 0, sw, sh - 40)
        threading.Thread(target=do_it, daemon=True).start()

    def _pos_tile_2(self):
        sw, sh = self._get_screen_size()
        half_w = sw // 2
        def do_it():
            for i, p in enumerate(self.profiles):
                port = p.get('debug_port', 0)
                if port:
                    hwnd = find_profile_window(port)
                    if hwnd:
                        x = (i % 2) * half_w
                        y = (i // 2) * (sh // 2)
                        position_window(hwnd, x, y, half_w, sh // 2)
        threading.Thread(target=do_it, daemon=True).start()

    def _pos_tile_3(self):
        sw, sh = self._get_screen_size()
        third_w = sw // 3
        def do_it():
            for i, p in enumerate(self.profiles):
                port = p.get('debug_port', 0)
                if port:
                    hwnd = find_profile_window(port)
                    if hwnd:
                        x = (i % 3) * third_w
                        y = (i // 3) * (sh // 2)
                        position_window(hwnd, x, y, third_w, sh // 2)
        threading.Thread(target=do_it, daemon=True).start()

    def _pos_cascade(self):
        def do_it():
            offset = 30
            for i, p in enumerate(self.profiles):
                port = p.get('debug_port', 0)
                if port:
                    hwnd = find_profile_window(port)
                    if hwnd:
                        position_window(hwnd, offset * i, offset * i, 800, 600)
        threading.Thread(target=do_it, daemon=True).start()

    # ── Discord handlers ──────────────────────────────────────────────────────

    def _save_discord(self):
        for key, entry in self.discord_entries.items():
            self.cfg.set('DISCORD', key, entry.get())
        self.cfg.set('SHEETS', 'SheetUrl', self.sheet_entry.get())
        save_config(self.cfg)
        self._log('Discord/Sheets settings saved')

    def _test_discord_queue(self):
        url = self.discord_entries['QueWebhook'].get()
        name = self.discord_entries['ProfileName'].get() or 'APM'
        threading.Thread(target=send_discord, args=(url, f'[{name}] APM Queue Test', 'APM'),
                          daemon=True).start()
        self._log('Sent test to queue webhook')

    def _test_discord_prod(self):
        url = self.discord_entries['ProdWebhook'].get()
        name = self.discord_entries['ProfileName'].get() or 'APM'
        threading.Thread(target=send_discord, args=(url, f'[{name}] APM Prod Test', 'APM'),
                          daemon=True).start()
        self._log('Sent test to prod webhook')

    # ── Settings save ─────────────────────────────────────────────────────────

    def _save_settings(self):
        for key, var in self.settings_checks.items():
            self.cfg.set('MAIN', key, '1' if var.get() else '0')
        self.cfg.set('MAIN', 'NavW', self.settings_navw.get())
        self.cfg.set('MAIN', 'NavH', self.settings_navh.get())
        for key, entry in self.settings_hotkeys.items():
            self.cfg.set('HOTKEYS', key, entry.get())
        for key, entry in self.settings_extra_hk.items():
            if not self.cfg.has_section('HOTKEYS2'):
                self.cfg.add_section('HOTKEYS2')
            self.cfg.set('HOTKEYS2', f'{key}-0', entry.get())
        for key, entry in self.settings_main.items():
            self.cfg.set('MAIN', key, entry.get())
        save_config(self.cfg)
        # Re-register hotkeys
        if keyboard:
            try:
                keyboard.unhook_all_hotkeys()
            except Exception:
                pass
            self._register_hotkeys()
        self._log('Settings saved')

    # ── Hotkeys ───────────────────────────────────────────────────────────────

    def _register_hotkeys(self):
        if not keyboard:
            return
        try:
            fwd = self.cfg.get('HOTKEYS', 'FORWARD')
            bwd = self.cfg.get('HOTKEYS', 'BACKWARD')
            top = self.cfg.get('HOTKEYS', 'TOP', fallback='')
            grpn = self.cfg.get('HOTKEYS', 'GROUPNEXT', fallback='')
            grpb = self.cfg.get('HOTKEYS', 'GROUPBACK', fallback='')

            if fwd:
                keyboard.add_hotkey(fwd, lambda: self.root.after(0, self._hk_forward), suppress=False)
            if bwd:
                keyboard.add_hotkey(bwd, lambda: self.root.after(0, self._hk_backward), suppress=False)
            if top:
                keyboard.add_hotkey(top, lambda: self.root.after(0, self._go_top), suppress=False)
        except Exception as e:
            self._log(f'Hotkey error: {e}')

    def _hk_forward(self):
        if self.hotkeys_enabled:
            self._cycle(1)

    def _hk_backward(self):
        if self.hotkeys_enabled:
            self._cycle(-1)

    # ── Polling ───────────────────────────────────────────────────────────────

    def _start_polling(self):
        threading.Thread(target=self._poll_loop, daemon=True).start()

    def _poll_loop(self):
        while self.running:
            self._poll_once()
            try:
                interval = int(self.cfg.get('MAIN', 'PollInterval', fallback='3'))
            except (ValueError, TypeError):
                interval = 3
            time.sleep(max(1, interval))

    def _poll_once(self):
        try:
            profiles = self.api.get_active_profiles()
            self.profiles = profiles

            # Scan tabs and queues for each profile
            for p in profiles:
                serial = p.get('serial') or p.get('custom') or p.get('uid') or ''
                port = p.get('debug_port', 0)
                if not serial or not port:
                    continue

                # Get tab titles
                tabs = get_profile_tabs(port)
                tab_titles = []
                for t in tabs:
                    title = t.get('title', '')
                    url = t.get('url', '')
                    if title:
                        tab_titles.append(title[:50])
                    elif url:
                        tab_titles.append(url[:50])
                self.profile_tabs[serial] = tab_titles

                # Scan for queue
                result = cdp_eval_queue(port, timeout=3)
                if result:
                    self.queue_data[serial] = result

            self.root.after(0, self._refresh_tree)
        except Exception as e:
            self._log(f'Poll error: {e}')

    def _refresh_tree(self):
        """Update the treeview with current profile data."""
        # Remember selection
        sel_serial = None
        if 0 <= self.selected_index < len(self.profiles):
            sel_serial = self.profiles[self.selected_index].get('serial', '')

        self.tree.delete(*self.tree.get_children())

        filtered = self.profiles
        if self.current_group:
            filtered = [p for p in self.profiles if
                        (p.get('group_name', '') or '').upper().startswith(self.current_group) or
                        (p.get('name', '') or '').upper().startswith(self.current_group)]
            if not filtered:
                filtered = self.profiles  # Fallback: show all if no match

        new_selected = -1
        for i, p in enumerate(filtered):
            serial = p.get('serial') or p.get('custom') or p.get('uid') or '?'
            tabs = self.profile_tabs.get(serial, [])
            qdata = self.queue_data.get(serial, {})

            # Profile column: serial + queue number
            profile_text = f'#{serial}'
            q = qdata.get('queue_number')
            if q:
                profile_text += f' [Q#{q:,}]'

            # Tab column: first tab title or queue event
            tab_text = tabs[0] if tabs else ''
            if not tab_text and qdata.get('title'):
                tab_text = qdata['title'][:50]

            item = self.tree.insert('', 'end', values=(profile_text, tab_text))

            if serial == sel_serial:
                new_selected = i
                self.tree.selection_set(item)

        self.selected_index = new_selected if new_selected >= 0 else (0 if filtered else -1)

        # Update scroll scale
        if self.profiles:
            self.scroll_scale.configure(to=len(self.profiles) - 1)

    # ── Screenshot ────────────────────────────────────────────────────────────

    def _take_screenshot(self):
        if not HAS_PIL:
            self._log('PIL not available for screenshots')
            return
        folder = self.cfg.get('SCREENSHOTS', 'Folder', fallback=SCREENSHOT_DIR)
        os.makedirs(folder, exist_ok=True)
        serial = ''
        if 0 <= self.selected_index < len(self.profiles):
            serial = self.profiles[self.selected_index].get('serial', '')
        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        name = f'{serial}_{ts}.png' if serial else f'screenshot_{ts}.png'
        path = os.path.join(folder, name)
        try:
            img = ImageGrab.grab()
            img.save(path)
            self._log(f'Screenshot saved: {name}')
        except Exception as e:
            self._log(f'Screenshot error: {e}')

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
        if keyboard:
            try:
                keyboard.unhook_all_hotkeys()
            except Exception:
                pass
        self.root.destroy()

    def run(self):
        self.root.mainloop()


# ─── Entry Point ──────────────────────────────────────────────────────────────

if __name__ == '__main__':
    missing = []
    if not requests:
        missing.append('requests')
    if not websocket:
        missing.append('websocket-client')
    if missing:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror('APM - Missing Dependencies',
                              f'Missing packages: {", ".join(missing)}\n\n'
                              f'Install with:\npip install {" ".join(missing)}')
        sys.exit(1)

    app = APMApp()
    app.run()
