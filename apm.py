"""
APM - AdsPower Manager v1.0
Standalone desktop app for managing AdsPower browser profiles.
Similar to MLM (MultiLogin Manager) but for AdsPower.
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
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


# ─── Config ───────────────────────────────────────────────────────────────────

def get_base_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

BASE_DIR = get_base_dir()
CONFIG_PATH = os.path.join(BASE_DIR, 'config.ini')

def load_config():
    cfg = configparser.ConfigParser()
    cfg.read(CONFIG_PATH, encoding='utf-8')
    return cfg

def save_config(cfg):
    with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
        cfg.write(f)


# ─── AdsPower API ─────────────────────────────────────────────────────────────

class AdsPowerAPI:
    def __init__(self, port=50325):
        self.bases = [
            f'http://127.0.0.1:{port}',
            f'http://local.adspower.net:{port}'
        ]

    def _get(self, path, timeout=4):
        for base in self.bases:
            try:
                r = requests.get(base + path, timeout=timeout)
                if r.ok:
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

        # Check if active endpoint has serial numbers
        has_serial = any(p.get('serial_number') or p.get('serialnumber') for p in raw)

        if has_serial:
            return self._normalize_profiles(raw)

        # Cross-reference with user/list for serial numbers
        active_ids = set()
        for p in raw:
            uid = str(p.get('user_id', '')).strip()
            if uid:
                active_ids.add(uid)

        user_data = self._get('/api/v1/user/list?page=1&page_size=100')
        if not user_data:
            return self._normalize_profiles(raw)

        all_users = user_data.get('data', {})
        if isinstance(all_users, dict):
            all_users = all_users.get('list', [])
        if not isinstance(all_users, list):
            return self._normalize_profiles(raw)

        # Merge active info (debug_port, ws) with user/list info (serial)
        merged = []
        for u in all_users:
            uid = str(u.get('user_id', '')).strip()
            if active_ids and uid not in active_ids:
                continue
            # Copy debug port/ws from active list
            for a in raw:
                if str(a.get('user_id', '')).strip() == uid:
                    if not u.get('debug_port') and a.get('debug_port'):
                        u['debug_port'] = a['debug_port']
                    if not u.get('ws') and a.get('ws'):
                        u['ws'] = a['ws']
                    break
            merged.append(u)

        return self._normalize_profiles(merged if merged else raw)

    def _normalize_profiles(self, raw):
        profiles = []
        for p in raw:
            serial = str(p.get('serial_number') or p.get('serialnumber') or '').strip()
            custom = str(p.get('custom_user_id') or p.get('customUserId') or '').strip()
            uid = str(p.get('user_id') or p.get('userId') or p.get('id') or '').strip()
            name = str(p.get('name') or p.get('profile_name') or (f'#{serial}' if serial else uid) or '')
            debug_port = 0
            if p.get('debug_port'):
                try:
                    debug_port = int(p['debug_port'])
                except (ValueError, TypeError):
                    pass
            if not debug_port:
                ws_url = ''
                ws = p.get('ws', {})
                if isinstance(ws, dict):
                    ws_url = ws.get('puppeteer') or ws.get('selenium') or ''
                if ws_url:
                    try:
                        from urllib.parse import urlparse
                        debug_port = int(urlparse(ws_url).port or 0)
                    except Exception:
                        m = re.search(r':(\d+)/', ws_url)
                        if m:
                            debug_port = int(m.group(1))

            profiles.append({
                'serial': serial,
                'custom': custom,
                'uid': uid,
                'name': name,
                'debug_port': debug_port,
            })
        return profiles


# ─── CDP Queue Scanner ────────────────────────────────────────────────────────

TM_HOST_RE = re.compile(
    r'ticketmaster\.com|livenation\.com|queue-it\.net|ticketmaster\.ca|ticketmaster\.co\.uk',
    re.I
)

QUEUE_PATTERNS = [
    re.compile(r'you\s+are\s+(?:now\s+)?in\s+the\s+queue\s*#?\s*([\d,]{1,8})', re.I),
    re.compile(r'\bin\s+the\s+queue\s*#\s*([\d,]{1,8})', re.I),
    re.compile(r'([\d,]{1,8})\s+people\s+ahead\s+of\s+you', re.I),
    re.compile(r'([\d,]{1,8})\s+people?\s+ahead', re.I),
    re.compile(r'you\s+are\s+(?:now\s+)?(?:number\s+|#\s*)?([\d,]{1,8})\s+in', re.I),
    re.compile(r'(?:queue\s+)?position\s+(?:is\s+)?(?:number\s+|#\s*)?([\d,]{1,8})', re.I),
    re.compile(r'there\s+are\s+([\d,]{1,8})\s+people', re.I),
    re.compile(r'#([\d,]{1,8})\s+in\s+(?:line|queue)', re.I),
]

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
/\\bin\\s+the\\s+queue\\s*#\\s*([\\d,]+)/i,
/([\\d,]+)\\s+people\\s+ahead\\s+of\\s+you/i,
/([\\d,]+)\\s+people?\\s+ahead/i,
/you\\s+are\\s+(?:now\\s+)?(?:number\\s+|#\\s*)?([\\d,]+)\\s+in/i,
/(?:queue\\s+)?position\\s+(?:is\\s+)?(?:number\\s+|#\\s*)?([\\d,]+)/i,
/there\\s+are\\s+([\\d,]+)\\s+people/i];
for(var j=0;j<pts.length;j++){var mm=txt.match(pts[j]);if(mm){var v=p(mm[1]);if(v)return{q:v,u:location.href,t:document.title};}}
return null;
})()"""


def parse_queue_from_title(title):
    """Extract queue number from tab title like '922 | Event Name'."""
    if not title:
        return None
    m = re.match(r'^\s*#?([\d,]{1,8})\s*\|', title)
    if m:
        n = int(m.group(1).replace(',', ''))
        if 0 < n < 100000000:
            return n
    for pat in QUEUE_PATTERNS:
        m = pat.search(title)
        if m:
            n = int(m.group(1).replace(',', ''))
            if 0 < n < 100000000:
                return n
    return None


def event_title_from_url(url):
    """Extract a human-readable event name from a URL."""
    try:
        from urllib.parse import urlparse, parse_qs, unquote
        u = urlparse(url)
        parts = [unquote(p).strip() for p in u.path.split('/') if p.strip()]
        skip = {'event', 'events', 'queue', 'checkout', 'tickets', 'ticket',
                'signup', 'thewaitingroom', 'waitingroom'}
        for p in reversed(parts):
            if p.lower() not in skip:
                label = p.replace('-', ' ').replace('_', ' ').title()
                if label and not label.isdigit():
                    return label
        qs = parse_qs(u.query)
        eid = (qs.get('e') or qs.get('eventId') or [''])[0]
        if eid:
            return f'Event {eid[:8]}'
        host = u.hostname or ''
        host = re.sub(r'^www\.', '', host).split('.')[0]
        if host.lower() in ('queue', 'signup'):
            return 'Queue'
        return host.title() if host else ''
    except Exception:
        return ''


def clean_event_title(raw, url=''):
    """Clean up a raw page title into a short event name."""
    fallback = event_title_from_url(url)
    if not raw:
        return fallback
    cleaned = raw
    cleaned = re.sub(r'^\s*#?[\d,]{1,8}\s*\|\s*', '', cleaned)
    cleaned = re.sub(r'\s*[|\-\u2013\u2014]\s*ticketmaster.*$', '', cleaned, flags=re.I)
    cleaned = re.sub(r'\s*[|\-\u2013\u2014]\s*livenation.*$', '', cleaned, flags=re.I)
    cleaned = re.sub(r'\s*[|\-\u2013\u2014]\s*queue.*$', '', cleaned, flags=re.I)
    cleaned = cleaned.strip()
    if not cleaned or re.match(r'^\d[\d,\s]*$', cleaned):
        return fallback
    if re.match(r'^(queue|waiting room|please wait|hold tight)$', cleaned, re.I):
        return fallback
    return cleaned


def cdp_eval_queue(debug_port, timeout=4):
    """
    Connect to a profile's CDP, find TM/queue tabs, evaluate queue number.
    Returns dict {queue_number, url, title, event} or None.
    """
    if not debug_port or debug_port <= 0:
        return None
    try:
        # Get tab list
        r = requests.get(f'http://127.0.0.1:{debug_port}/json', timeout=2)
        if not r.ok:
            return None
        tabs = r.json()
    except Exception:
        return None

    if not isinstance(tabs, list):
        return None

    # Find TM/queue tabs first, then fall back to any tab
    tm_tabs = [t for t in tabs if t.get('url') and TM_HOST_RE.search(t['url'])
               and t.get('webSocketDebuggerUrl')]
    if not tm_tabs:
        # Try any non-internal tab
        tm_tabs = [t for t in tabs if t.get('url')
                   and not re.match(r'^(devtools:|chrome:|chrome-extension:)', t['url'], re.I)
                   and t.get('webSocketDebuggerUrl')][:2]

    for tab in tm_tabs:
        ws_url = tab.get('webSocketDebuggerUrl', '')
        if not ws_url:
            continue

        # Try evaluating JS via WebSocket
        result = _ws_eval(ws_url, QUEUE_EVAL_JS, timeout=timeout)
        if result and isinstance(result, dict) and result.get('q'):
            url = result.get('u', tab.get('url', ''))
            title = result.get('t', tab.get('title', ''))
            return {
                'queue_number': result['q'],
                'url': url,
                'title': title,
                'event': clean_event_title(title, url),
            }

        # Fallback: check tab title
        q = parse_queue_from_title(tab.get('title', ''))
        if q:
            url = tab.get('url', '')
            title = tab.get('title', '')
            return {
                'queue_number': q,
                'url': url,
                'title': title,
                'event': clean_event_title(title, url),
            }

    return None


def _ws_eval(ws_url, expression, timeout=4):
    """Evaluate a JS expression via CDP WebSocket. Returns the value or None."""
    try:
        ws = websocket.create_connection(ws_url, timeout=timeout)
    except Exception:
        return None
    try:
        msg_id = 90000 + int(time.time() * 1000) % 10000
        ws.send(json.dumps({
            'id': msg_id,
            'method': 'Runtime.evaluate',
            'params': {
                'expression': expression,
                'returnByValue': True,
            }
        }))
        ws.settimeout(timeout)
        while True:
            raw = ws.recv()
            if not raw:
                break
            msg = json.loads(raw)
            if msg.get('id') == msg_id:
                val = (msg.get('result', {}).get('result', {}) or {}).get('value')
                return val
    except Exception:
        pass
    finally:
        try:
            ws.close()
        except Exception:
            pass
    return None


def cdp_open_tab(debug_port, url, timeout=5):
    """Open a new tab in a profile browser via CDP."""
    if not debug_port or debug_port <= 0:
        return False
    try:
        r = requests.get(f'http://127.0.0.1:{debug_port}/json/version', timeout=2)
        if not r.ok:
            return False
        data = r.json()
        ws_url = data.get('webSocketDebuggerUrl', '')
        if not ws_url:
            return False
    except Exception:
        return False

    try:
        ws = websocket.create_connection(ws_url, timeout=timeout)
    except Exception:
        return False
    try:
        msg_id = 80000 + int(time.time() * 1000) % 10000
        ws.send(json.dumps({
            'id': msg_id,
            'method': 'Target.createTarget',
            'params': {'url': url}
        }))
        ws.settimeout(timeout)
        while True:
            raw = ws.recv()
            if not raw:
                break
            msg = json.loads(raw)
            if msg.get('id') == msg_id:
                return bool(msg.get('result', {}).get('targetId'))
    except Exception:
        pass
    finally:
        try:
            ws.close()
        except Exception:
            pass
    return False


# ─── Window Management (Windows-specific) ────────────────────────────────────

def find_profile_window(debug_port):
    """Find the Windows HWND for an AdsPower profile browser by its debug port."""
    if not HAS_WIN32 or not debug_port:
        return None

    # Get the PID listening on this debug port by checking /json endpoint
    target_pids = set()
    try:
        r = requests.get(f'http://127.0.0.1:{debug_port}/json/version', timeout=2)
        if r.ok:
            # The process listening on debug_port is the browser
            import subprocess
            result = subprocess.run(
                ['netstat', '-ano'],
                capture_output=True, text=True, timeout=3
            )
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

    # Find all windows belonging to these PIDs, pick the main browser window
    best_hwnd = None
    best_area = 0

    def enum_callback(hwnd, _):
        nonlocal best_hwnd, best_area
        if not win32gui.IsWindowVisible(hwnd):
            return True
        try:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            # Check if this PID or its parent is our target
            if pid not in target_pids:
                return True
            title = win32gui.GetWindowText(hwnd)
            if not title:
                return True
            rect = win32gui.GetWindowRect(hwnd)
            w = rect[2] - rect[0]
            h = rect[3] - rect[1]
            area = w * h
            if area > best_area:
                best_area = area
                best_hwnd = hwnd
        except Exception:
            pass
        return True

    try:
        win32gui.EnumWindows(enum_callback, None)
    except Exception:
        pass

    return best_hwnd


def activate_window(hwnd):
    """Bring a window to the foreground."""
    if not HAS_WIN32 or not hwnd:
        return
    try:
        # Restore if minimized
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        win32gui.SetForegroundWindow(hwnd)
    except Exception:
        pass


def position_profile_window(hwnd, apm_rect):
    """Position a profile browser window to the right of the APM window."""
    if not HAS_WIN32 or not hwnd:
        return
    try:
        screen_w = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
        screen_h = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)
        # Place browser window to the right of APM
        apm_right = apm_rect[2] if apm_rect else 400
        browser_x = apm_right
        browser_y = 0
        browser_w = screen_w - browser_x
        browser_h = screen_h - 40  # Leave space for taskbar

        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)

        win32gui.SetWindowPos(
            hwnd, win32con.HWND_TOP,
            browser_x, browser_y, browser_w, browser_h,
            win32con.SWP_SHOWWINDOW
        )
        win32gui.SetForegroundWindow(hwnd)
    except Exception:
        pass


# ─── Discord Integration ─────────────────────────────────────────────────────

def send_discord_webhook(webhook_url, content, username='APM'):
    """Send a message to a Discord webhook."""
    if not webhook_url:
        return
    try:
        requests.post(webhook_url, json={
            'content': content,
            'username': username,
        }, timeout=5)
    except Exception:
        pass


# ─── Google Sheets Integration ────────────────────────────────────────────────

def push_to_sheets(sheet_url, data):
    """Push queue data to a Google Apps Script webhook."""
    if not sheet_url:
        return
    try:
        requests.post(sheet_url, json=data, timeout=5)
    except Exception:
        pass


# ─── Screenshot ───────────────────────────────────────────────────────────────

def take_screenshot(folder, profile_name=''):
    """Take a screenshot of the active window and save it."""
    if not HAS_PIL:
        return None
    try:
        os.makedirs(folder, exist_ok=True)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        name = f'{profile_name}_{timestamp}.png' if profile_name else f'screenshot_{timestamp}.png'
        path = os.path.join(folder, name)
        img = ImageGrab.grab()
        img.save(path)
        return path
    except Exception:
        return None


# ─── Main GUI ─────────────────────────────────────────────────────────────────

class APMApp:
    def __init__(self):
        self.cfg = load_config()
        self.api = AdsPowerAPI(
            port=int(self.cfg.get('MAIN', 'AdsPowerPort', fallback='50325'))
        )
        self.profiles = []
        self.queue_data = {}  # serial -> {queue_number, url, title, event}
        self.selected_index = -1
        self.running = True
        self.poll_interval = int(self.cfg.get('MAIN', 'PollInterval', fallback='3'))

        self._build_gui()
        self._register_hotkeys()
        self._start_polling()

    def _build_gui(self):
        self.root = tk.Tk()
        self.root.title('APM - AdsPower Manager v1.0')
        self.root.configure(bg='#1a1a2e')

        # Window size and position from config
        w = int(self.cfg.get('MAIN', 'GUIW', fallback='380'))
        h = int(self.cfg.get('MAIN', 'GUIH', fallback='700'))
        x = int(self.cfg.get('MAIN', 'GUIX', fallback='20'))
        y = int(self.cfg.get('MAIN', 'GUIY', fallback='60'))
        self.root.geometry(f'{w}x{h}+{x}+{y}')
        self.root.minsize(320, 400)
        self.root.resizable(True, True)

        # Always on top
        self.always_on_top = self.cfg.get('MAIN', 'AlwaysOnTop', fallback='1') == '1'
        self.root.attributes('-topmost', self.always_on_top)

        # Dark theme colors
        BG = '#1a1a2e'
        BG2 = '#16213e'
        BG3 = '#0f3460'
        FG = '#e0e0e0'
        FG2 = '#888'
        ACCENT = '#e94560'
        GOLD = '#ffd700'
        GREEN = '#00e676'

        # ── Header ──
        header = tk.Frame(self.root, bg=BG, pady=8)
        header.pack(fill='x', padx=10)

        title_frame = tk.Frame(header, bg=BG)
        title_frame.pack(side='left')
        tk.Label(title_frame, text='APM', font=('Consolas', 18, 'bold'),
                 fg=ACCENT, bg=BG).pack(side='left')
        tk.Label(title_frame, text=' v1.0', font=('Consolas', 10),
                 fg=FG2, bg=BG).pack(side='left', pady=(6, 0))

        btn_frame = tk.Frame(header, bg=BG)
        btn_frame.pack(side='right')

        self.pin_var = tk.StringVar(value='\u25a0' if self.always_on_top else '\u25a1')
        self.pin_btn = tk.Button(btn_frame, text='\u25a0' if self.always_on_top else '\u25a1',
                                  font=('Consolas', 12), fg=GOLD if self.always_on_top else FG2,
                                  bg=BG, bd=0, cursor='hand2',
                                  command=self._toggle_always_on_top)
        self.pin_btn.pack(side='left', padx=4)

        tk.Button(btn_frame, text='\u2699', font=('Consolas', 14), fg=FG2,
                  bg=BG, bd=0, cursor='hand2',
                  command=self._open_settings).pack(side='left', padx=4)

        # ── Status bar ──
        status_frame = tk.Frame(self.root, bg=BG2, pady=4, padx=10)
        status_frame.pack(fill='x', padx=10, pady=(0, 5))

        self.status_dot = tk.Label(status_frame, text='\u25cf', font=('Consolas', 10),
                                    fg='#ff4444', bg=BG2)
        self.status_dot.pack(side='left')
        self.status_label = tk.Label(status_frame, text=' Connecting...',
                                      font=('Consolas', 9), fg=FG2, bg=BG2)
        self.status_label.pack(side='left')
        self.profile_count = tk.Label(status_frame, text='0 profiles',
                                       font=('Consolas', 9), fg=FG2, bg=BG2)
        self.profile_count.pack(side='right')

        # ── Navigation buttons ──
        nav_frame = tk.Frame(self.root, bg=BG, pady=4)
        nav_frame.pack(fill='x', padx=10)

        self.prev_btn = tk.Button(nav_frame, text='\u25c0 PREV', font=('Consolas', 10, 'bold'),
                                   fg='#fff', bg='#7b2ff7', activebackground='#6320d0',
                                   bd=0, padx=12, pady=4, cursor='hand2',
                                   command=lambda: self._cycle_profile(-1))
        self.prev_btn.pack(side='left', expand=True, fill='x', padx=(0, 4))

        self.next_btn = tk.Button(nav_frame, text='NEXT \u25b6', font=('Consolas', 10, 'bold'),
                                   fg='#fff', bg='#7b2ff7', activebackground='#6320d0',
                                   bd=0, padx=12, pady=4, cursor='hand2',
                                   command=lambda: self._cycle_profile(1))
        self.next_btn.pack(side='right', expand=True, fill='x', padx=(4, 0))

        # ── URL bar ──
        url_frame = tk.Frame(self.root, bg=BG, pady=4)
        url_frame.pack(fill='x', padx=10)

        self.url_entry = tk.Entry(url_frame, font=('Consolas', 9), bg=BG2,
                                   fg=FG, insertbackground=FG, bd=1,
                                   relief='solid', highlightthickness=0)
        self.url_entry.pack(side='left', fill='x', expand=True, ipady=4)
        self.url_entry.insert(0, 'https://')

        self.open_btn = tk.Button(url_frame, text='Open All', font=('Consolas', 9, 'bold'),
                                   fg='#fff', bg=GREEN, activebackground='#00c853',
                                   bd=0, padx=8, pady=3, cursor='hand2',
                                   command=self._open_url_in_all)
        self.open_btn.pack(side='right', padx=(6, 0))

        # ── Profile list ──
        list_frame = tk.Frame(self.root, bg=BG)
        list_frame.pack(fill='both', expand=True, padx=10, pady=(5, 10))

        # Canvas + scrollbar for custom profile cards
        self.canvas = tk.Canvas(list_frame, bg=BG, highlightthickness=0, bd=0)
        scrollbar = tk.Scrollbar(list_frame, orient='vertical', command=self.canvas.yview)
        self.scroll_frame = tk.Frame(self.canvas, bg=BG)

        self.scroll_frame.bind('<Configure>',
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox('all')))
        self.canvas.create_window((0, 0), window=self.scroll_frame, anchor='nw')
        self.canvas.configure(yscrollcommand=scrollbar.set)

        self.canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        # Mouse wheel scrolling
        self.canvas.bind('<MouseWheel>',
            lambda e: self.canvas.yview_scroll(int(-1 * (e.delta / 120)), 'units'))
        self.canvas.bind('<Enter>',
            lambda e: self.canvas.bind_all('<MouseWheel>',
                lambda ev: self.canvas.yview_scroll(int(-1 * (ev.delta / 120)), 'units')))
        self.canvas.bind('<Leave>',
            lambda e: self.canvas.unbind_all('<MouseWheel>'))

        # ── Bottom bar ──
        bottom = tk.Frame(self.root, bg=BG2, pady=6, padx=10)
        bottom.pack(fill='x', side='bottom')

        tk.Button(bottom, text='\U0001f4f7', font=('Segoe UI Emoji', 11),
                  fg=FG, bg=BG2, bd=0, cursor='hand2',
                  command=self._take_screenshot).pack(side='left', padx=4)
        tk.Button(bottom, text='\U0001f4e4', font=('Segoe UI Emoji', 11),
                  fg=FG, bg=BG2, bd=0, cursor='hand2',
                  command=self._push_to_discord).pack(side='left', padx=4)
        self.last_update_label = tk.Label(bottom, text='--:--:--',
                                           font=('Consolas', 9), fg=FG2, bg=BG2)
        self.last_update_label.pack(side='right')

        # Save position on close
        self.root.protocol('WM_DELETE_WINDOW', self._on_close)

    def _toggle_always_on_top(self):
        self.always_on_top = not self.always_on_top
        self.root.attributes('-topmost', self.always_on_top)
        self.pin_btn.configure(
            text='\u25a0' if self.always_on_top else '\u25a1',
            fg='#ffd700' if self.always_on_top else '#888'
        )
        self.cfg.set('MAIN', 'AlwaysOnTop', '1' if self.always_on_top else '0')
        save_config(self.cfg)

    def _register_hotkeys(self):
        if not keyboard:
            return
        try:
            fwd = self.cfg.get('HOTKEYS', 'FORWARD', fallback='ctrl+shift+right')
            bwd = self.cfg.get('HOTKEYS', 'BACKWARD', fallback='ctrl+shift+left')
            keyboard.add_hotkey(fwd, lambda: self.root.after(0, self._cycle_profile, 1),
                                suppress=False)
            keyboard.add_hotkey(bwd, lambda: self.root.after(0, self._cycle_profile, -1),
                                suppress=False)
        except Exception as e:
            print(f'Hotkey registration failed: {e}')

    def _cycle_profile(self, direction):
        if not self.profiles:
            return
        if self.selected_index < 0:
            self.selected_index = 0
        else:
            self.selected_index = (self.selected_index + direction) % len(self.profiles)
        profile = self.profiles[self.selected_index]
        self._switch_to_profile(profile)
        self._render_profiles()

    def _switch_to_profile(self, profile):
        """Bring the profile's browser window to the foreground and position it."""
        if not HAS_WIN32:
            return
        debug_port = profile.get('debug_port', 0)
        if not debug_port:
            return

        hwnd = find_profile_window(debug_port)
        if not hwnd:
            return

        # Get APM window rect to position browser next to it
        try:
            apm_hwnd = self.root.winfo_id()
            # Get the top-level window handle
            import ctypes
            apm_hwnd = ctypes.windll.user32.GetParent(apm_hwnd)
            if apm_hwnd:
                apm_rect = win32gui.GetWindowRect(apm_hwnd)
            else:
                apm_rect = (0, 0, 400, 700)
        except Exception:
            apm_rect = (0, 0, 400, 700)

        position_profile_window(hwnd, apm_rect)

    def _open_url_in_all(self):
        url = self.url_entry.get().strip()
        if not url or url == 'https://':
            return
        count = 0
        for profile in self.profiles:
            port = profile.get('debug_port', 0)
            if port and cdp_open_tab(port, url):
                count += 1
        self.status_label.configure(text=f' Opened in {count}/{len(self.profiles)} profiles')

    def _render_profiles(self):
        """Render profile cards in the scrollable list."""
        # Clear existing
        for widget in self.scroll_frame.winfo_children():
            widget.destroy()

        BG = '#1a1a2e'
        BG_CARD = '#16213e'
        BG_SELECTED = '#0f3460'
        FG = '#e0e0e0'
        FG2 = '#888'
        ACCENT = '#e94560'
        GOLD = '#ffd700'
        GREEN = '#00e676'
        CYAN = '#00bcd4'

        canvas_width = self.canvas.winfo_width() or 360

        for i, profile in enumerate(self.profiles):
            serial = profile.get('serial', '') or profile.get('custom', '') or profile.get('uid', '')
            name = profile.get('name', serial)
            is_selected = (i == self.selected_index)
            bg = BG_SELECTED if is_selected else BG_CARD

            card = tk.Frame(self.scroll_frame, bg=bg, pady=6, padx=8, cursor='hand2')
            card.pack(fill='x', pady=2, padx=2)
            card.bind('<Button-1>', lambda e, idx=i: self._on_profile_click(idx))

            # Top row: serial/name + queue number
            top_row = tk.Frame(card, bg=bg)
            top_row.pack(fill='x')
            top_row.bind('<Button-1>', lambda e, idx=i: self._on_profile_click(idx))

            serial_display = f'#{serial}' if serial else name
            serial_lbl = tk.Label(top_row, text=serial_display, font=('Consolas', 11, 'bold'),
                                   fg=FG, bg=bg, anchor='w')
            serial_lbl.pack(side='left')
            serial_lbl.bind('<Button-1>', lambda e, idx=i: self._on_profile_click(idx))

            # Queue number
            qdata = self.queue_data.get(serial, {})
            q_num = qdata.get('queue_number')
            if q_num:
                q_color = GREEN if q_num < 100 else GOLD if q_num < 1000 else ACCENT
                q_lbl = tk.Label(top_row, text=f'Q#{q_num:,}', font=('Consolas', 11, 'bold'),
                                  fg=q_color, bg=bg)
                q_lbl.pack(side='right')
                q_lbl.bind('<Button-1>', lambda e, idx=i: self._on_profile_click(idx))

            # Bottom row: event + link
            event = qdata.get('event', '')
            url = qdata.get('url', '')

            if event or url:
                bot_row = tk.Frame(card, bg=bg)
                bot_row.pack(fill='x', pady=(2, 0))
                bot_row.bind('<Button-1>', lambda e, idx=i: self._on_profile_click(idx))

                if event:
                    ev_lbl = tk.Label(bot_row, text=event[:40], font=('Consolas', 8),
                                       fg=CYAN, bg=bg, anchor='w')
                    ev_lbl.pack(side='left')
                    ev_lbl.bind('<Button-1>', lambda e, idx=i: self._on_profile_click(idx))

                if url:
                    url_short = url[:60] + '...' if len(url) > 60 else url
                    url_lbl = tk.Label(bot_row, text=url_short, font=('Consolas', 7),
                                        fg=FG2, bg=bg, anchor='e')
                    url_lbl.pack(side='right')
                    url_lbl.bind('<Button-1>', lambda e, idx=i: self._on_profile_click(idx))

            # Selection indicator
            if is_selected:
                indicator = tk.Frame(card, bg=ACCENT, width=3)
                indicator.place(x=0, y=0, relheight=1)

    def _on_profile_click(self, index):
        self.selected_index = index
        if 0 <= index < len(self.profiles):
            self._switch_to_profile(self.profiles[index])
        self._render_profiles()

    def _start_polling(self):
        """Start background thread to poll AdsPower API and scan queues."""
        def poll_loop():
            while self.running:
                try:
                    profiles = self.api.get_active_profiles()
                    self.profiles = profiles

                    # Scan each profile for queue numbers
                    for p in profiles:
                        serial = p.get('serial', '') or p.get('custom', '') or p.get('uid', '')
                        if not serial:
                            continue
                        debug_port = p.get('debug_port', 0)
                        if not debug_port:
                            continue
                        result = cdp_eval_queue(debug_port, timeout=3)
                        if result:
                            self.queue_data[serial] = result

                    # Update GUI on main thread
                    self.root.after(0, self._update_gui, len(profiles))
                except Exception as e:
                    self.root.after(0, self._update_status, False, str(e))

                time.sleep(self.poll_interval)

        t = threading.Thread(target=poll_loop, daemon=True)
        t.start()

    def _update_gui(self, count):
        self._update_status(True, f' Connected')
        self.profile_count.configure(text=f'{count} profile{"s" if count != 1 else ""}')
        self.last_update_label.configure(text=datetime.now().strftime('%H:%M:%S'))
        self._render_profiles()

    def _update_status(self, connected, text=''):
        self.status_dot.configure(fg='#00e676' if connected else '#ff4444')
        if text:
            self.status_label.configure(text=text)

    def _take_screenshot(self):
        folder = self.cfg.get('SCREENSHOTS', 'Folder', fallback='Screenshots')
        if not os.path.isabs(folder):
            folder = os.path.join(BASE_DIR, folder)
        serial = ''
        if 0 <= self.selected_index < len(self.profiles):
            serial = self.profiles[self.selected_index].get('serial', '')
        path = take_screenshot(folder, serial)
        if path:
            self.status_label.configure(text=f' Screenshot saved')

    def _push_to_discord(self):
        webhook = self.cfg.get('DISCORD', 'QueWebhook', fallback='')
        profile_name = self.cfg.get('DISCORD', 'ProfileName', fallback='')
        if not webhook:
            self.status_label.configure(text=' No Discord webhook set')
            return
        lines = []
        for p in self.profiles:
            serial = p.get('serial', '') or p.get('uid', '')
            qdata = self.queue_data.get(serial, {})
            q = qdata.get('queue_number')
            event = qdata.get('event', '')
            if q:
                lines.append(f'#{serial}: Q#{q:,} - {event}')
            else:
                lines.append(f'#{serial}: No queue')
        content = f'**{profile_name or "APM"}** Queue Update:\n' + '\n'.join(lines)
        threading.Thread(target=send_discord_webhook, args=(webhook, content), daemon=True).start()
        self.status_label.configure(text=' Sent to Discord')

    def _open_settings(self):
        """Open settings window."""
        win = tk.Toplevel(self.root)
        win.title('APM Settings')
        win.configure(bg='#1a1a2e')
        win.geometry('400x500')
        win.attributes('-topmost', True)

        BG = '#1a1a2e'
        BG2 = '#16213e'
        FG = '#e0e0e0'
        FG2 = '#888'

        canvas = tk.Canvas(win, bg=BG, highlightthickness=0)
        canvas.pack(fill='both', expand=True, padx=10, pady=10)

        frame = tk.Frame(canvas, bg=BG)
        canvas.create_window((0, 0), window=frame, anchor='nw')

        row = 0
        entries = {}

        sections = [
            ('DISCORD', [
                ('QueWebhook', 'Queue Webhook URL'),
                ('ProdWebhook', 'Prod Webhook URL'),
                ('ProfileName', 'Profile Name'),
            ]),
            ('SHEETS', [
                ('SheetUrl', 'Google Sheets URL'),
            ]),
            ('HOTKEYS', [
                ('FORWARD', 'Forward Hotkey'),
                ('BACKWARD', 'Backward Hotkey'),
            ]),
            ('MAIN', [
                ('PollInterval', 'Poll Interval (sec)'),
                ('AdsPowerPort', 'AdsPower Port'),
            ]),
        ]

        for section_name, fields in sections:
            tk.Label(frame, text=section_name, font=('Consolas', 10, 'bold'),
                     fg='#e94560', bg=BG, anchor='w').grid(
                row=row, column=0, columnspan=2, sticky='w', pady=(10, 4))
            row += 1

            for key, label in fields:
                tk.Label(frame, text=label, font=('Consolas', 9),
                         fg=FG2, bg=BG, anchor='w').grid(
                    row=row, column=0, sticky='w', padx=(10, 5))
                entry = tk.Entry(frame, font=('Consolas', 9), bg=BG2, fg=FG,
                                  insertbackground=FG, bd=1, width=30)
                val = self.cfg.get(section_name, key, fallback='')
                entry.insert(0, val)
                entry.grid(row=row, column=1, sticky='ew', padx=5, pady=2)
                entries[(section_name, key)] = entry
                row += 1

        def save_settings():
            for (section, key), entry in entries.items():
                if not self.cfg.has_section(section):
                    self.cfg.add_section(section)
                self.cfg.set(section, key, entry.get())
            save_config(self.cfg)
            # Re-register hotkeys
            if keyboard:
                keyboard.unhook_all_hotkeys()
                self._register_hotkeys()
            win.destroy()
            self.status_label.configure(text=' Settings saved')

        tk.Button(frame, text='Save', font=('Consolas', 10, 'bold'),
                  fg='#fff', bg='#00e676', activebackground='#00c853',
                  bd=0, padx=20, pady=6, cursor='hand2',
                  command=save_settings).grid(
            row=row, column=0, columnspan=2, pady=20)

    def _on_close(self):
        """Save window position and exit."""
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


if __name__ == '__main__':
    # Check dependencies
    missing = []
    if not requests:
        missing.append('requests')
    if not websocket:
        missing.append('websocket-client')
    if missing:
        print(f'Missing required packages: {", ".join(missing)}')
        print(f'Install with: pip install {" ".join(missing)}')
        input('Press Enter to exit...')
        sys.exit(1)

    app = APMApp()
    app.run()
