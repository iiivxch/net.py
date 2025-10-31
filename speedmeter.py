import os
import sys
import json
import threading
import time
import datetime
import psutil
from dataclasses import dataclass, field
from typing import Tuple, Optional, Dict
import queue
from collections import deque

# GUI and Tray
import tkinter as tk
from tkinter import ttk
from PIL import Image, ImageDraw, ImageFont
import pystray
# Optional: Windows Startup shortcut (pywin32). If unavailable, feature is disabled gracefully.
try:
    import win32com.client as _win32_client
except Exception:
    _win32_client = None

# -------------------------------
# Configuration
# -------------------------------
APP_NAME = "SpeedMeter"
POLL_INTERVAL_SEC = 0.5
ICON_SIZE = 64  # Tray icon render size before OS scaling
DATA_DIR = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), APP_NAME)
DATA_FILE = os.path.join(DATA_DIR, "data.json")
CONFIG_FILE = os.path.join(DATA_DIR, "config.json")

# Windows Startup shortcut helpers
STARTUP_DIR = os.path.join(os.environ.get("APPDATA", ""), r"Microsoft\Windows\Start Menu\Programs\Startup")
STARTUP_SHORTCUT = os.path.join(STARTUP_DIR, f"{APP_NAME}.lnk")

def is_frozen() -> bool:
    return getattr(sys, 'frozen', False) is True

def get_executable_and_args() -> Tuple[str, str]:
    if is_frozen():
        return sys.executable, ""
    # run with python on the script
    script = os.path.abspath(__file__)
    return sys.executable, f'"{script}"'

def enable_autostart_shortcut():
    if _win32_client is None:
        return False
    try:
        os.makedirs(STARTUP_DIR, exist_ok=True)
        shell = _win32_client.Dispatch('WScript.Shell')
        shortcut = shell.CreateShortcut(STARTUP_SHORTCUT)
        target, args = get_executable_and_args()
        shortcut.TargetPath = target
        if args:
            shortcut.Arguments = args
        shortcut.WorkingDirectory = os.path.dirname(target if is_frozen() else os.path.abspath(__file__))
        shortcut.WindowStyle = 7  # Minimized
        shortcut.IconLocation = target
        shortcut.save()
        return True
    except Exception:
        return False

def disable_autostart_shortcut():
    try:
        if os.path.exists(STARTUP_SHORTCUT):
            os.remove(STARTUP_SHORTCUT)
    except Exception:
        pass

def is_autostart_enabled() -> bool:
    try:
        return os.path.exists(STARTUP_SHORTCUT)
    except Exception:
        return False

# Try to pick a reasonably legible font
def get_font(size: int) -> ImageFont.FreeTypeFont:
    # Try common fonts
    candidates = [
        "C:\\Windows\\Fonts\\segoeuib.ttf",  # Segoe UI Semibold
        "C:\\Windows\\Fonts\\segoeui.ttf",
        "C:\\Windows\\Fonts\\arial.ttf",
    ]
    for path in candidates:
        if os.path.exists(path):
            try:
                return ImageFont.truetype(path, size)
            except Exception:
                pass
    # Fallback
    return ImageFont.load_default()

# -------------------------------
# Utility
# -------------------------------
def human_readable_speed(bytes_per_sec: float) -> str:
    # Show K, M, G with one decimal
    units = ["B", "K", "M", "G", "T"]
    v = float(bytes_per_sec)
    idx = 0
    while v >= 1024.0 and idx < len(units) - 1:
        v /= 1024.0
        idx += 1
    if idx == 0:
        return f"{int(v)}{units[idx]}"
    return f"{v:.1f}{units[idx]}"

# General speed formatter based on unit preference
def format_speed(bytes_per_sec: float, unit: str) -> str:
    # Support exactly three units: MBps (MB/s), Mbps, Kbps
    key = (unit or "MBps").strip()
    if key == "MBps":
        return f"{(bytes_per_sec) / 1_000_000.0:.2f} MB/s"
    if key == "Mbps":
        return f"{(bytes_per_sec * 8.0) / 1_000_000.0:.2f} Mbps"
    if key == "Kbps":
        return f"{(bytes_per_sec * 8.0) / 1_000.0:.0f} Kbps"
    # Fallback to MB/s if an unknown unit is passed
    return f"{(bytes_per_sec) / 1_000_000.0:.2f} MB/s"

# -------------------------------
# Utility (date keys and data dir)
# -------------------------------
def today_key(dt: Optional[datetime.datetime] = None) -> str:
    dt = dt or datetime.datetime.now()
    return dt.strftime("%Y-%m-%d")

def month_key(dt: Optional[datetime.datetime] = None) -> str:
    dt = dt or datetime.datetime.now()
    return dt.strftime("%Y-%m")

def ensure_data_dir():
    os.makedirs(DATA_DIR, exist_ok=True)

# -------------------------------
# Config persistence
# -------------------------------
class Config:
    def __init__(self):
        # Defaults
        self.auto_show_overlay: bool = False
        self.overlay_pos: Tuple[int, int] = (0, 0)
        self.theme: str = "system"  # 'system' | 'light' | 'dark'
        self.units: str = "MBps"  # default to MB/s per user request
        self.smoothing_seconds: float = 0.0  # 0 disables EMA
        self.nic_name: Optional[str] = None  # None => auto
        self.overlay_transparent: bool = True  # allow disabling transparency if overlay seems invisible
        self.overlay_compact: bool = False  # compact size and font for the overlay
        self.overlay_opacity: float = 0.7  # 0..1, default 70%
        self.autostart: bool = False  # create a Startup shortcut when enabled

    def load(self):
        ensure_data_dir()
        if not os.path.exists(CONFIG_FILE):
            return
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
            self.auto_show_overlay = bool(data.get('auto_show_overlay', self.auto_show_overlay))
            pos = data.get('overlay_pos', list(self.overlay_pos))
            if isinstance(pos, list) and len(pos) == 2:
                self.overlay_pos = (int(pos[0]), int(pos[1]))
            self.theme = str(data.get('theme', self.theme))
            self.units = str(data.get('units', self.units))
            self.smoothing_seconds = float(data.get('smoothing_seconds', self.smoothing_seconds))
            self.nic_name = data.get('nic_name', self.nic_name)
            self.overlay_transparent = bool(data.get('overlay_transparent', self.overlay_transparent))
            self.overlay_compact = bool(data.get('overlay_compact', self.overlay_compact))
            try:
                self.overlay_opacity = float(data.get('overlay_opacity', self.overlay_opacity))
            except Exception:
                pass
            self.autostart = bool(data.get('autostart', self.autostart))
        except Exception:
            pass

    def save(self):
        ensure_data_dir()
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump({
                    'auto_show_overlay': self.auto_show_overlay,
                    'overlay_pos': list(self.overlay_pos),
                    'theme': self.theme,
                    'units': self.units,
                    'smoothing_seconds': self.smoothing_seconds,
                    'nic_name': self.nic_name,
                    'overlay_transparent': self.overlay_transparent,
                    'overlay_compact': self.overlay_compact,
                    'overlay_opacity': self.overlay_opacity,
                    'autostart': self.autostart,
                }, f)
        except Exception:
            pass

# -------------------------------
# Data Models
# -------------------------------
@dataclass
class UsageTotals:
    download_bytes: int = 0
    upload_bytes: int = 0

@dataclass
class UsageStore:
    # Store usage by day: { "YYYY-MM-DD": {"down": int, "up": int} }
    by_day: Dict[str, Dict[str, int]] = field(default_factory=dict)

    def add_usage(self, dt: datetime.datetime, down_delta: int, up_delta: int):
        dkey = today_key(dt)
        if dkey not in self.by_day:
            self.by_day[dkey] = {"down": 0, "up": 0}
        self.by_day[dkey]["down"] += max(0, down_delta)
        self.by_day[dkey]["up"] += max(0, up_delta)

    def get_today(self, dt: Optional[datetime.datetime] = None) -> UsageTotals:
        dkey = today_key(dt)
        if dkey in self.by_day:
            d = self.by_day[dkey]
            return UsageTotals(d.get("down", 0), d.get("up", 0))
        return UsageTotals(0, 0)

    def get_month(self, dt: Optional[datetime.datetime] = None) -> UsageTotals:
        mkey = month_key(dt)
        total_down = 0
        total_up = 0
        for day, vals in self.by_day.items():
            if day.startswith(mkey):
                total_down += vals.get("down", 0)
                total_up += vals.get("up", 0)
        return UsageTotals(total_down, total_up)

    def to_json(self) -> dict:
        return {"by_day": self.by_day}

    @staticmethod
    def from_json(data: dict) -> "UsageStore":
        return UsageStore(by_day=data.get("by_day", {}))

    def clear_today(self, dt: Optional[datetime.datetime] = None):
        dkey = today_key(dt)
        if dkey in self.by_day:
            self.by_day[dkey] = {"down": 0, "up": 0}

# -------------------------------
# Persistence
# -------------------------------
class UsageRepository:
    def __init__(self, path: str):
        self.path = path
        self.lock = threading.Lock()
        ensure_data_dir()

    def load(self) -> UsageStore:
        with self.lock:
            if not os.path.exists(self.path):
                return UsageStore()
            try:
                with open(self.path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                return UsageStore.from_json(data)
            except Exception:
                return UsageStore()

    def save(self, store: UsageStore):
        with self.lock:
            tmp = self.path + ".tmp"
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(store.to_json(), f)
            os.replace(tmp, self.path)

# -------------------------------
# Network Monitor
# -------------------------------
class NetMonitor(threading.Thread):
    def __init__(self, repo: UsageRepository, on_update=None, interval=POLL_INTERVAL_SEC, config: Optional[Config]=None):
        super().__init__(daemon=True)
        self.repo = repo
        self.on_update = on_update
        self.interval = interval
        self._stop_evt = threading.Event()
        # Track last counters per active interface
        self._last_pernic: Dict[str, Tuple[int, int]] = {}
        self._included_ifaces: set[str] = set()
        self._last_iface_refresh = 0.0
        self.current_down_bps = 0.0
        self.current_up_bps = 0.0

        # EMA smoothing
        self._ema_down_bps = None
        self._ema_up_bps = None
        self.config = config or Config()

        # Rolling average windows (10s)
        self._hist_down = deque()  # (timestamp, bps)
        self._hist_up = deque()    # (timestamp, bps)
        self.avg_down_bps_10s = 0.0
        self.avg_up_bps_10s = 0.0

        self.store = self.repo.load()

    def stop(self):
        self._stop_evt.set()

    def run(self):
        last_time = time.time()
        # Initialize interface set
        self._refresh_ifaces()
        while not self._stop_evt.is_set():
            time.sleep(self.interval)
            now = time.time()
            elapsed = max(1e-6, now - last_time)
            last_time = now

            # Periodically refresh included interfaces (e.g., every 10 seconds)
            if now - self._last_iface_refresh > 10.0:
                self._refresh_ifaces()
                self._last_iface_refresh = now

            pernic = psutil.net_io_counters(pernic=True)
            down_delta = 0
            up_delta = 0
            for name in self._included_ifaces:
                c = pernic.get(name)
                if not c:
                    continue
                prev = self._last_pernic.get(name, (c.bytes_recv, c.bytes_sent))
                d_down = max(0, c.bytes_recv - prev[0])
                d_up = max(0, c.bytes_sent - prev[1])
                self._last_pernic[name] = (c.bytes_recv, c.bytes_sent)
                down_delta += d_down
                up_delta += d_up

            # Bytes/sec (raw)
            raw_down_bps = max(0.0, down_delta / elapsed)
            raw_up_bps = max(0.0, up_delta / elapsed)

            # Smoothing via EMA if enabled
            s = max(0.0, float(self.config.smoothing_seconds or 0.0))
            if s > 0:
                alpha = min(1.0, max(0.01, self.interval / s))
                if self._ema_down_bps is None:
                    self._ema_down_bps = raw_down_bps
                    self._ema_up_bps = raw_up_bps
                else:
                    self._ema_down_bps = (alpha * raw_down_bps) + (1 - alpha) * self._ema_down_bps
                    self._ema_up_bps = (alpha * raw_up_bps) + (1 - alpha) * self._ema_up_bps
                self.current_down_bps = self._ema_down_bps
                self.current_up_bps = self._ema_up_bps
            else:
                self.current_down_bps = raw_down_bps
                self.current_up_bps = raw_up_bps

            # Update 10s rolling averages using raw instantaneous values
            # Keep last 10 seconds of history
            tcut = now - 10.0
            self._hist_down.append((now, raw_down_bps))
            self._hist_up.append((now, raw_up_bps))
            while self._hist_down and self._hist_down[0][0] < tcut:
                self._hist_down.popleft()
            while self._hist_up and self._hist_up[0][0] < tcut:
                self._hist_up.popleft()
            if self._hist_down:
                self.avg_down_bps_10s = sum(v for _, v in self._hist_down) / len(self._hist_down)
            else:
                self.avg_down_bps_10s = 0.0
            if self._hist_up:
                self.avg_up_bps_10s = sum(v for _, v in self._hist_up) / len(self._hist_up)
            else:
                self.avg_up_bps_10s = 0.0

            # Persist usage
            dt = datetime.datetime.now()
            self.store.add_usage(dt, down_delta, up_delta)
            # Save occasionally to reduce writes
            if int(now) % 10 == 0:
                self.repo.save(self.store)

            if self.on_update:
                try:
                    self.on_update(self.current_down_bps, self.current_up_bps, self.store)
                except Exception:
                    pass

        # Final save on exit
        try:
            self.repo.save(self.store)
        except Exception:
            pass

    def _refresh_ifaces(self):
        # Select active, non-virtual, non-loopback interfaces
        stats = psutil.net_if_stats()
        all_counters = psutil.net_io_counters(pernic=True)
        exclude_prefixes = (
            'Loopback', 'Software Loopback', 'isatap', 'Teredo', 'vEthernet', 'VMware',
            'VirtualBox', 'Npcap', 'NPF_', 'WAN Miniport', 'Bluetooth', 'Hyper-V'
        )
        included = set()
        for name, st in stats.items():
            if not st.isup:
                continue
            if name.startswith(exclude_prefixes):
                continue
            if name in all_counters:
                included.add(name)
        # If a specific NIC is configured, restrict to that
        if self.config.nic_name and self.config.nic_name in included:
            included = {self.config.nic_name}

        # Initialize last counters for any new iface
        for name in included:
            c = all_counters.get(name)
            if c and name not in self._last_pernic:
                self._last_pernic[name] = (c.bytes_recv, c.bytes_sent)
        # Drop removed ifaces
        for name in list(self._last_pernic.keys()):
            if name not in included:
                self._last_pernic.pop(name, None)
        self._included_ifaces = included

# -------------------------------
# UI Thread Host (single Tk root)
# -------------------------------
class UIThread:
    def __init__(self):
        self._thread: Optional[threading.Thread] = None
        self._ready = threading.Event()
        self._queue: "queue.Queue[callable]" = queue.Queue()
        self.root: Optional[tk.Tk] = None
        self._running = False

    def start(self):
        if self._thread and self._thread.is_alive():
            return
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()
        self._ready.wait(timeout=5)

    def call(self, fn):
        if not self._running:
            return
        self._queue.put(fn)

    def stop(self):
        def _quit():
            if self.root:
                try:
                    self.root.quit()
                except Exception:
                    pass
        self.call(_quit)

    def _pump(self):
        try:
            while True:
                fn = self._queue.get_nowait()
                try:
                    fn()
                except Exception:
                    pass
        except queue.Empty:
            pass
        if self.root:
            self.root.after(50, self._pump)

    def _run(self):
        self.root = tk.Tk()
        self.root.withdraw()
        self._running = True
        self._ready.set()
        self._pump()
        self.root.mainloop()
        self._running = False

# -------------------------------
# Tray Icon Rendering
# -------------------------------
class TrayIconRenderer:
    def __init__(self):
        self.font_small = get_font(18)
        self.font_tiny = get_font(16)

    def render(self, down_bps: float, up_bps: float, unit: str) -> Image.Image:
        # Create icon with two-line text: D: x  / U: y
        # Use colors to differentiate
        img = Image.new("RGBA", (ICON_SIZE, ICON_SIZE), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)

        # Background rounded rectangle for contrast
        radius = 12
        bg_color = (25, 25, 25, 255)
        draw.rounded_rectangle([0, 0, ICON_SIZE - 1, ICON_SIZE - 1], radius, fill=bg_color)

        # Text lines
        d_text = f"D {format_speed(down_bps, unit)}"
        u_text = f"U {format_speed(up_bps, unit)}"

        # Colors
        d_color = (80, 200, 120, 255)   # green-ish
        u_color = (255, 165, 0, 255)    # orange

        # Measure and place centered
        d_bbox = draw.textbbox((0, 0), d_text, font=self.font_small)
        d_w, d_h = d_bbox[2] - d_bbox[0], d_bbox[3] - d_bbox[1]
        u_bbox = draw.textbbox((0, 0), u_text, font=self.font_small)
        u_w, u_h = u_bbox[2] - u_bbox[0], u_bbox[3] - u_bbox[1]

        # Margins
        top = 6
        spacing = 2
        # Adjust if text overflows; fallback to tiny font
        if d_w > ICON_SIZE - 8 or u_w > ICON_SIZE - 8:
            d_bbox = draw.textbbox((0, 0), d_text, font=self.font_tiny)
            d_w, d_h = d_bbox[2] - d_bbox[0], d_bbox[3] - d_bbox[1]
            u_bbox = draw.textbbox((0, 0), u_text, font=self.font_tiny)
            u_w, u_h = u_bbox[2] - u_bbox[0], u_bbox[3] - u_bbox[1]
            font_d = self.font_tiny
            font_u = self.font_tiny
        else:
            font_d = self.font_small
            font_u = self.font_small

        d_x = (ICON_SIZE - d_w) // 2
        u_x = (ICON_SIZE - u_w) // 2
        d_y = top
        u_y = d_y + d_h + spacing

        draw.text((d_x, d_y), d_text, font=font_d, fill=d_color)
        draw.text((u_x, u_y), u_text, font=font_u, fill=u_color)

        # Outline
        draw.rounded_rectangle([0, 0, ICON_SIZE - 1, ICON_SIZE - 1], radius, outline=(60, 60, 60, 255), width=1)

        return img

# -------------------------------
# Dashboard Window
# -------------------------------
class Dashboard:
    def __init__(self, monitor: NetMonitor, ui: UIThread):
        self.monitor = monitor
        self.ui = ui
        self.win: Optional[tk.Toplevel] = None
        self.labels: Dict[str, ttk.Label] = {}

    def open(self):
        def _open():
            if self.win and tk.Toplevel.winfo_exists(self.win):
                try:
                    self.win.deiconify()
                    self.win.lift()
                    return
                except Exception:
                    pass
            self.win = tk.Toplevel(self.ui.root)
            self.win.title(f"{APP_NAME} - Dashboard")
            self.win.geometry("420x240")
            self.win.protocol("WM_DELETE_WINDOW", self._on_close)

            frm = ttk.Frame(self.win, padding=12)
            frm.pack(fill="both", expand=True)

            lbl_title = ttk.Label(frm, text="Real-time Speed", font=("Segoe UI", 12, "bold"))
            lbl_title.pack(anchor="w")

            self.labels["down"] = ttk.Label(frm, text="Download: 0 Mbps", font=("Segoe UI", 11))
            self.labels["down"].pack(anchor="w", pady=(4, 0))

            self.labels["up"] = ttk.Label(frm, text="Upload: 0 Mbps", font=("Segoe UI", 11))
            self.labels["up"].pack(anchor="w")

            ttk.Separator(frm, orient="horizontal").pack(fill="x", pady=8)

            lbl_usage = ttk.Label(frm, text="Data Usage", font=("Segoe UI", 12, "bold"))
            lbl_usage.pack(anchor="w")

            self.labels["today"] = ttk.Label(frm, text="Today: Down 0, Up 0, Total 0", font=("Segoe UI", 11))
            self.labels["today"].pack(anchor="w", pady=(4, 0))

            self.labels["month"] = ttk.Label(frm, text="This Month: Down 0, Up 0, Total 0", font=("Segoe UI", 11))
            self.labels["month"].pack(anchor="w")

            self._schedule_update()
        self.ui.start()
        self.ui.call(_open)

    def _on_close(self):
        if self.win is not None:
            self.win.withdraw()

    def _update_loop(self):
        # Update labels from monitor state
        unit = getattr(self.monitor.config, 'units', 'Mbps')
        d = format_speed(self.monitor.current_down_bps, unit)
        u = format_speed(self.monitor.current_up_bps, unit)
        d_avg = format_speed(self.monitor.avg_down_bps_10s, unit)
        u_avg = format_speed(self.monitor.avg_up_bps_10s, unit)
        if "down" in self.labels:
            self.labels["down"].configure(text=f"Download: {d}  (avg10: {d_avg})")
        if "up" in self.labels:
            self.labels["up"].configure(text=f"Upload:   {u}  (avg10: {u_avg})")

        today = self.monitor.store.get_today()
        month = self.monitor.store.get_month()

        def fmt_bytes(n: int) -> str:
            units = ["B", "KB", "MB", "GB", "TB"]
            v = float(n)
            idx = 0
            while v >= 1024.0 and idx < len(units) - 1:
                v /= 1024.0
                idx += 1
            if idx == 0:
                return f"{int(v)} {units[idx]}"
            return f"{v:.2f} {units[idx]}"

        t_total = today.download_bytes + today.upload_bytes
        m_total = month.download_bytes + month.upload_bytes

        if "today" in self.labels:
            self.labels["today"].configure(
                text=f"Today: Down {fmt_bytes(today.download_bytes)}, Up {fmt_bytes(today.upload_bytes)}, Total {fmt_bytes(t_total)}"
            )
        if "month" in self.labels:
            self.labels["month"].configure(
                text=f"This Month: Down {fmt_bytes(month.download_bytes)}, Up {fmt_bytes(month.upload_bytes)}, Total {fmt_bytes(m_total)}"
            )

        if self.win is not None and self.ui.root is not None:
            self.ui.root.after(1000, self._update_loop)

    def _schedule_update(self):
        self.ui.call(self._update_loop)

# -------------------------------
# Mini Overlay (always-on-top, frameless)
# -------------------------------
class MiniOverlay:
    def __init__(self, monitor: NetMonitor, ui: UIThread, config: Config):
        self.monitor = monitor
        self.ui = ui
        self.config = config
        self.win: Optional[tk.Toplevel] = None
        self.canvas: Optional[tk.Canvas] = None
        self.visible = False
        self.width = 300
        self.height = 44
        # Theme colors
        self._apply_theme()
        self.round = 14
        self.offset_x = 20
        self.offset_y = 80
        self._drag = {"x": 0, "y": 0}
        self._ctx_menu: Optional[tk.Menu] = None

    def _apply_theme(self):
        theme = (self.monitor.config.theme or 'system').lower()
        if theme == 'light':
            self.bg = "#F2F3F5"
            self.fg_down = "#0B8A3D"
            self.fg_up = "#C05A00"
            self.fg_text = "#111111"
        else:  # system/dark => dark as default
            self.bg = "#202225"
            self.fg_down = "#50C878"
            self.fg_up = "#FFA500"
            self.fg_text = "#FFFFFF"

    def _apply_orientation_size(self):
        # Horizontal-only layout; adjust for compact mode
        if getattr(self.monitor.config, 'overlay_compact', False):
            self.width = 200
            self.height = 28
            self.round = 10
        else:
            self.width = 300
            self.height = 44
            self.round = 14

    def _apply_opacity(self):
        # Apply window opacity using Tk attribute, clamped between 0.2 and 1.0
        try:
            alpha = float(self.config.overlay_opacity)
            alpha = max(0.2, min(alpha, 1.0))
            self.win.attributes('-alpha', alpha)
        except Exception:
            pass

    def open(self):
        if self.visible:
            return
        self.visible = True
        self.ui.start()
        self.ui.call(self._run_window)

    def close(self):
        self.visible = False
        if self.win:
            try:
                self.win.destroy()
            except Exception:
                pass
            self.win = None

    def toggle(self):
        if self.visible:
            self.close()
        else:
            self.open()

    def _run_window(self):
        # Adjust size based on orientation before creating window
        self._apply_orientation_size()
        self.win = tk.Toplevel(self.ui.root)
        self.win.overrideredirect(True)
        self.win.attributes("-topmost", True)
        # Apply transparency based on config; when disabled, use opaque background
        if self.config.overlay_transparent:
            self.win.attributes("-transparentcolor", "#010101")
        # Apply default opacity
        self._apply_opacity()

        sw = self.win.winfo_screenwidth()
        sh = self.win.winfo_screenheight()
        # Use persisted position if set
        px, py = self.config.overlay_pos
        if px == 0 and py == 0:
            x = max(0, sw - self.width - self.offset_x)
            y = max(0, sh - self.height - self.offset_y)
        else:
            x, y = px, py
        # Clamp inside visible bounds
        x = max(0, min(x, sw - self.width))
        y = max(0, min(y, sh - self.height))
        self.win.geometry(f"{self.width}x{self.height}+{x}+{y}")

        canvas_bg = "#010101" if self.config.overlay_transparent else self.bg
        self.canvas = tk.Canvas(self.win, width=self.width, height=self.height, highlightthickness=0, bd=0, bg=canvas_bg)
        self.canvas.pack(fill="both", expand=True)

        # Dragging support
        self.canvas.bind("<Button-1>", self._start_move)
        self.canvas.bind("<B1-Motion>", self._on_move)
        self.canvas.bind("<Button-3>", self._on_right_click)
        self.canvas.bind("<ButtonRelease-1>", self._on_release)

        self._redraw()
        if self.ui.root is not None:
            self.ui.root.after(500, self._tick)
        try:
            self.win.lift()
        except Exception:
            pass

    def _start_move(self, event):
        self._drag["x"] = event.x
        self._drag["y"] = event.y

    def _on_move(self, event):
        if not self.win:
            return
        x = self.win.winfo_x() + (event.x - self._drag["x"]) if hasattr(event, 'x') else self.win.winfo_x()
        y = self.win.winfo_y() + (event.y - self._drag["y"]) if hasattr(event, 'y') else self.win.winfo_y()
        self.win.geometry(f"+{x}+{y}")

    def _on_right_click(self, event):
        # Right click closes the overlay; can be reopened from tray
        self.close()

    def _on_release(self, event):
        # Persist new position
        try:
            self.config.overlay_pos = (self.win.winfo_x(), self.win.winfo_y())
            self.config.save()
        except Exception:
            pass

    def _rounded_rect(self, x1, y1, x2, y2, r, color):
        # Draw a rounded rectangle on the canvas
        if not self.canvas:
            return
        try:
            self.canvas.create_rectangle(x1+r, y1, x2-r, y2, fill=color, outline=color)
            self.canvas.create_rectangle(x1, y1+r, x2, y2-r, fill=color, outline=color)
            self.canvas.create_oval(x1, y1, x1+2*r, y1+2*r, fill=color, outline=color)
            self.canvas.create_oval(x2-2*r, y1, x2, y1+2*r, fill=color, outline=color)
            self.canvas.create_oval(x1, y2-2*r, x1+2*r, y2, fill=color, outline=color)
            self.canvas.create_oval(x2-2*r, y2-2*r, x2, y2, fill=color, outline=color)
        except Exception:
            # Fallback to simple rectangle if anything goes wrong
            self.canvas.create_rectangle(x1, y1, x2, y2, fill=color, outline=color)

    def _redraw(self):
        if not self.visible or not self.win:
            return
        try:
            self.canvas.delete("all")
            # Background
            self._rounded_rect(1, 1, self.width-1, self.height-1, self.round, self.bg)
            unit = getattr(self.monitor.config, 'units', 'Mbps')
            d = format_speed(self.monitor.current_down_bps, unit)
            u = format_speed(self.monitor.current_up_bps, unit)
            # Single-line, side-by-side (horizontal only)
            mid_y = self.height // 2
            if getattr(self.monitor.config, 'overlay_compact', False):
                pad = 8
                font = ("Segoe UI", 9, "bold")
            else:
                pad = 12
                font = ("Segoe UI", 11, "bold")
            self.canvas.create_text(pad, mid_y, text=f"D {d}", anchor="w", fill=self.fg_down, font=font)
            self.canvas.create_text(self.width-pad, mid_y, text=f"U {u}", anchor="e", fill=self.fg_up, font=font)
        except Exception:
            # Ignore drawing errors to keep the ticker running
            pass

    def _tick(self):
        if not self.visible or not self.win:
            return
        self._redraw()
        if self.ui.root is not None:
            self.ui.root.after(500, self._tick)

# -------------------------------
# Application
# -------------------------------
class SpeedMeterApp:
    def __init__(self):
        self.repo = UsageRepository(DATA_FILE)
        self.config = Config()
        self.config.load()
        self.monitor = NetMonitor(self.repo, on_update=self._on_monitor_update, interval=POLL_INTERVAL_SEC, config=self.config)
        self.tray_icon: Optional[pystray.Icon] = None
        self.renderer = TrayIconRenderer()
        self.icon_image = self.renderer.render(0.0, 0.0, self.config.units)
        self.ui = UIThread()
        self.dashboard = Dashboard(self.monitor, self.ui)
        self.overlay = MiniOverlay(self.monitor, self.ui, self.config)
        self._tray_update_lock = threading.Lock()

    def _on_monitor_update(self, down_bps: float, up_bps: float, store: UsageStore):
        # Update tray icon image and tooltip
        with self._tray_update_lock:
            try:
                self.icon_image = self.renderer.render(down_bps, up_bps, self.config.units)
                if self.tray_icon is not None:
                    # update icon and tooltip
                    self.tray_icon.icon = self.icon_image
                    self.tray_icon.title = (
                        f"D {format_speed(down_bps, self.config.units)} (avg {format_speed(self.monitor.avg_down_bps_10s, self.config.units)}) | "
                        f"U {format_speed(up_bps, self.config.units)} (avg {format_speed(self.monitor.avg_up_bps_10s, self.config.units)})"
                    )
            except Exception:
                pass

    def _on_open_dashboard(self, icon, item):
        # Called from tray menu
        threading.Thread(target=self.dashboard.open, daemon=True).start()

    def _on_preferences(self, icon, item):
        # Open Preferences window in UI thread
        def _open_prefs():
            win = tk.Toplevel(self.ui.root)
            win.title(f"{APP_NAME} - Preferences")
            win.geometry("420x360")
            frm = ttk.Frame(win, padding=12)
            frm.pack(fill="both", expand=True)

            # Units
            ttk.Label(frm, text="Units", font=("Segoe UI", 10, "bold")).grid(row=0, column=0, sticky="w")
            units_var = tk.StringVar(value=self.config.units)
            ttk.Combobox(frm, textvariable=units_var, values=["MBps", "Mbps", "Kbps"], state="readonly").grid(row=0, column=1, sticky="ew")

            # Smoothing
            ttk.Label(frm, text="Smoothing (seconds)").grid(row=1, column=0, sticky="w", pady=(6,0))
            smooth_var = tk.StringVar(value=str(self.config.smoothing_seconds or 0))
            ttk.Entry(frm, textvariable=smooth_var, width=10).grid(row=1, column=1, sticky="w", pady=(6,0))

            # Theme
            ttk.Label(frm, text="Theme", font=("Segoe UI", 10, "bold")).grid(row=2, column=0, sticky="w", pady=(10,0))
            theme_var = tk.StringVar(value=self.config.theme)
            ttk.Combobox(frm, textvariable=theme_var, values=["system", "dark", "light"], state="readonly").grid(row=2, column=1, sticky="ew", pady=(10,0))

            # Auto show overlay
            auto_var = tk.BooleanVar(value=self.config.auto_show_overlay)
            ttk.Checkbutton(frm, text="Auto show overlay on startup", variable=auto_var).grid(row=3, column=0, columnspan=2, sticky="w", pady=(10,0))

            # NIC selector
            ttk.Label(frm, text="Network Interface", font=("Segoe UI", 10, "bold")).grid(row=4, column=0, sticky="w", pady=(10,0))
            nic_var = tk.StringVar(value=self.config.nic_name or "(All)")
            try:
                nics = ["(All)"] + [n for n, st in psutil.net_if_stats().items() if st.isup]
            except Exception:
                nics = ["(All)"]
            ttk.Combobox(frm, textvariable=nic_var, values=nics, state="readonly").grid(row=4, column=1, sticky="ew")

            # Transparency toggle
            trans_var = tk.BooleanVar(value=self.config.overlay_transparent)
            ttk.Checkbutton(frm, text="Transparent corners (recommended)", variable=trans_var).grid(row=5, column=0, columnspan=2, sticky="w", pady=(10,0))

            # Compact mode
            compact_var = tk.BooleanVar(value=self.config.overlay_compact)
            ttk.Checkbutton(frm, text="Compact overlay (smaller)", variable=compact_var).grid(row=6, column=0, columnspan=2, sticky="w", pady=(6,0))

            # Opacity
            opacity_var = tk.StringVar(value=str(self.config.overlay_opacity))
            ttk.Label(frm, text="Opacity").grid(row=7, column=0, sticky="w")
            ttk.Entry(frm, textvariable=opacity_var, width=10).grid(row=7, column=1, sticky="w")

            # Autostart on Windows
            autostart_var = tk.BooleanVar(value=self.config.autostart or is_autostart_enabled())
            ttk.Checkbutton(frm, text="Start with Windows", variable=autostart_var).grid(row=8, column=0, columnspan=2, sticky="w", pady=(6,0))

            # Buttons
            btns = ttk.Frame(frm)
            btns.grid(row=9, column=0, columnspan=2, pady=(16,0), sticky="e")

            def on_save():
                self.config.units = units_var.get()
                try:
                    self.config.smoothing_seconds = float(smooth_var.get())
                except Exception:
                    self.config.smoothing_seconds = 0.0
                self.config.theme = theme_var.get()
                self.config.auto_show_overlay = bool(auto_var.get())
                sel = nic_var.get()
                self.config.nic_name = None if sel == "(All)" else sel
                self.config.overlay_transparent = bool(trans_var.get())
                self.config.overlay_compact = bool(compact_var.get())
                try:
                    self.config.overlay_opacity = float(opacity_var.get())
                except Exception:
                    pass
                # Autostart apply
                want_autostart = bool(autostart_var.get())
                self.config.autostart = want_autostart
                try:
                    if want_autostart:
                        enable_autostart_shortcut()
                    else:
                        disable_autostart_shortcut()
                except Exception:
                    pass
                self.config.save()
                # Apply live
                self.monitor.config = self.config
                self.monitor._ema_down_bps = None
                self.monitor._ema_up_bps = None
                self.overlay._apply_theme()
                # Refresh iface selection immediately
                self.monitor._refresh_ifaces()
                # If overlay is visible, reopen to apply size/font adjustments
                if self.overlay.visible:
                    try:
                        self.overlay.close()
                    except Exception:
                        pass
                    self.overlay.open()
                win.destroy()

            ttk.Button(btns, text="Save", command=on_save).pack(side="right", padx=(6,0))
            ttk.Button(btns, text="Cancel", command=win.destroy).pack(side="right")

            # layout stretch
            frm.columnconfigure(1, weight=1)

        self.ui.start()
        self.ui.call(_open_prefs)

    def _on_toggle_overlay(self, icon, item):
        try:
            self.overlay.toggle()
        except Exception:
            pass

    def _on_clear_today(self, icon, item):
        try:
            self.monitor.store.clear_today()
            self.repo.save(self.monitor.store)
        except Exception:
            pass

    def _on_quit(self, icon, item):
        try:
            self.monitor.stop()
        except Exception:
            pass
        try:
            icon.stop()
        except Exception:
            pass
        # Ensure process exits
        os._exit(0)

    def run(self):
        self.monitor.start()

        menu = pystray.Menu(
            pystray.MenuItem("Open Dashboard", self._on_open_dashboard, default=True),
            pystray.MenuItem("Preferences...", self._on_preferences),
            pystray.MenuItem(
                "Show Mini Overlay",
                self._on_toggle_overlay,
                checked=lambda item: self.overlay.visible,
            ),
            pystray.MenuItem("Clear Today's Data", self._on_clear_today),
            pystray.MenuItem("Quit", self._on_quit),
        )
        self.tray_icon = pystray.Icon(
            APP_NAME,
            icon=self.icon_image,
            title="Initializing...",
            menu=menu,
        )

        # Some platforms allow setup callback
        def setup(icon: pystray.Icon):
            # Attach double-click (not universally supported; fallback is default menu item)
            try:
                icon.visible = True
            except Exception:
                pass

        # Auto-show overlay if configured
        if self.config.auto_show_overlay:
            try:
                sw = self.overlay.win.winfo_screenwidth()
                sh = self.overlay.win.winfo_screenheight()
                px, py = self.config.overlay_pos
                if px == 0 and py == 0:
                    x = max(0, sw - self.overlay.width - self.overlay.offset_x)
                    y = max(0, sh - self.overlay.height - self.overlay.offset_y)
                else:
                    x, y = px, py
                x = max(0, min(x, sw - self.overlay.width))
                y = max(0, min(y, sh - self.overlay.height))
                self.overlay.win.geometry(f"{self.overlay.width}x{self.overlay.height}+{x}+{y}")
                self.overlay.open()
            except Exception:
                pass

        self.tray_icon.run(setup=setup)

# -------------------------------
# Entry
# -------------------------------
if __name__ == "__main__":
    # On Windows, avoid tkinter main loop blocking the tray:
    # We keep tray as main loop and open dashboard in a separate thread as needed.
    app = SpeedMeterApp()
    try:
        app.run()
    except KeyboardInterrupt:
        pass
