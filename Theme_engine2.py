#!/usr/bin/env python3
"""
Theme Engine v4.3
Fixes:
  1. wndproc LPARAM overflow on 64-bit Windows (was ctypes.wintypes.LPARAM=32bit)
  2. MPV invalid 'audio' option (replaced with 'ao'='null'/'wasapi')

NEW in v4.3: Config-file system
  • Place  theme_engine.json  next to this script to run fully headless.
  • When a config is found the GUI panel is never opened; all changes are
    applied silently (wallpaper, taskbar colour, overlay, audio visualizer).
  • Run with  --write-config  to generate a fully-annotated template.
  • Run with  --config path/to/file.json  to use a custom config path.

Auto-installs all required dependencies on first run (Python packages + mpv DLLs
+ Visual C++ Redistributable).  A dark-themed progress window is shown during setup.
After installation a marker file '.deps_ok' is written next to the script so the
installer is skipped on every subsequent launch.
"""

import sys, os

# ══════════════════════════════════════════════════════════════════════════════
#  CONFIG SYSTEM  (stdlib only — loaded before any 3rd-party import)
# ══════════════════════════════════════════════════════════════════════════════

# Default configuration — every key is documented here.
_DEFAULT_CONFIG = {
    # ── Headless mode ────────────────────────────────────────────────────────
    # When True (or when this file exists), the GUI panel is never opened.
    # Set to False to still show the panel even when a config file is present.
    "headless": True,

    # ── Video source ─────────────────────────────────────────────────────────
    # Absolute or relative path to the wallpaper video.
    # Relative paths are resolved from the script directory.
    # Leave as "" to skip video (overlay/taskbar colour still works).
    "video_path": "",

    # Seconds to wait before loading the video (useful at system startup
    # so Explorer/Desktop has time to initialise WorkerW).
    "startup_delay": 0,

    # Loop the video indefinitely.
    "loop_video": True,

    # Pass system audio through mpv (heard via speakers).
    # Set False for a silent wallpaper.
    "video_audio": False,

    # ── Adaptive colour ───────────────────────────────────────────────────────
    # Apply the dominant video colour to the Windows taskbar & title bars.
    "adaptive_taskbar_color": True,

    # "dynamic"  — colour updates as the video plays (extracted every ~6 s).
    # "static"   — always use static_color below.
    "palette_mode": "dynamic",

    # Used when palette_mode = "static".  Format: [R, G, B]  0–255.
    "static_color": [70, 130, 220],

    # ── Clock / date overlay ──────────────────────────────────────────────────
    "show_clock": True,
    "show_seconds": True,

    # Font family name (must be installed on the system).
    "clock_font": "Segoe UI Light",

    # Clock text colour + opacity.  Format: [R, G, B]  (opacity set separately).
    "clock_color": [255, 255, 255],

    # 0–255.  220 = slightly transparent white.
    "clock_opacity": 220,

    # Clock position as a fraction of screen size.
    # [0.5, 0.52] = horizontally centred, a little below the vertical midpoint.
    "clock_position": [0.5, 0.52],

    # Sparkle cursor trail.
    "show_sparkle": True,

    # ── Audio visualizer ──────────────────────────────────────────────────────
    "show_visualizer": True,

    # Automatically start WASAPI loopback capture on load.
    "auto_start_audio": True,

    # Visualizer follows the dominant video colour automatically.
    "viz_auto_color": True,

    # Manual bar colour when viz_auto_color = False.  Format: [R, G, B]
    "viz_color": [120, 210, 255],

    # One of: "bars" | "slim" | "mirror" | "wave" | "dots" | "circle"
    "viz_style": "bars",

    # When True the visualizer has its own drag-position independent of the clock.
    "viz_detached": False,

    # Visualizer position (only used when viz_detached = True).
    # Fraction of screen size, same format as clock_position.
    "viz_position": [0.5, 0.80],

    # ── Behaviour ─────────────────────────────────────────────────────────────
    # Pause the wallpaper when a fullscreen app covers the desktop.
    "pause_on_fullscreen": True,
}

_CONFIG_COMMENT = """\
// theme_engine.json — Theme Engine v4.3 configuration
// ─────────────────────────────────────────────────────────────────────────────
// Place this file next to theme_engine.py (or the .exe) to run headlessly.
// All GUI controls map 1-to-1 with the keys below.
// JSON does not support comments — remove these lines before using.
//
// Quick-start:
//   1. Set "video_path" to your wallpaper video.
//   2. Leave everything else at defaults.
//   3. The panel will not open; everything happens silently in the background.
// ─────────────────────────────────────────────────────────────────────────────
"""

def _get_config_path(argv):
    """Return (config_path, write_template_mode)."""
    script_dir = os.path.dirname(os.path.abspath(
        sys.executable if getattr(sys, "frozen", False) else __file__))

    # --write-config  →  write template and exit
    if "--write-config" in argv:
        return os.path.join(script_dir, "theme_engine.json"), True

    # --config path/to/file.json
    if "--config" in argv:
        idx = argv.index("--config")
        if idx + 1 < len(argv):
            return argv[idx + 1], False

    # Default: look next to script / exe
    return os.path.join(script_dir, "theme_engine.json"), False


def _write_config_template(path):
    """Write a human-readable annotated JSON template."""
    import json
    lines = [_CONFIG_COMMENT]
    raw   = json.dumps(_DEFAULT_CONFIG, indent=4)
    lines.append(raw)
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print(f"[Config] Template written to: {path}")
    print("[Config] Edit it, then restart Theme Engine.")


def _load_config(path):
    """
    Load and parse the JSON config.  Returns a dict merged with defaults,
    or None if the file does not exist (→ normal GUI mode).
    """
    if not os.path.exists(path):
        return None
    import json, re
    try:
        with open(path, "r", encoding="utf-8") as f:
            raw = f.read()
        # Strip // … line comments so the user can annotate their config
        raw_clean = re.sub(r'//[^\n]*', '', raw)
        user_cfg  = json.loads(raw_clean)
    except Exception as e:
        # Config exists but is malformed — warn and fall back to GUI
        print(f"[Config] WARNING: could not parse {path}: {e}")
        print("[Config] Falling back to GUI mode.")
        return None

    # Merge: start with defaults, overlay user values
    cfg = dict(_DEFAULT_CONFIG)
    cfg.update(user_cfg)

    # Resolve relative video_path
    if cfg["video_path"]:
        script_dir = os.path.dirname(os.path.abspath(
            sys.executable if getattr(sys, "frozen", False) else __file__))
        if not os.path.isabs(cfg["video_path"]):
            cfg["video_path"] = os.path.join(script_dir, cfg["video_path"])

    return cfg


# Resolve config early so the rest of the module can branch on _CFG
_CONFIG_PATH, _WRITE_CONFIG_MODE = _get_config_path(sys.argv)
_CFG = None   # populated after _bootstrap() if file exists

# ══════════════════════════════════════════════════════════════════════════════
#  FIRST-RUN AUTO-INSTALLER  (stdlib only — runs before any 3rd-party import)
# ══════════════════════════════════════════════════════════════════════════════

def _bootstrap():
    """
    First-run installer. Runs every launch but completes in <10 ms once set up.
    """
    from pathlib import Path
    import ctypes, os as _os, sys as _sys

    if getattr(_sys, "frozen", False):
        return

    script_dir = Path(__file__).parent
    marker     = script_dir / ".deps_ok"

    _MPV_DLL_NAMES = ("libmpv-2.dll", "mpv-2.dll", "mpv-1.dll")

    def _find_mpv_dll():
        for name in _MPV_DLL_NAMES:
            p = script_dir / name
            if p.exists():
                return p
        return None

    def _mpv_loads():
        dll = _find_mpv_dll()
        if not dll:
            return False, f"DLL not found. Expected one of {_MPV_DLL_NAMES} next to the script."
        _os.environ["PATH"] = str(script_dir) + _os.pathsep + _os.environ.get("PATH", "")
        try:
            ctypes.CDLL(str(dll))
            return True, dll.name
        except OSError as e:
            return False, f"{dll.name} failed to load: {e}"

    dll_ok, dll_msg = _mpv_loads()
    if not dll_ok:
        # In headless mode print to stderr and exit without a dialog
        if _load_config(_CONFIG_PATH) is not None:
            print(f"[Setup] FATAL: {dll_msg}", file=sys.stderr)
            _sys.exit(1)
        import tkinter as tk
        from tkinter import messagebox
        _r = tk.Tk(); _r.withdraw()
        messagebox.showerror(
            "Theme Engine — Missing File",
            f"libmpv-2.dll could not be loaded:\n  {dll_msg}\n\n"
            f"Make sure  libmpv-2.dll  is in the same folder as the script:\n"
            f"  {script_dir}\n\n"
            "This file should have been included with the download.\n"
            "If it is missing, contact the developer."
        )
        _r.destroy()
        _sys.exit(1)

    print(f"[Setup] {dll_msg} — OK")

    if marker.exists():
        return

    # In headless / write-config mode skip the interactive installer UI
    headless_install = (_WRITE_CONFIG_MODE or
                        os.path.exists(_CONFIG_PATH))

    import subprocess, threading
    if not headless_install:
        import tkinter as tk
        from tkinter import ttk, messagebox

    PIP_PACKAGES = [
        ("PyQt5",            "PyQt5  (GUI framework)"),
        ("numpy",            "numpy  (array math)"),
        ("opencv-python",    "opencv-python  (video frames / colour extraction)"),
        ("python-mpv",       "python-mpv  (Python binding for libmpv-2.dll)"),
        ("pyaudiowpatch",    "pyaudiowpatch  (WASAPI loopback — no permissions needed)"),
        ("sounddevice",      "sounddevice  (audio fallback)"),
    ]

    if headless_install:
        # Silent pip install — no GUI
        print("[Setup] First run: installing packages silently…")
        _sys.stdout.flush()
        subprocess.run([_sys.executable, "-m", "pip", "install", "--upgrade", "pip", "-q"],
                       capture_output=True)
        for pkg, label in PIP_PACKAGES:
            print(f"[Setup] Installing {label}…")
            _sys.stdout.flush()
            r = subprocess.run([_sys.executable, "-m", "pip", "install", pkg, "-q"],
                               capture_output=True, text=True)
            if r.returncode != 0:
                print(f"[Setup] FAILED: pip install {pkg}\n{r.stderr[-400:]}", file=_sys.stderr)
                _sys.exit(1)
        marker.write_text("ok")
        print("[Setup] All packages installed. Re-launching…")
        import subprocess as _sp
        _sp.Popen([_sys.executable] + _sys.argv)
        _sys.exit(0)

    # ── Interactive progress window (GUI mode only) ───────────────────────────
    _BG, _FG, _ACC, _SUB = "#12121f", "#dde1f0", "#7c6fff", "#444466"
    _OK, _WAIT, _ERR_COL = "#44dd88", "#555577", "#ff6655"

    root = tk.Tk()
    root.title("Theme Engine — First-Run Setup")
    root.geometry("520x310")
    root.resizable(False, False)
    root.configure(bg=_BG)
    root.protocol("WM_DELETE_WINDOW", lambda: None)

    root.update_idletasks()
    sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
    root.geometry(f"520x310+{(sw-520)//2}+{(sh-310)//2}")

    tk.Label(root, text="Theme Engine — First-Run Setup",
             font=("Segoe UI", 12, "bold"), bg=_BG, fg=_ACC).pack(pady=(18, 2))
    tk.Label(root,
             text="Installing Python packages. This only happens once.",
             font=("Segoe UI", 9), bg=_BG, fg=_SUB).pack(pady=(0, 8))

    _step_vars, _step_lbls = [], []
    chk = tk.Frame(root, bg=_BG); chk.pack(fill="x", padx=32, pady=(0, 6))
    for _, label in PIP_PACKAGES:
        row = tk.Frame(chk, bg=_BG); row.pack(anchor="w", pady=1)
        sv = tk.StringVar(value="○"); _step_vars.append(sv)
        il = tk.Label(row, textvariable=sv, font=("Segoe UI", 10), width=2, bg=_BG, fg=_WAIT)
        il.pack(side="left")
        tl = tk.Label(row, text=label, font=("Segoe UI", 8), bg=_BG, fg=_SUB, anchor="w")
        tl.pack(side="left")
        _step_lbls.append((il, tl))

    def _tick(i, ok=True):
        def _do():
            _step_vars[i].set("✓" if ok else "✗")
            c = _OK if ok else _ERR_COL
            _step_lbls[i][0].config(fg=c)
            _step_lbls[i][1].config(fg=_FG if ok else _ERR_COL)
        root.after(0, _do)

    def _run(i):
        def _do():
            _step_vars[i].set("▶")
            _step_lbls[i][0].config(fg=_ACC)
            _step_lbls[i][1].config(fg=_FG)
        root.after(0, _do)

    _sty = ttk.Style(); _sty.theme_use("default")
    _sty.configure("WE.Horizontal.TProgressbar",
                   troughcolor="#22223a", background=_ACC, thickness=10, borderwidth=0)
    bar = ttk.Progressbar(root, length=460, mode="determinate",
                          style="WE.Horizontal.TProgressbar")
    bar.pack(pady=(4, 2))

    status_var = tk.StringVar(value="Starting…")
    tk.Label(root, textvariable=status_var, font=("Segoe UI", 8, "bold"), bg=_BG, fg=_FG).pack()
    detail_var = tk.StringVar(value="")
    tk.Label(root, textvariable=detail_var, font=("Consolas", 7), bg=_BG, fg="#333355").pack()

    err_holder = [None]

    def _ui(msg, pct, detail=""):
        def _do():
            status_var.set(msg); bar["value"] = pct
            if detail: detail_var.set(detail[:100])
            root.update_idletasks()
        root.after(0, _do)

    def _worker():
        try:
            _ui("Upgrading pip…", 2)
            subprocess.run([_sys.executable, "-m", "pip", "install", "--upgrade", "pip", "-q"],
                           capture_output=True)
            n = len(PIP_PACKAGES)
            for i, (pkg, _) in enumerate(PIP_PACKAGES):
                _run(i); pct = 5 + int((i / n) * 90)
                _ui(f"Installing {pkg}…", pct, detail=f"pip install {pkg}")
                r = subprocess.run([_sys.executable, "-m", "pip", "install", pkg, "-q"],
                                   capture_output=True, text=True)
                if r.returncode != 0:
                    _tick(i, ok=False)
                    raise RuntimeError(f"pip install {pkg} failed:\n{r.stderr[-400:]}")
                _tick(i, ok=True)
            marker.write_text("ok")
            _ui("All done! Starting Theme Engine…", 100)
        except Exception as exc:
            err_holder[0] = exc
        finally:
            root.after(1200, root.destroy)

    threading.Thread(target=_worker, daemon=True).start()
    root.mainloop()

    if err_holder[0]:
        try:
            messagebox.showerror(
                "Setup Failed",
                f"Could not install a required package:\n\n{err_holder[0]}\n\n"
                "Try running manually:\n"
                "  pip install PyQt5 numpy opencv-python python-mpv sounddevice"
            )
        except Exception:
            print(f"[Setup] FAILED: {err_holder[0]}")
        _sys.exit(1)

    print("[Setup] Re-launching with installed packages…")
    import subprocess as _sp
    _sp.Popen([_sys.executable] + _sys.argv)
    _sys.exit(0)


_bootstrap()   # Must be called before any third-party import

# Handle --write-config here, after bootstrap ensured stdlib is ok
if _WRITE_CONFIG_MODE:
    _write_config_template(_CONFIG_PATH)
    sys.exit(0)

# Load config (None = no file found = normal GUI mode)
_CFG = _load_config(_CONFIG_PATH)
_HEADLESS = _CFG is not None and _CFG.get("headless", True)

if _CFG:
    print(f"[Config] Loaded: {_CONFIG_PATH}")
    print(f"[Config] Headless mode: {_HEADLESS}")

# ══════════════════════════════════════════════════════════════════════════════
#  END OF AUTO-INSTALLER — normal application code begins below
# ══════════════════════════════════════════════════════════════════════════════

import sys, os, ctypes, ctypes.wintypes, threading, time, winreg, tempfile
from pathlib import Path

if getattr(sys, "frozen", False):
    _exe_dir    = os.path.dirname(sys.executable)
    _meipass    = sys._MEIPASS
    _script_dir = _exe_dir
    for _d in (_exe_dir, _meipass):
        if _d not in os.environ["PATH"]:
            os.environ["PATH"] = _d + os.pathsep + os.environ["PATH"]
    import traceback as _tb
    _crash_log = os.path.join(_exe_dir, "ThemeEngine_crash.log")
    def _excepthook(exc_type, exc_value, exc_tb):
        msg = "".join(_tb.format_exception(exc_type, exc_value, exc_tb))
        try:
            with open(_crash_log, "w") as _f:
                _f.write(msg)
        except Exception:
            pass
        from PyQt5.QtWidgets import QApplication, QMessageBox
        _app = QApplication.instance() or QApplication(sys.argv)
        QMessageBox.critical(None, "Theme Engine — Crash",
            f"Theme Engine encountered an error:\n\n{exc_value}\n\n"
            f"Full details saved to:\n{_crash_log}")
        sys.exit(1)
    sys.excepthook = _excepthook
else:
    _script_dir = os.path.dirname(os.path.abspath(__file__))

if _script_dir not in os.environ["PATH"]:
    os.environ["PATH"] = _script_dir + os.pathsep + os.environ["PATH"]

_common_mpv_paths = [
    r"C:\Program Files\mpv", r"C:\Program Files (x86)\mpv", r"C:\mpv",
    r"C:\ProgramData\chocolatey\lib\mpvio.install\tools",
    r"C:\ProgramData\chocolatey\bin",
    os.path.expanduser(r"~\AppData\Local\Programs\mpv"),
    os.path.expanduser(r"~\scoop\apps\mpv\current"),
    os.path.expanduser(r"~\scoop\shims"),
]
for _p in _common_mpv_paths:
    if os.path.exists(_p) and _p not in os.environ["PATH"]:
        os.environ["PATH"] = _p + os.pathsep + os.environ["PATH"]

import cv2
import numpy as np

AUDIO_BACKEND = None
sd = None

try:
    import pyaudiowpatch as _pawp
    _p = _pawp.PyAudio()
    try:    _p.get_host_api_info_by_type(_pawp.paWASAPI)
    finally: _p.terminate()
    AUDIO_BACKEND = "wpatch"
    print("[Audio] pyaudiowpatch WASAPI loopback available")
except Exception as _e:
    print(f"[Audio] pyaudiowpatch not available ({_e}), trying sounddevice")
    try:
        import sounddevice as sd
        AUDIO_BACKEND = "sounddevice"
        print("[Audio] sounddevice fallback active")
    except ImportError:
        print("[Audio] No audio backend available — visualizer disabled")

AUDIO_OK = AUDIO_BACKEND is not None

MPV_OK = False
mpv = None

try:
    import mpv
    MPV_OK = True
    print("[mpv] Loaded successfully")
except Exception as e:
    print(f"[mpv] Failed to load: {e}")

from PyQt5.QtWidgets import (
    QFontComboBox,
    QApplication, QWidget, QLabel, QPushButton, QVBoxLayout, QHBoxLayout,
    QFileDialog, QCheckBox, QFontComboBox, QSlider, QColorDialog,
    QScrollArea, QSystemTrayIcon, QMenu, QSizePolicy, QMessageBox
)
from PyQt5.QtCore import Qt, QTimer, pyqtSignal, pyqtSlot, QRect, QThread, QObject, QDateTime, QPointF
from PyQt5.QtGui import (
    QImage, QPixmap, QPainter, QColor, QFont, QPainterPath,
    QBrush, QCursor, QPen, QLinearGradient, QIcon
)

_WND_CLASS_REGISTERED = False
_WNDPROC_CB = None

def set_dpi_aware():
    try: ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except: pass

def get_screen_size():
    u = ctypes.windll.user32
    return u.GetSystemMetrics(0), u.GetSystemMetrics(1)

def get_foreground_class():
    hwnd = ctypes.windll.user32.GetForegroundWindow()
    buf  = ctypes.create_unicode_buffer(256)
    ctypes.windll.user32.GetClassNameW(hwnd, buf, 256)
    return buf.value

def is_admin():
    try: return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except: return False

def make_tray_icon_pixmap(r, g, b):
    px = QPixmap(32, 32); px.fill(Qt.transparent)
    p = QPainter(px); p.setRenderHint(QPainter.Antialiasing)
    p.setBrush(QBrush(QColor(r,g,b,230))); p.setPen(Qt.NoPen)
    p.drawEllipse(1,1,30,30)
    p.setBrush(QBrush(QColor(255,255,255,240)))
    tri = QPainterPath()
    tri.moveTo(11,8); tri.lineTo(11,24); tri.lineTo(24,16); tri.closeSubpath()
    p.drawPath(tri); p.end()
    return px

# ══════════════════════════════════════════════
#  COLOR EXTRACTOR
# ══════════════════════════════════════════════

class ColorExtractor:
    def __init__(self):
        self._colors = [(30,80,160)]*6
        self._lock   = threading.Lock()
        self._last   = 0.0

    def extract(self, frame):
        if time.time() - self._last < 6.0:
            return self.get_colors()
        self._last = time.time()
        try:
            small = cv2.resize(frame, (64, 64))
            rgb   = cv2.cvtColor(small, cv2.COLOR_BGR2RGB)
            px    = np.ascontiguousarray(rgb.reshape(-1, 3), dtype=np.float32)
            crit = (cv2.TERM_CRITERIA_EPS | cv2.TERM_CRITERIA_MAX_ITER, 12, 1.0)
            _, labels, centers = cv2.kmeans(px, 6, None, crit, 4, cv2.KMEANS_PP_CENTERS)
            counts = np.bincount(labels.flatten(), minlength=6)
            scores = []
            for i, (rc, gc, bc) in enumerate(centers):
                mx, mn = max(rc,gc,bc)/255.0, min(rc,gc,bc)/255.0
                lum    = (mx + mn) / 2.0
                sat    = (mx - mn) / (1.0 - abs(2*lum - 1) + 1e-9) if mx > 0 else 0.0
                cov    = counts[i] / float(len(labels))
                lum_ok = 1.0 if 0.05 < lum < 0.92 else 0.1
                scores.append(sat * (cov ** 0.5) * lum_ok)
            order  = np.argsort(-np.array(scores))
            colors = [(int(centers[i][0]), int(centers[i][1]), int(centers[i][2]))
                      for i in order]
            with self._lock: self._colors = colors
            return colors
        except Exception:
            return self.get_colors()

    def get_colors(self):
        with self._lock: return list(self._colors)

    def get_accent(self):
        best, bs = self.get_colors()[0], -1.0
        for r, g, b in self.get_colors():
            mx, mn = max(r,g,b)/255.0, min(r,g,b)/255.0
            lum = (mx + mn) / 2.0
            sat = (mx - mn) / (1.0 - abs(2*lum - 1) + 1e-9) if mx > mn else 0.0
            if 0.1 < lum < 0.9 and sat > 0.10:
                sc = sat * sat * (1.0 - abs(lum - 0.50))
                if sc > bs: bs, best = sc, (r, g, b)
        r, g, b = best
        hi, lo = max(r,g,b), min(r,g,b)
        if hi > lo:
            mid   = (hi + lo) / 2
            scale = min(255.0 / max(hi, 1), 1.5)
            r = int(max(0, min(255, mid + (r - mid) * scale)))
            g = int(max(0, min(255, mid + (g - mid) * scale)))
            b = int(max(0, min(255, mid + (b - mid) * scale)))
        lum_out = (max(r,g,b) + min(r,g,b)) / 2
        if lum_out < 55:
            factor = 55 / max(lum_out, 1)
            r = min(255, int(r * factor))
            g = min(255, int(g * factor))
            b = min(255, int(b * factor))
        return (r, g, b)

    def get_glow(self):
        r, g, b = self.get_accent()
        return (min(255,r+80), min(255,g+80), min(255,b+80))

# ══════════════════════════════════════════════
#  WINDOWS THEMER
# ══════════════════════════════════════════════

class WindowsThemer:
    @staticmethod
    def _make_palette(r, g, b):
        levels = [0.10, 0.20, 0.35, 0.55, 0.70, 0.80, 0.90, 0.96]
        data = bytearray()
        for lv in levels:
            if lv <= 0.55:
                t  = lv / 0.55
                pr = int(r * t); pg = int(g * t); pb = int(b * t)
            else:
                t  = (lv - 0.55) / 0.45
                pr = int(r + (255 - r) * t); pg = int(g + (255 - g) * t); pb = int(b + (255 - b) * t)
            data += bytes([max(0,min(255,pb)), max(0,min(255,pg)), max(0,min(255,pr)), 0xFF])
        return bytes(data)

    def __init__(self):
        self._last_rgb = (-1, -1, -1)

    def apply(self, r, g, b):
        lr, lg, lb = self._last_rgb
        if abs(r-lr) + abs(g-lg) + abs(b-lb) < 6:
            return
        self._last_rgb = (r, g, b)
        reg_abgr = ctypes.c_uint32(0xFF000000 | (b<<16) | (g<<8) | r).value
        palette  = self._make_palette(r, g, b)
        u32      = ctypes.windll.user32
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Explorer\Accent",
                0, winreg.KEY_WRITE) as k:
                winreg.SetValueEx(k, "AccentColor",      0, winreg.REG_DWORD,  reg_abgr)
                winreg.SetValueEx(k, "AccentColorMenu",  0, winreg.REG_DWORD,  reg_abgr)
                winreg.SetValueEx(k, "StartColorMenu",   0, winreg.REG_DWORD,  reg_abgr)
                winreg.SetValueEx(k, "AccentPalette",    0, winreg.REG_BINARY, palette)
        except Exception as e:
            print(f"[Theme] Accent registry: {e}")
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize",
                0, winreg.KEY_WRITE) as k:
                winreg.SetValueEx(k, "ColorPrevalence",        0, winreg.REG_DWORD, 1)
                winreg.SetValueEx(k, "TaskbarColorPrevalence", 0, winreg.REG_DWORD, 1)
        except Exception as e:
            print(f"[Theme] Personalize: {e}")
        try:
            dwmapi = ctypes.windll.dwmapi
            dwmapi.DwmSetColorizationColor.restype  = ctypes.c_long
            dwmapi.DwmSetColorizationColor.argtypes = [ctypes.c_uint32, ctypes.c_bool]
            dwmapi.DwmSetColorizationColor(reg_abgr, False)
        except Exception as e:
            print(f"[Theme] DWM: {e}")
        try:
            WM_SETTINGCHANGE = 0x001A
            WM_THEMECHANGED  = 0x031A
            buf = ctypes.create_unicode_buffer("ImmersiveColorSet")
            res = ctypes.wintypes.DWORD(0)
            tray = u32.FindWindowW("Shell_TrayWnd", None)
            if tray:
                u32.SendMessageTimeoutW(tray, WM_SETTINGCHANGE, 0, buf, 0x0002, 500, ctypes.byref(res))
                u32.SendMessageTimeoutW(tray, WM_THEMECHANGED,  0, 0,   0x0002, 500, None)
            secondary = u32.FindWindowW("Shell_SecondaryTrayWnd", None)
            while secondary:
                u32.SendMessageTimeoutW(secondary, WM_SETTINGCHANGE, 0, buf, 0x0002, 500, ctypes.byref(res))
                u32.SendMessageTimeoutW(secondary, WM_THEMECHANGED,  0, 0,   0x0002, 500, None)
                secondary = u32.FindWindowExW(None, secondary, "Shell_SecondaryTrayWnd", None)
        except Exception as e:
            print(f"[Theme] Taskbar poke: {e}")

    def restore(self): self.apply(0, 120, 215)

# ══════════════════════════════════════════════
#  DESKTOP EMBEDDER
# ══════════════════════════════════════════════

class DesktopEmbedder:
    def __init__(self):
        self._target = 0

    def _cls(self, hwnd):
        buf = ctypes.create_unicode_buffer(64)
        ctypes.windll.user32.GetClassNameW(hwnd, buf, 64)
        return buf.value

    def _send_spawn_msg(self, pm):
        u32 = ctypes.windll.user32
        res = ctypes.wintypes.DWORD(0)
        u32.SendMessageTimeoutW(pm, 0x052C, 0xD, 0x1, 0x0002, 2000, ctypes.byref(res))
        time.sleep(0.25)
        u32.SendMessageTimeoutW(pm, 0x052C, 0,   0,   0x0002, 2000, ctypes.byref(res))
        time.sleep(0.25)

    def _find_target(self):
        u32 = ctypes.windll.user32
        pm  = u32.FindWindowW("Progman", None)
        if not pm:
            print("[Embed] Progman not found"); return 0
        self._send_spawn_msg(pm)
        PROC = ctypes.WINFUNCTYPE(ctypes.wintypes.BOOL, ctypes.wintypes.HWND, ctypes.wintypes.LPARAM)
        wallpaper_ww = [0]
        def cb(hwnd, _):
            defview = u32.FindWindowExW(hwnd, None, "SHELLDLL_DefView", None)
            if defview:
                cls = self._cls(hwnd)
                if cls == "WorkerW":
                    sib = u32.FindWindowExW(None, hwnd, "WorkerW", None)
                    wallpaper_ww[0] = sib if sib else hwnd
                elif cls == "Progman":
                    wallpaper_ww[0] = hwnd
            return True
        u32.EnumWindows(PROC(cb), 0)
        if wallpaper_ww[0]:
            return wallpaper_ww[0]
        return pm

    def _prep_for_child(self, hwnd):
        u32 = ctypes.windll.user32
        u32.GetWindowLongPtrW.restype  = ctypes.c_ssize_t
        u32.GetWindowLongPtrW.argtypes = [ctypes.wintypes.HWND, ctypes.c_int]
        u32.SetWindowLongPtrW.restype  = ctypes.c_ssize_t
        u32.SetWindowLongPtrW.argtypes = [ctypes.wintypes.HWND, ctypes.c_int, ctypes.c_ssize_t]
        GWL_STYLE, GWL_EXSTYLE = -16, -20
        WS_POPUP=0x80000000; WS_CHILD=0x40000000; WS_CAPTION=0x00C00000; WS_THICKFRAME=0x00040000
        style   = u32.GetWindowLongPtrW(hwnd, GWL_STYLE)
        exstyle = u32.GetWindowLongPtrW(hwnd, GWL_EXSTYLE)
        new_style   = ctypes.c_ssize_t((style & ~(WS_POPUP | WS_CAPTION | WS_THICKFRAME)) | WS_CHILD).value
        new_exstyle = ctypes.c_ssize_t(exstyle & ~0x00040000).value
        u32.SetWindowLongPtrW(hwnd, GWL_STYLE,   new_style)
        u32.SetWindowLongPtrW(hwnd, GWL_EXSTYLE, new_exstyle)
        u32.SetWindowPos(hwnd, 0, 0, 0, 0, 0, 0x0001|0x0002|0x0004|0x0020)

    def _force_show(self, hwnd, sw, sh):
        u32 = ctypes.windll.user32
        parent = u32.GetParent(hwnd)
        defview = u32.FindWindowExW(parent, None, "SHELLDLL_DefView", None) if parent else None
        insert_after = defview if defview else 1
        u32.SetWindowPos(hwnd, insert_after, 0, 0, sw, sh, 0x0010|0x0040|0x0020)
        u32.ShowWindow(hwnd, 5)
        u32.UpdateWindow(hwnd)

    def embed(self, hwnd, sw, sh):
        u32 = ctypes.windll.user32
        self._prep_for_child(hwnd)
        for attempt in range(5):
            target = self._find_target()
            if target:
                result = u32.SetParent(hwnd, target)
                err    = ctypes.GetLastError()
                self._force_show(hwnd, sw, sh)
                actual = u32.GetParent(hwnd)
                anc    = u32.GetAncestor(hwnd, 1)
                print(f"[Embed] attempt {attempt+1}: result={result:#x} err={err} target={target:#010x} parent={actual:#010x} ancestor={anc:#010x}")
                if actual == target or anc == target:
                    self._target = target
                    print("[Embed] SUCCESS"); return True
            time.sleep(0.4)
        print("[Embed] FAILED"); return False

    def detach(self, hwnd):
        ctypes.windll.user32.SetParent(hwnd, 0)
        ctypes.windll.user32.ShowWindow(hwnd, 0)

    def get_target(self): return self._target

# ══════════════════════════════════════════════
#  AUDIO VISUALIZER
# ══════════════════════════════════════════════

class AudioVisualizer(QObject):
    bars_updated = pyqtSignal(list)
    N = 48

    def __init__(self):
        super().__init__()
        self._running  = False
        self._smoothed = [0.0] * self.N
        self._peak     = [1e-4] * self.N
        self._lock     = threading.Lock()
        self._queue    = None
        self._sr       = 48000

    def start(self):
        if not AUDIO_OK or self._running: return
        import queue
        self._queue   = queue.SimpleQueue()
        self._running = True
        threading.Thread(target=self._run,     daemon=True).start()
        threading.Thread(target=self._emitter, daemon=True).start()

    def stop(self):
        self._running = False

    def _emitter(self):
        import time as _time
        SILENCE_TIMEOUT = 0.30
        ZERO_DECAY      = 0.72
        last_data       = _time.monotonic()
        while self._running:
            try:
                bars = self._queue.get(timeout=0.033)
                last_data = _time.monotonic()
                self.bars_updated.emit(bars)
            except Exception:
                elapsed = _time.monotonic() - last_data
                if elapsed > SILENCE_TIMEOUT:
                    with self._lock:
                        self._smoothed = [v * ZERO_DECAY for v in self._smoothed]
                        zeroed = list(self._smoothed)
                    if any(v > 0.002 for v in zeroed):
                        self.bars_updated.emit(zeroed)
                    else:
                        self.bars_updated.emit([0.0] * self.N)

    def _fft_to_bars(self, audio_mono):
        FFT_N=4096; SR=self._sr; F_LO=40.0; F_HI=18000.0
        DECAY=0.88; ATTACK=0.55; EQ_RISE=0.015; EQ_FALL=0.002
        windowed = audio_mono * np.hanning(len(audio_mono))
        fft_full = np.abs(np.fft.rfft(windowed, n=FFT_N))
        n_bins   = len(fft_full)
        bin_lo   = max(1, int(F_LO  * FFT_N / SR))
        bin_hi   = min(n_bins - 1, int(F_HI * FFT_N / SR))
        usable   = fft_full[bin_lo:bin_hi]
        n_use    = len(usable)
        if n_use < self.N: return None
        log_lo   = np.log1p(0); log_hi = np.log1p(n_use)
        edges    = [int(np.expm1(log_lo + (log_hi - log_lo) * i / self.N)) for i in range(self.N + 1)]
        raw_bars = []
        for i in range(self.N):
            lo = edges[i]; hi = max(lo + 1, edges[i + 1]); hi = min(hi, n_use)
            raw_bars.append(float(np.mean(usable[lo:hi])))
        with self._lock:
            for i, v in enumerate(raw_bars):
                if v > self._peak[i]:
                    self._peak[i] = self._peak[i] * (1 - EQ_RISE) + v * EQ_RISE
                else:
                    self._peak[i] = max(1e-6, self._peak[i] * (1 - EQ_FALL))
            eq_bars = [v / self._peak[i] for i, v in enumerate(raw_bars)]
            mx = max(eq_bars) if max(eq_bars) > 1e-8 else 1e-8
            eq_bars = [b / mx for b in eq_bars]
            for i, (nw, old) in enumerate(zip(eq_bars, self._smoothed)):
                self._smoothed[i] = (nw * ATTACK + old * (1 - ATTACK) if nw > old else old * DECAY)
            return list(self._smoothed)

    def _process(self, raw_bytes, n_channels, sampwidth=4):
        try:
            arr = np.frombuffer(raw_bytes, dtype=np.float32)
            if arr.size == 0: return None
            if n_channels > 1: arr = arr.reshape(-1, n_channels).mean(axis=1)
            return self._fft_to_bars(arr)
        except Exception as e:
            print(f"[Audio] _process: {e}")
            return None

    def _push(self, bars):
        if self._queue is None: return
        try:    self._queue.get_nowait()
        except Exception: pass
        self._queue.put(bars)

    def _run_wpatch(self):
        import pyaudiowpatch as pyaudio
        p = pyaudio.PyAudio()
        try:
            wasapi = p.get_host_api_info_by_type(pyaudio.paWASAPI)
            default_out = p.get_device_info_by_index(wasapi["defaultOutputDevice"])
            loopback_dev = None
            for dev in p.get_loopback_device_info_generator():
                if default_out["name"] in dev["name"]: loopback_dev = dev; break
            if loopback_dev is None:
                for dev in p.get_loopback_device_info_generator(): loopback_dev = dev; break
            if loopback_dev is None:
                print("[Audio] pyaudiowpatch: no loopback device found"); return False
            ch = int(loopback_dev["maxInputChannels"]); sr = int(loopback_dev["defaultSampleRate"])
            self._sr = sr; chunk = 1024
            stream = p.open(format=pyaudio.paFloat32, channels=ch, rate=sr, input=True,
                            input_device_index=loopback_dev["index"], frames_per_buffer=chunk)
            while self._running:
                try:
                    data = stream.read(chunk, exception_on_overflow=False)
                    bars = self._process(data, ch)
                    if bars: self._push(bars)
                except Exception as e:
                    print(f"[Audio] wpatch read: {e}"); break
            stream.stop_stream(); stream.close(); return True
        except Exception as e:
            print(f"[Audio] wpatch error: {e}"); return False
        finally:
            p.terminate()

    def _find_loopback_sd(self):
        try:
            devices  = sd.query_devices(); hostapis = sd.query_hostapis()
            wasapi_idx = next((i for i, a in enumerate(hostapis) if 'wasapi' in a['name'].lower()), None)
            if wasapi_idx is not None:
                out_idx = hostapis[wasapi_idx].get('default_output_device', -1)
                if out_idx < 0:
                    out_idx = next((i for i, d in enumerate(devices)
                                    if d['hostapi'] == wasapi_idx and d['max_output_channels'] > 0), -1)
                if out_idx >= 0:
                    d  = devices[out_idx]; ch = min(max(d.get('max_output_channels', 2), 1), 2)
                    sr = int(d.get('default_samplerate', 48000))
                    return out_idx, ch, True
            for i, d in enumerate(devices):
                nm = d['name'].lower()
                if d['max_input_channels'] > 0 and any(k in nm for k in ('stereo mix','loopback','what u hear','wave out mix')):
                    return i, min(d['max_input_channels'], 2), False
        except Exception as e:
            print(f"[Audio] sd device query: {e}")
        return None, 1, False

    def _run_sounddevice(self):
        dev, nch, wasapi_lb = self._find_loopback_sd()
        if dev is None:
            print("[Audio] sounddevice: no capture device found"); return False
        try:
            sr = int(sd.query_devices(dev)['default_samplerate'])
        except Exception:
            sr = 48000
        self._sr = sr
        def cb(indata, frames, t, status):
            if not self._running: return
            try:
                audio = indata.mean(axis=1) if indata.ndim > 1 else indata.flatten()
                bars  = self._fft_to_bars(audio)
                if bars: self._push(bars)
            except Exception as e:
                print(f"[Audio] sd cb: {e}")
        kw = dict(device=dev, channels=nch, samplerate=sr, blocksize=2048, callback=cb)
        if wasapi_lb:
            try: kw['extra_settings'] = sd.WasapiSettings(loopback=True)
            except Exception: pass
        try:
            with sd.InputStream(**kw):
                while self._running: time.sleep(0.05)
            return True
        except Exception as e:
            print(f"[Audio] sounddevice stream error: {e}"); return False

    def _run(self):
        if AUDIO_BACKEND == "wpatch":
            ok = self._run_wpatch()
            if not ok and sd is not None: self._run_sounddevice()
        elif AUDIO_BACKEND == "sounddevice":
            self._run_sounddevice()
        while self._running: time.sleep(0.1)

    def can_capture(self):
        if AUDIO_BACKEND == "wpatch":
            import pyaudiowpatch as pyaudio
            p = pyaudio.PyAudio()
            try:
                p.get_host_api_info_by_type(pyaudio.paWASAPI)
                for _ in p.get_loopback_device_info_generator(): return True
                return False
            except Exception: return False
            finally: p.terminate()
        elif AUDIO_BACKEND == "sounddevice":
            dev, _, _ = self._find_loopback_sd(); return dev is not None
        return False

# ══════════════════════════════════════════════
#  OVERLAY
# ══════════════════════════════════════════════

class Overlay(QWidget):
    VIZ_STYLES = ["bars", "slim", "mirror", "wave", "dots", "circle"]

    def __init__(self, sw, sh):
        super().__init__()
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.Tool | Qt.WindowStaysOnBottomHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setAttribute(Qt.WA_ShowWithoutActivating)
        self.setGeometry(0, 0, sw, sh)
        self._sw, self._sh = sw, sh
        self.font_family = "Segoe UI Light"
        self.text_color  = QColor(255, 255, 255, 220)
        self.show_secs   = True
        self.show_clock  = True
        self.cx          = sw // 2
        self.cy          = int(sh * 0.52)
        self._bars       = [0.0] * 48
        self._bar_col    = QColor(120, 210, 255, 210)
        self.show_viz    = True
        self.viz_style   = "bars"
        self.viz_detached = False
        self.vx          = sw // 2
        self.vy          = int(sh * 0.80)
        self._drag        = None
        self._drag_origin = None
        self._sparkles         = []
        self._last_mouse       = None
        self._sparkles_enabled = True
        QTimer(self, timeout=self.update,          interval=33).start()
        QTimer(self, timeout=self._poll_cursor,    interval=33).start()
        QTimer(self, timeout=self._age_sparkles,   interval=33).start()

    def apply_config(self, cfg):
        """Apply all overlay-related keys from a config dict."""
        sw, sh = self._sw, self._sh

        self.show_clock  = cfg.get("show_clock", True)
        self.show_secs   = cfg.get("show_seconds", True)
        self.font_family = cfg.get("clock_font", "Segoe UI Light")

        col = cfg.get("clock_color", [255, 255, 255])
        op  = cfg.get("clock_opacity", 220)
        self.text_color = QColor(col[0], col[1], col[2], op)

        pos = cfg.get("clock_position", [0.5, 0.52])
        self.cx = int(sw * pos[0])
        self.cy = int(sh * pos[1])

        self._sparkles_enabled = cfg.get("show_sparkle", True)

        self.show_viz      = cfg.get("show_visualizer", True)
        self.viz_detached  = cfg.get("viz_detached", False)
        style = cfg.get("viz_style", "bars")
        self.viz_style = style if style in self.VIZ_STYLES else "bars"

        vcol = cfg.get("viz_color", [120, 210, 255])
        self._bar_col = QColor(vcol[0], vcol[1], vcol[2], 210)

        vpos = cfg.get("viz_position", [0.5, 0.80])
        self.vx = int(sw * vpos[0])
        self.vy = int(sh * vpos[1])

    def set_bars(self, b):         self._bars = b;         self.update()
    def set_bar_col(self, r,g,b):  self._bar_col = QColor(r,g,b,210); self.update()
    def set_text_col(self, c):     self.text_color = c;   self.update()
    def set_viz_style(self, s):    self.viz_style = s if s in self.VIZ_STYLES else "bars"; self.update()

    def _poll_cursor(self):
        if not getattr(self, '_sparkles_enabled', True): return
        gp = QCursor.pos(); p = self.mapFromGlobal(gp)
        if self._last_mouse and (abs(p.x()-self._last_mouse.x()) > 2 or abs(p.y()-self._last_mouse.y()) > 2):
            self._spawn_sparkles(p.x(), p.y(), count=3)
        self._last_mouse = p

    def _spawn_sparkles(self, x, y, count=6):
        import random, math
        bc = self._bar_col
        for _ in range(count):
            angle = random.uniform(0, math.tau); speed = random.uniform(0.5, 2.8); life = random.randint(18, 42)
            dr = random.randint(-30, 30); dg = random.randint(-30, 30); db = random.randint(-30, 30)
            self._sparkles.append({"x": x, "y": y, "vx": math.cos(angle)*speed,
                "vy": math.sin(angle)*speed - random.uniform(0.3, 1.2), "life": life, "maxlife": life,
                "r": max(0,min(255,bc.red()+dr)), "g": max(0,min(255,bc.green()+dg)), "b": max(0,min(255,bc.blue()+db))})
        if len(self._sparkles) > 300: self._sparkles = self._sparkles[-300:]

    def _age_sparkles(self):
        alive = []
        for s in self._sparkles:
            s["x"] += s["vx"]; s["y"] += s["vy"]; s["vy"] += 0.08; s["life"] -= 1
            if s["life"] > 0: alive.append(s)
        self._sparkles = alive

    def paintEvent(self, _):
        p = QPainter(self); p.setRenderHints(QPainter.Antialiasing | QPainter.TextAntialiasing)
        self._draw_sparkles(p)
        if self.show_clock: self._draw_clock(p)
        if self.show_viz:   self._draw_viz(p)
        p.end()

    def _draw_sparkles(self, p):
        if not self._sparkles: return
        p.setPen(Qt.NoPen)
        for s in self._sparkles:
            t = s["life"] / s["maxlife"]; size = max(1.0, t * 5.0); alpha = int(t * 220)
            c = QColor(s["r"], s["g"], s["b"], alpha); p.setBrush(QBrush(c))
            cx, cy = s["x"], s["y"]; p.drawEllipse(QPointF(cx, cy), size * 0.9, size * 0.9)
            if t > 0.5:
                arm = size * 1.6; pc = QColor(255, 255, 255, int(t * 160)); p.setBrush(QBrush(pc))
                p.drawEllipse(QPointF(cx, cy - arm * 0.5), size * 0.25, size * 0.6)
                p.drawEllipse(QPointF(cx, cy + arm * 0.5), size * 0.25, size * 0.6)
                p.drawEllipse(QPointF(cx - arm * 0.5, cy), size * 0.6, size * 0.25)
                p.drawEllipse(QPointF(cx + arm * 0.5, cy), size * 0.6, size * 0.25)

    def _draw_clock(self, p):
        now  = QDateTime.currentDateTime()
        day  = now.toString("dddd").upper(); date = now.toString("d MMM yyyy")
        tstr = now.toString("HH:mm:ss") if self.show_secs else now.toString("HH:mm")
        cx, cy = self.cx, self.cy
        def draw_text(text, pt, y_off, tracking=6):
            f = QFont(self.font_family, pt, QFont.Light); f.setLetterSpacing(QFont.AbsoluteSpacing, tracking)
            p.setFont(f); rect = QRect(cx - 700, cy + y_off - pt, 1400, pt * 2 + 10)
            p.setPen(QColor(0, 0, 0, 120))
            p.drawText(rect.adjusted(2, 3, 2, 3), Qt.AlignHCenter | Qt.AlignVCenter, text)
            p.setPen(self.text_color)
            p.drawText(rect, Qt.AlignHCenter | Qt.AlignVCenter, text)
        draw_text(day, 58, -55, 18); draw_text(date, 19, 20, 4); draw_text(tstr, 30, 68, 6)

    def _viz_anchor(self):
        if self.viz_detached: return self.vx, self.vy
        return self.cx, self.cy

    def _draw_viz(self, p):
        if not any(b > 0.001 for b in self._bars): return
        style = self.viz_style
        if   style == "bars":   self._viz_bars(p, thin=False)
        elif style == "slim":   self._viz_bars(p, thin=True)
        elif style == "mirror": self._viz_mirror(p)
        elif style == "wave":   self._viz_wave(p)
        elif style == "dots":   self._viz_dots(p)
        elif style == "circle": self._viz_circle(p)

    def _bar_geometry(self, thin=False):
        sw, sh  = self._sw, self._sh; cx, cy = self._viz_anchor(); N = len(self._bars)
        BAR_W   = max(3 if thin else 6, int(sw * (0.004 if thin else 0.008)))
        GAP     = max(2 if thin else 2,  int(sw * (0.004 if thin else 0.003)))
        MAX_H   = int(sh * 0.16)
        y_base  = cy if self.viz_detached else max(MAX_H + 8, cy - 84 - 18)
        total_w = N * (BAR_W + GAP) - GAP; x0 = cx - total_w // 2
        return x0, y_base, BAR_W, GAP, MAX_H

    def _viz_bars(self, p, thin=False):
        x0, y_base, BAR_W, GAP, MAX_H = self._bar_geometry(thin)
        for i, val in enumerate(self._bars):
            h = max(2, int(val * MAX_H)); x = x0 + i * (BAR_W + GAP); y = y_base - h
            gc = QColor(self._bar_col); gc.setAlpha(30)
            p.setBrush(QBrush(gc)); p.setPen(Qt.NoPen); p.drawRoundedRect(x - 2, y - 3, BAR_W + 4, h + 6, 3, 3)
            grad = QLinearGradient(x, y + h, x, y); bc = QColor(self._bar_col)
            tc = QColor(min(255, bc.red()+80), min(255,bc.green()+80), min(255,bc.blue()+80), 250)
            grad.setColorAt(0.0, bc); grad.setColorAt(1.0, tc); p.setBrush(QBrush(grad))
            r = 1 if thin else 2; p.drawRoundedRect(x, y, BAR_W, h, r, r)
            hl = QColor(255, 255, 255, 55); p.setBrush(QBrush(hl))
            p.drawRoundedRect(x + 1, y, BAR_W - 2, min(3, h), 1, 1)

    def _viz_mirror(self, p):
        x0, y_base, BAR_W, GAP, MAX_H = self._bar_geometry(); half = MAX_H // 2
        for i, val in enumerate(self._bars):
            h = max(1, int(val * half)); x = x0 + i * (BAR_W + GAP)
            bc = QColor(self._bar_col); tc = QColor(min(255,bc.red()+80),min(255,bc.green()+80),min(255,bc.blue()+80),250)
            p.setPen(Qt.NoPen)
            grad = QLinearGradient(x, y_base, x, y_base - h)
            grad.setColorAt(0.0, bc); grad.setColorAt(1.0, tc); p.setBrush(QBrush(grad))
            p.drawRoundedRect(x, y_base - h, BAR_W, h, 2, 2)
            grad2 = QLinearGradient(x, y_base, x, y_base + h)
            grad2.setColorAt(0.0, bc); grad2.setColorAt(1.0, tc); p.setBrush(QBrush(grad2))
            p.drawRoundedRect(x, y_base, BAR_W, h, 2, 2)

    def _viz_wave(self, p):
        x0, y_base, BAR_W, GAP, MAX_H = self._bar_geometry(); N = len(self._bars)
        pts = [(x0 + i * (BAR_W + GAP) + BAR_W // 2, y_base - int(val * MAX_H))
               for i, val in enumerate(self._bars)]
        if len(pts) < 2: return
        path = QPainterPath(); path.moveTo(pts[0][0], pts[0][1])
        for i in range(1, len(pts)):
            x0p, y0p = pts[i-1]; x1p, y1p = pts[i]; cx1 = x0p + (x1p - x0p) * 0.5
            path.cubicTo(cx1, y0p, cx1, y1p, x1p, y1p)
        fill_path = QPainterPath(path)
        fill_path.lineTo(pts[-1][0], y_base); fill_path.lineTo(pts[0][0], y_base); fill_path.closeSubpath()
        bc = QColor(self._bar_col); bc.setAlpha(60); p.setBrush(QBrush(bc)); p.setPen(Qt.NoPen); p.drawPath(fill_path)
        pen_col = QColor(self._bar_col); pen_col.setAlpha(220)
        p.setPen(QPen(pen_col, 2.5, Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin)); p.setBrush(Qt.NoBrush); p.drawPath(path)
        p.setPen(Qt.NoPen); dc = QColor(255, 255, 255, 180); p.setBrush(QBrush(dc))
        for i, (x, y) in enumerate(pts):
            if self._bars[i] > 0.3: p.drawEllipse(QPointF(x, y), 2.5, 2.5)

    def _viz_dots(self, p):
        x0, y_base, BAR_W, GAP, MAX_H = self._bar_geometry(); DOT = max(4, BAR_W); STEP = DOT + 2
        for i, val in enumerate(self._bars):
            n_dots = max(1, int(val * MAX_H / STEP)); cx_dot = x0 + i * (BAR_W + GAP) + BAR_W // 2
            for d in range(n_dots):
                t = d / max(n_dots, 1); y_dot = y_base - d * STEP - DOT // 2; alpha = int(180 + t * 60)
                bc = QColor(self._bar_col)
                r = min(255, bc.red()+int(t*80)); g = min(255, bc.green()+int(t*80)); b = min(255, bc.blue()+int(t*80))
                col = QColor(r, g, b, alpha); gc = QColor(r, g, b, 40)
                p.setBrush(QBrush(gc)); p.setPen(Qt.NoPen); p.drawEllipse(QPointF(cx_dot, y_dot), DOT * 0.9, DOT * 0.9)
                p.setBrush(QBrush(col)); p.drawEllipse(QPointF(cx_dot, y_dot), DOT * 0.55, DOT * 0.55)

    def _viz_circle(self, p):
        import math
        cx, cy = self._viz_anchor(); N = len(self._bars); sh = self._sh
        R_MIN = int(sh * 0.06); R_MAX = int(sh * 0.18); bc_qt = QColor(self._bar_col)
        if not self.viz_detached: cy = cy - 84 - 18 - R_MAX - 10
        for i, val in enumerate(self._bars):
            angle = math.tau * i / N - math.pi / 2; r_out = R_MIN + int(val * (R_MAX - R_MIN))
            x1 = cx + math.cos(angle) * R_MIN; y1 = cy + math.sin(angle) * R_MIN
            x2 = cx + math.cos(angle) * r_out;  y2 = cy + math.sin(angle) * r_out
            t = val
            r = min(255, bc_qt.red()+int(t*80)); g = min(255, bc_qt.green()+int(t*80)); b = min(255, bc_qt.blue()+int(t*80))
            lw = max(2.0, 3.0 * val + 1.0)
            p.setPen(QPen(QColor(r, g, b, int(140+t*100)), lw, Qt.SolidLine, Qt.RoundCap))
            p.drawLine(QPointF(x1, y1), QPointF(x2, y2))
        p.setPen(Qt.NoPen); c2 = QColor(bc_qt); c2.setAlpha(180); p.setBrush(QBrush(c2))
        p.drawEllipse(QPointF(cx, cy), R_MIN * 0.4, R_MIN * 0.4)

    def mousePressEvent(self, e):
        if e.button() != Qt.LeftButton: return
        pos = e.pos()
        if self.viz_detached and self.show_viz and self._hit_viz(pos): self._drag = "viz"
        elif self.show_clock and self._hit_clock(pos): self._drag = "clock"
        elif self.show_viz and self._hit_viz(pos): self._drag = "clock"
        self._drag_origin = pos

    def mouseMoveEvent(self, e):
        if not self._drag or not self._drag_origin: return
        d = e.pos() - self._drag_origin; self._drag_origin = e.pos()
        if self._drag == "clock":
            self.cx = max(200, min(self._sw - 200, self.cx + d.x()))
            self.cy = max(100, min(self._sh - 100, self.cy + d.y()))
        elif self._drag == "viz":
            self.vx = max(100, min(self._sw - 100, self.vx + d.x()))
            self.vy = max(50,  min(self._sh - 50,  self.vy + d.y()))
        self.update()

    def mouseReleaseEvent(self, _):
        self._drag = None; self._drag_origin = None

    def _hit_viz(self, pos):
        ax, ay = self._viz_anchor(); sw = self._sw
        half_w = len(self._bars) * max(6, int(sw * 0.008)) // 2 + 20
        return abs(pos.x() - ax) < half_w and abs(pos.y() - ay) < int(self._sh * 0.25)

    def _hit_clock(self, pos):
        return abs(pos.x() - self.cx) < 350 and abs(pos.y() - self.cy) < 120

# ══════════════════════════════════════════════
#  MPV WALLPAPER
# ══════════════════════════════════════════════

class MpvWallpaper:
    def __init__(self):
        sw,sh = get_screen_size()
        self._sw,self._sh = sw,sh
        self._hwnd   = None
        self._player = None
        self._worker = None
        self._worker_stop = None
        self._embedded = False
        self._target   = 0
        self._reembed_cb = None
        self._path = None
        self._watchdog_stop   = None
        self._watchdog_thread = None
        self.frame_ready_cb = None

    def is_playing(self):
        if not self._player: return False
        try: return not self._player.pause
        except: return False

    def load(self, path, enable_audio=False):
        self.stop()
        if not MPV_OK: return False
        self._path = path
        _pm = ctypes.windll.user32.FindWindowW("Progman", None)
        emb = DesktopEmbedder()
        if _pm: emb._send_spawn_msg(_pm)
        target = emb._find_target()
        if not target:
            print("[mpv] Could not find WorkerW target"); return False
        sw, sh = self._sw, self._sh
        if not self._create_hwnd_as_child(target, sw, sh): return False
        u32 = ctypes.windll.user32
        defview = u32.FindWindowExW(target, None, "SHELLDLL_DefView", None)
        insert_after = defview if defview else 1
        u32.SetWindowPos(self._hwnd, insert_after, 0, 0, sw, sh, 0x0010|0x0040|0x0020)
        u32.ShowWindow(self._hwnd, 5); u32.UpdateWindow(self._hwnd)
        time.sleep(0.15)
        self._embedded_target = target
        try:
            opts = {
                'wid'                    : int(self._hwnd),
                'loop_file'              : 'inf',
                'force_window'           : 'immediate',
                'vo'                     : 'gpu-next',
                'gpu_api'                : 'd3d11',
                'gpu_context'            : 'd3d11',
                'hwdec'                  : 'd3d11va',
                'keepaspect'             : 'yes',
                'panscan'                : '1.0',
                'input_default_bindings' : False,
                'input_vo_keyboard'      : False,
                'cursor_autohide'        : 'no',
                'ao'                     : 'wasapi' if enable_audio else 'null',
                'volume'                 : 100 if enable_audio else 0,
                'mute'                   : 'no'  if enable_audio else 'yes',
            }
            self._player = mpv.MPV(**opts)
            self._player.play(path)
            time.sleep(0.3)
            self._start_extractor()
            return True
        except Exception as e:
            print(f"[mpv] load failed: {e}")
            self._destroy_hwnd(); return False

    def pause(self):
        if self._player:
            try: self._player.pause = True
            except: pass

    def resume(self):
        if self._player:
            try: self._player.pause = False
            except: pass

    def stop(self):
        self._stop_watchdog()
        if self._worker_stop: self._worker_stop.set()
        if self._worker and self._worker.is_alive(): self._worker.join(2)
        self._worker = self._worker_stop = None
        player = self._player; self._player = None
        if player:
            done = threading.Event()
            def _kill():
                try: player.terminate()
                except: pass
                try: player.wait_for_shutdown(timeout=2)
                except: pass
                done.set()
            t = threading.Thread(target=_kill, daemon=True); t.start(); done.wait(timeout=2.5)
            time.sleep(0.1)
        self._destroy_hwnd()
        self._embedded = False; self._target = 0

    def show(self):
        if self._hwnd: ctypes.windll.user32.ShowWindow(self._hwnd, 4)
    def hide(self):
        if self._hwnd: ctypes.windll.user32.ShowWindow(self._hwnd, 0)
    def winId(self): return self._hwnd or 0
    def close(self): self.stop()

    def set_embedded(self, state, target=0):
        self._embedded = state; self._target = target
        if state and self._reembed_cb: self._start_watchdog()

    def _start_watchdog(self):
        self._stop_watchdog()
        self._watchdog_stop = threading.Event()
        self._watchdog_thread = threading.Thread(target=self._watchdog_run, daemon=True)
        self._watchdog_thread.start()

    def _stop_watchdog(self):
        if self._watchdog_stop: self._watchdog_stop.set(); self._watchdog_stop=None

    def _watchdog_run(self):
        u32 = ctypes.windll.user32; last_tray = u32.FindWindowW('Shell_TrayWnd', None)
        last_parent = self._target
        while self._watchdog_stop and not self._watchdog_stop.wait(1.0):
            if not self._embedded or not self._hwnd: break
            needs = False; tray = u32.FindWindowW('Shell_TrayWnd', None)
            if tray and tray != last_tray: last_tray=tray; needs=True
            cp = u32.GetParent(self._hwnd)
            if cp != last_parent: last_parent=cp; needs=True
            if needs and self._reembed_cb:
                try: self._reembed_cb()
                except: pass

    def _create_hwnd_as_child(self, parent_hwnd, sw, sh):
        global _WND_CLASS_REGISTERED
        if not _WND_CLASS_REGISTERED:
            if not self._register_wnd_class(): return False
        hInst = ctypes.windll.kernel32.GetModuleHandleW(None)
        hwnd = ctypes.windll.user32.CreateWindowExW(
            0, "WallpaperEngineHost", "WallpaperEngineVideo",
            0x40000000|0x02000000|0x04000000|0x10000000,
            0, 0, sw, sh, parent_hwnd, None, hInst, None)
        if not hwnd:
            print(f"[mpv] CreateWindowExW (child) failed: {ctypes.GetLastError()}"); return False
        self._hwnd = hwnd
        return True

    def _register_wnd_class(self):
        global _WND_CLASS_REGISTERED, _WNDPROC_CB
        WNDPROC_TYPE = ctypes.WINFUNCTYPE(
            ctypes.c_ssize_t, ctypes.wintypes.HWND, ctypes.wintypes.UINT,
            ctypes.c_size_t, ctypes.c_ssize_t)
        _DefWndProc = ctypes.windll.user32.DefWindowProcW
        _DefWndProc.restype  = ctypes.c_ssize_t
        _DefWndProc.argtypes = [ctypes.wintypes.HWND, ctypes.wintypes.UINT, ctypes.c_size_t, ctypes.c_ssize_t]
        def wndproc(hwnd, msg, wp, lp): return _DefWndProc(hwnd, msg, wp, lp)
        _WNDPROC_CB = WNDPROC_TYPE(wndproc)

        class WNDCLASSEXW(ctypes.Structure):
            _fields_ = [("cbSize", ctypes.wintypes.UINT), ("style", ctypes.wintypes.UINT),
                        ("lpfnWndProc", WNDPROC_TYPE), ("cbClsExtra", ctypes.c_int),
                        ("cbWndExtra", ctypes.c_int), ("hInstance", ctypes.wintypes.HANDLE),
                        ("hIcon", ctypes.wintypes.HANDLE), ("hCursor", ctypes.wintypes.HANDLE),
                        ("hbrBackground", ctypes.wintypes.HANDLE), ("lpszMenuName", ctypes.wintypes.LPCWSTR),
                        ("lpszClassName", ctypes.wintypes.LPCWSTR), ("hIconSm", ctypes.wintypes.HANDLE)]

        wc = WNDCLASSEXW()
        wc.cbSize = ctypes.sizeof(WNDCLASSEXW); wc.lpfnWndProc = _WNDPROC_CB
        wc.hInstance = ctypes.windll.kernel32.GetModuleHandleW(None); wc.hbrBackground = 0
        wc.lpszClassName = "WallpaperEngineHost"
        res = ctypes.windll.user32.RegisterClassExW(ctypes.byref(wc))
        if res == 0:
            err = ctypes.GetLastError()
            if err == 1410: _WND_CLASS_REGISTERED = True; return True
            print(f"[mpv] RegisterClassExW failed: {err}"); return False
        _WND_CLASS_REGISTERED = True; return True

    def _destroy_hwnd(self):
        if self._hwnd: ctypes.windll.user32.DestroyWindow(self._hwnd); self._hwnd=None

    def _start_extractor(self):
        self._worker_stop = threading.Event(); path = self._path
        def run():
            cap = None; time.sleep(1.5)
            while not self._worker_stop.wait(3.5):
                if not self._player or not self.frame_ready_cb: continue
                try:
                    if cap is None: cap = cv2.VideoCapture(path)
                    if not cap or not cap.isOpened(): cap = None; continue
                    try:
                        pos = self._player.time_pos
                        if pos is not None: cap.set(cv2.CAP_PROP_POS_MSEC, float(pos) * 1000)
                    except Exception: pass
                    ret, frame = cap.read()
                    if ret and frame is not None: self.frame_ready_cb(frame)
                except Exception as e: print(f"[Extract] {e}")
            if cap: cap.release()
        self._worker = threading.Thread(target=run, daemon=True); self._worker.start()

# ══════════════════════════════════════════════
#  HEADLESS ENGINE  (no GUI panel)
# ══════════════════════════════════════════════

class HeadlessEngine(QObject):
    """
    Applies all settings from _CFG without opening any window.
    Runs as a background Qt object; the only UI is the system tray icon.
    """
    _frame_signal = pyqtSignal(object)

    def __init__(self, cfg):
        super().__init__()
        self._cfg    = cfg
        self._ext    = ColorExtractor()
        self._themer = WindowsThemer()
        self._wp     = MpvWallpaper()
        self._audio  = AudioVisualizer()
        self._audio_on = False

        sw, sh = get_screen_size()
        self._ov = Overlay(sw, sh)
        self._ov.apply_config(cfg)

        self._frame_signal.connect(self._on_frame)
        self._wp.frame_ready_cb = self._frame_signal.emit
        self._wp._reembed_cb    = lambda: QTimer.singleShot(0, self._reembed)

        self._audio.bars_updated.connect(self._on_bars)

        self._playing  = False
        self._embedded = False

        self._build_tray()

        delay = int(cfg.get("startup_delay", 0))
        if delay > 0:
            print(f"[Config] Startup delay: {delay}s")
            QTimer.singleShot(delay * 1000, self._start)
        else:
            QTimer.singleShot(200, self._start)

        # Periodic colour refresh (every 8 s)
        QTimer(self, timeout=self._refresh_theme, interval=8000).start()
        # Fullscreen auto-pause
        QTimer(self, timeout=self._check_desktop,  interval=800).start()

    def _build_tray(self):
        r, g, b = self._cfg.get("static_color", [70, 130, 220])
        self._tray = QSystemTrayIcon(QIcon(make_tray_icon_pixmap(r, g, b)), self)
        self._tray.setToolTip("Theme Engine (headless)")
        menu = QMenu()
        act_quit = menu.addAction("Quit Theme Engine")
        act_show_cfg = menu.addAction(f"Config: {os.path.basename(_CONFIG_PATH)}")
        act_show_cfg.setEnabled(False)
        act_quit.triggered.connect(self._quit)
        self._tray.setContextMenu(menu)
        self._tray.activated.connect(
            lambda r: self._tray.showMessage(
                "Theme Engine", "Running headlessly. Right-click to quit.",
                QSystemTrayIcon.Information, 2000)
            if r == QSystemTrayIcon.Trigger else None)
        self._tray.show()

    def _start(self):
        cfg = self._cfg
        vpath = cfg.get("video_path", "")

        if vpath and os.path.exists(vpath):
            print(f"[Headless] Loading video: {vpath}")
            enable_audio = cfg.get("video_audio", False)
            ok = self._wp.load(vpath, enable_audio=enable_audio)
            if ok:
                self._playing = True
                hw = self._wp.winId()
                if hw:
                    self._wp.set_embedded(True, getattr(self._wp, '_embedded_target', 0))
                    self._ov.show()
                    self._ov.raise_()
                    self._embedded = True
                    print("[Headless] Wallpaper active")
                # Auto-start audio
                QTimer.singleShot(1500, self._maybe_start_audio)
                # Force theme after extractor has had time to sample
                QTimer.singleShot(5000, self._force_apply_theme)
            else:
                print("[Headless] WARNING: video failed to load — check MPV/path")
        elif vpath:
            print(f"[Headless] WARNING: video_path not found: {vpath}")
        else:
            print("[Headless] No video_path set — running overlay/colour only")
            self._ov.show()

        # Apply static colour immediately if in static mode
        if cfg.get("palette_mode", "dynamic") == "static":
            sc = cfg.get("static_color", [70, 130, 220])
            self._apply_accent(sc[0], sc[1], sc[2])

    def _maybe_start_audio(self):
        cfg = self._cfg
        if not cfg.get("auto_start_audio", True): return
        if not AUDIO_OK or self._audio_on: return
        if self._audio.can_capture():
            self._audio.start()
            self._audio_on = True
            print("[Headless] Audio visualizer started")
        else:
            print("[Headless] No loopback capture available")

    @pyqtSlot(object)
    def _on_frame(self, frame):
        colors = self._ext.extract(frame)
        if self._cfg.get("viz_auto_color", True):
            self._ov.set_bar_col(*self._ext.get_glow())

    @pyqtSlot(list)
    def _on_bars(self, bars):
        self._ov.set_bars(bars)

    def _apply_accent(self, r, g, b):
        if self._cfg.get("adaptive_taskbar_color", True):
            threading.Thread(target=self._themer.apply, args=(r, g, b), daemon=True).start()

    def _refresh_theme(self):
        if not self._playing: return
        if self._cfg.get("palette_mode", "dynamic") != "dynamic": return
        acc = self._ext.get_accent()
        self._apply_accent(*acc)
        # Update tray icon colour
        self._tray.setIcon(QIcon(make_tray_icon_pixmap(*acc)))

    def _force_apply_theme(self):
        if self._cfg.get("palette_mode", "dynamic") != "dynamic": return
        acc = self._ext.get_accent()
        self._apply_accent(*acc)

    def _reembed(self):
        if not self._embedded or not self._wp._path: return
        sw, sh = get_screen_size(); hw = self._wp.winId()
        if not hw: return
        emb = DesktopEmbedder()
        ok  = emb.embed(hw, sw, sh)
        if ok: self._wp.set_embedded(True, emb.get_target())

    def _check_desktop(self):
        if not self._cfg.get("pause_on_fullscreen", True): return
        if not self._playing: return
        hwnd = ctypes.windll.user32.GetForegroundWindow()
        if not hwnd: return
        rc = ctypes.wintypes.RECT(); ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rc))
        sw, sh = get_screen_size()
        covers = rc.left<=0 and rc.top<=0 and rc.right>=sw and rc.bottom>=sh
        shell  = get_foreground_class() in ("Progman", "WorkerW", "")
        playing = self._wp.is_playing()
        if covers and not shell and playing:         self._wp.pause()
        elif (not covers or shell) and not playing:  self._wp.resume()

    def _quit(self):
        self._audio.stop()
        self._wp.stop()
        self._ov.close()
        self._themer.restore()
        self._tray.hide()
        QApplication.instance().quit()

# ══════════════════════════════════════════════
#  WIDGETS  (GUI mode only)
# ══════════════════════════════════════════════

class DropZone(QLabel):
    video_chosen = pyqtSignal(str)
    EXTS = ('.mp4','.avi','.mkv','.mov','.wmv','.webm','.m4v','.flv')
    FILT = "Video files (*.mp4 *.avi *.mkv *.mov *.wmv *.webm *.m4v *.flv)"
    def __init__(self, mpv_available=True):
        super().__init__()
        self._mpv_ok = mpv_available
        self.setAcceptDrops(True); self.setAlignment(Qt.AlignCenter); self.setFixedHeight(88); self._idle()
    def _idle(self):
        self.setText("🎬   Drop video here  ·  or click to browse" if self._mpv_ok else "⚠️   MPV not loaded — click for help")
        self.setStyleSheet("QLabel{border:1.5px dashed rgba(255,255,255,45);border-radius:12px;color:rgba(255,255,255,110);font-size:12px;padding:6px;}")
    def _hover(self):
        if not self._mpv_ok: return
        self.setStyleSheet("QLabel{border:1.5px dashed rgba(140,220,255,200);border-radius:12px;background:rgba(140,220,255,12);color:white;font-size:12px;padding:6px;}")
    def mousePressEvent(self,_):
        if not self._mpv_ok:
            QMessageBox.information(self,"MPV Setup","Put libmpv-2.dll next to this script.\nInstall: https://aka.ms/vs/17/release/vc_redist.x64.exe"); return
        p,_=QFileDialog.getOpenFileName(self,"Choose video","",self.FILT)
        if p: self.video_chosen.emit(p)
    def dragEnterEvent(self,e):
        if not self._mpv_ok: return
        if e.mimeData().hasUrls() and e.mimeData().urls()[0].toLocalFile().lower().endswith(self.EXTS):
            e.acceptProposedAction(); self._hover()
    def dragLeaveEvent(self,_): self._idle()
    def dropEvent(self,e):
        if not self._mpv_ok: return
        u=e.mimeData().urls()[0].toLocalFile()
        if u.lower().endswith(self.EXTS): self.video_chosen.emit(u)
        self._idle()

class PaletteBar(QWidget):
    def __init__(self):
        super().__init__()
        self._c=[(20,50,120)]*6; self._t=list(self._c); self._dynamic=True
        QTimer(self,timeout=self._step,interval=120).start(); self.setFixedHeight(36)
    def update_colors(self,c):
        if self._dynamic: self._t=(c+[(20,20,20)]*6)[:6]
    def set_dynamic(self,v): self._dynamic=v
    def set_static_colors(self,c): self._t=(c+[(20,20,20)]*6)[:6]; self._c=list(self._t)
    def _step(self):
        ch=False
        for i,(c,t) in enumerate(zip(self._c,self._t)):
            nc=tuple(int(cv+(tv-cv)*.18) for cv,tv in zip(c,t))
            if nc!=c: self._c[i]=nc; ch=True
        if ch: self.update()
    def paintEvent(self,_):
        p=QPainter(self); p.setRenderHint(QPainter.Antialiasing)
        w,h=self.width(),self.height(); sw=(w-8)//6; x=4
        for r,g,b in self._c:
            p.setBrush(QBrush(QColor(r,g,b))); p.setPen(Qt.NoPen)
            p.drawRoundedRect(x,2,sw-3,h-4,6,6); x+=sw
        p.end()

class IconBtn(QPushButton):
    def __init__(self,txt,tip=""):
        super().__init__(txt); self.setToolTip(tip); self.setFixedSize(38,38)
        self.setCursor(QCursor(Qt.PointingHandCursor)); self._acc=(70,130,220); self._on=False; self._re()
    def set_accent(self,r,g,b): self._acc=(r,g,b); self._re()
    def set_active(self,v):     self._on=v; self._re()
    def _re(self):
        r,g,b=self._acc
        if self._on:
            self.setStyleSheet(f"QPushButton{{background:rgb({r},{g},{b});color:white;border:none;border-radius:10px;font-size:14px;font-weight:700;}}QPushButton:hover{{background:rgb({min(255,r+30)},{min(255,g+30)},{min(255,b+30)});}}")
        else:
            self.setStyleSheet(f"QPushButton{{background:rgba({r},{g},{b},30);color:white;border:1.5px solid rgba({r},{g},{b},80);border-radius:10px;font-size:14px;font-weight:700;}}QPushButton:hover{{background:rgba({r},{g},{b},70);}}QPushButton:disabled{{color:rgba(255,255,255,30);border-color:rgba(255,255,255,18);background:rgba(255,255,255,5);}}")

class WideBtn(QPushButton):
    def __init__(self,txt):
        super().__init__(txt); self.setCursor(QCursor(Qt.PointingHandCursor)); self.setFixedHeight(46); self._sc(70,130,220)
    def _sc(self,r,g,b):
        self._acc=(r,g,b); r2,g2,b2=max(0,r-40),max(0,g-40),max(0,b-40)
        self.setStyleSheet(f"QPushButton{{background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 rgb({r},{g},{b}),stop:1 rgb({r2},{g2},{b2}));color:white;border:none;border-radius:11px;font-size:13px;font-weight:700;}}QPushButton:hover{{background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 rgb({min(255,r+30)},{min(255,g+30)},{min(255,b+30)}),stop:1 rgb({r},{g},{b}));}}QPushButton:disabled{{background:rgba(255,255,255,8);color:rgba(255,255,255,35);}}")
    def set_accent(self,r,g,b): self._sc(r,g,b)

class StatusBand(QLabel):
    def __init__(self):
        super().__init__("Ready"); self.setAlignment(Qt.AlignCenter); self.setFixedHeight(28); self._r,self._g,self._b=70,130,220; self._re()
    def set_accent(self,r,g,b): self._r,self._g,self._b=r,g,b; self._re()
    def _re(self):
        r,g,b=self._r,self._g,self._b
        self.setStyleSheet(f"QLabel{{background:rgba({r},{g},{b},22);color:rgb({min(255,r+110)},{min(255,g+110)},{min(255,b+110)});border-radius:6px;font-size:11px;font-weight:600;padding:0 10px;}}")
    def ok(self,m):      self.setText(f"✅  {m}")
    def info(self,m):    self.setText(f"ℹ  {m}")
    def err(self,m):     self.setText(f"⚠  {m}")
    def playing(self,m): self.setText(f"▶  {m}")
    def paused(self,m):  self.setText(f"⏸  {m}")

class Divider(QWidget):
    def __init__(self,r=70,g=130,b=220):
        super().__init__(); self.setFixedHeight(1); self._c=QColor(r,g,b,50)
    def set_color(self,r,g,b): self._c=QColor(r,g,b,50); self.update()
    def paintEvent(self,_):
        p=QPainter(self); p.fillRect(0,0,self.width(),1,self._c); p.end()

def section(txt,r=70,g=130,b=220):
    lbl=QLabel(txt)
    lbl.setStyleSheet(f"color:rgba({min(255,r+130)},{min(255,g+130)},{min(255,b+130)},210);font-size:10px;font-weight:800;letter-spacing:1.2px;")
    lbl.setFixedHeight(20); return lbl

# ══════════════════════════════════════════════
#  MAIN PANEL  (GUI mode)
# ══════════════════════════════════════════════

class Panel(QWidget):
    _frame_signal = pyqtSignal(object)

    def __init__(self):
        super().__init__()
        self._ext    = ColorExtractor()
        self._themer = WindowsThemer()
        self._emb    = DesktopEmbedder()
        self._wp     = MpvWallpaper()
        self._frame_signal.connect(self._on_frame)
        self._wp.frame_ready_cb  = self._frame_signal.emit
        self._wp._reembed_cb     = lambda: QTimer.singleShot(0, self._do_reembed)

        sw,sh = get_screen_size()
        self._ov = Overlay(sw,sh)

        self._audio    = AudioVisualizer()
        self._audio.bars_updated.connect(self._on_bars)
        self._audio_on = False

        self._embedded = False
        self._acc      = (55,120,200)
        self._vpath    = None
        self._playing  = False
        self._upause   = False
        self._palette_dynamic = True

        self._build_ui()
        self._build_tray()

        if is_admin(): self._status.err("Running as Admin — embed may fail!")

        QTimer(self, timeout=self._refresh_theme, interval=8000).start()
        QTimer(self, timeout=self._check_desktop,  interval=800).start()

        if not MPV_OK: QTimer.singleShot(100, self._show_mpv_warning)

    def _show_mpv_warning(self):
        has_dll = os.path.exists(os.path.join(_script_dir,"libmpv-2.dll"))
        msg = QMessageBox(self); msg.setWindowTitle("MPV Not Loaded"); msg.setIcon(QMessageBox.Warning)
        if has_dll:
            msg.setText("libmpv-2.dll found but failed to load.")
            msg.setInformativeText("Install Visual C++ Redistributables:\nhttps://aka.ms/vs/17/release/vc_redist.x64.exe\n\nThen restart.")
        else:
            msg.setText("libmpv-2.dll not found.")
            msg.setInformativeText(
                "1. Download from:\n   https://github.com/shinchiro/mpv-winbuild-cmake/releases\n\n"
                f"2. Place libmpv-2.dll in:\n   {_script_dir}\n\n"
                "3. Install VC++ Redist:\n   https://aka.ms/vs/17/release/vc_redist.x64.exe\n\n"
                "4. Restart this app")
        msg.exec_()

    def _build_tray(self):
        r,g,b=self._acc
        self._tray=QSystemTrayIcon(QIcon(make_tray_icon_pixmap(r,g,b)),self)
        self._tray.setToolTip("Theme Engine")
        menu=QMenu()
        self._tm_show =menu.addAction("Show Panel")
        self._tm_pause=menu.addAction("Pause")
        menu.addSeparator()
        self._tm_quit =menu.addAction("Quit")
        self._tm_show.triggered.connect(self._show_panel)
        self._tm_pause.triggered.connect(self._tray_pause_toggle)
        self._tm_quit.triggered.connect(self._real_quit)
        self._tray.setContextMenu(menu)
        self._tray.activated.connect(lambda r: self._show_panel() if r==QSystemTrayIcon.Trigger else None)
        self._tray.show()

    def _show_panel(self): self.show(); self.raise_(); self.activateWindow()
    def _tray_pause_toggle(self):
        if self._upause: self._play()
        else:            self._pause()
    def _real_quit(self): self._cleanup(); QApplication.instance().quit()
    def _cleanup(self):
        if self._embedded: self._emb.detach(self._wp.winId())
        self._audio.stop(); self._wp.stop(); self._ov.close()
        self._themer.restore(); self._tray.hide()
    def closeEvent(self,e):
        e.ignore(); self.hide()
        self._tray.showMessage("Theme Engine","Running in background — right-click tray to quit.",QSystemTrayIcon.Information,2500)

    def _build_ui(self):
        self.setWindowTitle("Theme Engine v4.3")
        self.setFixedWidth(460); self.setWindowFlags(Qt.Window)
        ml=QVBoxLayout(self); ml.setContentsMargins(0,0,0,0); ml.setSpacing(0)
        self._hdr=QWidget(); self._hdr.setFixedHeight(62)
        hl=QHBoxLayout(self._hdr); hl.setContentsMargins(18,0,16,0); hl.setSpacing(0)
        logo=QLabel("▶"); logo.setStyleSheet("font-size:22px;color:white;padding-right:10px;")
        title=QLabel("Theme Engine"); title.setStyleSheet("font-size:17px;font-weight:800;color:white;letter-spacing:.5px;")
        ver=QLabel("v4.3"); ver.setStyleSheet("font-size:10px;color:rgba(255,255,255,130);padding-left:6px;padding-top:5px;")
        hl.addWidget(logo); hl.addWidget(title); hl.addWidget(ver); hl.addStretch()
        ml.addWidget(self._hdr)
        scroll=QScrollArea(); scroll.setWidgetResizable(True); scroll.setFrameShape(QScrollArea.NoFrame)
        body=QWidget(); bl=QVBoxLayout(body); bl.setContentsMargins(16,14,16,18); bl.setSpacing(0)
        def gap(n=8): s=QWidget(); s.setFixedHeight(n); return s

        # Config info band (shown when a config file exists but headless=False)
        if _CFG is not None:
            cfg_band = QLabel(f"⚙  Config loaded: {os.path.basename(_CONFIG_PATH)}")
            cfg_band.setStyleSheet("background:rgba(124,111,255,30);color:rgba(200,190,255,200);border-radius:6px;font-size:11px;padding:4px 10px;")
            bl.addWidget(cfg_band); bl.addWidget(gap(6))

        bl.addWidget(section("VIDEO SOURCE")); bl.addWidget(gap(6))
        self._drop=DropZone(mpv_available=MPV_OK); self._drop.video_chosen.connect(self._load)
        bl.addWidget(self._drop); bl.addWidget(gap(4))
        self._vid_lbl=QLabel("No video loaded"); self._vid_lbl.setAlignment(Qt.AlignCenter)
        self._vid_lbl.setStyleSheet("color:rgba(255,255,255,80);font-size:11px;")
        bl.addWidget(self._vid_lbl); bl.addWidget(gap(10))
        bl.addWidget(section("PLAYBACK")); bl.addWidget(gap(6))
        tr=QWidget(); trl=QHBoxLayout(tr); trl.setContentsMargins(0,0,0,0); trl.setSpacing(6)
        self._b_play=IconBtn("▶","Play"); self._b_pause=IconBtn("⏸","Pause"); self._b_stop=IconBtn("⏹","Stop")
        for b in (self._b_play,self._b_pause,self._b_stop): b.setEnabled(False); trl.addWidget(b)
        self._b_play.clicked.connect(self._play); self._b_pause.clicked.connect(self._pause); self._b_stop.clicked.connect(self._stop)
        self._b_embed=WideBtn("🖥  Set as Live Wallpaper"); self._b_embed.setEnabled(False)
        self._b_embed.clicked.connect(self._toggle_embed); trl.addWidget(self._b_embed,1)
        bl.addWidget(tr); bl.addWidget(gap(10))
        self._status=StatusBand(); bl.addWidget(self._status); bl.addWidget(gap(14))
        bl.addWidget(Divider(*self._acc)); bl.addWidget(gap(10))
        bl.addWidget(section("ADAPTIVE COLOUR PALETTE")); bl.addWidget(gap(6))
        self._pal=PaletteBar(); bl.addWidget(self._pal); bl.addWidget(gap(6))
        pr=QWidget(); prl=QHBoxLayout(pr); prl.setContentsMargins(0,0,0,0); prl.setSpacing(12)
        from PyQt5.QtWidgets import QRadioButton, QButtonGroup
        self._rad_dyn  = QRadioButton("Dynamic (follows video)")
        self._rad_stat = QRadioButton("Static colour")
        self._rad_dyn.setChecked(True)
        self._pal_grp  = QButtonGroup(self); self._pal_grp.addButton(self._rad_dyn,0); self._pal_grp.addButton(self._rad_stat,1)
        self._b_statcol= IconBtn("■","Pick static colour"); self._b_statcol.setFixedSize(32,30)
        self._b_statcol.setEnabled(False); self._b_statcol.clicked.connect(self._pick_static_col)
        prl.addWidget(self._rad_dyn); prl.addWidget(self._rad_stat); prl.addWidget(self._b_statcol); prl.addStretch()
        self._pal_grp.buttonClicked.connect(self._on_palette_mode)
        bl.addWidget(pr); bl.addWidget(gap(4))
        self._chk_adapt=QCheckBox("Apply colour to Windows taskbar & title bars"); self._chk_adapt.setChecked(True)
        bl.addWidget(self._chk_adapt); bl.addWidget(gap(4))
        self._b_taskbar_test=WideBtn("⚙️  Open Taskbar Colour Settings (one-time setup)")
        self._b_taskbar_test.clicked.connect(self._open_taskbar_settings)
        bl.addWidget(self._b_taskbar_test); bl.addWidget(gap(14))
        bl.addWidget(Divider(*self._acc)); bl.addWidget(gap(10))
        bl.addWidget(section("CLOCK & DATE OVERLAY")); bl.addWidget(gap(8))
        cr=QWidget(); crl=QHBoxLayout(cr); crl.setContentsMargins(0,0,0,0); crl.setSpacing(10)
        self._chk_clock=QCheckBox("Show clock"); self._chk_clock.setChecked(True)
        self._chk_secs =QCheckBox("Show seconds"); self._chk_secs.setChecked(True)
        crl.addWidget(self._chk_clock); crl.addWidget(self._chk_secs); crl.addStretch()
        self._chk_clock.stateChanged.connect(lambda v: setattr(self._ov,'show_clock',bool(v)))
        self._chk_secs.stateChanged.connect(lambda v: setattr(self._ov,'show_secs',bool(v)))
        bl.addWidget(cr); bl.addWidget(gap(7))
        fr=QWidget(); frl=QHBoxLayout(fr); frl.setContentsMargins(0,0,0,0); frl.setSpacing(8)
        fl=QLabel("Font:"); fl.setFixedWidth(34); frl.addWidget(fl)
        self._font_cb=QFontComboBox(); self._font_cb.setCurrentFont(QFont("Segoe UI Light")); self._font_cb.setFixedHeight(30)
        self._font_cb.currentFontChanged.connect(lambda f: setattr(self._ov,'font_family',f.family()))
        frl.addWidget(self._font_cb,1); bl.addWidget(fr); bl.addWidget(gap(7))
        co=QWidget(); col=QHBoxLayout(co); col.setContentsMargins(0,0,0,0); col.setSpacing(8)
        self._b_tcol=IconBtn("A","Text Color"); self._b_tcol.setFixedSize(32,30); self._b_tcol.clicked.connect(self._pick_text)
        col.addWidget(self._b_tcol); col.addWidget(QLabel("Opacity:"))
        self._sld_op=QSlider(Qt.Horizontal); self._sld_op.setRange(40,255); self._sld_op.setValue(220)
        self._sld_op.valueChanged.connect(self._on_opacity); col.addWidget(self._sld_op,1)
        bl.addWidget(co); bl.addWidget(gap(5))
        self._chk_sparkle = QCheckBox("Show sparkle cursor trail")
        self._chk_sparkle.setChecked(True)
        self._chk_sparkle.stateChanged.connect(
            lambda v: (setattr(self._ov, '_sparkles_enabled', bool(v)),
                       setattr(self._ov, '_sparkles', []) if not v else None))
        bl.addWidget(self._chk_sparkle); bl.addWidget(gap(4))
        hint=QLabel("💡 Drag clock/viz on desktop to reposition"); hint.setStyleSheet("color:rgba(255,255,255,65);font-size:10px;font-style:italic;")
        bl.addWidget(hint); bl.addWidget(gap(14))
        bl.addWidget(Divider(*self._acc)); bl.addWidget(gap(10))
        bl.addWidget(section("MUSIC VISUALIZER")); bl.addWidget(gap(8))
        vr=QWidget(); vrl=QHBoxLayout(vr); vrl.setContentsMargins(0,0,0,0); vrl.setSpacing(10)
        self._chk_viz=QCheckBox("Show"); self._chk_viz.setChecked(True)
        self._chk_vauto=QCheckBox("Auto-colour"); self._chk_vauto.setChecked(True)
        self._chk_vdetach=QCheckBox("Detach from clock"); self._chk_vdetach.setChecked(False)
        vrl.addWidget(self._chk_viz); vrl.addWidget(self._chk_vauto)
        vrl.addWidget(self._chk_vdetach); vrl.addStretch()
        self._chk_viz.stateChanged.connect(lambda v: setattr(self._ov,'show_viz',bool(v)))
        self._chk_vdetach.stateChanged.connect(self._on_viz_detach)
        bl.addWidget(vr); bl.addWidget(gap(5))
        from PyQt5.QtWidgets import QComboBox
        sr2=QWidget(); srl=QHBoxLayout(sr2); srl.setContentsMargins(0,0,0,0); srl.setSpacing(8)
        srl.addWidget(QLabel("Style:"))
        self._viz_style_cb=QComboBox()
        for s in ["Bars","Slim bars","Mirror","Wave","Dots","Circle"]:
            self._viz_style_cb.addItem(s)
        self._viz_style_cb.setFixedHeight(28)
        self._viz_style_cb.currentIndexChanged.connect(self._on_viz_style)
        srl.addWidget(self._viz_style_cb,1)
        bl.addWidget(sr2); bl.addWidget(gap(7))
        vr2=QWidget(); vrl2=QHBoxLayout(vr2); vrl2.setContentsMargins(0,0,0,0); vrl2.setSpacing(8)
        self._b_bcol=IconBtn("🎨","Bar Color"); self._b_bcol.setFixedSize(32,30); self._b_bcol.clicked.connect(self._pick_bar)
        vrl2.addWidget(self._b_bcol)
        if AUDIO_OK:
            self._b_audio=WideBtn("▶  Start Audio Capture"); self._b_audio.clicked.connect(self._toggle_audio); vrl2.addWidget(self._b_audio,1)
        else:
            vrl2.addWidget(QLabel("  pip install sounddevice"),1)
        bl.addWidget(vr2); bl.addWidget(gap(4))
        bl.addWidget(QLabel("WASAPI loopback · no permissions needed · drag viz to reposition when detached"))
        bl.addWidget(gap(14))
        bl.addWidget(Divider(*self._acc)); bl.addWidget(gap(10))
        bl.addWidget(section("BEHAVIOUR")); bl.addWidget(gap(8))
        self._chk_autopause=QCheckBox("Pause when a fullscreen app covers the desktop"); self._chk_autopause.setChecked(True)
        self._chk_loop=QCheckBox("Loop video"); self._chk_loop.setChecked(True)
        bl.addWidget(self._chk_autopause); bl.addWidget(gap(4))
        bl.addWidget(self._chk_loop);      bl.addWidget(gap(16))
        scroll.setWidget(body); ml.addWidget(scroll)
        self._apply_theme(*self._acc); self.resize(460,780)

    def _apply_theme(self,r,g,b):
        self._acc=(r,g,b); dr,dg,db=max(0,r-60),max(0,g-60),max(0,b-60)
        bg=f"rgb({max(0,r-110)},{max(0,g-110)},{max(0,b-110)})"
        self.setStyleSheet(f"""
            QWidget{{background:{bg};color:rgba(255,255,255,185);}}
            QScrollArea{{background:transparent;border:none;}}
            QScrollBar:vertical{{background:rgba(255,255,255,12);width:6px;border-radius:3px;}}
            QScrollBar::handle:vertical{{background:rgba({r},{g},{b},120);border-radius:3px;}}
            QCheckBox{{color:rgba(255,255,255,185);font-size:12px;spacing:8px;}}
            QCheckBox::indicator{{width:15px;height:15px;border-radius:4px;border:1.5px solid rgba({r},{g},{b},130);background:rgba({r},{g},{b},12);}}
            QCheckBox::indicator:checked{{background:rgb({r},{g},{b});border-color:rgb({r},{g},{b});}}
            QLabel{{color:rgba(255,255,255,170);background:transparent;}}
            QFontComboBox{{background:rgba({r},{g},{b},22);border:1.5px solid rgba({r},{g},{b},80);border-radius:7px;color:white;padding:2px 8px;font-size:12px;}}
            QSlider::groove:horizontal{{background:rgba({r},{g},{b},40);height:4px;border-radius:2px;}}
            QSlider::handle:horizontal{{background:rgb({r},{g},{b});width:14px;height:14px;margin:-5px 0;border-radius:7px;}}
        """)
        self._hdr.setStyleSheet(f"background:qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 rgb({r},{g},{b}),stop:1 rgb({dr},{dg},{db}));")
        for w in (self._b_play,self._b_pause,self._b_stop): w.set_accent(r,g,b)
        self._b_embed.set_accent(r,g,b); self._status.set_accent(r,g,b)
        if hasattr(self,'_b_tcol'):  self._b_tcol.set_accent(r,g,b)
        if hasattr(self,'_b_bcol'):  self._b_bcol.set_accent(r,g,b)
        if AUDIO_OK and hasattr(self,'_b_audio'): self._b_audio.set_accent(r,g,b)
        if hasattr(self,'_b_taskbar_test'): self._b_taskbar_test.set_accent(r,g,b)
        if hasattr(self,'_pal'):  self._pal.update_colors(self._ext.get_colors())
        if hasattr(self,'_tray'): self._tray.setIcon(QIcon(make_tray_icon_pixmap(r,g,b)))

    @pyqtSlot(object)
    def _on_frame(self,frame):
        colors=self._ext.extract(frame); self._pal.update_colors(colors)
        if hasattr(self,'_chk_vauto') and self._chk_vauto.isChecked(): self._ov.set_bar_col(*self._ext.get_glow())

    @pyqtSlot(list)
    def _on_bars(self,bars): self._ov.set_bars(bars)

    def _refresh_theme(self):
        if not hasattr(self,'_chk_adapt') or not self._chk_adapt.isChecked() or not self._playing: return
        if not self._palette_dynamic: return
        acc = self._ext.get_accent()
        dr,dg,db = abs(acc[0]-self._acc[0]), abs(acc[1]-self._acc[1]), abs(acc[2]-self._acc[2])
        if dr+dg+db < 30: return
        self._apply_theme(*acc)
        threading.Thread(target=self._themer.apply, args=acc, daemon=True).start()

    def _check_desktop(self):
        if not self._chk_autopause.isChecked() or not self._playing or self._upause: return
        hwnd=ctypes.windll.user32.GetForegroundWindow()
        if not hwnd: return
        rc=ctypes.wintypes.RECT(); ctypes.windll.user32.GetWindowRect(hwnd,ctypes.byref(rc))
        sw,sh=get_screen_size()
        covers=(rc.left<=0 and rc.top<=0 and rc.right>=sw and rc.bottom>=sh)
        shell=get_foreground_class() in ("Progman","WorkerW","")
        playing=self._wp.is_playing()
        if covers and not shell and playing:          self._wp.pause(); self._tm_pause.setText("Resume")
        elif (not covers or shell) and not playing:   self._wp.resume(); self._tm_pause.setText("Pause")

    def _load(self,path):
        if not os.path.exists(path): self._status.err("File not found"); return
        if self._audio_on:
            self._audio.stop(); self._audio_on = False
            if hasattr(self,'_b_audio'): self._b_audio.setText("▶  Start Audio Capture")
        self._embedded = False
        self._vpath=path; self._vid_lbl.setText(f"📹  {Path(path).name}")
        self._status.info("Loading..."); QApplication.processEvents()
        ok = self._wp.load(path)
        self._on_load_done(ok)

    def _on_load_done(self, ok):
        if ok:
            self._playing=True; self._upause=False
            for b in (self._b_play,self._b_pause,self._b_stop): b.setEnabled(True)
            self._b_embed.setEnabled(True); self._status.ok("Loaded"); self._b_play.set_active(True)
            QTimer.singleShot(500,  self._toggle_embed)
            QTimer.singleShot(1500, self._auto_start_audio)
            QTimer.singleShot(5000, self._force_apply_theme)
        else:
            self._status.err("Could not open video — is libmpv-2.dll present?")

    def _force_apply_theme(self):
        if not self._palette_dynamic: return
        if not hasattr(self, '_chk_adapt') or not self._chk_adapt.isChecked(): return
        acc = self._ext.get_accent()
        self._apply_theme(*acc)
        threading.Thread(target=self._themer.apply, args=acc, daemon=True).start()

    def _play(self):
        self._upause=False; self._wp.resume(); self._status.playing("Playing")
        self._b_play.set_active(True); self._b_pause.set_active(False); self._tm_pause.setText("Pause")

    def _pause(self):
        self._upause=True; self._wp.pause(); self._status.paused("Paused")
        self._b_play.set_active(False); self._b_pause.set_active(True); self._tm_pause.setText("Resume")

    def _stop(self):
        self._playing=False; self._upause=False; self._wp.stop(); self._status.info("Stopped")
        for b in (self._b_play,self._b_pause): b.set_active(False)
        if self._embedded: self._toggle_embed()

    def _do_reembed(self):
        if not self._embedded or not self._vpath: return
        sw,sh=get_screen_size(); hw=self._wp.winId()
        if not hw: return
        ok=self._emb.embed(hw,sw,sh)
        if ok: self._wp.set_embedded(True,self._emb.get_target()); self._status.ok("Re-embedded")
        else:  self._status.err("Re-embed failed")

    def _toggle_embed(self):
        if not self._vpath: return
        if not self._embedded:
            hw = self._wp.winId()
            if hw:
                self._wp.set_embedded(True, getattr(self._wp, '_embedded_target', 0))
                self._ov.show(); self._ov.raise_(); self._ov.update()
                self._embedded = True
                self._b_embed.setText("✕  Detach Wallpaper"); self._b_embed.set_accent(*self._acc)
                self._status.ok("Wallpaper active!")
            else:
                self._status.err("No window handle — reload the video")
        else:
            self._wp.set_embedded(False)
            hw = self._wp.winId()
            if hw:
                ctypes.windll.user32.SetParent(hw, 0)
                ctypes.windll.user32.ShowWindow(hw, 0)
            self._ov.hide(); self._embedded = False
            self._b_embed.setText("🖥  Set as Live Wallpaper"); self._b_embed.set_accent(*self._acc)
            self._status.info("Detached")

    def _open_taskbar_settings(self):
        import subprocess
        subprocess.Popen(['explorer.exe', 'ms-settings:colors'])
        QMessageBox.information(self, "One-time Setup",
            "Windows Settings has opened to the Colours page.\n\n"
            "Turn ON:\n"
            "  • 'Show accent colour on title bars and window borders'\n"
            "  • 'Show accent colour on Start and taskbar'\n\n"
            "After enabling those, the taskbar will follow the video colour automatically.\n"
            "You only need to do this once.")

    def _on_palette_mode(self, btn):
        dynamic = (self._pal_grp.id(btn) == 0)
        self._palette_dynamic = dynamic; self._pal.set_dynamic(dynamic)
        self._b_statcol.setEnabled(not dynamic)
        if dynamic: self._pal.update_colors(self._ext.get_colors())

    def _pick_static_col(self):
        c = QColorDialog.getColor(QColor(*self._acc), self, "Static Palette Colour")
        if not c.isValid(): return
        r, g, b = c.red(), c.green(), c.blue()
        static_colors = [(r,g,b),(min(255,r+40),min(255,g+40),min(255,b+40)),
                         (max(0,r-40),max(0,g-40),max(0,b-40)),(min(255,r+80),min(255,g+15),min(255,b+15)),
                         (min(255,r+15),min(255,g+80),min(255,b+15)),(min(255,r+15),min(255,g+15),min(255,b+80))]
        self._pal.set_static_colors(static_colors)
        if self._chk_adapt.isChecked():
            self._apply_theme(r, g, b)
            threading.Thread(target=self._themer.apply, args=(r,g,b), daemon=True).start()

    def _pick_text(self):
        c=QColorDialog.getColor(self._ov.text_color,self,"Text Color",QColorDialog.ShowAlphaChannel)
        if c.isValid(): self._ov.set_text_col(c)

    def _pick_bar(self):
        c=QColorDialog.getColor(self._ov._bar_col,self,"Bar Color",QColorDialog.ShowAlphaChannel)
        if c.isValid(): self._ov.set_bar_col(c.red(),c.green(),c.blue())

    def _on_opacity(self,v):
        c=self._ov.text_color; c.setAlpha(v); self._ov.text_color=c; self._ov.update()

    def _auto_start_audio(self):
        if not AUDIO_OK or self._audio_on: return
        if self._audio.can_capture():
            self._toggle_audio(); self._status.ok("Audio visualizer started (WASAPI loopback)")
        else:
            print("[Audio] No loopback source available — visualizer inactive")

    def _on_viz_style(self, idx):
        styles = ["bars","slim","mirror","wave","dots","circle"]
        self._ov.set_viz_style(styles[idx] if idx < len(styles) else "bars")

    def _on_viz_detach(self, state):
        self._ov.viz_detached = bool(state); self._ov.update()

    def _toggle_audio(self):
        if not self._audio_on:
            self._audio.start(); self._audio_on=True
            if hasattr(self,'_b_audio'): self._b_audio.setText("⏹  Stop Capture")
        else:
            self._audio.stop(); self._audio_on=False
            if hasattr(self,'_b_audio'): self._b_audio.setText("▶  Start Audio Capture")
            self._ov.set_bars([0.0]*48)

# ══════════════════════════════════════════════
#  LAUNCH HELPERS
# ══════════════════════════════════════════════

def _detach_from_console():
    if getattr(sys, "frozen", False): return
    if ctypes.windll.kernel32.GetConsoleWindow() == 0: return
    if 'pythonw' in sys.executable.lower(): return
    pythonw = sys.executable.replace('python.exe', 'pythonw.exe')
    if not os.path.exists(pythonw):
        pythonw = os.path.join(os.path.dirname(sys.executable), 'pythonw.exe')
    if not os.path.exists(pythonw): pythonw = sys.executable
    import subprocess
    DETACHED_PROCESS=0x00000008; CREATE_NEW_PROC_GROUP=0x00000200
    subprocess.Popen([pythonw] + sys.argv,
        creationflags=DETACHED_PROCESS|CREATE_NEW_PROC_GROUP, close_fds=True,
        cwd=os.getcwd(), stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    sys.exit(0)

def _create_desktop_shortcut():
    marker = os.path.join(_script_dir, '.shortcut_created')
    if os.path.exists(marker): return
    try:
        desktop = os.path.join(os.path.expanduser('~'), 'Desktop')
        try:
            import winreg as _wr
            with _wr.OpenKey(_wr.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders") as k:
                desktop = _wr.QueryValueEx(k, "Desktop")[0]
        except Exception: pass
        shortcut_path = os.path.join(desktop, 'Theme Engine.lnk')
        pythonw = sys.executable.replace('python.exe', 'pythonw.exe')
        if not os.path.exists(pythonw): pythonw = sys.executable
        script_path = os.path.abspath(__file__)
        ps = (f'$s=(New-Object -ComObject WScript.Shell).CreateShortcut("{shortcut_path}");'
              f'$s.TargetPath="{pythonw}";$s.Arguments=\'"{script_path}"\';'
              f'$s.WorkingDirectory="{_script_dir}";$s.IconLocation="{script_path},0";'
              f'$s.Description="Live Theme Engine";$s.Save()')
        import subprocess
        subprocess.run(['powershell','-WindowStyle','Hidden','-NoProfile','-Command',ps],
                       creationflags=0x08000000, timeout=10)
        with open(marker, 'w') as f: f.write('1')
        print(f"[Setup] Desktop shortcut created: {shortcut_path}")
    except Exception as e:
        print(f"[Setup] Shortcut creation failed: {e}")

def _first_run_taskbar_prompt():
    # Skip in headless mode — no dialogs
    if _HEADLESS: return
    marker = os.path.join(_script_dir, '.taskbar_setup_shown')
    if os.path.exists(marker): return
    try:
        with open(marker, 'w') as f: f.write('1')
    except Exception: pass
    QTimer.singleShot(2000, lambda: _show_taskbar_first_run())

def _show_taskbar_first_run():
    import subprocess
    reply = QMessageBox.question(None, "Enable Taskbar Colour (one-time)",
        "For the taskbar to change colour with the wallpaper, Windows needs\n"
        "one setting enabled manually (required by Windows 11).\n\n"
        "Click Yes to open Settings now.\n\n"
        "In Settings → Personalisation → Colours, turn ON:\n"
        "  • Show accent colour on title bars\n"
        "  • Show accent colour on Start and taskbar\n\n"
        "You only need to do this once.",
        QMessageBox.Yes | QMessageBox.Cancel)
    if reply == QMessageBox.Yes:
        subprocess.Popen(['explorer.exe', 'ms-settings:colors'])


def main():
    _detach_from_console()
    _create_desktop_shortcut()
    set_dpi_aware()
    QApplication.setQuitOnLastWindowClosed(False)
    app = QApplication(sys.argv); app.setStyle("Fusion")

    if _HEADLESS:
        # ── Headless / config-driven mode ────────────────────────────────────
        print("[Config] Starting in headless mode — no GUI window")
        engine = HeadlessEngine(_CFG)
        # Keep a reference so it isn't garbage-collected
        app._engine = engine
        sys.exit(app.exec_())
    else:
        # ── Normal GUI mode ───────────────────────────────────────────────────
        panel = Panel()
        # If a config exists but headless=False, pre-apply its overlay settings
        if _CFG is not None:
            panel._ov.apply_config(_CFG)
            # Pre-load video if specified
            vpath = _CFG.get("video_path", "")
            if vpath and os.path.exists(vpath):
                QTimer.singleShot(300, lambda: panel._load(vpath))
        panel.show()
        _first_run_taskbar_prompt()
        sys.exit(app.exec_())


if __name__ == "__main__":
    main()
