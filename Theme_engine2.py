#!/usr/bin/env python3
"""
Theme Engine v4.2
Fixes:
  1. wndproc LPARAM overflow on 64-bit Windows (was ctypes.wintypes.LPARAM=32bit)
  2. MPV invalid 'audio' option (replaced with 'ao'='null'/'wasapi')

Auto-installs all required dependencies on first run (Python packages + mpv DLLs
+ Visual C++ Redistributable).  A dark-themed progress window is shown during setup.
After installation a marker file '.deps_ok' is written next to the script so the
installer is skipped on every subsequent launch.
"""

import sys, os

# ══════════════════════════════════════════════════════════════════════════════
#  FIRST-RUN AUTO-INSTALLER  (stdlib only — runs before any 3rd-party import)
# ══════════════════════════════════════════════════════════════════════════════

def _bootstrap():
    """
    First-run installer. Runs every launch but completes in <10 ms once set up.

    What it does:
      1. Verifies libmpv-2.dll is present next to the script and loads correctly.
         (The DLL is bundled with the script — no download needed.)
      2. On first run (or if packages are missing): installs Python packages via pip.
         A dark-themed progress window is shown. This only happens once.
      3. Writes a .deps_ok marker so subsequent launches skip the pip install.

    Ships with: libmpv-2.dll  (the complete mpv media engine — only file needed)

    NOTE: When running as a compiled exe (PyInstaller frozen build), this function
    returns immediately. All packages are already bundled inside the exe — pip
    install and re-launching are not needed and must NOT happen (re-launching
    would spawn infinite copies of the exe).
    """
    from pathlib import Path
    import ctypes, os as _os, sys as _sys

    # ── Frozen exe guard (PyInstaller) ───────────────────────────────────────
    # sys.frozen is set to True by PyInstaller's bootloader.
    # When frozen: packages are bundled, DLL is extracted to the temp dir,
    # PATH is already set up. Nothing to install, nothing to re-launch.
    if getattr(_sys, "frozen", False):
        return

    script_dir = Path(__file__).parent
    marker     = script_dir / ".deps_ok"

    # ── Step 1: verify libmpv-2.dll is present and loads ────────────────────
    # python-mpv accepts any of: libmpv-2.dll, mpv-2.dll, mpv-1.dll
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
        # DLL missing or broken — show a clear error and exit.
        # The user needs to place libmpv-2.dll next to the script.
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

    # ── Step 2: fast path if pip packages already installed ─────────────────
    if marker.exists():
        return

    # ── Step 3: first run — install pip packages ────────────────────────────
    import subprocess, threading
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

    # Package checklist
    _step_vars, _step_lbls = [], []
    chk = tk.Frame(root, bg=_BG)
    chk.pack(fill="x", padx=32, pady=(0, 6))
    for _, label in PIP_PACKAGES:
        row = tk.Frame(chk, bg=_BG)
        row.pack(anchor="w", pady=1)
        sv = tk.StringVar(value="○")
        _step_vars.append(sv)
        il = tk.Label(row, textvariable=sv, font=("Segoe UI", 10),
                      width=2, bg=_BG, fg=_WAIT)
        il.pack(side="left")
        tl = tk.Label(row, text=label, font=("Segoe UI", 8),
                      bg=_BG, fg=_SUB, anchor="w")
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
    tk.Label(root, textvariable=status_var,
             font=("Segoe UI", 8, "bold"), bg=_BG, fg=_FG).pack()

    detail_var = tk.StringVar(value="")
    tk.Label(root, textvariable=detail_var,
             font=("Consolas", 7), bg=_BG, fg="#333355").pack()

    err_holder = [None]

    def _ui(msg, pct, detail=""):
        def _do():
            status_var.set(msg)
            bar["value"] = pct
            if detail: detail_var.set(detail[:100])
            root.update_idletasks()
        root.after(0, _do)

    def _worker():
        try:
            _ui("Upgrading pip…", 2)
            subprocess.run(
                [_sys.executable, "-m", "pip", "install", "--upgrade", "pip", "-q"],
                capture_output=True
            )
            n = len(PIP_PACKAGES)
            for i, (pkg, _) in enumerate(PIP_PACKAGES):
                _run(i)
                pct = 5 + int((i / n) * 90)
                _ui(f"Installing {pkg}…", pct, detail=f"pip install {pkg}")
                r = subprocess.run(
                    [_sys.executable, "-m", "pip", "install", pkg, "-q"],
                    capture_output=True, text=True
                )
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

    # Re-launch so freshly installed packages are importable
    print("[Setup] Re-launching with installed packages…")
    import subprocess as _sp
    _sp.Popen([_sys.executable] + _sys.argv)
    _sys.exit(0)


_bootstrap()   # Must be called before any third-party import

# ══════════════════════════════════════════════════════════════════════════════
#  END OF AUTO-INSTALLER — normal application code begins below
# ══════════════════════════════════════════════════════════════════════════════

import sys, os, ctypes, ctypes.wintypes, threading, time, winreg, tempfile
from pathlib import Path

# When frozen by PyInstaller, __file__ points into the temp extraction dir
# (sys._MEIPASS).  User-visible files (shortcuts, markers) must go next to
# the exe — use sys.executable for that.  DLLs/resources are in _MEIPASS.
if getattr(sys, "frozen", False):
    _exe_dir    = os.path.dirname(sys.executable)          # next to ThemeEngine.exe
    _meipass    = sys._MEIPASS                              # bundled resources / DLLs
    _script_dir = _exe_dir                                  # user-facing files go here
    # Make sure both dirs are on PATH so libmpv-2.dll is found by ctypes
    for _d in (_exe_dir, _meipass):
        if _d not in os.environ["PATH"]:
            os.environ["PATH"] = _d + os.pathsep + os.environ["PATH"]
    # Write crash log next to the exe so silent errors become visible
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

# Audio capture backend — prefer pyaudiowpatch (true WASAPI loopback, zero
# permissions), fall back to sounddevice (may need Stereo Mix on some systems).
AUDIO_BACKEND = None   # "wpatch" | "sounddevice" | None
sd = None

try:
    import pyaudiowpatch as _pawp
    # Quick sanity-check: confirm WASAPI host API is available
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
    print("[mpv] Download libmpv-2.dll from https://github.com/shinchiro/mpv-winbuild-cmake/releases")
    print("[mpv] Place it next to this script, then install: https://aka.ms/vs/17/release/vc_redist.x64.exe")

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
            # Use 64x64 for better colour coverage
            small = cv2.resize(frame, (64, 64))
            rgb   = cv2.cvtColor(small, cv2.COLOR_BGR2RGB)
            px    = np.ascontiguousarray(rgb.reshape(-1, 3), dtype=np.float32)

            # k-means: more attempts and iterations to avoid grey local minima
            crit = (cv2.TERM_CRITERIA_EPS | cv2.TERM_CRITERIA_MAX_ITER, 12, 1.0)
            _, labels, centers = cv2.kmeans(
                px, 6, None, crit, 4, cv2.KMEANS_PP_CENTERS)

            counts = np.bincount(labels.flatten(), minlength=6)

            # Score each cluster by saturation * sqrt(coverage)
            # This ensures vivid colours beat large grey areas
            scores = []
            for i, (rc, gc, bc) in enumerate(centers):
                mx, mn = max(rc,gc,bc)/255.0, min(rc,gc,bc)/255.0
                lum    = (mx + mn) / 2.0
                sat    = (mx - mn) / (1.0 - abs(2*lum - 1) + 1e-9) if mx > 0 else 0.0
                cov    = counts[i] / float(len(labels))
                # Penalise very dark (<5% L) and very light (>92% L) clusters
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
        """Return the most vivid (saturated, mid-luminance) colour from the palette,
        boosted to ensure it's never too dark or muddy for taskbar use."""
        best, bs = self.get_colors()[0], -1.0
        for r, g, b in self.get_colors():
            mx, mn = max(r,g,b)/255.0, min(r,g,b)/255.0
            lum = (mx + mn) / 2.0
            sat = (mx - mn) / (1.0 - abs(2*lum - 1) + 1e-9) if mx > mn else 0.0
            if 0.1 < lum < 0.9 and sat > 0.10:
                sc = sat * sat * (1.0 - abs(lum - 0.50))
                if sc > bs: bs, best = sc, (r, g, b)
        # Boost saturation so the accent is vivid on the taskbar
        r, g, b = best
        hi, lo = max(r,g,b), min(r,g,b)
        if hi > lo:
            mid   = (hi + lo) / 2
            scale = min(255.0 / max(hi, 1), 1.5)
            r = int(max(0, min(255, mid + (r - mid) * scale)))
            g = int(max(0, min(255, mid + (g - mid) * scale)))
            b = int(max(0, min(255, mid + (b - mid) * scale)))
        # Lift floor: taskbar should never look nearly black
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
        """
        Generate the 32-byte AccentPalette binary value Windows 11 requires.
        It is 8 RGBA colours (BGRA byte order) from darkest to lightest shade.
        Without this, Windows ignores AccentColor on Win11 even if it's written.
        """
        # Lightness levels for each of the 8 palette slots (0=darkest, 7=lightest)
        levels = [0.10, 0.20, 0.35, 0.55, 0.70, 0.80, 0.90, 0.96]
        data = bytearray()
        for lv in levels:
            # Interpolate between black and the accent colour, then towards white
            if lv <= 0.55:
                # Dark half: black → accent
                t  = lv / 0.55
                pr = int(r * t)
                pg = int(g * t)
                pb = int(b * t)
            else:
                # Light half: accent → white
                t  = (lv - 0.55) / 0.45
                pr = int(r + (255 - r) * t)
                pg = int(g + (255 - g) * t)
                pb = int(b + (255 - b) * t)
            pr = max(0, min(255, pr))
            pg = max(0, min(255, pg))
            pb = max(0, min(255, pb))
            data += bytes([pb, pg, pr, 0xFF])   # BGRA order
        return bytes(data)

    def __init__(self):
        self._last_rgb = (-1, -1, -1)   # cache to skip no-op calls

    def apply(self, r, g, b):
        # Skip if colour hasn't changed meaningfully (avoids redundant DWM calls)
        lr, lg, lb = self._last_rgb
        if abs(r-lr) + abs(g-lg) + abs(b-lb) < 6:
            return
        self._last_rgb = (r, g, b)

        # Both registry AND DwmSetColorizationColor use ABGR (0xAABBGGRR)
        reg_abgr = ctypes.c_uint32(0xFF000000 | (b<<16) | (g<<8) | r).value
        palette  = self._make_palette(r, g, b)
        u32      = ctypes.windll.user32

        # 1. Write AccentColor + full AccentPalette (Win11 needs both)
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

        # 2. Ensure taskbar colour is enabled (write only once at startup ideally,
        #    but cheap enough to repeat — no visual side-effect)
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize",
                0, winreg.KEY_WRITE) as k:
                winreg.SetValueEx(k, "ColorPrevalence",        0, winreg.REG_DWORD, 1)
                winreg.SetValueEx(k, "TaskbarColorPrevalence", 0, winreg.REG_DWORD, 1)
        except Exception as e:
            print(f"[Theme] Personalize: {e}")

        # 3. DWM live title-bar colour — updates own app title bars instantly,
        #    no broadcast needed, no other apps are disturbed.
        try:
            dwmapi = ctypes.windll.dwmapi
            dwmapi.DwmSetColorizationColor.restype  = ctypes.c_long
            dwmapi.DwmSetColorizationColor.argtypes = [ctypes.c_uint32, ctypes.c_bool]
            dwmapi.DwmSetColorizationColor(reg_abgr, False)
        except Exception as e:
            print(f"[Theme] DWM: {e}")

        # 4. Poke ONLY the taskbar shell windows directly.
        #    We deliberately do NOT broadcast WM_SETTINGCHANGE or WM_THEMECHANGED
        #    to 0xFFFF (all windows) — that forces every running application to
        #    repaint its chrome simultaneously, which causes the visible glitching.
        #    Poking Shell_TrayWnd + SecondaryTaskbar children is enough for the
        #    taskbar to pick up the new accent colour without disturbing other apps.
        try:
            WM_SETTINGCHANGE = 0x001A
            WM_THEMECHANGED  = 0x031A
            buf = ctypes.create_unicode_buffer("ImmersiveColorSet")
            res = ctypes.wintypes.DWORD(0)

            # Primary taskbar
            tray = u32.FindWindowW("Shell_TrayWnd", None)
            if tray:
                u32.SendMessageTimeoutW(tray, WM_SETTINGCHANGE, 0, buf,
                                        0x0002, 500, ctypes.byref(res))
                u32.SendMessageTimeoutW(tray, WM_THEMECHANGED,  0, 0,
                                        0x0002, 500, None)

            # Secondary taskbars (multi-monitor)
            secondary = u32.FindWindowW("Shell_SecondaryTrayWnd", None)
            while secondary:
                u32.SendMessageTimeoutW(secondary, WM_SETTINGCHANGE, 0, buf,
                                        0x0002, 500, ctypes.byref(res))
                u32.SendMessageTimeoutW(secondary, WM_THEMECHANGED,  0, 0,
                                        0x0002, 500, None)
                secondary = u32.FindWindowExW(None, secondary,
                                              "Shell_SecondaryTrayWnd", None)
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
        """
        Send the magic Progman message that makes Explorer split WorkerW into
        two layers: one hosting SHELLDLL_DefView (icons) and one below for the
        wallpaper.  Params MUST be 0xD / 0x1 — using 0 / 0 causes Explorer on
        Win10/11 to partially handle the message without actually creating the
        second layer.  We fire both variants for maximum compatibility.
        """
        u32 = ctypes.windll.user32
        res = ctypes.wintypes.DWORD(0)
        u32.SendMessageTimeoutW(pm, 0x052C, 0xD, 0x1, 0x0002, 2000, ctypes.byref(res))
        time.sleep(0.25)
        u32.SendMessageTimeoutW(pm, 0x052C, 0,   0,   0x0002, 2000, ctypes.byref(res))
        time.sleep(0.25)

    def _find_target(self):
        """
        Replicate Steam Wallpaper Engine's targeting logic:
          1. Send 0x052C (with 0xD/0x1 params) to Progman to create wallpaper layer
          2. Enumerate top-level windows for a WorkerW containing SHELLDLL_DefView
             (that's the icon layer)
          3. Wallpaper target = FindWindowEx(NULL, icon_workerw, "WorkerW", NULL)
             i.e. the sibling WorkerW immediately below it in Z-order
          4. If DefView is in Progman directly, use Progman as target
          5. Last resort: fall back to Progman
        """
        u32 = ctypes.windll.user32
        pm  = u32.FindWindowW("Progman", None)
        if not pm:
            print("[Embed] Progman not found"); return 0

        self._send_spawn_msg(pm)

        PROC = ctypes.WINFUNCTYPE(ctypes.wintypes.BOOL,
                                  ctypes.wintypes.HWND,
                                  ctypes.wintypes.LPARAM)
        wallpaper_ww = [0]

        def cb(hwnd, _):
            defview = u32.FindWindowExW(hwnd, None, "SHELLDLL_DefView", None)
            if defview:
                cls = self._cls(hwnd)
                print(f"[Embed] DefView found inside {cls} {hwnd:#010x}")
                if cls == "WorkerW":
                    # Sibling WorkerW below the icon layer = wallpaper slot
                    sib = u32.FindWindowExW(None, hwnd, "WorkerW", None)
                    if sib:
                        print(f"[Embed] Wallpaper WorkerW (sibling) = {sib:#010x}")
                        wallpaper_ww[0] = sib
                    else:
                        print(f"[Embed] No sibling WorkerW found, using icon WorkerW")
                        wallpaper_ww[0] = hwnd
                elif cls == "Progman":
                    print(f"[Embed] DefView in Progman — using Progman as target")
                    wallpaper_ww[0] = hwnd
            return True

        u32.EnumWindows(PROC(cb), 0)

        if wallpaper_ww[0]:
            return wallpaper_ww[0]

        # Diagnostics if still nothing found
        with_dv = []; without_dv = []
        def cb2(hwnd, _):
            if self._cls(hwnd) == "WorkerW":
                (with_dv if u32.FindWindowExW(hwnd, None, "SHELLDLL_DefView", None)
                         else without_dv).append(hwnd)
            return True
        u32.EnumWindows(PROC(cb2), 0)
        print(f"[Embed] WARN: DefView not found in any top-level window")
        print(f"[Embed] WorkerW with DefView   = {[f'{h:#x}' for h in with_dv]}")
        print(f"[Embed] WorkerW without DefView= {[f'{h:#x}' for h in without_dv]}")
        print(f"[Embed] Falling back to Progman {pm:#010x}")
        return pm

    def _prep_for_child(self, hwnd):
        """
        SetParent only works if the window has WS_CHILD style.
        WS_POPUP windows: GetParent always returns 0 even after SetParent.
        Solution: strip WS_POPUP, add WS_CHILD, then SetParent will stick.

        IMPORTANT: SetWindowLongPtrW takes LONG_PTR (signed 64-bit on x64).
        Must declare argtypes explicitly — otherwise ctypes overflows on
        values with the high bit set (e.g. 0x40000000 is fine, but
        ~0x80000000 produces a large positive int that overflows c_long).
        """
        u32 = ctypes.windll.user32

        # Declare correct argtypes for SetWindowLongPtrW / GetWindowLongPtrW
        u32.GetWindowLongPtrW.restype  = ctypes.c_ssize_t
        u32.GetWindowLongPtrW.argtypes = [ctypes.wintypes.HWND, ctypes.c_int]
        u32.SetWindowLongPtrW.restype  = ctypes.c_ssize_t
        u32.SetWindowLongPtrW.argtypes = [ctypes.wintypes.HWND, ctypes.c_int, ctypes.c_ssize_t]

        GWL_STYLE, GWL_EXSTYLE = -16, -20
        WS_POPUP      = 0x80000000
        WS_CHILD      = 0x40000000
        WS_CAPTION    = 0x00C00000
        WS_THICKFRAME = 0x00040000

        style   = u32.GetWindowLongPtrW(hwnd, GWL_STYLE)
        exstyle = u32.GetWindowLongPtrW(hwnd, GWL_EXSTYLE)

        # Mask operations in Python produce arbitrary-precision ints.
        # Cast to c_ssize_t (signed 64-bit) before passing to SetWindowLongPtrW.
        new_style   = ctypes.c_ssize_t((style & ~(WS_POPUP | WS_CAPTION | WS_THICKFRAME)) | WS_CHILD).value
        new_exstyle = ctypes.c_ssize_t(exstyle & ~0x00040000).value  # remove WS_EX_APPWINDOW

        u32.SetWindowLongPtrW(hwnd, GWL_STYLE,   new_style)
        u32.SetWindowLongPtrW(hwnd, GWL_EXSTYLE, new_exstyle)
        # FRAMECHANGED tells Windows to re-read the style
        u32.SetWindowPos(hwnd, 0, 0, 0, 0, 0, 0x0001|0x0002|0x0004|0x0020)
        print(f"[Embed] style {style:#010x} -> {new_style:#010x}")

    def _force_show(self, hwnd, sw, sh):
        u32 = ctypes.windll.user32
        parent = u32.GetParent(hwnd)
        defview = u32.FindWindowExW(parent, None, "SHELLDLL_DefView", None) if parent else None
        insert_after = defview if defview else 1  # below DefView, or HWND_BOTTOM
        u32.SetWindowPos(hwnd, insert_after, 0, 0, sw, sh, 0x0010|0x0040|0x0020)
        u32.ShowWindow(hwnd, 5)
        u32.UpdateWindow(hwnd)

    def embed(self, hwnd, sw, sh):
        print(f"[Embed] hwnd={hwnd:#010x} {sw}x{sh}")
        u32 = ctypes.windll.user32

        # Convert WS_POPUP -> WS_CHILD FIRST — otherwise GetParent always returns 0
        self._prep_for_child(hwnd)

        for attempt in range(5):
            target = self._find_target()
            if target:
                result = u32.SetParent(hwnd, target)
                err    = ctypes.GetLastError()
                self._force_show(hwnd, sw, sh)
                actual = u32.GetParent(hwnd)
                # Also try GetAncestor as a backup verification
                anc = u32.GetAncestor(hwnd, 1)  # GA_PARENT=1
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
    """
    Captures system audio output using the best available backend:

    1. pyaudiowpatch  — true WASAPI loopback (same API Wallpaper Engine uses).
       Captures whatever plays through your speakers/headphones.
       Zero Windows permissions needed, no Stereo Mix, no driver config.

    2. sounddevice    — fallback.  Uses PortAudio WASAPI loopback if available,
       otherwise tries Stereo Mix / 'What U Hear' named devices.
    """
    bars_updated = pyqtSignal(list)
    N = 48

    def __init__(self):
        super().__init__()
        self._running  = False
        self._smoothed = [0.0] * self.N
        self._peak     = [1e-4] * self.N   # per-bar long-term peak for equalisation
        self._lock     = threading.Lock()
        self._queue    = None
        self._sr       = 48000             # filled in by each backend before calling _fft_to_bars

    # ── public ────────────────────────────────────────────────────────────────
    def start(self):
        if not AUDIO_OK or self._running: return
        import queue
        self._queue   = queue.SimpleQueue()
        self._running = True
        threading.Thread(target=self._run,     daemon=True).start()
        threading.Thread(target=self._emitter, daemon=True).start()

    def stop(self):
        self._running = False

    # ── internal ──────────────────────────────────────────────────────────────
    def _emitter(self):
        """Emit bars to the UI thread. If no audio data arrives for 300 ms
        (stream idle / silence), rapidly decay bars to zero so the display
        clears instead of freezing at the last frame."""
        import time as _time
        SILENCE_TIMEOUT = 0.30   # seconds before we start zeroing
        ZERO_DECAY      = 0.72   # multiplier per 33 ms tick when silent
        last_data       = _time.monotonic()

        while self._running:
            try:
                bars = self._queue.get(timeout=0.033)
                last_data = _time.monotonic()
                self.bars_updated.emit(bars)
            except Exception:
                # Nothing arrived in 33 ms — check for silence
                elapsed = _time.monotonic() - last_data
                if elapsed > SILENCE_TIMEOUT:
                    with self._lock:
                        self._smoothed = [v * ZERO_DECAY for v in self._smoothed]
                        zeroed = list(self._smoothed)
                    # Stop emitting once all bars are effectively zero
                    if any(v > 0.002 for v in zeroed):
                        self.bars_updated.emit(zeroed)
                    else:
                        self.bars_updated.emit([0.0] * self.N)

    def _fft_to_bars(self, audio_mono):
        """
        Convert a mono float32 numpy array → smoothed, equalised bar list.

        Three fixes for the "first bar always high" problem:
          1. Frequency floor/ceiling: map bars to 40 Hz–18 kHz only.
             Bin 0 is DC (always non-zero even in silence) — excluded.
          2. Per-bar auto-equaliser: each bar tracks its own long-term peak
             and normalises against it.  All bars self-level over a few seconds
             regardless of whether bass or treble dominates the source material.
          3. Global normalise AFTER per-bar eq, so the loudest bar = 1.0 but
             relative heights are already equalised.
        """
        FFT_N   = 4096
        SR      = self._sr
        F_LO    = 40.0          # Hz — skip DC + sub-bass rumble
        F_HI    = 18000.0       # Hz — above hearing threshold / Nyquist safe
        DECAY   = 0.88          # smoothing decay  (lower = snappier fall)
        ATTACK  = 0.55          # smoothing attack (lower = snappier rise)
        EQ_RISE = 0.015         # how fast per-bar peak rises  (fast)
        EQ_FALL = 0.002         # how fast per-bar peak decays (slow → equalises gradually)

        windowed = audio_mono * np.hanning(len(audio_mono))
        fft_full = np.abs(np.fft.rfft(windowed, n=FFT_N))
        n_bins   = len(fft_full)

        # Frequency of each FFT bin
        freqs = np.fft.rfftfreq(FFT_N, d=1.0 / SR)

        # Bin indices for our frequency range
        bin_lo = max(1, int(F_LO  * FFT_N / SR))   # max(1,...) skips DC bin 0
        bin_hi = min(n_bins - 1, int(F_HI * FFT_N / SR))

        usable = fft_full[bin_lo:bin_hi]
        n_use  = len(usable)
        if n_use < self.N:
            return None

        # Log-spaced bin edges within [bin_lo, bin_hi]
        log_lo = np.log1p(0)
        log_hi = np.log1p(n_use)
        edges  = [int(np.expm1(log_lo + (log_hi - log_lo) * i / self.N))
                  for i in range(self.N + 1)]

        raw_bars = []
        for i in range(self.N):
            lo = edges[i]
            hi = max(lo + 1, edges[i + 1])
            hi = min(hi, n_use)
            raw_bars.append(float(np.mean(usable[lo:hi])))

        with self._lock:
            # Per-bar auto-equaliser: track slow-moving peak per band
            for i, v in enumerate(raw_bars):
                if v > self._peak[i]:
                    self._peak[i] = self._peak[i] * (1 - EQ_RISE) + v * EQ_RISE
                else:
                    self._peak[i] = max(1e-6, self._peak[i] * (1 - EQ_FALL))

            eq_bars = [v / self._peak[i] for i, v in enumerate(raw_bars)]

            # Global normalise so loudest bar = 1.0
            mx = max(eq_bars) if max(eq_bars) > 1e-8 else 1e-8
            eq_bars = [b / mx for b in eq_bars]

            # Temporal smoothing
            for i, (nw, old) in enumerate(zip(eq_bars, self._smoothed)):
                self._smoothed[i] = (nw * ATTACK + old * (1 - ATTACK)
                                     if nw > old else old * DECAY)
            return list(self._smoothed)

    def _process(self, raw_bytes, n_channels, sampwidth=4):
        """Convert raw float32 bytes from pyaudiowpatch → bars."""
        try:
            arr = np.frombuffer(raw_bytes, dtype=np.float32)
            if arr.size == 0:
                return None
            if n_channels > 1:
                arr = arr.reshape(-1, n_channels).mean(axis=1)
            return self._fft_to_bars(arr)
        except Exception as e:
            print(f"[Audio] _process: {e}")
            return None

    def _push(self, bars):
        """Drop oldest entry if consumer is slow, then push latest."""
        if self._queue is None:
            return
        try:    self._queue.get_nowait()
        except Exception: pass
        self._queue.put(bars)

    # ── backend: pyaudiowpatch ─────────────────────────────────────────────────
    def _run_wpatch(self):
        import pyaudiowpatch as pyaudio
        p = pyaudio.PyAudio()
        try:
            # Find the default speakers loopback device
            wasapi = p.get_host_api_info_by_type(pyaudio.paWASAPI)
            default_out = p.get_device_info_by_index(wasapi["defaultOutputDevice"])

            loopback_dev = None
            for dev in p.get_loopback_device_info_generator():
                # Match by name — loopback entries mirror the output device name
                if default_out["name"] in dev["name"]:
                    loopback_dev = dev
                    break
            # If no name match, take any loopback
            if loopback_dev is None:
                for dev in p.get_loopback_device_info_generator():
                    loopback_dev = dev
                    break

            if loopback_dev is None:
                print("[Audio] pyaudiowpatch: no loopback device found")
                return False

            ch = int(loopback_dev["maxInputChannels"])
            sr = int(loopback_dev["defaultSampleRate"])
            self._sr = sr
            chunk = 1024
            print(f"[Audio] wpatch loopback: [{loopback_dev['index']}] "
                  f"{loopback_dev['name']}  {ch}ch  {sr}Hz")

            stream = p.open(
                format=pyaudio.paFloat32,
                channels=ch,
                rate=sr,
                input=True,
                input_device_index=loopback_dev["index"],
                frames_per_buffer=chunk,
            )
            print("[Audio] WASAPI loopback stream open — no permissions needed")
            while self._running:
                try:
                    data = stream.read(chunk, exception_on_overflow=False)
                    bars = self._process(data, ch)
                    if bars:
                        self._push(bars)
                except Exception as e:
                    print(f"[Audio] wpatch read: {e}")
                    break
            stream.stop_stream()
            stream.close()
            return True
        except Exception as e:
            print(f"[Audio] wpatch error: {e}")
            return False
        finally:
            p.terminate()

    # ── backend: sounddevice (fallback) ────────────────────────────────────────
    def _find_loopback_sd(self):
        """Return (device_idx, channels, use_wasapi_loopback) for sounddevice."""
        try:
            devices  = sd.query_devices()
            hostapis = sd.query_hostapis()
            wasapi_idx = next((i for i, a in enumerate(hostapis)
                               if 'wasapi' in a['name'].lower()), None)
            if wasapi_idx is not None:
                out_idx = hostapis[wasapi_idx].get('default_output_device', -1)
                if out_idx < 0:
                    out_idx = next((i for i, d in enumerate(devices)
                                    if d['hostapi'] == wasapi_idx
                                    and d['max_output_channels'] > 0), -1)
                if out_idx >= 0:
                    d  = devices[out_idx]
                    ch = min(max(d.get('max_output_channels', 2), 1), 2)
                    sr = int(d.get('default_samplerate', 48000))
                    print(f"[Audio] sd WASAPI loopback → [{out_idx}] {d['name']} {ch}ch {sr}Hz")
                    return out_idx, ch, True
            # Named Stereo Mix fallback
            for i, d in enumerate(devices):
                nm = d['name'].lower()
                if d['max_input_channels'] > 0 and any(
                        k in nm for k in ('stereo mix','loopback','what u hear','wave out mix')):
                    print(f"[Audio] sd Stereo Mix → [{i}] {d['name']}")
                    return i, min(d['max_input_channels'], 2), False
        except Exception as e:
            print(f"[Audio] sd device query: {e}")
        return None, 1, False

    def _run_sounddevice(self):
        dev, nch, wasapi_lb = self._find_loopback_sd()
        if dev is None:
            print("[Audio] sounddevice: no capture device found")
            return False
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
                if bars:
                    self._push(bars)
            except Exception as e:
                print(f"[Audio] sd cb: {e}")

        kw = dict(device=dev, channels=nch, samplerate=sr, blocksize=2048, callback=cb)
        if wasapi_lb:
            try:
                kw['extra_settings'] = sd.WasapiSettings(loopback=True)
            except Exception:
                pass
        try:
            with sd.InputStream(**kw):
                print(f"[Audio] sounddevice stream running")
                while self._running: time.sleep(0.05)
            return True
        except Exception as e:
            print(f"[Audio] sounddevice stream error: {e}")
            return False

    # ── dispatcher ────────────────────────────────────────────────────────────
    def _run(self):
        if AUDIO_BACKEND == "wpatch":
            ok = self._run_wpatch()
            if not ok and sd is not None:
                print("[Audio] Falling back to sounddevice")
                self._run_sounddevice()
        elif AUDIO_BACKEND == "sounddevice":
            self._run_sounddevice()
        # idle until stopped
        while self._running: time.sleep(0.1)

    # ── probe (used by auto-start logic) ──────────────────────────────────────
    def can_capture(self):
        """Return True if a loopback source is available without user config."""
        if AUDIO_BACKEND == "wpatch":
            import pyaudiowpatch as pyaudio
            p = pyaudio.PyAudio()
            try:
                wasapi = p.get_host_api_info_by_type(pyaudio.paWASAPI)
                for _ in p.get_loopback_device_info_generator():
                    return True
                return False
            except Exception:
                return False
            finally:
                p.terminate()
        elif AUDIO_BACKEND == "sounddevice":
            dev, _, _ = self._find_loopback_sd()
            return dev is not None
        return False

# ══════════════════════════════════════════════
#  OVERLAY
# ══════════════════════════════════════════════

class Overlay(QWidget):
    """
    Transparent desktop overlay (WindowStaysOnBottomHint) that draws:
      • Clock / date / day  (draggable)
      • Audio visualizer    (6 styles, optionally detached from clock)
      • Sparkle trail       (follows mouse cursor)

    Viz styles:  bars | slim | mirror | wave | dots | circle
    """

    VIZ_STYLES = ["bars", "slim", "mirror", "wave", "dots", "circle"]

    def __init__(self, sw, sh):
        super().__init__()
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.Tool | Qt.WindowStaysOnBottomHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setAttribute(Qt.WA_ShowWithoutActivating)
        self.setGeometry(0, 0, sw, sh)
        self._sw, self._sh = sw, sh

        # ── Clock
        self.font_family = "Segoe UI Light"
        self.text_color  = QColor(255, 255, 255, 220)
        self.show_secs   = True
        self.show_clock  = True
        self.cx          = sw // 2
        self.cy          = int(sh * 0.52)

        # ── Visualizer
        self._bars       = [0.0] * 48
        self._bar_col    = QColor(120, 210, 255, 210)
        self.show_viz    = True
        self.viz_style   = "bars"        # one of VIZ_STYLES
        self.viz_detached = False         # when True, viz has its own anchor
        self.vx          = sw // 2       # viz x when detached
        self.vy          = int(sh * 0.80)# viz y when detached (near bottom)

        # ── Drag state (tracks which widget is being dragged)
        self._drag        = None          # "clock" | "viz" | None
        self._drag_origin = None

        # ── Sparkle particles  [{x,y,vx,vy,life,maxlife,r,g,b}, ...]
        self._sparkles         = []
        self._last_mouse       = None
        self._sparkles_enabled = True

        # ── Timers
        QTimer(self, timeout=self.update,          interval=33).start()   # ~30 fps
        QTimer(self, timeout=self._poll_cursor,    interval=33).start()   # sparkle feed
        QTimer(self, timeout=self._age_sparkles,   interval=33).start()   # particle physics

    # ── public setters ────────────────────────────────────────────────────────
    def set_bars(self, b):         self._bars = b;         self.update()
    def set_bar_col(self, r,g,b):  self._bar_col = QColor(r,g,b,210); self.update()
    def set_text_col(self, c):     self.text_color = c;   self.update()
    def set_viz_style(self, s):    self.viz_style = s if s in self.VIZ_STYLES else "bars"; self.update()

    # ── sparkle machinery ─────────────────────────────────────────────────────
    def _poll_cursor(self):
        """Called every 33 ms — spawn sparkles at cursor even if no mouse event fires
        (happens when another app is in front but desktop is partially visible)."""
        if not getattr(self, '_sparkles_enabled', True):
            return
        from PyQt5.QtGui import QCursor
        gp = QCursor.pos()
        p  = self.mapFromGlobal(gp)
        if self._last_mouse and (abs(p.x()-self._last_mouse.x()) > 2 or
                                  abs(p.y()-self._last_mouse.y()) > 2):
            self._spawn_sparkles(p.x(), p.y(), count=3)
        self._last_mouse = p

    def _spawn_sparkles(self, x, y, count=6):
        import random, math
        bc = self._bar_col
        for _ in range(count):
            angle = random.uniform(0, math.tau)
            speed = random.uniform(0.5, 2.8)
            life  = random.randint(18, 42)
            # colour: bar colour ± small variation
            dr = random.randint(-30, 30)
            dg = random.randint(-30, 30)
            db = random.randint(-30, 30)
            self._sparkles.append({
                "x": x, "y": y,
                "vx": math.cos(angle) * speed,
                "vy": math.sin(angle) * speed - random.uniform(0.3, 1.2),
                "life": life, "maxlife": life,
                "r": max(0,min(255, bc.red()   + dr)),
                "g": max(0,min(255, bc.green() + dg)),
                "b": max(0,min(255, bc.blue()  + db)),
            })
        # cap total particles
        if len(self._sparkles) > 300:
            self._sparkles = self._sparkles[-300:]

    def _age_sparkles(self):
        alive = []
        for s in self._sparkles:
            s["x"]    += s["vx"]
            s["y"]    += s["vy"]
            s["vy"]   += 0.08      # gravity
            s["life"] -= 1
            if s["life"] > 0:
                alive.append(s)
        self._sparkles = alive

    # ── paint ─────────────────────────────────────────────────────────────────
    def paintEvent(self, _):
        p = QPainter(self)
        p.setRenderHints(QPainter.Antialiasing | QPainter.TextAntialiasing)
        self._draw_sparkles(p)
        if self.show_clock:             self._draw_clock(p)
        if self.show_viz:               self._draw_viz(p)
        p.end()

    def _draw_sparkles(self, p):
        if not self._sparkles:
            return
        p.setPen(Qt.NoPen)
        for s in self._sparkles:
            t    = s["life"] / s["maxlife"]          # 1→0 as particle ages
            size = max(1.0, t * 5.0)
            alpha = int(t * 220)
            # star shape: small cross
            c = QColor(s["r"], s["g"], s["b"], alpha)
            p.setBrush(QBrush(c))
            cx, cy = s["x"], s["y"]
            # draw a tiny rotated diamond
            p.drawEllipse(QPointF(cx, cy), size * 0.9, size * 0.9)
            # cross arms (sparkle twinkle)
            if t > 0.5:
                arm = size * 1.6
                pc  = QColor(255, 255, 255, int(t * 160))
                p.setBrush(QBrush(pc))
                p.drawEllipse(QPointF(cx, cy - arm * 0.5), size * 0.25, size * 0.6)
                p.drawEllipse(QPointF(cx, cy + arm * 0.5), size * 0.25, size * 0.6)
                p.drawEllipse(QPointF(cx - arm * 0.5, cy), size * 0.6, size * 0.25)
                p.drawEllipse(QPointF(cx + arm * 0.5, cy), size * 0.6, size * 0.25)

    def _draw_clock(self, p):
        now  = QDateTime.currentDateTime()
        day  = now.toString("dddd").upper()
        date = now.toString("d MMM yyyy")
        tstr = now.toString("HH:mm:ss") if self.show_secs else now.toString("HH:mm")
        cx, cy = self.cx, self.cy

        def draw_text(text, pt, y_off, tracking=6):
            f = QFont(self.font_family, pt, QFont.Light)
            f.setLetterSpacing(QFont.AbsoluteSpacing, tracking)
            p.setFont(f)
            rect = QRect(cx - 700, cy + y_off - pt, 1400, pt * 2 + 10)
            p.setPen(QColor(0, 0, 0, 120))
            p.drawText(rect.adjusted(2, 3, 2, 3), Qt.AlignHCenter | Qt.AlignVCenter, text)
            p.setPen(self.text_color)
            p.drawText(rect, Qt.AlignHCenter | Qt.AlignVCenter, text)

        draw_text(day, 58, -55, 18)
        draw_text(date, 19, 20, 4)
        draw_text(tstr, 30, 68, 6)

    # ── viz dispatcher ────────────────────────────────────────────────────────
    def _viz_anchor(self):
        """Return (cx, cy) for the visualizer — either clock anchor or own anchor."""
        if self.viz_detached:
            return self.vx, self.vy
        # Attached: sit above the clock block
        return self.cx, self.cy

    def _draw_viz(self, p):
        if not any(b > 0.001 for b in self._bars):
            return
        style = self.viz_style
        if   style == "bars":   self._viz_bars(p, thin=False)
        elif style == "slim":   self._viz_bars(p, thin=True)
        elif style == "mirror": self._viz_mirror(p)
        elif style == "wave":   self._viz_wave(p)
        elif style == "dots":   self._viz_dots(p)
        elif style == "circle": self._viz_circle(p)

    def _bar_geometry(self, thin=False):
        """Return (x0, y_base, bar_w, gap, max_h) for linear styles."""
        sw, sh  = self._sw, self._sh
        cx, cy  = self._viz_anchor()
        N       = len(self._bars)
        BAR_W   = max(3 if thin else 6, int(sw * (0.004 if thin else 0.008)))
        GAP     = max(2 if thin else 2,  int(sw * (0.004 if thin else 0.003)))
        MAX_H   = int(sh * 0.16)

        if self.viz_detached:
            # detached: bars grow upward from vy
            y_base = cy
        else:
            # attached: above clock day label top ≈ cy - 84, plus gap
            y_base = max(MAX_H + 8, cy - 84 - 18)

        total_w = N * (BAR_W + GAP) - GAP
        x0      = cx - total_w // 2
        return x0, y_base, BAR_W, GAP, MAX_H

    def _viz_bars(self, p, thin=False):
        x0, y_base, BAR_W, GAP, MAX_H = self._bar_geometry(thin)
        for i, val in enumerate(self._bars):
            h  = max(2, int(val * MAX_H))
            x  = x0 + i * (BAR_W + GAP)
            y  = y_base - h
            gc = QColor(self._bar_col); gc.setAlpha(30)
            p.setBrush(QBrush(gc)); p.setPen(Qt.NoPen)
            p.drawRoundedRect(x - 2, y - 3, BAR_W + 4, h + 6, 3, 3)
            grad = QLinearGradient(x, y + h, x, y)
            bc   = QColor(self._bar_col)
            tc   = QColor(min(255, bc.red()+80), min(255,bc.green()+80), min(255,bc.blue()+80), 250)
            grad.setColorAt(0.0, bc); grad.setColorAt(1.0, tc)
            p.setBrush(QBrush(grad))
            r = 1 if thin else 2
            p.drawRoundedRect(x, y, BAR_W, h, r, r)
            hl = QColor(255, 255, 255, 55)
            p.setBrush(QBrush(hl))
            p.drawRoundedRect(x + 1, y, BAR_W - 2, min(3, h), 1, 1)

    def _viz_mirror(self, p):
        """Bars grow both upward and downward from a centre line."""
        x0, y_base, BAR_W, GAP, MAX_H = self._bar_geometry()
        half = MAX_H // 2
        for i, val in enumerate(self._bars):
            h  = max(1, int(val * half))
            x  = x0 + i * (BAR_W + GAP)
            bc = QColor(self._bar_col)
            tc = QColor(min(255,bc.red()+80),min(255,bc.green()+80),min(255,bc.blue()+80),250)
            p.setPen(Qt.NoPen)
            # upward
            grad = QLinearGradient(x, y_base, x, y_base - h)
            grad.setColorAt(0.0, bc); grad.setColorAt(1.0, tc)
            p.setBrush(QBrush(grad))
            p.drawRoundedRect(x, y_base - h, BAR_W, h, 2, 2)
            # downward (mirror)
            grad2 = QLinearGradient(x, y_base, x, y_base + h)
            grad2.setColorAt(0.0, bc); grad2.setColorAt(1.0, tc)
            p.setBrush(QBrush(grad2))
            p.drawRoundedRect(x, y_base, BAR_W, h, 2, 2)

    def _viz_wave(self, p):
        """Smooth cubic spline through bar tops — Wallpaper Engine wave style."""
        from PyQt5.QtGui import QPainterPath
        x0, y_base, BAR_W, GAP, MAX_H = self._bar_geometry()
        N   = len(self._bars)
        pts = []
        for i, val in enumerate(self._bars):
            h = int(val * MAX_H)
            x = x0 + i * (BAR_W + GAP) + BAR_W // 2
            pts.append((x, y_base - h))

        if len(pts) < 2:
            return

        # Build smooth path via cubic bezier control points
        path = QPainterPath()
        path.moveTo(pts[0][0], pts[0][1])
        for i in range(1, len(pts)):
            x0p, y0p = pts[i-1]
            x1p, y1p = pts[i]
            cx1 = x0p + (x1p - x0p) * 0.5
            path.cubicTo(cx1, y0p, cx1, y1p, x1p, y1p)

        # Filled wave under the line
        fill_path = QPainterPath(path)
        last_x = pts[-1][0]
        fill_path.lineTo(last_x,  y_base)
        fill_path.lineTo(pts[0][0], y_base)
        fill_path.closeSubpath()

        bc   = QColor(self._bar_col); bc.setAlpha(60)
        p.setBrush(QBrush(bc)); p.setPen(Qt.NoPen)
        p.drawPath(fill_path)

        # Stroke the line
        pen_col = QColor(self._bar_col); pen_col.setAlpha(220)
        from PyQt5.QtGui import QPen
        p.setPen(QPen(pen_col, 2.5, Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin))
        p.setBrush(Qt.NoBrush)
        p.drawPath(path)

        # Bright dot at each peak
        p.setPen(Qt.NoPen)
        dc = QColor(255, 255, 255, 180)
        p.setBrush(QBrush(dc))
        for x, y in pts:
            if self._bars[pts.index((x,y))] > 0.3:
                p.drawEllipse(QPointF(x, y), 2.5, 2.5)

    def _viz_dots(self, p):
        """Column of stacked glowing dots per band."""
        x0, y_base, BAR_W, GAP, MAX_H = self._bar_geometry()
        DOT   = max(4, BAR_W)
        STEP  = DOT + 2
        for i, val in enumerate(self._bars):
            n_dots = max(1, int(val * MAX_H / STEP))
            cx_dot = x0 + i * (BAR_W + GAP) + BAR_W // 2
            for d in range(n_dots):
                t     = d / max(n_dots, 1)
                y_dot = y_base - d * STEP - DOT // 2
                alpha = int(180 + t * 60)
                bc    = QColor(self._bar_col)
                r     = min(255, bc.red()   + int(t * 80))
                g     = min(255, bc.green() + int(t * 80))
                b     = min(255, bc.blue()  + int(t * 80))
                col   = QColor(r, g, b, alpha)
                # outer glow
                gc    = QColor(r, g, b, 40)
                p.setBrush(QBrush(gc)); p.setPen(Qt.NoPen)
                p.drawEllipse(QPointF(cx_dot, y_dot), DOT * 0.9, DOT * 0.9)
                # dot
                p.setBrush(QBrush(col))
                p.drawEllipse(QPointF(cx_dot, y_dot), DOT * 0.55, DOT * 0.55)

    def _viz_circle(self, p):
        """Radial spectrum — bars radiate outward from a centre point."""
        import math
        from PyQt5.QtGui import QPen
        cx, cy = self._viz_anchor()
        N      = len(self._bars)
        sh     = self._sh
        R_MIN  = int(sh * 0.06)
        R_MAX  = int(sh * 0.18)
        bc_qt  = QColor(self._bar_col)

        if not self.viz_detached:
            # Shift up so circle sits above the clock
            cy = cy - 84 - 18 - R_MAX - 10

        for i, val in enumerate(self._bars):
            angle  = math.tau * i / N - math.pi / 2   # start at top
            r_out  = R_MIN + int(val * (R_MAX - R_MIN))
            x1     = cx + math.cos(angle) * R_MIN
            y1     = cy + math.sin(angle) * R_MIN
            x2     = cx + math.cos(angle) * r_out
            y2     = cy + math.sin(angle) * r_out
            t      = val
            r      = min(255, bc_qt.red()   + int(t * 80))
            g      = min(255, bc_qt.green() + int(t * 80))
            b      = min(255, bc_qt.blue()  + int(t * 80))
            alpha  = int(140 + t * 100)
            lw     = max(2.0, 3.0 * val + 1.0)
            pen    = QPen(QColor(r, g, b, alpha), lw, Qt.SolidLine, Qt.RoundCap)
            p.setPen(pen)
            p.drawLine(QPointF(x1, y1), QPointF(x2, y2))

        # Centre dot
        p.setPen(Qt.NoPen)
        c2 = QColor(bc_qt); c2.setAlpha(180)
        p.setBrush(QBrush(c2))
        p.drawEllipse(QPointF(cx, cy), R_MIN * 0.4, R_MIN * 0.4)

    # ── drag ──────────────────────────────────────────────────────────────────
    def _hit_viz(self, pos):
        """True if pos is near the visualizer anchor."""
        ax, ay = self._viz_anchor()
        sw = self._sw
        half_w = len(self._bars) * max(6, int(sw * 0.008)) // 2 + 20
        return abs(pos.x() - ax) < half_w and abs(pos.y() - ay) < int(self._sh * 0.25)

    def _hit_clock(self, pos):
        return (abs(pos.x() - self.cx) < 350 and
                abs(pos.y() - self.cy) < 120)

    def mousePressEvent(self, e):
        if e.button() != Qt.LeftButton:
            return
        pos = e.pos()
        if self.viz_detached and self.show_viz and self._hit_viz(pos):
            self._drag = "viz"
        elif self.show_clock and self._hit_clock(pos):
            self._drag = "clock"
        elif self.show_viz and self._hit_viz(pos):
            self._drag = "clock"   # attached: drag whole block
        self._drag_origin = pos

    def mouseMoveEvent(self, e):
        if not self._drag or not self._drag_origin:
            return
        d = e.pos() - self._drag_origin
        self._drag_origin = e.pos()
        if self._drag == "clock":
            self.cx = max(200, min(self._sw - 200, self.cx + d.x()))
            self.cy = max(100, min(self._sh - 100, self.cy + d.y()))
            if not self.viz_detached:
                pass   # viz follows automatically via _viz_anchor()
        elif self._drag == "viz":
            self.vx = max(100, min(self._sw - 100, self.vx + d.x()))
            self.vy = max(50,  min(self._sh - 50,  self.vy + d.y()))
        self.update()

    def mouseReleaseEvent(self, _):
        self._drag = None
        self._drag_origin = None


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
        """
        Critical order: embed window into WorkerW FIRST, then start MPV.
        MPV initialises its D3D11 swapchain at player creation time.
        If the window is reparented after that, the swapchain parent changes
        and MPV stops rendering. By embedding first, MPV renders directly
        into the already-parented child window from the start.
        """
        self.stop()
        if not MPV_OK: return False
        self._path = path

        # Step 1: find WorkerW target
        # Force WorkerW creation in case Explorer just started/restarted
        _pm = ctypes.windll.user32.FindWindowW("Progman", None)
        emb = DesktopEmbedder()
        if _pm:
            emb._send_spawn_msg(_pm)  # uses correct 0xD/0x1 params
        target = emb._find_target()
        if not target:
            print("[mpv] Could not find WorkerW target"); return False

        sw, sh = self._sw, self._sh

        # Step 2: create window already as WS_CHILD of the target
        if not self._create_hwnd_as_child(target, sw, sh):
            return False

        # Step 3: fix Z-order and show window before MPV initialises its swapchain.
        #
        # When DefView lives inside Progman (common on Win11), Progman's children are:
        #   [bottom] static wallpaper paint  ← Progman draws this itself
        #   [  ...] our video window          ← must be HERE (above bg, below icons)
        #   [ top ] SHELLDLL_DefView          ← desktop icons, must stay on top
        #
        # SetWindowPos(hwnd, hInsertAfter) places hwnd BELOW hInsertAfter in Z-order.
        # So passing the DefView handle puts our window just under the icons layer.
        # If no DefView is found we fall back to HWND_BOTTOM (1).
        u32 = ctypes.windll.user32
        defview = u32.FindWindowExW(target, None, "SHELLDLL_DefView", None)
        if defview:
            # Insert our window immediately below DefView so icons stay on top
            insert_after = defview
            print(f"[mpv] Placing video below DefView {defview:#010x}")
        else:
            insert_after = 1  # HWND_BOTTOM fallback
            print(f"[mpv] DefView not found in parent, using HWND_BOTTOM")
        u32.SetWindowPos(self._hwnd, insert_after, 0, 0, sw, sh, 0x0010|0x0040|0x0020)
        u32.ShowWindow(self._hwnd, 5)
        u32.UpdateWindow(self._hwnd)
        time.sleep(0.15)  # let Windows finish layout before MPV swapchain init
        self._embedded_target = target

        # Step 4: start MPV — it now renders into a properly parented child window
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
            print(f"[mpv] Creating player with wid={self._hwnd:#010x} (child of {target:#010x})")
            self._player = mpv.MPV(**opts)
            self._player.play(path)
            print(f"[mpv] Playing: {path}")
            print(f"[mpv] HWND: {self._hwnd:#010x}")
            print(f"[mpv] Parent: {ctypes.windll.user32.GetParent(self._hwnd):#010x}")
            # Give mpv a moment to open the file — don't block on wait_for_property
            # as it can hang if the video-params event never fires (e.g. bad codec).
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
        # Stop frame extractor first so cv2.VideoCapture is released
        if self._worker_stop: self._worker_stop.set()
        if self._worker and self._worker.is_alive(): self._worker.join(2)
        self._worker = self._worker_stop = None

        player = self._player
        self._player = None
        if player:
            # Terminate mpv in a background thread so we never block the main thread.
            # We then wait at most 2s for it to finish before proceeding.
            done = threading.Event()
            def _kill():
                try: player.terminate()
                except: pass
                try: player.wait_for_shutdown(timeout=2)
                except: pass
                done.set()
            t = threading.Thread(target=_kill, daemon=True); t.start()
            done.wait(timeout=2.5)   # never block Qt main thread longer than this
            time.sleep(0.1)          # let OS release D3D11 swapchain

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
        u32 = ctypes.windll.user32
        last_tray   = u32.FindWindowW('Shell_TrayWnd', None)
        last_parent = self._target
        print("[Watchdog] Started")
        while self._watchdog_stop and not self._watchdog_stop.wait(1.0):
            if not self._embedded or not self._hwnd: break
            needs = False
            tray = u32.FindWindowW('Shell_TrayWnd', None)
            if tray and tray != last_tray:
                print(f"[Watchdog] Explorer restarted"); last_tray=tray; needs=True
            cp = u32.GetParent(self._hwnd)
            if cp != last_parent:
                print(f"[Watchdog] Parent changed"); last_parent=cp; needs=True
            if needs and self._reembed_cb:
                try: self._reembed_cb()
                except: pass
        print("[Watchdog] Stopped")

    def _create_hwnd(self):
        global _WND_CLASS_REGISTERED, _WNDPROC_CB
        if not _WND_CLASS_REGISTERED:
            if not self._register_wnd_class(): return False
        hInst = ctypes.windll.kernel32.GetModuleHandleW(None)
        hwnd  = ctypes.windll.user32.CreateWindowExW(
            0, "WallpaperEngineHost", "WallpaperEngineVideo",
            0x80000000|0x02000000,             # WS_POPUP|WS_CLIPCHILDREN (no WS_VISIBLE yet)
            0,0,self._sw,self._sh, None,None,hInst,None)
        if not hwnd:
            print(f"[mpv] CreateWindowExW failed: {ctypes.GetLastError()}"); return False
        self._hwnd = hwnd
        print(f"[mpv] Created HWND {hwnd:#010x}")
        return True

    def _create_hwnd_as_child(self, parent_hwnd, sw, sh):
        """Create the video window directly as a WS_CHILD of parent_hwnd.
        This avoids the reparent-after-render problem entirely."""
        global _WND_CLASS_REGISTERED
        if not _WND_CLASS_REGISTERED:
            if not self._register_wnd_class(): return False
        hInst = ctypes.windll.kernel32.GetModuleHandleW(None)
        # WS_CHILD=0x40000000, WS_CLIPCHILDREN=0x02000000,
        # WS_CLIPSIBLINGS=0x04000000, WS_VISIBLE=0x10000000
        hwnd = ctypes.windll.user32.CreateWindowExW(
            0, "WallpaperEngineHost", "WallpaperEngineVideo",
            0x40000000|0x02000000|0x04000000|0x10000000,
            0, 0, sw, sh,
            parent_hwnd, None, hInst, None)
        if not hwnd:
            err = ctypes.GetLastError()
            print(f"[mpv] CreateWindowExW (child) failed: {err}")
            return False
        self._hwnd = hwnd
        print(f"[mpv] Created child HWND {hwnd:#010x} under {parent_hwnd:#010x}")
        return True

    def _register_wnd_class(self):
        global _WND_CLASS_REGISTERED, _WNDPROC_CB

        # ─── FIX: On 64-bit Windows LPARAM/WPARAM are 64-bit values.
        # Using ctypes.wintypes.LPARAM (32-bit) causes OverflowError.
        # Must use ctypes.c_ssize_t (LRESULT/LPARAM) and c_size_t (WPARAM).
        WNDPROC_TYPE = ctypes.WINFUNCTYPE(
            ctypes.c_ssize_t,       # LRESULT  (return)
            ctypes.wintypes.HWND,   # hWnd
            ctypes.wintypes.UINT,   # uMsg
            ctypes.c_size_t,        # WPARAM   (unsigned 64-bit on x64)
            ctypes.c_ssize_t,       # LPARAM   (signed   64-bit on x64)
        )

        # Pre-declare argtypes so ctypes won't silently truncate values
        _DefWndProc = ctypes.windll.user32.DefWindowProcW
        _DefWndProc.restype  = ctypes.c_ssize_t
        _DefWndProc.argtypes = [
            ctypes.wintypes.HWND,
            ctypes.wintypes.UINT,
            ctypes.c_size_t,
            ctypes.c_ssize_t,
        ]

        def wndproc(hwnd, msg, wp, lp):
            return _DefWndProc(hwnd, msg, wp, lp)

        _WNDPROC_CB = WNDPROC_TYPE(wndproc)

        class WNDCLASSEXW(ctypes.Structure):
            _fields_ = [
                ("cbSize",        ctypes.wintypes.UINT),
                ("style",         ctypes.wintypes.UINT),
                ("lpfnWndProc",   WNDPROC_TYPE),
                ("cbClsExtra",    ctypes.c_int),
                ("cbWndExtra",    ctypes.c_int),
                ("hInstance",     ctypes.wintypes.HANDLE),
                ("hIcon",         ctypes.wintypes.HANDLE),
                ("hCursor",       ctypes.wintypes.HANDLE),
                ("hbrBackground", ctypes.wintypes.HANDLE),
                ("lpszMenuName",  ctypes.wintypes.LPCWSTR),
                ("lpszClassName", ctypes.wintypes.LPCWSTR),
                ("hIconSm",       ctypes.wintypes.HANDLE),
            ]

        wc = WNDCLASSEXW()
        wc.cbSize       = ctypes.sizeof(WNDCLASSEXW)
        wc.lpfnWndProc  = _WNDPROC_CB
        wc.hInstance    = ctypes.windll.kernel32.GetModuleHandleW(None)
        wc.hbrBackground= 0
        wc.lpszClassName= "WallpaperEngineHost"

        res = ctypes.windll.user32.RegisterClassExW(ctypes.byref(wc))
        if res == 0:
            err = ctypes.GetLastError()
            if err == 1410:   # ERROR_CLASS_ALREADY_EXISTS — that's fine
                _WND_CLASS_REGISTERED = True; return True
            print(f"[mpv] RegisterClassExW failed: {err}"); return False
        _WND_CLASS_REGISTERED = True; return True

    def _destroy_hwnd(self):
        if self._hwnd:
            ctypes.windll.user32.DestroyWindow(self._hwnd); self._hwnd=None

    def _start_extractor(self):
        self._worker_stop = threading.Event()
        path = self._path
        def run():
            cap = None
            time.sleep(1.5)
            while not self._worker_stop.wait(3.5):
                if not self._player or not self.frame_ready_cb:
                    continue
                try:
                    if cap is None:
                        cap = cv2.VideoCapture(path)
                    if not cap or not cap.isOpened():
                        cap = None; continue
                    # Sync seek to current mpv position so colours match the scene
                    try:
                        pos = self._player.time_pos
                        if pos is not None:
                            cap.set(cv2.CAP_PROP_POS_MSEC, float(pos) * 1000)
                    except Exception:
                        pass
                    ret, frame = cap.read()
                    if ret and frame is not None:
                        self.frame_ready_cb(frame)
                except Exception as e:
                    print(f"[Extract] {e}")
            if cap:
                cap.release()
        self._worker = threading.Thread(target=run, daemon=True)
        self._worker.start()

# ══════════════════════════════════════════════
#  WIDGETS
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
            nc=tuple(int(cv+(tv-cv)*.18) for cv,tv in zip(c,t))  # faster lerp
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
#  MAIN PANEL
# ══════════════════════════════════════════════

class Panel(QWidget):
    _frame_signal = pyqtSignal(object)   # thread-safe bridge: worker → main thread

    def __init__(self):
        super().__init__()
        self._ext    = ColorExtractor()
        self._themer = WindowsThemer()
        self._emb    = DesktopEmbedder()
        self._wp     = MpvWallpaper()
        # Use a real Qt signal so the worker thread can safely deliver frames to
        # the main thread.  QTimer.singleShot called from a non-main thread is
        # not safe in Qt5 and silently drops events.
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
        self._palette_dynamic = True   # True=follows video, False=locked to static colour

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
        self.setWindowTitle("Theme Engine v4.2")
        self.setFixedWidth(460); self.setWindowFlags(Qt.Window)
        ml=QVBoxLayout(self); ml.setContentsMargins(0,0,0,0); ml.setSpacing(0)
        self._hdr=QWidget(); self._hdr.setFixedHeight(62)
        hl=QHBoxLayout(self._hdr); hl.setContentsMargins(18,0,16,0); hl.setSpacing(0)
        logo=QLabel("▶"); logo.setStyleSheet("font-size:22px;color:white;padding-right:10px;")
        title=QLabel("Theme Engine"); title.setStyleSheet("font-size:17px;font-weight:800;color:white;letter-spacing:.5px;")
        ver=QLabel("v4.2"); ver.setStyleSheet("font-size:10px;color:rgba(255,255,255,130);padding-left:6px;padding-top:5px;")
        hl.addWidget(logo); hl.addWidget(title); hl.addWidget(ver); hl.addStretch()
        ml.addWidget(self._hdr)
        scroll=QScrollArea(); scroll.setWidgetResizable(True); scroll.setFrameShape(QScrollArea.NoFrame)
        body=QWidget(); bl=QVBoxLayout(body); bl.setContentsMargins(16,14,16,18); bl.setSpacing(0)
        def gap(n=8): s=QWidget(); s.setFixedHeight(n); return s
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
        # Dynamic / Static radio row
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
        # Taskbar test/diagnose button
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
        # Style picker row
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
        # Only rebuild QSS + taskbar if colour shifted more than ~15 units (avoids thrashing)
        dr,dg,db = abs(acc[0]-self._acc[0]), abs(acc[1]-self._acc[1]), abs(acc[2]-self._acc[2])
        if dr+dg+db < 30: return   # only update on meaningful colour shift
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
        # Clean up any currently running session before starting a new one
        if self._audio_on:
            self._audio.stop(); self._audio_on = False
            if hasattr(self,'_b_audio'): self._b_audio.setText("▶  Start Audio Capture")
        self._embedded = False
        self._vpath=path; self._vid_lbl.setText(f"📹  {Path(path).name}")
        self._status.info("Loading..."); QApplication.processEvents()
        # load() must run on the main thread — Win32 CreateWindowExW requires it.
        # stop() already terminates mpv in a background thread (max 2.5s wait).
        ok = self._wp.load(path)
        self._on_load_done(ok)

    def _on_load_done(self, ok):
        if ok:
            self._playing=True; self._upause=False
            for b in (self._b_play,self._b_pause,self._b_stop): b.setEnabled(True)
            self._b_embed.setEnabled(True); self._status.ok("Loaded"); self._b_play.set_active(True)
            QTimer.singleShot(500,  self._toggle_embed)
            QTimer.singleShot(1500, self._auto_start_audio)
            # Force taskbar colour update after extractor has had time to sample a frame
            QTimer.singleShot(5000, self._force_apply_theme)
        else:
            self._status.err("Could not open video — is libmpv-2.dll present?")

    def _force_apply_theme(self):
        """Apply taskbar colour immediately, bypassing the delta-change guard."""
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
            # load() already created the window as a child of WorkerW and started MPV.
            # Just show the overlay and update UI state.
            hw = self._wp.winId()
            if hw:
                self._wp.set_embedded(True, getattr(self._wp, '_embedded_target', 0))
                self._ov.show()
                self._ov.raise_()   # ensure overlay is above the video window
                self._ov.update()
                self._embedded = True
                self._b_embed.setText("✕  Detach Wallpaper"); self._b_embed.set_accent(*self._acc)
                self._status.ok("Wallpaper active!")
            else:
                self._status.err("No window handle — reload the video")
        else:
            self._wp.set_embedded(False)
            # Detach: reparent back to desktop and hide
            hw = self._wp.winId()
            if hw:
                ctypes.windll.user32.SetParent(hw, 0)
                ctypes.windll.user32.ShowWindow(hw, 0)
            self._ov.hide(); self._embedded = False
            self._b_embed.setText("🖥  Set as Live Wallpaper"); self._b_embed.set_accent(*self._acc)
            self._status.info("Detached")

    def _open_taskbar_settings(self):
        """Open Windows Settings directly to the Colours page for one-time manual enable,
        then immediately apply the current accent colour once the user returns."""
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
        self._palette_dynamic = dynamic
        self._pal.set_dynamic(dynamic)
        self._b_statcol.setEnabled(not dynamic)
        if dynamic:
            # Re-apply current video colours immediately
            self._pal.update_colors(self._ext.get_colors())

    def _pick_static_col(self):
        c = QColorDialog.getColor(QColor(*self._acc), self, "Static Palette Colour")
        if not c.isValid(): return
        r, g, b = c.red(), c.green(), c.blue()
        static_colors = [(r,g,b),
                         (min(255,r+40),min(255,g+40),min(255,b+40)),
                         (max(0,r-40),max(0,g-40),max(0,b-40)),
                         (min(255,r+80),min(255,g+15),min(255,b+15)),
                         (min(255,r+15),min(255,g+80),min(255,b+15)),
                         (min(255,r+15),min(255,g+15),min(255,b+80))]
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
        """Called after a video loads. Auto-starts the audio visualizer.
        pyaudiowpatch gives us zero-permission WASAPI loopback — no Stereo Mix,
        no driver config, no dialog needed."""
        if not AUDIO_OK or self._audio_on:
            return
        if self._audio.can_capture():
            self._toggle_audio()
            self._status.ok("Audio visualizer started (WASAPI loopback)")
        else:
            print("[Audio] No loopback source available — visualizer inactive")

    def _on_viz_style(self, idx):
        styles = ["bars","slim","mirror","wave","dots","circle"]
        self._ov.set_viz_style(styles[idx] if idx < len(styles) else "bars")

    def _on_viz_detach(self, state):
        self._ov.viz_detached = bool(state)
        self._ov.update()

    def _toggle_audio(self):
        if not self._audio_on:
            self._audio.start(); self._audio_on=True
            if hasattr(self,'_b_audio'): self._b_audio.setText("⏹  Stop Capture")
        else:
            self._audio.stop(); self._audio_on=False
            if hasattr(self,'_b_audio'): self._b_audio.setText("▶  Start Audio Capture")
            self._ov.set_bars([0.0]*48)

def _detach_from_console():
    """
    If the user launched us from cmd/PowerShell, re-launch using pythonw.exe
    (which never opens a console window) as a fully detached process, then
    exit the terminal-attached copy.  This means closing the terminal will
    no longer kill the wallpaper.
    """
    # When frozen (PyInstaller exe) the exe is already a GUI app with no
    # console.  Re-launching would spawn another copy — skip entirely.
    if getattr(sys, "frozen", False):
        return

    # Already detached if no console window exists
    if ctypes.windll.kernel32.GetConsoleWindow() == 0:
        return
    # Already running under pythonw
    if 'pythonw' in sys.executable.lower():
        return

    pythonw = sys.executable.replace('python.exe', 'pythonw.exe')
    if not os.path.exists(pythonw):
        pythonw = os.path.join(os.path.dirname(sys.executable), 'pythonw.exe')
    if not os.path.exists(pythonw):
        pythonw = sys.executable  # fallback: same exe but still detach

    import subprocess
    DETACHED_PROCESS      = 0x00000008
    CREATE_NEW_PROC_GROUP = 0x00000200
    subprocess.Popen(
        [pythonw] + sys.argv,
        creationflags=DETACHED_PROCESS | CREATE_NEW_PROC_GROUP,
        close_fds=True,
        cwd=os.getcwd(),
        stdin=subprocess.DEVNULL,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
    sys.exit(0)   # kill the console-attached copy


def _create_desktop_shortcut():
    """Create a desktop shortcut on first run using PowerShell WScript.Shell.
    Uses a marker file to skip on subsequent runs."""
    marker = os.path.join(_script_dir, '.shortcut_created')
    if os.path.exists(marker):
        return
    try:
        desktop = os.path.join(os.path.expanduser('~'), 'Desktop')
        # Try the localised shell folder path as fallback
        try:
            import winreg as _wr
            with _wr.OpenKey(_wr.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders") as k:
                desktop = _wr.QueryValueEx(k, "Desktop")[0]
        except Exception:
            pass

        shortcut_path = os.path.join(desktop, 'Theme Engine.lnk')
        pythonw = sys.executable.replace('python.exe', 'pythonw.exe')
        if not os.path.exists(pythonw):
            pythonw = sys.executable
        script_path = os.path.abspath(__file__)

        ps = (
            f'$s=(New-Object -ComObject WScript.Shell).CreateShortcut("{shortcut_path}");'
            f'$s.TargetPath="{pythonw}";'
            f'$s.Arguments=\'"{script_path}"\';'
            f'$s.WorkingDirectory="{_script_dir}";'
            f'$s.IconLocation="{script_path},0";'
            f'$s.Description="Live Theme Engine";'
            f'$s.Save()'
        )
        import subprocess
        subprocess.run(
            ['powershell', '-WindowStyle', 'Hidden', '-NoProfile', '-Command', ps],
            creationflags=0x08000000,  # CREATE_NO_WINDOW
            timeout=10
        )
        # Write marker so we don't repeat this next launch
        with open(marker, 'w') as f:
            f.write('1')
        print(f"[Setup] Desktop shortcut created: {shortcut_path}")
    except Exception as e:
        print(f"[Setup] Shortcut creation failed: {e}")


def _first_run_taskbar_prompt():
    """On very first run, show a one-time message telling the user to enable
    taskbar colour in Settings — since Windows 11 requires it manually once."""
    marker = os.path.join(_script_dir, '.taskbar_setup_shown')
    if os.path.exists(marker):
        return
    try:
        with open(marker, 'w') as f:
            f.write('1')
    except Exception:
        pass
    # Delay so the main window is visible first
    QTimer.singleShot(2000, lambda: _show_taskbar_first_run())


def _show_taskbar_first_run():
    import subprocess
    reply = QMessageBox.question(
        None, "Enable Taskbar Colour (one-time)",
        "For the taskbar to change colour with the wallpaper, Windows needs\n"
        "one setting enabled manually (required by Windows 11).\n\n"
        "Click Yes to open Settings now.\n\n"
        "In Settings → Personalisation → Colours, turn ON:\n"
        "  • Show accent colour on title bars\n"
        "  • Show accent colour on Start and taskbar\n\n"
        "You only need to do this once.",
        QMessageBox.Yes | QMessageBox.Cancel
    )
    if reply == QMessageBox.Yes:
        subprocess.Popen(['explorer.exe', 'ms-settings:colors'])


def main():
    _detach_from_console()   # must be first — before Qt even starts
    _create_desktop_shortcut()
    set_dpi_aware()
    QApplication.setQuitOnLastWindowClosed(False)
    app=QApplication(sys.argv); app.setStyle("Fusion")
    panel=Panel(); panel.show()
    _first_run_taskbar_prompt()
    sys.exit(app.exec_())

if __name__=="__main__":
    main()
