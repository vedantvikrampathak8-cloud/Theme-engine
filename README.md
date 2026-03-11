# 🎨 Theme Engine v4.2

A live wallpaper engine for Windows that plays video as your desktop background and dynamically adapts your system accent colour, taskbar, and title bars to match the mood of whatever is playing.
---

## ⬇️ Download

| File | Description |
|---|---|
| [**ThemeEngine.exe**](https://github.com/vedantvikrampathak8-cloud/Theme-engine/releases/latest/download/ThemeEngine.exe) | Ready-to-run exe — no install needed |
| [**libmpv-2.dll**](https://github.com/vedantvikrampathak8-cloud/Theme-engine/releases/latest/download/libmpv-2.dll) | mpv media engine (place next to the exe) |

👉 **[View all releases](https://github.com/vedantvikrampathak8-cloud/Theme-engine/releases/latest)**

Place both files in the same folder and run `ThemeEngine.exe`. That's it.

> **Requires:** Windows 10/11 64-bit · Visual C++ Redistributable 2022 x64 ([download](https://aka.ms/vs/17/release/vc_redist.x64.exe) — most PCs already have it)

---

## Features

### 🎬 Live Video Wallpaper
- Plays any video file (MP4, MKV, WEBM, AVI, GIF, and more) directly on your desktop using **libmpv**
- Hardware-accelerated via D3D11 / DXVA2 — negligible CPU usage
- Renders underneath desktop icons, just like Steam Wallpaper Engine
- Loop, pause, resume controls
- Auto-pauses when a fullscreen app covers the desktop

### 🎨 Adaptive Colour System
- Extracts dominant colours from the video in real time using K-means clustering
- Applies the accent colour to:
  - Windows taskbar
  - Window title bars (via DWM)
  - The Theme Engine UI itself
- **Dynamic mode** — colour follows the video automatically
- **Static mode** — lock to a colour of your choice
- Flicker-free: only pokes the taskbar directly, never broadcasts to every window

### 🕐 Clock & Date Overlay
- Large, elegant clock with day, date, and time rendered on the desktop
- Fully draggable — click and drag anywhere on the desktop to reposition
- Customisable font (any installed font), text colour, and opacity
- Toggle seconds on/off

### 📊 Music Visualizer
Six visualizer styles to choose from:

| Style | Description |
|---|---|
| **Bars** | Classic gradient bars growing upward |
| **Slim** | Narrow bars with wider gaps — minimal look |
| **Mirror** | Bars grow both up and down from a centre line |
| **Wave** | Smooth spline through bar tops with filled area underneath |
| **Dots** | Stacked glowing circles per frequency band |
| **Circle** | Radial spectrum radiating outward from a centre point |

- **WASAPI loopback** audio capture — captures whatever plays through your speakers with **zero permissions needed**, no Stereo Mix setup, no driver configuration (same API used by Steam Wallpaper Engine)
- Falls back to sounddevice / Stereo Mix on older systems
- Per-bar auto-equaliser — all frequency bands self-level so bass doesn't dominate
- Bars decay to zero quickly when audio stops — no frozen bars
- Auto-colours from the wallpaper, or pick a static colour
- **Detachable** from the clock — drag independently when detached

### ✨ Sparkle Cursor Trail
- Particle system follows your mouse cursor on the desktop
- Sparkles inherit the current visualizer colour and fade with physics (gravity, velocity)
- Fully toggleable

---

## Requirements

### To run `ThemeEngine.exe` (end users)
- Windows 10 or Windows 11 (64-bit)
- Visual C++ Redistributable 2022 x64 — [download here](https://aka.ms/vs/17/release/vc_redist.x64.exe) (most PCs already have it)
- That's it — `ThemeEngine.exe` is fully self-contained

### To run from source (`wallpaper_engine2.py`)
- Python 3.10+ (64-bit)
- `libmpv-2.dll` placed in the same folder as the script
- Python packages (auto-installed on first run):
  ```
  PyQt5
  numpy
  opencv-python
  python-mpv
  pyaudiowpatch
  sounddevice
  ```

### To build the exe (`build_theme_engine.py`)
- Python 3.10+ (64-bit)
- `wallpaper_engine2.py` and `libmpv-2.dll` in the same folder
- PyInstaller (auto-installed by the build script)

---

## Getting Started

### Running from source
1. Place `wallpaper_engine2.py` and `libmpv-2.dll` in the same folder
2. Run:
   ```
   python wallpaper_engine2.py
   ```
3. On first run, a setup window will install the required Python packages automatically (one time only, takes ~1 minute)
4. Once loaded, click **Browse** to pick a video file, then **▶ Play**

### Building the exe
1. Place `build_theme_engine.py`, `wallpaper_engine2.py`, and `libmpv-2.dll` in the same folder
2. Run:
   ```
   python build_theme_engine.py
   ```
3. The output is `dist/ThemeEngine.exe` — upload this file (and `libmpv-2.dll`) to the GitHub Release
4. 
---

## Usage

### Setting a live wallpaper
1. Click **Browse** and select a video file
2. Click **▶ Play** — the video appears on your desktop
3. Click **🖥 Set as Live Wallpaper** to embed it behind the desktop icons

### Repositioning the clock and visualizer
- **Clock:** click and drag anywhere near the clock on your desktop
- **Visualizer (attached):** drag the clock — the visualizer moves with it
- **Visualizer (detached):** tick *Detach from clock* in the panel, then drag the visualizer independently

### Enabling the music visualizer
1. Play some audio (music, video, anything through your speakers)
2. Click **▶ Start Audio Capture** — bars appear automatically
3. The visualizer uses WASAPI loopback — no extra setup required on Windows 10/11

### Adaptive colour
- **Dynamic** (default): accent colour follows the dominant colour of the current video frame every 8 seconds
- **Static**: select *Static colour* and click the colour picker to lock to a specific colour
- Tick *Apply colour to Windows taskbar & title bars* to theme the system UI

---

## Panel Sections

| Section | Controls |
|---|---|
| **Video Source** | File browser, play/pause/stop |
| **Playback** | Embed/detach wallpaper, pause status |
| **Adaptive Colour Palette** | Dynamic/static mode, colour picker, taskbar toggle |
| **Clock & Date Overlay** | Show/hide, seconds toggle, font, colour, opacity, sparkle toggle |
| **Music Visualizer** | Show/hide, style, detach, auto-colour, colour picker, capture start/stop |
| **Behaviour** | Auto-pause on fullscreen, loop video |

---

## Known Limitations

- **Windows only** — uses Win32 APIs (DWM, WASAPI, WorkerW embedding)
- The wallpaper renders behind desktop icons but above the static wallpaper; this is the same method used by Steam Wallpaper Engine
- If Explorer crashes and restarts, click **🖥 Set as Live Wallpaper** again to re-embed
- On some systems the adaptive colour change may take up to 8 seconds to update the taskbar after a scene change

---

## Project Structure

```
theme-engine/
├── wallpaper_engine2.py     # Main application (run directly or build into exe)
├── build_theme_engine.py    # PyInstaller build script → produces ThemeEngine.exe
├── README.md                # This file
│
│   (not in repo — attached to GitHub Release)
├── ThemeEngine.exe          # Built exe — upload as release asset
└── libmpv-2.dll             # mpv media engine — upload as release asset
```

---

## Acknowledgements

- [mpv](https://mpv.io/) — the media engine powering video playback
- [shinchiro/mpv-winbuild-cmake](https://github.com/shinchiro/mpv-winbuild-cmake) — Windows builds of libmpv
- [python-mpv](https://github.com/jaseg/python-mpv) — Python bindings for libmpv
- [pyaudiowpatch](https://github.com/s0d3s/PyAudioWPatch) — WASAPI loopback audio capture
- [PyQt5](https://riverbankcomputing.com/software/pyqt/) — UI framework
