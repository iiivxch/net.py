# SpeedMeter

A lightweight Windows tray app and mini overlay that shows real-time internet speed (Download/Upload) with a compact UI.

## Features
- Horizontal mini overlay (movable, always-on-top, opacity control)
- Units: MBps, Mbps, Kbps
- 10s rolling average in Dashboard and tray tooltip
- Network interface selection
- Daily/monthly data usage
- Start with Windows (autostart toggle)
- EXE build with PyInstaller and GitHub Actions

## Requirements (dev)
- Python 3.10+ on Windows
- `pip install -r requirements.txt`

## Run
```bash
python net.py
