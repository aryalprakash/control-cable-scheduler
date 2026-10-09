# Control Cable Scheduler

Desktop app (Python + Tkinter) to extract terminal/ferrule data from panel drawings (PDF/DXF), review external wiring, generate a control cable schedule, and export to Excel.

## Download pre-built apps (GitHub Actions)

1. Open [Actions](https://github.com/aryalprakash/control-cable-scheduler/actions) on this repo
2. Open the latest **Build executables** run (green check)
3. Download artifacts:
   - **macos-app** → unzip → `Control Cable Scheduler.app`
   - **windows-exe** → unzip → run `Control Cable Scheduler.exe`

**Unsigned builds:** macOS may block the app — right-click → Open, or allow it under System Settings → Privacy & Security. Windows SmartScreen may warn — choose More info → Run anyway.

## Run from source

```bash
python3 -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

Requires Python 3.10+ with Tkinter.

## Build locally

**macOS / Linux**

```bash
chmod +x build.sh
./build.sh
# → dist/Control Cable Scheduler.app  (macOS)
```

**Windows (PowerShell)**

```powershell
.\build.ps1
# → dist\Control Cable Scheduler\Control Cable Scheduler.exe
```

Or manually:

```bash
pip install -r requirements-build.txt
pyinstaller --noconfirm control_cable_scheduler.spec
```

## Workflow

1. Create or open a project (`.ccs`)
2. Enter substation → add voltage levels and bays
3. Select a bay → upload drawings (PDF/DXF) → Extract information
4. Review External TB / ferrule → Generate cable schedule → Export to Excel
