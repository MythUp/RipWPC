# RipWPC

Simple Tkinter UI to start/stop both the Windows Parental Controls service (`WpcMonSvc`) and its process (`WpcMon.exe`). A single button reflects the current state and toggles both service and program together. The console window is hidden on launch.

## Requirements
- Windows
- Python 3.10+ (tested with the standard library `tkinter`)
- Dependencies:
  - `psutil`

Install dependencies:
```cmd
pip install psutil
```

## Run
From the repository root:
```cmd
python RipWPC.py
```
The UI will appear and the console window is hidden automatically.

## Build (PyInstaller)
To produce an elevated, console-less single executable:
```cmd
pyinstaller --name RipWPC --uac-admin --onefile --noconsole --add-data "manifest.json;." RipWPC.py
```
Artifacts will be in the `dist/` directory. If PyInstaller reports an issue with the obsolete `typing` backport, uninstall it first:
```cmd
python -m pip uninstall typing
```
