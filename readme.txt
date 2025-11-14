=========================================================
DAVINCI ECU AUTOMATION — DEPLOYMENT & RUNBOOK
=========================================================

0) OVERVIEW
------------
Automates DAVINCI ECU workflow: launch → open file → wait for processing → save .mod.
Script names:
  - davinci_automation.py  (core automation)
  - agent.py               (local API bridge for website button)
Target OS: Windows 10 or 11
Requires Python 3.12 x64 and GUI desktop access.

---------------------------------------------------------
1) FILES PROVIDED
---------------------------------------------------------
C:\Program Files\DAVINCI\
    davinci_automation.py
    agent.py
    loading.png   (optional template image)
C:\davinci_automation\   (log directory, created automatically)
C:\ecu_files\original\   (input folder)
C:\ecu_files\modified\   (output folder)

---------------------------------------------------------
2) PREREQUISITES
---------------------------------------------------------
- Windows 10 or 11, desktop unlocked
- DAVINCI installed (C:\Program Files\DAVINCI\davinci.exe)
- Python 3.12 x64 installed with “Add to PATH” ticked
- Display scaling = 100%
- Disable sleep/screensaver
- Writable directories: C:\davinci_automation\, C:\ecu_files\original\, C:\ecu_files\modified\
- GUI session active (not minimized or RDP-hidden)

---------------------------------------------------------
3) FOLDER LAYOUT
---------------------------------------------------------
C:\Program Files\DAVINCI\
    davinci.exe
    davinci_automation.py
    agent.py
    loading.png
C:\davinci_automation\
    davinci_automation.log  (created automatically)
C:\ecu_files\
    original\   →  incoming BINs
    modified\   →  generated .mod files

---------------------------------------------------------
4) ONE-TIME SETUP COMMANDS
---------------------------------------------------------
py -0p
py -3.12 -V

C:
mkdir C:\davinci_venv
py -3.12 -m venv C:\davinci_venv

C:\davinci_venv\Scripts\activate

python -m pip install --upgrade pip wheel
pip install --only-binary=:all: numpy opencv-python
pip install pyautogui pygetwindow pillow fastapi uvicorn requests

mkdir C:\davinci_automation
mkdir C:\ecu_files\original C:\ecu_files\modified

---------------------------------------------------------
5) RUNNING THE AUTOMATION (MANUAL)
---------------------------------------------------------
C:\davinci_venv\Scripts\activate
cd /d C:\davinci_automation

python "C:\Program Files\DAVINCI\davinci_automation.py" ^
  --exe "C:\Program Files\DAVINCI\davinci.exe" ^
  --open-direct ^
  --input "C:\ecu_files\original\vw_golf_edc17.bin" ^
  --outdir "C:\ecu_files\modified" ^
  --timeout-load 90 ^
  --timeout-process 600 ^
  --timeout-save 120

type C:\davinci_automation\davinci_automation.log

---------------------------------------------------------
6) WHAT THE SCRIPT DOES
---------------------------------------------------------
1. Launches DAVINCI directly with the input BIN file.
2. Waits for DAVINCI window to appear.
3. Waits for processing to finish (“Save Mod File” dialog).
4. Saves .mod output to C:\ecu_files\modified.
5. Logs every step to C:\davinci_automation\davinci_automation.log.

---------------------------------------------------------
7) CONFIGURATION OPTIONS
---------------------------------------------------------
--exe             Path to DAVINCI executable
--input           Input ECU file (.bin/.ori)
--outdir          Output folder for .mod
--timeout-load    Wait for main window (sec)
--timeout-process Wait for processing completion (sec)
--timeout-save    Wait for Save dialog (sec)

Dialog title constants (inside script):
DEFAULT_MAIN_TITLE_HINT = "DAVINCI"
SAVE_DIALOG_TITLE_HINT  = "Save Mod File"

---------------------------------------------------------
8) OPERATIONAL TIPS
---------------------------------------------------------
- Keep Windows user logged in and display unlocked.
- Move mouse to top-left corner to abort safely.
- Avoid multitasking or minimizing the window.
- Disable DAVINCI auto-updates/popups.
- Use 100% display scaling.

---------------------------------------------------------
9) TROUBLESHOOTING
---------------------------------------------------------
A) pip not recognized
    → Use: python -m pip install ...
B) No space left on device
    → Clean %temp%, empty Recycle Bin
C) NumPy/OpenCV build errors
    → Install Python 3.12 x64, not 3.14
D) ModuleNotFoundError: pyautogui
    → Activate venv or call its python explicitly
E) PermissionError on log file
    → Logs must be written to C:\davinci_automation
F) DAVINCI not opening
    → Check correct path or try Run as Administrator
G) “Save Mod File” dialog never appears
    → Increase --timeout-process
H) Scaling issues
    → Set 100% DPI
I) UAC prompts
    → Run Command Prompt as Administrator

---------------------------------------------------------
10) VERIFICATION
---------------------------------------------------------
- .mod file appears in C:\ecu_files\modified
- C:\davinci_automation\davinci_automation.log ends with “SUCCESS”
- Repeat test with another BIN

---------------------------------------------------------
11) MAC NOTE
---------------------------------------------------------
Run inside Windows VM (Parallels or remote Windows PC). macOS cannot drive DAVINCI natively.

---------------------------------------------------------
12) SUPPORT DATA FOR DEBUG
---------------------------------------------------------
- C:\davinci_automation\davinci_automation.log
- Screenshot of DAVINCI window
- python -V
- pip list
- Windows scaling value

---------------------------------------------------------
13) QUICK RUN BLOCK (MANUAL)
---------------------------------------------------------
C:\davinci_venv\Scripts\activate
cd /d C:\davinci_automation
python "C:\Program Files\DAVINCI\davinci_automation.py" ^
  --exe "C:\Program Files\DAVINCI\davinci.exe" ^
  --open-direct ^
  --input "C:\ecu_files\original\vw_golf_edc17.bin" ^
  --outdir "C:\ecu_files\modified" ^
  --timeout-load 90 ^
  --timeout-process 600 ^
  --timeout-save 120
type C:\davinci_automation\davinci_automation.log

---------------------------------------------------------
14) LOCAL AGENT (agent.py)
---------------------------------------------------------
Purpose:
- Lets website buttons trigger DAVINCI automatically.
- Listens on http://127.0.0.1:8765
- Accepts /process_upload (POST with file)

Manual Run:
---------------------------------------------------------
C:\davinci_venv\Scripts\activate
cd /d "C:\Program Files\DAVINCI"
C:\davinci_venv\Scripts\uvicorn.exe agent:APP --host 127.0.0.1 --port 8765 --reload

Health Check:
---------------------------------------------------------
curl http://127.0.0.1:8765/health
Expected:
{"ok": true}

Process Test (Upload method):
---------------------------------------------------------
curl -F "file=@C:\ecu_files\original\vw_golf_edc17.bin" http://127.0.0.1:8765/process_upload

Output .mod file appears in:
C:\ecu_files\modified

Log output:
C:\davinci_automation\davinci_automation.log

---------------------------------------------------------
15) OPTIONAL AUTOSTART
---------------------------------------------------------
echo C:\davinci_venv\Scripts\uvicorn.exe agent:APP --host 127.0.0.1 --port 8765 > C:\davinci_automation\start_agent.cmd

schtasks /Create /TN "DavinciLocalAgent" /TR "C:\davinci_automation\start_agent.cmd" /SC ONLOGON /RU "USERNAME"

Agent runs at every login automatically.

---------------------------------------------------------
16) SUMMARY
---------------------------------------------------------
- davinci_automation.py handles GUI workflow.
- agent.py exposes local REST API.
- Website triggers → agent downloads file → launches DAVINCI → waits → saves .mod.
- Logs and outputs in C:\davinci_automation and C:\ecu_files\modified.
- Requires open desktop session and venv Python environment.

---------------------------------------------------------
END OF RUNBOOK (C: VERSION, WITH AGENT)
---------------------------------------------------------