# agent.py — open DaVinci, select BRAND→ECU, feed BIN (no save)

from fastapi import FastAPI, UploadFile, File, Form
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import Response
from pathlib import Path
import subprocess, shutil, json, sys, re, requests, base64, os

APP = FastAPI()
APP.add_middleware(
    CORSMiddleware,
    allow_origins=["*"], allow_methods=["*"], allow_headers=["*"], max_age=86400
)

EXE     = r"C:\Program Files\DAVINCI\davinci.exe"
SCRIPT  = r"C:\Program Files\DAVINCI\davinci_automation.py"
PYTHON  = r"C:\davinci_venv\Scripts\python.exe"
INDIR   = Path(r"C:\ecu_files\original")
WORKDIR = Path(r"C:\davinci_automation")
for p in (INDIR, WORKDIR): p.mkdir(parents=True, exist_ok=True)

@APP.get("/health")
def health(): 
    return {"ok": True}

@APP.options("/process_upload")
def options_upload(): 
    return Response(status_code=200)

def _normalize_services(s: str) -> str:
    """
    Accepts:
      - raw string: 'Stage 0, DPF OFF, EGR Off'
      - JSON array: '["Stage 0","DPF OFF","EGR Off"]'
      - dict-object-y strings -> comma-join values
    Returns a flat string for the automation script.
    """
    if not s:
        return ""
    s = s.strip()
    try:
        if s.startswith("[") or s.startswith("{"):
            obj = json.loads(s)
            if isinstance(obj, list):
                return ", ".join(str(x) for x in obj)
            if isinstance(obj, dict):
                return ", ".join(str(v) for v in obj.values())
    except Exception:
        pass
    return s

@APP.post("/process_upload")
async def process_upload(
    file: UploadFile = File(...),
    filename: str = Form("input.bin"),
    brand: str = Form(...),
    ecu: str = Form(...),
    services: str = Form(""),   # <- now accepted and forwarded
    task_id: str = Form(None),
):
    brand_clean = (brand or "").strip()
    ecu_clean   = (ecu or "").strip()
    services_norm = _normalize_services(services)
    task_id_in = (task_id or "").strip()

    # persist BIN with its exact original name
    original_name = Path(file.filename or filename or "input.bin").name
    bin_path = INDIR / original_name
    with open(bin_path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    # build command
    cmd = [
        PYTHON, SCRIPT,
        "--exe", EXE,
        "--input", str(bin_path),
        "--brand", brand_clean,
        "--ecu", ecu_clean,
        "--services", services_norm,        # <- pass services
    ]

    # echo inputs to server console for debugging
    print(f"[agent] exe={EXE}")
    print(f"[agent] input={bin_path}")
    print(f"[agent] brand={brand_clean} | ecu={ecu_clean}")
    print(f"[agent] services={services_norm}")
    print(f"[agent] task_id={task_id_in}")
    sys.stdout.flush()

    r = subprocess.run(cmd, cwd=str(WORKDIR), capture_output=True, text=True)
    ok = (r.returncode == 0)

    # Extract saved path from the automation script's stdout/stderr
    saved_path = None
    out = r.stdout or ""
    err = r.stderr or ""
    for pat in (r"SAVED_PATH:(?P<p>.+)", r"SAVED:(?P<p>.+)", r"SAVING_TO:(?P<p>.+)"):
        m = re.search(pat, out) or re.search(pat, err)
        if m:
            saved_path = m.group("p").strip().strip('"')
            break
    print(f"[agent] saved_path={saved_path}")
    sys.stdout.flush()

    # If task_id and saved_path are present, notify backend
    save_reply_status = None
    save_reply_body = None
    if task_id_in and saved_path:
        try:
            with open(saved_path, "rb") as f:
                b64 = base64.b64encode(f.read()).decode("ascii")

            payload = {
                "task_id": str(task_id_in),
                "saved_path": saved_path,   # optional, for extension only
                "file_b64": b64,
            }
            resp = requests.post(
                "https://stagingbackend.guru-host.co.uk/api/davinci/save_reply",
                json=payload,
                timeout=30,
            )
            print(resp.status_code, resp.text)
        except Exception as e:
            print(f"[agent] save_reply error: {e}")
    else:
        print("[agent] save_reply skipped (missing task_id or saved_path)")

    return {
        "ok": ok,
        "opened_file": str(bin_path),
        "saved_path": saved_path,
        "metadata": {
            "brand": brand_clean,
            "ecu": ecu_clean,
            "services": services_norm,   # <- include in API response
            "original_filename": original_name,
        },
        "task_id": task_id_in,
        "save_reply": {"status": save_reply_status, "body": save_reply_body},
        "stdout_tail": (r.stdout or "")[-2000:],
        "stderr_tail": (r.stderr or "")[-2000:],
        "note": "Open-only. Script selects Brand→ECU, feeds BIN, and applies requested toggles.",
    }