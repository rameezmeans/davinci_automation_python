# agent.py — receive upload → run full DaVinci automation → apply services → save mod file → upload result

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
        "--services", services_norm,        
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
# agent.py — polling worker: fetch tasks → download BINs → run DaVinci automation → upload results

from pathlib import Path
import subprocess, shutil, json, sys, re, requests, base64, os, time, logging

EXE     = r"C:\Program Files\DAVINCI\davinci.exe"
SCRIPT  = r"C:\Program Files\DAVINCI\davinci_automation.py"
PYTHON  = r"C:\davinci_venv\Scripts\python.exe"
INDIR   = Path(r"C:\ecu_files\original")
WORKDIR = Path(r"C:\davinci_automation")

for p in (INDIR, WORKDIR):
    try:
        p.mkdir(parents=True, exist_ok=True)
    except Exception:
        pass

API_FILES_URL = "https://stagingbackend.guru-host.co.uk/api/davinci/files"
API_SAVE_REPLY_URL = "https://stagingbackend.guru-host.co.uk/api/davinci/save_reply"

logging.basicConfig(
    filename=str(Path("C:/davinci_automation/agent.log")),
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
)


def _normalize_services(s):
    """Normalize the `services` value from the API into a flat string.

    Accepts list/dict/string/JSON-string and returns something like:
      'Stage 0, DPF OFF, EGR OFF'
    """
    if s is None:
        return ""
    if isinstance(s, list):
        return ", ".join(str(x) for x in s)
    if isinstance(s, dict):
        return ", ".join(str(v) for v in s.values())

    s = str(s).strip()
    if not s:
        return ""

    # Try to decode JSON if the backend sent a JSON-encoded string
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


def _download_file(url: str, dest: Path) -> bool:
    """Download a file from `url` to `dest`. Returns True on success."""
    try:
        logging.info(f"Downloading {url} -> {dest}")
        with requests.get(url, stream=True, timeout=60) as r:
            r.raise_for_status()
            with open(dest, "wb") as f:
                for chunk in r.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
        return True
    except Exception as e:
        logging.error(f"Download failed for {url}: {e}")
        return False


def _run_automation(bin_path: Path, brand: str, ecu: str, services: str):
    """Call davinci_automation.py with the given parameters.

    Returns (ok, saved_path, stdout, stderr).
    """
    brand_clean = (brand or "").strip()
    ecu_clean = (ecu or "").strip()
    services_norm = _normalize_services(services)

    cmd = [
        PYTHON, SCRIPT,
        "--exe", EXE,
        "--input", str(bin_path),
        "--brand", brand_clean,
        "--ecu", ecu_clean,
        "--services", services_norm,
    ]

    logging.info(
        f"Running automation for {bin_path} | "
        f"brand={brand_clean} ecu={ecu_clean} services={services_norm}"
    )

    r = subprocess.run(cmd, cwd=str(WORKDIR), capture_output=True, text=True)
    ok = (r.returncode == 0)

    out = r.stdout or ""
    err = r.stderr or ""
    logging.info(f"automation stdout (tail): {out[-500:]}")
    logging.info(f"automation stderr (tail): {err[-500:]}")

    saved_path = None
    for pat in (r"SAVED_PATH:(?P<p>.+)", r"SAVED:(?P<p>.+)", r"SAVING_TO:(?P<p>.+)"):
        m = re.search(pat, out) or re.search(pat, err)
        if m:
            saved_path = m.group("p").strip().strip('"')
            break

    if not ok:
        logging.error(f"Automation failed (code={r.returncode}) for {bin_path}")
    if not saved_path:
        logging.warning(f"Automation did not report saved path for {bin_path}")

    return ok, saved_path, out, err


def _post_save_reply(task_id, saved_path: str):
    """POST the modified file back to the backend as base64 for the given task_id."""
    if not (task_id and saved_path):
        logging.info("save_reply skipped (missing task_id or saved_path)")
        return
    try:
        with open(saved_path, "rb") as f:
            b64 = base64.b64encode(f.read()).decode("ascii")

        payload = {
            "task_id": str(task_id),
            "saved_path": saved_path,
            "file_b64": b64,
        }
        logging.info(f"Posting save_reply for task_id={task_id}")
        resp = requests.post(API_SAVE_REPLY_URL, json=payload, timeout=30)
        logging.info(f"save_reply response: {resp.status_code} {resp.text[:500]}")
    except Exception as e:
        logging.error(f"save_reply error for task_id={task_id}: {e}")


def process_task(task: dict):
    """Process a single task from the /api/davinci/files endpoint."""
    try:
        task_id = task.get("task_id")
        file_url = task.get("file")
        file_name = task.get("file_name") or "input.bin"
        brand = task.get("brand") or ""
        ecu = task.get("ecu") or ""
        services = task.get("services") or ""

        logging.info(
            f"Processing task_id={task_id}, file={file_url}, file_name={file_name}"
        )

        if not file_url:
            logging.error(f"Task {task_id}: missing file URL, skipping")
            return

        bin_path = INDIR / Path(file_name).name
        if not _download_file(file_url, bin_path):
            logging.error(f"Task {task_id}: download failed, skipping automation")
            return

        ok, saved_path, out, err = _run_automation(bin_path, brand, ecu, services)

        if ok and saved_path:
            _post_save_reply(task_id, saved_path)
        else:
            logging.error(f"Task {task_id}: automation failed or no saved_path")
    except Exception as e:
        logging.error(f"Unhandled error while processing task {task}: {e}")


def poll_forever(interval_seconds: int = 120):
    """Main loop: poll the backend every `interval_seconds` seconds.

    For each returned task, download the file, run DaVinci automation,
    and push the result back via /api/davinci/save_reply.
    """
    logging.info(
        f"Starting DaVinci polling worker. Interval={interval_seconds}s, files_dir={INDIR}"
    )

    while True:
        try:
            logging.info(f"Polling {API_FILES_URL}")
            resp = requests.get(API_FILES_URL, timeout=30)
            resp.raise_for_status()

            try:
                tasks = resp.json()
            except Exception as e:
                logging.error(
                    f"Failed to decode JSON from files API: {e} | body={resp.text[:500]}"
                )
                tasks = []

            if not tasks:
                logging.info("No tasks returned.")
            else:
                logging.info(f"Received {len(tasks)} task(s)")
                for task in tasks:
                    process_task(task)
        except Exception as e:
            logging.error(f"Top-level polling error: {e}")

        time.sleep(interval_seconds)


if __name__ == "__main__":
    try:
        poll_forever()
    except KeyboardInterrupt:
        print("Stopped by user.")