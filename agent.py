# agent.py — polling worker: fetch tasks → download BINs → run DaVinci automation → upload results

from pathlib import Path
import subprocess
import json
import sys
import re
import requests
import base64
import os
import time
import logging

# Paths and configuration
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

API_STAGING_FILES_URL = "https://backend-staging.ecutech.gr/api/davinci/files"
API_PRODUCTION_FILES_URL = "https://backend.ecutech.gr/api/davinci/files"

API_STAGING_SAVE_REPLY_URL = "https://backend-staging.ecutech.gr/api/davinci/save_reply"
API_PRODUCTION_SAVE_REPLY_URL = "https://backend.ecutech.gr/api/davinci/save_reply"

API_STAGING_FAILURE_REPLY_URL = "https://backend-staging.ecutech.gr/api/davinci/failure"
API_PRODUCTION_FAILURE_REPLY_URL = "https://backend.ecutech.gr/api/davinci/failure"

logging.basicConfig(
    filename=str(Path("C:/davinci_automation/agent.log")),
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
)

def _select_failure_url(on_dev):
    """
    Pick the correct failure URL based on `on_dev` flag from the task.
    on_dev = 1 (or truthy '1') → staging; otherwise → production.
    """
    flag = str(on_dev).strip()
    if flag == "1":
        return API_STAGING_FAILURE_REPLY_URL
    return API_PRODUCTION_FAILURE_REPLY_URL

def _post_failure(task_id, message: str, on_dev):
    """POST a failure reason back to the backend for the given task_id."""
    if not task_id or not (message or "").strip():
        logging.info("failure skipped (missing task_id or message)")
        return

    payload = {
        "task_id": str(task_id),
        "message": message,
    }
    try:
        url = _select_failure_url(on_dev)
        logging.info(f"Posting failure to {url} for task_id={task_id} | message={message}")
        print(f"[AGENT] Posting failure for task_id={task_id}: {message}", flush=True)
        resp = requests.post(url, json=payload, timeout=30)
        print(f"[AGENT] failure response: {resp.status_code}", flush=True)
        logging.info(f"failure response: {resp.status_code} {resp.text[:500]}")
    except Exception as e:
        logging.error(f"failure post error for task_id={task_id}: {e}")

def _select_save_reply_url(on_dev):
    """
    Pick the correct save_reply URL based on `on_dev` flag from the task.
    on_dev = 1 (or truthy '1') → staging; otherwise → production.
    """
    flag = str(on_dev).strip()
    if flag == "1":
        return API_STAGING_SAVE_REPLY_URL
    return API_PRODUCTION_SAVE_REPLY_URL


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

def _extract_automation_error(out: str, err: str) -> str:
    """Pull a structured AUTMATION_ERROR line out of automation output."""
    combined = (out or "") + "\n" + (err or "")
    m = re.search(r"AUTOMATION_ERROR:\s*(.+)", combined)
    if m:
        return m.group(1).strip()
    # fallback: first 500 chars of combined output
    return combined.strip()[:500] if combined.strip() else "DaVinci automation failed (no further details)."

def _run_automation(bin_path: Path, brand: str, ecu: str, services):
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

    print(f"[AGENT] Running automation for {bin_path} | brand={brand_clean} ecu={ecu_clean} services={services_norm}", flush=True)
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
    print(f"[AGENT] automation returncode={r.returncode}", flush=True)

    saved_path = None
    for pat in (r"SAVED_PATH:(?P<p>.+)", r"SAVED:(?P<p>.+)", r"SAVING_TO:(?P<p>.+)"):
        m = re.search(pat, out) or re.search(pat, err)
        if m:
            saved_path = m.group("p").strip().strip('"')
            break

    if saved_path:
        print(f"[AGENT] Detected saved_path: {saved_path}", flush=True)
    else:
        print("[AGENT] No saved_path detected in automation output", flush=True)

    if not ok:
        logging.error(f"Automation failed (code={r.returncode}) for {bin_path}")
    if not saved_path:
        logging.warning(f"Automation did not report saved path for {bin_path}")

    error_message = None
    if (not ok) or (not saved_path):
        error_message = _extract_automation_error(out, err)

    return ok, saved_path, out, err, error_message


def _post_save_reply(task_id, saved_path: str, on_dev):
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
        print(f"[AGENT] Uploading result for task_id={task_id} from {saved_path}", flush=True)
        url = _select_save_reply_url(on_dev)
        logging.info(f"Posting save_reply to {url} for task_id={task_id}")
        resp = requests.post(url, json=payload, timeout=30)
        print(f"[AGENT] save_reply response: {resp.status_code}", flush=True)
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
        on_dev = task.get("on_dev") or ""

        on_dev = "1" if str(on_dev).strip() == "1" else "0"

        # Early validation for missing brand/ECU
        if not brand.strip() or not ecu.strip():
            msg = f"Missing brand or ECU for task {task_id} (brand='{brand}', ecu='{ecu}')"
            logging.error(msg)
            print(f"[AGENT] {msg}", flush=True)
            _post_failure(task_id, msg, on_dev)
            return

        print(f"[AGENT] Processing task_id={task_id} | file_name={file_name} | brand={brand} | ecu={ecu}", flush=True)
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
        print(f"[AGENT] Downloaded file to {bin_path}", flush=True)

        ok, saved_path, out, err, error_message = _run_automation(bin_path, brand, ecu, services)
        print(f"[AGENT] Automation finished for task_id={task_id} | ok={ok} | saved_path={saved_path}", flush=True)

        if ok and saved_path:
            _post_save_reply(task_id, saved_path, on_dev)
            print(f"[AGENT] Completed task_id={task_id}", flush=True)
        else:
            logging.error(f"Task {task_id}: automation failed or no saved_path")
            print(f"[AGENT] Task {task_id} FAILED (ok={ok}, saved_path={saved_path})", flush=True)

            failure_reason = error_message or f"DaVinci automation failed or did not produce a saved file (ok={ok}, saved_path={saved_path})."
            _post_failure(task_id, failure_reason, on_dev)
            
    except Exception as e:
        logging.error(f"Unhandled error while processing task {task}: {e}")


def _fetch_all_tasks():
    """
    Poll both staging and production /api/davinci/files endpoints.
    Returns a single flat list of task dicts.
    If a task does not contain `on_dev`, it will be set based on the source
    (staging → '1', production → '0').
    """
    all_tasks = []

    sources = [
        ("staging", API_STAGING_FILES_URL, "1"),
        ("production", API_PRODUCTION_FILES_URL, "0"),
    ]

    for label, url, default_on_dev in sources:
        try:
            logging.info(f"Polling {label} files URL: {url}")
            resp = requests.get(url, timeout=30)
            resp.raise_for_status()

            try:
                data = resp.json()
            except Exception as e:
                logging.error(
                    f"Failed to decode JSON from {label} files API: {e} | body={resp.text[:500]}"
                )
                data = []

            if not data:
                logging.info(f"{label}: no tasks returned.")
                continue

            logging.info(f"{label}: received {len(data)} task(s)")
            for task in data:
                # If backend didn't set on_dev explicitly, infer from source.
                if "on_dev" not in task or (task["on_dev"] in (None, "", 0, 1) and str(task["on_dev"]).strip() == ""):
                    task["on_dev"] = default_on_dev
                all_tasks.append(task)

        except Exception as e:
            logging.error(f"{label} files polling error: {e}")

    return all_tasks


def poll_forever(interval_seconds: int = 120):
    """Main loop: poll the backend every `interval_seconds` seconds.

    For each returned task, download the file, run DaVinci automation,
    and push the result back via /api/davinci/save_reply.
    """
    logging.info(
        f"Starting DaVinci polling worker. Interval={interval_seconds}s, files_dir={INDIR}"
    )

    while True:
        print("[AGENT] --- Poll cycle start ---", flush=True)
        try:
            tasks = _fetch_all_tasks()
            print(f"[AGENT] Polling done, got {len(tasks)} task(s) (staging+production)", flush=True)

            if not tasks:
                logging.info("No tasks returned.")
                print("[AGENT] No tasks returned this cycle.", flush=True)
            else:
                logging.info(f"Received {len(tasks)} task(s)")
                print(f"[AGENT] Processing {len(tasks)} task(s) from queue", flush=True)
                for task in tasks:
                    process_task(task)
        except Exception as e:
            logging.error(f"Top-level polling error: {e}")
            print(f"[AGENT] Top-level polling error: {e}", flush=True)

        time.sleep(interval_seconds)
        print(f"[AGENT] Sleeping {interval_seconds} seconds before next poll...", flush=True)


if __name__ == "__main__":
    print(">>> AGENT: STARTED. Polling for tasks...", flush=True)
    try:
        poll_forever()
    except KeyboardInterrupt:
        print("Stopped by user.")