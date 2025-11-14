# davinci_automation.py — Launch/attach DaVinci → select BRAND → double-click ECU → wait for Open dialog → STOP
# deps: pip install pywinauto

import argparse, sys, time, subprocess
from pathlib import Path
from pywinauto.application import Application
from pywinauto import Desktop
from pywinauto.timings import TimeoutError as UIATimeout
from pywinauto.keyboard import send_keys
from pywinauto.controls.uiawrapper import UIAWrapper
from pywinauto.uia_defines import IUIA
from pywinauto import clipboard
import logging
import threading
logging.basicConfig(filename=str(Path("C:/davinci_automation/davinci_automation.log")),
                    level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")


MAIN_TITLES = ["DaVinci DPF EGR DTC", "DaVinci"]
# Common Save dialog titles (multi-locale)
OPEN_HINTS  = ["Open", "Öffnen", "Abrir", "Открытие", "Open File", "Select file", "Original Files"]
SAVE_HINTS  = ["Save Mod File", "Save As", "Save", "Speichern", "Guardar", "Сохранение", "Save file", "Save Modified File"]
# Helper to detect top-level file dialogs even if not a dialog or with custom titles

# Treat multiple brand inputs as VAG (DaVinci groups them under one node)
VAG_BRANDS = {"volkswagen", "vw", "audi", "seat", "skoda"}

def effective_brand(brand: str) -> str:
    """Map various inputs to the brand label used by DaVinci's tree.
    For VAG cluster (Volkswagen/Audi/SEAT/Skoda/VW) return 'VAG'.
    Otherwise return the original string unchanged.
    """
    b = (brand or "").strip().lower()
    if b in VAG_BRANDS:
        return "VAG"
    return brand

def launch_if_needed(exe: Path):
    if not exe.exists():
        raise FileNotFoundError(f"DaVinci not found: {exe}")
    if any(Desktop(backend="uia").windows(title_re=t) for t in MAIN_TITLES):
        return
    subprocess.Popen([str(exe)], shell=False, cwd=str(exe.parent))
    time.sleep(2)
    logging.info("launched/attached")

def connect_window(timeout=25):
    t0 = time.time()
    while time.time() - t0 < timeout:
        for t in MAIN_TITLES:
            try:
                app = Application(backend="uia").connect(title_re=t, timeout=2)
                win = app.window(title_re=t)
                win.set_focus()
                return app, win
            except Exception:
                pass
        time.sleep(0.4)
    raise RuntimeError("DaVinci window not found. Match elevation (Admin vs non-Admin).")

def get_tree(win):
    try:
        tr = win.child_window(control_type="Tree")
        if tr.exists():
            return tr.wrapper_object()
    except Exception:
        pass
    for c in win.descendants():
        try:
            if c.element_info.control_type == "Tree":
                return c.wrapper_object()
        except Exception:
            pass
    # Fallback: focus right pane so keystrokes can work later
    try:
        win.set_focus()
        r = win.rectangle()
        x = int(r.left + r.width() * 0.83); y = r.top + 120
        win.click_input(coords=(x - r.left, y - r.top))
        return None
    except Exception:
        pass
    raise RuntimeError("Brand/ECU list not found. Ensure main screen visible; Windows scaling 100%.")

def select_brand_ecu_ui(tree, brand: str, ecu: str):
    logging.info(f"selecting brand={brand} ecu={ecu}")
    eff_brand = effective_brand(brand)
    b = eff_brand.strip().lower()
    e = ecu.strip().lower()
    if not b or not e:
        raise ValueError("brand and ecu required")

    # Expand roots
    try:
        for r in tree.roots():
            try: r.expand()
            except Exception: pass
    except Exception:
        pass

    # Locate brand
    brand_node = None
    for n in tree.descendants():
        try:
            txt = n.window_text().strip().lower()
            if txt == b or b in txt:
                brand_node = n; break
        except Exception:
            pass
    if not brand_node:
        raise RuntimeError(f"Brand not found: {eff_brand}")

    try:
        brand_node.select(); brand_node.expand()
    except Exception:
        pass
    time.sleep(0.2)

    # Locate ECU under brand
    ecu_node = None
    for d in brand_node.descendants():
        try:
            txt = d.window_text().strip().lower()
            if txt == e or e in txt:
                ecu_node = d; break
        except Exception:
            pass
    if not ecu_node:
        raise RuntimeError(f"ECU not found under {eff_brand}: {ecu}")

    # Double-click ECU to trigger the Open dialog
    try:
        ecu_node.double_click_input()
    except Exception:
        # Fallback to coordinate-based double-click
        rect = ecu_node.rectangle()
        cx = int((rect.left + rect.right) / 2)
        cy = int((rect.top + rect.bottom) / 2)
        ecu_node.double_click_input(coords=(cx - rect.left, cy - rect.top))
    time.sleep(0.7)

    # After double-click, a confirmation popup may appear; press Enter to dismiss it
    time.sleep(0.5)
    try:
        send_keys("{ENTER}")
    except Exception:
        pass

    # Sometimes a second confirmation appears or focus is elsewhere; press Enter again and try closing info dialog.
    time.sleep(0.3)
    try:
        send_keys("{ENTER}")
    except Exception:
        pass
    # Attempt to close any 'Info' blocker that may have appeared
    try:
        maybe_close_info_dialog(timeout=3)
    except Exception:
        pass
    logging.info("brand/ecu selection done")

def select_brand_ecu_keys(win, brand: str, ecu: str):
    logging.info(f"selecting brand={brand} ecu={ecu}")
    win.set_focus()
    send_keys("{HOME}")
    time.sleep(0.1)
    eff_brand = effective_brand(brand)
    send_keys(eff_brand, with_spaces=True, pause=0.02)
    send_keys("{ENTER}")
    send_keys("{RIGHT}")
    time.sleep(0.2)
    send_keys(ecu, with_spaces=True, pause=0.02)
    # Double-click via keyboard: ENTER twice
    send_keys("{ENTER}")
    time.sleep(0.15)
    send_keys("{ENTER}")
    time.sleep(0.2)
    # Dismiss the first info dialog that appears after ECU activation
    try:
        maybe_close_info_dialog(timeout=3)
    except Exception:
        pass
    logging.info("brand/ecu selection done")

def maybe_close_info_dialog(timeout=6):
    """Close blocking info dialogs (e.g., 'BDM READ IS REQUIRED') so the Open dialog can appear."""
    t0 = time.time()
    while time.time() - t0 < timeout:
        # Common 'Info' dialog titles
        for title in ["INFO", "Info", "Information"]:
            try:
                dlg = Desktop(backend="uia").window(title=title, control_type="Window")
                if dlg.exists(timeout=0.2):
                    try:
                        dlg.set_focus()
                    except Exception:
                        pass
                    # Click OK if present, otherwise press Enter
                    try:
                        ok = dlg.child_window(title_re="^(OK|Ok|ok)$", control_type="Button")
                        if ok.exists(timeout=0.2):
                            ok.click_input()
                        else:
                            send_keys("{ENTER}")
                    except Exception:
                        send_keys("{ENTER}")
                    time.sleep(0.2)
                    return True
            except Exception:
                pass
        # Fallback: scan any dialog containing specific message text
        try:
            for w in Desktop(backend="uia").windows():
                try:
                    txt = (w.window_text() or "").lower()
                except Exception:
                    txt = ""
                if "bdm read is required" in txt:
                    try:
                        w.set_focus()
                    except Exception:
                        pass
                    try:
                        ok = w.child_window(title_re="^(OK|Ok|ok)$", control_type="Button")
                        if ok.exists(timeout=0.2):
                            ok.click_input()
                        else:
                            send_keys("{ENTER}")
                    except Exception:
                        send_keys("{ENTER}")
                    time.sleep(0.2)
                    return True
        except Exception:
            pass
        time.sleep(0.2)
    return False


# Helper: type folder and filename into the Open dialog's File name input, then submit
def type_folder_and_filename(folder: str, filename: str):
    """Assumes focus is already in the 'File name' input. Types full path, then TAB and ENTER."""
    full_path = str(Path(folder) / filename)
    try:
        # Clear any existing text first
        send_keys("^a{BACKSPACE}")
        time.sleep(0.05)
    except Exception:
        pass
    send_keys(full_path, with_spaces=True)
    time.sleep(0.05)
    # TAB to move focus to Open button (or equivalent), then ENTER
    send_keys("{TAB}{ENTER}")


# Wait for Save dialog and prepend modified_dir to filename, then submit (background threadable)
def wait_for_save_dialog_then_prepend_modified(modified_dir=r"C:\ecu_files\modified", timeout=600):
    """
    Wait until a Save-like dialog appears (same common file dialog, but with a prefilled 'File name'),
    then prepend `modified_dir` to the current filename and submit (TAB then ENTER).
    Runs non-blocking if started in a background thread.
    """
    t0 = time.time()
    while time.time() - t0 < timeout:
        try:
            dlg = find_any_open_dialog()
            if dlg is None:
                time.sleep(0.2)
                continue
            # Try to locate the 'File name' edit (usually the last Edit control)
            edit = None
            try:
                edits = dlg.descendants(control_type="Edit")
                if edits:
                    edit = edits[-1]
            except Exception:
                edit = None
            if not edit:
                time.sleep(0.2)
                continue
            # Focus and read current value
            try:
                edit.set_focus()
            except Exception:
                pass
            time.sleep(0.05)
            # Copy text from the edit; fall back to .get_value() / window_text()
            current_text = ""
            try:
                send_keys("^a^c")
                time.sleep(0.05)
                current_text = (clipboard.GetData() or "").strip()
            except Exception:
                pass
            if not current_text:
                try:
                    current_text = (edit.get_value() or "").strip()
                except Exception:
                    try:
                        current_text = (edit.window_text() or "").strip()
                    except Exception:
                        current_text = ""
            # If there's no filename yet, keep watching
            if not current_text:
                time.sleep(0.2)
                continue
            # Use only the base name part
            try:
                fname = Path(current_text).name
            except Exception:
                fname = current_text
            if not fname:
                time.sleep(0.2)
                continue
            # Build the new full path in the modified folder
            full_out = str(Path(modified_dir) / fname)
            # Replace the text and submit
            try:
                send_keys("^a{BACKSPACE}")
                time.sleep(0.05)
            except Exception:
                pass
            send_keys(full_out, with_spaces=True)
            time.sleep(0.05)
            send_keys("{TAB}{ENTER}")
            logging.info(f"Save dialog handled: wrote {full_out}")
            return True
        except Exception:
            # swallow and keep watching
            time.sleep(0.2)
    logging.info("Save dialog watcher timed out without action.")
    return False

def invoke_open_dialog(win):
    """Try multiple ways to trigger the Open-file dialog from the main window/menu."""
    try:
        win.set_focus()
    except Exception:
        pass

    # 1) Try UIA menu search for 'File' -> 'Open'
    try:
        # Some DaVinci builds expose a MenuBar with MenuItems
        menubars = win.descendants(control_type="MenuBar")
        for mb in menubars:
            try:
                file_menu = None
                for mi in mb.descendants(control_type="MenuItem"):
                    name = (mi.window_text() or "").lower()
                    if name in ["file", "&file", "archivo", "datei", "fichier", "файл", "文件", "arquivo"]:
                        file_menu = mi
                        break
                if file_menu:
                    try:
                        file_menu.select()
                    except Exception:
                        file_menu.click_input()
                    time.sleep(0.2)
                    # find Open-like item
                    for mi in file_menu.descendants(control_type="MenuItem"):
                        n = (mi.window_text() or "").lower()
                        if any(k in n for k in ["open", "öffnen", "abrir", "ouvrir", "открыть", "打开"]):
                            try:
                                mi.select()
                            except Exception:
                                mi.click_input()
                            time.sleep(0.6)
                            return True
            except Exception:
                continue
    except Exception:
        pass

    # 2) Keyboard menu accelerator: Alt+F then O
    try:
        send_keys("%fo")
        time.sleep(0.6)
    except Exception:
        pass

    # 3) Ctrl+O as last resort
    try:
        send_keys("^o")
        time.sleep(0.6)
    except Exception:
        pass
    return False

def nudge_open_dialog(win, retries=3):
    """Proactively try to surface the Open dialog if the app swallowed the first double-click.
    Sends ENTER (confirm), CTRL+O (common Open shortcut), and attempts to dismiss any 'Info' dialog."""
    logging.info("nudging for Open dialog")
    try:
        win.set_focus()
    except Exception:
        pass
    for _ in range(max(1, retries)):
        try:
            # Confirm/advance any blocking prompt
            send_keys("{ENTER}")
            time.sleep(0.25)
            # Try common File→Open shortcut
            send_keys("^o")
            try:
                invoke_open_dialog(win)
            except Exception:
                pass
        except Exception:
            pass
        # Close possible info dialog
        try:
            maybe_close_info_dialog(timeout=0.6)
        except Exception:
            pass
        time.sleep(0.6)


# Helper to find an embedded Open dialog within the DaVinci window
def find_open_dialog_within(win):
    """Some builds host the file picker as a child Pane/Window inside the main window.
    Search within the DaVinci window for a container having an Edit ('File name') and an 'Open' button."""
    try:
        # First, any child window/pane that looks like a file picker
        candidates = []
        for ct in ("Window", "Pane", "Group"):
            try:
                candidates.extend(win.descendants(control_type=ct))
            except Exception:
                continue

        # Heuristic scoring
        for c in candidates:
            try:
                # must contain at least one Edit
                edits = c.descendants(control_type="Edit")
                if not edits:
                    continue

                # look for an 'Open' button sibling OR a combobox/toolbar that looks like the address bar
                has_open_btn = False
                for btn in c.descendants(control_type="Button"):
                    name = (btn.window_text() or "").lower()
                    if any(k in name for k in ["open", "öffnen", "abrir", "ouvrir", "открыть", "打开"]):
                        has_open_btn = True
                        break

                # also accept if we find a 'File name' edit or filename label nearby
                has_filename = False
                for e in edits:
                    nm = (e.window_text() or "").lower()
                    if any(k in nm for k in ["file name", "dateiname", "nombre", "nome", "nome do arquivo", "nome del file", "имя файла"]):
                        has_filename = True
                        break

                if has_open_btn or has_filename:
                    return c.wrapper_object()
            except Exception:
                continue
    except Exception:
        pass
    return None

def find_any_open_dialog():
    """Return a wrapper for any likely Open-file dialog."""
    # Known titles first
    for hint in OPEN_HINTS:
        try:
            d = Desktop(backend="uia").window(title_re=hint, control_type="Window")
            if d.exists(timeout=0.2):
                return d.wrapper_object()
        except Exception:
            pass

    # Generic top-level detector
    dlg = find_top_level_file_dialog()
    if dlg is not None:
        return dlg

    # Heuristic over visible windows: must have Edit and an Open-like button
    try:
        for w in Desktop(backend="uia").windows():
            try:
                if not w.is_visible():
                    continue
            except Exception:
                pass
            try:
                edits = w.descendants(control_type="Edit")
                if not edits:
                    continue
                for btn in w.descendants(control_type="Button"):
                    name = (btn.window_text() or "").lower()
                    if any(k in name for k in ["open", "öffnen", "abrir", "ouvrir", "открыть", "打开"]):
                        return w.wrapper_object()
            except Exception:
                continue
    except Exception:
        pass

    # Last resort: any visible window containing a 'File name' text label
    try:
        for w in Desktop(backend="uia").windows():
            try:
                if not w.is_visible():
                    continue
            except Exception:
                pass
            try:
                for t in w.descendants(control_type="Text"):
                    txt = (t.window_text() or "").lower()
                    if any(k in txt for k in ["file name", "dateiname", "nombre", "nome", "имя файла"]):
                        return w.wrapper_object()
            except Exception:
                continue
    except Exception:
        pass

    return None


# --- Save/Open dialog helpers ---
def find_top_level_file_dialog():
    """Detect a common file dialog hosted as a top-level window (UIA or legacy)."""
    # Legacy common dialog (#32770)
    try:
        dlg = Desktop(backend="win32").window(class_name="#32770")
        if dlg.exists(timeout=0.2):
            return dlg.wrapper_object()
    except Exception:
        pass
    # UIA windows that look like file pickers (have Edit and an Open/Save button)
    try:
        for w in Desktop(backend="uia").windows():
            try:
                edits = w.descendants(control_type="Edit")
            except Exception:
                edits = []
            if not edits:
                continue
            try:
                for b in w.descendants(control_type="Button"):
                    n = (b.window_text() or "").lower()
                    if any(k in n for k in ["open", "öffnen", "abrir", "ouvrir", "открыть", "打开",
                                            "save", "speichern", "guardar", "salvar"]):
                        return w.wrapper_object()
            except Exception:
                continue
    except Exception:
        pass
    return None

def _set_folder_via_address_bar(folder: Path):
    # Works in both Open/Save dialogs: Alt+D focuses address bar
    send_keys("%d")
    time.sleep(0.2)
    send_keys(str(folder), with_spaces=True)
    send_keys("{ENTER}")
    time.sleep(0.35)

def _accept_open_dialog(dlg) -> bool:
    """Activate Open without relying on Enter focus."""
    try:
        dlg.set_focus()
    except Exception:
        pass
    variants = ["Open", "&Open", "Öffnen", "Abrir", "Открыть"]
    try:
        for lab in variants:
            try:
                btn = dlg.child_window(title=lab, control_type="Button")
                if btn.exists():
                    w = btn.wrapper_object()
                    try:
                        if hasattr(w, "invoke"):
                            w.invoke(); return True
                    except Exception:
                        pass
                    try:
                        w.set_focus(); send_keys(" "); return True
                    except Exception:
                        pass
            except Exception:
                continue
        for b in dlg.descendants(control_type="Button"):
            try:
                t = (b.window_text() or "").strip().lower()
                if t.startswith("open") or t.startswith("öffnen") or t.startswith("abrir") or t.startswith("откры"):
                    w = b.wrapper_object()
                    try:
                        if hasattr(w, "invoke"):
                            w.invoke(); return True
                    except Exception:
                        pass
                    try:
                        w.set_focus(); send_keys(" "); return True
                    except Exception:
                        pass
            except Exception:
                continue
    except Exception:
        pass
    try:
        dlg.type_keys("%o"); return True
    except Exception:
        pass
    try:
        iui = IUIA()
        for _ in range(120):
            try:
                foc_el = iui.get_focused_element()
                if foc_el is not None:
                    foc = UIAWrapper(foc_el)
                    ctl = (foc.element_info.control_type or "")
                    name = (foc.window_text() or "").strip().lower()
                    if ctl == "Button" and (name.startswith("open") or name.startswith("öffnen")
                                            or name.startswith("abrir") or name.startswith("откры")):
                        send_keys(" "); return True
            except Exception:
                pass
            send_keys("{TAB}")
            time.sleep(0.05)
    except Exception:
        pass
    try:
        dlg.type_keys("{ENTER}"); return True
    except Exception:
        return False

def _locate_possible_open_dialog():
    for hint in OPEN_HINTS:
        try:
            dlg = Desktop(backend="uia").window(title_re=hint, control_type="Window")
            if dlg.exists(timeout=0.2):
                return dlg
        except Exception:
            pass
    try:
        for w in Desktop(backend="uia").windows():
            try:
                if w.is_dialog():
                    for b in w.descendants(control_type="Button"):
                        t = (b.window_text() or "").strip().lower()
                        if t.startswith("open") or t.startswith("öffnen") or t.startswith("abrir") or t.startswith("откры"):
                            return w
            except Exception:
                continue
    except Exception:
        pass
    try:
        dlg = Desktop(backend="win32").window(class_name="#32770")
        if dlg.exists(timeout=0.2):
            return dlg
    except Exception:
        pass
    return None

def wait_save_dialog(timeout=120):
    """Wait for DaVinci's Save/Save As dialog after user clicks Save."""
    t0 = time.time()
    while time.time() - t0 < timeout:
        for hint in SAVE_HINTS:
            try:
                dlg = Desktop(backend="uia").window(title_re=hint, control_type="Window")
                if dlg.exists(timeout=0.4):
                    return ("uia", dlg)
            except Exception:
                pass
        try:
            for w in Desktop(backend="uia").windows():
                try:
                    if w.is_dialog() and "open" not in (w.window_text() or "").lower():
                        if any(btn.window_text().strip().lower().startswith("save")
                               for btn in w.descendants(control_type="Button")):
                            return ("uia", w)
                except Exception:
                    pass
        except Exception:
            pass
        time.sleep(0.25)
    try:
        dlg = Desktop(backend="win32").window(class_name="#32770")
        if dlg.exists(timeout=1.5):
            return ("win32", dlg)
    except Exception:
        pass
    raise UIATimeout("Save dialog did not appear.")

def _read_filename_from_dialog_uia(dlg):
    try:
        e = dlg.child_window(auto_id="1148", control_type="Edit")
        if e.exists():
            return e.get_value()
    except Exception:
        pass
    try:
        edits = [c for c in dlg.descendants(control_type="Edit")]
        if edits:
            return edits[-1].get_value()
    except Exception:
        pass
    return None

def _wait_dialog_gone(dlg, timeout=15) -> bool:
    t0 = time.time()
    while time.time() - t0 < timeout:
        try:
            if not dlg.exists(timeout=0.2):
                return True
        except Exception:
            return True
        time.sleep(0.1)
    return False

def _best_filename_from_dialog(dlg, fallback_name: str) -> str:
    try:
        e = dlg.child_window(auto_id="1148", control_type="Edit")
        if e.exists():
            val = e.get_value() or ""
            return val.strip() or fallback_name
    except Exception:
        pass
    try:
        edits = [c for c in dlg.descendants(control_type="Edit")]
        if edits:
            try:
                val = edits[-1].get_value() or ""
                return val.strip() or fallback_name
            except Exception:
                pass
    except Exception:
        pass
    return fallback_name

def _accept_save_dialog(dlg) -> bool:
    """Ensure Save is activated without relying on Enter focus."""
    try:
        dlg.set_focus()
    except Exception:
        pass
    variants = ["Save", "&Save", "Speichern", "Guardar", "Сохранить", "Salvar"]
    try:
        for lab in variants:
            try:
                btn = dlg.child_window(title=lab, control_type="Button")
                if btn.exists():
                    w = btn.wrapper_object()
                    try:
                        if hasattr(w, "invoke"):
                            w.invoke(); return True
                    except Exception:
                        pass
                    try:
                        w.set_focus(); send_keys(" "); return True
                    except Exception:
                        pass
            except Exception:
                continue
        for b in dlg.descendants(control_type="Button"):
            try:
                t = (b.window_text() or "").strip().lower()
                if t.startswith("save"):
                    w = b.wrapper_object()
                    try:
                        if hasattr(w, "invoke"):
                            w.invoke(); return True
                    except Exception:
                        pass
                    try:
                        w.set_focus(); send_keys(" "); return True
                    except Exception:
                        pass
            except Exception:
                continue
    except Exception:
        pass
    try:
        dlg.type_keys("%s"); return True
    except Exception:
        pass
    try:
        iui = IUIA()
        for _ in range(120):
            try:
                foc_el = iui.get_focused_element()
                if foc_el is not None:
                    foc = UIAWrapper(foc_el)
                    ctl = (foc.element_info.control_type or "")
                    name = (foc.window_text() or "").strip().lower()
                    if ctl == "Button" and name.startswith("save"):
                        send_keys(" "); return True
            except Exception:
                pass
            send_keys("{TAB}")
            time.sleep(0.05)
    except Exception:
        pass
    return False

def save_to_folder_uia(dlg, target_folder: Path, filename_fallback: str):
    target_folder.mkdir(parents=True, exist_ok=True)
    try:
        _set_folder_via_address_bar(target_folder)
    except Exception:
        pass
    fname = _read_filename_from_dialog_uia(dlg) or filename_fallback
    try:
        edit = dlg.child_window(auto_id="1148", control_type="Edit")
        if edit.exists():
            e = edit.wrapper_object()
            e.set_focus(); e.select()
            e.type_keys(fname, with_spaces=True)
        else:
            edits = [c for c in dlg.descendants(control_type="Edit")]
            if edits:
                e = edits[-1].wrapper_object()
                e.set_focus(); e.select()
                e.type_keys(fname, with_spaces=True)
    except Exception:
        pass
    return _accept_save_dialog(dlg)

def save_to_folder_win32(dlg, target_folder: Path, filename_fallback: str):
    target_folder.mkdir(parents=True, exist_ok=True)
    try:
        _set_folder_via_address_bar(target_folder)
    except Exception:
        pass
    try:
        edit = dlg.child_window(class_name="Edit")
        if edit.exists():
            e = edit.wrapper_object()
            try:
                cur = e.window_text() or ""
            except Exception:
                cur = ""
            if not cur.strip():
                cur = filename_fallback
            e.set_focus(); e.type_keys("^a{BACKSPACE}")
            e.type_keys(cur, with_spaces=True)
            time.sleep(0.2)
            return _accept_save_dialog(dlg)
    except Exception:
        pass
    return False

def maybe_confirm_overwrite(timeout=8):
    t0 = time.time()
    while time.time() - t0 < timeout:
        for backend in ("uia", "win32"):
            try:
                desk = Desktop(backend=backend)
                for w in desk.windows():
                    title = (w.window_text() or "").lower()
                    if any(k in title for k in ["confirm save as", "overwrite", "replace", "bestätigen"]):
                        for yes in ["Yes", "&Yes", "Ja", "Sí", "Да"]:
                            try:
                                b = w.child_window(title=yes, control_type="Button") if backend=="uia" else w.child_window(title=yes)
                                if b.exists():
                                    b.wrapper_object().click_input()
                                    return True
                            except Exception:
                                pass
                        try:
                            w.type_keys("%y")
                        except Exception:
                            try: w.type_keys("{ENTER}")
                            except Exception: pass
                        return True
            except Exception:
                pass
        time.sleep(0.25)
    return False

def maybe_click_yes_popup(timeout=8):
    """
    Look for any popup/dialog with a 'Yes' button and click it.
    Use after double-clicking in DaVinci when it shows a confirmation.
    """
    t0 = time.time()
    while time.time() - t0 < timeout:
        for backend in ("uia", "win32"):
            try:
                desk = Desktop(backend=backend)
                for w in desk.windows():
                    title = (w.window_text() or "").strip()
                    # You can narrow this down later if needed
                    # For now: any small dialog with a Yes button
                    try:
                        for yes_label in ["Yes", "&Yes", "Ja", "Sí", "Да"]:
                            try:
                                btn = (
                                    w.child_window(title=yes_label, control_type="Button")
                                    if backend == "uia"
                                    else w.child_window(title=yes_label)
                                )
                                if btn.exists(timeout=0.2):
                                    btn.wrapper_object().click_input()
                                    logging.info(f"Clicked YES on popup: '{title}'")
                                    return True
                            except Exception:
                                continue
                    except Exception:
                        continue
            except Exception:
                pass
        time.sleep(0.25)
    logging.info("maybe_click_yes_popup: no YES popup detected within timeout.")
    return False


def focus_save_filename_edit(save_win):
    """Focus the 'File name' input of a Save dialog and report status."""
    try:
        edits = save_win.descendants(control_type="Edit")
    except Exception:
        edits = []
    if not edits:
        return None
    try:
        edit = edits[-1]
    except Exception:
        return None
    try:
        edit.set_focus()
    except Exception:
        pass
    # Tell logs and stdout explicitly that we're in the File name box
    logging.info("At Save dialog: focus moved to 'File name' input.")
    print("STATUS: Save dialog detected; focus is in File name.")
    return edit

# --- Robust filename edit locator ---
def find_filename_edit(dlg):
    """
    Return the UIA wrapper for the 'File name' EDIT control, trying the canonical locations:
      1) auto_id=1148 Edit
      2) auto_id=1148 ComboBox → child Edit
      3) the last Edit under any ComboBox that is near a 'File name' label
      4) fallback: the last visible Edit in the dialog
    """
    # 1) Canonical Edit with auto_id=1148
    try:
        e = dlg.child_window(auto_id="1148", control_type="Edit")
        if e.exists(timeout=0.2):
            return e
    except Exception:
        pass

    # 2) ComboBox with auto_id=1148 then inner Edit
    try:
        cb = dlg.child_window(auto_id="1148", control_type="ComboBox")
        if cb.exists(timeout=0.2):
            try:
                inner_edits = cb.descendants(control_type="Edit")
                if inner_edits:
                    return inner_edits[-1]
            except Exception:
                pass
    except Exception:
        pass

    # 3) Heuristic: find a 'File name' label and grab a nearby Edit/ComboBox→Edit
    try:
        labels = []
        for t in dlg.descendants(control_type="Text"):
            txt = (t.window_text() or "").strip().lower()
            if any(k in txt for k in ["file name", "dateiname", "nombre", "nome", "имя файла"]):
                labels.append(t)
        if labels:
            # pick the first visible label, search siblings/descendants for an Edit
            lab = labels[0]
            parent = lab.parent()
            if parent:
                try:
                    # prefer an Edit inside a ComboBox sibling
                    for cb in parent.descendants(control_type="ComboBox"):
                        try:
                            inner_edits = cb.descendants(control_type="Edit")
                            if inner_edits:
                                return inner_edits[-1]
                        except Exception:
                            continue
                    # otherwise, any Edit sibling
                    edits = parent.descendants(control_type="Edit")
                    if edits:
                        return edits[-1]
                except Exception:
                    pass
    except Exception:
        pass

    # 4) Last resort: the last Edit visible in the dialog
    try:
        edits = [e for e in dlg.descendants(control_type="Edit")]
        if edits:
            return edits[-1]
    except Exception:
        pass
    return None

# Helper: get current filename from the Save dialog's File name edit, robustly
def get_current_filename_from_edit(dlg) -> str:
    """
    Robustly read the current text in the 'File name' field.
    Priority: UIA value from the correct Edit → clipboard → window_text.
    Includes small waits and multiple attempts to handle slow dialogs.
    """
    edit = find_filename_edit(dlg)
    # Try to focus the exact edit
    if edit is not None:
        try:
            edit.set_focus()
        except Exception:
            pass

    raw = ""
    # Try UIA value first (most reliable, no keyboard required)
    for _ in range(3):
        if edit is not None:
            try:
                v = (edit.get_value() or "").strip()
                if v:
                    raw = v
                    break
            except Exception:
                pass
        time.sleep(0.05)

    # Clipboard fallback (Ctrl+A, Ctrl+C)
    if not raw:
        try:
            send_keys("^a^c")
            time.sleep(0.06)
            raw = (clipboard.GetData() or "").strip()
        except Exception:
            raw = ""

    # window_text final fallback
    if not raw and edit is not None:
        try:
            raw = (edit.window_text() or "").strip()
        except Exception:
            raw = ""

    # Normalize: keep only the last path segment (the filename)
    if raw:
        try:
            # strip any wrapping quotes that Windows sometimes adds
            if raw.startswith('"') and raw.endswith('"') and len(raw) > 1:
                raw = raw[1:-1]
            name_only = Path(raw).name
            return name_only
        except Exception:
            return raw
    return ""

# Helper: save into the modified dir using the existing filename from the Save dialog
def save_into_modified_dir_with_existing_name(dlg, modified_dir=r"C:\\ecu_files\\modified") -> str | None:
    """
    Assumes we are already at the Save dialog and focus is (or will be) in the 'File name' field.
    Reads the current filename, prepends modified_dir, types full absolute path, then TAB and ENTER.
    Returns the final absolute path string if we attempted to save, else None.
    """
    # Ensure focus is in the filename edit
    _ = focus_save_filename_edit(dlg)
    time.sleep(0.2)

    # Read current proposed filename
    current = get_current_filename_from_edit(dlg)
    try:
        fname_only = Path(current).name if current else "modified.bin"
    except Exception:
        fname_only = current or "modified.bin"

    out_dir = Path(modified_dir)
    try:
        out_dir.mkdir(parents=True, exist_ok=True)
    except Exception:
        pass
    final_path = str(out_dir / fname_only)

    # Replace text in the edit and submit
    # Prefer dialog-scoped typing, with fallback to global send_keys
    typed = False
    try:
        dlg.type_keys("^a{BACKSPACE}")
        time.sleep(0.05)
        dlg.type_keys(final_path, with_spaces=True)
        time.sleep(0.05)
        dlg.type_keys("{TAB}{ENTER}")
        typed = True
    except Exception:
        try:
            send_keys("^a{BACKSPACE}")
            time.sleep(0.05)
            send_keys(final_path, with_spaces=True)
            time.sleep(0.05)
            send_keys("{TAB}{ENTER}")
            typed = True
        except Exception:
            typed = False

    if typed:
        logging.info(f"Save Mod File: typed full path → {final_path}")
        print(f"SAVING_TO:{final_path}")
        return final_path
    else:
        logging.info("Save Mod File: failed to type into File name edit.")
        return None


#
# Helper: navigate via address bar to target folder, keep existing filename, then Save
def save_via_address_bar_using_existing_name(dlg, target_folder=r"C:\\ecu_files\\modified") -> str | None:
    """
    Jump to target_folder via address bar (Alt+D), then focus File name (Alt+N),
    read the filename using UIA (with clipboard fallback), and save via Alt+S.
    """
    out_dir = Path(target_folder)
    try:
        out_dir.mkdir(parents=True, exist_ok=True)
    except Exception:
        pass

    # 1) Address bar jump
    jumped = False
    try:
        dlg.type_keys('%d')
        time.sleep(0.2)
        dlg.type_keys(str(out_dir), with_spaces=True)
        dlg.type_keys('{ENTER}')
        time.sleep(0.45)
        jumped = True
    except Exception:
        try:
            send_keys('%d')
            time.sleep(0.2)
            send_keys(str(out_dir), with_spaces=True)
            send_keys('{ENTER}')
            time.sleep(0.45)
            jumped = True
        except Exception:
            jumped = False

    # 2) Focus File name via Alt+N and read it robustly
    try:
        dlg.type_keys('%n')
        time.sleep(0.15)
    except Exception:
        try:
            send_keys('%n')
            time.sleep(0.15)
        except Exception:
            pass

    fname_only = get_current_filename_from_edit(dlg).strip()
    if not fname_only:
        # As a last resort, fall back to a safe default
        fname_only = "modified.bin"

    # Emit a precise diagnostic (consumed by agent.py if needed)
    print(f"RAW_FILENAME:{fname_only}")

    final_path = str(out_dir / fname_only)

    # 3) Save via Alt+S (do not retype name; keep whatever is there)
    saved = False
    try:
        dlg.type_keys('%s')
        saved = True
    except Exception:
        try:
            send_keys('%s')
            saved = True
        except Exception:
            saved = False

    if saved:
        logging.info(f"Save via address bar + Alt+N + Alt+S → {final_path}")
        print(f"SAVED_PATH:{final_path}")
        return final_path
    else:
        logging.info("Save via address bar + Alt+N + Alt+S failed to fire.")
        return None


# --- Helper: trigger DaVinci's 'Save Mod File' after services are applied ---
def trigger_save_mod_file(win):
    """
    Trigger DaVinci's 'Save Mod File' after services are applied.

    Strategy:
      1) Try clicking a toolbar/button whose text contains 'Save Mod'
         or looks like a Save button for the mod file.
      2) Fallback: send Alt+S (%s) and then Ctrl+S (^s) as accelerators.
    """
    try:
        win.set_focus()
    except Exception:
        pass

    # 1) Try to find and click a Save button on the main window
    try:
        for btn in win.descendants(control_type="Button"):
            try:
                label = (btn.window_text() or "").strip().lower()
            except Exception:
                continue

            if not label:
                continue

            # Prefer explicit 'Save Mod File', but accept generic 'Save' buttons too
            if "save mod" in label or ("save" in label and "file" in label) or label.startswith("save"):
                try:
                    btn.click_input()
                    logging.info(f"Clicked Save button on main window: '{label}'")
                    time.sleep(0.6)
                    return True
                except Exception as e:
                    logging.info(f"Failed clicking Save button '{label}': {e}")
                    # continue trying other strategies
    except Exception as e:
        logging.info(f"trigger_save_mod_file: error while scanning buttons: {e}")

    # 2) Fallback: keyboard accelerators
    # Try Alt+S (menu mnemonic for 'Save Mod File' in some builds)
    try:
        send_keys("%s")
        logging.info("trigger_save_mod_file: sent Alt+S to main window.")
        time.sleep(0.6)
        return True
    except Exception:
        pass

    # Try Ctrl+S as a generic Save shortcut
    try:
        send_keys("^s")
        logging.info("trigger_save_mod_file: sent Ctrl+S to main window.")
        time.sleep(0.6)
        return True
    except Exception:
        pass

    logging.info("trigger_save_mod_file: could not reliably trigger Save Mod File.")
    return False

######## this is the part where I will implement the solution automation########

SERVICE_LABELS = {
    "DPF": "DPF",
    "EGR": "EGR",
    "TVA": "TVA",
    "LAMBDA": "LAMBDA",
    "MAF": "MAF",
    "FLAPS": "FLAPS",
    "STARTSTOP": "STARTSTOP",
    "ADBLUE": "ADBLUE",
    "READINESS": "READINESS",
}

def parse_services_map(services: str) -> dict[str, str]:
    """
    'DPF OFF, EGR OFF' → {'DPF': 'OFF', 'EGR': 'OFF'}
    case-insensitive, ignores unknown tokens.
    """
    result = {}
    if not services:
        return result

    parts = [p.strip() for p in services.replace(";", ",").split(",") if p.strip()]
    for p in parts:
        up = p.upper()
        for key in SERVICE_LABELS.keys():
            if key in up:
                if "OFF" in up:
                    result[key] = "OFF"
                elif "ON" in up:
                    result[key] = "ON"
                break
    return result


def click_toggle_by_label(win, label_text: str):
    """
    Find the Text control 'DPF' / 'EGR' etc, then DOUBLE-CLICK slightly to its LEFT
    (where the ON/OFF slider is in your layout).
    """
    label_upper = label_text.upper()
    lbl = None

    # Find the label
    for t in win.descendants(control_type="Text"):
        try:
            txt = (t.window_text() or "").strip().upper()
        except Exception:
            continue
        if txt == label_upper:
            lbl = t
            break

    if not lbl:
        logging.info(f"Service label not found: {label_text}")
        return False

    try:
        win_rect = win.rectangle()
        r = lbl.rectangle()
        # Click a bit to the LEFT of the text, middle vertically
        offset = 40  # adjust if needed
        x = r.left - offset
        y = (r.top + r.bottom) // 2

        # Clamp inside window bounds so we don't click outside
        if x < win_rect.left + 5:
            x = win_rect.left + 5

        # DOUBLE-CLICK instead of single click
        win.double_click_input(coords=(x - win_rect.left, y - win_rect.top))
        logging.info(f"Double-clicked toggle LEFT of {label_text} at ({x},{y})")
        return True
    except Exception as e:
        logging.info(f"Failed double-clicking toggle for {label_text}: {e}")
        return False

def apply_services(win, services: str):
    """
    Apply requested services (e.g. 'DPF OFF, EGR OFF') by clicking toggles.
    Now uses DOUBLE-CLICK per label and hits ENTER after each toggle
    to dismiss any blocking popup.
    """
    svc_map = parse_services_map(services)
    if not svc_map:
        logging.info("No services requested; skipping toggle automation.")
        return

    logging.info(f"Applying services: {svc_map}")
    try:
        win.set_focus()
    except Exception:
        pass

    time.sleep(0.3)

    for key, state in svc_map.items():
        # Right now we assume default = ON, so:
        #   - OFF → double-click once to disable
        #   - ON  → do nothing (already ON)
        if state == "OFF":
            clicked = click_toggle_by_label(win, SERVICE_LABELS[key])
            time.sleep(0.15)
            if clicked:
                # handle any popup by pressing ENTER
                try:
                    send_keys("{ENTER}")
                except Exception:
                    pass
                # short pause so DaVinci can process it
                time.sleep(0.2)
        # if you ever need to force ON, you can also add logic here

def after_file_loaded_double_click_and_confirm(win, wait_before=5.0):
    """
    After BIN is loaded:
      1) wait a bit
      2) double-click in the central work area
      3) wait for a popup and click YES if present
    """
    # 1) wait for DaVinci to finish loading the file
    time.sleep(wait_before)

    # 2) double-click roughly in the center (you can tweak this later)
    try:
        r = win.rectangle()
        cx = int((r.left + r.right) / 2)
        cy = int((r.top + r.bottom) / 2)
        win.double_click_input(coords=(cx - r.left, cy - r.top))
        logging.info(f"After-load double-click at ({cx},{cy})")
    except Exception as e:
        logging.info(f"after_file_loaded_double_click_and_confirm: double-click failed: {e}")

    # 3) wait for and click YES on any confirmation popup
    try:
        maybe_click_yes_popup(timeout=8)
    except Exception:
        pass


######## end of solution automation########
def run(exe: Path, brand: str, ecu: str, input_path: str | None = None, services: str = ""):
    """Launch/attach DaVinci, select brand+ECU, and stop once the Open dialog appears.
    input_path and services are accepted for compatibility but not used here.
    """
    launch_if_needed(exe)
    logging.info("launched/attached")
    app, win = connect_window()
    tree = get_tree(win)
    logging.info(f"selecting brand={brand} ecu={ecu}")
    if tree is None:
        select_brand_ecu_keys(win, brand, ecu)
    else:
        select_brand_ecu_ui(tree, brand, ecu)
    logging.info("brand/ecu selection done")
    # Close the info dialog synchronously once; focus should now be in 'File name'
    try:
        maybe_close_info_dialog(timeout=5)
    except Exception:
        pass

    if not input_path:
        raise RuntimeError("Missing --input path: required to derive the filename.")

    filename = Path(input_path).name
    # As soon as the info dialog closes, immediately start typing the path and filename.
    type_folder_and_filename(r"C:\ecu_files\original", filename)
    print("OK: typed full path and submitted.")

    # 1) Wait for file to load (5s), double-click, and accept YES popup
    try:
        after_file_loaded_double_click_and_confirm(win)
    except Exception as e:
        logging.info(f"after_file_loaded_double_click_and_confirm failed: {e}")
    
    # 2) Apply services (DPF OFF, EGR OFF, etc.)
    try:
        apply_services(win, services)
    except Exception as e:
        logging.info(f"apply_services failed: {e}")

    # 3) Open Save Mod menu via Alt+M and confirm with ENTER twice
    try:
        win.set_focus()
    except Exception:
        pass
    try:
        send_keys('%m')
        time.sleep(0.3)
        send_keys('{ENTER}')
        time.sleep(0.2)
        send_keys('{ENTER}')
        time.sleep(0.5)
        logging.info('Triggered Save via Alt+M + ENTER + ENTER')
    except Exception as e:
        logging.info(f"Alt+M Save sequence failed: {e}")

    
    # PASSIVE WAIT for the Save dialog — handle save as required
    logging.info("Passive wait: watching for Save dialog to appear …")
    try:
        backend, sdlg = wait_save_dialog(timeout=180)  # explicitly detects 'Save Mod File' now
        try:
            sdlg.set_focus()
        except Exception:
            pass

        # Navigate using the address bar to C:\\ecu_files\\modified and keep existing filename
        saved_path = save_via_address_bar_using_existing_name(sdlg, target_folder=r"C:\\ecu_files\\modified")
        if saved_path:
            print(f"SAVED_PATH:{saved_path}")
        else:
            # Ensure the agent still receives a line it can parse
            # Try to read the filename again and synthesize the path
            fallback_name = get_current_filename_from_edit(sdlg) or "modified.bin"
            print(f"SAVED_PATH:{str(Path(r'C:\\\\ecu_files\\\\modified') / Path(fallback_name).name)}")

        # Handle optional overwrite prompts and wait for the dialog to close
        try:
            maybe_confirm_overwrite(timeout=8)
        except Exception:
            pass
        try:
            _wait_dialog_gone(sdlg, timeout=20)
        except Exception:
            pass

        
    except UIATimeout:
        logging.info("No Save dialog detected within timeout; continuing.")

def parse_args():
    p = argparse.ArgumentParser(
        description=(
            "DaVinci: select BRAND → double-click ECU → wait Open dialog → stop.\n"
            "Note: --input and --services are working now."
        )
    )
    p.add_argument("--exe", required=True, help="Path to davinci.exe")
    p.add_argument("--brand", required=True, help="Brand as shown in DaVinci (e.g., BMW)")
    p.add_argument("--ecu", required=True, help="ECU as shown under the brand (e.g., Bosch MEVD17.2)")
    # Back-compat only — these values are parsed but unused in this script
    p.add_argument("--input", help="(ignored) Full path to the BIN to open in the dialog")
    p.add_argument("--services", default="", help="(ignored) Services string e.g. 'Stage 0, DPF OFF'")
    return p.parse_args()

if __name__ == "__main__":
    try:
        a = parse_args()
        run(Path(a.exe), a.brand, a.ecu, input_path=a.input, services=a.services)
        sys.exit(0)
    except UIATimeout as e:
        print("ERROR:", str(e)); sys.exit(2)
    except Exception as e:
        print("ERROR:", str(e)); sys.exit(1)