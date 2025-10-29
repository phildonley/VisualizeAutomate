#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PRODUCTION VERSION - Enhanced with:
✓ Better output folder dialog handling (paste + click Select Folder button)
✓ Navigation controls in guided mode (skip forward/back)
✓ Improved close sequence with retries
✓ Longer waits for dialog opening
✓ All previous fixes maintained
"""

import os, re, sys, json, time, argparse
from dataclasses import dataclass
from typing import Dict, Optional

import pandas as pd
import pyperclip
import psutil
import keyboard
import mouse
import win32api, win32con, win32gui

try:
    import pythoncom, win32com.client
    HAVE_COM = True
except:
    HAVE_COM = False

OUTPUT_ROOT = r"C:\Users\Phillip.Donley\Downloads\Render Folder"
REQUIRED_CAM_SUFFIXES = ("103", "105", "107", "109", "111")
VK_F = 0x46

# --- Focus/Window helpers ---
def _get_fg_title():
    try:
        hwnd = win32gui.GetForegroundWindow()
        return win32gui.GetWindowText(hwnd) or ""
    except Exception:
        return ""

def _wait_for_dialog_title(substrs=("Open",), timeout=20.0, poll=0.25):
    """Wait until the foreground window title contains any of substrs."""
    end = time.time() + timeout
    while time.time() < end:
        t = _get_fg_title()
        if any(s.lower() in t.lower() for s in substrs):
            return True
        time.sleep(poll)
    return False

# Updated with new output folder steps
GUIDED_STEPS = [
    "import_ok_btn",           # 1. Import Settings OK button
    "camera_tab",              # 2. Camera tab
    "plus_tab",                # 3. Plus (+) icon for importing
    "import_cameras_btn",      # 4. Import cameras button
    "old_cam_1",               # 5. First old camera to delete
    "old_cam_2",               # 6. Second old camera to delete
    "cam_103",                 # 7. Camera 103 to start centering
    "wizard_next_or_render",   # 8. Next/Render button in wizard
    "job_name_textbox",        # 9. Job name text field
    "output_folder_btn",       # 10. Browse/... button for output folder
    "folder_select_btn",       # 11. "Select Folder" button in dialog (USING CORRECT NAME!)
    "cameras_dropdown",        # 12. Cameras dropdown in wizard
    "cameras_select_all",      # 13. Select All option
    "cameras_dropdown_close",  # 14. Close dropdown (click elsewhere)
    "render_no_save_btn",      # 15. No button when closing render window
    "project_no_save_btn"      # 16. No button when closing project
]
UI_POINTS_PATH = "ui_points.json"

class Logger:
    def __init__(self, v=False): self.v = v
    def _p(self, l, m): print(f"[{time.strftime('%H:%M:%S')}] {l:<5s} {m}"); sys.stdout.flush()
    def info(self, m): self._p("INFO", m)
    def warn(self, m): self._p("WARN", m)
    def error(self, m): self._p("ERROR", m)
    def dbg(self, m):
        if self.v: self._p("DBG", m)

log = Logger(False)

def send_hw_key(vk, ds=0.03, us=0.03):
    win32api.keybd_event(vk, 0, 0, 0); time.sleep(ds)
    win32api.keybd_event(vk, 0, win32con.KEYEVENTF_KEYUP, 0); time.sleep(us)

def sanitize_job_name(tms, part):
    tms = str(tms).strip()
    if re.fullmatch(r"\d+\.0", tms): tms = tms[:-2]
    part = str(part).strip().replace(".", "_")
    return (f"{tms}_{part}").replace(".", "_")

def is_visualize_running():
    for p in psutil.process_iter(attrs=["name"]):
        try:
            if p.info["name"] and "Visualize" in p.info["name"]: return True
        except: pass
    return False

def get_visualize_hwnd():
    def cb(hwnd, res):
        if win32gui.IsWindowVisible(hwnd):
            t = win32gui.GetWindowText(hwnd)
            if t and "Visualize" in t and "Open" not in t and "Import" not in t:
                res.append((hwnd, t))
    r = []
    win32gui.EnumWindows(cb, r)
    return r[0] if r else (None, None)

def focus_visualize():
    log.info("[FOCUS] Clicking Visualize...")
    v = get_visualize_hwnd()
    if v and len(v) == 2 and v[0]:
        h, t = v
        try:
            rect = win32gui.GetWindowRect(h)
            x = (rect[0] + rect[2]) // 2
            y = (rect[1] + rect[3]) // 2
            mouse.move(x, y, absolute=True, duration=0.1)
            time.sleep(0.2)
            mouse.click()
            time.sleep(0.5)
            win32gui.ShowWindow(h, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(h)
            time.sleep(1.0)
            log.info("[FOCUS] ✓")
            return True
        except Exception as e:
            log.warn(f"[FOCUS] {e}")
    return False

class UIPointsIO:
    def __init__(self, p=UI_POINTS_PATH):
        self.path = p; self.points = {}
    def load(self):
        if os.path.exists(self.path):
            self.points = json.load(open(self.path))
            log.info(f"[UI] Loaded {len(self.points)} points")
    def save(self): json.dump(self.points, open(self.path, "w"), indent=2)
    def set_point(self, l, x, y): self.points[l] = {"x": x, "y": y}
    def has(self, l): return l in self.points
    def get(self, l): d = self.points[l]; return d["x"], d["y"]

class GuidedRecorder:
    """
    Enhanced guided recorder with navigation controls:
    - Ctrl+Shift+Space: Capture current point and advance
    - Ctrl+Shift+Right: Skip forward without capturing
    - Ctrl+Shift+Left: Go back one step
    - Ctrl+Shift+Q: Save and quit
    """
    def __init__(self, io):
        self.io = io
        self.io.load()
        self.idx = 0
        # Find first uncaptured step
        for i, l in enumerate(GUIDED_STEPS):
            if not self.io.has(l):
                self.idx = i
                break
        else:
            self.idx = len(GUIDED_STEPS)  # All captured
        self.running = True
        self._bind()
    
    def _bind(self):
        keyboard.add_hotkey("ctrl+shift+space", self.cap)
        keyboard.add_hotkey("ctrl+shift+right", self.skip_forward)
        keyboard.add_hotkey("ctrl+shift+left", self.skip_back)
        keyboard.add_hotkey("ctrl+shift+q", self.fin)
    
    def cap(self):
        """Capture current point and advance"""
        if self.idx >= len(GUIDED_STEPS):
            log.warn("All steps captured!")
            return
        l = GUIDED_STEPS[self.idx]
        x, y = mouse.get_position()
        self.io.set_point(l, x, y)
        log.info(f"✓ [{self.idx+1}/{len(GUIDED_STEPS)}] {l} at ({x},{y})")
        self.idx += 1
    
    def skip_forward(self):
        """Skip to next step without capturing"""
        if self.idx < len(GUIDED_STEPS) - 1:
            self.idx += 1
            log.info(f"→ Skipped to step {self.idx+1}")
        else:
            log.warn("Already at last step")
    
    def skip_back(self):
        """Go back one step"""
        if self.idx > 0:
            self.idx -= 1
            log.info(f"← Back to step {self.idx+1}")
        else:
            log.warn("Already at first step")
    
    def fin(self):
        """Save and quit"""
        self.io.save()
        self.running = False
        log.info("✓ Saved and exiting!")
    
    def run(self):
        log.info("="*80)
        log.info("=== GUIDED UI CAPTURE MODE ===")
        log.info("="*80)
        log.info("")
        log.info("CONTROLS:")
        log.info("  Ctrl+Shift+Space  = Capture current point & advance")
        log.info("  Ctrl+Shift+Right  = Skip forward (don't capture)")
        log.info("  Ctrl+Shift+Left   = Go back one step")
        log.info("  Ctrl+Shift+Q      = Save & quit")
        log.info("")
        log.info("STEPS TO CAPTURE:")
        for i, s in enumerate(GUIDED_STEPS, 1):
            status = "✓" if self.io.has(s) else " "
            log.info(f"  [{status}] {i:2d}. {s}")
        log.info("")
        log.info("="*80)
        
        while self.running:
            if self.idx < len(GUIDED_STEPS):
                l = GUIDED_STEPS[self.idx]
                log.info(f">>> STEP {self.idx+1}/{len(GUIDED_STEPS)}: {l}")
            else:
                log.info(">>> ALL STEPS CAPTURED! Press Ctrl+Shift+Q to save and quit.")
            time.sleep(3)

class PDMClient:
    def __init__(self, vn):
        if not HAVE_COM: raise RuntimeError("No COM")
        self.vn = vn; self.v = None
    def login(self):
        pythoncom.CoInitialize()
        self.v = win32com.client.Dispatch("ConisioLib.EdmVault")
        self.v.LoginAuto(self.vn, 0)
        if not self.v.IsLoggedIn: raise RuntimeError("Login fail")
        log.info(f"[PDM] ✓ {self.vn}")
    def ensure_local(self, vp)->Optional[str]:
        if not self.v: self.login()
        fp, fn = os.path.dirname(vp), os.path.basename(vp)
        try: fld = self.v.GetFolderFromPath(fp)
        except Exception as e: log.warn(f"[PDM] {e}"); return None
        try: f = fld.GetFile(fn)
        except Exception as e: log.warn(f"[PDM] {e}"); return None
        try: f.GetFileCopy(0, 0, fld.ID, 0, "")
        except: pass
        try: lp = str(f.GetLocalPath(fld.ID))
        except:
            try: lp = os.path.join(str(fld.LocalPath), fn)
            except: lp = None
        if lp and os.path.exists(lp):
            log.info(f"[PDM] ✓ {lp}")
            return lp
        return None

class RenderWatcher:
    def __init__(self, root, settle=20):
        self.root = root; self.settle = settle
    def _cand(self, jn):
        if not os.path.isdir(self.root): return []
        ex = os.path.join(self.root, jn); p = []
        if os.path.isdir(ex): p.append(ex)
        for d in os.listdir(self.root):
            f = os.path.join(self.root, d)
            if os.path.isdir(f) and jn.lower() in d.lower() and f not in p: p.append(f)
        return p
    def wait_dir(self, jn, to=300):
        log.info(f"[WATCH] Waiting for {jn}...")
        e = time.time() + to
        while time.time() < e:
            c = self._cand(jn)
            if c: log.info(f"[WATCH] ✓ {c[0]}"); return c[0]
            time.sleep(2)
        log.error("[WATCH] ✗ Timeout"); return None
    def wait_five(self, jd)->bool:
        log.info(f"[WATCH] Looking in: {jd}")
        log.info("[WATCH] Waiting for 5 renders...")
        req = set(REQUIRED_CAM_SUFFIXES)
        
        if not os.path.exists(jd):
            log.error(f"[WATCH] ✗ Directory doesn't exist: {jd}")
            return False
        
        for attempt in range(300):
            found = {}
            try:
                files = os.listdir(jd)
                if attempt % 30 == 0:
                    log.info(f"[WATCH] Found {len(files)} files in folder...")
                
                for f in files:
                    if f.lower().endswith((".jpg", ".jpeg")):
                        for s in req:
                            if s in f:
                                found[s] = os.path.join(jd, f)
                                log.dbg(f"[WATCH] Found {s}: {f}")
                                break
            except Exception as e:
                log.warn(f"[WATCH] Error listing files: {e}")
                pass
            
            if len(found) >= 5:
                log.info(f"[WATCH] ✓ Found all 5 renders! Checking stability...")
                snap = {s: os.path.getsize(p) for s, p in found.items()}
                log.info(f"[WATCH] Waiting {self.settle}s for stability...")
                time.sleep(self.settle)
                
                ok = True
                for s, p in found.items():
                    try:
                        ns = os.path.getsize(p)
                        if ns != snap[s]:
                            log.warn(f"[WATCH] {s} still growing ({snap[s]}→{ns})")
                            ok = False
                    except:
                        ok = False
                
                if ok:
                    log.info("[WATCH] ✓ All stable!")
                    return True
                else:
                    log.info("[WATCH] Still rendering, waiting...")
            
            time.sleep(2)
        
        log.error("[WATCH] ✗ Timeout")
        return False

class VisualizeDriver:
    def __init__(self, io):
        self.io = io
        self.io.load()
    
    def _has(self, l): return self.io.has(l)
    
    def _click(self, l, d=1.0):
        if not self._has(l): raise RuntimeError(f"No {l}")
        x, y = self.io.get(l)
        mouse.move(x, y, absolute=True, duration=0.1)
        time.sleep(0.2)
        mouse.click()
        time.sleep(d)
    
    def _dbl(self, l):
        if not self._has(l): raise RuntimeError(f"No {l}")
        x, y = self.io.get(l)
        mouse.move(x, y, absolute=True, duration=0.1)
        time.sleep(0.2)
        mouse.double_click()
        time.sleep(0.5)
    
    def open_file(self, p):
        log.info(f"[OPEN] {p}")
        if not os.path.exists(p):
            log.error(f"[OPEN] ✗ Not found: {p}")
            return False
        
        if not focus_visualize():
            log.error("[OPEN] ✗ Can't focus")
            return False
        
        # Open file dialog
        log.info("[OPEN] Opening File menu...")
        keyboard.send("f")
        time.sleep(5)  # Wait for File menu to fully open
        
        log.info("[OPEN] Pressing Ctrl+O...")
        keyboard.send("ctrl+o")
        time.sleep(5)  # Wait for Open dialog (fast with PDM local checkout)
        
        # Paste path and open
        log.info("[OPEN] Copying filepath to clipboard...")
        pyperclip.copy(p)
        time.sleep(2)  # Wait for clipboard
        
        log.info("[OPEN] Pasting filepath...")
        keyboard.send("ctrl+v")
        time.sleep(5)  # Wait for paste to fully complete
        
        log.info("[OPEN] Pressing Enter to open...")
        keyboard.send("enter")
        time.sleep(10)  # Wait for file to start opening
        
        # Wait for Import Settings dialog and click OK
        log.info("[OPEN] Waiting for Import Settings dialog to fully appear...")
        time.sleep(90)  # keep your long wait; Visualize needs it
        
        if self._has("import_ok_btn"):
            # OPTIONAL: preview-move so you can visually verify the saved point is right
            x, y = self.io.get("import_ok_btn")
            log.info(f"[OPEN] Preview OK at ({x},{y})")
            mouse.move(x, y, absolute=True, duration=0.1)
            time.sleep(0.6)  # brief pause so you can see cursor over the button
        
            log.info("[OPEN] Clicking Import Settings OK...")
            mouse.click()
            time.sleep(2.5)  # give the modal time to close
        else:
            log.warn("[OPEN] No import_ok_btn saved; skipping click")

        
        # Wait for file to load
        log.info("[OPEN] Waiting for file to load...")
        time.sleep(10)
        
        # Focus viewport and click to ensure it's active
        log.info("[OPEN] Focusing viewport...")
        focus_visualize()
        time.sleep(1)
        
        # Click in center of window to activate viewport
        v = get_visualize_hwnd()
        if v and len(v) == 2 and v[0]:
            h, t = v
            try:
                rect = win32gui.GetWindowRect(h)
                center_x = (rect[0] + rect[2]) // 2
                center_y = (rect[1] + rect[3]) // 2
                mouse.move(center_x, center_y, absolute=True, duration=0.1)
                time.sleep(0.2)
                mouse.click()
                time.sleep(1)
                log.info("[OPEN] Viewport activated")
            except:
                log.warn("[OPEN] Could not click viewport center")
        
        log.info("[OPEN] ✓ File loaded and ready")
        return True

    def import_cams(self):
        log.info("[CAM] Importing...")
        req = ["camera_tab", "plus_tab", "import_cameras_btn"]
        if any(not self._has(p) for p in req):
            raise RuntimeError(f"Missing: {req}")

        # Open the import dialog
        self._click("camera_tab", d=1)
        time.sleep(2.0)
        self._click("plus_tab", d=1)
        time.sleep(2.0)
        self._click("import_cameras_btn", d=0.8)  # single click, short delay
        time.sleep(0.2) 
        
        # Wait for the File Open dialog to be foreground
        log.info("[CAM] Waiting for 'Open' dialog to foreground...")
        if not _wait_for_dialog_title(("Open", "Select", "Browse"), timeout=12):
            log.warn("[CAM] Open dialog not detected by title; proceeding cautiously after delay")
            time.sleep(2.0)

        # Ensure the 'File name' field gets focus: Alt+N targets that control in common dialogs
        log.info("[CAM] Focusing 'File name' box (Alt+N)...")
        keyboard.send("alt+n"); time.sleep(0.4)

        # Paste the camera list
        cs = "\"103\" \"105\" \"107\" \"109\" \"111\""
        pyperclip.copy(cs); time.sleep(0.2)
        log.info(f"[CAM] Pasting camera list: {cs}")
        keyboard.send("ctrl+v"); time.sleep(0.6)

        # Confirm open
        log.info("[CAM] Pressing Enter to import...")
        keyboard.send("enter"); time.sleep(3.0)

        log.info("[CAM] ✓")

    def del_old_cams(self):
        log.info("[CAM] Deleting old...")
        for i, l in enumerate(["old_cam_1", "old_cam_2"], 1):
            if self._has(l):
                self._dbl(l); time.sleep(1)
                keyboard.send("delete"); time.sleep(2)
        log.info("[CAM] ✓")

    def center_cams(self):
        log.info("[CAM] Centering...")
        if not self._has("cam_103"): raise RuntimeError("No cam_103")
        
        self._dbl("cam_103"); time.sleep(2)
        
        log.info("[CAM] Cam 1"); send_hw_key(VK_F); time.sleep(2)
        
        log.info("[CAM] Cam 2")
        keyboard.send("right"); time.sleep(0.5)
        keyboard.send("right"); time.sleep(0.5)
        keyboard.send("enter"); time.sleep(1)
        send_hw_key(VK_F); time.sleep(2)
        
        log.info("[CAM] Cam 3")
        keyboard.send("down"); time.sleep(0.5)
        keyboard.send("enter"); time.sleep(1)
        send_hw_key(VK_F); time.sleep(2)
        
        log.info("[CAM] Cam 4")
        keyboard.send("left"); time.sleep(0.5)
        keyboard.send("enter"); time.sleep(1)
        send_hw_key(VK_F); time.sleep(2)
        
        log.info("[CAM] Cam 5")
        keyboard.send("down"); time.sleep(0.5)
        keyboard.send("enter"); time.sleep(1)
        send_hw_key(VK_F); time.sleep(2)
        
        log.info("[CAM] ✓")

    def render(self, jn):
        log.info(f"[WIZ] Starting render wizard for: {jn}")
        
        # Start render wizard
        keyboard.send("ctrl+r")
        time.sleep(3)
        
        # Click Next 4 times
        log.info("[WIZ] Clicking Next 4x...")
        for i in range(4):
            self._click("wizard_next_or_render", d=2)
            log.info(f"[WIZ] Next {i+1}/4")
        
        # Set job name
        log.info("[WIZ] Setting job name...")
        self._click("job_name_textbox", d=1)
        keyboard.send("ctrl+a")
        time.sleep(0.5)
        pyperclip.copy(jn)
        keyboard.send("ctrl+v")
        time.sleep(1)
        log.info(f"[WIZ] Job name set: {jn}")
        
        # === OUTPUT FOLDER ===
        log.info("[WIZ] Setting output folder...")
        self._click("output_folder_btn", d=1.5)   # give dialog time to open
        
        log.info("[WIZ] Waiting for folder dialog to become foreground...")
        if not _wait_for_dialog_title(("Browse", "Select", "Folder"), timeout=20):
            log.warn("[WIZ] Dialog not detected; adding fallback wait")
            time.sleep(4)
        
        log.info("[WIZ] Allowing dialog to stabilize...")
        time.sleep(2.0)
        
        pyperclip.copy(OUTPUT_ROOT)
        keyboard.send("ctrl+v")
        time.sleep(1.5)
        
        log.info("[WIZ] Navigating to folder...")
        keyboard.send("enter")
        time.sleep(2.5)
        
        log.info("[WIZ] Confirming Select Folder via Enter...")
        keyboard.send("enter")
        time.sleep(2.0)
        
        if _wait_for_dialog_title(("Browse", "Select", "Folder"), timeout=2.0):
            if self._has("folder_select_btn"):
                x, y = self.io.get("folder_select_btn")
                log.info(f"[WIZ] Clicking folder_select_btn at ({x}, {y})")
                mouse.move(x, y, absolute=True, duration=0.1)
                time.sleep(0.4)
                mouse.click()
                time.sleep(2.5)
        
        log.info("[WIZ] ✓ Output folder set")


        # Select cameras
        log.info("[WIZ] Selecting cameras...")
        self._click("cameras_dropdown", d=2)
        time.sleep(0.4)
        log.info("[WIZ] Clicking Select All...")
        self._click("cameras_select_all", d=2)
        time.sleep(0.4)
        log.info("[WIZ] Closing dropdown...")
        time.sleep(0.4)
        keyboard.send("esc")
        log.info("[WIZ] ✓ All cameras selected")
        time.sleep(8.0)
        
        # Start render
        log.info("[WIZ] Starting render...")
        self._click("wizard_next_or_render", d=3)
        log.info("[WIZ] ✓ Render started!")

    def close(self):
        """
        Enhanced close sequence with retries.
        Closes render window and project without saving.
        """
        log.info("[CLOSE] Starting close sequence...")
        
        # === STEP 1: Close render window ===
        log.info("[CLOSE] Closing render window...")
        for attempt in range(3):
            keyboard.send("ctrl+w")
            time.sleep(2)
            
            # Try clicking No button if we have it
            if self._has("render_no_save_btn"):
                log.info(f"[CLOSE] Clicking render No button (attempt {attempt+1})...")
                try:
                    self._click("render_no_save_btn", d=2)
                    log.info("[CLOSE] ✓ Render window closed")
                    break
                except:
                    log.warn("[CLOSE] Click failed, retrying...")
            else:
                # Fallback: press N key
                log.info(f"[CLOSE] Pressing N key (attempt {attempt+1})...")
                keyboard.send("n")
                time.sleep(2)
                break
        
        # Small pause between closes
        time.sleep(1)
        
        # === STEP 2: Close project ===
        log.info("[CLOSE] Closing project...")
        for attempt in range(3):
            keyboard.send("ctrl+w")
            time.sleep(2)
            
            # Try clicking No button if we have it
            if self._has("project_no_save_btn"):
                log.info(f"[CLOSE] Clicking project No button (attempt {attempt+1})...")
                try:
                    self._click("project_no_save_btn", d=2)
                    log.info("[CLOSE] ✓ Project closed")
                    break
                except:
                    log.warn("[CLOSE] Click failed, retrying...")
            else:
                # Fallback: press N key
                log.info(f"[CLOSE] Pressing N key (attempt {attempt+1})...")
                keyboard.send("n")
                time.sleep(2)
                break
        
        # === SAFETY CLEANUP ===
        log.info("[CLOSE] Safety cleanup...")
        keyboard.send("escape")
        time.sleep(0.5)
        keyboard.send("escape")
        time.sleep(0.5)
        
        # One more Ctrl+W + N just in case
        keyboard.send("ctrl+w")
        time.sleep(1)
        keyboard.send("n")
        time.sleep(1)
        
        log.info("[CLOSE] ✓ Close sequence complete")

def read_excel(ep):
    for i in range(5):
        try: df=pd.read_excel(ep, engine="openpyxl"); break
        except PermissionError: time.sleep(1.5)
    else: raise
    
    if "A" not in df.columns and len(df.columns) >= 11:
        df.rename(columns={df.columns[0]:"A", df.columns[9]:"J", df.columns[10]:"K"}, inplace=True)
    
    for c in ("A","J","K"):
        if c not in df.columns: raise RuntimeError(f"Missing {c}")
    
    for idx,row in df.iterrows():
        yield {"A": row.get("A",""), "J": row.get("J",""), "K": row.get("K",""), "_index": idx}

def process(d, w, r, jdt, pdm):
    pt=str(r["A"]).strip()
    tms=str(r["K"]).strip()
    orig=str(r["J"]).strip()
    
    if not orig or not os.path.isabs(orig):
        log.info(f"[SKIP] Row {r['_index']} - No valid path")
        return
    
    up = orig
    if pdm:
        loc = pdm.ensure_local(orig)
        if loc and os.path.exists(loc): up = loc
    
    jn = sanitize_job_name(tms, pt)
    
    log.info("")
    log.info("="*80)
    log.info(f"JOB: {jn}")
    log.info("="*80)
    
    if not d.open_file(up):
        log.error("[ABORT] Failed to open file")
        return
    
    d.import_cams()
    d.del_old_cams()
    d.center_cams()
    d.render(jn)
    
    jdir = w.wait_dir(jn, jdt)
    if not jdir:
        log.error("[ABORT] Render folder not found")
        d.close()
        return
    
    if not w.wait_five(jdir):
        log.error("[ABORT] Not all renders completed")
        d.close()
        return
    
    d.close()
    log.info(f"✓✓✓ COMPLETED: {jn}")

def main():
    global log
    p=argparse.ArgumentParser(
        description="Solidworks Visualize Automation Script",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
EXAMPLES:
  # Capture UI points (guided mode with navigation)
  python visualize_automator.py --listen-guided
  
  # Run automation
  python visualize_automator.py --excel "C:\\path\\to\\file.xlsx" --pdm-vault "YourVaultName"
  
  # Run with verbose logging
  python visualize_automator.py --excel "C:\\path\\to\\file.xlsx" --verbose
        """
    )
    m=p.add_mutually_exclusive_group()
    m.add_argument("--listen-guided", action="store_true",
                   help="Enter guided mode to capture UI coordinates")
    p.add_argument("--excel", type=str,
                   help="Path to Excel file with parts list")
    p.add_argument("--verbose", action="store_true",
                   help="Enable verbose debug logging")
    p.add_argument("--jobdir-timeout", type=int, default=300,
                   help="Timeout in seconds for render folder to appear (default: 300)")
    p.add_argument("--settle-seconds", type=int, default=20,
                   help="Seconds to wait for files to stabilize (default: 20)")
    p.add_argument("--pdm-vault", type=str,
                   help="PDM vault name for file retrieval")
    a=p.parse_args()
    log=Logger(a.verbose)
    
    io=UIPointsIO()
    
    if a.listen_guided:
        GuidedRecorder(io).run()
        return
    
    if not is_visualize_running():
        log.warn("WARNING: Visualize doesn't appear to be running!")
    
    if not a.excel:
        log.error("ERROR: --excel argument is required for automation mode")
        log.error("Use --help for usage information")
        sys.exit(2)
    
    pdm=None
    if a.pdm_vault and HAVE_COM:
        try:
            pdm=PDMClient(a.pdm_vault)
            pdm.login()
        except Exception as e:
            log.error(f"[PDM] Failed to login: {e}")

    w=RenderWatcher(OUTPUT_ROOT, a.settle_seconds)
    d=VisualizeDriver(io)
    
    log.info("="*80)
    log.info("STARTING AUTOMATION")
    log.info("="*80)
    
    for row in read_excel(a.excel):
        try:
            process(d, w, row, a.jobdir_timeout, pdm)
        except KeyboardInterrupt:
            log.error("STOPPED BY USER (Ctrl+C)")
            break
        except Exception as e:
            log.error(f"ERROR on row {row.get('_index')}: {e}")
            import traceback
            traceback.print_exc()
    
    log.info("="*80)
    log.info("AUTOMATION COMPLETE")
    log.info("="*80)

if __name__=="__main__": main()
