#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PRODUCTION VERSION - Complete with all fixes:
✓ DPI awareness to prevent coordinate scaling
✓ Fixed render wizard timing (4 clicks, no Enter keys)
✓ Alt+F close sequence with proper waits
✓ PDM session management and prefetch optimization
✓ All timing and functionality preserved
"""

# CRITICAL: Set DPI awareness BEFORE any other imports
import ctypes
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)  # PROCESS_PER_MONITOR_DPI_AWARE
except:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except:
        pass

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
    "folder_select_btn",       # 11. "Select Folder" button in dialog
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
        self._session_count = 0  # Track open operations
    
    def login(self):
        pythoncom.CoInitialize()
        self.v = win32com.client.Dispatch("ConisioLib.EdmVault")
        self.v.LoginAuto(self.vn, 0)
    
    def ensure_session(self):
        """
        Called before each file operation to ensure PDM session is active.
        Re-initializes if stale.
        """
        self._session_count += 1
        if self._session_count % 50 == 0:  # Every 50 opens, refresh
            log.info("[PDM] Refreshing session...")
            try:
                # Test if session is alive
                self.v.RootFolderPath
            except:
                # Session is dead, re-login
                log.warn("[PDM] Session stale, re-logging in...")
                self.login()
    
    def preflight_local(self, p):
        """
        Pre-fetch and warm cache for a file before Visualize opens it.
        This prevents Visualize's slow internal PDM search.
        """
        if not self.v or not os.path.isabs(p):
            return p
        
        try:
            # Ensure session is fresh
            self.ensure_session()
            
            # Get folder and file
            dn = os.path.dirname(p)
            bn = os.path.basename(p)
            fo = self.v.GetFolderFromPath(dn)
            if not fo:
                log.dbg(f"[PDM] Folder not in vault: {dn}")
                return p
            
            fi = fo.GetFile(bn)
            if not fi:
                log.dbg(f"[PDM] File not in vault: {bn}")
                return p
            
            # Get latest version
            log.info(f"[PDM] Pre-fetching: {bn}")
            fi.GetFileCopy(0)
            
            # Get local cache path
            lp = fi.GetLocalPath(fo.ID)
            if lp and os.path.exists(lp):
                # Warm OS cache by reading first few bytes
                try:
                    with open(lp, 'rb') as f:
                        f.read(4096)
                    log.dbg(f"[PDM] Cache warmed: {lp}")
                except:
                    pass
                return lp
            
            return p
            
        except Exception as e:
            log.warn(f"[PDM] Preflight error: {e}")
            return p
    
    def ensure_local(self, p):
        """Legacy method - now just calls preflight_local"""
        return self.preflight_local(p)

class RenderWatcher:
    def __init__(self, root, settle=20):
        self.root = root
        self.settle = settle
        self._found_cache = {}  # Reset tracking
    
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
        # Clear cache for new job
        self._found_cache = {}
        
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
                                if s not in self._found_cache:
                                    log.info(f"[WATCH] Found {s}: {f}")
                                    self._found_cache[s] = f
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
                        log.warn(f"[WATCH] {s} disappeared?")
                        ok = False
                
                if ok:
                    log.info("[WATCH] ✓ All renders stable!")
                    return True
            
            time.sleep(1)
        
        log.error("[WATCH] ✗ Timeout waiting for renders")
        return False

class VisualizeDriver:
    def __init__(self, io):
        self.io = io; self.io.load()
        
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
        log.info(f"[OPEN] Opening: {os.path.basename(p)}")
        
        # Focus and prepare
        focus_visualize()
        time.sleep(0.5)
        
        log.info("[OPEN] Opening file dialog (Ctrl+O)...")
        keyboard.send("ctrl+o")
        time.sleep(6)  # Your extended wait for dialog
        
        log.info("[OPEN] Waiting for Open dialog...")
        if _wait_for_dialog_title(timeout=10):
            log.info("[OPEN] Dialog detected")
            time.sleep(2)
        else:
            log.warn("[OPEN] Dialog not detected; continuing")
            time.sleep(5)
        
        log.info("[OPEN] Typing filepath...")
        pyperclip.copy("")
        time.sleep(0.5)
        pyperclip.copy(p)
        time.sleep(0.5)
        keyboard.send("ctrl+v")
        time.sleep(5)  # Wait for paste to fully complete
        
        log.info("[OPEN] Pressing Enter to open...")
        keyboard.send("enter")
        time.sleep(10)  # Wait for file to start opening
        
        # Wait for Import Settings dialog and click OK
        log.info("[OPEN] Waiting for Import Settings dialog to fully appear...")
        time.sleep(90)  # Keep your long wait
        
        if self._has("import_ok_btn"):
            # Preview-move so you can visually verify the saved point is right
            x, y = self.io.get("import_ok_btn")
            log.info(f"[OPEN] Preview OK at ({x},{y})")
            mouse.move(x, y, absolute=True, duration=0.1)
            time.sleep(0.6)  # Brief pause so you can see cursor over the button
        
            log.info("[OPEN] Clicking Import Settings OK...")
            mouse.click()
            time.sleep(2.5)  # Give the modal time to close
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
                time.sleep(0.3)
                mouse.click()
                time.sleep(0.5)
            except:
                pass
        
        log.info("[OPEN] ✓ File loaded and ready")
        return True

    def import_cams(self):
        log.info("[CAM] Importing cameras...")
        self._click("camera_tab", d=1)
        time.sleep(0.5)
        self._click("plus_tab", d=1)
        time.sleep(0.5)
        self._click("import_cameras_btn", d=0.8)  # Single click, short delay
        log.info("[CAM] Waiting for camera import to complete...")
        time.sleep(8)
        
        # Focus viewport again after import
        log.info("[CAM] Re-focusing viewport after import...")
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
                time.sleep(0.5)
            except:
                pass
        
        log.info("[CAM] ✓")

    def del_old_cams(self):
        log.info("[CAM] Deleting old cameras...")
        
        # First camera
        if self._has("old_cam_1"):
            self._click("old_cam_1", d=0.8)
            keyboard.send("delete")
            time.sleep(1)
        
        # Second camera
        if self._has("old_cam_2"):
            self._click("old_cam_2", d=0.8)
            keyboard.send("delete")
            time.sleep(1)
        
        log.info("[CAM] ✓")
    
    def center_cams(self):
        log.info("[CAM] Centering...")
        
        # Check for 'cam_103' first
        if self._has("cam_103"):
            log.info("[CAM] Found cam_103")
            self._dbl("cam_103")
        else:
            log.warn("[CAM] No cam_103 saved; skipping camera center")
            return
        
        # Check for 'viewport_canvas' for additional setup
        if self._has("viewport_canvas"):
            log.info("[CAM] Clicking viewport_canvas...")
            x, y = self.io.get("viewport_canvas")
            mouse.move(x, y, absolute=True, duration=0.1)
            time.sleep(0.4)
            mouse.click()
            time.sleep(0.5)
        
        # Center view
        send_hw_key(VK_F); time.sleep(2)
        
        log.info("[CAM] ✓")

    def render(self, jn):
        log.info(f"[WIZ] Starting render wizard for: {jn}")
        
        # Start render wizard
        keyboard.send("ctrl+r")
        time.sleep(3)
        
        # Click Next exactly 4 times with small pauses between clicks
        log.info("[WIZ] Clicking Next 4 times to reach Job Settings page...")
        for i in range(4):
            if not self._has("wizard_next_or_render"):
                log.error("[WIZ] No wizard_next_or_render point saved")
                return
            
            log.info(f"[WIZ] Next click {i+1}/4...")
            self._click("wizard_next_or_render", d=0.3)  # 0.3 second pause after each click
            
        # Now we should be on the Job Settings page
        time.sleep(1.0)  # Give the final page a moment to load completely
        
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
        self._click("output_folder_btn", d=1.5)   # Give dialog time to open
        
        log.info("[WIZ] Waiting for folder dialog to become foreground...")
        if not _wait_for_dialog_title(("Browse", "Select", "Folder"), timeout=20):
            log.warn("[WIZ] Dialog not detected; adding fallback wait")
            time.sleep(4)
        
        log.info("[WIZ] Allowing dialog to stabilize...")
        time.sleep(2.0)
        
        # Focus the address bar so paste goes to the right place
        log.info("[WIZ] Focusing address bar (Alt+D)...")
        keyboard.send("alt+d")
        time.sleep(0.4)
        
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
        Close render window and project using Alt+F menu approach.
        IMPORTANT: This is called AFTER renders complete!
        """
        log.info("[CLOSE] Starting close sequence...")
        
        # === STEP 1: Close render window ===
        log.info("[CLOSE] Step 1: Closing render window...")
        
        # Ensure Visualize has focus
        focus_visualize()
        time.sleep(1.0)
        
        # Open File menu with Alt+F
        log.info("[CLOSE] Opening File menu (Alt+F)...")
        keyboard.send("alt+f")
        time.sleep(1.0)  # Wait for menu to open
        
        # Close with Ctrl+W
        log.info("[CLOSE] Closing window (Ctrl+W)...")
        keyboard.send("ctrl+w")
        time.sleep(2.0)  # Wait for save dialog to appear
        
        # Handle save dialog - click No or press N
        if self._has("render_no_save_btn"):
            log.info("[CLOSE] Clicking 'No' button for render window...")
            try:
                time.sleep(1.0)
                self._click("render_no_save_btn", d=2.0)
                log.info("[CLOSE] ✓ Render window closed via button")
            except:
                log.warn("[CLOSE] Button click failed, using keyboard...")
                keyboard.send("n")
                time.sleep(1.0)
        else:
            log.info("[CLOSE] Pressing 'N' key...")
            keyboard.send("n")
            time.sleep(1.0)
        
        # === IMPORTANT: 10 second pause between closes ===
        log.info("[CLOSE] Waiting 10 seconds before closing project...")
        time.sleep(10.0)
        
        # === STEP 2: Close project file ===
        log.info("[CLOSE] Step 2: Closing project file...")
        
        # Ensure Visualize still has focus
        focus_visualize()
        time.sleep(1.0)
        
        # Open File menu with Alt+F
        log.info("[CLOSE] Opening File menu (Alt+F)...")
        keyboard.send("alt+f")
        time.sleep(1.0)  # Wait for menu to open
        
        # Close with Ctrl+W
        log.info("[CLOSE] Closing project (Ctrl+W)...")
        keyboard.send("ctrl+w")
        time.sleep(2.0)  # Wait for save dialog to appear
        
        # Handle save dialog - click No or press N
        if self._has("project_no_save_btn"):
            log.info("[CLOSE] Clicking 'No' button for project...")
            try:
                time.sleep(1.0)
                self._click("project_no_save_btn", d=2.0)
                log.info("[CLOSE] ✓ Project closed via button")
            except:
                log.warn("[CLOSE] Button click failed, using keyboard...")
                keyboard.send("n")
                time.sleep(1.0)
        else:
            log.info("[CLOSE] Pressing 'N' key...")
            keyboard.send("n")
            time.sleep(1.0)
        
        # === Final cleanup ===
        log.info("[CLOSE] Final cleanup...")
        time.sleep(2.0)
        keyboard.send("escape")  # Clear any lingering dialogs
        time.sleep(0.5)
        
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
        # Use preflight_local for PDM optimization
        loc = pdm.preflight_local(orig)
        if loc and os.path.exists(loc): 
            up = loc
    
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
