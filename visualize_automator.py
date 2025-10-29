#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SOLIDWORKS Visualize Automation Script
Uses exact coordinates from ui_points.json with NO modifications
"""

import os
import sys
import json
import time
import argparse
import pyperclip
import keyboard
import mouse
import win32gui
import win32api
import win32con
import pandas as pd
import psutil

# Try to import COM for PDM support
try:
    import pythoncom
    import win32com.client
    HAVE_COM = True
except ImportError:
    HAVE_COM = False
    print("Warning: COM support not available - PDM functionality disabled")

# Configuration
OUTPUT_ROOT = r"C:\Users\Phillip.Donley\Downloads\Render Folder"
UI_POINTS_PATH = "ui_points.json"

# Automation sequence steps
AUTOMATION_STEPS = [
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

class Logger:
    """Simple logger with timestamps"""
    def __init__(self, verbose=False):
        self.verbose = verbose
    
    def _print(self, level, msg):
        timestamp = time.strftime('%H:%M:%S')
        print(f"[{timestamp}] {level:<5s} {msg}")
        sys.stdout.flush()
    
    def info(self, msg):
        self._print("INFO", msg)
    
    def warn(self, msg):
        self._print("WARN", msg)
    
    def error(self, msg):
        self._print("ERROR", msg)
    
    def debug(self, msg):
        if self.verbose:
            self._print("DEBUG", msg)

# Global logger instance
log = Logger()

class UIPoints:
    """Handles loading and accessing UI coordinates"""
    def __init__(self, path=UI_POINTS_PATH):
        self.path = path
        self.points = {}
        self.load()
    
    def load(self):
        """Load coordinates from JSON file"""
        if os.path.exists(self.path):
            with open(self.path, 'r') as f:
                self.points = json.load(f)
            log.info(f"Loaded {len(self.points)} UI points from {self.path}")
        else:
            log.error(f"UI points file not found: {self.path}")
            raise FileNotFoundError(f"Missing {self.path}")
    
    def get(self, name):
        """Get coordinates for a named UI element"""
        if name in self.points:
            point = self.points[name]
            return point["x"], point["y"]
        else:
            raise KeyError(f"UI point '{name}' not found")
    
    def has(self, name):
        """Check if a UI point exists"""
        return name in self.points
    
    def save(self):
        """Save coordinates to JSON file"""
        with open(self.path, 'w') as f:
            json.dump(self.points, f, indent=2)
        log.info(f"Saved {len(self.points)} UI points to {self.path}")
    
    def set_point(self, name, x, y):
        """Set coordinates for a UI element"""
        self.points[name] = {"x": x, "y": y}

class GuidedRecorder:
    """
    Interactive mode to capture UI coordinates
    Controls:
    - Ctrl+Shift+Space: Capture current point and advance
    - Ctrl+Shift+Right: Skip forward without capturing
    - Ctrl+Shift+Left: Go back one step
    - Ctrl+Shift+Q: Save and quit
    """
    def __init__(self, ui_points):
        self.ui = ui_points
        self.steps = AUTOMATION_STEPS
        self.current_index = 0
        self.running = True
        
        # Find first uncaptured step
        for i, step in enumerate(self.steps):
            if not self.ui.has(step):
                self.current_index = i
                break
        else:
            self.current_index = len(self.steps)  # All captured
        
        self._bind_keys()
    
    def _bind_keys(self):
        """Bind keyboard shortcuts"""
        keyboard.add_hotkey("ctrl+shift+space", self.capture_current)
        keyboard.add_hotkey("ctrl+shift+right", self.skip_forward)
        keyboard.add_hotkey("ctrl+shift+left", self.skip_back)
        keyboard.add_hotkey("ctrl+shift+q", self.save_and_quit)
    
    def capture_current(self):
        """Capture current mouse position and advance"""
        if self.current_index >= len(self.steps):
            log.warn("All steps captured!")
            return
        
        step_name = self.steps[self.current_index]
        x, y = mouse.get_position()
        self.ui.set_point(step_name, x, y)
        log.info(f"✓ [{self.current_index+1}/{len(self.steps)}] {step_name} at ({x}, {y})")
        self.current_index += 1
    
    def skip_forward(self):
        """Skip to next step without capturing"""
        if self.current_index < len(self.steps) - 1:
            self.current_index += 1
            log.info(f"→ Skipped to step {self.current_index+1}/{len(self.steps)}")
        else:
            log.warn("Already at last step")
    
    def skip_back(self):
        """Go back one step"""
        if self.current_index > 0:
            self.current_index -= 1
            log.info(f"← Back to step {self.current_index+1}/{len(self.steps)}")
        else:
            log.warn("Already at first step")
    
    def save_and_quit(self):
        """Save captured points and exit"""
        self.ui.save()
        self.running = False
        log.info("✓ Saved and exiting!")
    
    def run(self):
        """Run the guided capture mode"""
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
        for i, step in enumerate(self.steps, 1):
            status = "✓" if self.ui.has(step) else " "
            log.info(f"  [{status}] {i:2d}. {step}")
        log.info("")
        log.info("="*80)
        
        while self.running:
            if self.current_index < len(self.steps):
                step_name = self.steps[self.current_index]
                log.info(f">>> STEP {self.current_index+1}/{len(self.steps)}: {step_name}")
            else:
                log.info(">>> ALL STEPS CAPTURED! Press Ctrl+Shift+Q to save and quit.")
            time.sleep(3)

class PDMClient:
    """SOLIDWORKS PDM vault client for file management"""
    def __init__(self, vault_name):
        if not HAVE_COM:
            raise RuntimeError("COM support not available for PDM")
        self.vault_name = vault_name
        self.vault = None
        
    def login(self):
        """Login to PDM vault"""
        log.info(f"[PDM] Logging into vault: {self.vault_name}")
        try:
            pythoncom.CoInitialize()
            self.vault = win32com.client.Dispatch("ConisioLib.EdmVault")
            self.vault.LoginAuto(self.vault_name, 0)
            log.info("[PDM] Login successful")
        except Exception as e:
            log.error(f"[PDM] Login failed: {e}")
            raise
    
    def ensure_local(self, file_path):
        """Ensure file is available locally, download if needed"""
        if not self.vault:
            return file_path
            
        try:
            # Get folder and file interfaces
            folder_path = os.path.dirname(file_path)
            file_name = os.path.basename(file_path)
            
            folder = self.vault.GetFolderFromPath(folder_path)
            if not folder:
                log.warn(f"[PDM] Folder not in vault: {folder_path}")
                return file_path
            
            file = folder.GetFile(file_name)
            if not file:
                log.warn(f"[PDM] File not in vault: {file_name}")
                return file_path
            
            # Get latest version
            log.info(f"[PDM] Getting latest version of: {file_name}")
            file.GetFileCopy(0)  # 0 = parent window handle
            
            # Return local cache path
            local_path = file.GetLocalPath(folder.ID)
            log.info(f"[PDM] Local path: {local_path}")
            return local_path if local_path else file_path
            
        except Exception as e:
            log.warn(f"[PDM] Error getting file: {e}")
            return file_path

class VisualizeAutomation:
    """Main automation driver for SOLIDWORKS Visualize"""
    def __init__(self, ui_points):
        self.ui = ui_points
        self.click_duration = 0.1  # Mouse movement duration
        self.click_delay = 0.5     # Delay after clicking
    
    def click_point(self, point_name, wait_after=None):
        """Click a UI element using exact coordinates from JSON"""
        try:
            x, y = self.ui.get(point_name)
            log.info(f"Clicking {point_name} at ({x}, {y})")
            
            # Move mouse to exact coordinates - NO MODIFICATIONS
            mouse.move(x, y, absolute=True, duration=self.click_duration)
            time.sleep(0.2)  # Small delay to ensure mouse has settled
            
            # Perform click
            mouse.click()
            
            # Wait after click
            wait_time = wait_after if wait_after else self.click_delay
            time.sleep(wait_time)
            
        except Exception as e:
            log.error(f"Failed to click {point_name}: {e}")
            raise
    
    def type_text(self, text):
        """Type text using keyboard"""
        log.info(f"Typing: {text}")
        keyboard.write(text)
        time.sleep(0.3)
    
    def send_keys(self, keys):
        """Send keyboard shortcut"""
        log.info(f"Sending keys: {keys}")
        keyboard.send(keys)
        time.sleep(0.5)
    
    def paste_text(self, text):
        """Copy text to clipboard and paste"""
        log.info(f"Pasting: {text}")
        pyperclip.copy(text)
        time.sleep(0.2)
        keyboard.send("ctrl+v")
        time.sleep(0.3)
    
    def focus_visualize(self):
        """Ensure Visualize window is focused"""
        log.info("Focusing Visualize window...")
        
        def callback(hwnd, windows):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if title and "Visualize" in title and "Open" not in title and "Import" not in title:
                    windows.append((hwnd, title))
        
        windows = []
        win32gui.EnumWindows(callback, windows)
        
        if windows:
            hwnd, title = windows[0]
            try:
                # Restore if minimized
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                # Bring to front
                win32gui.SetForegroundWindow(hwnd)
                time.sleep(1.0)
                log.info(f"Focused: {title}")
                return True
            except Exception as e:
                log.warn(f"Failed to focus window: {e}")
        
        return False
    
    def open_file(self, filepath):
        """Open a file in Visualize"""
        log.info(f"Opening file: {filepath}")
        
        # Focus Visualize
        self.focus_visualize()
        
        # Open file dialog
        self.send_keys("ctrl+o")
        time.sleep(2.0)  # Wait for dialog
        
        # Type filename
        self.paste_text(filepath)
        time.sleep(1.0)
        
        # Press Enter to open
        self.send_keys("enter")
        time.sleep(5.0)  # Wait for file to load
        
        # Click Import OK button
        self.click_point("import_ok_btn", wait_after=3.0)
        
        return True
    
    def import_cameras(self):
        """Import cameras"""
        log.info("Importing cameras...")
        
        # Click camera tab
        self.click_point("camera_tab", wait_after=1.0)
        
        # Click plus tab
        self.click_point("plus_tab", wait_after=1.0)
        
        # Click import cameras button
        self.click_point("import_cameras_btn", wait_after=3.0)
        
        # Wait for import to complete
        time.sleep(2.0)
    
    def delete_old_cameras(self):
        """Delete old cameras"""
        log.info("Deleting old cameras...")
        
        # Delete first old camera
        self.click_point("old_cam_1", wait_after=0.5)
        self.send_keys("delete")
        time.sleep(1.0)
        
        # Delete second old camera
        self.click_point("old_cam_2", wait_after=0.5)
        self.send_keys("delete")
        time.sleep(1.0)
    
    def center_cameras(self):
        """Center cameras on model"""
        log.info("Centering cameras...")
        
        # Start with camera 103
        self.click_point("cam_103", wait_after=1.0)
        
        # Center view with F key
        self.send_keys("f")
        time.sleep(1.0)
        
        # Could add more camera operations here if needed
    
    def start_render_wizard(self, job_name):
        """Start render wizard and configure settings"""
        log.info("Starting render wizard...")
        
        # Open render wizard
        self.send_keys("ctrl+shift+r")
        time.sleep(3.0)
        
        # Click Next to go to settings
        self.click_point("wizard_next_or_render", wait_after=2.0)
        
        # Set job name
        self.click_point("job_name_textbox", wait_after=0.5)
        self.send_keys("ctrl+a")  # Select all
        time.sleep(0.3)
        self.type_text(job_name)
        time.sleep(1.0)
        
        # Set output folder
        self.click_point("output_folder_btn", wait_after=2.0)
        
        # In folder dialog, paste path
        output_path = os.path.join(OUTPUT_ROOT, job_name)
        self.paste_text(output_path)
        time.sleep(1.0)
        
        # Click Select Folder button
        self.click_point("folder_select_btn", wait_after=2.0)
        
        # Select all cameras
        self.click_point("cameras_dropdown", wait_after=1.0)
        self.click_point("cameras_select_all", wait_after=0.5)
        self.send_keys("escape")  # Close dropdown
        time.sleep(1.0)
        
        # Start render
        self.click_point("wizard_next_or_render", wait_after=3.0)
        
        log.info("Render started!")
    
    def close_project(self):
        """Close render window and project without saving"""
        log.info("Closing project...")
        
        # Close render window
        self.send_keys("ctrl+w")
        time.sleep(2.0)
        
        # Try clicking No button for render window
        if self.ui.has("render_no_save_btn"):
            try:
                self.click_point("render_no_save_btn", wait_after=2.0)
            except:
                # Fallback to keyboard
                self.send_keys("n")
                time.sleep(1.0)
        else:
            self.send_keys("n")
            time.sleep(1.0)
        
        # Close project
        self.send_keys("ctrl+w")
        time.sleep(2.0)
        
        # Try clicking No button for project
        if self.ui.has("project_no_save_btn"):
            try:
                self.click_point("project_no_save_btn", wait_after=2.0)
            except:
                # Fallback to keyboard
                self.send_keys("n")
                time.sleep(1.0)
        else:
            self.send_keys("n")
            time.sleep(1.0)
        
        log.info("Project closed")
    
    def wait_for_renders(self, job_name, timeout=300):
        """Wait for render files to complete"""
        log.info(f"Waiting for renders to complete...")
        
        output_dir = os.path.join(OUTPUT_ROOT, job_name)
        start_time = time.time()
        
        # Wait for directory to exist
        while not os.path.exists(output_dir):
            if time.time() - start_time > timeout:
                log.error("Timeout waiting for output directory")
                return False
            time.sleep(5)
        
        log.info(f"Output directory created: {output_dir}")
        
        # Wait for 5 camera files (103, 105, 107, 109, 111)
        expected_files = 5
        last_count = 0
        stable_time = 0
        
        while time.time() - start_time < timeout:
            # Count image files
            files = [f for f in os.listdir(output_dir) 
                    if f.lower().endswith(('.png', '.jpg', '.jpeg'))]
            
            if len(files) != last_count:
                last_count = len(files)
                stable_time = time.time()
                log.info(f"Found {len(files)} render files")
            
            # If we have all files and they've been stable for 20 seconds
            if len(files) >= expected_files and (time.time() - stable_time) > 20:
                log.info("All renders completed!")
                return True
            
            time.sleep(5)
        
        log.error("Timeout waiting for renders to complete")
        return False

def sanitize_job_name(tms, part):
    """Create a clean job name from TMS and part"""
    tms = str(tms).strip()
    if tms.endswith('.0'):
        tms = tms[:-2]
    
    part = str(part).strip().replace(".", "_")
    job_name = f"{tms}_{part}".replace(".", "_")
    
    return job_name

def process_row(automation, row_data, pdm_client=None):
    """Process a single row from Excel"""
    part = str(row_data.get("A", "")).strip()
    tms = str(row_data.get("K", "")).strip()
    filepath = str(row_data.get("J", "")).strip()
    
    if not filepath or not os.path.isabs(filepath):
        log.warn(f"Skipping row - no valid filepath")
        return False
    
    # Handle PDM file if client is available
    actual_filepath = filepath
    if pdm_client:
        try:
            local_path = pdm_client.ensure_local(filepath)
            if local_path and os.path.exists(local_path):
                actual_filepath = local_path
                log.info(f"[PDM] Using local cache: {actual_filepath}")
        except Exception as e:
            log.warn(f"[PDM] Failed to get file from vault: {e}")
            # Continue with original path
    
    job_name = sanitize_job_name(tms, part)
    
    log.info("="*80)
    log.info(f"Processing job: {job_name}")
    log.info(f"File: {actual_filepath}")
    log.info("="*80)
    
    try:
        # Execute automation sequence
        automation.open_file(actual_filepath)
        automation.import_cameras()
        automation.delete_old_cameras()
        automation.center_cameras()
        automation.start_render_wizard(job_name)
        
        # Wait for renders to complete
        if not automation.wait_for_renders(job_name):
            log.error("Render failed to complete")
            return False
        
        # Close project
        automation.close_project()
        
        log.info(f"✓ Completed job: {job_name}")
        return True
        
    except Exception as e:
        log.error(f"Failed to process job: {e}")
        import traceback
        traceback.print_exc()
        return False

def read_excel(filepath):
    """Read Excel file and yield rows"""
    try:
        df = pd.read_excel(filepath, engine='openpyxl')
        
        # Ensure we have the required columns
        if "A" not in df.columns and len(df.columns) >= 11:
            # Rename columns by position
            df.columns = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K"] + list(df.columns[11:])
        
        # Check for required columns
        required = ["A", "J", "K"]
        for col in required:
            if col not in df.columns:
                raise ValueError(f"Missing required column: {col}")
        
        # Yield rows
        for idx, row in df.iterrows():
            yield {
                "A": row.get("A", ""),
                "J": row.get("J", ""),
                "K": row.get("K", ""),
                "_index": idx
            }
            
    except Exception as e:
        log.error(f"Failed to read Excel file: {e}")
        raise

def is_visualize_running():
    """Check if SOLIDWORKS Visualize is running"""
    for proc in psutil.process_iter(['name']):
        try:
            if proc.info['name'] and 'Visualize' in proc.info['name']:
                return True
        except:
            pass
    return False

def main():
    """Main entry point"""
    global log
    
    parser = argparse.ArgumentParser(
        description="SOLIDWORKS Visualize Automation Script",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
EXAMPLES:
  # Capture UI points in guided mode
  python visualize_automator.py --guided-capture
  
  # Run automation on Excel file
  python visualize_automator.py --excel "C:\\path\\to\\parts.xlsx"
  
  # Run with PDM vault
  python visualize_automator.py --excel "C:\\path\\to\\parts.xlsx" --pdm-vault "YourVaultName"
  
  # Run with verbose logging
  python visualize_automator.py --excel "C:\\path\\to\\parts.xlsx" --verbose
  
  # Test single click
  python visualize_automator.py --test-click import_ok_btn
        """
    )
    
    # Create mutually exclusive group for modes
    mode_group = parser.add_mutually_exclusive_group()
    mode_group.add_argument("--guided-capture", action="store_true",
                           help="Enter guided mode to capture UI coordinates")
    mode_group.add_argument("--excel", type=str,
                           help="Path to Excel file with parts list")
    mode_group.add_argument("--test-click", type=str,
                           help="Test clicking a specific UI element")
    
    parser.add_argument("--verbose", action="store_true",
                       help="Enable verbose logging")
    parser.add_argument("--pdm-vault", type=str,
                       help="PDM vault name for file retrieval")
    
    args = parser.parse_args()
    
    # Set up logger
    log = Logger(verbose=args.verbose)
    
    # Check if Visualize is running
    if not is_visualize_running():
        log.warn("SOLIDWORKS Visualize doesn't appear to be running!")
        response = input("Continue anyway? (y/n): ")
        if response.lower() != 'y':
            sys.exit(1)
    
    # Load UI points
    try:
        ui_points = UIPoints()
    except FileNotFoundError:
        log.error("UI points file not found. Make sure ui_points.json is in the current directory.")
        sys.exit(1)
    
    # Create automation instance
    automation = VisualizeAutomation(ui_points)
    
    # Test mode - click a single element
    if args.test_click:
        log.info(f"Test mode - clicking {args.test_click}")
        try:
            automation.focus_visualize()
            automation.click_point(args.test_click)
            log.info("Test click completed")
        except Exception as e:
            log.error(f"Test click failed: {e}")
        sys.exit(0)
    
    # Check Excel argument
    if not args.excel:
        log.error("ERROR: --excel argument is required")
        parser.print_help()
        sys.exit(1)
    
    if not os.path.exists(args.excel):
        log.error(f"Excel file not found: {args.excel}")
        sys.exit(1)
    
    # Initialize PDM client if requested
    pdm_client = None
    if args.pdm_vault:
        if not HAVE_COM:
            log.error("COM support not available - cannot use PDM vault")
            log.error("Make sure pywin32 is installed: pip install pywin32")
            sys.exit(1)
        
        try:
            pdm_client = PDMClient(args.pdm_vault)
            pdm_client.login()
            log.info(f"[PDM] Connected to vault: {args.pdm_vault}")
        except Exception as e:
            log.error(f"[PDM] Failed to connect to vault: {e}")
            response = input("Continue without PDM? (y/n): ")
            if response.lower() != 'y':
                sys.exit(1)
            pdm_client = None
    
    # Process Excel file
    log.info("="*80)
    log.info("STARTING AUTOMATION")
    log.info(f"Excel file: {args.excel}")
    log.info(f"Output root: {OUTPUT_ROOT}")
    if pdm_client:
        log.info(f"PDM vault: {args.pdm_vault}")
    log.info("="*80)
    
    success_count = 0
    error_count = 0
    
    try:
        for row in read_excel(args.excel):
            try:
                if process_row(automation, row, pdm_client):
                    success_count += 1
                else:
                    error_count += 1
            except KeyboardInterrupt:
                log.warn("Stopped by user (Ctrl+C)")
                break
            except Exception as e:
                log.error(f"Error processing row {row.get('_index', '?')}: {e}")
                error_count += 1
    
    except Exception as e:
        log.error(f"Fatal error: {e}")
    
    # Summary
    log.info("="*80)
    log.info("AUTOMATION COMPLETE")
    log.info(f"Successful: {success_count}")
    log.info(f"Errors: {error_count}")
    log.info("="*80)

if __name__ == "__main__":
    main()
