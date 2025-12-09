"""
VibeVoice Startup Tray Application
A lightweight system tray app that:
- Starts automatically on Windows login
- Monitors the SAPI pipe server status
- Auto-starts the server if configured
- Provides quick access to start/stop/status
"""

import ctypes
import os
import sys
import subprocess
import threading
import time
import json
from pathlib import Path

# Ensure parent paths are available
SCRIPT_DIR = Path(__file__).parent.absolute()
SAPI_DIR = SCRIPT_DIR.parent
PROJECT_DIR = SAPI_DIR.parent
sys.path.insert(0, str(PROJECT_DIR))

import tkinter as tk
from tkinter import ttk, messagebox

# Windows-specific
if sys.platform == "win32":
    import win32file
    import pywintypes
else:
    print("Error: Windows only")
    sys.exit(1)

# Constants
PIPE_NAME = r"\\.\pipe\vibevoice"
CONFIG_FILE = SCRIPT_DIR / "vibevoice_tray_config.json"
PID_FILE = PROJECT_DIR / "server.pid"
SERVER_SCRIPT = PROJECT_DIR / "demo" / "sapi_pipe_server.py"

# Default config
DEFAULT_CONFIG = {
    "auto_start_server": True,
    "device": "cuda:0",
    "minimize_to_tray": True,
    "start_minimized": True,
}


def load_config() -> dict:
    """Load config from JSON file."""
    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE, "r") as f:
                cfg = json.load(f)
                # Merge with defaults for any missing keys
                return {**DEFAULT_CONFIG, **cfg}
        except:
            pass
    return DEFAULT_CONFIG.copy()


def save_config(cfg: dict):
    """Save config to JSON file."""
    with open(CONFIG_FILE, "w") as f:
        json.dump(cfg, f, indent=2)


def is_pipe_available() -> bool:
    """Check if the named pipe is responding."""
    try:
        h = win32file.CreateFile(
            PIPE_NAME,
            win32file.GENERIC_READ | win32file.GENERIC_WRITE,
            0, None, 3, 0, None
        )
        win32file.CloseHandle(h)
        return True
    except pywintypes.error:
        return False


def get_server_pid() -> int | None:
    """Get PID from file if server is running."""
    if PID_FILE.exists():
        try:
            pid = int(PID_FILE.read_text().strip())
            # Verify it's actually running
            result = subprocess.run(
                ["tasklist", "/FI", f"PID eq {pid}", "/FO", "CSV"],
                capture_output=True, text=True
            )
            if str(pid) in result.stdout:
                return pid
        except:
            pass
    return None


def start_server(device: str) -> int | None:
    """Start the SAPI pipe server, return PID."""
    if not SERVER_SCRIPT.exists():
        return None

    try:
        proc = subprocess.Popen(
            [sys.executable, str(SERVER_SCRIPT), "--device", device],
            cwd=str(PROJECT_DIR),
            creationflags=subprocess.CREATE_NEW_PROCESS_GROUP | subprocess.CREATE_NO_WINDOW
        )
        PID_FILE.write_text(str(proc.pid))
        return proc.pid
    except Exception as e:
        print(f"Failed to start server: {e}")
        return None


def stop_server() -> bool:
    """Stop the server by PID."""
    pid = get_server_pid()
    if pid:
        try:
            subprocess.run(["taskkill", "/F", "/PID", str(pid)], capture_output=True)
            if PID_FILE.exists():
                PID_FILE.unlink()
            return True
        except:
            pass
    return False


class VibeVoiceTrayApp:
    """Mini UI for startup/tray management."""

    def __init__(self, root: tk.Tk, start_minimized: bool = False):
        self.root = root
        self.root.title("VibeVoice TTS")
        self.root.geometry("380x320")
        self.root.resizable(False, False)

        self.config = load_config()
        self.monitoring = True
        self.server_pid = None

        self._setup_ui()
        self._start_monitor_thread()

        # Handle window close
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        # Start minimized if configured
        if start_minimized or self.config.get("start_minimized", False):
            self.root.withdraw()
            # Show briefly then hide (for tray icon to work)
            self.root.after(100, self._minimize_to_tray)

        # Auto-start server if configured
        if self.config.get("auto_start_server", False):
            self.root.after(500, self._auto_start_server)

    def _setup_ui(self):
        """Build the UI."""
        main = ttk.Frame(self.root, padding=15)
        main.pack(fill=tk.BOTH, expand=True)

        # Title
        ttk.Label(main, text="VibeVoice TTS", font=("Segoe UI", 14, "bold")).pack(pady=(0, 10))

        # Status Frame
        status_frame = ttk.LabelFrame(main, text="Status", padding=10)
        status_frame.pack(fill=tk.X, pady=5)

        self.status_indicator = tk.Canvas(status_frame, width=16, height=16, highlightthickness=0)
        self.status_indicator.pack(side=tk.LEFT, padx=(0, 10))
        self._draw_indicator("gray")

        self.status_label = ttk.Label(status_frame, text="Checking...", font=("Segoe UI", 10))
        self.status_label.pack(side=tk.LEFT)

        self.pid_label = ttk.Label(status_frame, text="", font=("Segoe UI", 9), foreground="gray")
        self.pid_label.pack(side=tk.RIGHT)

        # Control Buttons
        btn_frame = ttk.Frame(main)
        btn_frame.pack(fill=tk.X, pady=15)

        self.start_btn = ttk.Button(btn_frame, text="Start Server", command=self._start_server)
        self.start_btn.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)

        self.stop_btn = ttk.Button(btn_frame, text="Stop Server", command=self._stop_server, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)

        # Settings Frame
        settings_frame = ttk.LabelFrame(main, text="Settings", padding=10)
        settings_frame.pack(fill=tk.X, pady=5)

        # Auto-start checkbox
        self.auto_start_var = tk.BooleanVar(value=self.config.get("auto_start_server", True))
        ttk.Checkbutton(
            settings_frame, text="Auto-start server on login",
            variable=self.auto_start_var, command=self._save_settings
        ).pack(anchor=tk.W)

        # Start minimized checkbox
        self.start_min_var = tk.BooleanVar(value=self.config.get("start_minimized", True))
        ttk.Checkbutton(
            settings_frame, text="Start minimized",
            variable=self.start_min_var, command=self._save_settings
        ).pack(anchor=tk.W)

        # Device selection
        device_frame = ttk.Frame(settings_frame)
        device_frame.pack(fill=tk.X, pady=(10, 0))
        ttk.Label(device_frame, text="Device:").pack(side=tk.LEFT)

        self.device_var = tk.StringVar(value=self.config.get("device", "cuda:0"))
        self.device_combo = ttk.Combobox(device_frame, textvariable=self.device_var, width=20)
        self.device_combo.pack(side=tk.LEFT, padx=10)
        self.device_combo.bind("<<ComboboxSelected>>", lambda e: self._save_settings())
        self._load_devices()

        # Bottom buttons
        bottom_frame = ttk.Frame(main)
        bottom_frame.pack(fill=tk.X, pady=(15, 0))

        ttk.Button(bottom_frame, text="Open Full Manager", command=self._open_manager).pack(side=tk.LEFT)
        ttk.Button(bottom_frame, text="Hide", command=self._minimize_to_tray).pack(side=tk.RIGHT)

    def _draw_indicator(self, color: str):
        """Draw status indicator circle."""
        self.status_indicator.delete("all")
        self.status_indicator.create_oval(2, 2, 14, 14, fill=color, outline="")

    def _load_devices(self):
        """Load available CUDA devices."""
        devices = ["cuda:0", "cpu"]
        try:
            result = subprocess.run(
                ["nvidia-smi", "--query-gpu=index,name", "--format=csv,noheader"],
                capture_output=True, text=True
            )
            if result.returncode == 0:
                devices = ["cpu"]
                for line in result.stdout.strip().split('\n'):
                    if line.strip():
                        idx, name = line.split(',', 1)
                        devices.append(f"cuda:{idx.strip()}")
        except:
            pass
        self.device_combo['values'] = devices

    def _save_settings(self):
        """Save current settings to config."""
        self.config["auto_start_server"] = self.auto_start_var.get()
        self.config["start_minimized"] = self.start_min_var.get()
        self.config["device"] = self.device_var.get()
        save_config(self.config)

    def _start_monitor_thread(self):
        """Start background thread to monitor server status."""
        def monitor():
            while self.monitoring:
                self._update_status()
                time.sleep(2)

        thread = threading.Thread(target=monitor, daemon=True)
        thread.start()

    def _update_status(self):
        """Update UI with current server status."""
        pid = get_server_pid()
        pipe_ok = is_pipe_available()

        def update_ui():
            if pid and pipe_ok:
                self._draw_indicator("#22c55e")  # Green
                self.status_label.config(text="Server Running")
                self.pid_label.config(text=f"PID: {pid}")
                self.start_btn.config(state=tk.DISABLED)
                self.stop_btn.config(state=tk.NORMAL)
            elif pid:
                self._draw_indicator("#eab308")  # Yellow - starting
                self.status_label.config(text="Starting...")
                self.pid_label.config(text=f"PID: {pid}")
                self.start_btn.config(state=tk.DISABLED)
                self.stop_btn.config(state=tk.NORMAL)
            else:
                self._draw_indicator("#ef4444")  # Red
                self.status_label.config(text="Server Stopped")
                self.pid_label.config(text="")
                self.start_btn.config(state=tk.NORMAL)
                self.stop_btn.config(state=tk.DISABLED)

            self.server_pid = pid

        self.root.after(0, update_ui)

    def _auto_start_server(self):
        """Auto-start server if not running."""
        if not get_server_pid():
            self._start_server()

    def _start_server(self):
        """Start the server."""
        device = self.device_var.get()
        self.status_label.config(text="Starting...")
        self._draw_indicator("#eab308")

        def do_start():
            pid = start_server(device)
            if pid:
                self.root.after(0, lambda: self.status_label.config(text=f"Started (PID: {pid})"))
            else:
                self.root.after(0, lambda: messagebox.showerror("Error", "Failed to start server"))

        threading.Thread(target=do_start, daemon=True).start()

    def _stop_server(self):
        """Stop the server."""
        self.status_label.config(text="Stopping...")
        stop_server()
        self._update_status()

    def _minimize_to_tray(self):
        """Hide window (minimize to tray behavior)."""
        self.root.withdraw()

    def _open_manager(self):
        """Open the full installer/manager."""
        manager_path = SCRIPT_DIR / "vibevoice_installer.py"
        if manager_path.exists():
            subprocess.Popen([sys.executable, str(manager_path)], cwd=str(SCRIPT_DIR))

    def _on_close(self):
        """Handle window close - minimize instead of exit."""
        if self.config.get("minimize_to_tray", True):
            self._minimize_to_tray()
        else:
            self._quit()

    def _quit(self):
        """Actually quit the application."""
        self.monitoring = False
        self.root.quit()
        self.root.destroy()

    def show(self):
        """Show the window."""
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()


def add_to_startup(enable: bool = True):
    """Add/remove from Windows startup via registry."""
    import winreg

    key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
    app_name = "VibeVoiceTTS"

    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_ALL_ACCESS)

        if enable:
            # Create startup entry
            script_path = str(SCRIPT_DIR / "vibevoice_tray.py")
            cmd = f'"{sys.executable}" "{script_path}" --minimized'
            winreg.SetValueEx(key, app_name, 0, winreg.REG_SZ, cmd)
            print(f"Added to startup: {cmd}")
        else:
            # Remove startup entry
            try:
                winreg.DeleteValue(key, app_name)
                print("Removed from startup")
            except FileNotFoundError:
                pass

        winreg.CloseKey(key)
        return True
    except Exception as e:
        print(f"Startup registry error: {e}")
        return False


def is_in_startup() -> bool:
    """Check if app is in Windows startup."""
    import winreg

    key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
    app_name = "VibeVoiceTTS"

    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_READ)
        winreg.QueryValueEx(key, app_name)
        winreg.CloseKey(key)
        return True
    except:
        return False


def main():
    import argparse

    parser = argparse.ArgumentParser(description="VibeVoice Tray Application")
    parser.add_argument("--minimized", action="store_true", help="Start minimized")
    parser.add_argument("--add-startup", action="store_true", help="Add to Windows startup")
    parser.add_argument("--remove-startup", action="store_true", help="Remove from Windows startup")
    args = parser.parse_args()

    # Handle startup commands
    if args.add_startup:
        if add_to_startup(True):
            print("Successfully added to Windows startup")
        sys.exit(0)

    if args.remove_startup:
        if add_to_startup(False):
            print("Successfully removed from Windows startup")
        sys.exit(0)

    # Start the GUI
    root = tk.Tk()

    # Set icon if available
    icon_path = SCRIPT_DIR / "vibevoice.ico"
    if icon_path.exists():
        root.iconbitmap(str(icon_path))

    app = VibeVoiceTrayApp(root, start_minimized=args.minimized)

    # Bind show on taskbar click (when minimized)
    def on_map(event):
        app.show()

    root.bind("<Map>", on_map)

    root.mainloop()


if __name__ == "__main__":
    main()
