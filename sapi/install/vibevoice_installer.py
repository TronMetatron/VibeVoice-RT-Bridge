"""
VibeVoice SAPI Installer & Diagnostic Tool
Run as Administrator to install, check, and manage VibeVoice TTS components.
"""

import ctypes
import os
import sys
import subprocess
import winreg
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from pathlib import Path
import threading
import time

# Constants
CLSID = "{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}"
PIPE_NAME = r"\\.\pipe\vibevoice"
SERVICE_NAME = "VibeVoiceTTS"
SERVER_SCRIPT_NAME = "sapi_pipe_server.py"

# Voice definitions
VOICES = [
    {"name": "VibeVoice Carter", "id": "en-Carter_man", "gender": "Male", "token": "VibeVoice-Carter"},
    {"name": "VibeVoice Davis", "id": "en-Davis_man", "gender": "Male", "token": "VibeVoice-Davis"},
    {"name": "VibeVoice Emma", "id": "en-Emma_woman", "gender": "Female", "token": "VibeVoice-Emma"},
    {"name": "VibeVoice Frank", "id": "en-Frank_man", "gender": "Male", "token": "VibeVoice-Frank"},
    {"name": "VibeVoice Grace", "id": "en-Grace_woman", "gender": "Female", "token": "VibeVoice-Grace"},
    {"name": "VibeVoice Mike", "id": "en-Mike_man", "gender": "Male", "token": "VibeVoice-Mike"},
    {"name": "VibeVoice Samuel", "id": "in-Samuel_man", "gender": "Male", "token": "VibeVoice-Samuel"},
]

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def run_as_admin():
    if sys.argv[0].endswith('.py'):
        script = sys.argv[0]
        params = ' '.join([f'"{arg}"' for arg in sys.argv[1:]])
        ctypes.windll.shell32.ShellExecuteW(
            None, "runas", sys.executable, f'"{script}" {params}', None, 1
        )
    sys.exit()

class VibeVoiceInstaller:
    def __init__(self, root):
        self.root = root
        self.root.title("VibeVoice SAPI Installer (Enhanced)")
        self.root.geometry("850x750")
        
        # Paths
        self.script_dir = Path(__file__).parent.absolute()
        self.sapi_dir = self.script_dir.parent
        self.project_dir = self.sapi_dir.parent
        self.dll_path = self.sapi_dir / "VibeVoiceSAPI" / "bin" / "Release" / "VibeVoiceSAPI.dll"
        # If bin/Release doesn't exist, try x64/Release (Visual Studio default)
        if not self.dll_path.exists():
            self.dll_path = self.sapi_dir / "VibeVoiceSAPI" / "x64" / "Release" / "VibeVoiceSAPI.dll"
            
        self.service_script = self.project_dir / "service" / "vibevoice_service.py"
        self.pipe_server_script = self.project_dir / "demo" / "sapi_pipe_server.py"
        self.pid_file = self.project_dir / "server.pid"

        # State
        self.server_process = None
        self.external_pid = None
        
        self.setup_ui()
        self.load_gpu_list()
        self.refresh_status()

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def setup_ui(self):
        style = ttk.Style()
        style.configure("Bold.TLabel", font=("Segoe UI", 10, "bold"))
        
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Header
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 10))
        
        title_label = ttk.Label(header_frame, text="VibeVoice SAPI Manager", font=("Segoe UI", 16, "bold"))
        title_label.pack(side=tk.LEFT)
        
        admin_text = "Running as Admin" if is_admin() else "NOT Admin (Restricted)"
        admin_color = "green" if is_admin() else "red"
        ttk.Label(header_frame, text=admin_text, foreground=admin_color, font=("Segoe UI", 10, "bold")).pack(side=tk.RIGHT)

        # Tabs
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill=tk.BOTH, expand=True, pady=5)

        self.status_tab = ttk.Frame(notebook, padding="10")
        self.actions_tab = ttk.Frame(notebook, padding="10")
        self.log_tab = ttk.Frame(notebook, padding="10")
        
        notebook.add(self.status_tab, text="Status")
        notebook.add(self.actions_tab, text="Actions")
        notebook.add(self.log_tab, text="Logs")

        self.setup_status_tab()
        self.setup_actions_tab()
        self.setup_log_tab()

        # Footer
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(10, 0))
        ttk.Button(btn_frame, text="Refresh", command=self.refresh_status).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Install Everything", command=self.install_all).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Exit", command=self.on_closing).pack(side=tk.RIGHT, padx=5)

    def setup_status_tab(self):
        # 1. DLL Status
        f1 = ttk.LabelFrame(self.status_tab, text="DLL Registration", padding=5)
        f1.pack(fill=tk.X, pady=5)
        self.lbl_dll_path = ttk.Label(f1, text="Checking...")
        self.lbl_dll_path.pack(anchor=tk.W)
        self.lbl_dll_reg = ttk.Label(f1, text="Checking...", font=("Segoe UI", 9, "bold"))
        self.lbl_dll_reg.pack(anchor=tk.W)

        # 2. Server Status
        f2 = ttk.LabelFrame(self.status_tab, text="Server Status", padding=5)
        f2.pack(fill=tk.X, pady=5)
        self.lbl_server = ttk.Label(f2, text="Checking...", font=("Segoe UI", 11, "bold"))
        self.lbl_server.pack(anchor=tk.W)
        self.lbl_pipe = ttk.Label(f2, text="Checking pipe...")
        self.lbl_pipe.pack(anchor=tk.W)
        self.lbl_pid = ttk.Label(f2, text="")
        self.lbl_pid.pack(anchor=tk.W)

        # 3. Voices
        f3 = ttk.LabelFrame(self.status_tab, text="Registered Voices", padding=5)
        f3.pack(fill=tk.BOTH, expand=True, pady=5)
        
        cols = ("Name", "ID", "Status")
        self.tree = ttk.Treeview(f3, columns=cols, show="headings", height=8)
        self.tree.heading("Name", text="Voice Name")
        self.tree.heading("ID", text="Model ID")
        self.tree.heading("Status", text="Registry Status")
        self.tree.column("Name", width=150)
        self.tree.column("ID", width=150)
        self.tree.column("Status", width=100)
        self.tree.pack(fill=tk.BOTH, expand=True)

    def setup_actions_tab(self):
        # Server Controls
        f_server = ttk.LabelFrame(self.actions_tab, text="Server Control", padding=10)
        f_server.pack(fill=tk.X, pady=5)

        # GPU Selection
        row_frame = ttk.Frame(f_server)
        row_frame.pack(fill=tk.X, pady=5)
        ttk.Label(row_frame, text="Select GPU:").pack(side=tk.LEFT)
        
        self.gpu_var = tk.StringVar(value="cuda:0")
        self.gpu_combo = ttk.Combobox(row_frame, textvariable=self.gpu_var, state="readonly", width=40)
        self.gpu_combo.pack(side=tk.LEFT, padx=10)
        
        btn_frame = ttk.Frame(f_server)
        btn_frame.pack(fill=tk.X, pady=5)
        self.btn_start = ttk.Button(btn_frame, text="▶ Start Server", command=self.start_server)
        self.btn_start.pack(side=tk.LEFT, padx=5)
        self.btn_stop = ttk.Button(btn_frame, text="■ Stop Server", command=self.stop_server)
        self.btn_stop.pack(side=tk.LEFT, padx=5)

        # DLL Controls
        f_dll = ttk.LabelFrame(self.actions_tab, text="Registry / DLL", padding=10)
        f_dll.pack(fill=tk.X, pady=5)
        ttk.Button(f_dll, text="Register DLL", command=self.register_dll).pack(side=tk.LEFT, padx=5)
        ttk.Button(f_dll, text="Unregister DLL", command=self.unregister_dll).pack(side=tk.LEFT, padx=5)
        ttk.Button(f_dll, text="Register Voices", command=self.register_voices).pack(side=tk.LEFT, padx=5)
        ttk.Button(f_dll, text="Fix/Clean Registry", command=self.fix_registry).pack(side=tk.LEFT, padx=5)

        # Startup Controls
        f_startup = ttk.LabelFrame(self.actions_tab, text="Startup", padding=10)
        f_startup.pack(fill=tk.X, pady=5)

        self.startup_status_label = ttk.Label(f_startup, text="Checking...")
        self.startup_status_label.pack(side=tk.LEFT, padx=5)

        self.btn_add_startup = ttk.Button(f_startup, text="Enable Auto-Start", command=self.add_to_startup)
        self.btn_add_startup.pack(side=tk.LEFT, padx=5)

        self.btn_remove_startup = ttk.Button(f_startup, text="Disable Auto-Start", command=self.remove_from_startup)
        self.btn_remove_startup.pack(side=tk.LEFT, padx=5)

        ttk.Button(f_startup, text="Open Tray App", command=self.open_tray_app).pack(side=tk.LEFT, padx=5)

        # Tools
        f_tools = ttk.LabelFrame(self.actions_tab, text="Tools", padding=10)
        f_tools.pack(fill=tk.X, pady=5)
        ttk.Button(f_tools, text="Open Speech Settings", command=self.open_settings).pack(side=tk.LEFT, padx=5)
        ttk.Button(f_tools, text="Open Legacy Control Panel", command=self.open_legacy_cpl).pack(side=tk.LEFT, padx=5)

    def setup_log_tab(self):
        self.log_text = scrolledtext.ScrolledText(self.log_tab, height=15)
        self.log_text.pack(fill=tk.BOTH, expand=True)

    def log(self, msg):
        ts = time.strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{ts}] {msg}\n")
        self.log_text.see(tk.END)

    # --- GPU DETECTION ---
    def load_gpu_list(self):
        """Detect GPUs using nvidia-smi."""
        gpus = ["cpu"]
        try:
            # Run nvidia-smi to get index and name
            cmd = ["nvidia-smi", "--query-gpu=index,name", "--format=csv,noheader"]
            result = subprocess.run(cmd, capture_output=True, text=True)
            if result.returncode == 0:
                lines = result.stdout.strip().split('\n')
                for line in lines:
                    if line.strip():
                        idx, name = line.split(',', 1)
                        # Format: "cuda:0 (RTX 5090)"
                        gpus.append(f"cuda:{idx.strip()} ({name.strip()})")
        except Exception:
            # Fallback if nvidia-smi fails
            gpus.extend(["cuda:0", "cuda:1"])
            
        # Update combo box
        # Remove duplicates and prioritize CUDA
        final_list = sorted(list(set(gpus)), key=lambda x: 0 if "cuda" in x else 1)
        self.gpu_combo['values'] = final_list
        if len(final_list) > 0:
            self.gpu_combo.current(0)
        self.log(f"Detected Devices: {final_list}")

    # --- PROCESS MANAGEMENT (THE FIX) ---
    def get_saved_pid(self):
        """Read PID from server.pid file."""
        if self.pid_file.exists():
            try:
                return int(self.pid_file.read_text().strip())
            except:
                return None
        return None

    def is_process_running(self, pid):
        """Check if a process with PID is actually running."""
        if not pid: return False
        try:
            # Use tasklist to filter by PID
            cmd = ["tasklist", "/FI", f"PID eq {pid}", "/FO", "CSV"]
            out = subprocess.check_output(cmd).decode()
            return str(pid) in out
        except:
            return False

    def refresh_status(self):
        # 1. Check DLL
        if self.dll_path.exists():
            self.lbl_dll_path.config(text=f"DLL Found: {self.dll_path.name}", foreground="green")
        else:
            self.lbl_dll_path.config(text="DLL NOT FOUND (Build Solution in Release x64)", foreground="red")

        # 2. Check Registry
        if self.check_com_registered():
            self.lbl_dll_reg.config(text="COM Registered: YES", foreground="green")
        else:
            self.lbl_dll_reg.config(text="COM Registered: NO", foreground="red")

        # 3. Check Server (Robust)
        saved_pid = self.get_saved_pid()
        is_running = False
        
        if saved_pid and self.is_process_running(saved_pid):
            self.external_pid = saved_pid
            is_running = True
        else:
            self.external_pid = None
            if self.pid_file.exists():
                try: self.pid_file.unlink() # Clean up stale file
                except: pass

        # 4. Check Pipe
        pipe_ok = False
        try:
            import win32file
            h = win32file.CreateFile(PIPE_NAME, win32file.GENERIC_READ, 0, None, 3, 0, None)
            win32file.CloseHandle(h)
            pipe_ok = True
        except:
            pass

        if is_running:
            self.lbl_server.config(text="✓ Server Running", foreground="green")
            self.lbl_pid.config(text=f"PID: {self.external_pid}")
            self.btn_start.state(['disabled'])
            self.btn_stop.state(['!disabled'])
        else:
            self.lbl_server.config(text="✗ Server Stopped", foreground="red")
            self.lbl_pid.config(text="")
            self.btn_start.state(['!disabled'])
            self.btn_stop.state(['disabled'])

        if pipe_ok:
            self.lbl_pipe.config(text="✓ Pipe Connected", foreground="green")
        else:
            self.lbl_pipe.config(text="Waiting for pipe...", foreground="orange")

        # 5. Voices
        self.refresh_voices_tree()

        # 6. Startup status
        self.refresh_startup_status()

    def start_server(self):
        if not self.pipe_server_script.exists():
            messagebox.showerror("Error", "Server script not found.")
            return

        # Get selected device from string (e.g. "cuda:0 (RTX 5090)" -> "cuda:0")
        raw_selection = self.gpu_var.get()
        device_arg = raw_selection.split(' ')[0]

        self.log(f"Starting server on {device_arg}...")

        try:
            # CREATE_NEW_PROCESS_GROUP is important for detached processes
            self.server_process = subprocess.Popen(
                [sys.executable, str(self.pipe_server_script), "--device", device_arg],
                cwd=str(self.project_dir),
                creationflags=subprocess.CREATE_NEW_PROCESS_GROUP
            )
            
            # Save PID immediately
            self.pid_file.write_text(str(self.server_process.pid))
            self.log(f"Server started with PID {self.server_process.pid}")
            
            # Start a thread to monitor output (optional, keeping it simple here)
            time.sleep(1)
            self.refresh_status()
            
        except Exception as e:
            self.log(f"Error starting: {e}")
            messagebox.showerror("Error", str(e))

    def stop_server(self):
        # Prefer internal handle, fallback to file
        pid_to_kill = None
        
        if self.server_process:
            pid_to_kill = self.server_process.pid
        elif self.external_pid:
            pid_to_kill = self.external_pid

        if pid_to_kill:
            self.log(f"Stopping PID {pid_to_kill}...")
            try:
                # Force kill
                subprocess.run(["taskkill", "/F", "/PID", str(pid_to_kill)], capture_output=True)
                self.server_process = None
                self.external_pid = None
                if self.pid_file.exists():
                    self.pid_file.unlink()
                self.log("Server stopped.")
            except Exception as e:
                self.log(f"Failed to kill: {e}")
        
        self.refresh_status()

    # --- REGISTRY HELPERS ---
    def check_com_registered(self):
        try:
            winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, f"SOFTWARE\\Classes\\CLSID\\{CLSID}")
            return True
        except:
            return False

    def register_dll(self):
        if not is_admin(): return messagebox.showerror("Admin", "Run as Administrator required.")
        ret = subprocess.call(["regsvr32", "/s", str(self.dll_path)])
        if ret == 0: self.log("DLL Registered.")
        else: self.log("DLL Registration Failed.")
        self.refresh_status()

    def unregister_dll(self):
        if not is_admin(): return messagebox.showerror("Admin", "Run as Administrator required.")
        subprocess.call(["regsvr32", "/u", "/s", str(self.dll_path)])
        self.log("DLL Unregistered.")
        self.refresh_status()

    def register_voices(self):
        if not is_admin(): return messagebox.showerror("Admin", "Run as Administrator required.")
        self.log("Registering voices...")
        try:
            # Register in both legacy SAPI and OneCore (Windows 11) locations
            registry_bases = [
                "SOFTWARE\\Microsoft\\Speech\\Voices\\Tokens",           # Legacy SAPI5
                "SOFTWARE\\Microsoft\\Speech_OneCore\\Voices\\Tokens",   # Windows 11 OneCore
            ]

            for base in registry_bases:
                for v in VOICES:
                    try:
                        key = winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, f"{base}\\{v['token']}")
                        winreg.SetValueEx(key, "", 0, winreg.REG_SZ, v['name'])
                        winreg.SetValueEx(key, "CLSID", 0, winreg.REG_SZ, CLSID)
                        winreg.SetValueEx(key, "VoiceId", 0, winreg.REG_SZ, v['id'])

                        attr = winreg.CreateKey(key, "Attributes")
                        winreg.SetValueEx(attr, "Name", 0, winreg.REG_SZ, v['name'])
                        winreg.SetValueEx(attr, "Gender", 0, winreg.REG_SZ, v['gender'])
                        winreg.SetValueEx(attr, "Language", 0, winreg.REG_SZ, "409")
                        # Additional attributes for OneCore compatibility
                        winreg.SetValueEx(attr, "Age", 0, winreg.REG_SZ, "Adult")
                        winreg.SetValueEx(attr, "Vendor", 0, winreg.REG_SZ, "VibeVoice")
                        winreg.CloseKey(attr)
                        winreg.CloseKey(key)
                    except Exception as e:
                        self.log(f"Warning: Could not register {v['token']} in {base}: {e}")

            self.log("Voices registered in legacy SAPI and OneCore locations.")
        except Exception as e:
            self.log(f"Error: {e}")
        self.refresh_status()

    def fix_registry(self):
        # Remove old stuff
        self.unregister_dll()
        self.register_dll()
        self.register_voices()
        self.log("Registry reset complete.")

    def refresh_voices_tree(self):
        for i in self.tree.get_children(): self.tree.delete(i)
        for v in VOICES:
            try:
                winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, f"SOFTWARE\\Microsoft\\Speech\\Voices\\Tokens\\{v['token']}")
                status = "OK"
            except:
                status = "Missing"
            self.tree.insert("", "end", values=(v['name'], v['id'], status))

    def install_all(self):
        self.register_dll()
        self.register_voices()
        self.start_server()

    def open_settings(self):
        os.system("start ms-settings:speech")

    def open_legacy_cpl(self):
        # Opens the old specific SAPI control panel
        # Use rundll32 to properly launch .cpl files
        cpl_path = r'C:\Windows\System32\Speech\SpeechUX\sapi.cpl'
        if os.path.exists(cpl_path):
            subprocess.Popen(['rundll32.exe', 'shell32.dll,Control_RunDLL', cpl_path])
        else:
            # Try alternative method
            subprocess.Popen(['control.exe', 'speech'], shell=True)
            self.log("sapi.cpl not found, trying control panel speech")

    # --- STARTUP MANAGEMENT ---
    def is_in_startup(self) -> bool:
        """Check if tray app is in Windows startup."""
        try:
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Run",
                0, winreg.KEY_READ
            )
            winreg.QueryValueEx(key, "VibeVoiceTTS")
            winreg.CloseKey(key)
            return True
        except:
            return False

    def refresh_startup_status(self):
        """Update startup status display."""
        if self.is_in_startup():
            self.startup_status_label.config(text="Auto-Start: ON", foreground="green")
            self.btn_add_startup.state(['disabled'])
            self.btn_remove_startup.state(['!disabled'])
        else:
            self.startup_status_label.config(text="Auto-Start: OFF", foreground="gray")
            self.btn_add_startup.state(['!disabled'])
            self.btn_remove_startup.state(['disabled'])

    def add_to_startup(self):
        """Add tray app to Windows startup."""
        try:
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Run",
                0, winreg.KEY_ALL_ACCESS
            )
            tray_script = self.script_dir / "vibevoice_tray.py"
            cmd = f'"{sys.executable}" "{tray_script}" --minimized'
            winreg.SetValueEx(key, "VibeVoiceTTS", 0, winreg.REG_SZ, cmd)
            winreg.CloseKey(key)
            self.log("Added to Windows startup")
            self.refresh_startup_status()
        except Exception as e:
            self.log(f"Failed to add to startup: {e}")
            messagebox.showerror("Error", f"Failed to add to startup: {e}")

    def remove_from_startup(self):
        """Remove tray app from Windows startup."""
        try:
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Run",
                0, winreg.KEY_ALL_ACCESS
            )
            winreg.DeleteValue(key, "VibeVoiceTTS")
            winreg.CloseKey(key)
            self.log("Removed from Windows startup")
            self.refresh_startup_status()
        except FileNotFoundError:
            self.log("Not in startup (already removed)")
            self.refresh_startup_status()
        except Exception as e:
            self.log(f"Failed to remove from startup: {e}")

    def open_tray_app(self):
        """Open the tray app."""
        tray_script = self.script_dir / "vibevoice_tray.py"
        if tray_script.exists():
            subprocess.Popen([sys.executable, str(tray_script)], cwd=str(self.script_dir))
            self.log("Opened tray app")
        else:
            messagebox.showerror("Error", "Tray app not found")

    def on_closing(self):
        # Don't kill server on exit automatically, just update state
        self.root.destroy()

if __name__ == "__main__":
    if not is_admin():
        # Re-run as admin
        run_as_admin()
    
    root = tk.Tk()
    app = VibeVoiceInstaller(root)
    root.mainloop()