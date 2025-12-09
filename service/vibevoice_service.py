"""
VibeVoice Windows Service

This service runs the VibeVoice TTS server as a Windows service,
providing SAPI integration through a named pipe interface.

Installation:
    python vibevoice_service.py install
    python vibevoice_service.py start

Removal:
    python vibevoice_service.py stop
    python vibevoice_service.py remove

Requirements:
    pip install pywin32
"""

import os
import sys
import time
import logging
import threading
from pathlib import Path

# Add parent directory for imports
sys.path.insert(0, str(Path(__file__).parent.parent))
sys.path.insert(0, str(Path(__file__).parent.parent / "demo"))

import win32serviceutil
import win32service
import win32event
import servicemanager


# Configure logging
LOG_DIR = Path(os.environ.get("PROGRAMDATA", "C:\\ProgramData")) / "VibeVoice"
LOG_DIR.mkdir(parents=True, exist_ok=True)
LOG_FILE = LOG_DIR / "vibevoice_service.log"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE),
        logging.StreamHandler(),
    ],
)
logger = logging.getLogger("VibeVoiceService")


class VibeVoiceService(win32serviceutil.ServiceFramework):
    """Windows service for VibeVoice TTS."""

    _svc_name_ = "VibeVoiceTTS"
    _svc_display_name_ = "VibeVoice Text-to-Speech Service"
    _svc_description_ = (
        "Provides high-quality neural text-to-speech synthesis "
        "through the VibeVoice model. Used by SAPI-compatible applications."
    )

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.stop_event = win32event.CreateEvent(None, 0, 0, None)
        self.server = None
        self.server_thread = None

    def SvcStop(self):
        """Handle service stop request."""
        logger.info("Service stop requested")
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.stop_event)

        if self.server:
            self.server.stop()

    def SvcDoRun(self):
        """Main service entry point."""
        servicemanager.LogMsg(
            servicemanager.EVENTLOG_INFORMATION_TYPE,
            servicemanager.PYS_SERVICE_STARTED,
            (self._svc_name_, ""),
        )

        logger.info("VibeVoice service starting...")
        self.main()

    def main(self):
        """Main service logic."""
        try:
            # Change to demo directory for imports
            os.chdir(Path(__file__).parent.parent / "demo")

            # Import the pipe server
            from sapi_pipe_server import SAPIPipeServer

            # Configuration
            model_path = os.environ.get(
                "VIBEVOICE_MODEL_PATH",
                "microsoft/VibeVoice-Realtime-0.5B"
            )
            device = os.environ.get("VIBEVOICE_DEVICE", "cuda")
            inference_steps = int(os.environ.get("VIBEVOICE_INFERENCE_STEPS", "5"))

            logger.info(f"Configuration:")
            logger.info(f"  Model: {model_path}")
            logger.info(f"  Device: {device}")
            logger.info(f"  Inference steps: {inference_steps}")

            # Create and start the server
            self.server = SAPIPipeServer(
                model_path=model_path,
                device=device,
                inference_steps=inference_steps,
            )

            # Load model (this takes a while)
            logger.info("Loading model...")
            self.server.load_model()
            logger.info("Model loaded successfully")

            # Run server in a separate thread
            self.server_thread = threading.Thread(
                target=self.server.run,
                daemon=True,
            )
            self.server_thread.start()

            logger.info("Service is running")

            # Wait for stop event
            while True:
                result = win32event.WaitForSingleObject(self.stop_event, 1000)
                if result == win32event.WAIT_OBJECT_0:
                    break

            logger.info("Service stopped")

        except Exception as e:
            logger.error(f"Service error: {e}", exc_info=True)
            servicemanager.LogErrorMsg(f"VibeVoice service error: {e}")


def install_service():
    """Install the service with proper arguments."""
    # Get the Python executable path
    python_exe = sys.executable

    # Get the script path
    script_path = os.path.abspath(__file__)

    # Install using pythonservice.exe wrapper
    try:
        win32serviceutil.InstallService(
            VibeVoiceService._svc_name_,
            VibeVoiceService._svc_name_,
            VibeVoiceService._svc_display_name_,
            startType=win32service.SERVICE_DEMAND_START,
            description=VibeVoiceService._svc_description_,
        )
        print(f"Service '{VibeVoiceService._svc_name_}' installed successfully.")
    except Exception as e:
        print(f"Failed to install service: {e}")
        raise


if __name__ == "__main__":
    if len(sys.argv) == 1:
        # Running as service
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(VibeVoiceService)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        # Command line mode
        win32serviceutil.HandleCommandLine(VibeVoiceService)
