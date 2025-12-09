"""
VibeVoice SAPI Named Pipe Server

Fixes:
1. 'cannot pickle _thread.lock' crash (via SafeEvent)
2. Audio cutoffs (via Silence Padding)
3. Windows Settings visibility (via Low Integrity Security Descriptor)
4. Thread safety for concurrent requests
5. Proper error handling for empty text and invalid voices
"""

import argparse
import os
import struct
import sys
import threading
import traceback
import time
from pathlib import Path

# Add parent directory for imports
sys.path.insert(0, str(Path(__file__).parent.parent))

# Windows-specific imports
if sys.platform == "win32":
    import win32pipe
    import win32file
    import win32security
    import pywintypes
else:
    print("Error: This server is designed for Windows only.")
    sys.exit(1)


PIPE_NAME = r"\\.\pipe\vibevoice"
SAMPLE_RATE = 24000
BUFFER_SIZE = 65536

# Error codes
ERR_SUCCESS = 0
ERR_EMPTY_TEXT = 1
ERR_INVALID_VOICE = 2
ERR_MODEL_ERROR = 3
ERR_UNKNOWN = 99

# Request flags (for future extensibility)
FLAG_NONE = 0x00000000
FLAG_NO_SILENCE_PAD = 0x00000001  # Skip silence padding if client handles it


# --- FIX FOR CRASH: SafeEvent ---
class SafeEvent(threading.Event):
    """
    A wrapper around threading.Event that can be deepcopied safely.
    The transformers library tries to deepcopy the generation config, 
    and if a standard Event is inside, it crashes on the _thread.lock.
    This class prevents that by returning itself during a copy.
    """
    def __deepcopy__(self, memo):
        # Return self instead of trying to copy the lock
        return self
    
    def __reduce__(self):
        # Fallback for pickling
        return (SafeEvent, ())


class SAPIPipeServer:
    def __init__(self, model_path: str, device: str = "cuda", inference_steps: int = 5):
        self.model_path = model_path
        self.device = device
        self.inference_steps = inference_steps
        self.tts_service = None
        self.running = False
        self._lock = threading.Lock()
        
        # Security Attributes for OneCore/Settings App visibility
        self._security_attributes = self._create_low_integrity_security_attributes()

    def _create_low_integrity_security_attributes(self):
        """
        Creates a Security Descriptor that allows Low Integrity processes (AppContainers/OneCore)
        to read/write to the pipe. Required for Windows Settings visibility.
        """
        try:
            sd = win32security.SECURITY_DESCRIPTOR()
            sd.Initialize()
            sd.SetSecurityDescriptorDacl(1, None, 0) # Allow Everyone
            
            sa = win32security.SECURITY_ATTRIBUTES()
            sa.SECURITY_DESCRIPTOR = sd
            sa.bInheritHandle = 1
            return sa
        except Exception as e:
            print(f"[Warning] Security init failed: {e}")
            return None

    def _normalize_text(self, text: str) -> str:
        """Fixes newlines causing cutoffs and adds punctuation."""
        text = text.strip()
        text = text.replace("\r\n", " ").replace("\n", " ")
        text = text.replace("'", "'").replace("'", "'").replace(""", '"').replace(""", '"')

        if not text:
            return ""

        if text[-1] not in '.!?,:;':
            text = text + '.'

        return text

    def load_model(self):
        from web.app import StreamingTTSService
        print(f"[SAPI Server] Loading model from {self.model_path} on {self.device}...")
        self.tts_service = StreamingTTSService(
            model_path=self.model_path,
            device=self.device,
            inference_steps=self.inference_steps,
        )
        self.tts_service.load()
        print(f"[SAPI Server] Model loaded. Ready.")

    def create_pipe(self):
        return win32pipe.CreateNamedPipe(
            PIPE_NAME,
            win32pipe.PIPE_ACCESS_DUPLEX,
            win32pipe.PIPE_TYPE_BYTE | win32pipe.PIPE_READMODE_BYTE | win32pipe.PIPE_WAIT,
            win32pipe.PIPE_UNLIMITED_INSTANCES,
            BUFFER_SIZE,
            BUFFER_SIZE,
            0,
            self._security_attributes
        )

    def write_audio_chunk(self, pipe, chunk: bytes):
        length_data = struct.pack('<I', len(chunk))
        win32file.WriteFile(pipe, length_data)
        if len(chunk) > 0:
            win32file.WriteFile(pipe, chunk)

    def write_end_of_stream(self, pipe):
        win32file.WriteFile(pipe, struct.pack('<I', 0))

    def write_error(self, pipe, error_code: int, message: str):
        try:
            win32file.WriteFile(pipe, struct.pack('<I', 0xFFFFFFFF))
            win32file.WriteFile(pipe, struct.pack('<I', error_code))
            msg_bytes = message.encode('utf-8')[:255]
            msg_padded = msg_bytes + b'\x00' * (256 - len(msg_bytes))
            win32file.WriteFile(pipe, msg_padded)
        except:
            pass

    def _resolve_voice(self, voice_id: str) -> tuple[str | None, bool]:
        """
        Resolve voice ID to a valid voice key.
        Returns (voice_key, found) tuple.
        """
        if not voice_id:
            # No voice specified, use default (first available)
            return None, True

        # Exact match
        if voice_id in self.tts_service.voice_presets:
            return voice_id, True

        # Partial match (case-insensitive)
        for k in self.tts_service.voice_presets:
            if voice_id.lower() in k.lower():
                return k, True

        # No match found
        return None, False

    def handle_client(self, pipe):
        try:
            # --- READ REQUEST ---
            hr, data = win32file.ReadFile(pipe, 4)
            if len(data) < 4:
                return
            text_len = struct.unpack('<I', data)[0]
            if text_len == 0:
                self.write_error(pipe, ERR_EMPTY_TEXT, "Empty text length")
                return

            hr, data = win32file.ReadFile(pipe, text_len)
            text = data.decode('utf-16-le')

            hr, data = win32file.ReadFile(pipe, 32)
            voice_id = data.rstrip(b'\x00').decode('ascii', errors='ignore')

            # Read flags (4 bytes) - used for future extensibility
            hr, data = win32file.ReadFile(pipe, 4)
            flags = struct.unpack('<I', data)[0] if len(data) >= 4 else FLAG_NONE

            print(f"[Request] {text[:40]}{'...' if len(text) > 40 else ''} (voice={voice_id}, flags=0x{flags:08X})")

            # --- VALIDATE AND PREPARE ---
            text = self._normalize_text(text)

            # Check for empty text after normalization
            if not text:
                self.write_error(pipe, ERR_EMPTY_TEXT, "Text is empty after normalization")
                return

            # Resolve Voice ID (thread-safe read)
            with self._lock:
                voice_key, voice_found = self._resolve_voice(voice_id)

            if not voice_found:
                available = ", ".join(self.tts_service.voice_presets.keys())
                self.write_error(pipe, ERR_INVALID_VOICE, f"Voice '{voice_id}' not found. Available: {available}")
                return

            # USE SAFE EVENT (Fixes the crash)
            stop_event = SafeEvent()

            # --- STREAM (thread-safe) ---
            chunk_count = 0
            with self._lock:
                for audio_chunk in self.tts_service.stream(
                    text=text,
                    voice_key=voice_key,
                    stop_event=stop_event
                ):
                    pcm_bytes = self.tts_service.chunk_to_pcm16(audio_chunk)
                    self.write_audio_chunk(pipe, pcm_bytes)
                    chunk_count += 1

            # --- FIX FOR CUTOFFS: PAD SILENCE ---
            # SAPI sometimes drops the last buffer. We push 300ms of silence to flush it.
            # Can be disabled via FLAG_NO_SILENCE_PAD if client handles buffering.
            if not (flags & FLAG_NO_SILENCE_PAD):
                silence_samples = int(SAMPLE_RATE * 0.3)
                silence_bytes = b'\x00' * (silence_samples * 2)
                self.write_audio_chunk(pipe, silence_bytes)

            # End Stream
            self.write_end_of_stream(pipe)

            # Force flush to ensure SAPI gets the data before we close
            try:
                win32file.FlushFileBuffers(pipe)
            except pywintypes.error:
                pass

            print(f"[Done] Sent {chunk_count} chunks.")

        except pywintypes.error as e:
            # Pipe errors (client disconnected, etc.) - log but don't write error
            print(f"[Pipe Error] {e}")
        except Exception as e:
            print(f"[Error] {e}")
            traceback.print_exc()
            try:
                self.write_error(pipe, ERR_MODEL_ERROR, str(e)[:200])
            except:
                pass

    def run(self):
        self.running = True
        print(f"[SAPI Server] Listening on {PIPE_NAME}")
        print(f"[Info] Features: Thread-safe, Silence Padding, Low Integrity Access")
        print(f"[Info] Available voices: {', '.join(self.tts_service.voice_presets.keys())}")

        while self.running:
            try:
                pipe = self.create_pipe()
                win32pipe.ConnectNamedPipe(pipe, None)

                # Handle connection in a thread
                t = threading.Thread(target=self._handle_safe, args=(pipe,), daemon=True)
                t.start()

            except pywintypes.error as e:
                if e.winerror == 232:  # ERROR_NO_DATA - pipe closing
                    continue
                print(f"[Server Loop Error] {e}")
                time.sleep(1)
            except Exception as e:
                print(f"[Server Loop Error] {e}")
                time.sleep(1)

    def _handle_safe(self, pipe):
        try:
            self.handle_client(pipe)
        finally:
            try:
                win32pipe.DisconnectNamedPipe(pipe)
                win32file.CloseHandle(pipe)
            except:
                pass

    def stop(self):
        self.running = False

def main():
    parser = argparse.ArgumentParser(description="VibeVoice SAPI Named Pipe Server")
    parser.add_argument("--model_path", type=str, default="microsoft/VibeVoice-Realtime-0.5B",
                        help="Path to the VibeVoice model")
    parser.add_argument("--device", type=str, default="cuda",
                        help="Device to run on (cuda, cpu, mps)")
    parser.add_argument("--inference_steps", type=int, default=5,
                        help="Number of diffusion inference steps (default: 5)")
    args = parser.parse_args()

    # Change to directory of script
    os.chdir(Path(__file__).parent)

    server = SAPIPipeServer(
        model_path=args.model_path,
        device=args.device,
        inference_steps=args.inference_steps
    )
    try:
        server.load_model()
        server.run()
    except KeyboardInterrupt:
        print("\n[SAPI Server] Shutting down...")
        server.stop()


if __name__ == "__main__":
    main()