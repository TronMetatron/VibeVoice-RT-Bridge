"""
Test client for the VibeVoice SAPI Named Pipe Server.

This script connects to the pipe server, sends a TTS request, and saves the audio.

Usage:
    python test_pipe_client.py --text "Hello, this is a test." --voice Carter --output test.wav
"""

import argparse
import struct
import sys
import wave
from pathlib import Path

if sys.platform == "win32":
    import win32file
    import win32pipe
    import pywintypes
else:
    print("Error: This script requires Windows")
    sys.exit(1)


PIPE_NAME = r"\\.\pipe\vibevoice"
SAMPLE_RATE = 24000


def send_tts_request(text: str, voice_id: str = "", flags: int = 0) -> bytes:
    """
    Send a TTS request to the pipe server and receive audio data.

    Args:
        text: Text to synthesize
        voice_id: Voice preset ID (e.g., "en-Carter_man" or just "Carter")
        flags: Reserved flags (currently unused)

    Returns:
        bytes: PCM audio data (16-bit signed, 24kHz, mono)
    """
    # Connect to the named pipe
    try:
        pipe = win32file.CreateFile(
            PIPE_NAME,
            win32file.GENERIC_READ | win32file.GENERIC_WRITE,
            0,  # No sharing
            None,  # Default security
            win32file.OPEN_EXISTING,
            0,  # Default attributes
            None,  # No template
        )
    except pywintypes.error as e:
        if e.winerror == 2:  # ERROR_FILE_NOT_FOUND
            raise ConnectionError(
                f"Cannot connect to pipe server at {PIPE_NAME}. "
                "Make sure sapi_pipe_server.py is running."
            )
        raise

    try:
        # Encode text as UTF-16LE
        text_bytes = text.encode('utf-16-le')
        text_length = len(text_bytes)

        # Prepare voice ID (32 bytes, null-padded ASCII)
        voice_bytes = voice_id.encode('ascii')[:31]
        voice_padded = voice_bytes + b'\x00' * (32 - len(voice_bytes))

        # Build request
        request = (
            struct.pack('<I', text_length) +  # Text length
            text_bytes +                       # Text (UTF-16LE)
            voice_padded +                     # Voice ID (32 bytes)
            struct.pack('<I', flags)           # Flags
        )

        # Send request
        win32file.WriteFile(pipe, request)

        # Read audio chunks
        audio_data = bytearray()

        while True:
            # Read chunk length
            result, length_bytes = win32file.ReadFile(pipe, 4)
            if len(length_bytes) < 4:
                raise IOError("Failed to read chunk length")

            chunk_length = struct.unpack('<I', length_bytes)[0]

            # Check for end of stream
            if chunk_length == 0:
                break

            # Check for error marker
            if chunk_length == 0xFFFFFFFF:
                # Read error code
                result, error_code_bytes = win32file.ReadFile(pipe, 4)
                error_code = struct.unpack('<I', error_code_bytes)[0]

                # Read error message
                result, error_msg_bytes = win32file.ReadFile(pipe, 256)
                error_msg = error_msg_bytes.rstrip(b'\x00').decode('utf-8', errors='ignore')

                raise RuntimeError(f"Server error {error_code}: {error_msg}")

            # Read chunk data
            result, chunk_data = win32file.ReadFile(pipe, chunk_length)
            if len(chunk_data) < chunk_length:
                raise IOError(f"Failed to read chunk: expected {chunk_length}, got {len(chunk_data)}")

            audio_data.extend(chunk_data)

        return bytes(audio_data)

    finally:
        win32file.CloseHandle(pipe)


def save_wav(audio_data: bytes, output_path: str, sample_rate: int = SAMPLE_RATE):
    """Save PCM audio data as a WAV file."""
    with wave.open(output_path, 'wb') as wav_file:
        wav_file.setnchannels(1)  # Mono
        wav_file.setsampwidth(2)  # 16-bit
        wav_file.setframerate(sample_rate)
        wav_file.writeframes(audio_data)


def main():
    parser = argparse.ArgumentParser(description="Test client for VibeVoice SAPI Pipe Server")
    parser.add_argument(
        "--text",
        type=str,
        default="Hello! This is a test of the VibeVoice text to speech system.",
        help="Text to synthesize",
    )
    parser.add_argument(
        "--voice",
        type=str,
        default="Carter",
        help="Voice preset (e.g., Carter, Emma, Davis)",
    )
    parser.add_argument(
        "--output",
        type=str,
        default="pipe_test_output.wav",
        help="Output WAV file path",
    )
    args = parser.parse_args()

    print(f"Connecting to {PIPE_NAME}...")
    print(f"Text: {args.text}")
    print(f"Voice: {args.voice}")

    try:
        audio_data = send_tts_request(args.text, args.voice)
        print(f"Received {len(audio_data)} bytes of audio")

        # Calculate duration
        duration = len(audio_data) / (SAMPLE_RATE * 2)  # 2 bytes per sample
        print(f"Audio duration: {duration:.2f} seconds")

        # Save to file
        save_wav(audio_data, args.output)
        print(f"Saved to {args.output}")

    except ConnectionError as e:
        print(f"Connection error: {e}")
        sys.exit(1)
    except RuntimeError as e:
        print(f"Server error: {e}")
        sys.exit(1)
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
