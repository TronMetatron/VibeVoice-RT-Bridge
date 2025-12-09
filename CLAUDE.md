# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

VibeVoice-Realtime is a lightweight (0.5B parameter) real-time text-to-speech model supporting streaming text input and long-form speech generation. It produces initial audible speech in ~300ms and is designed for low-latency generation. This repository contains the streaming/realtime variant only (single speaker, English).

## Common Commands

### Installation
```bash
pip install -e .
```

### Run WebSocket Demo
```bash
python demo/vibevoice_realtime_demo.py --model_path microsoft/VibeVoice-Realtime-0.5B --port 3000 --device cuda
```

On Windows with batch file:
```batch
run_demo.bat
```

### Inference from File
```bash
python demo/realtime_model_inference_from_file.py \
    --model_path microsoft/VibeVoice-Realtime-0.5B \
    --txt_path demo/text_examples/1p_vibevoice.txt \
    --speaker_name Carter \
    --device cuda
```

## Architecture

### Core Components

**Model (`vibevoice/modular/`):**
- `modeling_vibevoice_streaming_inference.py` - Main inference model (`VibeVoiceStreamingForConditionalGenerationInference`). Uses a staged forward approach: `forward_lm()` for base text LM, `forward_tts_lm()` for TTS LM with diffusion, and `generate()` for full streaming pipeline
- `modeling_vibevoice_streaming.py` - Base model definition (`VibeVoiceStreamingModel`)
- `configuration_vibevoice_streaming.py` - Model configuration
- `modular_vibevoice_diffusion_head.py` - Diffusion head for speech generation
- `modular_vibevoice_tokenizer.py` - Acoustic tokenizer with streaming cache support
- `streamer.py` - `AudioStreamer` for real-time audio chunk delivery

**Processor (`vibevoice/processor/`):**
- `vibevoice_streaming_processor.py` - `VibeVoiceStreamingProcessor` handles text tokenization and audio I/O
- `vibevoice_tokenizer_processor.py` - Audio normalization and file handling

**Demo (`demo/`):**
- `vibevoice_realtime_demo.py` - Launches FastAPI/uvicorn WebSocket server
- `web/app.py` - WebSocket endpoint (`/stream`) and `StreamingTTSService` class
- `realtime_model_inference_from_file.py` - Batch inference from text files
- `VoiceMapper` class handles speaker name to voice file path mapping

### Generation Pipeline

The model uses windowed text prefill with interleaved speech generation:
1. Text is fed in windows of 5 tokens (`TTS_TEXT_WINDOW_SIZE`)
2. After each text window, 6 speech latents are sampled (`TTS_SPEECH_WINDOW_SIZE`)
3. Speech latents are decoded to audio chunks via acoustic tokenizer
4. Audio chunks are streamed via `AudioStreamer` if provided
5. Binary EOS classifier determines when to stop generation

### Key Constants
- Sample rate: 24kHz
- Acoustic tokenizer frame rate: 7.5 Hz
- Default diffusion inference steps: 5
- Default CFG scale: 1.5

## Device Support

- **CUDA**: Uses bfloat16 + flash_attention_2 (recommended)
- **MPS (Apple Silicon)**: Uses float32 + SDPA
- **CPU**: Uses float32 + SDPA

## Voice Presets

Voice presets are `.pt` files in `demo/voices/streaming_model/` containing pre-computed KV cache for voice prompts. Available voices: Carter, Davis, Emma, Frank, Grace, Mike (English), Samuel (Indian English).

## WebSocket API

Endpoint: `ws://localhost:3000/stream`

Query parameters:
- `text` - Text to synthesize
- `voice` - Voice preset name (e.g., "en-Carter_man")
- `cfg` - CFG scale (default: 1.5)
- `steps` - Diffusion inference steps (default: 5)

REST endpoints:
- `GET /` - Web UI
- `GET /config` - Returns available voices and default voice

## Limitations

- English only (other languages produce unpredictable results)
- Very short inputs (<3 words) may cause instability
- Does not support code, mathematical formulas, or uncommon symbols
- Single speaker only (use multi-speaker variants for conversations)

## SAPI Integration

The `sapi/` directory contains a Windows SAPI5 TTS engine that allows any SAPI-compatible application to use VibeVoice voices.

### Architecture

```
SAPI Application → VibeVoiceSAPI.dll (C++) → Named Pipe → Python Server → VibeVoice Model
```

### Components

**C++ SAPI DLL (`sapi/VibeVoiceSAPI/`):**
- `VibeVoiceSAPI.cpp` - Implements `ISpTTSEngine` and `ISpObjectWithToken` interfaces
- `VibeVoiceSAPI.h` - Header with `PipeClient` class for named pipe communication
- Communicates with Python server via `\\.\pipe\vibevoice`

**Python Pipe Server (`demo/sapi_pipe_server.py`):**
- Named pipe server using win32pipe
- Reuses `StreamingTTSService` from `web/app.py`
- Protocol: 4-byte length prefix + UTF-16LE text + 32-byte voice ID

**Windows Service (`service/vibevoice_service.py`):**
- Runs pipe server as a Windows service
- Install: `python vibevoice_service.py install`
- Start: `net start VibeVoiceTTS`

### Installation

1. Build the DLL in Visual Studio (Release x64)
2. Run `sapi/install/install.bat` as Administrator
3. Start the service or run `run_sapi_server.bat`

### Testing

```bash
# Start the pipe server manually
python demo/sapi_pipe_server.py --model_path microsoft/VibeVoice-Realtime-0.5B

# Test with pipe client
python demo/test_pipe_client.py --text "Hello world" --voice Carter --output test.wav
```

### Registered Voices

After installation, these voices appear in Windows Settings > Speech:
- VibeVoice Carter (Male)
- VibeVoice Davis (Male)
- VibeVoice Emma (Female)
- VibeVoice Frank (Male)
- VibeVoice Grace (Female)
- VibeVoice Mike (Male)
- VibeVoice Samuel (Male)

## Dependencies

Key dependencies from pyproject.toml:
- transformers==4.51.3 (specific version required)
- torch, accelerate, diffusers
- fastapi, uvicorn (for WebSocket demo)
- flash-attn (optional, for CUDA acceleration)
- pywin32 (for SAPI pipe server and Windows service)
