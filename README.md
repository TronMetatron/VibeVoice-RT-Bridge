<div align="center">

# VibeVoice-RT-Bridge

**Windows SAPI5 Bridge for Microsoft's VibeVoice Realtime TTS**

[![Based on VibeVoice](https://img.shields.io/badge/Based%20on-Microsoft%20VibeVoice-blue?logo=microsoft)](https://github.com/microsoft/VibeVoice)
[![Windows](https://img.shields.io/badge/Platform-Windows-0078D6?logo=windows)](https://www.microsoft.com/windows)
[![SAPI5](https://img.shields.io/badge/Interface-SAPI5-green)](https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ee125663(v=vs.85))

</div>

## What is this?

This project adds **Windows SAPI5 integration** to [Microsoft's VibeVoice-Realtime-0.5B](https://github.com/microsoft/VibeVoice) model, allowing any SAPI-compatible Windows application to use VibeVoice voices for text-to-speech.

**Features:**
- SAPI5 COM DLL bridge (C++) that connects to the VibeVoice model
- Named pipe server for efficient IPC between SAPI and Python
- System tray app for auto-starting on Windows login
- Installer/Manager UI for easy setup
- 7 high-quality voices: Carter, Davis, Emma, Frank, Grace, Mike, Samuel

## Architecture

```
Windows App (SAPI) --> VibeVoiceSAPI.dll --> Named Pipe --> Python Server --> VibeVoice Model (GPU)
```

## Quick Start

1. **Install dependencies:**
   ```bash
   pip install -e .
   ```

2. **Start the SAPI server:**
   ```bash
   python demo/sapi_pipe_server.py --device cuda:0
   ```

3. **Run the installer** (as Administrator) to register voices:
   ```bash
   python sapi/install/vibevoice_installer.py
   ```

4. **Enable auto-start** (optional):
   ```bash
   python sapi/install/vibevoice_tray.py --add-startup
   ```

## Credits

- **VibeVoice Model**: [Microsoft Research](https://github.com/microsoft/VibeVoice) - The underlying TTS model
- **SAPI Bridge**: [TronMetatron](https://github.com/TronMetatron) - Windows SAPI5 integration, DLL, pipe server, tray app

---

<details>
<summary><strong>Original VibeVoice README (click to expand)</strong></summary>

<div align="center">

## VibeVoice: Open-Source Frontier Voice AI
[![Project Page](https://img.shields.io/badge/Project-Page-blue?logo=microsoft)](https://microsoft.github.io/VibeVoice)
[![Hugging Face](https://img.shields.io/badge/HuggingFace-Collection-orange?logo=huggingface)](https://huggingface.co/collections/microsoft/vibevoice-68a2ef24a875c44be47b034f)
[![Technical Report](https://img.shields.io/badge/Technical-Report-red?logo=adobeacrobatreader)](https://arxiv.org/pdf/2508.19205)

</div>

<div align="center">
<picture>
  <source media="(prefers-color-scheme: dark)" srcset="Figures/VibeVoice_logo_white.png">
  <img src="Figures/VibeVoice_logo.png" alt="VibeVoice Logo" width="300">
</picture>
</div>


### Overview

VibeVoice is a novel framework designed for generating **expressive**, **long-form**, **multi-speaker** conversational audio, such as podcasts, from text. It addresses significant challenges in traditional Text-to-Speech (TTS) systems, particularly in scalability, speaker consistency, and natural turn-taking.

VibeVoice currently includes two model variants:

- **Long-form multi-speaker model**: Synthesizes conversational/single-speaker speech up to **90 minutes** with up to **4 distinct speakers**, surpassing the typical 1–2 speaker limits of many prior models.
- **[Realtime streaming TTS model](docs/vibevoice-realtime-0.5b.md)**: Produces initial audible speech in ~**300 ms** and supports **streaming text input** for single-speaker **real-time** speech generation; designed for low-latency generation.

A core innovation of VibeVoice is its use of continuous speech tokenizers (Acoustic and Semantic) operating at an ultra-low frame rate of 7.5 Hz. These tokenizers efficiently preserve audio fidelity while significantly boosting computational efficiency for processing long sequences. VibeVoice employs a [next-token diffusion](https://arxiv.org/abs/2412.08635) framework, leveraging a Large Language Model (LLM) to understand textual context and dialogue flow, and a diffusion head to generate high-fidelity acoustic details.


<p align="left">
  <img src="Figures/MOS-preference.png" alt="MOS Preference Results" height="260px">
  <img src="Figures/VibeVoice.jpg" alt="VibeVoice Overview" height="250px" style="margin-right: 10px;">
</p>


### 🎵 Demo Examples


**Video Demo**

We produced this video with [Wan2.2](https://github.com/Wan-Video/Wan2.2). We sincerely appreciate the Wan-Video team for their great work.

**English**
<div align="center">

https://github.com/user-attachments/assets/0967027c-141e-4909-bec8-091558b1b784

</div>


**Chinese**
<div align="center">

https://github.com/user-attachments/assets/322280b7-3093-4c67-86e3-10be4746c88f

</div>

**Cross-Lingual**
<div align="center">

https://github.com/user-attachments/assets/838d8ad9-a201-4dde-bb45-8cd3f59ce722

</div>

**Spontaneous Singing**
<div align="center">

https://github.com/user-attachments/assets/6f27a8a5-0c60-4f57-87f3-7dea2e11c730

</div>


**Long Conversation with 4 people**
<div align="center">

https://github.com/user-attachments/assets/a357c4b6-9768-495c-a576-1618f6275727

</div>

For more examples, see the [Project Page](https://microsoft.github.io/VibeVoice).



## Risks and limitations

While efforts have been made to optimize it through various techniques, it may still produce outputs that are unexpected, biased, or inaccurate. VibeVoice inherits any biases, errors, or omissions produced by its base model (specifically, Qwen2.5 1.5b in this release).
Potential for Deepfakes and Disinformation: High-quality synthetic speech can be misused to create convincing fake audio content for impersonation, fraud, or spreading disinformation. Users must ensure transcripts are reliable, check content accuracy, and avoid using generated content in misleading ways. Users are expected to use the generated content and to deploy the models in a lawful manner, in full compliance with all applicable laws and regulations in the relevant jurisdictions. It is best practice to disclose the use of AI when sharing AI-generated content.

English and Chinese only: Transcripts in languages other than English or Chinese may result in unexpected audio outputs.

Non-Speech Audio: The model focuses solely on speech synthesis and does not handle background noise, music, or other sound effects.

Overlapping Speech: The current model does not explicitly model or generate overlapping speech segments in conversations.

We do not recommend using VibeVoice in commercial or real-world applications without further testing and development. This model is intended for research and development purposes only. Please use responsibly.

</details>
