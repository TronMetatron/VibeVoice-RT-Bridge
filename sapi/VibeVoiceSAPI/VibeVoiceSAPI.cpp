// VibeVoiceSAPI.cpp
// SAPI5 TTS Engine implementation for VibeVoice

#include "VibeVoiceSAPI.h"
#include <strsafe.h>

// Instantiate GUIDs - define storage for our custom GUIDs
// {A1B2C3D4-E5F6-7890-ABCD-EF1234567890}
const GUID CLSID_VibeVoiceTTSEngine_Impl =
    { 0xa1b2c3d4, 0xe5f6, 0x7890, { 0xab, 0xcd, 0xef, 0x12, 0x34, 0x56, 0x78, 0x90 } };

// {A1B2C3D4-E5F6-7890-ABCD-EF1234567891}
const GUID LIBID_VibeVoiceSAPILib_Impl =
    { 0xa1b2c3d4, 0xe5f6, 0x7890, { 0xab, 0xcd, 0xef, 0x12, 0x34, 0x56, 0x78, 0x91 } };

// ATL Module instance
CVibeVoiceSAPIModule _AtlModule;

// Register the COM object with ATL - required for DllRegisterServer to work
OBJECT_ENTRY_AUTO(CLSID_VibeVoiceTTSEngine, CVibeVoiceTTSEngine)

//=============================================================================
// PipeClient Implementation
//=============================================================================

PipeClient::PipeClient()
    : m_hPipe(INVALID_HANDLE_VALUE)
{
}

PipeClient::~PipeClient()
{
    Disconnect();
}

HRESULT PipeClient::Connect()
{
    if (m_hPipe != INVALID_HANDLE_VALUE) {
        return S_OK;  // Already connected
    }

    // Wait for pipe to be available
    if (!WaitNamedPipeW(PIPE_NAME, PIPE_TIMEOUT_MS)) {
        DWORD err = GetLastError();
        if (err == ERROR_SEM_TIMEOUT) {
            return HRESULT_FROM_WIN32(ERROR_TIMEOUT);
        }
        return HRESULT_FROM_WIN32(err);
    }

    // Open the pipe
    m_hPipe = CreateFileW(
        PIPE_NAME,
        GENERIC_READ | GENERIC_WRITE,
        0,              // No sharing
        NULL,           // Default security
        OPEN_EXISTING,
        0,              // Default attributes
        NULL            // No template
    );

    if (m_hPipe == INVALID_HANDLE_VALUE) {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    // Set pipe to message mode (optional, we use byte mode)
    DWORD mode = PIPE_READMODE_BYTE;
    SetNamedPipeHandleState(m_hPipe, &mode, NULL, NULL);

    return S_OK;
}

void PipeClient::Disconnect()
{
    if (m_hPipe != INVALID_HANDLE_VALUE) {
        CloseHandle(m_hPipe);
        m_hPipe = INVALID_HANDLE_VALUE;
    }
}

HRESULT PipeClient::ReadExact(void* buffer, DWORD count)
{
    BYTE* ptr = static_cast<BYTE*>(buffer);
    DWORD remaining = count;

    while (remaining > 0) {
        DWORD bytesRead = 0;
        if (!ReadFile(m_hPipe, ptr, remaining, &bytesRead, NULL)) {
            return HRESULT_FROM_WIN32(GetLastError());
        }
        if (bytesRead == 0) {
            return HRESULT_FROM_WIN32(ERROR_BROKEN_PIPE);
        }
        ptr += bytesRead;
        remaining -= bytesRead;
    }

    return S_OK;
}

HRESULT PipeClient::WriteExact(const void* buffer, DWORD count)
{
    const BYTE* ptr = static_cast<const BYTE*>(buffer);
    DWORD remaining = count;

    while (remaining > 0) {
        DWORD bytesWritten = 0;
        if (!WriteFile(m_hPipe, ptr, remaining, &bytesWritten, NULL)) {
            return HRESULT_FROM_WIN32(GetLastError());
        }
        ptr += bytesWritten;
        remaining -= bytesWritten;
    }

    return S_OK;
}

HRESULT PipeClient::StreamTTS(
    LPCWSTR text,
    LPCSTR voiceId,
    AudioChunkCallback callback,
    void* callbackContext,
    volatile bool* cancelFlag)
{
    HRESULT hr;

    // Connect if not already connected
    hr = Connect();
    if (FAILED(hr)) {
        return hr;
    }

    // Prepare text as UTF-16LE
    size_t textLen = wcslen(text);
    DWORD textBytes = static_cast<DWORD>(textLen * sizeof(wchar_t));

    // Prepare voice ID (32 bytes, null-padded)
    char voiceIdPadded[32] = { 0 };
    if (voiceId) {
        StringCchCopyA(voiceIdPadded, 32, voiceId);
    }

    // Build request
    // [4 bytes] text_length
    // [N bytes] text (UTF-16LE)
    // [32 bytes] voice_id
    // [4 bytes] flags

    hr = WriteExact(&textBytes, 4);
    if (FAILED(hr)) goto cleanup;

    hr = WriteExact(text, textBytes);
    if (FAILED(hr)) goto cleanup;

    hr = WriteExact(voiceIdPadded, 32);
    if (FAILED(hr)) goto cleanup;

    {
        DWORD flags = 0;
        hr = WriteExact(&flags, 4);
        if (FAILED(hr)) goto cleanup;
    }

    // Read audio chunks
    while (true) {
        // Check for cancellation
        if (cancelFlag && *cancelFlag) {
            hr = E_ABORT;
            goto cleanup;
        }

        // Read chunk length
        DWORD chunkLength = 0;
        hr = ReadExact(&chunkLength, 4);
        if (FAILED(hr)) goto cleanup;

        // End of stream
        if (chunkLength == 0) {
            break;
        }

        // Error marker
        if (chunkLength == 0xFFFFFFFF) {
            DWORD errorCode = 0;
            char errorMsg[256] = { 0 };

            ReadExact(&errorCode, 4);
            ReadExact(errorMsg, 256);

            // Map error codes to HRESULTs
            switch (errorCode) {
            case ERR_EMPTY_TEXT:
                hr = E_INVALIDARG;
                break;
            case ERR_INVALID_VOICE:
                hr = SPERR_VOICE_NOT_FOUND;
                break;
            case ERR_MODEL_ERROR:
                hr = E_FAIL;
                break;
            default:
                hr = E_UNEXPECTED;
                break;
            }
            goto cleanup;
        }

        // Sanity check chunk size
        if (chunkLength > PIPE_BUFFER_SIZE * 10) {
            hr = E_UNEXPECTED;
            goto cleanup;
        }

        // Read chunk data
        std::vector<BYTE> chunkData(chunkLength);
        hr = ReadExact(chunkData.data(), chunkLength);
        if (FAILED(hr)) goto cleanup;

        // Call the callback with the audio data
        if (callback) {
            callback(chunkData.data(), chunkLength, callbackContext);
        }
    }

    hr = S_OK;

cleanup:
    Disconnect();
    return hr;
}


//=============================================================================
// CVibeVoiceTTSEngine Implementation
//=============================================================================

CVibeVoiceTTSEngine::CVibeVoiceTTSEngine()
{
}

CVibeVoiceTTSEngine::~CVibeVoiceTTSEngine()
{
}

//-----------------------------------------------------------------------------
// ISpObjectWithToken::SetObjectToken
// Called by SAPI when the engine is created, provides access to registry data
//-----------------------------------------------------------------------------
STDMETHODIMP CVibeVoiceTTSEngine::SetObjectToken(ISpObjectToken* pToken)
{
    if (!pToken) {
        return E_INVALIDARG;
    }

    m_cpToken = pToken;

    // Read voice ID from registry (stored under the token's Attributes key)
    CSpDynamicString dstrVoiceId;
    HRESULT hr = pToken->GetStringValue(L"VoiceId", &dstrVoiceId);
    if (SUCCEEDED(hr) && dstrVoiceId) {
        // Convert to ASCII for the pipe protocol
        int len = WideCharToMultiByte(CP_ACP, 0, dstrVoiceId, -1, NULL, 0, NULL, NULL);
        if (len > 0) {
            m_voiceId.resize(len);
            WideCharToMultiByte(CP_ACP, 0, dstrVoiceId, -1, &m_voiceId[0], len, NULL, NULL);
            // Remove null terminator from string
            if (!m_voiceId.empty() && m_voiceId.back() == '\0') {
                m_voiceId.pop_back();
            }
        }
    }

    return S_OK;
}

//-----------------------------------------------------------------------------
// ISpObjectWithToken::GetObjectToken
//-----------------------------------------------------------------------------
STDMETHODIMP CVibeVoiceTTSEngine::GetObjectToken(ISpObjectToken** ppToken)
{
    if (!ppToken) {
        return E_POINTER;
    }

    if (!m_cpToken) {
        return E_UNEXPECTED;
    }

    *ppToken = m_cpToken;
    (*ppToken)->AddRef();
    return S_OK;
}

//-----------------------------------------------------------------------------
// ISpTTSEngine::GetOutputFormat
// Tells SAPI what audio format we produce
//-----------------------------------------------------------------------------
STDMETHODIMP CVibeVoiceTTSEngine::GetOutputFormat(
    const GUID* pTargetFmtId,
    const WAVEFORMATEX* pTargetWaveFormatEx,
    GUID* pOutputFormatId,
    WAVEFORMATEX** ppCoMemOutputWaveFormatEx)
{
    if (!pOutputFormatId || !ppCoMemOutputWaveFormatEx) {
        return E_POINTER;
    }

    // Allocate WAVEFORMATEX structure
    WAVEFORMATEX* pWfx = static_cast<WAVEFORMATEX*>(
        CoTaskMemAlloc(sizeof(WAVEFORMATEX)));
    if (!pWfx) {
        return E_OUTOFMEMORY;
    }

    // Fill in our format: 24kHz, 16-bit, mono PCM
    pWfx->wFormatTag = WAVE_FORMAT_PCM;
    pWfx->nChannels = NUM_CHANNELS;
    pWfx->nSamplesPerSec = SAMPLE_RATE;
    pWfx->wBitsPerSample = BITS_PER_SAMPLE;
    pWfx->nBlockAlign = (NUM_CHANNELS * BITS_PER_SAMPLE) / 8;
    pWfx->nAvgBytesPerSec = SAMPLE_RATE * pWfx->nBlockAlign;
    pWfx->cbSize = 0;

    *pOutputFormatId = SPDFID_WaveFormatEx;
    *ppCoMemOutputWaveFormatEx = pWfx;

    return S_OK;
}

//-----------------------------------------------------------------------------
// ISpTTSEngine::Speak
// Main synthesis method - receives text, outputs audio
//-----------------------------------------------------------------------------
STDMETHODIMP CVibeVoiceTTSEngine::Speak(
    DWORD dwSpeakFlags,
    REFGUID rguidFormatId,
    const WAVEFORMATEX* pWaveFormatEx,
    const SPVTEXTFRAG* pTextFragList,
    ISpTTSEngineSite* pOutputSite)
{
    if (!pTextFragList || !pOutputSite) {
        return E_INVALIDARG;
    }

    // Extract all text from the fragment list
    std::wstring fullText = ExtractText(pTextFragList);
    if (fullText.empty()) {
        return S_OK;  // Nothing to speak
    }

    // Set up the audio callback context
    AudioContext ctx;
    ctx.pOutputSite = pOutputSite;
    ctx.audioOffset = 0;
    ctx.cancelled = false;

    // Stream TTS from the Python server
    HRESULT hr = m_pipeClient.StreamTTS(
        fullText.c_str(),
        m_voiceId.c_str(),
        AudioCallback,
        &ctx,
        &ctx.cancelled
    );

    return hr;
}

//-----------------------------------------------------------------------------
// ExtractText - Combines all text fragments into a single string
//-----------------------------------------------------------------------------
std::wstring CVibeVoiceTTSEngine::ExtractText(const SPVTEXTFRAG* pTextFragList)
{
    std::wstring result;

    for (const SPVTEXTFRAG* pFrag = pTextFragList; pFrag != NULL; pFrag = pFrag->pNext) {
        // Only process spoken text (not silence, spell-out, etc.)
        if (pFrag->State.eAction == SPVA_Speak) {
            if (pFrag->pTextStart && pFrag->ulTextLen > 0) {
                result.append(pFrag->pTextStart, pFrag->ulTextLen);
            }
        }
        else if (pFrag->State.eAction == SPVA_Silence) {
            // Could insert a pause marker, but for now we just skip
        }
    }

    return result;
}

//-----------------------------------------------------------------------------
// AudioCallback - Called for each audio chunk from the server
//-----------------------------------------------------------------------------
void CVibeVoiceTTSEngine::AudioCallback(const BYTE* data, DWORD size, void* context)
{
    AudioContext* ctx = static_cast<AudioContext*>(context);
    if (!ctx || !ctx->pOutputSite) {
        return;
    }

    // Check if SAPI wants us to abort
    // GetActions() returns DWORD directly, not via pointer
    DWORD actions = ctx->pOutputSite->GetActions();
    if (actions & SPVES_ABORT) {
        ctx->cancelled = true;
        return;
    }

    // Write audio data to SAPI
    ULONG bytesWritten = 0;
    HRESULT hr = ctx->pOutputSite->Write(data, size, &bytesWritten);

    if (hr == SP_AUDIO_STOPPED) {
        ctx->cancelled = true;
        return;
    }

    ctx->audioOffset += bytesWritten;
}


//=============================================================================
// DLL Entry Points
//=============================================================================

extern "C" BOOL WINAPI DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved)
{
    return _AtlModule.DllMain(dwReason, lpReserved);
}

STDAPI DllCanUnloadNow()
{
    return _AtlModule.DllCanUnloadNow();
}

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv)
{
    return _AtlModule.DllGetClassObject(rclsid, riid, ppv);
}

STDAPI DllRegisterServer()
{
    // Register COM object
    HRESULT hr = _AtlModule.DllRegisterServer(FALSE);
    if (FAILED(hr)) {
        return hr;
    }

    // Note: Voice token registration is done separately by install script
    // because we need to register multiple voices with different attributes

    return S_OK;
}

STDAPI DllUnregisterServer()
{
    // Note: Voice token cleanup is done by uninstall script

    return _AtlModule.DllUnregisterServer(FALSE);
}

STDAPI DllInstall(BOOL bInstall, LPCWSTR pszCmdLine)
{
    HRESULT hr = E_FAIL;

    if (bInstall) {
        hr = DllRegisterServer();
        if (FAILED(hr)) {
            DllUnregisterServer();
        }
    }
    else {
        hr = DllUnregisterServer();
    }

    return hr;
}
