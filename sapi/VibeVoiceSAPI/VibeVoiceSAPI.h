// VibeVoiceSAPI.h
// SAPI5 TTS Engine for VibeVoice
// Implements ISpTTSEngine and ISpObjectWithToken interfaces

#pragma once

// Suppress deprecated API warnings from sphelper.h
#pragma warning(disable: 4996)

#include <windows.h>
#include <sapi.h>
#include <sapiddk.h>
#include <sphelper.h>
#include <atlbase.h>
#include <atlcom.h>
#include <string>
#include <vector>

#include "resource.h"

// {A1B2C3D4-E5F6-7890-ABCD-EF1234567890}
// Generate a new GUID for your installation using guidgen.exe or uuidgen
extern const GUID CLSID_VibeVoiceTTSEngine_Impl;
#define CLSID_VibeVoiceTTSEngine CLSID_VibeVoiceTTSEngine_Impl

// Pipe communication constants
constexpr wchar_t PIPE_NAME[] = L"\\\\.\\pipe\\vibevoice";
constexpr DWORD PIPE_BUFFER_SIZE = 65536;
constexpr DWORD PIPE_TIMEOUT_MS = 30000;  // 30 second timeout

// Audio format: 24kHz, 16-bit, mono
constexpr DWORD SAMPLE_RATE = 24000;
constexpr WORD BITS_PER_SAMPLE = 16;
constexpr WORD NUM_CHANNELS = 1;

// Error codes from Python server
constexpr DWORD ERR_EMPTY_TEXT = 1;
constexpr DWORD ERR_INVALID_VOICE = 2;
constexpr DWORD ERR_MODEL_ERROR = 3;
constexpr DWORD ERR_UNKNOWN = 99;


// Forward declarations
class CVibeVoiceTTSEngine;


//-----------------------------------------------------------------------------
// PipeClient - Handles communication with the Python TTS server
//-----------------------------------------------------------------------------
class PipeClient
{
public:
    PipeClient();
    ~PipeClient();

    // Connect to the named pipe server
    HRESULT Connect();

    // Disconnect from the pipe
    void Disconnect();

    // Check if connected
    bool IsConnected() const { return m_hPipe != INVALID_HANDLE_VALUE; }

    // Send TTS request and stream audio back via callback
    // Returns S_OK on success, error HRESULT on failure
    typedef void (*AudioChunkCallback)(const BYTE* data, DWORD size, void* context);
    HRESULT StreamTTS(
        LPCWSTR text,
        LPCSTR voiceId,
        AudioChunkCallback callback,
        void* callbackContext,
        volatile bool* cancelFlag = nullptr
    );

private:
    HANDLE m_hPipe;

    // Read exactly 'count' bytes from the pipe
    HRESULT ReadExact(void* buffer, DWORD count);

    // Write exactly 'count' bytes to the pipe
    HRESULT WriteExact(const void* buffer, DWORD count);
};


//-----------------------------------------------------------------------------
// CVibeVoiceTTSEngine - Main SAPI TTS Engine implementation
//-----------------------------------------------------------------------------
class ATL_NO_VTABLE CVibeVoiceTTSEngine :
    public CComObjectRootEx<CComMultiThreadModel>,
    public CComCoClass<CVibeVoiceTTSEngine, &CLSID_VibeVoiceTTSEngine>,
    public ISpTTSEngine,
    public ISpObjectWithToken
{
public:
    CVibeVoiceTTSEngine();
    ~CVibeVoiceTTSEngine();

    DECLARE_REGISTRY_RESOURCEID(IDR_VIBEVOICE)
    DECLARE_NOT_AGGREGATABLE(CVibeVoiceTTSEngine)

    BEGIN_COM_MAP(CVibeVoiceTTSEngine)
        COM_INTERFACE_ENTRY(ISpTTSEngine)
        COM_INTERFACE_ENTRY(ISpObjectWithToken)
    END_COM_MAP()

    DECLARE_PROTECT_FINAL_CONSTRUCT()

    HRESULT FinalConstruct() { return S_OK; }
    void FinalRelease() {}

    // ISpObjectWithToken
    STDMETHODIMP SetObjectToken(ISpObjectToken* pToken) override;
    STDMETHODIMP GetObjectToken(ISpObjectToken** ppToken) override;

    // ISpTTSEngine
    STDMETHODIMP Speak(
        DWORD dwSpeakFlags,
        REFGUID rguidFormatId,
        const WAVEFORMATEX* pWaveFormatEx,
        const SPVTEXTFRAG* pTextFragList,
        ISpTTSEngineSite* pOutputSite
    ) override;

    STDMETHODIMP GetOutputFormat(
        const GUID* pTargetFmtId,
        const WAVEFORMATEX* pTargetWaveFormatEx,
        GUID* pOutputFormatId,
        WAVEFORMATEX** ppCoMemOutputWaveFormatEx
    ) override;

private:
    CComPtr<ISpObjectToken> m_cpToken;
    std::string m_voiceId;  // Voice ID from registry (e.g., "en-Carter_man")
    PipeClient m_pipeClient;

    // Helper to extract all text from SPVTEXTFRAG linked list
    std::wstring ExtractText(const SPVTEXTFRAG* pTextFragList);

    // Audio callback context
    struct AudioContext {
        ISpTTSEngineSite* pOutputSite;
        ULONGLONG audioOffset;
        volatile bool cancelled;
    };

    // Static callback for audio chunks
    static void AudioCallback(const BYTE* data, DWORD size, void* context);
};


// {A1B2C3D4-E5F6-7890-ABCD-EF1234567891}
// Type library GUID
extern const GUID LIBID_VibeVoiceSAPILib_Impl;
#define LIBID_VibeVoiceSAPILib LIBID_VibeVoiceSAPILib_Impl

//-----------------------------------------------------------------------------
// Module declaration for ATL
//-----------------------------------------------------------------------------
class CVibeVoiceSAPIModule : public ATL::CAtlDllModuleT<CVibeVoiceSAPIModule>
{
public:
    DECLARE_LIBID(LIBID_VibeVoiceSAPILib)
};

extern class CVibeVoiceSAPIModule _AtlModule;
