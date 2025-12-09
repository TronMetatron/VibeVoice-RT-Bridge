# register_voices.ps1
# Registers VibeVoice voices in the Windows SAPI registry

#Requires -RunAsAdministrator

param(
    [string]$DllPath = "",
    [switch]$Uninstall
)

# CLSID for our TTS engine (must match VibeVoiceSAPI.h)
$EngineCLSID = "{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}"

# Voice definitions
$Voices = @(
    @{
        Name = "VibeVoice Carter"
        VoiceId = "en-Carter_man"
        Gender = "Male"
        Age = "Adult"
    },
    @{
        Name = "VibeVoice Davis"
        VoiceId = "en-Davis_man"
        Gender = "Male"
        Age = "Adult"
    },
    @{
        Name = "VibeVoice Emma"
        VoiceId = "en-Emma_woman"
        Gender = "Female"
        Age = "Adult"
    },
    @{
        Name = "VibeVoice Frank"
        VoiceId = "en-Frank_man"
        Gender = "Male"
        Age = "Adult"
    },
    @{
        Name = "VibeVoice Grace"
        VoiceId = "en-Grace_woman"
        Gender = "Female"
        Age = "Adult"
    },
    @{
        Name = "VibeVoice Mike"
        VoiceId = "en-Mike_man"
        Gender = "Male"
        Age = "Adult"
    },
    @{
        Name = "VibeVoice Samuel"
        VoiceId = "in-Samuel_man"
        Gender = "Male"
        Age = "Adult"
    }
)

# Registry base path for SAPI voices
$VoicesPath = "HKLM:\SOFTWARE\Microsoft\Speech\Voices\Tokens"

function Register-Voice {
    param(
        [hashtable]$Voice
    )

    $tokenName = "VibeVoice-$($Voice.VoiceId -replace '-','_' -replace '_.*','')"
    $tokenPath = Join-Path $VoicesPath $tokenName
    $attrPath = Join-Path $tokenPath "Attributes"

    Write-Host "Registering voice: $($Voice.Name) at $tokenPath"

    # Create token key
    if (-not (Test-Path $tokenPath)) {
        New-Item -Path $tokenPath -Force | Out-Null
    }

    # Set token values
    Set-ItemProperty -Path $tokenPath -Name "(Default)" -Value $Voice.Name
    Set-ItemProperty -Path $tokenPath -Name "CLSID" -Value $EngineCLSID
    Set-ItemProperty -Path $tokenPath -Name "VoiceId" -Value $Voice.VoiceId

    # Create attributes key
    if (-not (Test-Path $attrPath)) {
        New-Item -Path $attrPath -Force | Out-Null
    }

    # Set attributes
    Set-ItemProperty -Path $attrPath -Name "Name" -Value $Voice.Name
    Set-ItemProperty -Path $attrPath -Name "Gender" -Value $Voice.Gender
    Set-ItemProperty -Path $attrPath -Name "Age" -Value $Voice.Age
    Set-ItemProperty -Path $attrPath -Name "Vendor" -Value "Microsoft Research"
    Set-ItemProperty -Path $attrPath -Name "Language" -Value "409"  # English US

    Write-Host "  Registered successfully" -ForegroundColor Green
}

function Unregister-Voice {
    param(
        [hashtable]$Voice
    )

    $tokenName = "VibeVoice-$($Voice.VoiceId -replace '-','_' -replace '_.*','')"
    $tokenPath = Join-Path $VoicesPath $tokenName

    if (Test-Path $tokenPath) {
        Write-Host "Removing voice: $($Voice.Name)"
        Remove-Item -Path $tokenPath -Recurse -Force
        Write-Host "  Removed successfully" -ForegroundColor Yellow
    }
}

# Main logic
if ($Uninstall) {
    Write-Host "Unregistering VibeVoice voices..." -ForegroundColor Cyan
    foreach ($voice in $Voices) {
        Unregister-Voice -Voice $voice
    }
    Write-Host "`nVoices unregistered." -ForegroundColor Green
}
else {
    # Register DLL if path provided
    if ($DllPath -and (Test-Path $DllPath)) {
        Write-Host "Registering COM DLL: $DllPath" -ForegroundColor Cyan
        $result = Start-Process -FilePath "regsvr32.exe" -ArgumentList "/s `"$DllPath`"" -Wait -PassThru
        if ($result.ExitCode -eq 0) {
            Write-Host "  DLL registered successfully" -ForegroundColor Green
        }
        else {
            Write-Host "  DLL registration failed with exit code: $($result.ExitCode)" -ForegroundColor Red
            exit 1
        }
    }

    Write-Host "`nRegistering VibeVoice voices..." -ForegroundColor Cyan
    foreach ($voice in $Voices) {
        Register-Voice -Voice $voice
    }

    Write-Host "`nVoices registered successfully!" -ForegroundColor Green
    Write-Host "`nAvailable voices:"
    foreach ($voice in $Voices) {
        Write-Host "  - $($Voice.Name) ($($Voice.Gender))"
    }
}
