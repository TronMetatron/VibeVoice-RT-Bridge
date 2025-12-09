"""
Register VibeVoice voices in Windows 11 OneCore registry.
Run this script as Administrator.
"""
import ctypes
import sys
import winreg

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

if not is_admin():
    print("ERROR: This script must be run as Administrator!")
    print("Right-click and select 'Run as administrator'")
    input("Press Enter to exit...")
    sys.exit(1)

CLSID = '{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}'
VOICES = [
    {'name': 'VibeVoice Carter', 'id': 'en-Carter_man', 'gender': 'Male', 'token': 'VibeVoice-Carter'},
    {'name': 'VibeVoice Davis', 'id': 'en-Davis_man', 'gender': 'Male', 'token': 'VibeVoice-Davis'},
    {'name': 'VibeVoice Emma', 'id': 'en-Emma_woman', 'gender': 'Female', 'token': 'VibeVoice-Emma'},
    {'name': 'VibeVoice Frank', 'id': 'en-Frank_man', 'gender': 'Male', 'token': 'VibeVoice-Frank'},
    {'name': 'VibeVoice Grace', 'id': 'en-Grace_woman', 'gender': 'Female', 'token': 'VibeVoice-Grace'},
    {'name': 'VibeVoice Mike', 'id': 'en-Mike_man', 'gender': 'Male', 'token': 'VibeVoice-Mike'},
    {'name': 'VibeVoice Samuel', 'id': 'in-Samuel_man', 'gender': 'Male', 'token': 'VibeVoice-Samuel'},
]

print("Registering VibeVoice voices in Windows 11 OneCore...")
print()

base = r'SOFTWARE\Microsoft\Speech_OneCore\Voices\Tokens'
success_count = 0

for v in VOICES:
    try:
        key = winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, f"{base}\\{v['token']}")
        winreg.SetValueEx(key, '', 0, winreg.REG_SZ, v['name'])
        winreg.SetValueEx(key, 'CLSID', 0, winreg.REG_SZ, CLSID)
        winreg.SetValueEx(key, 'VoiceId', 0, winreg.REG_SZ, v['id'])

        attr = winreg.CreateKey(key, 'Attributes')
        winreg.SetValueEx(attr, 'Name', 0, winreg.REG_SZ, v['name'])
        winreg.SetValueEx(attr, 'Gender', 0, winreg.REG_SZ, v['gender'])
        winreg.SetValueEx(attr, 'Language', 0, winreg.REG_SZ, '409')
        winreg.SetValueEx(attr, 'Age', 0, winreg.REG_SZ, 'Adult')
        winreg.SetValueEx(attr, 'Vendor', 0, winreg.REG_SZ, 'VibeVoice')
        winreg.CloseKey(attr)
        winreg.CloseKey(key)
        print(f"  [OK] {v['name']}")
        success_count += 1
    except Exception as e:
        print(f"  [FAILED] {v['name']}: {e}")

print()
print(f"Registered {success_count}/{len(VOICES)} voices.")
print()
print("NOTE: You may need to restart Windows or the Speech service")
print("for the voices to appear in Windows 11 Settings.")
print()
input("Press Enter to exit...")
