from comtypes import GUID

# Refer:
#   https://github.com/AndreMiras/pycaw/blob/develop/pycaw/constants.py
CLSID_MMDeviceEnumerator = GUID('{BCDE0395-E52F-467C-8E3D-C4579291692E}')

class EDataFlow:
    # Refer:
    #   https://learn.microsoft.com/ja-jp/windows/win32/api/mmdeviceapi/ne-mmdeviceapi-edataflow
    eRender = 0
    eCompute = 1
    eAll = 2


class ERole:
    # Refer:
    #   https://learn.microsoft.com/ja-jp/windows/win32/api/mmdeviceapi/ne-mmdeviceapi-erole
    eConsole = 0
    eMultimedia = 1
    eCommunications = 2


class DeviceState:
    # Refer:
    #   https://learn.microsoft.com/ja-jp/windows/win32/coreaudio/device-state-xxx-constants
    ACTIVE = 0x01
    DISABLED = 0x02
    NOTPRESENT = 0x04
    UNPLUGGED = 0x08


class STGM:
    # Refer:
    #   https://learn.microsoft.com/ja-jp/windows/win32/stg/stgm-constants
    STGM_READ = 0
    STGM_WRITE = 1
    STGM_READWRITE = 2

