import comtypes
from pycaw.api.mmdeviceapi import IMMDeviceEnumerator, PROPERTYKEY
from pycaw.api.endpointvolume import IAudioEndpointVolume
import core_audio_constants


class CoreAudio:
    """
    Core Audio API wrap class
    """

    def __init__(self):
        pass

    def __del__(self):
        pass

    def get_default_audio_device_id(self, role: int) -> str:
        """
        Return the default audio device ID with the following process.

        1. CoInitialize()
        2. IMMDeviceEnumerator = CoCreateInstance(...)
        3. IMMDevice = IMMDeviceEnumerator::GetDefaultAudioEndpoint(...)
        4. id = IMMDevice::GetId()
        5. CoUninitialize()
        """

        comtypes.CoInitialize()

        device_enumerator = comtypes.CoCreateInstance(
            core_audio_constants.CLSID_MMDeviceEnumerator,
            IMMDeviceEnumerator,
            comtypes.CLSCTX_INPROC_SERVER,
        )

        device = device_enumerator.GetDefaultAudioEndpoint(
            core_audio_constants.EDataFlow.eRender,
            role,
            # core_audio_constants.ERole.eConsole == 0 or
            # core_audio_constants.ERole.eMultimedia == 1 or
            # core_audio_constants.ERole.eCommunications == 2
            # Actually, on the current windows version, the same device is returned regardless of the role.
        )

        id = device.GetId()

        comtypes.CoUninitialize()

        return id

    def audio_device_id_list(self) -> list:
        """
        Enumerate Core Audio devices and return a list of GUIDs with the following process.

        1. CoInitialize()
        2. IMMDeviceEnumerator = CoCreateInstance(...)
        3. IMMDeviceCollection = IMMDeviceEnumerator::EnumAudioEndpoints(...)
        4. IMMDevice = IMMDeviceCollection::Item(i)
        5. id = IMMDevice::GetId()
        6. CoUninitialize()
        """

        comtypes.CoInitialize()

        device_enumerator = comtypes.CoCreateInstance(
            core_audio_constants.CLSID_MMDeviceEnumerator,
            IMMDeviceEnumerator,
            comtypes.CLSCTX_INPROC_SERVER,
        )

        collections = device_enumerator.EnumAudioEndpoints( # type: ignore
            core_audio_constants.EDataFlow.eRender,
            core_audio_constants.DeviceState.ACTIVE,
            # const.DeviceState.ACTIVE | const.DeviceState.UNPLUGGED,
        )

        devices = []

        count = collections.GetCount()
        for i in range(count):
            device = collections.Item(i)

            # Refer:
            #   https://github.com/AndreMiras/pycaw/blob/develop/pycaw/utils.py
            # 
            # property_store = device.OpenPropertyStore(const.STGM.STGM_READ)
            # property_count = property_store.GetCount()
            # for j in range(property_count):
            #     key = property_store.GetAt(j)
            #     val = property_store.GetValue(key)
            #     value = val.GetValue()
            #     print(f'{key=}, {val=}, {value=}')

            id = device.GetId()
            devices.append(id)

        comtypes.CoUninitialize()

        return devices

    def get_friendly_name(self, device_id) -> str:
        """
        Return the friendly name of the device from the device ID with the following process.

        1. CoInitialize()
        2. IMMDeviceEnumerator = CoCreateInstance(...)
        3. IMMDevice = IMMDeviceEnumerator::GetDevice(ID)
        4. IPropertyStore = IMMDevice::OpenPropertyStore(STGM_READ)
        5. PROPERTYKEY = {A45C254E-DF1C-4EFD-8020-67D146A850E0}, 14
        6. value = IPropertyStore::GetValue(PROPERTYKEY)
        7. friendly_name = value.GetValue()
        8. CoUninitialize()
        """

        comtypes.CoInitialize()

        device_enumerator = comtypes.CoCreateInstance(
            core_audio_constants.CLSID_MMDeviceEnumerator,
            IMMDeviceEnumerator,
            comtypes.CLSCTX_INPROC_SERVER,
        )

        device = device_enumerator.GetDevice(device_id) # type: ignore
        property_store = device.OpenPropertyStore(core_audio_constants.STGM.STGM_READ)

        # Refer:
        #   https://github.com/AndreMiras/pycaw/blob/develop/pycaw/utils.py
    
        key = PROPERTYKEY()
        key.fmtid = comtypes.GUID('{A45C254E-DF1C-4EFD-8020-67D146A850E0}')
        key.pid = 14

        value = property_store.GetValue(comtypes.pointer(key))
        friendly_name = value.GetValue()

        comtypes.CoUninitialize()

        return friendly_name

    def get_volume(self, device_id):
        """
        Return the master volume of the specified device.

        1. CoInitialize()
        2. IMMDeviceEnumerator = CoCreateInstance(...)
        3. IMMDevice = IMMDeviceEnumerator::GetDevice(ID)
        4. IUnknown = IMMDevice::Activate(...)
        5. IAudioEndpointVolume = IUnknown::QueryInterface(IAudioEndpointVolume)
        6. volume = IAudioEndpointVolume::GetMasterVolumeLevelScalar()
        7. IAudioEndpointVolume::Release()
        8. CoUninitialize()
        """

        comtypes.CoInitialize()

        device_enumerator = comtypes.CoCreateInstance(
            core_audio_constants.CLSID_MMDeviceEnumerator,
            IMMDeviceEnumerator,
            comtypes.CLSCTX_INPROC_SERVER,
        )

        device = device_enumerator.GetDevice(device_id) # type: ignore
        audio_endpoint_volume = device.Activate(
            IAudioEndpointVolume._iid_, # type: ignore
            comtypes.CLSCTX_ALL,
            None,
        )

        endpoint_volume = audio_endpoint_volume.QueryInterface(IAudioEndpointVolume)

        volume = endpoint_volume.GetMasterVolumeLevelScalar()

        audio_endpoint_volume.Release()

        comtypes.CoUninitialize()

        return volume

    def get_mute(self, device_id):
        """
        Return the mute state of the specified device.

        1. CoInitialize()
        2. IMMDeviceEnumerator = CoCreateInstance(...)
        3. IMMDevice = IMMDeviceEnumerator::GetDevice(ID)
        4. IUnknown = IMMDevice::Activate(...)
        5. IAudioEndpointVolume = IUnknown::QueryInterface(IAudioEndpointVolume)
        6. mute = IAudioEndpointVolume::GetMute()
        7. IAudioEndpointVolume::Release()
        8. CoUninitialize()
        """

        comtypes.CoInitialize()

        device_enumerator = comtypes.CoCreateInstance(
            core_audio_constants.CLSID_MMDeviceEnumerator,
            IMMDeviceEnumerator,
            comtypes.CLSCTX_INPROC_SERVER,
        )

        device = device_enumerator.GetDevice(device_id) # type: ignore
        audio_endpoint_volume = device.Activate(
            IAudioEndpointVolume._iid_, # type: ignore
            comtypes.CLSCTX_ALL,
            None,
        )

        endpoint_volume = audio_endpoint_volume.QueryInterface(IAudioEndpointVolume)

        mute = endpoint_volume.GetMute()

        audio_endpoint_volume.Release()

        comtypes.CoUninitialize()

        return True if mute==1 else False

    def set_volume(self, device_id, volume: float):
        """
        Set the master volume of the specified device.

        1. CoInitialize()
        2. IMMDeviceEnumerator = CoCreateInstance(...)
        3. IMMDevice = IMMDeviceEnumerator::GetDevice(ID)
        4. IUnknown = IMMDevice::Activate(...)
        5. IAudioEndpointVolume = IUnknown::QueryInterface(IAudioEndpointVolume)
        6. volume = IAudioEndpointVolume::SetMasterVolumeLevelScalar(volume)
        7. IAudioEndpointVolume::Release()
        8. CoUninitialize()
        """

        comtypes.CoInitialize()

        device_enumerator = comtypes.CoCreateInstance(
            core_audio_constants.CLSID_MMDeviceEnumerator,
            IMMDeviceEnumerator,
            comtypes.CLSCTX_INPROC_SERVER,
        )

        device = device_enumerator.GetDevice(device_id) # type: ignore
        audio_endpoint_volume = device.Activate(
            IAudioEndpointVolume._iid_, # type: ignore
            comtypes.CLSCTX_ALL,
            None,
        )

        endpoint_volume = audio_endpoint_volume.QueryInterface(IAudioEndpointVolume)

        guid = device_id.split('}.')[1]
        endpoint_volume.SetMasterVolumeLevelScalar(volume, None)

        audio_endpoint_volume.Release()

        comtypes.CoUninitialize()

    def set_mute(self, device_id, mute: bool):
        """
        Set the mute state of the specified device.

        1. CoInitialize()
        2. IMMDeviceEnumerator = CoCreateInstance(...)
        3. IMMDevice = IMMDeviceEnumerator::GetDevice(ID)
        4. IUnknown = IMMDevice::Activate(...)
        5. IAudioEndpointVolume = IUnknown::QueryInterface(IAudioEndpointVolume)
        6. mute = IAudioEndpointVolume::SetMute(mute)
        7. IAudioEndpointVolume::Release()
        8. CoUninitialize()
        """

        comtypes.CoInitialize()

        device_enumerator = comtypes.CoCreateInstance(
            core_audio_constants.CLSID_MMDeviceEnumerator,
            IMMDeviceEnumerator,
            comtypes.CLSCTX_INPROC_SERVER,
        )

        device = device_enumerator.GetDevice(device_id) # type: ignore
        audio_endpoint_volume = device.Activate(
            IAudioEndpointVolume._iid_, # type: ignore
            comtypes.CLSCTX_ALL,
            None,
        )

        endpoint_volume = audio_endpoint_volume.QueryInterface(IAudioEndpointVolume)

        guid = device_id.split('}.')[1]
        endpoint_volume.SetMute(mute, None)

        audio_endpoint_volume.Release()

        comtypes.CoUninitialize()

