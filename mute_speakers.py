from core_audio import CoreAudio
import core_audio_constants

from mute_log import mute_log


def mute_current_speaker(mute: bool, vol: bool, log: bool):
    ca = CoreAudio()
    device0 = ca.get_default_audio_device_id(
        core_audio_constants.ERole.eConsole
    )
    device_name = ca.get_friendly_name(device0)

    if mute: ca.set_mute(device0, True)
    if vol: ca.set_volume(device0, 0)
    if log: mute_log(f'{device_name} - Mute({mute}) Volume({vol})')


def mute_all_speakers(mute: bool, vol: bool, log: bool):
    ca = CoreAudio()
    devices = ca.audio_device_id_list()
    for device in devices:
        device_name = ca.get_friendly_name(device)

        if mute: ca.set_mute(device, True)
        if vol: ca.set_volume(device, 0)
        if log: mute_log(f'{device_name} - Mute({mute}) Volume({vol})')


if __name__ == '__main__':
    ca = CoreAudio()

    # Get default render device
    device0 = ca.get_default_audio_device_id(
        core_audio_constants.ERole.eConsole
    )
    print(ca.get_friendly_name(device0))
    device1 = ca.get_default_audio_device_id(
        core_audio_constants.ERole.eMultimedia
    )
    print(ca.get_friendly_name(device1))
    device2 = ca.get_default_audio_device_id(
        core_audio_constants.ERole.eCommunications
    )
    print(ca.get_friendly_name(device2))

    volume = ca.get_volume(device0)
    print(volume)
    mute = ca.get_mute(device0)
    print(mute)

    ca.set_volume(device0, 0)
    ca.set_mute(device0, True)

    volume = ca.get_volume(device0)
    print(volume)
    mute = ca.get_mute(device0)
    print(mute)

    # Enumerate all render devices
    devices = ca.audio_device_id_list()
    for device in devices:
        print(ca.get_friendly_name(device))

