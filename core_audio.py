import uuid
import comtypes
from comtypes import GUID
from comtypes import COMObject
from pycaw.api.mmdeviceapi import IMMDeviceEnumerator, IMMNotificationClient, PROPERTYKEY
from pycaw.api.endpointvolume import IAudioEndpointVolume, IAudioEndpointVolumeCallback
import core_audio_constants


S_OK = 0
MY_UUID = '{'+str(uuid.uuid4())+'}' # {xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx}


class COMErrorException(Exception):
    # Detail information are not implemented. Just for raising exception.
    pass


class DeviceChangedCallback(COMObject):
    """
    IMMNotificationClient interface class
    """

    _com_interfaces_ = (IMMNotificationClient,)

    def __init__(self, render_callback=None, capture_callback=None):
        super().__init__()
        self.render_callback  = render_callback  # callback function for render device
        self.capture_callback = capture_callback # callback function for capture device

    def OnDefaultDeviceChanged(self, flow_id, role_id, default_device_id):
        pass
        return S_OK
    def OnDeviceAdded(self, added_device_id):
        pass
        return S_OK
    def OnDeviceRemoved(self, removed_device_id):
        pass
        return S_OK
    def OnDeviceStateChanged(self, device_id, new_state_id):
        state = {
            core_audio_constants.DeviceState.ACTIVE:     'ACTIVE',
            core_audio_constants.DeviceState.DISABLED:   'DISABLED',
            core_audio_constants.DeviceState.NOTPRESENT: 'NOTPRESENT',
            core_audio_constants.DeviceState.UNPLUGGED:  'UNPLUGGED',
        }
        # print(f'OnDeviceStateChanged : {device_id=}, new_state={state[new_state_id]}') # _FOR_DEBUG_

        render  = '{0.0.0.00000000}'
        capture = '{0.0.1.00000000}'

        if device_id.startswith(render):
            # render device
            if self.render_callback:
                self.render_callback(device_id, new_state_id)
        elif device_id.startswith(capture):
            # capture device
            if self.capture_callback:
                self.capture_callback(device_id, new_state_id)
        pass
        return S_OK
    def OnPropertyValueChanged(self, device_id, property_struct):
        pass
        return S_OK


class VolumeChangedCallback(COMObject):
    """
    IAudioEndpointVolumeCallback interface class
    """

    # Refer:
    #   https://learn.microsoft.com/ja-jp/windows/win32/api/endpointvolume/ns-endpointvolume-audio_volume_notification_data
    # 
    # typedef struct AUDIO_VOLUME_NOTIFICATION_DATA {
    #   GUID  guidEventContext;
    #   BOOL  bMuted;
    #   float fMasterVolume;
    #   UINT  nChannels;
    #   float afChannelVolumes[1];
    # } AUDIO_VOLUME_NOTIFICATION_DATA, *PAUDIO_VOLUME_NOTIFICATION_DATA;

    def __init__(self, callback=None):
        self.callback = callback

    _com_interfaces_ = (IAudioEndpointVolumeCallback,)

    def OnNotify(self, pNotify):
        # breakpoint()
        notify_data = pNotify.contents
        guid = notify_data.guidEventContext
        bMuted = notify_data.bMuted
        fMasterVolume = notify_data.fMasterVolume
        nChannels = notify_data.nChannels
        ChannelVolumes = list(notify_data.afChannelVolumes)
        pass
        if self.callback:
            # ChannelVolumes is an array, and only nChannels elements are validated.
            self.callback(guid, bMuted, fMasterVolume, nChannels, ChannelVolumes[:nChannels])
        return S_OK


class CoreAudio:
    """
    Core Audio API wrap class
    """

    def __init__(self):
        # If the IAudioEndpointVolume interface is released, the callback function for volume change notification will also be released.
        # To avoid releasing the callback function, it is as an instance variable.
        self.audio_endpoint_volume = None
        # It sets to True, when the volume change callback function is registered.
        self.volume_change_callback_registered = False

    def __del__(self):
        if self.volume_change_callback_registered:
            # Error : Not released
            pass
        if self.audio_endpoint_volume:
            self.audio_endpoint_volume.Release()

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

    def register_device_change_callback(self, callback):
        """
        Register a callback function to receive device state change notifications.

        1. CoInitialize()
        2. IMMDeviceEnumerator = CoCreateInstance(...)
        3. IMMDeviceEnumerator::RegisterEndpointNotificationCallback(callback)
        4. CoUninitialize()
        """

        comtypes.CoInitialize()

        device_enumerator = comtypes.CoCreateInstance(
            core_audio_constants.CLSID_MMDeviceEnumerator,
            IMMDeviceEnumerator,
            comtypes.CLSCTX_INPROC_SERVER,
        )

        ret = device_enumerator.RegisterEndpointNotificationCallback(callback) # type: ignore
        pass

        comtypes.CoUninitialize()

    def unregister_device_change_callback(self, callback):
        """
        Unregister a callback function to receive device state change notifications.

        1. CoInitialize()
        2. IMMDeviceEnumerator = CoCreateInstance(...)
        3. IMMDeviceEnumerator::UnregisterEndpointNotificationCallback(callback)
        4. CoUninitialize()
        """

        comtypes.CoInitialize()

        device_enumerator = comtypes.CoCreateInstance(
            core_audio_constants.CLSID_MMDeviceEnumerator,
            IMMDeviceEnumerator,
            comtypes.CLSCTX_INPROC_SERVER,
        )

        ret = device_enumerator.UnregisterEndpointNotificationCallback(callback) # type: ignore
        pass

        comtypes.CoUninitialize()

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

        if self.audio_endpoint_volume is None:
            device_enumerator = comtypes.CoCreateInstance(
                core_audio_constants.CLSID_MMDeviceEnumerator,
                IMMDeviceEnumerator,
                comtypes.CLSCTX_INPROC_SERVER,
            )

            device = device_enumerator.GetDevice(device_id) # type: ignore
            self.audio_endpoint_volume = device.Activate(
                IAudioEndpointVolume._iid_, # type: ignore
                comtypes.CLSCTX_ALL,
                None,
            )

        endpoint_volume = self.audio_endpoint_volume.QueryInterface(IAudioEndpointVolume)

        volume = endpoint_volume.GetMasterVolumeLevelScalar()

        # Don't release the IAudioEndpointVolume interface here.
        # self.audio_endpoint_volume.Release()

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

        if self.audio_endpoint_volume is None:
            device_enumerator = comtypes.CoCreateInstance(
                core_audio_constants.CLSID_MMDeviceEnumerator,
                IMMDeviceEnumerator,
                comtypes.CLSCTX_INPROC_SERVER,
            )

            device = device_enumerator.GetDevice(device_id) # type: ignore
            self.audio_endpoint_volume = device.Activate(
                IAudioEndpointVolume._iid_, # type: ignore
                comtypes.CLSCTX_ALL,
                None,
            )

        endpoint_volume = self.audio_endpoint_volume.QueryInterface(IAudioEndpointVolume)

        mute = endpoint_volume.GetMute()

        # Don't release the IAudioEndpointVolume interface here.
        # self.audio_endpoint_volume.Release()

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

        if self.audio_endpoint_volume is None:
            device_enumerator = comtypes.CoCreateInstance(
                core_audio_constants.CLSID_MMDeviceEnumerator,
                IMMDeviceEnumerator,
                comtypes.CLSCTX_INPROC_SERVER,
            )

            device = device_enumerator.GetDevice(device_id) # type: ignore
            self.audio_endpoint_volume = device.Activate(
                IAudioEndpointVolume._iid_, # type: ignore
                comtypes.CLSCTX_ALL,
                None,
            )

        endpoint_volume = self.audio_endpoint_volume.QueryInterface(IAudioEndpointVolume)

        guid = device_id.split('}.')[1]
        endpoint_volume.SetMasterVolumeLevelScalar(volume, GUID(MY_UUID)) # GUID(guid))
        # It is possible to set GUID_NULL as a GUID to pass to the IAudioEndpointVolumeCallback::OnNotify.
        # It's a trial to use a unique GUID to distinguish internal changes to others.

        # Don't release the IAudioEndpointVolume interface here.
        # self.audio_endpoint_volume.Release()

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

        if self.audio_endpoint_volume is None:
            device_enumerator = comtypes.CoCreateInstance(
                core_audio_constants.CLSID_MMDeviceEnumerator,
                IMMDeviceEnumerator,
                comtypes.CLSCTX_INPROC_SERVER,
            )

            device = device_enumerator.GetDevice(device_id) # type: ignore
            self.audio_endpoint_volume = device.Activate(
                IAudioEndpointVolume._iid_, # type: ignore
                comtypes.CLSCTX_ALL,
                None,
            )

        endpoint_volume = self.audio_endpoint_volume.QueryInterface(IAudioEndpointVolume)

        guid = device_id.split('}.')[1]
        endpoint_volume.SetMute(mute, GUID(MY_UUID)) # GUID(guid))
        # It is possible to set GUID_NULL as a GUID to pass to the IAudioEndpointVolumeCallback::OnNotify.
        # It's a trial to use a unique GUID to distinguish internal changes to others.

        # self.audio_endpoint_volume.Release()

        comtypes.CoUninitialize()

    def register_volume_change_callback(self, device_id, callback):
        """
        Register a callback function to receive volume change notifications for the specified device.

        1. CoInitialize()
        2. IMMDeviceEnumerator = CoCreateInstance(...)
        3. IMMDevice = IMMDeviceEnumerator::GetDevice(ID)
        4. IUnknown = IMMDevice::Activate(...)
        5. IAudioEndpointVolume = IUnknown::QueryInterface(IAudioEndpointVolume)
        6. IAudioEndpointVolume::RegisterControlChangeNotify(callback)
        7. IAudioEndpointVolume::Release()
        8. CoUninitialize()
        """

        if self.volume_change_callback_registered:
            # Error : Already registered
            return

        comtypes.CoInitialize()

        if self.audio_endpoint_volume is None:
            device_enumerator = comtypes.CoCreateInstance(
                core_audio_constants.CLSID_MMDeviceEnumerator,
                IMMDeviceEnumerator,
                comtypes.CLSCTX_INPROC_SERVER,
            )

            device = device_enumerator.GetDevice(device_id) # type: ignore
            self.audio_endpoint_volume = device.Activate(
                IAudioEndpointVolume._iid_, # type: ignore
                comtypes.CLSCTX_ALL,
                None,
            )

        endpoint_volume = self.audio_endpoint_volume.QueryInterface(IAudioEndpointVolume)

        ret = endpoint_volume.RegisterControlChangeNotify(callback)

        # Don't release the IAudioEndpointVolume interface here.
        # self.audio_endpoint_volume.Release()

        comtypes.CoUninitialize()

        self.volume_change_callback_registered = True

    def unregister_volume_change_callback(self, device_id, callback):
        """
        Unregister a callback function to receive volume change notifications for the specified device.

        1. CoInitialize()
        2. IMMDeviceEnumerator = CoCreateInstance(...)
        3. IMMDevice = IMMDeviceEnumerator::GetDevice(ID)
        4. IUnknown = IMMDevice::Activate(...)
        5. IAudioEndpointVolume = IUnknown::QueryInterface(IAudioEndpointVolume)
        6. IAudioEndpointVolume::UnregisterControlChangeNotify(callback)
        7. IAudioEndpointVolume::Release()
        8. CoUninitialize()
        """

        if not self.volume_change_callback_registered:
            # Error : Not registered
            return

        if self.audio_endpoint_volume is None:
            return

        comtypes.CoInitialize()

        # device_enumerator = comtypes.CoCreateInstance(
        #     const.CLSID_MMDeviceEnumerator,
        #     IMMDeviceEnumerator,
        #     comtypes.CLSCTX_INPROC_SERVER,
        # )
        # 
        # device = device_enumerator.GetDevice(device_id)
        # self.audio_endpoint_volume = device.Activate(
        #     IAudioEndpointVolume._iid_,
        #     comtypes.CLSCTX_ALL,
        #     None,
        # )

        endpoint_volume = self.audio_endpoint_volume.QueryInterface(IAudioEndpointVolume)

        ret = endpoint_volume.UnregisterControlChangeNotify(callback)
        pass

        self.audio_endpoint_volume.Release()
        # _CAUTION_ : If it is called from callback function, the following line causes deadlock
        self.audio_endpoint_volume = None

        comtypes.CoUninitialize()

        self.volume_change_callback_registered = False

    def release(self):
        if self.audio_endpoint_volume:
            self.audio_endpoint_volume.Release()
        self.audio_endpoint_volume = None


