import multiprocessing
import time
import wave
from ctypes import POINTER, cast

import comtypes
from comtypes import CLSCTX_ALL, COMError
from pycaw.api.audioclient import WAVEFORMATEX
from pycaw.api.mmdeviceapi import IMMDevice, IMMDeviceEnumerator
from pycaw.pycaw import IAudioClient

import core_audio_constants
from audio_render_client import IAudioRenderClient


class Play:
    PLAY = 1
    PAUSE = 2
    STOP = 0


class Playing:
    PLAYING = 1
    FINISH = 0


def _get_device_by_id(id) -> IMMDevice:
    """
    Get IMMDevice by GUIDs with the following process.

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

    mmdevice = None

    count = collections.GetCount()
    for i in range(count):
        mmdevice = collections.Item(i)
        if id == mmdevice.GetId():
            break

    comtypes.CoUninitialize()

    return mmdevice


def _play_audio(device_id, wav_file, play, playing, event):
    """
    Play an audio file executed by a process.
    """

    # print('Start Process') # _FOR_DEBUG_
    wav = wave.open(wav_file, 'rb')
    frame_rate   = wav.getframerate() # frame rate (ex. 44100=44.1kHz)
    channels     = wav.getnchannels() # number of channels (monaural: 1, stereo: 2)
    sample_width = wav.getsampwidth() # sample byte size (ex. 16bit=2byte)
    frames       = wav.getnframes()   # number of frames
    data_size    = frames * channels * sample_width # data byte size

    # print('Init COM') # _FOR_DEBUG_
    comtypes.CoInitialize()

    audio_device = _get_device_by_id(device_id)

    interface = audio_device.Activate(
        IAudioClient._iid_, # type: ignore
        CLSCTX_ALL,
        None
    )
    audio_client = cast(interface, POINTER(IAudioClient))

    wav_format_ex = WAVEFORMATEX()
    wav_format_ex.wFormatTag      = 1
    wav_format_ex.nChannels       = channels
    wav_format_ex.nSamplesPerSec  = frame_rate
    wav_format_ex.wBitsPerSample  = sample_width * 8
    wav_format_ex.nBlockAlign     = wav_format_ex.nChannels * wav_format_ex.wBitsPerSample // 8
    wav_format_ex.nAvgBytesPerSec = wav_format_ex.nSamplesPerSec * wav_format_ex.nBlockAlign

    BUFFER_SIZE_IN_SECONDS = 2.0
    REFTIMES_PER_SEC = 10_000_000
    requestedSoundBufferDuration = int(REFTIMES_PER_SEC * BUFFER_SIZE_IN_SECONDS)
    stream_flags:int = core_audio_constants.AUDCLNT_STREAMFLAGS_RATEADJUST
    audio_client.Initialize(
        core_audio_constants.AUDCLNT_SHAREMODE.AUDCLNT_SHAREMODE_SHARED,
        stream_flags,
        requestedSoundBufferDuration,
        0,
        comtypes.pointer(wav_format_ex),
        None
    )

    service = audio_client.GetService(IAudioRenderClient._iid_) # type: ignore
    audio_render_client = cast(service, POINTER(IAudioRenderClient))

    num_buffer_frames = audio_client.GetBufferSize()

    audio_client.Start()

    frame_chunk_size = 1024 # This number is changeable.
    wav_play_data = 0
    frames_to_write = 0

    stream_available = True

    # print('Start loop') # _FOR_DEBUG_

    while wav_play_data < data_size:
        try:
            event.wait()

            if play.value == Play.PLAY:
                pass
            elif play.value == Play.PAUSE:
                # continue
                pass
            elif play.value == Play.STOP:
                break

            # The padding frame means the data in the buffer and has not been played yet.
            buffer_padding_frames = audio_client.GetCurrentPadding()

            # If the buffer padding is less than frame_chunk_size, write frame_chunk_size frames.
            if buffer_padding_frames < frame_chunk_size:
                frames_to_write = frame_chunk_size
            else:
                # Enough data is in the buffer. Not need to write more. Wait for the buffer to be played.
                continue

            # Get the WASAPI buffer to play.
            buffer_to_play = audio_render_client.GetBuffer(frames_to_write)
            # Read the wav data.
            wav_frame_data = wav.readframes(frames_to_write)

            byte_len = len(wav_frame_data) # The byte size of actual read wav data.
            for i in range(byte_len):
                # Copy the wav data to the WASAPI buffer to play.
                buffer_to_play[i] = wav_frame_data[i]
            audio_render_client.ReleaseBuffer(frames_to_write, 0)
            wav_play_data += byte_len

        except COMError as e:
            # IAudioRenderClient can't be used anymore
            # possibly, the device is disconnected before finish playing
            stream_available = False
            break

    # print('Loop break') # _FOR_DEBUG_

    try:
        while audio_client.GetCurrentPadding() > 0:
            time.sleep(0.1)
    except COMError as e:
        pass

    # print('Finish playing') # _FOR_DEBUG_

    playing.value = Playing.FINISH

    try:
        audio_client.Stop()
        audio_client.Release()
        audio_render_client.Release()
    except COMError as e:
        pass

    comtypes.CoUninitialize()

    # print('Exit Playing') # _FOR_DEBUG_


class AudioPlayer:
    def __init__(self):
        self.play = multiprocessing.Value('i', Play.STOP)
        self.playing = multiprocessing.Value('i', Playing.FINISH)
        self.play_process = None
        self.event = multiprocessing.Event()

    def play_audio(self, device_id, wav_file):
        """
        Play an audio file.
        
        Args:
            device_id (str): The audio device ID strings.
            wav_file (str): The path of the WAV file.
        """

        if self.play_process is None:
            # Playing process is not started.

            self.play.value = Play.PLAY
            self.playing.value = Playing.PLAYING
            self.event.set()
            self.play_process = multiprocessing.Process(target=_play_audio, args=(device_id, wav_file, self.play, self.playing, self.event))
            self.play_process.start()
        else:
            # PAUSE
            self.play.value = Play.PLAY
            self.event.set()

    def pause_audio(self):
        self.play.value = Play.PAUSE
        self.event.clear()

    def stop_audio(self):
        self.play.value = Play.STOP
        self.event.set()
        while self.is_playing:
            time.sleep(0.1)
        self.play_process = None

    def audio_finished(self):
        # If the audio is finished naturally, the process is finished but the instance variable is not cleared.
        # In this case, this method is needed to be called just to clear the variable.
        self.play_process = None
        self.event.set()

    @property
    def is_playing(self):
        return self.playing.value == Playing.PLAYING


