import os
import multiprocessing

import tkinter as tk
import tkinter.ttk as ttk
from tkinter import font, filedialog

from core_audio import CoreAudio, DeviceChangedCallback, VolumeChangedCallback
import core_audio_constants
from audio_player import AudioPlayer

from get_path import get_module_path

S_OK = 0


def icon_path() -> str:
    resource_path = get_module_path()
    path = os.path.join(resource_path, 'icons\\')
    return path


class MainWindow(tk.Frame):
    def __init__(self, root):
        super().__init__(root)
        self.root = root

        self._init_device_info()

        self.style = ttk.Style()
        themes = self.style.theme_names()
        # themes = ('winnative', 'clam', 'alt', 'default', 'classic', 'vista', 'xpnative')
        if 'winnative' in themes:
            self.style.theme_use('winnative')

        self._create_fonts()
        self._load_icons()
        self._create_frames()

        icon_file = icon_path() + 'headphone.ico'
        if os.path.exists(icon_file):
            self.root.iconbitmap(default=icon_file)

        self.root.protocol('WM_DELETE_WINDOW', self._exit)
        # idle timer ID
        self.after_id = None

    def _init_device_info(self):
        # Core Audio
        self.ca = CoreAudio()

        self.ca_audio_id_list = self.ca.audio_device_id_list() # Core Audio device ID List
        self.ca_selected_device_id = None                      # Selected Core Audio device ID

        self.device_notification = DeviceChangedCallback(render_callback=self.device_changed_callback)
        self.ca.register_device_change_callback(self.device_notification)
        self.volume_notification = VolumeChangedCallback(self.volume_changed_callback)

        # Audio Player
        self.audio_player = AudioPlayer()

    def _exit(self):
        # Stop playing, just in case
        self.audio_player.stop_audio()

        # UnRegister volume changed notifier
        if self.ca_selected_device_id:
            self.ca.unregister_volume_change_callback(self.ca_selected_device_id, self.volume_notification)
            self.volume_notification = None

        # UnRegister device changed notifier
        self.ca.unregister_device_change_callback(self.device_notification)
        self.root.quit()

    def _create_fonts(self):
        self.font12 = font.Font(root=self.root, family='', size=12)
        self.default_font = font.Font(root=self.root, family='', size=14)
        self.default_font_bold = font.Font(root=self.root, family='', size=14, weight='bold')
        self.root.option_add('*font', self.default_font)
        pass


    def _load_icons(self):
        self.icon_play = None
        self.icon_pause = None
        self.icon_stop = None
        self.icon_speaker = None
        self.icon_mute = None
        path = icon_path()
        if os.path.exists(path + 'play.png'):
            self.icon_play = tk.PhotoImage(file=path + 'play.png')
            self.icon_play = self.icon_play.subsample(2)
        if os.path.exists(path + 'pause.png'):
            self.icon_pause = tk.PhotoImage(file=path + 'pause.png')
            self.icon_pause = self.icon_pause.subsample(2)
        if os.path.exists(path + 'stop.png'):
            self.icon_stop = tk.PhotoImage(file=path + 'stop.png')
            self.icon_stop = self.icon_stop.subsample(2)
        if os.path.exists(path + 'speaker.png'):
            self.icon_speaker = tk.PhotoImage(file=path + 'speaker.png')
            self.icon_speaker = self.icon_speaker.subsample(2)
        if os.path.exists(path + 'mute.png'):
            self.icon_mute = tk.PhotoImage(file=path + 'mute.png')
            self.icon_mute = self.icon_mute.subsample(2)
        pass

    def _create_frames(self):
        # File selector
        self.frame_file_selector = tk.Frame(self.root)#, background='gray')
        self.frame_file_selector.place(x=10, y=10, width=780, height=80)
        # Speaker list
        self.frame_speaker_list = tk.Frame(self.root)#, background='blue')
        self.frame_speaker_list.place(x=10, y=100, width=450, height=110)
        # Speaker volume
        self.frame_speaker_volume = tk.Frame(self.root)#, background='green')
        self.frame_speaker_volume.place(x=480, y=100, width=310, height=60)
        # Play/Pause/Stop buttons
        self.frame_play_pause_stop = tk.Frame(self.root)#, background='maroon')
        self.frame_play_pause_stop.place(x=480, y=170, width=240, height=40)
        pass
        self._create_widgets()
        pass

    def _create_widgets(self):
        # File selector
        self._create_file_selector(self.frame_file_selector)
        # Speaker list
        self._create_speaker_list(self.frame_speaker_list)
        # Speaker volume
        self._create_speaker_volume(self.frame_speaker_volume)
        # Play/Pause/Stop buttons
        self._create_play_buttons(self.frame_play_pause_stop)
        pass

    def _create_file_selector(self, parent):
        # Label
        self.wav_label = tk.Label(parent, text='Select a wav file to play:', width=30, anchor=tk.W, font=self.default_font_bold)
        self.wav_label.place(x=0, y=0)
        # Entry
        self.wav_entry = tk.Entry(parent, font=self.font12)
        self.wav_entry.place(x=0, y=30, width=665, height=40)
        # Button
        self.wav_button = tk.Button(parent, text='Browse...', font=self.default_font_bold, command=self._on_browse)
        self.wav_button.place(x=675, y=30)
        pass

    def _create_speaker_list(self, parent):
        # Label
        self.speaker_label = tk.Label(parent, text='Select a speaker to sound:', width=30, anchor=tk.W, font=self.default_font_bold)
        self.speaker_label.place(x=0, y=0)
        # Frame (Listbox + Scrollbar)
        self.speaker_frame = tk.Frame(parent)
        self.speaker_frame.place(x=0, y=30, width=450, height=80)
        # Scrollbar
        self.scroll = ttk.Scrollbar(self.speaker_frame, orient=tk.VERTICAL)
        self.scroll.place(x=425, y=0, height=80, width=25)
        # Listbox
        self.speaker_list = tk.Listbox(self.speaker_frame, selectmode=tk.SINGLE, activestyle='none', yscrollcommand=self.scroll.set, font=self.font12, border=1)
        self.speaker_list.place(x=0, y=0, width=430, height=80)
        # Add speaker names to the list
        for id in self.ca_audio_id_list:
            friendly_name = self.ca.get_friendly_name(id)
            self.speaker_list.insert(tk.END, friendly_name)
            pass
        self.scroll.config(command=self.speaker_list.yview)
        self.speaker_list.bind('<<ListboxSelect>>', self._on_select_speaker)
        pass
        # Refresh button # _FOR_DEBUG_
        # self.refresh_button = tk.Button(parent, text='Refresh', font=('Arial',10), command=self._on_refresh_speaker_list)
        # self.refresh_button.place(x=380, y=0, width=65)
        pass

    def _create_speaker_volume(self, parent):
        # Button : Mute
        self.mute = tk.Button(parent, image=self.icon_speaker, command=self._on_mute)
        self.mute.place(x=0, y=10, width=40, height=40)
        # Scale : Volume
        self.volume_var = tk.IntVar()
        self.volume_var.set(0)
        self.volume_scale = tk.Scale(parent, from_=0, to=100, variable=self.volume_var, resolution=1, showvalue=True, length=260, width=20, orient=tk.HORIZONTAL, command=self._on_volume)
        self.volume_scale.place(x=50, y=0)
        # Disable controls
        self.mute.config(state=tk.DISABLED)
        self.volume_scale.config(state=tk.DISABLED)
        pass

    def _create_play_buttons(self, parent):
        # Button : Play
        self.play_button = tk.Button(parent, image=self.icon_play, command=self._on_play)
        self.play_button.place(x=20, y=0, width=40, height=40)
        # Button : Pause
        self.pause_button = tk.Button(parent, image=self.icon_pause, command=self._on_pause)
        self.pause_button.place(x=100, y=0, width=40, height=40)
        # Button : Stop
        self.stop_button = tk.Button(parent, image=self.icon_stop, command=self._on_stop)
        self.stop_button.place(x=180, y=0, width=40, height=40)
        # Disable controls
        self.play_button.config(state=tk.DISABLED)
        self.pause_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.DISABLED)
        pass

    def _on_browse(self):
        file_path = filedialog.askopenfilename(filetypes=[('Wav Files', '*.wav')])
        if file_path:
            self.wav_entry.delete(0, tk.END)
            self.wav_entry.insert(0, file_path)
        pass

    def _on_select_speaker(self, event):
        selected = self.speaker_list.curselection()
        if selected:
            n = selected[0]
        else:
            return

        # Release
        if self.volume_notification:
            # UnRegister volume changed notifier
            self.ca.unregister_volume_change_callback(self.ca_selected_device_id, self.volume_notification)
            self.volume_notification = None
        # Release audio_endpoint_volume
        self.ca.release()

        pass

        if n >= 0:
            # Selected device ID
            self.ca_selected_device_id = self.ca_audio_id_list[n]

            # Enable controls
            self.mute.config(state=tk.NORMAL)
            self.volume_scale.config(state=tk.NORMAL)

            # Set volume slider
            vol = self.ca.get_volume(self.ca_selected_device_id)
            self.volume_var.set(int((vol+0.005) * 100)) # needs round off

            # Set mute status
            mute = self.ca.get_mute(self.ca_selected_device_id)
            self.mute.config(image=self.icon_mute if mute else self.icon_speaker)

            # Play/Pause/Stop
            self.play_button.config(state=tk.NORMAL)
            self.pause_button.config(state=tk.DISABLED)
            self.stop_button.config(state=tk.DISABLED)

            # # Register volume changed notifier
            self.volume_notification = VolumeChangedCallback(self.volume_changed_callback)
            self.ca.register_volume_change_callback(self.ca_selected_device_id, self.volume_notification)

        pass

    def _on_mute(self):
        if self.ca_selected_device_id:
            mute = self.ca.get_mute(self.ca_selected_device_id)
            mute = not mute
            self.ca.set_mute(self.ca_selected_device_id, mute)
            self.mute.config(image=self.icon_mute if mute else self.icon_speaker)

    def _on_volume(self, event):
        if self.ca_selected_device_id:
            volume = self.volume_var.get() / 100
            self.ca.set_volume(self.ca_selected_device_id, volume)
            if volume == 0:
                mute = True
                self.ca.set_mute(self.ca_selected_device_id, mute)
                self.mute.config(image=self.icon_mute)

    def _on_refresh_speaker_list(self):
        if self.ca_selected_device_id:
            # UnRegister volume changed notifier
            self.ca.unregister_volume_change_callback(self.ca_selected_device_id, self.volume_notification)
            self.volume_notification = None
            # Release audio_endpoint_volume
            # _CAUTION_ : If it is called from callback function, the following line causes deadlock
            self.ca.release()
        self.ca_selected_device_id = None

        self.ca_audio_id_list = self.ca.audio_device_id_list()

        self.speaker_list.config(state=tk.NORMAL)
        self.speaker_list.delete(0, tk.END)
        for id in self.ca_audio_id_list:
            friendly_name = self.ca.get_friendly_name(id)
            self.speaker_list.insert(tk.END, friendly_name)
        self.speaker_list.selection_clear(0, tk.END)

        self.mute.config(state=tk.DISABLED)
        self.volume_scale.config(state=tk.DISABLED)
        self.play_button.config(state=tk.DISABLED)
        self.pause_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.DISABLED)

    def _on_play(self):
        if not self.ca_selected_device_id:
            return

        wav_file = self.wav_entry.get()
        if not wav_file:
            return
        if not os.path.exists(wav_file):
            return

        self.speaker_list.config(state=tk.DISABLED)
        self.play_button.config(state=tk.DISABLED)
        self.pause_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.NORMAL)

        # Play audio
        self.audio_player.play_audio(self.ca_selected_device_id, wav_file)

        # Start timer
        self.after_id = self.after(100, self._wait_finish)

    def _on_pause(self):
        self.audio_player.pause_audio()
        self.play_button.config(state=tk.NORMAL)
        self.pause_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)

    def _on_stop(self):
        self.audio_player.stop_audio()
        self.speaker_list.config(state=tk.NORMAL)
        self.play_button.config(state=tk.NORMAL)
        self.pause_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.DISABLED)
        # Stop timer
        if self.after_id:
            self.after_cancel(self.after_id)
            self.after_id = None

    def _wait_finish(self):
        if self.audio_player.is_playing:
            # Playing, ReStart timer
            self.after_id = self.after(100, self._wait_finish)
        else:
            self.audio_player.audio_finished()
            # Finish playing
            self.after_id = None
            self.speaker_list.config(state=tk.NORMAL)
            if self.ca_selected_device_id:
                self.play_button.config(state=tk.NORMAL)
            else:
                self.play_button.config(state=tk.DISABLED)
            self.pause_button.config(state=tk.DISABLED)
            self.stop_button.config(state=tk.DISABLED)
            # Stop timer
            if self.after_id:
                self.after_cancel(self.after_id)
                self.after_id = None

    def volume_changed_callback(self, guid, bMuted, fMasterVolume, nChannels, ChannelVolumes):
        """
        Callback function, called when the volume is changed.
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

        self.mute.config(image=self.icon_mute if bMuted else self.icon_speaker)
        self.volume_var.set(int((fMasterVolume+0.005) * 100)) # needs round off
        return S_OK

    def device_changed_callback(self, device_id, new_state_id):
        """
        Callback function, called when the device is changed.
        """

        # Refer:
        #   https://learn.microsoft.com/ja-jp/windows/win32/coreaudio/device-state-xxx-constants
        state = {
            core_audio_constants.DeviceState.ACTIVE:     'ACTIVE',
            core_audio_constants.DeviceState.DISABLED:   'DISABLED',
            core_audio_constants.DeviceState.NOTPRESENT: 'NOTPRESENT',
            core_audio_constants.DeviceState.UNPLUGGED:  'UNPLUGGED',
        }
        # print(f'render device changed : {device_id=}, new_state={state[new_state_id]}')

        # STOP Audio ...
        self.audio_player.stop_audio()

        # _CAUTION_ : The following line causes deadlock, if it calls here
        # It's important to call it from idle timer.
        # The _on_refresh_speaker_list() will try to release the Core Audio resources.
        # This function is called from the Core Audio.
        # So, if it calls directly here, it causes deadlock.
        # self._on_refresh_speaker_list()
        self.after(100, self._on_refresh_speaker_list)


    # It doesn't need, because it handles by idle timer
    # def play_finished_callback(self):
    #     """
    #     Callback function, called when the play is finished.
    #     """
    #     pass


def main():
    root = tk.Tk()
    root.title('Simple WASAPI Player')
    # root.geometry('800x220+50+50')
    root.geometry('800x220')
    root.resizable(False, False)

    # Create window
    main_window = MainWindow(root)
    root.mainloop()


if __name__ == '__main__':
    multiprocessing.freeze_support()
    main()


