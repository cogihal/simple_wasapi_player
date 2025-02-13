# Simple wav file player using WASAPI

This is a simple wav file player script using WASAPI and Tkinter GUI.

## Features

- Select speaker to play.
- Select wav file to play.
- Play / Pause / Stop.
- Change volume.


## Developing environments

- Windows 11 Pro
- Python 3.12.7  
  If you use python 3.13, Tkinter will not work.  
	See details : https://github.com/python/cpython/issues/125235
- comtypes==1.4.9
- pycaw==20240210


## Caution

The following pycaw code should be edited.

pycaw/api/audioclient/depend.py

The type of the following variables should be changed from WORD to DWORD.
- nSamplesPerSec
- nAvgBytesPerSec

See also the following URL.

https://learn.microsoft.com/en-us/windows/win32/api/mmeapi/ns-mmeapi-waveformatex


## References

https://gist.github.com/kevinmoran/3d05e190fb4e7f27c1043a3ba321cede

