Eela Audio Logos mixstation

I am making a public download and guide available to keep this great little machine alive in the modern age.
As a lot of Eela Audio got sold off, a lot has been lost regarding software and documentation.
Firmware used to be send out to you. There was never a public download page for firmware and software.

Credits for the working manual to https://interstage.dk/Sider/Produkter/Brochure_manual/Eela/Eela_Logos_man.pdf
And also the brochure: https://interstage.dk/Sider/Produkter/Brochure_manual/Eela/Eela_Logos_bro.pdf

Both zips include the correct LogosTool version for that firmware version!

If anyone has other firmware and LogosTool versions, please send me a message so that I can create an as complete page as possible.

Using Windows 95:
Oracle VM with 95/98 - able to download and upload config, but not successful to firmware upgrade to Main V1.41 from V1.28
Create an ISO from the files inside the ZIP and mount it as CD inside windows 95/98. Install software.
Make sure you hook the correct COM port!

I assume that it's way easier to have a machine actually running Wind95/98.

Using Windows 10 (or newer):
Software does open, but unable to connect to device.

To install configuration software LogosTool on Windows 11, you need this tool to open 16-bit old Windows software. LogosTool itself is a 32bit program. So you just need it for the installer:
https://github.com/otya128/winevdm



Will update if I find a way to firmware upgrade.

============================

Eela Controller Bridge tool

I vibecoded a tool to translate the transport controls,jog/shuttle wheel and d-pad (2,4,6,8 and ON # LINE button on the numpad) to MIDI and/or key shortcuts.
This way you are able to use your Eela Logos in any audio/video software.
I added presets of a lot of software.
It should be cross platform, but I only tested it fully on Windows 11.

Eela Controller Bridge
Cross-platform Python GUI for sniffing RS-232 button packets from an Eela Logos / D902
and mapping them to keyboard shortcuts or MIDI messages.

Tested conceptually for Windows, macOS, Linux.

Install:
    py -3.12 -m pip install pyserial PySide6 pyautogui mido python-rtmidi

If you only want keyboard mappings and sniffing, python-rtmidi is optional:
    py -3.12 -m pip install pyserial PySide6 pyautogui mido

Run:
    py -3.12 eela_controller_bridge.py

Notes:
- When you open the software on Windows, it doesn't open on your task bar. It's in the hidden icons next to the time in on the taskbar.
- Do not name this file "serial.py" or "import serial.py".
- On macOS, keyboard output may require Accessibility permission.
- On Windows, create a virtual MIDI port with loopMIDI if you want MIDI output.
