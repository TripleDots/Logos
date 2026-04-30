"""
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
- Do not name this file "serial.py" or "import serial.py".
- On macOS, keyboard output may require Accessibility permission.
- On Windows, create a virtual MIDI port with loopMIDI if you want MIDI output.
"""

from __future__ import annotations

import json
import os
import sys
import time

# Windows DPI fix: must be set before importing PySide6/Qt.
# Prevents: SetProcessDpiAwarenessContext() failed: Access is denied.
if sys.platform.startswith("win"):
    # Windows / Qt DPI handling:
    # Some hosts (VS Code, elevated shells, Windows compatibility settings) set DPI awareness
    # before Qt starts. Then Qt logs SetProcessDpiAwarenessContext "Access is denied".
    # It is usually harmless, but noisy, so we force Qt to skip DPI handling and suppress
    # the qpa.window warning category before PySide6 is imported.
    os.environ["QT_QPA_PLATFORM"] = "windows:dpiawareness=0"
    os.environ["QT_ENABLE_HIGHDPI_SCALING"] = "0"
    os.environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "0"
    os.environ["QT_SCALE_FACTOR"] = "1"
    os.environ["QT_SCREEN_SCALE_FACTORS"] = "1"
    os.environ["QT_SCALE_FACTOR_ROUNDING_POLICY"] = "PassThrough"
    os.environ["QT_LOGGING_RULES"] = os.environ.get("QT_LOGGING_RULES", "") + ";qt.qpa.window=false"
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Optional

import serial
import serial.tools.list_ports

from PySide6.QtCore import QObject, QThread, Signal, Slot, Qt, QEvent
from PySide6.QtGui import QIcon, QAction, QStandardItemModel, QStandardItem, QFont
from PySide6.QtWidgets import (
    QApplication,
    QStyle,
    QComboBox,
    QFileDialog,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QCheckBox,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QDialog,
    QPushButton,
    QPlainTextEdit,
    QSpinBox,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
    QSystemTrayIcon,
    QMenu,
)

import pyautogui
import webbrowser

try:
    import mido
except Exception:  # pragma: no cover
    mido = None


KEYBOARD_ACTION = "Keyboard Shortcut"
MIDI_ACTION = "MIDI"
APP_NAME = "EelaControllerBridge"

DAW_PRESETS = {
    # DAW presets
    "DAW / Acid Pro": {
        "START": ("space", "start"),
        "STOP": ("enter", "stop"),
        "RECORD": ("ctrl+r", "record"),
        "UNDO": ("ctrl+z", "cc:70:127"),
        "DELETE": ("delete", "cc:71:127"),
        "SHIFT": ("shift", "cc:19:127"),
        "ZOOM IN": ("up", "cc:20:1"),
        "ZOOM OUT": ("down", "cc:21:1"),
        "LEFT MARKER": ("pageup", "cc:22:127"),
        "RIGHT MARKER": ("pagedown", "cc:23:127"),
        "D-PAD UP": ("up", "cc:24:127"),
        "D-PAD DOWN": ("down", "cc:25:127"),
        "D-PAD LEFT": ("left", "cc:26:127"),
        "D-PAD RIGHT": ("right", "cc:27:127"),
        "JOG LEFT": ("left", "cc:28:65"),
        "JOG RIGHT": ("right", "cc:28:63"),
        "SHUTTLE LEFT": ("j", "cc:30:65"),
        "SHUTTLE RIGHT": ("l", "cc:30:63"),
        "ON # LINE": ("q", "note:63"),
    },
    "DAW / Ableton": {
        "START": ("space", "start"),
        "STOP": ("space", "stop"),
        "RECORD": ("f9", "record"),
        "UNDO": ("ctrl+z", "cc:70:127"),
        "DELETE": ("delete", "cc:71:127"),
        "SHIFT": ("shift", "cc:19:127"),
        "ZOOM IN": ("plus", "cc:20:1"),
        "ZOOM OUT": ("minus", "cc:21:1"),
        "LEFT MARKER": ("ctrl+shift+left", "cc:22:127"),
        "RIGHT MARKER": ("ctrl+shift+right", "cc:23:127"),
        "D-PAD UP": ("up", "cc:24:127"),
        "D-PAD DOWN": ("down", "cc:25:127"),
        "D-PAD LEFT": ("left", "cc:26:127"),
        "D-PAD RIGHT": ("right", "cc:27:127"),
        "JOG LEFT": ("left", "cc:28:65"),
        "JOG RIGHT": ("right", "cc:28:63"),
        "SHUTTLE LEFT": ("j", "cc:30:65"),
        "SHUTTLE RIGHT": ("l", "cc:30:63"),
        "ON # LINE": ("enter", "note:63"),
        "ZOOM TO SELECTION": ("z", "cc:72:127"),
    },
    "DAW / Adobe Audition": {
        "START": ("space", "start"),
        "STOP": ("space", "stop"),
        "RECORD": ("shift+space", "record"),
        "UNDO": ("ctrl+z", "cc:70:127"),
        "DELETE": ("delete", "cc:71:127"),
        "SHIFT": ("shift", "cc:19:127"),
        "ZOOM IN": ("plus", "cc:20:1"),
        "ZOOM OUT": ("minus", "cc:21:1"),
        "LEFT MARKER": ("ctrl+left", "cc:22:127"),
        "RIGHT MARKER": ("ctrl+right", "cc:23:127"),
        "D-PAD UP": ("up", "cc:24:127"),
        "D-PAD DOWN": ("down", "cc:25:127"),
        "D-PAD LEFT": ("left", "cc:26:127"),
        "D-PAD RIGHT": ("right", "cc:27:127"),
        "JOG LEFT": ("left", "cc:28:65"),
        "JOG RIGHT": ("right", "cc:28:63"),
        "SHUTTLE LEFT": ("j", "cc:30:65"),
        "SHUTTLE RIGHT": ("l", "cc:30:63"),
        "ON # LINE": ("m", "note:63"),
        "ZOOM TO SELECTION": ("z", "cc:72:127"),
    },
    "DAW / Audacity": {
        "START": ("space", "start"),
        "STOP": ("space", "stop"),
        "RECORD": ("r", "record"),
        "UNDO": ("ctrl+z", "cc:70:127"),
        "DELETE": ("delete", "cc:71:127"),
        "SHIFT": ("shift", "cc:19:127"),
        "ZOOM IN": ("ctrl+1", "cc:20:1"),
        "ZOOM OUT": ("ctrl+3", "cc:21:1"),
        "LEFT MARKER": ("ctrl+alt+j", "cc:22:127"),
        "RIGHT MARKER": ("ctrl+alt+k", "cc:23:127"),
        "D-PAD UP": ("up", "cc:24:127"),
        "D-PAD DOWN": ("down", "cc:25:127"),
        "D-PAD LEFT": ("left", "cc:26:127"),
        "D-PAD RIGHT": ("right", "cc:27:127"),
        "JOG LEFT": ("left", "cc:28:65"),
        "JOG RIGHT": ("right", "cc:28:63"),
        "SHUTTLE LEFT": ("j", "cc:30:65"),
        "SHUTTLE RIGHT": ("l", "cc:30:63"),
        "ON # LINE": ("ctrl+b", "note:63"),
        "ZOOM TO SELECTION": ("ctrl+e", "cc:72:127"),
    },
    "DAW / Bitwig": {
        "START": ("space", "start"),
        "STOP": ("space", "stop"),
        "RECORD": ("r", "record"),
        "UNDO": ("ctrl+z", "cc:70:127"),
        "DELETE": ("delete", "cc:71:127"),
        "SHIFT": ("shift", "cc:19:127"),
        "ZOOM IN": ("plus", "cc:20:1"),
        "ZOOM OUT": ("minus", "cc:21:1"),
        "LEFT MARKER": ("ctrl+shift+left", "cc:22:127"),
        "RIGHT MARKER": ("ctrl+shift+right", "cc:23:127"),
        "D-PAD UP": ("up", "cc:24:127"),
        "D-PAD DOWN": ("down", "cc:25:127"),
        "D-PAD LEFT": ("left", "cc:26:127"),
        "D-PAD RIGHT": ("right", "cc:27:127"),
        "JOG LEFT": ("left", "cc:28:65"),
        "JOG RIGHT": ("right", "cc:28:63"),
        "SHUTTLE LEFT": ("j", "cc:30:65"),
        "SHUTTLE RIGHT": ("l", "cc:30:63"),
        "ON # LINE": ("enter", "note:63"),
        "ZOOM TO SELECTION": ("z", "cc:72:127"),
    },
    "DAW / Cakewalk": {
        "START": ("space", "start"),
        "STOP": ("space", "stop"),
        "RECORD": ("r", "record"),
        "UNDO": ("ctrl+z", "cc:70:127"),
        "DELETE": ("delete", "cc:71:127"),
        "SHIFT": ("shift", "cc:19:127"),
        "ZOOM IN": ("ctrl+right", "cc:20:1"),
        "ZOOM OUT": ("ctrl+left", "cc:21:1"),
        "LEFT MARKER": ("ctrl+pageup", "cc:22:127"),
        "RIGHT MARKER": ("ctrl+pagedown", "cc:23:127"),
        "D-PAD UP": ("up", "cc:24:127"),
        "D-PAD DOWN": ("down", "cc:25:127"),
        "D-PAD LEFT": ("left", "cc:26:127"),
        "D-PAD RIGHT": ("right", "cc:27:127"),
        "JOG LEFT": ("left", "cc:28:65"),
        "JOG RIGHT": ("right", "cc:28:63"),
        "SHUTTLE LEFT": ("j", "cc:30:65"),
        "SHUTTLE RIGHT": ("l", "cc:30:63"),
        "ON # LINE": ("m", "note:63"),
        "ZOOM TO SELECTION": ("f", "cc:72:127"),
    },
    "DAW / Cubase/Nuendo": {
        "START": ("space", "start"),
        "STOP": ("num0", "stop"),
        "RECORD": ("num*", "record"),
        "UNDO": ("ctrl+z", "cc:70:127"),
        "DELETE": ("delete", "cc:71:127"),
        "SHIFT": ("shift", "cc:19:127"),
        "ZOOM IN": ("h", "cc:20:1"),
        "ZOOM OUT": ("g", "cc:21:1"),
        "LEFT MARKER": ("p", "cc:22:127"),
        "RIGHT MARKER": ("shift+p", "cc:23:127"),
        "D-PAD UP": ("up", "cc:24:127"),
        "D-PAD DOWN": ("down", "cc:25:127"),
        "D-PAD LEFT": ("left", "cc:26:127"),
        "D-PAD RIGHT": ("right", "cc:27:127"),
        "JOG LEFT": ("left", "cc:28:65"),
        "JOG RIGHT": ("right", "cc:28:63"),
        "SHUTTLE LEFT": ("j", "cc:30:65"),
        "SHUTTLE RIGHT": ("l", "cc:30:63"),
        "ON # LINE": ("enter", "note:63"),
        "ZOOM TO SELECTION": ("shift+f", "cc:72:127"),
    },
    "DAW / FL Studio": {
        "START": ("space", "start"),
        "STOP": ("space", "stop"),
        "RECORD": ("r", "record"),
        "UNDO": ("ctrl+alt+z", "cc:70:127"),
        "DELETE": ("delete", "cc:71:127"),
        "SHIFT": ("shift", "cc:19:127"),
        "ZOOM IN": ("ctrl+plus", "cc:20:1"),
        "ZOOM OUT": ("ctrl+minus", "cc:21:1"),
        "LEFT MARKER": ("home", "cc:22:127"),
        "RIGHT MARKER": ("end", "cc:23:127"),
        "D-PAD UP": ("up", "cc:24:127"),
        "D-PAD DOWN": ("down", "cc:25:127"),
        "D-PAD LEFT": ("left", "cc:26:127"),
        "D-PAD RIGHT": ("right", "cc:27:127"),
        "JOG LEFT": ("left", "cc:28:65"),
        "JOG RIGHT": ("right", "cc:28:63"),
        "SHUTTLE LEFT": ("j", "cc:30:65"),
        "SHUTTLE RIGHT": ("l", "cc:30:63"),
        "ON # LINE": ("enter", "note:63"),
        "ZOOM TO SELECTION": ("ctrl+enter", "cc:72:127"),
    },
    "DAW / Logic": {
        "START": ("space", "start"),
        "STOP": ("space", "stop"),
        "RECORD": ("r", "record"),
        "UNDO": ("command+z", "cc:70:127"),
        "DELETE": ("delete", "cc:71:127"),
        "SHIFT": ("shift", "cc:19:127"),
        "ZOOM IN": ("command+right", "cc:20:1"),
        "ZOOM OUT": ("command+left", "cc:21:1"),
        "LEFT MARKER": ("comma", "cc:22:127"),
        "RIGHT MARKER": ("period", "cc:23:127"),
        "D-PAD UP": ("up", "cc:24:127"),
        "D-PAD DOWN": ("down", "cc:25:127"),
        "D-PAD LEFT": ("left", "cc:26:127"),
        "D-PAD RIGHT": ("right", "cc:27:127"),
        "JOG LEFT": ("left", "cc:28:65"),
        "JOG RIGHT": ("right", "cc:28:63"),
        "SHUTTLE LEFT": ("j", "cc:30:65"),
        "SHUTTLE RIGHT": ("l", "cc:30:63"),
        "ON # LINE": ("return", "note:63"),
        "ZOOM TO SELECTION": ("z", "cc:72:127"),
    },
    "DAW / ProTools": {
        "START": ("space", "start"),
        "STOP": ("space", "stop"),
        "RECORD": ("ctrl+space", "record"),
        "UNDO": ("ctrl+z", "cc:70:127"),
        "DELETE": ("backspace", "cc:71:127"),
        "SHIFT": ("shift", "cc:19:127"),
        "ZOOM IN": ("ctrl+]", "cc:20:1"),
        "ZOOM OUT": ("ctrl+[", "cc:21:1"),
        "LEFT MARKER": ("comma", "cc:22:127"),
        "RIGHT MARKER": ("period", "cc:23:127"),
        "D-PAD UP": ("up", "cc:24:127"),
        "D-PAD DOWN": ("down", "cc:25:127"),
        "D-PAD LEFT": ("left", "cc:26:127"),
        "D-PAD RIGHT": ("right", "cc:27:127"),
        "JOG LEFT": ("left", "cc:28:65"),
        "JOG RIGHT": ("right", "cc:28:63"),
        "SHUTTLE LEFT": ("j", "cc:30:65"),
        "SHUTTLE RIGHT": ("l", "cc:30:63"),
        "ON # LINE": ("enter", "note:63"),
        "ZOOM TO SELECTION": ("e", "cc:72:127"),
    },
    "DAW / Pyramix": {
        "START": ("space", "start"),
        "STOP": ("space", "stop"),
        "RECORD": ("r", "record"),
        "UNDO": ("ctrl+z", "cc:70:127"),
        "DELETE": ("delete", "cc:71:127"),
        "SHIFT": ("shift", "cc:19:127"),
        "ZOOM IN": ("plus", "cc:20:1"),
        "ZOOM OUT": ("minus", "cc:21:1"),
        "LEFT MARKER": ("home", "cc:22:127"),
        "RIGHT MARKER": ("end", "cc:23:127"),
        "D-PAD UP": ("up", "cc:24:127"),
        "D-PAD DOWN": ("down", "cc:25:127"),
        "D-PAD LEFT": ("left", "cc:26:127"),
        "D-PAD RIGHT": ("right", "cc:27:127"),
        "JOG LEFT": ("left", "cc:28:65"),
        "JOG RIGHT": ("right", "cc:28:63"),
        "SHUTTLE LEFT": ("j", "cc:30:65"),
        "SHUTTLE RIGHT": ("l", "cc:30:63"),
        "ON # LINE": ("m", "note:63"),
        "ZOOM TO SELECTION": ("f", "cc:72:127"),
    },
    "DAW / Reason": {
        "START": ("space", "start"),
        "STOP": ("space", "stop"),
        "RECORD": ("num*", "record"),
        "UNDO": ("ctrl+z", "cc:70:127"),
        "DELETE": ("delete", "cc:71:127"),
        "SHIFT": ("shift", "cc:19:127"),
        "ZOOM IN": ("plus", "cc:20:1"),
        "ZOOM OUT": ("minus", "cc:21:1"),
        "LEFT MARKER": ("home", "cc:22:127"),
        "RIGHT MARKER": ("end", "cc:23:127"),
        "D-PAD UP": ("up", "cc:24:127"),
        "D-PAD DOWN": ("down", "cc:25:127"),
        "D-PAD LEFT": ("left", "cc:26:127"),
        "D-PAD RIGHT": ("right", "cc:27:127"),
        "JOG LEFT": ("left", "cc:28:65"),
        "JOG RIGHT": ("right", "cc:28:63"),
        "SHUTTLE LEFT": ("j", "cc:30:65"),
        "SHUTTLE RIGHT": ("l", "cc:30:63"),
        "ON # LINE": ("enter", "note:63"),
        "ZOOM TO SELECTION": ("f", "cc:72:127"),
    },
    "DAW / Reaper": {
        "START": ("space", "start"),
        "STOP": ("space", "stop"),
        "RECORD": ("ctrl+r", "record"),
        "UNDO": ("ctrl+z", "cc:70:127"),
        "DELETE": ("delete", "cc:71:127"),
        "SHIFT": ("shift", "cc:19:127"),
        "ZOOM IN": ("plus", "cc:20:1"),
        "ZOOM OUT": ("minus", "cc:21:1"),
        "LEFT MARKER": ("shift+j", "cc:22:127"),
        "RIGHT MARKER": ("shift+k", "cc:23:127"),
        "D-PAD UP": ("up", "cc:24:127"),
        "D-PAD DOWN": ("down", "cc:25:127"),
        "D-PAD LEFT": ("left", "cc:26:127"),
        "D-PAD RIGHT": ("right", "cc:27:127"),
        "JOG LEFT": ("alt+left", "cc:28:65"),
        "JOG RIGHT": ("alt+right", "cc:28:63"),
        "SHUTTLE LEFT": ("j", "cc:30:65"),
        "SHUTTLE RIGHT": ("l", "cc:30:63"),
        "ON # LINE": ("m", "note:63"),
        "ZOOM TO SELECTION": ("z", "cc:72:127"),
    },
    "DAW / Samplitude": {
        "START": ("space", "start"),
        "STOP": ("space", "stop"),
        "RECORD": ("r", "record"),
        "UNDO": ("ctrl+z", "cc:70:127"),
        "DELETE": ("delete", "cc:71:127"),
        "SHIFT": ("shift", "cc:19:127"),
        "ZOOM IN": ("plus", "cc:20:1"),
        "ZOOM OUT": ("minus", "cc:21:1"),
        "LEFT MARKER": ("home", "cc:22:127"),
        "RIGHT MARKER": ("end", "cc:23:127"),
        "D-PAD UP": ("up", "cc:24:127"),
        "D-PAD DOWN": ("down", "cc:25:127"),
        "D-PAD LEFT": ("left", "cc:26:127"),
        "D-PAD RIGHT": ("right", "cc:27:127"),
        "JOG LEFT": ("alt+left", "cc:28:65"),
        "JOG RIGHT": ("alt+right", "cc:28:63"),
        "SHUTTLE LEFT": ("j", "cc:30:65"),
        "SHUTTLE RIGHT": ("l", "cc:30:63"),
        "ON # LINE": ("m", "note:63"),
        "ZOOM TO SELECTION": ("f", "cc:72:127"),
    },
    "DAW / Sequoia": {
        "START": ("space", "start"),
        "STOP": ("space", "stop"),
        "RECORD": ("r", "record"),
        "UNDO": ("ctrl+z", "cc:70:127"),
        "DELETE": ("delete", "cc:71:127"),
        "SHIFT": ("shift", "cc:19:127"),
        "ZOOM IN": ("plus", "cc:20:1"),
        "ZOOM OUT": ("minus", "cc:21:1"),
        "LEFT MARKER": ("home", "cc:22:127"),
        "RIGHT MARKER": ("end", "cc:23:127"),
        "D-PAD UP": ("up", "cc:24:127"),
        "D-PAD DOWN": ("down", "cc:25:127"),
        "D-PAD LEFT": ("left", "cc:26:127"),
        "D-PAD RIGHT": ("right", "cc:27:127"),
        "JOG LEFT": ("alt+left", "cc:28:65"),
        "JOG RIGHT": ("alt+right", "cc:28:63"),
        "SHUTTLE LEFT": ("j", "cc:30:65"),
        "SHUTTLE RIGHT": ("l", "cc:30:63"),
        "ON # LINE": ("m", "note:63"),
        "ZOOM TO SELECTION": ("f", "cc:72:127"),
    },
    "DAW / Sound Forge": {
        "START": ("space", "start"),
        "STOP": ("space", "stop"),
        "RECORD": ("ctrl+r", "record"),
        "UNDO": ("ctrl+z", "cc:70:127"),
        "DELETE": ("delete", "cc:71:127"),
        "SHIFT": ("shift", "cc:19:127"),
        "ZOOM IN": ("up", "cc:20:1"),
        "ZOOM OUT": ("down", "cc:21:1"),
        "LEFT MARKER": ("home", "cc:22:127"),
        "RIGHT MARKER": ("end", "cc:23:127"),
        "D-PAD UP": ("pageup", "cc:24:127"),
        "D-PAD DOWN": ("pagedown", "cc:25:127"),
        "D-PAD LEFT": ("left", "cc:26:127"),
        "D-PAD RIGHT": ("right", "cc:27:127"),
        "JOG LEFT": ("left", "cc:28:65"),
        "JOG RIGHT": ("right", "cc:28:63"),
        "SHUTTLE LEFT": ("j", "cc:30:65"),
        "SHUTTLE RIGHT": ("l", "cc:30:63"),
        "ON # LINE": ("m", "note:63"),
        "ZOOM TO SELECTION": ("f", "cc:72:127"),
    },
    "DAW / Studio One": {
        "START": ("space", "start"),
        "STOP": ("num0", "stop"),
        "RECORD": ("num*", "record"),
        "UNDO": ("ctrl+z", "cc:70:127"),
        "DELETE": ("delete", "cc:71:127"),
        "SHIFT": ("shift", "cc:19:127"),
        "ZOOM IN": ("e", "cc:20:1"),
        "ZOOM OUT": ("w", "cc:21:1"),
        "LEFT MARKER": ("comma", "cc:22:127"),
        "RIGHT MARKER": ("period", "cc:23:127"),
        "D-PAD UP": ("up", "cc:24:127"),
        "D-PAD DOWN": ("down", "cc:25:127"),
        "D-PAD LEFT": ("left", "cc:26:127"),
        "D-PAD RIGHT": ("right", "cc:27:127"),
        "JOG LEFT": ("alt+left", "cc:28:65"),
        "JOG RIGHT": ("alt+right", "cc:28:63"),
        "SHUTTLE LEFT": ("j", "cc:30:65"),
        "SHUTTLE RIGHT": ("l", "cc:30:63"),
        "ON # LINE": ("y", "note:63"),
        "ZOOM TO SELECTION": ("shift+z", "cc:72:127"),
    },
    "DAW / WaveLab": {
        "START": ("space", "start"),
        "STOP": ("space", "stop"),
        "RECORD": ("num*", "record"),
        "UNDO": ("ctrl+z", "cc:70:127"),
        "DELETE": ("delete", "cc:71:127"),
        "SHIFT": ("shift", "cc:19:127"),
        "ZOOM IN": ("h", "cc:20:1"),
        "ZOOM OUT": ("g", "cc:21:1"),
        "LEFT MARKER": ("home", "cc:22:127"),
        "RIGHT MARKER": ("end", "cc:23:127"),
        "D-PAD UP": ("pageup", "cc:24:127"),
        "D-PAD DOWN": ("pagedown", "cc:25:127"),
        "D-PAD LEFT": ("left", "cc:26:127"),
        "D-PAD RIGHT": ("right", "cc:27:127"),
        "JOG LEFT": ("alt+left", "cc:28:65"),
        "JOG RIGHT": ("alt+right", "cc:28:63"),
        "SHUTTLE LEFT": ("j", "cc:30:65"),
        "SHUTTLE RIGHT": ("l", "cc:30:63"),
        "ON # LINE": ("m", "note:63"),
        "ZOOM TO SELECTION": ("f", "cc:72:127"),
    },

    # NLE presets
    "NLE / DaVinci Resolve": {
        "START": ("space", "start"),
        "STOP": ("k", "stop"),
        "RECORD": ("f12", "record"),
        "UNDO": ("ctrl+z", "cc:70:127"),
        "DELETE": ("backspace", "cc:71:127"),
        "SHIFT": ("shift", "cc:19:127"),
        "ZOOM IN": ("plus", "cc:20:1"),
        "ZOOM OUT": ("minus", "cc:21:1"),
        "LEFT MARKER": ("up", "cc:22:127"),
        "RIGHT MARKER": ("down", "cc:23:127"),
        "D-PAD UP": ("pageup", "cc:24:127"),
        "D-PAD DOWN": ("pagedown", "cc:25:127"),
        "D-PAD LEFT": ("left", "cc:26:127"),
        "D-PAD RIGHT": ("right", "cc:27:127"),
        "JOG LEFT": ("j", "cc:28:65"),
        "JOG RIGHT": ("l", "cc:28:63"),
        "SHUTTLE LEFT": ("shift+j", "cc:30:65"),
        "SHUTTLE RIGHT": ("shift+l", "cc:30:63"),
        "ON # LINE": ("enter", "note:63"),
        "ZOOM TO SELECTION": ("shift+z", "cc:72:127"),
    },
    "NLE / Final Cut Pro": {
        "START": ("space", "start"),
        "STOP": ("k", "stop"),
        "RECORD": ("command+shift+2", "record"),
        "UNDO": ("command+z", "cc:70:127"),
        "DELETE": ("delete", "cc:71:127"),
        "SHIFT": ("shift", "cc:19:127"),
        "ZOOM IN": ("command+plus", "cc:20:1"),
        "ZOOM OUT": ("command+minus", "cc:21:1"),
        "LEFT MARKER": ("m", "cc:22:127"),
        "RIGHT MARKER": ("option+m", "cc:23:127"),
        "D-PAD UP": ("up", "cc:24:127"),
        "D-PAD DOWN": ("down", "cc:25:127"),
        "D-PAD LEFT": ("left", "cc:26:127"),
        "D-PAD RIGHT": ("right", "cc:27:127"),
        "JOG LEFT": ("j", "cc:28:65"),
        "JOG RIGHT": ("l", "cc:28:63"),
        "SHUTTLE LEFT": ("shift+j", "cc:30:65"),
        "SHUTTLE RIGHT": ("shift+l", "cc:30:63"),
        "ON # LINE": ("enter", "note:63"),
        "ZOOM TO SELECTION": ("shift+z", "cc:72:127"),
    },
    "NLE / Lightworks": {
        "START": ("space", "start"),
        "STOP": ("k", "stop"),
        "RECORD": ("r", "record"),
        "UNDO": ("ctrl+z", "cc:70:127"),
        "DELETE": ("delete", "cc:71:127"),
        "SHIFT": ("shift", "cc:19:127"),
        "ZOOM IN": ("plus", "cc:20:1"),
        "ZOOM OUT": ("minus", "cc:21:1"),
        "LEFT MARKER": ("i", "cc:22:127"),
        "RIGHT MARKER": ("o", "cc:23:127"),
        "D-PAD UP": ("pageup", "cc:24:127"),
        "D-PAD DOWN": ("pagedown", "cc:25:127"),
        "D-PAD LEFT": ("left", "cc:26:127"),
        "D-PAD RIGHT": ("right", "cc:27:127"),
        "JOG LEFT": ("j", "cc:28:65"),
        "JOG RIGHT": ("l", "cc:28:63"),
        "SHUTTLE LEFT": ("shift+j", "cc:30:65"),
        "SHUTTLE RIGHT": ("shift+l", "cc:30:63"),
        "ON # LINE": ("enter", "note:63"),
        "ZOOM TO SELECTION": ("z", "cc:72:127"),
    },
    "NLE / Media Composer": {
        "START": ("space", "start"),
        "STOP": ("k", "stop"),
        "RECORD": ("ctrl+r", "record"),
        "UNDO": ("ctrl+z", "cc:70:127"),
        "DELETE": ("delete", "cc:71:127"),
        "SHIFT": ("shift", "cc:19:127"),
        "ZOOM IN": ("plus", "cc:20:1"),
        "ZOOM OUT": ("minus", "cc:21:1"),
        "LEFT MARKER": ("i", "cc:22:127"),
        "RIGHT MARKER": ("o", "cc:23:127"),
        "D-PAD UP": ("pageup", "cc:24:127"),
        "D-PAD DOWN": ("pagedown", "cc:25:127"),
        "D-PAD LEFT": ("left", "cc:26:127"),
        "D-PAD RIGHT": ("right", "cc:27:127"),
        "JOG LEFT": ("j", "cc:28:65"),
        "JOG RIGHT": ("l", "cc:28:63"),
        "SHUTTLE LEFT": ("shift+j", "cc:30:65"),
        "SHUTTLE RIGHT": ("shift+l", "cc:30:63"),
        "ON # LINE": ("enter", "note:63"),
        "ZOOM TO SELECTION": ("t", "cc:72:127"),
    },
    "NLE / Premiere Pro": {
        "START": ("space", "start"),
        "STOP": ("k", "stop"),
        "RECORD": ("shift+space", "record"),
        "UNDO": ("ctrl+z", "cc:70:127"),
        "DELETE": ("delete", "cc:71:127"),
        "SHIFT": ("shift", "cc:19:127"),
        "ZOOM IN": ("equal", "cc:20:1"),
        "ZOOM OUT": ("minus", "cc:21:1"),
        "LEFT MARKER": ("i", "cc:22:127"),
        "RIGHT MARKER": ("o", "cc:23:127"),
        "D-PAD UP": ("pageup", "cc:24:127"),
        "D-PAD DOWN": ("pagedown", "cc:25:127"),
        "D-PAD LEFT": ("left", "cc:26:127"),
        "D-PAD RIGHT": ("right", "cc:27:127"),
        "JOG LEFT": ("j", "cc:28:65"),
        "JOG RIGHT": ("l", "cc:28:63"),
        "SHUTTLE LEFT": ("shift+j", "cc:30:65"),
        "SHUTTLE RIGHT": ("shift+l", "cc:30:63"),
        "ON # LINE": ("enter", "note:63"),
        "ZOOM TO SELECTION": ("backslash", "cc:72:127"),
    },
    "NLE / Vegas Pro": {
        "START": ("space", "start"),
        "STOP": ("k", "stop"),
        "RECORD": ("ctrl+r", "record"),
        "UNDO": ("ctrl+z", "cc:70:127"),
        "DELETE": ("delete", "cc:71:127"),
        "SHIFT": ("shift", "cc:19:127"),
        "ZOOM IN": ("up", "cc:20:1"),
        "ZOOM OUT": ("down", "cc:21:1"),
        "LEFT MARKER": ("i", "cc:22:127"),
        "RIGHT MARKER": ("o", "cc:23:127"),
        "D-PAD UP": ("pageup", "cc:24:127"),
        "D-PAD DOWN": ("pagedown", "cc:25:127"),
        "D-PAD LEFT": ("left", "cc:26:127"),
        "D-PAD RIGHT": ("right", "cc:27:127"),
        "JOG LEFT": ("j", "cc:28:65"),
        "JOG RIGHT": ("l", "cc:28:63"),
        "SHUTTLE LEFT": ("shift+j", "cc:30:65"),
        "SHUTTLE RIGHT": ("shift+l", "cc:30:63"),
        "ON # LINE": ("m", "note:63"),
        "ZOOM TO SELECTION": ("f", "cc:72:127"),
    },

    # Playout presets. These are conservative defaults because most playout systems use custom hotkey profiles.
    "Playout / AudioVault": {
        "START": ("f1", "start"),
        "STOP": ("f2", "stop"),
        "RECORD": ("f12", "record"),
        "UNDO": ("ctrl+z", "cc:70:127"),
        "DELETE": ("delete", "cc:71:127"),
        "SHIFT": ("shift", "cc:19:127"),
        "ZOOM IN": ("plus", "cc:20:1"),
        "ZOOM OUT": ("minus", "cc:21:1"),
        "LEFT MARKER": ("home", "cc:22:127"),
        "RIGHT MARKER": ("end", "cc:23:127"),
        "D-PAD UP": ("up", "cc:24:127"),
        "D-PAD DOWN": ("down", "cc:25:127"),
        "D-PAD LEFT": ("pageup", "cc:26:127"),
        "D-PAD RIGHT": ("pagedown", "cc:27:127"),
        "JOG LEFT": ("left", "cc:28:65"),
        "JOG RIGHT": ("right", "cc:28:63"),
        "SHUTTLE LEFT": ("shift+left", "cc:30:65"),
        "SHUTTLE RIGHT": ("shift+right", "cc:30:63"),
        "ON # LINE": ("enter", "note:63"),
        "ZOOM TO SELECTION": ("f", "cc:72:127"),
    },
    "Playout / mAirList": {
        "START": ("f1", "start"),
        "STOP": ("f2", "stop"),
        "RECORD": ("f12", "record"),
        "UNDO": ("ctrl+z", "cc:70:127"),
        "DELETE": ("delete", "cc:71:127"),
        "SHIFT": ("shift", "cc:19:127"),
        "ZOOM IN": ("plus", "cc:20:1"),
        "ZOOM OUT": ("minus", "cc:21:1"),
        "LEFT MARKER": ("home", "cc:22:127"),
        "RIGHT MARKER": ("end", "cc:23:127"),
        "D-PAD UP": ("up", "cc:24:127"),
        "D-PAD DOWN": ("down", "cc:25:127"),
        "D-PAD LEFT": ("left", "cc:26:127"),
        "D-PAD RIGHT": ("right", "cc:27:127"),
        "JOG LEFT": ("ctrl+left", "cc:28:65"),
        "JOG RIGHT": ("ctrl+right", "cc:28:63"),
        "SHUTTLE LEFT": ("shift+left", "cc:30:65"),
        "SHUTTLE RIGHT": ("shift+right", "cc:30:63"),
        "ON # LINE": ("enter", "note:63"),
        "ZOOM TO SELECTION": ("f", "cc:72:127"),
    },
    "Playout / Myriad": {
        "START": ("f1", "start"),
        "STOP": ("f2", "stop"),
        "RECORD": ("f12", "record"),
        "UNDO": ("ctrl+z", "cc:70:127"),
        "DELETE": ("delete", "cc:71:127"),
        "SHIFT": ("shift", "cc:19:127"),
        "ZOOM IN": ("plus", "cc:20:1"),
        "ZOOM OUT": ("minus", "cc:21:1"),
        "LEFT MARKER": ("home", "cc:22:127"),
        "RIGHT MARKER": ("end", "cc:23:127"),
        "D-PAD UP": ("up", "cc:24:127"),
        "D-PAD DOWN": ("down", "cc:25:127"),
        "D-PAD LEFT": ("pageup", "cc:26:127"),
        "D-PAD RIGHT": ("pagedown", "cc:27:127"),
        "JOG LEFT": ("left", "cc:28:65"),
        "JOG RIGHT": ("right", "cc:28:63"),
        "SHUTTLE LEFT": ("shift+left", "cc:30:65"),
        "SHUTTLE RIGHT": ("shift+right", "cc:30:63"),
        "ON # LINE": ("enter", "note:63"),
        "ZOOM TO SELECTION": ("f", "cc:72:127"),
    },
    "Playout / NexGen": {
        "START": ("f1", "start"),
        "STOP": ("f2", "stop"),
        "RECORD": ("f12", "record"),
        "UNDO": ("ctrl+z", "cc:70:127"),
        "DELETE": ("delete", "cc:71:127"),
        "SHIFT": ("shift", "cc:19:127"),
        "ZOOM IN": ("plus", "cc:20:1"),
        "ZOOM OUT": ("minus", "cc:21:1"),
        "LEFT MARKER": ("home", "cc:22:127"),
        "RIGHT MARKER": ("end", "cc:23:127"),
        "D-PAD UP": ("up", "cc:24:127"),
        "D-PAD DOWN": ("down", "cc:25:127"),
        "D-PAD LEFT": ("pageup", "cc:26:127"),
        "D-PAD RIGHT": ("pagedown", "cc:27:127"),
        "JOG LEFT": ("left", "cc:28:65"),
        "JOG RIGHT": ("right", "cc:28:63"),
        "SHUTTLE LEFT": ("shift+left", "cc:30:65"),
        "SHUTTLE RIGHT": ("shift+right", "cc:30:63"),
        "ON # LINE": ("enter", "note:63"),
        "ZOOM TO SELECTION": ("f", "cc:72:127"),
    },
    "Playout / OmniPlayer": {
        "START": ("f1", "start"),
        "STOP": ("f2", "stop"),
        "RECORD": ("f12", "record"),
        "UNDO": ("ctrl+z", "cc:70:127"),
        "DELETE": ("delete", "cc:71:127"),
        "SHIFT": ("shift", "cc:19:127"),
        "ZOOM IN": ("plus", "cc:20:1"),
        "ZOOM OUT": ("minus", "cc:21:1"),
        "LEFT MARKER": ("home", "cc:22:127"),
        "RIGHT MARKER": ("end", "cc:23:127"),
        "D-PAD UP": ("up", "cc:24:127"),
        "D-PAD DOWN": ("down", "cc:25:127"),
        "D-PAD LEFT": ("pageup", "cc:26:127"),
        "D-PAD RIGHT": ("pagedown", "cc:27:127"),
        "JOG LEFT": ("left", "cc:28:65"),
        "JOG RIGHT": ("right", "cc:28:63"),
        "SHUTTLE LEFT": ("shift+left", "cc:30:65"),
        "SHUTTLE RIGHT": ("shift+right", "cc:30:63"),
        "ON # LINE": ("enter", "note:63"),
        "ZOOM TO SELECTION": ("f", "cc:72:127"),
    },
    "Playout / PlayoutONE": {
        "START": ("f1", "start"),
        "STOP": ("f2", "stop"),
        "RECORD": ("f12", "record"),
        "UNDO": ("ctrl+z", "cc:70:127"),
        "DELETE": ("delete", "cc:71:127"),
        "SHIFT": ("shift", "cc:19:127"),
        "ZOOM IN": ("plus", "cc:20:1"),
        "ZOOM OUT": ("minus", "cc:21:1"),
        "LEFT MARKER": ("home", "cc:22:127"),
        "RIGHT MARKER": ("end", "cc:23:127"),
        "D-PAD UP": ("up", "cc:24:127"),
        "D-PAD DOWN": ("down", "cc:25:127"),
        "D-PAD LEFT": ("pageup", "cc:26:127"),
        "D-PAD RIGHT": ("pagedown", "cc:27:127"),
        "JOG LEFT": ("left", "cc:28:65"),
        "JOG RIGHT": ("right", "cc:28:63"),
        "SHUTTLE LEFT": ("shift+left", "cc:30:65"),
        "SHUTTLE RIGHT": ("shift+right", "cc:30:63"),
        "ON # LINE": ("enter", "note:63"),
        "ZOOM TO SELECTION": ("f", "cc:72:127"),
    },
    "Playout / RadioDJ": {
        "START": ("f1", "start"),
        "STOP": ("f2", "stop"),
        "RECORD": ("f12", "record"),
        "UNDO": ("ctrl+z", "cc:70:127"),
        "DELETE": ("delete", "cc:71:127"),
        "SHIFT": ("shift", "cc:19:127"),
        "ZOOM IN": ("plus", "cc:20:1"),
        "ZOOM OUT": ("minus", "cc:21:1"),
        "LEFT MARKER": ("home", "cc:22:127"),
        "RIGHT MARKER": ("end", "cc:23:127"),
        "D-PAD UP": ("up", "cc:24:127"),
        "D-PAD DOWN": ("down", "cc:25:127"),
        "D-PAD LEFT": ("pageup", "cc:26:127"),
        "D-PAD RIGHT": ("pagedown", "cc:27:127"),
        "JOG LEFT": ("left", "cc:28:65"),
        "JOG RIGHT": ("right", "cc:28:63"),
        "SHUTTLE LEFT": ("shift+left", "cc:30:65"),
        "SHUTTLE RIGHT": ("shift+right", "cc:30:63"),
        "ON # LINE": ("enter", "note:63"),
        "ZOOM TO SELECTION": ("f", "cc:72:127"),
    },
    "Playout / RCS Zetta": {
        "START": ("f1", "start"),
        "STOP": ("f2", "stop"),
        "RECORD": ("f12", "record"),
        "UNDO": ("ctrl+z", "cc:70:127"),
        "DELETE": ("delete", "cc:71:127"),
        "SHIFT": ("shift", "cc:19:127"),
        "ZOOM IN": ("plus", "cc:20:1"),
        "ZOOM OUT": ("minus", "cc:21:1"),
        "LEFT MARKER": ("home", "cc:22:127"),
        "RIGHT MARKER": ("end", "cc:23:127"),
        "D-PAD UP": ("up", "cc:24:127"),
        "D-PAD DOWN": ("down", "cc:25:127"),
        "D-PAD LEFT": ("pageup", "cc:26:127"),
        "D-PAD RIGHT": ("pagedown", "cc:27:127"),
        "JOG LEFT": ("left", "cc:28:65"),
        "JOG RIGHT": ("right", "cc:28:63"),
        "SHUTTLE LEFT": ("shift+left", "cc:30:65"),
        "SHUTTLE RIGHT": ("shift+right", "cc:30:63"),
        "ON # LINE": ("enter", "note:63"),
        "ZOOM TO SELECTION": ("f", "cc:72:127"),
    },
    "Playout / Simian": {
        "START": ("f1", "start"),
        "STOP": ("f2", "stop"),
        "RECORD": ("f12", "record"),
        "UNDO": ("ctrl+z", "cc:70:127"),
        "DELETE": ("delete", "cc:71:127"),
        "SHIFT": ("shift", "cc:19:127"),
        "ZOOM IN": ("plus", "cc:20:1"),
        "ZOOM OUT": ("minus", "cc:21:1"),
        "LEFT MARKER": ("home", "cc:22:127"),
        "RIGHT MARKER": ("end", "cc:23:127"),
        "D-PAD UP": ("up", "cc:24:127"),
        "D-PAD DOWN": ("down", "cc:25:127"),
        "D-PAD LEFT": ("pageup", "cc:26:127"),
        "D-PAD RIGHT": ("pagedown", "cc:27:127"),
        "JOG LEFT": ("left", "cc:28:65"),
        "JOG RIGHT": ("right", "cc:28:63"),
        "SHUTTLE LEFT": ("shift+left", "cc:30:65"),
        "SHUTTLE RIGHT": ("shift+right", "cc:30:63"),
        "ON # LINE": ("enter", "note:63"),
        "ZOOM TO SELECTION": ("f", "cc:72:127"),
    },
    "Playout / WideOrbit": {
        "START": ("f1", "start"),
        "STOP": ("f2", "stop"),
        "RECORD": ("f12", "record"),
        "UNDO": ("ctrl+z", "cc:70:127"),
        "DELETE": ("delete", "cc:71:127"),
        "SHIFT": ("shift", "cc:19:127"),
        "ZOOM IN": ("plus", "cc:20:1"),
        "ZOOM OUT": ("minus", "cc:21:1"),
        "LEFT MARKER": ("home", "cc:22:127"),
        "RIGHT MARKER": ("end", "cc:23:127"),
        "D-PAD UP": ("up", "cc:24:127"),
        "D-PAD DOWN": ("down", "cc:25:127"),
        "D-PAD LEFT": ("pageup", "cc:26:127"),
        "D-PAD RIGHT": ("pagedown", "cc:27:127"),
        "JOG LEFT": ("left", "cc:28:65"),
        "JOG RIGHT": ("right", "cc:28:63"),
        "SHUTTLE LEFT": ("shift+left", "cc:30:65"),
        "SHUTTLE RIGHT": ("shift+right", "cc:30:63"),
        "ON # LINE": ("enter", "note:63"),
        "ZOOM TO SELECTION": ("f", "cc:72:127"),
    },
}
SETTINGS_PATH = Path.home() / ".eela_controller_bridge_settings.json"

DEFAULT_MAPPINGS = [
    # name, hex pattern, action_type, shortcut_action, midi_action, shift_shortcut_action, shift_midi_action
    ("START", "00 00 00 f8", KEYBOARD_ACTION, "ctrl+space", "start", "shift+space", "continue"),
    ("STOP", "00 78 00 f8", KEYBOARD_ACTION, "space", "stop", "esc", "cc:90:127"),
    ("RECORD", "00 00 f8", KEYBOARD_ACTION, "ctrl+r", "note:60", "ctrl+shift+r", "cc:91:127"),
    ("SHIFT", "00 f8 80 f8", KEYBOARD_ACTION, "shift", "cc:10:127", "", "", True, "as_shift_layer"),
    ("ZOOM IN", "00 80 80 f8", KEYBOARD_ACTION, "ctrl+plus", "cc:20:127", "ctrl+shift+plus", "cc:92:127"),
    ("ZOOM OUT", "00 78 80 f8", KEYBOARD_ACTION, "ctrl+minus", "cc:21:127", "ctrl+shift+minus", "cc:93:127"),
    ("DELETE", "00 00 80 f8", KEYBOARD_ACTION, "delete", "note:61", "shift+delete", "cc:94:127"),
    ("UNDO", "00 80 00 f8 00", KEYBOARD_ACTION, "ctrl+z", "note:62", "ctrl+shift+z", "cc:95:127"),
    ("LEFT MARKER", "00 f8 00 f8", KEYBOARD_ACTION, "ctrl+left", "cc:22:127", "shift+left", "cc:96:127"),
    ("RIGHT MARKER", "00 80 00 f8", KEYBOARD_ACTION, "ctrl+right", "cc:23:127", "shift+right", "cc:97:127"),
    ("D-PAD UP", "00 78 80 f8", KEYBOARD_ACTION, "up", "cc:24:127", "pageup", "cc:98:127"),
    ("D-PAD DOWN", "00 00 80 f8", KEYBOARD_ACTION, "down", "cc:25:127", "pagedown", "cc:99:127"),
    ("D-PAD LEFT", "00 80 78 f8", KEYBOARD_ACTION, "left", "cc:26:127", "home", "cc:100:127"),
    ("D-PAD RIGHT", "00 f8 78 f8", KEYBOARD_ACTION, "right", "cc:27:127", "end", "cc:101:127"),
    ("ON # LINE", "00 78 f8", KEYBOARD_ACTION, "enter", "note:63", "shift+enter", "cc:102:127"),
    ("JOG LEFT", "78 78 f8", KEYBOARD_ACTION, "left", "cc:28:127", "ctrl+left", "cc:103:65"),
    ("JOG RIGHT", "78 f8 00", KEYBOARD_ACTION, "right", "cc:29:127", "ctrl+right", "cc:103:63"),
    ("SHUTTLE LEFT", "00 f8 00 78 f8 00", KEYBOARD_ACTION, "j", "cc:30:127", "shift+j", "cc:104:65"),
    ("SHUTTLE RIGHT", "00 80 00 78 80 00", KEYBOARD_ACTION, "l", "cc:31:127", "shift+l", "cc:104:63"),
]


@dataclass
class Mapping:
    name: str
    hex_pattern: str
    action_type: str = KEYBOARD_ACTION
    shortcut_action: str = ""  # e.g. ctrl+shift+z, space, enter
    midi_action: str = ""      # e.g. start, stop, note:60, cc:20:127
    shift_shortcut_action: str = ""
    shift_midi_action: str = ""
    enabled: bool = True
    shift_mode: str = "yes"

    @property
    def pattern_bytes(self) -> bytes:
        return bytes.fromhex(clean_hex(self.hex_pattern))

    @classmethod
    def from_dict(cls, data: dict) -> "Mapping":
        """Load both new mapping files and older files that used action_value."""
        action_type = normalize_action_type(data.get("action_type", KEYBOARD_ACTION))
        old_value = data.get("action_value", "")
        shortcut_action = data.get("shortcut_action", "")
        midi_action = data.get("midi_action", "")

        if old_value and not shortcut_action and not midi_action:
            if action_type == MIDI_ACTION:
                midi_action = old_value
            else:
                shortcut_action = old_value

        return cls(
            name=data.get("name", ""),
            hex_pattern=data.get("hex_pattern", ""),
            action_type=action_type,
            shortcut_action=shortcut_action,
            midi_action=midi_action,
            shift_shortcut_action=data.get("shift_shortcut_action", ""),
            shift_midi_action=data.get("shift_midi_action", ""),
            enabled=bool(data.get("enabled", True)),
            shift_mode=data.get("shift_mode", "as_shift_layer" if data.get("name", "").upper() == "SHIFT" else "yes"),
        )


def normalize_action_type(value: str) -> str:
    value = str(value).strip().lower()
    if value in {"midi", "midi action"}:
        return MIDI_ACTION
    return KEYBOARD_ACTION


def clean_hex(value: str) -> str:
    return value.replace("0x", "").replace(",", " ").replace("-", " ").strip()


def normalize_hex(data: bytes) -> str:
    return data.hex(" ")


def load_settings() -> dict:
    if not SETTINGS_PATH.exists():
        return {}
    try:
        with open(SETTINGS_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_settings(settings: dict):
    try:
        with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
            json.dump(settings, f, indent=2)
    except Exception:
        pass


def app_command() -> str:
    script = Path(__file__).resolve()
    python_exe = Path(sys.executable).resolve()
    return f'"{python_exe}" "{script}"'


def windows_startup_file() -> Path:
    startup = Path(os.environ.get("APPDATA", "")) / "Microsoft" / "Windows" / "Start Menu" / "Programs" / "Startup"
    return startup / f"{APP_NAME}.bat"


def linux_autostart_file() -> Path:
    return Path.home() / ".config" / "autostart" / f"{APP_NAME}.desktop"


def macos_plist_file() -> Path:
    return Path.home() / "Library" / "LaunchAgents" / f"com.eela.{APP_NAME}.plist"


def is_startup_enabled() -> bool:
    if sys.platform.startswith("win"):
        return windows_startup_file().exists()
    if sys.platform == "darwin":
        return macos_plist_file().exists()
    return linux_autostart_file().exists()


def set_startup_enabled(enabled: bool):
    command = app_command()

    if sys.platform.startswith("win"):
        path = windows_startup_file()
        path.parent.mkdir(parents=True, exist_ok=True)
        if enabled:
            path.write_text(f"@echo off\nstart \"\" {command}\n", encoding="utf-8")
        elif path.exists():
            path.unlink()
        return

    if sys.platform == "darwin":
        path = macos_plist_file()
        path.parent.mkdir(parents=True, exist_ok=True)
        if enabled:
            plist = f'''<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>Label</key>
    <string>com.eela.{APP_NAME}</string>
    <key>ProgramArguments</key>
    <array>
        <string>{sys.executable}</string>
        <string>{Path(__file__).resolve()}</string>
    </array>
    <key>RunAtLoad</key>
    <true/>
</dict>
</plist>
'''
            path.write_text(plist, encoding="utf-8")
        elif path.exists():
            path.unlink()
        return

    path = linux_autostart_file()
    path.parent.mkdir(parents=True, exist_ok=True)
    if enabled:
        desktop = f'''[Desktop Entry]
Type=Application
Name=Eela Controller Bridge
Exec={command}
X-GNOME-Autostart-enabled=true
Terminal=false
'''
        path.write_text(desktop, encoding="utf-8")
    elif path.exists():
        path.unlink()


class ShortcutCaptureDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Capture shortcut")
        self.setModal(True)
        self.setMinimumWidth(360)
        self.setFocusPolicy(Qt.StrongFocus)
        self.setFocus()
        self.captured_shortcut = ""
        self.pressed_keys: set[str] = set()

        layout = QVBoxLayout(self)
        info_text = "Press the shortcut now...\nExample: Ctrl + Shift + Z"
        self.info_label = QLabel(info_text)
        self.info_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.info_label)

        self.preview_label = QLabel("")
        self.preview_label.setAlignment(Qt.AlignCenter)
        self.preview_label.setStyleSheet("font-weight: bold; font-size: 16px;")
        layout.addWidget(self.preview_label)

        buttons = QHBoxLayout()
        cancel_btn = QPushButton("Cancel")
        clear_btn = QPushButton("Clear")
        use_btn = QPushButton("Use shortcut")

        for btn in (cancel_btn, clear_btn, use_btn):
            btn.setFocusPolicy(Qt.NoFocus)
        buttons.addWidget(cancel_btn)
        buttons.addWidget(clear_btn)
        buttons.addStretch(1)
        buttons.addWidget(use_btn)
        layout.addLayout(buttons)

        self.setFocus()

        cancel_btn.clicked.connect(self.reject)
        clear_btn.clicked.connect(self.clear_shortcut)
        use_btn.clicked.connect(self.accept_if_available)

    def keyPressEvent(self, event):
        key_name = self.qt_key_to_name(event)
        if key_name:
            self.captured_shortcut = key_name
            self.preview_label.setText(self.captured_shortcut)
        event.accept()
        return

    def keyReleaseEvent(self, event):
        event.accept()

    def qt_key_to_name(self, event) -> str:
        key = event.key()
        modifiers = event.modifiers()

        modifier_key_codes = {
            Qt.Key_Control,
            Qt.Key_Shift,
            Qt.Key_Alt,
            Qt.Key_Meta,
        }

        # If the user presses only a modifier, capture that modifier by itself.
        if key in modifier_key_codes:
            if key == Qt.Key_Control:
                return "ctrl"
            if key == Qt.Key_Shift:
                return "shift"
            if key == Qt.Key_Alt:
                return "alt"
            if key == Qt.Key_Meta:
                return "command" if sys.platform == "darwin" else "win"

        parts = []
        if modifiers & Qt.ControlModifier:
            parts.append("ctrl")
        if modifiers & Qt.ShiftModifier:
            parts.append("shift")
        if modifiers & Qt.AltModifier:
            parts.append("alt")
        if modifiers & Qt.MetaModifier:
            parts.append("command" if sys.platform == "darwin" else "win")

        special = {
            Qt.Key_Space: "space",
            Qt.Key_Return: "enter",
            Qt.Key_Enter: "enter",
            Qt.Key_Tab: "tab",
            Qt.Key_Escape: "esc",
            Qt.Key_Backspace: "backspace",
            Qt.Key_Delete: "delete",
            Qt.Key_Left: "left",
            Qt.Key_Right: "right",
            Qt.Key_Up: "up",
            Qt.Key_Down: "down",
            Qt.Key_Home: "home",
            Qt.Key_End: "end",
            Qt.Key_PageUp: "pgup",
            Qt.Key_PageDown: "pgdn",
            Qt.Key_Plus: "plus",
            Qt.Key_Minus: "minus",
            Qt.Key_BracketLeft: "[",
            Qt.Key_BracketRight: "]",
            Qt.Key_Asterisk: "*",
            Qt.Key_Slash: "/",
            Qt.Key_Backslash: "\\",
            Qt.Key_Comma: ",",
            Qt.Key_Period: ".",
        }

        if Qt.Key_F1 <= key <= Qt.Key_F35:
            main_key = f"f{key - Qt.Key_F1 + 1}"
        elif key in special:
            main_key = special[key]
        elif Qt.Key_A <= key <= Qt.Key_Z:
            main_key = chr(ord("a") + key - Qt.Key_A)
        elif Qt.Key_0 <= key <= Qt.Key_9:
            main_key = chr(ord("0") + key - Qt.Key_0)
        else:
            text = event.text().lower()
            main_key = text if text and text.isprintable() else ""

        if main_key:
            parts.append(main_key)

        # remove duplicates while preserving order
        result = []
        for part in parts:
            if part not in result:
                result.append(part)
        return "+".join(result)

    def update_preview(self):
        if self.captured_shortcut:
            self.preview_label.setText(self.captured_shortcut)

    def clear_shortcut(self):
        self.pressed_keys.clear()
        self.captured_shortcut = ""
        self.preview_label.setText("")

    def accept_if_available(self):
        if self.captured_shortcut:
            self.accept()


class SerialWorker(QObject):
    packet = Signal(bytes)
    status = Signal(str)
    error = Signal(str)

    def __init__(self, port: str, baudrate: int = 9600):
        super().__init__()
        self.port = port
        self.baudrate = baudrate
        self._running = False
        self._ser: Optional[serial.Serial] = None

    @Slot()
    def run(self):
        self._running = True
        try:
            self._ser = serial.Serial(
                self.port,
                baudrate=self.baudrate,
                bytesize=8,
                parity="N",
                stopbits=1,
                timeout=0.05,
                xonxoff=False,
                rtscts=False,
                dsrdtr=False,
            )
            self.status.emit(f"Listening on {self.port} @ {self.baudrate} baud")
        except Exception as exc:
            self.error.emit(f"Could not open serial port: {exc}")
            return

        while self._running:
            try:
                data = self._ser.read(64)
                if data:
                    self.packet.emit(data)
            except Exception as exc:
                self.error.emit(f"Serial read error: {exc}")
                break

        try:
            if self._ser and self._ser.is_open:
                self._ser.close()
        finally:
            self.status.emit("Serial stopped")

    def stop(self):
        self._running = False


class ActionRunner:
    def __init__(self):
        pyautogui.PAUSE = 0
        self.midi_out = None
        self.midi_port_name = ""

    def set_midi_port(self, port_name: str):
        self.midi_port_name = port_name
        if self.midi_out:
            try:
                self.midi_out.close()
            except Exception:
                pass
            self.midi_out = None

        if not port_name or mido is None:
            return

        self.midi_out = mido.open_output(port_name)

    def run(self, mapping: Mapping, shift_active: bool = False):
        if not mapping.enabled:
            return

        if normalize_action_type(mapping.action_type) == KEYBOARD_ACTION:
            action = mapping.shift_shortcut_action if shift_active and mapping.shift_shortcut_action else mapping.shortcut_action
            self.send_hotkey(action)
        else:
            action = mapping.shift_midi_action if shift_active and mapping.shift_midi_action else mapping.midi_action
            self.send_midi(action)

    def send_hotkey(self, combo: str):
        combo = combo.strip().lower()
        if not combo:
            return

        parts = [self.parse_key_name(p.strip()) for p in combo.split("+") if p.strip()]
        if not parts:
            return

        pyautogui.hotkey(*parts)

    def parse_key_name(self, token: str) -> str:
        aliases = {
            "control": "ctrl",
            "cmd": "command",
            "win": "win",
            "option": "alt",
            "return": "enter",
            "escape": "esc",
            "del": "delete",
            "pageup": "pgup",
            "pagedown": "pgdn",
            "plus": "+",
            "minus": "-",
            "equal": "=",
            "backslash": "\\",
            "comma": ",",
            "period": ".",
            "num0": "num0",
            "num1": "num1",
            "num2": "num2",
            "num3": "num3",
            "num4": "num4",
            "num5": "num5",
            "num6": "num6",
            "num7": "num7",
            "num8": "num8",
            "num9": "num9",
            "num*": "multiply",
            "numpad*": "multiply",
        }
        return aliases.get(token, token)

    def send_midi(self, value: str):
        if mido is None:
            raise RuntimeError("mido is not installed")
        if self.midi_out is None:
            raise RuntimeError("No MIDI output selected")

        value = value.strip().lower()
        if not value:
            return

        if value == "start":
            self.midi_out.send(mido.Message("start"))
        elif value == "stop":
            self.midi_out.send(mido.Message("stop"))
        elif value == "continue":
            self.midi_out.send(mido.Message("continue"))
        elif value == "record":
            # MMC Record Strobe: F0 7F 7F 06 06 F7
            # Many DAWs can learn or respond to this; otherwise users can remap it.
            self.midi_out.send(mido.Message("sysex", data=[0x7F, 0x7F, 0x06, 0x06]))
        elif value == "pause":
            # MMC Pause: F0 7F 7F 06 09 F7
            self.midi_out.send(mido.Message("sysex", data=[0x7F, 0x7F, 0x06, 0x09]))
        elif value.startswith("note:"):
            parts = value.split(":")
            note = int(parts[1])
            velocity = int(parts[2]) if len(parts) > 2 else 100
            self.midi_out.send(mido.Message("note_on", note=note, velocity=velocity))
        elif value.startswith("cc:"):
            _, control, val = value.split(":")
            self.midi_out.send(mido.Message("control_change", control=int(control), value=int(val)))
        else:
            raise ValueError("MIDI Action must be start, stop, continue, record, pause, note:60, or cc:20:127")


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Eela Controller Bridge")
        self.resize(1180, 680)

        self.settings = load_settings()
        self.mappings: list[Mapping] = [Mapping(*item) for item in DEFAULT_MAPPINGS]
        self.runner = ActionRunner()
        self.worker_thread: Optional[QThread] = None
        self.worker: Optional[SerialWorker] = None
        self.learn_row: Optional[int] = None
        self.last_trigger: dict[str, float] = {}
        self.shift_active = False
        self.shift_push_pattern = bytes.fromhex("00 f8 80 f8")
        self.shift_release_pattern = bytes.fromhex("00 f8 00")

        self.port_combo = QComboBox()
        self.refresh_ports_btn = QPushButton("Refresh")
        self.baud_spin = QSpinBox()
        self.baud_spin.setRange(300, 115200)
        self.baud_spin.setValue(9600)
        self.start_btn = QPushButton("Start sniffing")
        self.stop_btn = QPushButton("Stop")
        self.stop_btn.setEnabled(False)

        self.midi_combo = QComboBox()
        self.refresh_midi_btn = QPushButton("Refresh MIDI")

        self.preset_combo = QComboBox()
        self.populate_preset_combo()
        self.set_saved_preset_selection()
        self.apply_preset_btn = QPushButton("Apply preset")
        self.load_last_preset_checkbox = QCheckBox("Load last selected preset on startup")
        self.load_last_preset_checkbox.setChecked(bool(self.settings.get("load_last_preset_on_startup", False)))
        self.boot_options_label = QLabel("Boot options")
        self.boot_options_label.setStyleSheet("font-weight: bold;")

        self.startup_checkbox = QCheckBox("Load software on startup")
        self.startup_checkbox.setChecked(is_startup_enabled())

        self.auto_sniff_checkbox = QCheckBox("Start sniffing automatically")
        self.auto_sniff_checkbox.setChecked(bool(self.settings.get("start_sniffing_automatically", False)))

        self.start_background_checkbox = QCheckBox("Start in background")
        self.start_background_checkbox.setChecked(bool(self.settings.get("start_in_background", sys.platform.startswith("win"))))

        self.use_last_serial_checkbox = QCheckBox("Use last used Serial COM port")
        self.use_last_serial_checkbox.setChecked(bool(self.settings.get("use_last_serial_port", True)))

        self.status_label = QLabel("Ready")
        self.last_hex_label = QLabel("Last HEX: -")
        self.last_decoded_label = QLabel("Decoded: -")

        self.table = QTableWidget(0, 8)
        self.table.setHorizontalHeaderLabels([
            "Enabled",
            "Name",
            "HEX pattern",
            "Action type",
            "Shortcut Action",
            "MIDI Action",
            "Shift Shortcut",
            "Shift MIDI",
        ])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.verticalHeader().setDefaultSectionSize(28)
        self.table.cellDoubleClicked.connect(self.capture_shortcut_for_cell)

        self.add_btn = QPushButton("Add mapping")
        self.remove_btn = QPushButton("Remove selected")
        self.learn_btn = QPushButton("Learn selected")
        self.save_btn = QPushButton("Save mappings")
        self.load_btn = QPushButton("Load mappings")
        self.github_btn = QPushButton("Help and Info")
        self.github_btn.setCursor(Qt.PointingHandCursor)
        self.github_btn.setStyleSheet('''
            QPushButton {
                background: transparent;
                border: none;
                color: #1a73e8;
                text-align: left;
                padding: 0px;
                font-size: 12px;
            }
            QPushButton:hover {
                text-decoration: underline;
            }
        ''')
        self.github_btn.setMaximumWidth(90)

        self.log_box = QLineEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setPlaceholderText("Latest raw packet appears here")

        self.toggle_log_btn = QPushButton("Show console log")
        self.clear_log_btn = QPushButton("Clear log")
        self.clear_log_btn.setEnabled(False)
        self.console_log = QPlainTextEdit()
        self.console_log.setReadOnly(True)
        self.console_log.setMaximumBlockCount(2000)
        self.console_log.setVisible(False)
        self.clear_log_btn.setVisible(False)

        self.build_ui()
        self.connect_signals()
        self.refresh_ports()
        self.refresh_midi_ports()
        self.populate_table()
        self.setup_tray()

        if self.settings.get("start_in_background", sys.platform.startswith("win")):
            self.hide()
        else:
            self.show()

        if self.settings.get("load_last_preset_on_startup", False):
            self.apply_selected_preset()

        if self.settings.get("start_sniffing_automatically", False):
            self.start_serial()

    def setup_tray(self):
        self.tray_icon = QSystemTrayIcon(self)
        icon = self.style().standardIcon(QStyle.StandardPixmap.SP_ComputerIcon)
        self.tray_icon.setIcon(icon)
        self.setWindowIcon(icon)

        tray_menu = QMenu()

        show_action = QAction("Show", self)
        hide_action = QAction("Hide", self)
        quit_action = QAction("Quit", self)

        show_action.triggered.connect(self.show_normal)
        hide_action.triggered.connect(self.hide)
        quit_action.triggered.connect(self.force_quit)

        tray_menu.addAction(show_action)
        tray_menu.addAction(hide_action)
        tray_menu.addSeparator()
        tray_menu.addAction(quit_action)

        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.activated.connect(self.on_tray_activated)
        self.tray_icon.show()

    def on_tray_activated(self, reason):
        if reason == QSystemTrayIcon.Trigger:
            if self.isVisible():
                self.hide()
            else:
                self.show_normal()

    def show_normal(self):
        self.show()
        self.raise_()
        self.activateWindow()

    def force_quit(self):
        self.tray_icon.hide()
        QApplication.quit()

    def build_ui(self):
        central = QWidget()
        root = QVBoxLayout(central)

        serial_group = QGroupBox("Serial input")
        serial_layout = QGridLayout(serial_group)
        serial_layout.addWidget(QLabel("Port"), 0, 0)
        serial_layout.addWidget(self.port_combo, 0, 1)
        serial_layout.addWidget(self.refresh_ports_btn, 0, 2)
        serial_layout.addWidget(QLabel("Baud"), 1, 0)
        serial_layout.addWidget(self.baud_spin, 1, 1)
        serial_layout.addWidget(self.start_btn, 2, 1)
        serial_layout.addWidget(self.stop_btn, 2, 2)

        midi_group = QGroupBox("MIDI output, optional")
        midi_layout = QGridLayout(midi_group)
        midi_layout.addWidget(QLabel("Output"), 0, 0)
        midi_layout.addWidget(self.midi_combo, 0, 1)
        midi_layout.addWidget(self.refresh_midi_btn, 0, 2)

        preset_group = QGroupBox("Keymap Presets")
        preset_layout = QGridLayout(preset_group)
        preset_layout.addWidget(QLabel("Select preset"), 0, 0)
        preset_layout.addWidget(self.preset_combo, 0, 1)
        preset_layout.addWidget(self.apply_preset_btn, 0, 2)
        preset_layout.addWidget(self.load_last_preset_checkbox, 1, 1, 1, 2)

        boot_group = QGroupBox("Boot options")
        boot_layout = QVBoxLayout(boot_group)
        boot_layout.addWidget(self.startup_checkbox)
        boot_layout.addWidget(self.auto_sniff_checkbox)
        boot_layout.addWidget(self.start_background_checkbox)
        boot_layout.addWidget(self.use_last_serial_checkbox)

        right_column = QVBoxLayout()
        right_column.addWidget(midi_group)
        right_column.addWidget(boot_group)

        top = QHBoxLayout()
        top.addWidget(serial_group)
        top.addLayout(right_column)
        root.addLayout(top)

        root.addWidget(self.status_label)
        root.addWidget(self.last_hex_label)
        root.addWidget(self.last_decoded_label)
        root.addWidget(self.log_box)

        log_buttons = QHBoxLayout()
        log_buttons.addWidget(self.toggle_log_btn)
        log_buttons.addWidget(self.clear_log_btn)
        log_buttons.addStretch(1)
        root.addLayout(log_buttons)
        root.addWidget(self.console_log)
        root.addWidget(preset_group)

        root.addWidget(self.table)

        buttons = QHBoxLayout()
        buttons.addWidget(self.add_btn)
        buttons.addWidget(self.remove_btn)
        buttons.addWidget(self.learn_btn)
        buttons.addStretch(1)
        buttons.addWidget(self.load_btn)
        buttons.addWidget(self.save_btn)
        root.addLayout(buttons)

        help_text = QLabel(
            "Shortcut examples: ctrl+shift+z, space, enter, left, right, delete | "
            "MIDI examples: start, stop, continue, note:60, note:60:100, cc:20:127"
        )
        help_text.setWordWrap(True)

        help_row = QHBoxLayout()
        help_row.addWidget(help_text, 1)
        help_row.addWidget(self.github_btn, 0, Qt.AlignRight)

        root.addLayout(help_row)

        self.setCentralWidget(central)

    def populate_preset_combo(self):
        self.preset_combo.clear()
        model = QStandardItemModel()

        categories = {
            "DAW": [],
            "NLE": [],
            "Playout": [],
        }

        for preset_name in DAW_PRESETS.keys():
            if preset_name.startswith("DAW /"):
                categories["DAW"].append(preset_name)
            elif preset_name.startswith("NLE /"):
                categories["NLE"].append(preset_name)
            elif preset_name.startswith("Playout /"):
                categories["Playout"].append(preset_name)

        for category, items in categories.items():
            header = QStandardItem(category)
            header.setEnabled(False)
            header.setSelectable(False)
            font = QFont()
            font.setBold(True)
            header.setFont(font)
            model.appendRow(header)

            for preset in items:
                display_name = preset.split("/", 1)[1].strip()
                item = QStandardItem(f"   {display_name}")
                item.setData(preset, Qt.UserRole)
                model.appendRow(item)

        self.preset_combo.setModel(model)

    def set_saved_preset_selection(self):
        saved = self.settings.get("last_preset", "DAW / Reaper")
        for i in range(self.preset_combo.count()):
            data = self.preset_combo.itemData(i, Qt.UserRole)
            if data == saved:
                self.preset_combo.setCurrentIndex(i)
                return

    def save_preset_options(self):
        self.settings["load_last_preset_on_startup"] = self.load_last_preset_checkbox.isChecked()
        save_settings(self.settings)

    def capture_shortcut_for_cell(self, row: int, col: int):
        if col not in (4, 6):
            return
        dialog = ShortcutCaptureDialog(self)
        if dialog.exec() == QDialog.Accepted and dialog.captured_shortcut:
            self.table.setItem(row, col, QTableWidgetItem(dialog.captured_shortcut))
            self.read_table()
            self.status_label.setText(f"Captured shortcut: {dialog.captured_shortcut}")

    def apply_selected_preset(self):
        preset_name = self.preset_combo.currentData(Qt.UserRole)

        # Fallback for older combo items or if Qt.UserRole data is missing.
        if not preset_name:
            display_name = self.preset_combo.currentText().strip()
            if not display_name or display_name in {"DAW", "NLE", "Playout"}:
                return
            for full_name in DAW_PRESETS.keys():
                short_name = full_name.split("/", 1)[-1].strip()
                if display_name == short_name or display_name == full_name:
                    preset_name = full_name
                    break

        if not preset_name:
            self.status_label.setText("Preset not found")
            return

        preset = DAW_PRESETS.get(preset_name)
        if not preset:
            self.status_label.setText(f"Preset not found: {preset_name}")
            return

        self.read_table()
        changed = False
        for mapping in self.mappings:
            if mapping.name in preset:
                shortcut, midi = preset[mapping.name]
                mapping.shortcut_action = shortcut
                mapping.midi_action = midi
                changed = True

        if changed:
            self.settings["last_preset"] = preset_name
            save_settings(self.settings)
            self.populate_table()
            self.status_label.setText(f"Loaded preset: {preset_name}")
            self.append_console_log("-", f"Loaded preset: {preset_name}")
        else:
            self.status_label.setText(f"Preset loaded but no matching controls: {preset_name}")

    def append_console_log(self, hex_text: str, decoded: str = "unknown"):
        timestamp = time.strftime("%H:%M:%S")
        self.console_log.appendPlainText(f"[{timestamp}] HEX: {hex_text} | {decoded}")
        self.clear_log_btn.setEnabled(True)

    def toggle_console_log(self):
        visible = not self.console_log.isVisible()
        self.console_log.setVisible(visible)
        self.clear_log_btn.setVisible(visible)
        self.toggle_log_btn.setText("Hide console log" if visible else "Show console log")

    def clear_console_log(self):
        self.console_log.clear()
        self.clear_log_btn.setEnabled(False)

    def connect_signals(self):
        self.refresh_ports_btn.clicked.connect(self.refresh_ports)
        self.refresh_midi_btn.clicked.connect(self.refresh_midi_ports)
        self.start_btn.clicked.connect(self.start_serial)
        self.stop_btn.clicked.connect(self.stop_serial)
        self.add_btn.clicked.connect(self.add_mapping)
        self.remove_btn.clicked.connect(self.remove_selected)
        self.learn_btn.clicked.connect(self.learn_selected)
        self.save_btn.clicked.connect(self.save_mappings)
        self.load_btn.clicked.connect(self.load_mappings)
        self.apply_preset_btn.clicked.connect(self.apply_selected_preset)
        self.load_last_preset_checkbox.stateChanged.connect(self.save_preset_options)
        self.toggle_log_btn.clicked.connect(self.toggle_console_log)
        self.clear_log_btn.clicked.connect(self.clear_console_log)
        self.github_btn.clicked.connect(self.open_github)
        self.midi_combo.currentTextChanged.connect(self.change_midi_port)
        self.port_combo.currentIndexChanged.connect(self.save_selected_ports)
        self.startup_checkbox.stateChanged.connect(self.toggle_startup)
        self.auto_sniff_checkbox.stateChanged.connect(self.save_boot_options)
        self.start_background_checkbox.stateChanged.connect(self.save_boot_options)
        self.use_last_serial_checkbox.stateChanged.connect(self.save_boot_options)

    def open_github(self):
        webbrowser.open("https://github.com/TripleDots/Logos")

    def refresh_ports(self):
        self.port_combo.clear()
        last_port = self.settings.get("last_serial_port", "") if self.settings.get("use_last_serial_port", True) else ""
        selected_index = -1
        for index, port in enumerate(serial.tools.list_ports.comports()):
            self.port_combo.addItem(f"{port.device} - {port.description}", port.device)
            if port.device == last_port:
                selected_index = index
        if selected_index >= 0:
            self.port_combo.setCurrentIndex(selected_index)

    def refresh_midi_ports(self):
        self.midi_combo.clear()
        self.midi_combo.addItem("No MIDI", "")
        last_midi = self.settings.get("last_midi_port", "")
        selected_index = 0
        if mido is not None:
            try:
                for index, name in enumerate(mido.get_output_names(), start=1):
                    self.midi_combo.addItem(name, name)
                    if name == last_midi:
                        selected_index = index
            except Exception as exc:
                self.status_label.setText(f"MIDI unavailable: {exc}")
        self.midi_combo.setCurrentIndex(selected_index)

    def change_midi_port(self):
        port = self.midi_combo.currentData() or ""
        self.save_selected_ports()
        try:
            self.runner.set_midi_port(port)
            if port:
                self.status_label.setText(f"MIDI output selected: {port}")
        except Exception as exc:
            QMessageBox.warning(self, "MIDI error", str(exc))

    def save_selected_ports(self):
        self.settings["last_serial_port"] = self.port_combo.currentData() or ""
        self.settings["last_midi_port"] = self.midi_combo.currentData() or ""
        save_settings(self.settings)

    def save_boot_options(self):
        self.settings["start_sniffing_automatically"] = self.auto_sniff_checkbox.isChecked()
        self.settings["start_in_background"] = self.start_background_checkbox.isChecked()
        self.settings["use_last_serial_port"] = self.use_last_serial_checkbox.isChecked()
        save_settings(self.settings)

    def toggle_startup(self):
        enabled = self.startup_checkbox.isChecked()
        try:
            set_startup_enabled(enabled)
            self.save_boot_options()
            self.status_label.setText("Load software on startup enabled" if enabled else "Load software on startup disabled")
        except Exception as exc:
            self.startup_checkbox.blockSignals(True)
            self.startup_checkbox.setChecked(is_startup_enabled())
            self.startup_checkbox.blockSignals(False)
            QMessageBox.warning(self, "Startup error", str(exc))

    def populate_table(self):
        self.table.setRowCount(len(self.mappings))
        for row, mapping in enumerate(self.mappings):
            enabled_combo = QComboBox()
            if mapping.name.upper() == "SHIFT":
                enabled_combo.addItems(["Yes", "As Shift Layer", "No"])
                if getattr(mapping, "shift_mode", "yes") == "as_shift_layer":
                    enabled_combo.setCurrentText("As Shift Layer")
                else:
                    enabled_combo.setCurrentText("Yes" if mapping.enabled else "No")
            else:
                enabled_combo.addItems(["Yes", "No"])
                enabled_combo.setCurrentText("Yes" if mapping.enabled else "No")
            enabled_combo.currentTextChanged.connect(lambda _=None: self.read_table())

            name_item = QTableWidgetItem(mapping.name)
            hex_item = QTableWidgetItem(mapping.hex_pattern)

            action_combo = QComboBox()
            action_combo.addItems([KEYBOARD_ACTION, MIDI_ACTION])
            action_combo.setCurrentText(normalize_action_type(mapping.action_type))
            action_combo.currentTextChanged.connect(lambda _=None: self.read_table())

            shortcut_item = QTableWidgetItem(mapping.shortcut_action)
            midi_item = QTableWidgetItem(mapping.midi_action)
            shift_shortcut_item = QTableWidgetItem(mapping.shift_shortcut_action)
            shift_midi_item = QTableWidgetItem(mapping.shift_midi_action)

            self.table.setCellWidget(row, 0, enabled_combo)
            self.table.setItem(row, 1, name_item)
            self.table.setItem(row, 2, hex_item)
            self.table.setCellWidget(row, 3, action_combo)
            self.table.setItem(row, 4, shortcut_item)
            self.table.setItem(row, 5, midi_item)
            self.table.setItem(row, 6, shift_shortcut_item)
            self.table.setItem(row, 7, shift_midi_item)

    def read_table(self):
        mappings = []
        for row in range(self.table.rowCount()):
            enabled_widget = self.table.cellWidget(row, 0)
            enabled_text = enabled_widget.currentText().strip().lower() if isinstance(enabled_widget, QComboBox) else "yes"
            name = self.item_text(row, 1).strip()
            hex_pattern = self.item_text(row, 2).strip()
            combo = self.table.cellWidget(row, 3)
            action_type = normalize_action_type(combo.currentText()) if isinstance(combo, QComboBox) else KEYBOARD_ACTION

            if name and hex_pattern:
                mappings.append(
                    Mapping(
                        name=name,
                        hex_pattern=hex_pattern,
                        action_type=action_type,
                        shortcut_action=self.item_text(row, 4).strip(),
                        midi_action=self.item_text(row, 5).strip(),
                        shift_shortcut_action=self.item_text(row, 6).strip(),
                        shift_midi_action=self.item_text(row, 7).strip(),
                        enabled=(enabled_text in ("yes", "as shift layer")),
                        shift_mode=("as_shift_layer" if enabled_text == "as shift layer" else enabled_text.replace(" ", "_")),
                    )
                )
        self.mappings = mappings

    def item_text(self, row: int, col: int) -> str:
        item = self.table.item(row, col)
        return item.text() if item else ""

    def start_serial(self):
        port = self.port_combo.currentData()
        if not port:
            QMessageBox.warning(self, "No port", "Select a serial port first.")
            return

        self.read_table()
        self.worker_thread = QThread()
        self.worker = SerialWorker(port, self.baud_spin.value())
        self.worker.moveToThread(self.worker_thread)
        self.worker_thread.started.connect(self.worker.run)
        self.worker.packet.connect(self.handle_packet)
        self.worker.status.connect(self.status_label.setText)
        self.worker.error.connect(self.show_error)
        self.worker_thread.start()

        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)

    def stop_serial(self):
        if self.worker:
            self.worker.stop()
        if self.worker_thread:
            self.worker_thread.quit()
            self.worker_thread.wait(1000)
        self.worker = None
        self.worker_thread = None
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)

    @Slot(str)
    def show_error(self, text: str):
        self.status_label.setText(text)
        QMessageBox.warning(self, "Error", text)

    @Slot(bytes)
    def handle_packet(self, data: bytes):
        hex_text = normalize_hex(data)
        self.last_hex_label.setText(f"Last HEX: {hex_text}")
        self.log_box.setText(hex_text)

        self.read_table()
        shift_mapping = next((m for m in self.mappings if m.name.upper() == "SHIFT"), None)
        shift_is_layer = bool(shift_mapping and getattr(shift_mapping, "shift_mode", "yes") == "as_shift_layer")

        if shift_is_layer and data == self.shift_push_pattern:
            self.shift_active = True
            self.last_decoded_label.setText("Decoded: SHIFT layer active")
            self.append_console_log(hex_text, "SHIFT layer active")
            return

        if shift_is_layer and data == self.shift_release_pattern:
            self.shift_active = False
            self.last_decoded_label.setText("Decoded: SHIFT layer inactive")
            self.append_console_log(hex_text, "SHIFT layer inactive")
            return

        if self.learn_row is not None:
            self.table.setItem(self.learn_row, 2, QTableWidgetItem(hex_text))
            self.status_label.setText(f"Learned HEX for row {self.learn_row + 1}: {hex_text}")
            self.append_console_log(hex_text, f"learned row {self.learn_row + 1}")
            self.learn_row = None
            self.read_table()
            return

        matched = self.find_match(data)
        if not matched:
            self.last_decoded_label.setText("Decoded: unknown")
            self.append_console_log(hex_text, "unknown")
            return

        if matched.name.upper() == "SHIFT" and getattr(matched, "shift_mode", "yes") == "as_shift_layer":
            return

        active_value = self.get_active_action_value(matched)
        layer_name = "SHIFT" if self.shift_active else "NORMAL"
        decoded_text = f"{matched.name} [{layer_name}] -> {matched.action_type}: {active_value}"
        self.last_decoded_label.setText(f"Decoded: {decoded_text}")
        self.append_console_log(hex_text, decoded_text)

        key = matched.hex_pattern + (":shift" if self.shift_active else ":normal")
        now = time.time()
        if now - self.last_trigger.get(key, 0) < 0.10:
            return
        self.last_trigger[key] = now

        try:
            self.runner.run(matched, self.shift_active)
        except Exception as exc:
            self.status_label.setText(f"Action error: {exc}")

    def get_active_action_value(self, mapping: Mapping) -> str:
        if normalize_action_type(mapping.action_type) == KEYBOARD_ACTION:
            return mapping.shift_shortcut_action if self.shift_active and mapping.shift_shortcut_action else mapping.shortcut_action
        return mapping.shift_midi_action if self.shift_active and mapping.shift_midi_action else mapping.midi_action

    def find_match(self, data: bytes) -> Optional[Mapping]:
        for mapping in self.mappings:
            try:
                if mapping.enabled and data == mapping.pattern_bytes:
                    return mapping
            except ValueError:
                continue

        for mapping in self.mappings:
            try:
                pattern = mapping.pattern_bytes
                if mapping.enabled and pattern and pattern in data:
                    return mapping
            except ValueError:
                continue
        return None

    def add_mapping(self):
        self.read_table()
        self.mappings.append(Mapping("New Button", "", KEYBOARD_ACTION, "", "", "", ""))
        self.populate_table()

    def remove_selected(self):
        rows = sorted({idx.row() for idx in self.table.selectedIndexes()}, reverse=True)
        if not rows:
            return
        self.read_table()
        for row in rows:
            if 0 <= row < len(self.mappings):
                self.mappings.pop(row)
        self.populate_table()

    def learn_selected(self):
        row = self.table.currentRow()
        if row < 0:
            QMessageBox.information(self, "Learn", "Select a row first.")
            return
        self.learn_row = row
        self.status_label.setText("Learning: press a button on the Eela now...")

    def save_mappings(self):
        self.read_table()
        path, _ = QFileDialog.getSaveFileName(self, "Save mappings", "eela_mappings.json", "JSON files (*.json)")
        if not path:
            return
        with open(path, "w", encoding="utf-8") as f:
            json.dump([asdict(m) for m in self.mappings], f, indent=2)
        self.status_label.setText(f"Saved mappings to {Path(path).name}")

    def load_mappings(self):
        path, _ = QFileDialog.getOpenFileName(self, "Load mappings", "", "JSON files (*.json)")
        if not path:
            return
        with open(path, "r", encoding="utf-8") as f:
            raw = json.load(f)
        self.mappings = [Mapping.from_dict(item) for item in raw]
        self.populate_table()
        self.status_label.setText(f"Loaded mappings from {Path(path).name}")

    def closeEvent(self, event):
        if sys.platform.startswith("win"):
            event.ignore()
            self.hide()
            return

        self.save_selected_ports()
        self.stop_serial()
        super().closeEvent(event)


def ensure_qt_conf_for_windows():
    if not sys.platform.startswith("win"):
        return

    qt_conf_path = Path(__file__).resolve().parent / "qt.conf"
    qt_conf_text = "[Platforms]\nWindowsArguments = dpiawareness=0\n"

    try:
        existing_text = ""
        if qt_conf_path.exists():
            existing_text = qt_conf_path.read_text(encoding="utf-8", errors="ignore")
        if "WindowsArguments" not in existing_text:
            qt_conf_path.write_text(qt_conf_text, encoding="utf-8")
    except Exception:
        # If qt.conf cannot be written, the command-line platform argument still helps.
        pass


def main():
    ensure_qt_conf_for_windows()

    qt_args = list(sys.argv)
    if sys.platform.startswith("win") and "-platform" not in qt_args:
        qt_args += ["-platform", "windows:dpiawareness=0"]

    app = QApplication(qt_args)
    window = MainWindow()
    # MainWindow decides whether to show or start hidden based on settings.
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
