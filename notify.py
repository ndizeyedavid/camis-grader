"""
Native Windows toast notifications + taskbar flashing, with no extra packages.

How it works
  * register() writes HKCU\\Software\\Classes\\AppUserModelId\\<APP_ID> with a
    display name and icon. That is Microsoft's documented way for an unpackaged
    desktop app to own its toasts, so they show "CAMIS Marks" and its icon.
  * toast() asks the built-in Windows PowerShell (5.1, which can reach the WinRT
    toast API) to show the notification. It runs hidden, in the background,
    and any failure is only logged. A notification must never break marking.
  * flash() makes the app's taskbar button blink until the teacher clicks it.
"""

from __future__ import annotations

import base64
import ctypes
import os
import subprocess
import sys
import threading
from pathlib import Path
from typing import Optional
from xml.sax.saxutils import escape

APP_ID = "NESA.CAMISMarks"
_icon: Optional[Path] = None
_enabled = sys.platform == "win32"

POWERSHELL = Path(os.environ.get("SystemRoot", r"C:\Windows")) / "System32/WindowsPowerShell/v1.0/powershell.exe"

_SCRIPT = r"""
$ErrorActionPreference = 'Stop'
[Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] > $null
[Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime] > $null
$xml = New-Object Windows.Data.Xml.Dom.XmlDocument
$xml.LoadXml([Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('__XML__')))
$toast = [Windows.UI.Notifications.ToastNotification]::new($xml)
$toast.Tag = 'camis'
$toast.Group = 'camis'
[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier('__APPID__').Show($toast)
"""


def register(icon: Optional[Path]) -> None:
    """Give our toasts a name and icon, and tie this process to that identity."""
    global _icon, _enabled
    if sys.platform != "win32":
        _enabled = False
        return
    _icon = icon if icon and icon.exists() else None
    try:
        import winreg

        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, rf"Software\Classes\AppUserModelId\{APP_ID}") as key:
            winreg.SetValueEx(key, "DisplayName", 0, winreg.REG_SZ, "CAMIS Marks")
            if _icon:
                winreg.SetValueEx(key, "IconUri", 0, winreg.REG_SZ, str(_icon))
                winreg.SetValueEx(key, "IconBackgroundColor", 0, winreg.REG_SZ, "FF0A0A0A")
    except OSError as exc:
        print(f"notify: could not register app id: {exc}")
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_ID)
    except Exception as exc:
        print(f"notify: could not set app id: {exc}")


def set_enabled(value: bool) -> None:
    global _enabled
    _enabled = bool(value) and sys.platform == "win32"


def _toast_xml(title: str, body: str, kind: str) -> str:
    logo = (f'<image placement="appLogoOverride" src="{escape(_icon.as_uri())}"/>' if _icon else "")
    # Problems stay on screen until dismissed; good news goes away by itself.
    scenario = ' scenario="reminder"' if kind == "error" else ""
    actions = '<actions><action content="OK" arguments="dismiss" activationType="system"/></actions>' if kind == "error" else ""
    audio = '<audio src="ms-winsoundevent:Notification.Default"/>'
    return (
        f'<toast{scenario} duration="{"long" if kind != "info" else "short"}">'
        f'<visual><binding template="ToastGeneric">'
        f"<text>{escape(title)}</text><text>{escape(body)}</text>{logo}"
        f"</binding></visual>{actions}{audio}</toast>"
    )


def toast(title: str, body: str, kind: str = "info") -> None:
    """Show a Windows notification. kind: info | success | warning | error."""
    if not _enabled:
        return
    xml = _toast_xml(title, body, kind)
    script = (_SCRIPT.replace("__XML__", base64.b64encode(xml.encode("utf-8")).decode("ascii"))
              .replace("__APPID__", APP_ID))
    encoded = base64.b64encode(script.encode("utf-16-le")).decode("ascii")

    def run() -> None:
        try:
            result = subprocess.run(
                [str(POWERSHELL), "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
                 "-WindowStyle", "Hidden", "-EncodedCommand", encoded],
                capture_output=True, timeout=30,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
            if result.returncode != 0:
                print(f"notify: powershell failed: {result.stderr.decode(errors='replace')[:400]}")
        except Exception as exc:
            print(f"notify: {exc}")

    threading.Thread(target=run, name="toast", daemon=True).start()


class _FLASHWINFO(ctypes.Structure):
    _fields_ = [("cbSize", ctypes.c_uint), ("hwnd", ctypes.c_void_p), ("dwFlags", ctypes.c_uint),
                ("uCount", ctypes.c_uint), ("dwTimeout", ctypes.c_uint)]


def find_window(title: str) -> int:
    if sys.platform != "win32":
        return 0
    try:
        return int(ctypes.windll.user32.FindWindowW(None, title) or 0)
    except Exception:
        return 0


def is_foreground(hwnd: int) -> bool:
    try:
        return bool(hwnd) and int(ctypes.windll.user32.GetForegroundWindow() or 0) == hwnd
    except Exception:
        return False


def flash(hwnd: int) -> None:
    """Blink the taskbar button until the window is brought to the front."""
    if not hwnd or sys.platform != "win32":
        return
    try:
        FLASHW_ALL, FLASHW_TIMERNOFG = 0x3, 0xC
        info = _FLASHWINFO(ctypes.sizeof(_FLASHWINFO), hwnd, FLASHW_ALL | FLASHW_TIMERNOFG, 0, 0)
        ctypes.windll.user32.FlashWindowEx(ctypes.byref(info))
    except Exception as exc:
        print(f"notify: flash failed: {exc}")
