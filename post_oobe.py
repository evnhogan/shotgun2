"""Automate Windows OOBE setup screens with no third-party dependencies.

This script runs after ``OOBE.py`` to finish the out-of-box experience.
It selects English and US keyboard layout, skips network connection,
and stops so the user can enter a computer name like ``CS-123``.
After the name is entered it leaves the password blank and continues
until Windows boots to the desktop.

The automation uses ``ctypes`` to simulate key presses via the
``SendInput`` API. It requires running on Windows with a visible OOBE
window in the foreground.
"""

from __future__ import annotations

import argparse
import ctypes
import logging
import platform
import sys
import time

IS_WINDOWS = platform.system() == "Windows"
if IS_WINDOWS:
    USER32 = ctypes.windll.user32
else:
    USER32 = None
INPUT_KEYBOARD = 1
KEYEVENTF_KEYUP = 0x0002

# virtual-key codes used in the script
VK_RETURN = 0x0D
VK_TAB = 0x09
VK_SPACE = 0x20
VK_SHIFT = 0x10

# mapping for simple alphanumeric characters
CHAR_TO_VK = {
    **{chr(c): c for c in range(ord("0"), ord("9") + 1)},
    **{chr(c): c for c in range(ord("A"), ord("Z") + 1)},
    "-": 0xBD,
    "_": 0xBD,
}

logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler("post_oobe.log", mode="a", encoding="utf-8"),
    ],
)
LOGGER = logging.getLogger(__name__)


def ensure_oobe_window_foreground() -> None:
    """Log a warning if the active window does not appear to be OOBE."""
    if USER32 is None:
        return
    hwnd = USER32.GetForegroundWindow()
    buf = ctypes.create_unicode_buffer(256)
    USER32.GetWindowTextW(hwnd, buf, 256)
    title = buf.value.lower()
    if "out of box experience" not in title:
        LOGGER.warning("OOBE window not in foreground: %s", buf.value)


class KEYBDINPUT(ctypes.Structure):  # pylint: disable=too-few-public-methods
    """Structure passed to ``SendInput`` for keyboard events."""

    _fields_ = [
        ("wVk", ctypes.c_ushort),
        ("wScan", ctypes.c_ushort),
        ("dwFlags", ctypes.c_ulong),
        ("time", ctypes.c_ulong),
        ("dwExtraInfo", ctypes.c_void_p),
    ]


class INPUT(ctypes.Structure):  # pylint: disable=too-few-public-methods
    """General input structure for ``SendInput``."""

    _fields_ = [
        ("type", ctypes.c_ulong),
        ("ki", KEYBDINPUT),
    ]


def _send_vk(vk: int, flags: int = 0) -> None:
    """Send a single keyboard event using ``SendInput``."""
    if USER32 is None:
        return
    inp = INPUT(INPUT_KEYBOARD, KEYBDINPUT(vk, 0, flags, 0, None))
    sent = USER32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(inp))
    if sent != 1:
        LOGGER.warning("SendInput failed for vk=%s flags=%s", vk, flags)
    time.sleep(0.05)


def press_key(vk: int, *, shift: bool = False) -> None:
    """Press and release the given virtual-key code."""
    if shift:
        _send_vk(VK_SHIFT)
    _send_vk(vk, 0)
    _send_vk(vk, KEYEVENTF_KEYUP)
    if shift:
        _send_vk(VK_SHIFT, KEYEVENTF_KEYUP)


def type_text(text: str) -> None:
    """Type a string using virtual-key codes."""
    shift_chars = {"_"}
    for ch in text:
        upper = ch.upper()
        vk = CHAR_TO_VK.get(upper)
        if vk is None:
            LOGGER.warning("Skipping unsupported character: %r", ch)
            continue
        need_shift = (ch.isalpha() and ch.isupper()) or ch in shift_chars
        press_key(vk, shift=need_shift)
        time.sleep(0.05)


def oobe_flow(device_name: str) -> None:
    """Complete the Windows setup screens using keyboard automation."""
    if not IS_WINDOWS:
        LOGGER.error("This script is intended for Windows only.")
        return

    ensure_oobe_window_foreground()

    LOGGER.info("Starting OOBE automation sequence")
    # region selection (default English)
    press_key(VK_RETURN)
    time.sleep(5)
    # keyboard layout (default US)
    press_key(VK_RETURN)
    time.sleep(2)
    # skip second keyboard layout
    press_key(VK_TAB)
    press_key(VK_RETURN)
    time.sleep(2)
    # choose "I don't have internet"
    press_key(VK_TAB)
    press_key(VK_TAB)
    press_key(VK_RETURN)
    time.sleep(2)
    # continue with limited setup
    press_key(VK_TAB)
    press_key(VK_RETURN)
    time.sleep(5)
    # stop on device name screen
    LOGGER.info("Waiting for device name input")
    type_text(device_name)
    press_key(VK_RETURN)
    time.sleep(2)
    # skip password creation
    press_key(VK_RETURN)
    press_key(VK_RETURN)
    # accept remaining screens (privacy settings, etc.)
    for _ in range(6):
        time.sleep(2)
        press_key(VK_RETURN)
    LOGGER.info("Automation complete, awaiting desktop")
    time.sleep(30)


def main(argv: list[str] | None = None) -> None:
    """Entry point for command line execution."""
    parser = argparse.ArgumentParser(description="Automate OOBE after bypass")
    parser.add_argument("cs_number", help="Numeric identifier for computer")
    args = parser.parse_args(argv)
    if not args.cs_number.isdigit():
        parser.error("cs_number must be numeric")
    if not IS_WINDOWS:
        parser.error("This script only runs on Windows")
    oobe_flow(f"CS-{args.cs_number}")


if __name__ == "__main__":
    main()
