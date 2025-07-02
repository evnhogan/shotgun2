"""Automate the Windows OOBE bypass shortcut.

This script presses ``Shift+F10`` to open a command prompt and then types the
command ``OOBE/bypassnro``. It exits early if executed on a non-Windows
system or if ``pyautogui`` cannot be imported (such as when no graphical
environment is available). Use ``--dry-run`` to log actions without sending
keyboard input.
"""

import argparse
import sys
import time
import logging
import platform
import os

try:
    import pyautogui
except ImportError as exc:  # pragma: no cover - environment specific
    msg = [f"Failed to import pyautogui: {exc}"]
    if os.name != "nt" and "DISPLAY" not in os.environ:
        msg.append("DISPLAY environment variable is not set")
    msg.append(
        "Please ensure the library is installed and a graphical environment is available."
    )
    print(". ".join(msg))
    sys.exit(1)

pyautogui.FAILSAFE = False
pyautogui.PAUSE = 0.1


def open_command_prompt(delay: float = 2.0, *, dry_run: bool = False) -> None:
    """Open a command prompt via the Shift+F10 hotkey."""
    if dry_run:
        logger.info("[DRY RUN] Would press Shift+F10")
    else:
        pyautogui.hotkey("shift", "f10")
    time.sleep(delay)


def type_command(command: str, *, dry_run: bool = False) -> None:
    """Type a command and press Enter."""
    if dry_run:
        logger.info("[DRY RUN] Would type: %s", command)
    else:
        pyautogui.typewrite(command)
        pyautogui.press("enter")


logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
logger = logging.getLogger(__name__)


def main(argv: list[str] | None = None) -> None:
    """Entry point for the script."""
    argv = argv if argv is not None else sys.argv[1:]

    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--delay",
        type=float,
        default=2.0,
        help="Delay in seconds after opening the command prompt",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Log actions without sending any keyboard input",
    )
    parser.add_argument(
        "command",
        nargs="?",
        default="OOBE/bypassnro",
        help="Command to execute after the command prompt opens",
    )
    args = parser.parse_args(argv)

    if platform.system() != "Windows":
        logger.error("This script is intended to run on Windows.")
        sys.exit(1)

    logger.info("Opening command prompt with Shift+F10")
    try:
        open_command_prompt(args.delay, dry_run=args.dry_run)
        logger.info("Typing command: %s", args.command)
        type_command(args.command, dry_run=args.dry_run)
    except (pyautogui.FailSafeException, pyautogui.PyAutoGUIException) as exc:
        logger.exception("Automation failed: %s", exc)


if __name__ == "__main__":
    main()
