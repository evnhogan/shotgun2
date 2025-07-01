"""Automate the Windows OOBE bypass shortcut.

This script presses ``Shift+F10`` to open a command prompt and then types the
command ``OOBE/bypassnro``. It exits early if executed on a non-Windows system
or if ``pyautogui`` cannot be imported (such as when no graphical environment
is available).
"""

import argparse
import sys
import time
import logging
import platform

try:
    import pyautogui
except Exception as exc:  # pragma: no cover - environment specific
    print(
        f'Failed to import pyautogui: {exc}. '
        'Please ensure the library is installed and '
        'a graphical environment is available.'
    )
    sys.exit(1)

pyautogui.FAILSAFE = False
pyautogui.PAUSE = 0.1


def open_command_prompt(delay: float = 2.0) -> None:
    """Open a command prompt via the Shift+F10 hotkey."""
    pyautogui.hotkey('shift', 'f10')
    time.sleep(delay)


def type_command(command: str) -> None:
    """Type a command and press Enter."""
    pyautogui.typewrite(command)
    pyautogui.press('enter')


logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[logging.StreamHandler(sys.stdout)]
)
logger = logging.getLogger(__name__)


def main(argv: list[str] | None = None) -> None:
    """Entry point for the script."""
    argv = argv if argv is not None else sys.argv[1:]

    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        '--delay',
        type=float,
        default=2.0,
        help='Delay in seconds after opening the command prompt',
    )
    parser.add_argument(
        'command',
        nargs='?',
        default='OOBE/bypassnro',
        help='Command to execute after the command prompt opens',
    )
    args = parser.parse_args(argv)

    if platform.system() != 'Windows':
        logger.error('This script is intended to run on Windows.')
        return

    logger.info('Opening command prompt with Shift+F10')
    try:
        open_command_prompt(args.delay)
        logger.info('Typing command: %s', args.command)
        type_command(args.command)
    except Exception as exc:
        logger.exception('Automation failed: %s', exc)


if __name__ == '__main__':
    main()
