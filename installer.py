"""Automated installer runner for Windows machines."""

import sys
import json
import logging
import subprocess
import ctypes
import platform
import os
import urllib.request
import re
import tempfile
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent



STATE_FILE = BASE_DIR / 'installer_state.json'
LOG_FILE = BASE_DIR / 'installer.log'
SCHEDULED_TASK_NAME = 'ShotgunInstallerResume'



logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE, mode='a', encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)


def is_admin() -> bool:
    """Return ``True`` if running with administrator privileges."""

    if platform.system() != 'Windows':
        return False
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception as e:
        logger.error('Admin check failed: %s', e)
        return False


def run_as_admin() -> None:
    """Re-launch this script with administrator rights and exit."""
    logger.info('Re-launching script with administrator privileges')
    script = Path(__file__).resolve()
    params = ' '.join([f'"{script}"'] + [f'"{p}"' for p in sys.argv[1:]])
    ctypes.windll.shell32.ShellExecuteW(
        None,
        'runas',
        sys.executable,
        params,
        str(BASE_DIR),
        1,
    )
    sys.exit(0)


def load_state() -> dict:
    """Return persisted install state or an empty default."""
    if STATE_FILE.exists():
        try:
            with STATE_FILE.open('r', encoding='utf-8') as f:
                return json.load(f)
        except json.JSONDecodeError:
            logger.warning('State file corrupt, removing and starting fresh')
            try:
                STATE_FILE.unlink()
            except OSError as e:
                logger.error('Failed to delete corrupt state file: %s', e)
    return {'step': 0}


def save_state(state: dict) -> None:
    """Persist installer progress to the state file."""
    with STATE_FILE.open('w', encoding='utf-8') as f:
        json.dump(state, f)

# --- Restart handling ---


def create_resume_task() -> None:
    """Create a scheduled task to resume this installer after reboot."""
    cmd = [
        'schtasks', '/Create', '/F', '/TN', SCHEDULED_TASK_NAME,
        '/SC', 'ONSTART', '/RL', 'HIGHEST', '/RU', 'SYSTEM',
        '/TR',
        ' '.join(
            f'"{p}"' for p in [sys.executable, Path(__file__).resolve()]
        ),
    ]
    try:
        subprocess.run(cmd, check=True)
        logger.info('Scheduled task for resume created')
    except Exception as e:
        logger.error('Failed to create scheduled task: %s', e)


def remove_resume_task() -> None:
    """Delete the scheduled resume task if it exists."""
    cmd = ['schtasks', '/Delete', '/F', '/TN', SCHEDULED_TASK_NAME]
    subprocess.run(cmd, check=False)


def reboot_system() -> None:
    """Reboot the machine and schedule resume of the installer."""
    create_resume_task()
    logger.info('Rebooting system...')
    subprocess.run(['shutdown', '/r', '/t', '5', '/f'], check=True)


def is_reboot_pending() -> bool:
    """Return ``True`` if Windows signals that a reboot is pending."""
    reboot_keys = [
        (
            r'SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate'
            r'\Auto Update\RebootRequired'
        ),
        (
            r'SYSTEM\CurrentControlSet\Control\Session Manager'
            r'\PendingFileRenameOperations'
        ),
    ]
    try:
        import winreg
        for subkey in reboot_keys:
            try:
                winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, subkey)
                return True
            except FileNotFoundError:
                continue
    except Exception:
        logger.debug('winreg not available for reboot check')
    return False

# --- Windows Updates ---


def install_windows_updates() -> None:
    """Install all available Windows updates via PowerShell."""
    logger.info('Installing Windows updates...')
    powershell_cmd = (
        'Install-Module -Name PSWindowsUpdate -Force; '
        'Import-Module PSWindowsUpdate; '
        'Add-WUServiceManager -MicrosoftUpdate; '
        'Get-WindowsUpdate -Install -AcceptAll -MicrosoftUpdate -IgnoreReboot'
    )
    try:
        result = subprocess.run(
            [
                'powershell',
                '-NoProfile',
                '-ExecutionPolicy',
                'Bypass',
                '-Command',
                powershell_cmd,
            ],
            check=False,
            capture_output=True,
            text=True,
        )
        if result.returncode == 0:
            logger.info('Windows updates installed successfully')
        else:
            logger.error(
                'Windows update command failed with code %s', result.returncode
            )
            if result.stdout:
                logger.error('stdout: %s', result.stdout.strip())
            if result.stderr:
                logger.error('stderr: %s', result.stderr.strip())
    except Exception as e:
        logger.error('Windows update invocation failed: %s', e)

# --- Dell Updates ---


def download_latest_dcu() -> Path | None:
    """Download the latest Dell Command Update installer and return its path."""
    url_page = (
        'https://www.dell.com/support/kbdoc/en-us/000183146/dell-command-update'
    )
    try:
        with urllib.request.urlopen(url_page) as resp:
            html = resp.read().decode('utf-8', errors='ignore')
        match = re.search(
            r'https://dl\.dell\.com/[^"\s]*Command-Update[^"\s]*\.exe', html
        )
        if not match:
            logger.error('Could not find Dell Command Update download link')
            return None
        url = match.group(0)
        fd, tmp_path = tempfile.mkstemp(suffix='.exe')
        os.close(fd)
        logger.info('Downloading Dell Command Update from %s', url)
        urllib.request.urlretrieve(url, tmp_path)
        return Path(tmp_path)
    except Exception as exc:
        logger.error('Failed to download Dell Command Update: %s', exc)
        return None


def install_dell_updates() -> None:
    """Install firmware and driver updates via Dell Command Update."""
    logger.info('Installing Dell Command Update updates...')
    candidates = [
        Path('C:/Program Files/Dell/CommandUpdate/dcu-cli.exe'),
        Path('C:/Program Files (x86)/Dell/CommandUpdate/dcu-cli.exe'),
    ]
    dcu = next((p for p in candidates if p.exists()), None)
    if not dcu:
        logger.warning('Dell Command Update executable not found')
        installer = download_latest_dcu()
        if not installer:
            logger.error('Unable to download Dell Command Update')
            return
        try:
            subprocess.run([str(installer), '/s'], check=True)
        except Exception as e:
            logger.error('Failed to install Dell Command Update: %s', e)
            return
        dcu = next((p for p in candidates if p.exists()), None)
        if not dcu:
            logger.error('Dell Command Update installation did not create CLI')
            return
    try:
        result = subprocess.run(
            [str(dcu), '/applyUpdates', '/silent'],
            check=False,
            capture_output=True,
            text=True,
        )
        if result.returncode == 0:
            logger.info('Dell Command Update completed successfully')
        elif result.returncode == 2:
            logger.info('No Dell updates available')
        elif result.returncode == 102:
            logger.info('Dell Command Update requires a reboot to continue')
        else:
            logger.error(
                'Dell Command Update failed with exit code %s: %s',
                result.returncode,
                result.stdout.strip() or result.stderr.strip(),
            )
    except Exception as e:
        logger.error('Dell Command Update invocation failed: %s', e)

# --- Install files ---




def main() -> None:
    """Run update and installer steps, handling reboots as needed."""
    if platform.system() != 'Windows':
        logger.error('This installer only runs on Windows')
        return
    if not is_admin():
        run_as_admin()
    state = load_state()
    steps = [install_windows_updates, install_dell_updates]
    for idx, func in enumerate(steps, start=1):
        if state['step'] >= idx:
            continue
        func()
        state['step'] = idx
        save_state(state)
        if is_reboot_pending():
            save_state(state)
            reboot_system()
            return
    save_state(state)
    remove_resume_task()
    try:
        ctypes.windll.user32.MessageBoxW(
            0,
            'All Windows & Dell Command Updates Finished! Moving to the next phase...',
            'Updates Complete',
            0x40,
        )
    except Exception as e:
        logger.error('Failed to display completion message: %s', e)
    logger.info('Installation completed successfully')


if __name__ == '__main__':
    try:
        main()
    except Exception as exc:
        logger.exception('Fatal error: %s', exc)
        remove_resume_task()
        sys.exit(1)
