import sys
import json
import logging
import subprocess
import ctypes
import platform
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent

try:
    from tqdm import tqdm
except ImportError:
    tqdm = None

STATE_FILE = BASE_DIR / 'installer_state.json'
LOG_FILE = BASE_DIR / 'installer.log'
INSTALL_DIR = BASE_DIR / 'installers'
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


def is_admin():
    if platform.system() != 'Windows':
        return False
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception as e:
        logger.error('Admin check failed: %s', e)
        return False


def run_as_admin():
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


def load_state():
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
    return {'step': 0, 'completed_files': []}


def save_state(state):
    with STATE_FILE.open('w', encoding='utf-8') as f:
        json.dump(state, f)

# --- Restart handling ---


def create_resume_task():
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


def remove_resume_task():
    cmd = ['schtasks', '/Delete', '/F', '/TN', SCHEDULED_TASK_NAME]
    subprocess.run(cmd, check=False)


def reboot_system():
    create_resume_task()
    logger.info('Rebooting system...')
    subprocess.run(['shutdown', '/r', '/t', '5', '/f'], check=True)


def is_reboot_pending():
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


def install_windows_updates():
    logger.info('Installing Windows updates...')
    powershell_cmd = (
        'Install-Module -Name PSWindowsUpdate -Force; '
        'Import-Module PSWindowsUpdate; '
        'Add-WUServiceManager -MicrosoftUpdate; '
        'Get-WindowsUpdate -Install -AcceptAll -MicrosoftUpdate -IgnoreReboot'
    )
    try:
        subprocess.run(
            [
                'powershell',
                '-NoProfile',
                '-ExecutionPolicy',
                'Bypass',
                '-Command',
                powershell_cmd,
            ],
            check=True,
        )
    except Exception as e:
        logger.error('Windows update failed: %s', e)

# --- Dell Updates ---


def install_dell_updates():
    logger.info('Installing Dell Command Update updates...')
    dcu = Path('C:/Program Files/Dell/CommandUpdate/dcu-cli.exe')
    if not dcu.exists():
        logger.warning('Dell Command Update not found at %s', dcu)
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


def _run_installer_file(path: Path):
    """Execute an installer file and block until completion."""
    cmd = [str(path)]
    if path.suffix.lower() == '.msi':
        cmd = ['msiexec', '/i', str(path), '/qn', '/norestart']
    subprocess.run(cmd, check=True)


def run_installers(state):
    if not INSTALL_DIR.exists():
        logger.warning('Installer directory %s not found', INSTALL_DIR)
        return
    files = [f for f in sorted(INSTALL_DIR.glob('*')) if f.is_file()]
    total = len(files)
    bar = None
    if tqdm:
        bar = tqdm(
            total=total,
            initial=len(state['completed_files']),
            unit='file',
        )
    for f in files:
        if str(f) in state['completed_files']:
            continue
        try:
            logger.info('Running installer %s', f)
            _run_installer_file(f)
            state['completed_files'].append(str(f))
            save_state(state)
        except Exception as e:
            logger.error('Installer %s failed: %s', f, e)
        if bar:
            bar.update(1)
    if bar:
        bar.close()


def main():
    if platform.system() != 'Windows':
        logger.error('This installer only runs on Windows')
        return
    if not is_admin():
        run_as_admin()
    state = load_state()
    steps = [install_windows_updates, install_dell_updates, run_installers]
    for idx, func in enumerate(steps, start=1):
        if state['step'] >= idx:
            continue
        func(state) if func == run_installers else func()
        state['step'] = idx
        save_state(state)
        if is_reboot_pending():
            save_state(state)
            reboot_system()
            return
    save_state(state)
    remove_resume_task()
    logger.info('Installation completed successfully')


if __name__ == '__main__':
    try:
        main()
    except Exception as exc:
        logger.exception('Fatal error: %s', exc)
        remove_resume_task()
        sys.exit(1)
