# shotgun2
An all-in-one Python-based installer for C/S new hire computers

## Windows Automated Installer

`installer.py` automates Windows and Dell Command updates. The script maintains state between reboots and resumes automatically via a scheduled task.
### Usage
1. Run `python installer.py` from an elevated PowerShell or Command Prompt. The script will request administrator rights if not already elevated.
2. The program logs progress to `installer.log` and the console.
3. If a reboot is required, the machine will restart and continue from the last completed step using the `ShotgunInstallerResume` scheduled task.
4. After updates finish, a pop-up confirms completion before proceeding.

## Windows OOBE Bypass Utility

`OOBE.py` automates opening a command prompt during Windows setup and typing a command, defaulting to `OOBE/bypassnro`. This allows bypassing the requirement to connect to a network during the out-of-box experience.

### Usage
```bash
python OOBE.py [--delay SECONDS] [command]
```
- `--delay` specifies how long to wait after the command prompt opens (default 2 seconds).
- `command` is the command to run once the prompt appears; the default is `OOBE/bypassnro`.

The script exits early on non-Windows systems or if `pyautogui` cannot be loaded.

### Automatic launch from USB

Copy `OOBE.py`, `OOBE.bat` and `autorun.inf` to the root of a USB drive. When
the drive is inserted during the first boot, Windows AutoPlay can start the
batch file which runs `OOBE.py` automatically. Modern Windows versions may
require manually approving the action.

## Requirements

Install the following Python package before running either script:

```bash
pip install pyautogui
```

`pyautogui` is required for keyboard and mouse automation in `OOBE.py`.
