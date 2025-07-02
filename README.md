# shotgun2
An all-in-one Python-based installer for C/S new hire computers

## Windows Automated Installer

`installer.py` automates Windows and Dell Command updates. The script maintains state between reboots and resumes automatically via a scheduled task.
### Usage
1. Run `python installer.py` from an elevated PowerShell or Command Prompt. The script will request administrator rights if not already elevated.
2. The program logs progress to `installer.log` and the console.
3. If a reboot is required, the machine will restart and continue from the last completed step using the `ShotgunInstallerResume` scheduled task.
4. After updates finish, a pop-up confirms completion before proceeding.
5. To resume manually after an unexpected reboot, rerun `python installer.py`
   The script continues from the saved step in `installer_state.json`.
6. Use `--windows-updates` and `--dell-updates` to run only selected steps.

## Windows OOBE Bypass Utility

`OOBE.py` automates opening a command prompt during Windows setup and typing a command, defaulting to `OOBE/bypassnro`. This allows bypassing the requirement to connect to a network during the out-of-box experience.

### Usage
```bash
python OOBE.py [--delay SECONDS] [--dry-run] [command]
```
- `--delay` specifies how long to wait after the command prompt opens (default 2 seconds).
- `--dry-run` logs actions without sending any keyboard input.
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

## Automated OOBE Navigation

`post_oobe.py` continues the Windows setup after running `OOBE.py`. It uses only
built-in modules and simulates key presses via the Windows API. The script
selects English and the US keyboard layout, chooses to continue without
internet and stops at the device name screen. Provide a numeric identifier when
running the script and it enters `CS-<number>` automatically, leaves the
password blank and clicks through the remaining setup dialogs until reaching the
desktop. The program exits immediately if run on a non-Windows system or if the
device number is not numeric.

### Usage

```bash
python post_oobe.py 123
```

This would type `CS-123` when prompted for a device name.

