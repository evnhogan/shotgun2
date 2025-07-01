# shotgun2
An all-in-one Python-based installer for C/S new hire computers

## Windows Automated Installer

`installer.py` automates Windows and Dell Command updates, then executes installation packages in the `installers` directory. The script maintains state between reboots and resumes automatically via a scheduled task.
### Usage
1. Place any installation executables in the `installers` folder.
2. Run `python installer.py` from an elevated PowerShell or Command Prompt. The script will request administrator rights if not already elevated.
3. The program logs progress to `installer.log` and the console. If `tqdm` is installed, a progress bar is shown for installer files.
4. `.msi` packages are installed silently via `msiexec /qn /norestart` while other executables are run directly.
5. If a reboot is required, the machine will restart and continue from the last completed step using the `ShotgunInstallerResume` scheduled task.

## Windows OOBE Bypass Utility

`OOBE.py` automates opening a command prompt during Windows setup and typing a command, defaulting to `OOBE/bypassnro`. This allows bypassing the requirement to connect to a network during the out-of-box experience.

### Usage
```bash
python OOBE.py [--delay SECONDS] [command]
```
- `--delay` specifies how long to wait after the command prompt opens (default 2 seconds).
- `command` is the command to run once the prompt appears; the default is `OOBE/bypassnro`.

The script exits early on non-Windows systems or if `pyautogui` cannot be loaded.

## Requirements

Install the following Python packages before running either script:

```bash
pip install pyautogui tqdm
```

`pyautogui` is required for keyboard and mouse automation in `OOBE.py`, while
`tqdm` enables progress bars in `installer.py`.
