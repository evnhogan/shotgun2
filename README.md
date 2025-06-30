# shotgun2
An all-in-one Python-based installer for C/S new hire computers

## Windows Automated Installer

`installer.py` automates Windows and Dell Command updates, then executes installation packages in the `installers` directory. The script maintains state between reboots and resumes automatically via a scheduled task.
### Usage
1. Place any installation executables in the `installers` folder.
2. Run `python installer.py` from an elevated PowerShell or Command Prompt. The script will request administrator rights if not already elevated.
3. The program logs progress to `installer.log` and the console. If `tqdm` is installed, a progress bar is shown for installer files.
4. If a reboot is required, the machine will restart and continue from the last completed step using the `ShotgunInstallerResume` scheduled task.
