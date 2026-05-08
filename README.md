# Kyron

Kyron is a lightweight desktop auto clicker built with Python and Tkinter.

## Features

- Save and reload click/keyboard scripts
- Global hotkey toggle
- Randomized click position and delay
- Desktop UI with script management
- Windows `.exe` packaging support

## Run From Source

```powershell
pip install pynput pillow
python kyron.py
```

## Build `.exe`

```powershell
pyinstaller --noconfirm --clean --onefile --windowed --name Kyron --icon logo.ico --add-data "logo.ico;." --add-data "logo.png;." kyron.py
```

The packaged app will be created in `dist/Kyron.exe`.

## Notes

- Runtime scripts are stored in a local `scripts` folder next to the app.
- Build artifacts are excluded from git and should be distributed through GitHub Releases instead.
