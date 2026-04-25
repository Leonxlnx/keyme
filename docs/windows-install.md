# Windows Install

Keyme is Windows-only.

## Normal Install

1. Download the release ZIP from GitHub.
2. Extract the ZIP.
3. Double-click `Setup-Keyme.cmd`.

Setup installs Keyme to:

```text
%LOCALAPPDATA%\Keyme
```

It creates one Desktop shortcut:

```text
Keyme
```

It also creates a Startup shortcut so Keyme launches after Windows login.

## Uninstall

Run:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%LOCALAPPDATA%\Keyme\scripts\uninstall.ps1"
```

The uninstall script stops Keyme, removes shortcuts, and removes the installed app folder.
