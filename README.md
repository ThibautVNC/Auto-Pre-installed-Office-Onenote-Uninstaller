# Uninstall Pre-installed Office

Those pre-installed Microsoft 365 and OneNote packages on every new laptop are pretty annoying, and the official SaRA tool won't take them off your hands anymore. Worry no more.

A PowerShell script that removes the pre-installed **Microsoft 365** and **Microsoft OneNote** Click-to-Run packages that ship with most new OEM Windows machines, so you can deploy the licensed version your organisation actually uses.

It runs from a simple menu, shows exactly what it found before touching anything, and reports live progress while each package is being removed.

![Screenshot of the script running](image.png)

---

## Why

Every new machine shows up with a trial or consumer build of Microsoft 365 already baked in. Push your own Office deployment on top of that and you either get a failure or a lovely mixed install that nobody wants to troubleshoot at 4 PM on a Friday. Removing it through **Settings > Apps** works, but it is slow and it needs somebody clicking through dialogs on every single box.

So: this script digs the packages out of the registry and hands them to Microsoft's own Click-to-Run uninstaller with the UI switched off. You pick an option, it does the work, you get a clean machine.

---

## What it does

1. Scans both the 64-bit and 32-bit uninstall registry hives:
   - `HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall`
   - `HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall`
2. Filters for display names starting with `Microsoft 365 - ` or `Microsoft OneNote - ` (the Click-to-Run products).
3. Lists everything it found, with a count, before asking for confirmation.
4. Calls each package's own `UninstallString` with `displaylevel=False` so the removal runs silently — no Microsoft dialogs to click through.
5. Shows a live spinner with elapsed time per package, since a full Office removal can take several minutes and would otherwise look frozen.
6. Translates the uninstaller exit code into plain language and prints a summary table, flagging when a restart is required.

Nothing is removed until you confirm. Choosing **Rescan** after a run lets you verify the machine is actually clean.

---

## Features

- **Menu-driven** — start, rescan, change language, or exit. No arguments to remember.
- **Multilingual interface** — English (default), Dutch and French, switchable at runtime from the menu.
- **Live progress** — per-package counter (`[1/3]`), spinner and elapsed timer.
- **Readable exit codes** — `0`, `1641` and `3010` are reported as success, with `1641`/`3010` noted as *restart required*, instead of showing a bare number.
- **Summary report** — every package with its final status once the run finishes.
- **Safe by default** — read-only scan until you explicitly confirm the removal.

---

## Requirements

| | |
|---|---|
| OS | Windows 10 / 11 |
| PowerShell | 5.1 or later (7.x also works) |
| Rights | Administrator — enforced by `#Requires -RunAsAdministrator` |

---

## Usage

Open PowerShell **as Administrator** and run:

```powershell
.\uninstall_preinstalled_office.ps1
```

If script execution is blocked on the machine, either unblock the file:

```powershell
Unblock-File .\uninstall_preinstalled_office.ps1
```

…or run it for that session only:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\uninstall_preinstalled_office.ps1
```

### Menu

| Option | Action |
|---|---|
| `1` | Start removal (asks for confirmation first) |
| `2` | Rescan the registry |
| `3` | Switch language — English / Nederlands / Français |
| `0` | Exit |

---

## Exit codes

The Click-to-Run uninstaller returns standard Windows Installer codes. The script maps the common ones:

| Code | Meaning |
|---|---|
| `0` | Removed successfully |
| `1641` | Success, restart required |
| `3010` | Success, restart required |
| `1602` | Cancelled by the user |
| `17002` | Scenario interrupted — another Office operation was already running |

Anything else is printed with its raw code so you can look it up.

---

## Notes

- Only **Click-to-Run** packages are matched. MSI-based Office installations use a different naming pattern and are deliberately left alone.
- Hit a `17002`? Something Office-related was already running. Let it finish, rescan, try again.
- After a `1641` or `3010`, reboot before installing your own Office deployment. It saves you a support ticket later.

---

## Credit

**Thibaut VNC**
