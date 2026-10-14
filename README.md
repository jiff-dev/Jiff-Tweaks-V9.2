# Jiff Tweaks V9.2

Free Windows tweaks for gaming and low latency. **One PowerShell file. No installer, no separate app.**
Live dashboard, one-click profiles, a preview of the exact old and new value before you apply anything, and every change can be undone.

Community: https://discord.gg/fuegoscrims

## Run it

Open **PowerShell as Administrator** and paste:

```powershell
irm https://raw.githubusercontent.com/jiff-dev/Jiff-free-Tweaks-V9/main/Jiff-Tweaks.ps1 | iex
```

Or download `Jiff-Tweaks.ps1` and either double-click **Jiff Tweaks.bat**, or run:

```powershell
powershell -ExecutionPolicy Bypass -File .\Jiff-Tweaks.ps1
```

A window opens. If you are not an administrator it asks for permission first (use the same Windows account you play on).

## What you get

| | |
|---|---|
| **Cockpit** | Live CPU, GPU, memory and disk, three profiles (Safe, Competitive, Full), a tune score, suggested next tweaks, before/after numbers of your last run, driver latency health |
| **Live** | Per-core CPU, clock, GPU load/temp/VRAM (NVIDIA), disk and network throughput, DPC and interrupt time, top processes, optional ping/jitter/loss monitor |
| **Tweaks** | Performance, Services, Network, Debloat and Privacy, GPU (NVIDIA and AMD), Expert Lab. Each card: Preview (exact old and new values), Apply |
| **Changes** | Everything that was changed, with the original value next to it |
| **Restore** | One click puts every saved original back (removed Store apps can be reinstalled from the Store) |
| **Logs** | Every run, line by line |

## Use it from the console

```powershell
.\Jiff-Tweaks.ps1 -List                    # every tweak and whether it is on
.\Jiff-Tweaks.ps1 -Code reg_gaming         # print the code of one tweak (not shown in the window)
.\Jiff-Tweaks.ps1 -Preview svc_basic       # what it would change
.\Jiff-Tweaks.ps1 -Preset safe             # safe | competitive | full
.\Jiff-Tweaks.ps1 -Apply reg_gaming,privacy
.\Jiff-Tweaks.ps1 -Restore                 # put everything back
.\Jiff-Tweaks.ps1 -NoGui                   # text menu
.\Jiff-Tweaks.ps1 -SelfTest                # checks that everything the tweaks need works on this PC
```

Add `-Yes` to skip the confirmation question.

## Safe by design

- **It is this one file.** Open it in Notepad. The top of the file lists where everything is. Nothing else is downloaded or installed.
- **Originals are saved first** (`C:\JiffTweaks\journal.jsonl`, readable by administrators only). The first saved value is never overwritten, so Restore always goes back to the real original.
- **The window is local.** It talks to `127.0.0.1` only, needs a secret token made at launch, and can only start tweaks from the list in the file. It cannot run any other command.
- **No telemetry, no accounts.** The internet is used only if you switch on the ping monitor, press *Check for updates*, or apply the branding tweak (it downloads one wallpaper).
- Risky tweaks (aggressive, security, destructive) show a warning before they run. Defender and CPU mitigations are only touched in the Expert Lab.
- The Full profile creates a Windows restore point first.

## Profiles

- **Safe** - gaming and input registry, privacy, task cleanup, TCP settings, power plan. Nothing that breaks apps.
- **Competitive** - Safe plus latency/scheduler registry, Nagle off, adapter properties, USB/device power saving off, boot timer, basic services.
- **Full** - everything: advanced service tiers, GPU tweaks, debloat, OneDrive removal.

## If something does not work

Run `.\Jiff-Tweaks.ps1 -SelfTest` and read the PASS/FAIL lines. Each FAIL names the part (registry, services, scheduled tasks, powercfg, bcdedit, WMI, Store apps, network cmdlets, local server). Copy the output when you ask for help.

Notes: requires Windows 10/11 and Windows PowerShell 5.1 (built in) or PowerShell 7. Drop your own `JiffOSUltimate.pow` next to the script and the power plan tweak imports it instead of building one.

## For the repo owner

Upload `Jiff-Tweaks.ps1` and `version.txt` to the root of the `main` branch. The one-line command above reads the first, the *Check for updates* button reads the second. When you release a new version, change the number in `version.txt`.
