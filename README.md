# ServerCleanUpTools

Windows PowerShell utilities for server and share maintenance.

## Archive-OldFiles.ps1

Recursively finds files older than a chosen number of years (by last modified time), then **plans** or **moves** them to a separate archive location while **mirroring** the folder structure. Defaults to **preview-only** so you can review actions before committing.

After a successful run, settings are written to **`Archive-OldFiles.config.json`** next to the script (override with `-ConfigPath`, or skip with `-NoSaveConfig`). Each completed run also appends a short line block to **`Archive-OldFiles.run.log`** in the same folder (disable with **`-NoRunLog`**). **HTML** reports are written under the **archive** path, not the script directory. The next time you run the script with no paths on the command line, it loads the JSON and asks whether to use the saved values, then validates paths. Invalid entries are corrected one field at a time.

On a **file server**, when you are prompted for paths (not when all paths are passed on the command line), the script lists **published SMB disk shares** (via `Win32_Share`) so you can pick the **input** share, or **Other** to type a path. The **archive** location defaults to the local path of the share named **`Archive`** (change with `-ArchiveShareName`). If that share does not exist, you are prompted for the full archive folder path. Use **`-SkipShareMenu`** to disable the share picker and use plain prompts only.

Age is judged by **last write time** by default (Explorer “Date modified”). After a **copy or restore**, that date can be recent even when the file is old; use **`-AgeBasis CreationTime`** or **`-AgeBasis Earliest`** (older of creation vs last write) to match what you expect. Enumeration uses **`-Force`** so hidden/system files are included.

The **`param`** block is **minimal** (no `[Parameter]`, no `[switch]`) with validation in script (v1.3.0+) so Windows PowerShell does not fail with **argument types do not match** at bind time on some hosts.

- **Requirements:** Windows PowerShell 5.1+
- **Documentation / download page:** [technologist.services/tools/archive-files/](https://technologist.services/tools/archive-files/)

## Troubleshooting (Windows)

- **Download blocked:** `Unblock-File -Path .\Archive-OldFiles.ps1`
- **Execution policy:** `powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Archive-OldFiles.ps1` (add `-InputPath`, `-ArchivePath`, `-Years`), or `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned`
- **Wrong extension:** Ensure the file is `Archive-OldFiles.ps1`, not `.ps1.txt`
- **Get-Help:** From the script folder: `Get-Help .\Archive-OldFiles.ps1 -Full`

## Repository

https://github.com/hesseltined/ServerCleanUpTools
