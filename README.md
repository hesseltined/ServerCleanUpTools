# ServerCleanUpTools

Windows PowerShell utilities for server and share maintenance.

## Archive-OldFiles (Rust)

Native Windows binary port lives in **`archive-old-files/`** (Cargo crate `archive-old-files`). It reads/writes the same **`Archive-OldFiles.config.json`** and run log as the PowerShell script. On a Mac, use **`archive-old-files/setup-cross-compile.sh`** (installs **Homebrew LLVM**, **llvm-tools**, **cargo-xwin**, then runs **`cargo xwin build`**). See **`archive-old-files/README.md`** for prerequisites and manual steps.

Run on Windows: `archive-old-files.exe --help` (or pass `--input`, `--archive`, `--years`, etc.). Interactive prompts use the terminal when required options are omitted.

## Archive-OldFiles.ps1

Recursively finds files older than a chosen number of years (by last modified time), then **plans** or **moves** them to a separate archive location while **mirroring** the folder structure. Defaults to **preview-only** so you can review actions before committing.

After a successful run, settings are written to **`Archive-OldFiles.config.json`** next to the script (override with `-ConfigPath`, or skip with `-NoSaveConfig`). Each completed run also appends a short line block to **`Archive-OldFiles.run.log`** in the same folder (disable with **`-NoRunLog`**). **HTML** reports are written under the **archive** path, not the script directory. The next time you run the script with no paths on the command line, it loads the JSON and asks whether to use the saved values, then validates paths. Invalid entries are corrected one field at a time.

On a **file server**, when you are prompted for paths (not when all paths are passed on the command line), the script lists **published SMB disk shares** (via `Win32_Share`) so you can pick the **input** share, or **Other** to type a path. The **archive** location defaults to the local path of the share named **`Archive`** (change with `-ArchiveShareName`). If that share does not exist, you are prompted for the full archive folder path. Use **`-SkipShareMenu`** to disable the share picker and use plain prompts only.

**`-All`** runs a **preview-only** pass over **every** published disk share (same filter as the picker: Type 0, no trailing `$` in the name). **`-Commit` is ignored** (no moves). Pass **`ArchivePath`** and **`Years`** (or use saved config); **HTML** reports are written **one file per share** under the archive folder. Shares whose path **contains** the archive folder are skipped.

Age is judged by **last write time** by default (Explorer “Date modified”). You can use **`-AgeBasis LastAccessTime`** (last opened / NTFS last access; note that last-access updates can be disabled on a volume), **`-AgeBasis LatestWriteOrAccess`** (newer of modified and last access—only archives when both are older than the cutoff), **`-AgeBasis CreationTime`**, or **`-AgeBasis Earliest`** (older of creation vs last write, useful after copies/restores). Aliases include **Modified**, **Opened**, **ModifiedOrOpened**, and **A** / **B** / **C**. Enumeration uses **`-Force`** so hidden/system files are included.

In reports, owners that are unresolved SIDs or deleted accounts show as **No active user**.

The **`param`** block stays **mostly untyped** (avoids bind-time type mismatches on some hosts); **`[switch]$All`** is an exception so **`-All`** needs no value. Other validation runs after binding.

- **Requirements:** Windows PowerShell 5.1+
- **Documentation / download page:** [technologist.services/tools/archive-files/](https://technologist.services/tools/archive-files/)

## Troubleshooting (Windows)

- **Download blocked:** `Unblock-File -Path .\Archive-OldFiles.ps1`
- **Execution policy:** `powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Archive-OldFiles.ps1` (add `-InputPath`, `-ArchivePath`, `-Years`), or `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned`
- **Wrong extension:** Ensure the file is `Archive-OldFiles.ps1`, not `.ps1.txt`
- **Get-Help:** From the script folder: `Get-Help .\Archive-OldFiles.ps1 -Full`
- **`-All` “Missing an argument”:** Use a build **v1.7.1+** (switch parameter). On very old copies you could use `-All:$true`.

## Repository

https://github.com/hesseltined/ServerCleanUpTools
