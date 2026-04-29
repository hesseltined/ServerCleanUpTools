# archive-old-files

Rust port of `Archive-OldFiles.ps1`: age-based archival with the same JSON config and run log.

## Cross-compile on macOS → Windows x64 (MSVC)

You accept [Microsoft’s MSVC/SDK license](https://go.microsoft.com/fwlink/?LinkId=2086102) when using **cargo-xwin**.

### One-time prerequisites

1. **Homebrew** [brew.sh](https://brew.sh)
2. **LLVM / clang** (cargo-xwin prerequisite):

   ```bash
   brew install llvm
   ```

3. **Rust** with Windows target and LLVM tools:

   ```bash
   rustup target add x86_64-pc-windows-msvc
   rustup component add llvm-tools
   ```

4. **cargo-xwin**:

   ```bash
   cargo install --locked cargo-xwin
   ```

5. Put Homebrew LLVM on your `PATH` when building (Apple Silicon vs Intel):

   ```bash
   # Apple Silicon
   export PATH="/opt/homebrew/opt/llvm/bin:$PATH"
   # Intel Mac
   # export PATH="/usr/local/opt/llvm/bin:$PATH"
   ```

### Build

From this directory:

```bash
chmod +x setup-cross-compile.sh
./setup-cross-compile.sh
```

Or manually:

```bash
export PATH="/opt/homebrew/opt/llvm/bin:$PATH"   # adjust for Intel Mac
cargo xwin cache xwin    # optional; caches SDK/CRT
cargo xwin build --release --target x86_64-pc-windows-msvc
```

Output: `target/x86_64-pc-windows-msvc/release/archive-old-files.exe`

### Test

- **On Windows**: copy the `.exe` plus any config you need; run `archive-old-files.exe --help`.
- **On Mac** (smoke): `cargo build` only builds the stub error path on non-Windows; real behavior requires Windows or **Wine**:

  ```bash
  brew install --cask wine-stable   # optional
  cargo xwin test --target x86_64-pc-windows-msvc   # if wine works for your setup
  ```

### CLI quick reference

```text
archive-old-files.exe --input "D:\Data" --archive "D:\Arc" --years 7
archive-old-files.exe --input "D:\Data" --archive "D:\Arc" --years 7 --commit --output HTML
archive-old-files.exe --all --archive "D:\Reports" --years 7
archive-old-files.exe --test-email --config "C:\path\Archive-OldFiles.config.json"
```

Default config path: `Archive-OldFiles.config.json` next to the `.exe`.
