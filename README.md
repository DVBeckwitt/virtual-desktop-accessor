# virtual-desktop-accessor

This repository combines three related pieces:

- `src/`: the `winvd` Rust crate for Windows 11 virtual desktop access
- `dll/`: a `cdylib` wrapper that exports the C ABI used by AutoHotkey
- `quick-desktop-hotkeys.ahk`: a machine-specific AutoHotkey v1 script built on top of the DLL

The upstream Rust crate documentation still lives in [README-crate.md](./README-crate.md). This top-level README focuses on the workspace layout in this repository and the AutoHotkey workflow that is being maintained here.

## Requirements

- Windows 11 24H2 or newer
- Rust if you want to build the DLL from source
- AutoHotkey v1 for `quick-desktop-hotkeys.ahk`

## Repository Layout

- `src/`: core `winvd` Rust implementation
- `dll/`: builds `VirtualDesktopAccessor.dll`
- `quick-desktop-hotkeys.ahk`: personal desktop-navigation script with app-specific helpers

## Building

Build the full workspace:

```powershell
cargo build --workspace
```

Build release artifacts:

```powershell
cargo build --release --workspace
```

The DLL is produced by the `dll` crate and ends up at:

- `target\debug\VirtualDesktopAccessor.dll`
- `target\release\VirtualDesktopAccessor.dll`

If you only need a prebuilt DLL, you can also use the upstream [release downloads](https://github.com/Ciantic/VirtualDesktopAccessor/releases/).

## AutoHotkey Scripts

The `quick-desktop-hotkeys.ahk` script is intentionally local and expects a few machine-specific paths near the top of the file to be adjusted before use:

- `VDA_PATH`
- `RestartShortcutPath`
- `FallbackScriptPath`

It also contains app-specific title and process hints for the windows it manages.

### `quick-desktop-hotkeys.ahk` Highlights

- `Ctrl+1` through `Ctrl+0`: jump to desktops 1 through 10
- `Win+1` through `Win+0`: same desktop jump bindings on the Windows key
- `MButton + WheelUp/WheelDown`: switch to the next or previous desktop
- `Ctrl+Shift+WheelUp/WheelDown`: move the active window to the next or previous desktop
- `Shift+WheelUp/WheelDown`: cycle tabs in supported apps, otherwise fall back to normal scrolling
- `Ctrl+Alt+R`: reload the AutoHotkey script

There are additional app-launch and focus shortcuts in the script for the local workstation setup.

## Exported DLL Functions

All exported functions return `-1` on error unless otherwise noted.

```rust
fn GetCurrentDesktopNumber() -> i32
fn GetDesktopCount() -> i32
fn GetDesktopIdByNumber(number: i32) -> GUID
fn GetDesktopNumberById(desktop_id: GUID) -> i32
fn GetWindowDesktopId(hwnd: HWND) -> GUID
fn GetWindowDesktopNumber(hwnd: HWND) -> i32
fn IsWindowOnCurrentVirtualDesktop(hwnd: HWND) -> i32
fn MoveWindowToDesktopNumber(hwnd: HWND, desktop_number: i32) -> i32
fn GoToDesktopNumber(desktop_number: i32) -> i32
fn SetDesktopName(desktop_number: i32, in_name_ptr: *const i8) -> i32
fn GetDesktopName(desktop_number: i32, out_utf8_ptr: *mut u8, out_utf8_len: usize) -> i32
fn RegisterPostMessageHook(listener_hwnd: HWND, message_offset: u32) -> i32
fn UnregisterPostMessageHook(listener_hwnd: HWND) -> i32
fn IsPinnedWindow(hwnd: HWND) -> i32
fn PinWindow(hwnd: HWND) -> i32
fn UnPinWindow(hwnd: HWND) -> i32
fn IsPinnedApp(hwnd: HWND) -> i32
fn PinApp(hwnd: HWND) -> i32
fn UnPinApp(hwnd: HWND) -> i32
fn IsWindowOnDesktopNumber(hwnd: HWND, desktop_number: i32) -> i32
fn CreateDesktop() -> i32
fn RemoveDesktop(remove_desktop_number: i32, fallback_desktop_number: i32) -> i32
```
