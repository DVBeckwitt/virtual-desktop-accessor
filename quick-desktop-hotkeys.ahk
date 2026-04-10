; AutoHotkey v1 script
; Bind Ctrl+1..Ctrl+0 and Win+1..Win+0 to jump directly to virtual desktops 1..10 (0-indexed internally).

#NoEnv
#UseHook On
#InstallKeybdHook
#InstallMouseHook
#SingleInstance Force
SetWorkingDir, %A_ScriptDir%
SetBatchLines, -1
FileGetTime, LastScriptModTime, %A_ScriptFullPath%, M
RestartShortcutPath := A_AppData . "\Microsoft\Windows\Start Menu\Programs\Startup\quick-desktop-hotkeys.ahk - Shortcut.lnk"
StartupShortcutPath := RestartShortcutPath
FallbackScriptPath := "C:\Users\Kenpo\OneDrive\Documents\GitHub\virtual-desktop-accessor\quick-desktop-hotkeys.ahk"
RA_SIM_RUN_PATH := "C:\Users\Kenpo\OneDrive\Documents\GitHub\PhD Work\ra_sim\run_ra_sim.bat"
EnsureStartupShortcutExists()

; Load the DLL from your installed location
VDA_PATH := "C:\Users\Kenpo\OneDrive\VirtualDesktopAccessor-rust\VirtualDesktopAccessor.dll"
if (!FileExist(VDA_PATH)) {
    MsgBox, 16, VirtualDesktopAccessor, VirtualDesktopAccessor.dll was not found at:`n%VDA_PATH%
    ExitApp
}

hVirtualDesktopAccessor := DllCall("LoadLibrary", "Str", VDA_PATH, "Ptr")
if (!hVirtualDesktopAccessor) {
    MsgBox, 16, VirtualDesktopAccessor, Failed to load VirtualDesktopAccessor.dll.
    ExitApp
}

GetDesktopCountProc := DllCall("GetProcAddress", "Ptr", hVirtualDesktopAccessor, "AStr", "GetDesktopCount", "Ptr")
GoToDesktopNumberProc := DllCall("GetProcAddress", "Ptr", hVirtualDesktopAccessor, "AStr", "GoToDesktopNumber", "Ptr")
GetCurrentDesktopNumberProc := DllCall("GetProcAddress", "Ptr", hVirtualDesktopAccessor, "AStr", "GetCurrentDesktopNumber", "Ptr")
GetWindowDesktopNumberProc := DllCall("GetProcAddress", "Ptr", hVirtualDesktopAccessor, "AStr", "GetWindowDesktopNumber", "Ptr")
IsWindowOnCurrentVirtualDesktopProc := DllCall("GetProcAddress", "Ptr", hVirtualDesktopAccessor, "AStr", "IsWindowOnCurrentVirtualDesktop", "Ptr")
MoveWindowToDesktopNumberProc := DllCall("GetProcAddress", "Ptr", hVirtualDesktopAccessor, "AStr", "MoveWindowToDesktopNumber", "Ptr")

; Edit these tokens if your Codex executable/title differs.
CODEx_TITLE_HINT := "Codex"
CODEx_PROC_HINTS := "Codex.exe|codex.exe"
CODEx_RUN_HINTS := "Codex.exe|codex.exe"
OUTLOOK_TITLE_HINT := "Outlook"
OUTLOOK_PROC_HINTS := "OUTLOOK.EXE|outlook.exe|OLK.EXE|olk.exe"
OUTLOOK_RUN_HINT := "outlook.exe"
BRAVE_NIGHTLY_TITLE_HINT := "Brave Nightly"
BRAVE_NIGHTLY_PROC_HINTS := "brave.exe|Brave.exe|BraveBrowser.exe"
BRAVE_NIGHTLY_RUN_HINT := "brave.exe"
BRAVE_NIGHTLY_PATH_HINTS := "brave-browser-nightly|brave-browser\\nightly"
RUSTDESK_TITLE_HINTS := "RustDesk|RustDesk Remote Desktop"
RUSTDESK_PROC_HINTS := "RustDesk.exe|rustdesk.exe"
RUSTDESK_RUN_HINTS := "C:\Program Files\RustDesk\RustDesk.exe|C:\Program Files (x86)\RustDesk\RustDesk.exe|rustdesk.exe"
CHATGPT_TITLE_HINTS := "ChatGPT|ChatGPT Desktop|ChatGPT - OpenAI|OpenAI"
CHATGPT_PROC_HINTS := "ChatGPT.exe|chatgpt.exe|ChatGPTDesktop.exe"
CHATGPT_RUN_HINTS := A_AppData . "\..\Local\Programs\ChatGPT\ChatGPT.exe|" A_ProgramFiles . "\OpenAI\ChatGPT\ChatGPT.exe|" A_ProgramFiles . "\OpenAI\ChatGPT Desktop\ChatGPT.exe|chatgpt.exe"
DISCORD_TITLE_HINTS := "Discord|Discord Canary|Discord PTB"
DISCORD_PROC_HINTS := "Discord.exe|discord.exe|DiscordCanary.exe|DiscordPTB.exe"
DISCORD_VERSIONED_RUN_HINT := A_LocalAppData . "\Discord\app-1.0.9231\Discord.exe"
DISCORD_RUN_HINTS := A_LocalAppData . "\Discord\Update.exe|C:\Users\Kenpo\AppData\Local\Programs\Discord\Discord.exe|" . DISCORD_VERSIONED_RUN_HINT
DISCORD_PATH_HINTS := "discordapp|discord"
PHONE_LINK_RUN_HINT := "ms-phone:"
NEOVIM_TITLE_HINTS := "Neovim|Neovide|NeoVim"
NEOVIM_PROC_HINTS := "nvim.exe|nvim-qt.exe|Neovide.exe|neovim.exe|Neovim.exe|neovide.exe"
NEOVIM_RUN_HINTS := A_ProgramFiles . "\Neovim\bin\nvim-qt.exe|" A_ProgramFiles . "\Neovim\bin\nvim.exe|" A_LocalAppData . "\Programs\Neovim\bin\nvim-qt.exe|" A_LocalAppData . "\Programs\Neovim\bin\nvim.exe|nvim-qt.exe|nvim.exe"
CONVERT_SITE_URL := "https://p2r3.github.io/convert/"
YOUTUBE_MUSIC_URL := "https://music.youtube.com/"
YOUTUBE_MUSIC_SHORTCUT := "C:\Users\Kenpo\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\YouTube Music.lnk"
HotkeysEnabled := 1
PendingCtrlDPress := 0

SendPlainPageUp() {
    leftCtrlDown := GetKeyState("LCtrl", "P")
    rightCtrlDown := GetKeyState("RCtrl", "P")
    if (leftCtrlDown) {
        SendInput, {LCtrl up}
    }
    if (rightCtrlDown) {
        SendInput, {RCtrl up}
    }
    SendInput, ^{Home}
    if (leftCtrlDown) {
        SendInput, {LCtrl down}
    }
    if (rightCtrlDown) {
        SendInput, {RCtrl down}
    }
}

SendPlainPageDown() {
    leftCtrlDown := GetKeyState("LCtrl", "P")
    rightCtrlDown := GetKeyState("RCtrl", "P")
    if (leftCtrlDown) {
        SendInput, {LCtrl up}
    }
    if (rightCtrlDown) {
        SendInput, {RCtrl up}
    }
    SendInput, ^{End}
    if (leftCtrlDown) {
        SendInput, {LCtrl down}
    }
    if (rightCtrlDown) {
        SendInput, {RCtrl down}
    }
}

IsActiveCodexWindow() {
    global CODEx_PROC_HINTS, CODEx_TITLE_HINT

    hwnd := WinExist("A")
    if (!hwnd) {
        return false
    }

    WinGet, processName, ProcessName, ahk_id %hwnd%
    if (!ErrorLevel && processName != "") {
        StringLower, processNameLower, processName
        Loop, Parse, CODEx_PROC_HINTS, |
        {
            candidateProc := A_LoopField
            if (candidateProc = "") {
                continue
            }

            StringLower, candidateProcLower, candidateProc
            if (processNameLower = candidateProcLower) {
                return true
            }
        }
    }

    WinGetTitle, title, ahk_id %hwnd%
    if (!ErrorLevel && title != "" && InStr(title, CODEx_TITLE_HINT)) {
        return true
    }

    return false
}

HandleShiftWheel(direction) {
    leftShiftDown := GetKeyState("LShift", "P")
    rightShiftDown := GetKeyState("RShift", "P")
    if (leftShiftDown) {
        SendInput, {LShift up}
    }
    if (rightShiftDown) {
        SendInput, {RShift up}
    }
    if (IsActiveCodexWindow()) {
        if (direction = "up") {
            SendInput, ^+{[}
        } else {
            SendInput, ^+{]}
        }
    } else if (IsActiveTabScrollWindow()) {
        if (direction = "up") {
            SendInput, ^+{Tab}
        } else {
            SendInput, ^{Tab}
        }
    } else {
        if (direction = "up") {
            SendInput, {WheelUp}
        } else {
            SendInput, {WheelDown}
        }
    }
    if (leftShiftDown) {
        SendInput, {LShift down}
    }
    if (rightShiftDown) {
        SendInput, {RShift down}
    }
}

GoToDesktopIfExists(index) {
    global GetDesktopCountProc, GoToDesktopNumberProc
    desktopCount := DllCall(GetDesktopCountProc, "Int")
    if (index < 0 || index >= desktopCount) {
        return
    }
    DllCall(GoToDesktopNumberProc, "Int", index)
}

GetDesktopCount() {
    global GetDesktopCountProc
    return DllCall(GetDesktopCountProc, "Int")
}

GoToPrevDesktop() {
    global GetCurrentDesktopNumberProc
    current := DllCall(GetCurrentDesktopNumberProc, "Int")
    count := GetDesktopCount()
    if (count <= 1) {
        return
    }
    last := count - 1
    target := current - 1
    if (target < 0) {
        target := last
    }
    GoToDesktopIfExists(target)
}

GoToNextDesktop() {
    global GetCurrentDesktopNumberProc
    current := DllCall(GetCurrentDesktopNumberProc, "Int")
    count := GetDesktopCount()
    if (count <= 1) {
        return
    }
    target := current + 1
    if (target >= count) {
        target := 0
    }
    GoToDesktopIfExists(target)
}

IsActiveBraveWindow() {
    global BRAVE_NIGHTLY_PROC_HINTS, BRAVE_NIGHTLY_TITLE_HINT, BRAVE_NIGHTLY_PATH_HINTS
    WinGet, activeHwnd, ID, A
    if (!activeHwnd) {
        return false
    }

    WinGet, activeProcessName, ProcessName, ahk_id %activeHwnd%
    if (ErrorLevel || activeProcessName = "") {
        return false
    }

    StringLower, activeProcessNameLower, activeProcessName
    Loop, Parse, BRAVE_NIGHTLY_PROC_HINTS, |
    {
        processHint := A_LoopField
        if (processHint = "") {
            continue
        }
        StringLower, processHintLower, processHint
        if (processHintLower = activeProcessNameLower) {
            WinGetTitle, activeTitle, ahk_id %activeHwnd%
            activeTitleLower := ""
            if (activeTitle) {
                StringLower, activeTitleLower, activeTitle
                if (InStr(activeTitleLower, "brave")) {
                    if (InStr(activeTitleLower, "nightly")) {
                        return true
                    }
                    if (InStr(activeTitleLower, "brave nightly")) {
                        return true
                    }
                }
            }

            WinGet, activeProcessPath, ProcessPath, ahk_id %activeHwnd%
            if (activeProcessPath) {
                StringLower, activeProcessPathLower, activeProcessPath
                if (InStr(activeProcessPathLower, "brave")) {
                    if (InStr(activeProcessPathLower, "nightly")) {
                        return true
                    }
                    Loop, Parse, BRAVE_NIGHTLY_PATH_HINTS, |
                    {
                        pathHint := A_LoopField
                        if (pathHint = "") {
                            continue
                        }
                        if (InStr(activeProcessPathLower, pathHint)) {
                            return true
                        }
                    }
                }
            }
            return false
        }
    }
    return false
}

IsActiveDiscordWindow() {
    global DISCORD_PROC_HINTS, DISCORD_PATH_HINTS, DISCORD_TITLE_HINTS
    WinGet, activeHwnd, ID, A
    if (!activeHwnd) {
        return false
    }

    WinGet, activeProcessName, ProcessName, ahk_id %activeHwnd%
    if (ErrorLevel || activeProcessName = "") {
        return false
    }

    StringLower, activeProcessNameLower, activeProcessName
    Loop, Parse, DISCORD_PROC_HINTS, |
    {
        processHint := A_LoopField
        if (processHint = "") {
            continue
        }
        StringLower, processHintLower, processHint
        if (processHintLower = activeProcessNameLower) {
            return true
        }
    }

    WinGetTitle, activeTitle, ahk_id %activeHwnd%
    if (activeTitle) {
        StringLower, activeTitleLower, activeTitle
        if (InStr(activeTitleLower, "discord")) {
            return true
        }
    }

    WinGet, activeProcessPath, ProcessPath, ahk_id %activeHwnd%
    if (activeProcessPath) {
        StringLower, activeProcessPathLower, activeProcessPath
        Loop, Parse, DISCORD_PATH_HINTS, |
        {
            pathHint := A_LoopField
            if (pathHint = "") {
                continue
            }
            if (InStr(activeProcessPathLower, pathHint)) {
                return true
            }
        }
    }

    return false
}

IsActiveTabScrollWindow() {
    return IsActiveBraveWindow() || IsActiveDiscordWindow()
}

MoveActiveWindowAndSwitchTo(targetDesktop) {
    global MoveWindowToDesktopNumberProc, GoToDesktopNumberProc
    if (targetDesktop = "" || targetDesktop < 0) {
        return
    }

    WinGet, hwnd, ID, A
    if (!hwnd) {
        return
    }

    DllCall(MoveWindowToDesktopNumberProc, "Ptr", hwnd, "Int", targetDesktop)
    DllCall(GoToDesktopNumberProc, "Int", targetDesktop)
}

MoveActiveWindowToLeftDesktop() {
    global GetCurrentDesktopNumberProc
    if (GetDesktopCount() <= 1) {
        return
    }
    current := DllCall(GetCurrentDesktopNumberProc, "Int")
    target := current - 1
    if (target < 0) {
        target := GetDesktopCount() - 1
    }
    MoveActiveWindowAndSwitchTo(target)
}

MoveActiveWindowToRightDesktop() {
    global GetCurrentDesktopNumberProc
    if (GetDesktopCount() <= 1) {
        return
    }
    current := DllCall(GetCurrentDesktopNumberProc, "Int")
    count := GetDesktopCount()
    target := current + 1
    if (target >= count) {
        target := 0
    }
    MoveActiveWindowAndSwitchTo(target)
}

MoveToCurrentDesktop(hwnd) {
    global GetCurrentDesktopNumberProc, GetWindowDesktopNumberProc, MoveWindowToDesktopNumberProc, IsWindowOnCurrentVirtualDesktopProc
    if (!hwnd) {
        return 0
    }

    currentDesktop := DllCall(GetCurrentDesktopNumberProc, "Int")
    windowDesktop := DllCall(GetWindowDesktopNumberProc, "Ptr", hwnd, "Int")
    if (windowDesktop = currentDesktop || windowDesktop < 0) {
        return 1
    }

    DllCall(MoveWindowToDesktopNumberProc, "Ptr", hwnd, "Int", currentDesktop)
    Loop 20 {
        if (DllCall(IsWindowOnCurrentVirtualDesktopProc, "Ptr", hwnd, "Int") = 1) {
            return 1
        }
        Sleep, 25
    }
    return 0
}

FindCodexWindow() {
    global CODEx_PROC_HINTS, CODEx_TITLE_HINT

    ; Try direct exe matches first.
    Loop, Parse, CODEx_PROC_HINTS, |
    {
        proc := A_LoopField
        if (proc = "") {
            continue
        }
        WinGet, hwnd, ID, ahk_exe %proc%
        if (hwnd) {
            return hwnd
        }
    }

    DetectHiddenWindows, On
    WinGet, hwndList, List
    Loop, % hwndList {
        hwnd := hwndList%A_Index%
        WinGet, processName, ProcessName, ahk_id %hwnd%
        if (ErrorLevel || processName = "") {
            continue
        }

        StringLower, processNameLower, processName
        Loop, Parse, CODEx_PROC_HINTS, |
        {
            candidateLower := A_LoopField
            StringLower, candidateLower, candidateLower
            if (processNameLower = candidateLower) {
                return hwnd
            }
        }

        WinGetTitle, title, ahk_id %hwnd%
        if (title && InStr(title, CODEx_TITLE_HINT)) {
            return hwnd
        }
    }

    ; Final fallback by title only.
    WinGet, hwnd, ID, %CODEx_TITLE_HINT%
    return hwnd
}

LaunchCodex() {
    global CODEx_RUN_HINTS

    Loop, Parse, CODEx_RUN_HINTS, |
    {
        candidate := A_LoopField
        if (candidate = "") {
            continue
        }

        if (InStr(candidate, " ") && !RegExMatch(candidate, "^""")) {
            candidateRun := """" . candidate . """"
        } else {
            candidateRun := candidate
        }

        Run, %candidateRun%
        Sleep, 150
        return
    }
}

FocusCodex(pullToCurrent := 0) {
    global GetCurrentDesktopNumberProc, GetWindowDesktopNumberProc, GoToDesktopNumberProc, IsWindowOnCurrentVirtualDesktopProc
    hwnd := FindCodexWindow()
    if (!hwnd) {
        return
    }

    if (pullToCurrent) {
        MoveToCurrentDesktop(hwnd)
    } else {
        codexDesktop := DllCall(GetWindowDesktopNumberProc, "Ptr", hwnd, "Int")
        if (codexDesktop < 0) {
            return
        }
        currentDesktop := DllCall(GetCurrentDesktopNumberProc, "Int")

        if (codexDesktop != currentDesktop) {
            DllCall(GoToDesktopNumberProc, "Int", codexDesktop)
            Loop 20 {
                if (DllCall(IsWindowOnCurrentVirtualDesktopProc, "Ptr", hwnd, "Int") = 1) {
                    break
                }
                Sleep, 25
            }
        }
    }

    Loop 5 {
        WinShow, ahk_id %hwnd%
        WinActivate, ahk_id %hwnd%
        if (WinActive("ahk_id " . hwnd)) {
            break
        }
        Sleep, 40
    }
}

GetSelectionOrClipboardText() {
    oldClipboard := Clipboard
    oldClipboardAll := ClipboardAll

    Clipboard := ""
    SendInput, ^c
    ClipWait, 0.5
    if (ErrorLevel) {
        selectedText := ""
    } else {
        selectedText := Clipboard
    }
    Clipboard := oldClipboardAll

    if (Trim(selectedText) != "") {
        return selectedText
    }

    return Trim(oldClipboard)
}

OpenChatGPTWithSelectionOrClipboard() {
    textToSend := GetSelectionOrClipboardText()
    chatHwnd := FocusChatGPT(0)
    if (!chatHwnd) {
        return
    }
    WinActivate, ahk_id %chatHwnd%
    WinWaitActive, ahk_id %chatHwnd%, , 2
    if (ErrorLevel) {
        return
    }
    Sleep, 120
    SendInput, ^n
    Sleep, 100

    if (textToSend = "") {
        return
    }

    clipboardSaved := ClipboardAll
    Clipboard := textToSend
    SendInput, ^v
    Sleep, 80
    Clipboard := clipboardSaved
}

FindOutlookWindow() {
    global OUTLOOK_PROC_HINTS, OUTLOOK_TITLE_HINT

    Loop, Parse, OUTLOOK_PROC_HINTS, |
    {
        proc := A_LoopField
        if (proc = "") {
            continue
        }
        WinGet, hwnd, ID, ahk_exe %proc%
        if (hwnd) {
            return hwnd
        }
    }

    DetectHiddenWindows, On
    WinGet, hwndList, List
    Loop, % hwndList {
        hwnd := hwndList%A_Index%
        WinGet, processName, ProcessName, ahk_id %hwnd%
        if (ErrorLevel || processName = "") {
            continue
        }

        StringLower, processNameLower, processName
        Loop, Parse, OUTLOOK_PROC_HINTS, |
        {
            candidateLower := A_LoopField
            StringLower, candidateLower, candidateLower
            if (processNameLower = candidateLower) {
                return hwnd
            }
        }

        WinGetTitle, title, ahk_id %hwnd%
        if (title && InStr(title, OUTLOOK_TITLE_HINT)) {
            return hwnd
        }
    }

    WinGet, hwnd, ID, %OUTLOOK_TITLE_HINT%
    return hwnd
}

FocusOutlook(pullToCurrent := 0) {
    global GetCurrentDesktopNumberProc, GetWindowDesktopNumberProc, GoToDesktopNumberProc, IsWindowOnCurrentVirtualDesktopProc
    hwnd := FindOutlookWindow()
    if (!hwnd) {
        Run, %OUTLOOK_RUN_HINT%
        Loop 120 {
            hwnd := FindOutlookWindow()
            if (hwnd) {
                break
            }
            Sleep, 50
        }
    }
    if (!hwnd) {
        return
    }

    outlookDesktop := DllCall(GetWindowDesktopNumberProc, "Ptr", hwnd, "Int")
    if (pullToCurrent) {
        MoveToCurrentDesktop(hwnd)
    } else {
        if (outlookDesktop < 0) {
            return
        }
        currentDesktop := DllCall(GetCurrentDesktopNumberProc, "Int")

        if (outlookDesktop != currentDesktop) {
            DllCall(GoToDesktopNumberProc, "Int", outlookDesktop)
            Loop 20 {
                if (DllCall(IsWindowOnCurrentVirtualDesktopProc, "Ptr", hwnd, "Int") = 1) {
                    break
                }
                Sleep, 25
            }
        }
    }

    Loop 5 {
        WinShow, ahk_id %hwnd%
        WinActivate, ahk_id %hwnd%
        if (WinActive("ahk_id " . hwnd)) {
            break
        }
        Sleep, 40
    }
}

FindBraveNightlyWindow() {
    global BRAVE_NIGHTLY_PROC_HINTS, BRAVE_NIGHTLY_TITLE_HINT

    Loop, Parse, BRAVE_NIGHTLY_PROC_HINTS, |
    {
        proc := A_LoopField
        if (proc = "") {
            continue
        }
        WinGet, hwnd, ID, ahk_exe %proc%
        if (hwnd) {
            return hwnd
        }
    }

    DetectHiddenWindows, On
    WinGet, hwndList, List
    Loop, % hwndList {
        hwnd := hwndList%A_Index%
        WinGet, processName, ProcessName, ahk_id %hwnd%
        if (ErrorLevel || processName = "") {
            continue
        }

        StringLower, processNameLower, processName
        Loop, Parse, BRAVE_NIGHTLY_PROC_HINTS, |
        {
            candidateLower := A_LoopField
            StringLower, candidateLower, candidateLower
            if (processNameLower = candidateLower) {
                return hwnd
            }
        }

        WinGetTitle, title, ahk_id %hwnd%
        if (title && InStr(title, BRAVE_NIGHTLY_TITLE_HINT)) {
            return hwnd
        }
    }

    WinGet, hwnd, ID, %BRAVE_NIGHTLY_TITLE_HINT%
    return hwnd
}

FocusBraveNightlyOnSecondDesktop(pullToCurrent := 0) {
    global GetCurrentDesktopNumberProc, GetWindowDesktopNumberProc, GoToDesktopNumberProc, IsWindowOnCurrentVirtualDesktopProc, MoveWindowToDesktopNumberProc, BRAVE_NIGHTLY_RUN_HINT
    hwnd := FindBraveNightlyWindow()

    if (!hwnd) {
        Run, %BRAVE_NIGHTLY_RUN_HINT%
        Loop 120 {
            hwnd := FindBraveNightlyWindow()
            if (hwnd) {
                break
            }
            Sleep, 50
        }
    }
    if (!hwnd) {
        return
    }

    targetDesktop := 0
    braveDesktop := DllCall(GetWindowDesktopNumberProc, "Ptr", hwnd, "Int")
    if (pullToCurrent) {
        MoveToCurrentDesktop(hwnd)
    } else {
        if (braveDesktop != targetDesktop && braveDesktop >= 0) {
            DllCall(MoveWindowToDesktopNumberProc, "Ptr", hwnd, "Int", targetDesktop)
        }

        currentDesktop := DllCall(GetCurrentDesktopNumberProc, "Int")
        if (currentDesktop != targetDesktop) {
            DllCall(GoToDesktopNumberProc, "Int", targetDesktop)
        }
    }

    Loop 20 {
        if (DllCall(IsWindowOnCurrentVirtualDesktopProc, "Ptr", hwnd, "Int") = 1) {
            break
        }
        Sleep, 25
    }

    Loop 5 {
        WinShow, ahk_id %hwnd%
        WinActivate, ahk_id %hwnd%
        if (WinActive("ahk_id " . hwnd)) {
            break
        }
        Sleep, 40
    }
}

FindRustDeskWindow() {
    global RUSTDESK_PROC_HINTS, RUSTDESK_TITLE_HINTS

    Loop, Parse, RUSTDESK_PROC_HINTS, |
    {
        proc := A_LoopField
        if (proc = "") {
            continue
        }
        WinGet, hwnd, ID, ahk_exe %proc%
        if (hwnd) {
            return hwnd
        }
    }

    DetectHiddenWindows, On
    WinGet, hwndList, List
    Loop, % hwndList {
        hwnd := hwndList%A_Index%
        WinGet, processName, ProcessName, ahk_id %hwnd%
        if (ErrorLevel || processName = "") {
            continue
        }

        StringLower, processNameLower, processName
        Loop, Parse, RUSTDESK_PROC_HINTS, |
        {
            candidateLower := A_LoopField
            StringLower, candidateLower, candidateLower
            if (processNameLower = candidateLower) {
                return hwnd
            }
        }

        WinGetTitle, title, ahk_id %hwnd%
        if (title) {
            Loop, Parse, RUSTDESK_TITLE_HINTS, |
            {
                titleHint := A_LoopField
                if (titleHint && InStr(title, titleHint)) {
                    return hwnd
                }
            }
        }
    }

    oldMatchMode := A_TitleMatchMode
    SetTitleMatchMode, 2
    Loop, Parse, RUSTDESK_TITLE_HINTS, |
    {
        titleHint := A_LoopField
        if (titleHint = "") {
            continue
        }
        SetTitleMatchMode, 2
        WinGet, hwnd, ID, %titleHint%
        if (hwnd) {
            SetTitleMatchMode, %oldMatchMode%
            return hwnd
        }
    }
    SetTitleMatchMode, %oldMatchMode%
    return hwnd
}

FindChatGPTWindow() {
    global CHATGPT_PROC_HINTS, CHATGPT_TITLE_HINTS

    Loop, Parse, CHATGPT_PROC_HINTS, |
    {
        proc := A_LoopField
        if (proc = "") {
            continue
        }
        WinGet, hwnd, ID, ahk_exe %proc%
        if (hwnd) {
            return hwnd
        }
    }

    DetectHiddenWindows, On
    WinGet, hwndList, List
    Loop, % hwndList {
        hwnd := hwndList%A_Index%
        WinGet, processName, ProcessName, ahk_id %hwnd%
        if (ErrorLevel || processName = "") {
            continue
        }

        StringLower, processNameLower, processName
        Loop, Parse, CHATGPT_PROC_HINTS, |
        {
            candidateLower := A_LoopField
            StringLower, candidateLower, candidateLower
            if (processNameLower = candidateLower) {
                return hwnd
            }
        }

        WinGetTitle, title, ahk_id %hwnd%
        if (title) {
            Loop, Parse, CHATGPT_TITLE_HINTS, |
            {
                titleHint := A_LoopField
                if (titleHint && InStr(title, titleHint)) {
                    return hwnd
                }
            }
        }
    }

    oldMatchMode := A_TitleMatchMode
    SetTitleMatchMode, 2
    Loop, Parse, CHATGPT_TITLE_HINTS, |
    {
        titleHint := A_LoopField
        if (titleHint = "") {
            continue
        }
        WinGet, hwnd, ID, %titleHint%
        if (hwnd) {
            SetTitleMatchMode, %oldMatchMode%
            return hwnd
        }
    }
    SetTitleMatchMode, %oldMatchMode%
    return hwnd
}

LaunchRustDesk() {
    global RUSTDESK_RUN_HINTS
    Loop, Parse, RUSTDESK_RUN_HINTS, |
    {
        candidate := A_LoopField
        if (candidate = "") {
            continue
        }
        if (InStr(candidate, " ") && !RegExMatch(candidate, "^""")) {
            candidateRun := """" . candidate . """"
        } else {
            candidateRun := candidate
        }
        Run, %candidateRun%
        Sleep, 150
        return
    }
}

LaunchChatGPT() {
    global CHATGPT_RUN_HINTS
    Loop, Parse, CHATGPT_RUN_HINTS, |
    {
        candidate := A_LoopField
        if (candidate = "") {
            continue
        }
        if (FileExist(candidate)) {
            if (InStr(candidate, " ") && !RegExMatch(candidate, "^""")) {
                candidateRun := """" . candidate . """"
            } else {
                candidateRun := candidate
            }
            Run, %candidateRun%
            Sleep, 150
            return
        }
    }
    Run, chatgpt.exe
}

FocusRustDesk(pullToCurrent := 0) {
    global GetCurrentDesktopNumberProc, GetWindowDesktopNumberProc, GoToDesktopNumberProc, IsWindowOnCurrentVirtualDesktopProc
    hwnd := FindRustDeskWindow()

    if (!hwnd) {
        LaunchRustDesk()
        Loop 120 {
            hwnd := FindRustDeskWindow()
            if (hwnd) {
                break
            }
            Sleep, 50
        }
    }
    if (!hwnd) {
        return
    }

    targetDesktop := DllCall(GetWindowDesktopNumberProc, "Ptr", hwnd, "Int")
    if (pullToCurrent) {
        MoveToCurrentDesktop(hwnd)
    } else if (targetDesktop >= 0) {
        currentDesktop := DllCall(GetCurrentDesktopNumberProc, "Int")
        if (targetDesktop != currentDesktop) {
            DllCall(GoToDesktopNumberProc, "Int", targetDesktop)
        }
        Loop 20 {
            if (DllCall(IsWindowOnCurrentVirtualDesktopProc, "Ptr", hwnd, "Int") = 1) {
                break
            }
            Sleep, 25
        }
    }

    Loop 5 {
        WinShow, ahk_id %hwnd%
        WinActivate, ahk_id %hwnd%
        if (WinActive("ahk_id " . hwnd)) {
            break
        }
        Sleep, 40
    }
}

FocusChatGPT(pullToCurrent := 0) {
    global GetCurrentDesktopNumberProc, GetWindowDesktopNumberProc, GoToDesktopNumberProc, IsWindowOnCurrentVirtualDesktopProc
    hwnd := FindChatGPTWindow()

    if (!hwnd) {
        LaunchChatGPT()
        Loop 120 {
            hwnd := FindChatGPTWindow()
            if (hwnd) {
                break
            }
            Sleep, 50
        }
    }
    if (!hwnd) {
        return
    }

    targetDesktop := DllCall(GetWindowDesktopNumberProc, "Ptr", hwnd, "Int")
    if (pullToCurrent) {
        MoveToCurrentDesktop(hwnd)
    } else if (targetDesktop >= 0) {
        currentDesktop := DllCall(GetCurrentDesktopNumberProc, "Int")
        if (targetDesktop != currentDesktop) {
            DllCall(GoToDesktopNumberProc, "Int", targetDesktop)
        }
        Loop 20 {
            if (DllCall(IsWindowOnCurrentVirtualDesktopProc, "Ptr", hwnd, "Int") = 1) {
                break
            }
            Sleep, 25
        }
    }

    Loop 5 {
        WinShow, ahk_id %hwnd%
        WinActivate, ahk_id %hwnd%
        if (WinActive("ahk_id " . hwnd)) {
            break
        }
        Sleep, 40
    }
    return hwnd
}

FindDiscordWindow() {
    global DISCORD_PROC_HINTS, DISCORD_TITLE_HINTS

    Loop, Parse, DISCORD_PROC_HINTS, |
    {
        proc := A_LoopField
        if (proc = "") {
            continue
        }
        WinGet, hwnd, ID, ahk_exe %proc%
        if (hwnd) {
            return hwnd
        }
    }

    DetectHiddenWindows, On
    WinGet, hwndList, List
    Loop, % hwndList {
        hwnd := hwndList%A_Index%
        WinGet, processName, ProcessName, ahk_id %hwnd%
        if (ErrorLevel || processName = "") {
            continue
        }

        StringLower, processNameLower, processName
        Loop, Parse, DISCORD_PROC_HINTS, |
        {
            candidateLower := A_LoopField
            StringLower, candidateLower, candidateLower
            if (processNameLower = candidateLower) {
                return hwnd
            }
        }

        WinGetTitle, title, ahk_id %hwnd%
        if (title) {
            titleLower := title
            StringLower, titleLower, titleLower
            Loop, Parse, DISCORD_TITLE_HINTS, |
            {
                titleHint := A_LoopField
                if (titleHint = "") {
                    continue
                }
                titleHintLower := titleHint
                StringLower, titleHintLower, titleHintLower
                if (InStr(titleLower, titleHintLower)) {
                    return hwnd
                }
            }
        }
    }

    oldMatchMode := A_TitleMatchMode
    SetTitleMatchMode, 2
    Loop, Parse, DISCORD_TITLE_HINTS, |
    {
        titleHint := A_LoopField
        if (titleHint = "") {
            continue
        }
        WinGet, hwnd, ID, %titleHint%
        if (hwnd) {
            SetTitleMatchMode, %oldMatchMode%
            return hwnd
        }
    }
    SetTitleMatchMode, %oldMatchMode%
    return 0
}

LaunchDiscord() {
    global DISCORD_RUN_HINTS
    Loop, Parse, DISCORD_RUN_HINTS, |
    {
        candidate := A_LoopField
        if (candidate = "") {
            continue
        }
        if (FileExist(candidate)) {
            candidateRun := candidate
            if (InStr(candidateRun, " ") && !RegExMatch(candidateRun, "^""")) {
                candidateRun := """" . candidateRun . """"
            }
            candidateLower := candidate
            StringLower, candidateLower, candidateLower
            if (InStr(candidateLower, "update.exe")) {
                Run, % candidateRun . " --processStart Discord.exe"
            } else {
                Run, % candidateRun
            }
            Sleep, 200
            return
        }
    }

    Loop, Files, % A_LocalAppData . "\Discord\app-*\Discord.exe", F
    {
        candidateRun := A_LoopFileFullPath
        if (candidateRun != "") {
            if (InStr(candidateRun, " ") && !RegExMatch(candidateRun, "^""")) {
                candidateRun := """" . candidateRun . """"
            }
            Run, % candidateRun
            Sleep, 200
            return
        }
    }

    Run, discord.exe
}

FocusDiscord(pullToCurrent := 0) {
    global GetCurrentDesktopNumberProc, GetWindowDesktopNumberProc, GoToDesktopNumberProc, IsWindowOnCurrentVirtualDesktopProc
    hwnd := FindDiscordWindow()

    if (!hwnd) {
        LaunchDiscord()
        Loop 120 {
            hwnd := FindDiscordWindow()
            if (hwnd) {
                break
            }
            Sleep, 50
        }
    }

    if (!hwnd) {
        Run, discord://
        Loop 120 {
            hwnd := FindDiscordWindow()
            if (hwnd) {
                break
            }
            Sleep, 50
        }
    }

    if (!hwnd) {
        return
    }

    targetDesktop := DllCall(GetWindowDesktopNumberProc, "Ptr", hwnd, "Int")
    if (pullToCurrent) {
        MoveToCurrentDesktop(hwnd)
    } else if (targetDesktop >= 0) {
        currentDesktop := DllCall(GetCurrentDesktopNumberProc, "Int")
        if (targetDesktop != currentDesktop) {
            DllCall(GoToDesktopNumberProc, "Int", targetDesktop)
        }
        Loop 20 {
            if (DllCall(IsWindowOnCurrentVirtualDesktopProc, "Ptr", hwnd, "Int") = 1) {
                break
            }
            Sleep, 25
        }
    }

    Loop 5 {
        WinShow, ahk_id %hwnd%
        WinActivate, ahk_id %hwnd%
        if (WinActive("ahk_id " . hwnd)) {
            break
        }
        Sleep, 40
    }
}

FindNeovimWindow() {
    global NEOVIM_PROC_HINTS, NEOVIM_TITLE_HINTS

    Loop, Parse, NEOVIM_PROC_HINTS, |
    {
        proc := A_LoopField
        if (proc = "") {
            continue
        }
        WinGet, hwnd, ID, ahk_exe %proc%
        if (hwnd) {
            return hwnd
        }
    }

    DetectHiddenWindows, On
    WinGet, hwndList, List
    Loop, % hwndList {
        hwnd := hwndList%A_Index%
        WinGet, processName, ProcessName, ahk_id %hwnd%
        if (ErrorLevel || processName = "") {
            continue
        }

        StringLower, processNameLower, processName
        Loop, Parse, NEOVIM_PROC_HINTS, |
        {
            candidateLower := A_LoopField
            StringLower, candidateLower, candidateLower
            if (processNameLower = candidateLower) {
                return hwnd
            }
        }

        WinGetTitle, title, ahk_id %hwnd%
        if (title) {
            Loop, Parse, NEOVIM_TITLE_HINTS, |
            {
                titleHint := A_LoopField
                if (titleHint && InStr(title, titleHint)) {
                    return hwnd
                }
            }
        }
    }

    oldMatchMode := A_TitleMatchMode
    SetTitleMatchMode, 2
    Loop, Parse, NEOVIM_TITLE_HINTS, |
    {
        titleHint := A_LoopField
        if (titleHint = "") {
            continue
        }
        WinGet, hwnd, ID, %titleHint%
        if (hwnd) {
            SetTitleMatchMode, %oldMatchMode%
            return hwnd
        }
    }
    SetTitleMatchMode, %oldMatchMode%
    return hwnd
}

IsActiveNeovimWindow() {
    global NEOVIM_PROC_HINTS, NEOVIM_TITLE_HINTS

    hwnd := WinExist("A")
    if (!hwnd) {
        return false
    }

    WinGet, processName, ProcessName, ahk_id %hwnd%
    if (!ErrorLevel && processName != "") {
        StringLower, processNameLower, processName
        Loop, Parse, NEOVIM_PROC_HINTS, |
        {
            candidateProc := A_LoopField
            if (candidateProc = "") {
                continue
            }

            StringLower, candidateProcLower, candidateProc
            if (processNameLower = candidateProcLower) {
                return true
            }
        }
    }

    WinGetTitle, title, ahk_id %hwnd%
    if (!ErrorLevel && title != "") {
        Loop, Parse, NEOVIM_TITLE_HINTS, |
        {
            titleHint := A_LoopField
            if (titleHint && InStr(title, titleHint)) {
                return true
            }
        }
    }

    return false
}

LaunchNeovim() {
    global NEOVIM_RUN_HINTS
    Loop, Parse, NEOVIM_RUN_HINTS, |
    {
        candidate := A_LoopField
        if (candidate = "") {
            continue
        }
        if (FileExist(candidate)) {
            if (InStr(candidate, " ") && !RegExMatch(candidate, "^""")) {
                candidateRun := """" . candidate . """"
            } else {
                candidateRun := candidate
            }
            Run, %candidateRun%
            Sleep, 150
            return
        }
    }
    Run, nvim-qt.exe
}

FocusNeovim(pullToCurrent := 0) {
    global GetCurrentDesktopNumberProc, GetWindowDesktopNumberProc, GoToDesktopNumberProc, IsWindowOnCurrentVirtualDesktopProc
    hwnd := FindNeovimWindow()

    if (!hwnd) {
        LaunchNeovim()
        Loop 120 {
            hwnd := FindNeovimWindow()
            if (hwnd) {
                break
            }
            Sleep, 50
        }
    }
    if (!hwnd) {
        return
    }

    targetDesktop := DllCall(GetWindowDesktopNumberProc, "Ptr", hwnd, "Int")
    if (pullToCurrent) {
        MoveToCurrentDesktop(hwnd)
    } else if (targetDesktop >= 0) {
        currentDesktop := DllCall(GetCurrentDesktopNumberProc, "Int")
        if (targetDesktop != currentDesktop) {
            DllCall(GoToDesktopNumberProc, "Int", targetDesktop)
        }
        Loop 20 {
            if (DllCall(IsWindowOnCurrentVirtualDesktopProc, "Ptr", hwnd, "Int") = 1) {
                break
            }
            Sleep, 25
        }
    }

    Loop 5 {
        WinShow, ahk_id %hwnd%
        WinActivate, ahk_id %hwnd%
        if (WinActive("ahk_id " . hwnd)) {
            break
        }
        Sleep, 40
    }
}

GetActiveExplorerPath() {
    WinGet, activeHwnd, ID, A
    if (!activeHwnd) {
        return ""
    }

    WinGetClass, activeClass, ahk_id %activeHwnd%
    if (activeClass != "CabinetWClass" && activeClass != "ExploreWClass") {
        return ""
    }

    shellWindows := ComObjCreate("Shell.Application").Windows
    for window in shellWindows {
        try {
            if (window.HWND = activeHwnd) {
                return window.Document.Folder.Self.Path
            }
        } catch e {
            continue
        }
    }

    return ""
}

OpenPowerShellHere() {
    explorerPath := GetActiveExplorerPath()

    if (explorerPath = "") {
        Run, powershell.exe
        return
    }

    Run, powershell.exe -NoExit -NoProfile, %explorerPath%
}

OpenPhoneLink() {
    Run, %PHONE_LINK_RUN_HINT%
}

OpenRaSim() {
    global RA_SIM_RUN_PATH

    if (!FileExist(RA_SIM_RUN_PATH)) {
        MsgBox, 16, quick-desktop-hotkeys, RA sim launcher was not found at:`n%RA_SIM_RUN_PATH%
        return
    }

    Run, % """" . RA_SIM_RUN_PATH . """"
}

EnsureStartupShortcutExists() {
    global StartupShortcutPath

    if (FileExist(StartupShortcutPath)) {
        return
    }

    startupDir := A_AppData . "\Microsoft\Windows\Start Menu\Programs\Startup"
    if (!FileExist(startupDir)) {
        FileCreateDir, %startupDir%
    }

    shell := ComObjCreate("WScript.Shell")
    if (!IsObject(shell)) {
        return
    }

    shortcut := shell.CreateShortcut(StartupShortcutPath)
    if (!IsObject(shortcut)) {
        return
    }

    if (A_AhkPath) {
        shortcut.TargetPath := A_AhkPath
        shortcut.Arguments := """" . A_ScriptFullPath . """"
    } else {
        shortcut.TargetPath := A_ScriptFullPath
    }
    shortcut.WorkingDirectory := A_ScriptDir
    shortcut.Description := "quick-desktop-hotkeys"
    shortcut.Save()
}

OpenConvertSite() {
    global BRAVE_NIGHTLY_RUN_HINT, CONVERT_SITE_URL

    Run, % BRAVE_NIGHTLY_RUN_HINT . " --new-window """ . CONVERT_SITE_URL . """"
    if (ErrorLevel) {
        Run, % CONVERT_SITE_URL
    }
}

OpenYouTubeMusic() {
    global BRAVE_NIGHTLY_RUN_HINT, YOUTUBE_MUSIC_SHORTCUT, YOUTUBE_MUSIC_URL

    Run, % """" . YOUTUBE_MUSIC_SHORTCUT . """"
    if (ErrorLevel) {
        Run, % BRAVE_NIGHTLY_RUN_HINT . " --new-window """ . YOUTUBE_MUSIC_URL . """"
    }
}

TurnOffMonitors() {
    ; Launch the built-in black screen saver directly.
    screensaverPath := A_WinDir . "\System32\scrnsave.scr"
    if (!FileExist(screensaverPath)) {
        screensaverPath := A_WinDir . "\SysWOW64\scrnsave.scr"
    }
    if (FileExist(screensaverPath)) {
        Run, % screensaverPath . " /S"
        return
    }

    ; Fallback to the configured system screensaver.
    PostMessage, 0x0112, 0xF140, 0,, ahk_id 0xFFFF
}

SetHotkeysEnabled(state) {
    global HotkeysEnabled
    if (state = "toggle") {
        HotkeysEnabled := !HotkeysEnabled
    } else if (state = 0) {
        HotkeysEnabled := 0
    } else {
        HotkeysEnabled := 1
    }

    if (HotkeysEnabled) {
        ToolTip, quick-desktop-hotkeys enabled
    } else {
        ToolTip, quick-desktop-hotkeys disabled
    }
    SetTimer, __ClearHotkeyStateTooltip, -1200
}

__ClearHotkeyStateTooltip:
    ToolTip
return
__CtrlD_PageUp:
    global PendingCtrlDPress
    if (PendingCtrlDPress) {
        PendingCtrlDPress := 0
        SendPlainPageUp()
    }
    return
^!F10::
    SetHotkeysEnabled(0)
    return

^!F11::
    SetHotkeysEnabled(1)
    return

^!F12::
    SetHotkeysEnabled("toggle")
    return

#If HotkeysEnabled
^1::GoToDesktopIfExists(0)
^2::GoToDesktopIfExists(1)
^3::GoToDesktopIfExists(2)
^4::GoToDesktopIfExists(3)
^5::GoToDesktopIfExists(4)
^6::GoToDesktopIfExists(5)
^7::GoToDesktopIfExists(6)
^8::GoToDesktopIfExists(7)
^9::GoToDesktopIfExists(8)
^0::GoToDesktopIfExists(9)
#1::GoToDesktopIfExists(0)
#2::GoToDesktopIfExists(1)
#3::GoToDesktopIfExists(2)
#4::GoToDesktopIfExists(3)
#5::GoToDesktopIfExists(4)
#6::GoToDesktopIfExists(5)
#7::GoToDesktopIfExists(6)
#8::GoToDesktopIfExists(7)
#9::GoToDesktopIfExists(8)
#0::GoToDesktopIfExists(9)
#c::FocusCodex(0)
^#c::FocusCodex(1)
!-::
    SendInput, {U+2014}
    return
LAlt::
RAlt::
    return
LAlt & Tab::
RAlt & Tab::
    if GetKeyState("Shift", "P") {
        SendInput, {Alt down}{Shift down}{Tab}
        return
    }
    SendInput, {Alt down}{Tab}
    return
; Alt-based helpers outside Neovim.
#If (HotkeysEnabled && !IsActiveNeovimWindow())
!u::
    if (GetKeyState("Shift", "P")) {
        SendInput, ^z
        return
    }
    SendInput, {PgUp 4}
    return
!d::
    SendInput, {PgDn 4}
    return
#If HotkeysEnabled
#InputLevel 100
$^d::
    global PendingCtrlDPress
    if (PendingCtrlDPress) {
        PendingCtrlDPress := 0
        SetTimer, __CtrlD_PageUp, Off
        SendPlainPageDown()
        return
    }
    PendingCtrlDPress := 1
    SetTimer, __CtrlD_PageUp, -250
    return
#InputLevel 0
#m::FocusOutlook(0)
^#m::FocusOutlook(1)
#b::FocusBraveNightlyOnSecondDesktop(0)
^#b::FocusBraveNightlyOnSecondDesktop(1)
#w::FocusRustDesk(0)
^#w::FocusRustDesk(1)
#o::OpenPhoneLink()
!g::OpenRaSim()
#If (HotkeysEnabled && !GetKeyState("Alt", "P") && !GetKeyState("Ctrl", "P"))
#x::FocusChatGPT(0)
#z::OpenConvertSite()
#y::OpenYouTubeMusic()
#If HotkeysEnabled
!x::OpenChatGPTWithSelectionOrClipboard()
#!x::OpenChatGPTWithSelectionOrClipboard()
^#x::
    OpenChatGPTWithSelectionOrClipboard()
    return
#+x::FocusChatGPT(1)
#InputLevel 100
#f::
FocusDiscord(0)
return

^#f::
FocusDiscord(1)
return
#InputLevel 0
#n::FocusNeovim(0)
^#n::FocusNeovim(1)
#s::OpenPowerShellHere()
F12::TurnOffMonitors()
^!q::ExitApp
#If (HotkeysEnabled && GetKeyState("Ctrl", "P") && GetKeyState("Shift", "P") && !GetKeyState("MButton", "P"))
; Ctrl+Shift+wheel moves the active window between virtual desktops.
*WheelUp::
    MoveActiveWindowToRightDesktop()
    return
*WheelDown::
    MoveActiveWindowToLeftDesktop()
    return
#If (HotkeysEnabled && GetKeyState("Shift", "P") && !GetKeyState("Ctrl", "P") && !GetKeyState("MButton", "P"))
; Shift+wheel cycles tabs in supported apps, otherwise it falls back to plain scrolling.
*WheelUp::
    HandleShiftWheel("up")
    return
*WheelDown::
    HandleShiftWheel("down")
    return
#If HotkeysEnabled
MButton & WheelUp::
    if (GetKeyState("Ctrl", "P") && GetKeyState("Shift", "P")) {
        MoveActiveWindowToRightDesktop()
        return
    }
    if (GetKeyState("Ctrl", "P") && IsActiveTabScrollWindow()) {
        SendInput, ^+{Tab}
        return
    }
    GoToNextDesktop()
    return
MButton & WheelDown::
    if (GetKeyState("Ctrl", "P") && GetKeyState("Shift", "P")) {
        MoveActiveWindowToLeftDesktop()
        return
    }
    if (GetKeyState("Ctrl", "P") && IsActiveTabScrollWindow()) {
        SendInput, ^{Tab}
        return
    }
    GoToPrevDesktop()
    return
^!r::
    ReloadQuickDesktopScript()
    return

ReloadQuickDesktopScript() {
    global RestartShortcutPath, FallbackScriptPath

    scriptToRun := FallbackScriptPath
    if (FileExist(RestartShortcutPath)) {
        scriptToRun := RestartShortcutPath
    }

    Run, %scriptToRun%
    ExitApp
}
