#Requires AutoHotkey v2.0

#+F6::  ; Win + Shift + F6
{
    ; Step 1: Copy active cell in Excel and paste in Edge with single Tab
    try {
        excel := ComObjActive("Excel.Application")
        A_Clipboard := excel.ActiveCell.Text  ; Copy current cell's text
        ClipWait(1)  ; Wait for clipboard to contain data
    } catch {
        MsgBox "Excel not found or no active cell."
        return
    }

    if !WinExist("ahk_exe msedge.exe") {
        Run "msedge.exe"
        WinWaitActive("ahk_exe msedge.exe",, 5)  ; Wait up to 5 seconds for Edge
    } else {
        WinActivate("ahk_exe msedge.exe")  ; Switch to Edge
        WinWaitActive("ahk_exe msedge.exe",, 2)  ; Ensure Edge is active
    }

    Send "^v"  ; Paste clipboard content
    Send "{Tab}"  ; Press Tab once
    Sleep 100  ; Brief pause to ensure stability

    ; Step 2: Switch to Excel, move right, copy cell, switch to Edge, paste, press Tab
    WinActivate("ahk_exe excel.exe")  ; Switch back to Excel
    WinWaitActive("ahk_exe excel.exe",, 2)  ; Ensure Excel is active
    excel.ActiveCell.Offset(0, 1).Select  ; Move one cell to the right
    A_Clipboard := excel.ActiveCell.Text  ; Copy new cell's text
    ClipWait(1)

    WinActivate("ahk_exe msedge.exe")  ; Switch to Edge
    WinWaitActive("ahk_exe msedge.exe",, 2)
    Send "^v"  ; Paste
    Send "{Tab}"  ; Press Tab once
    Sleep 100

    ; Step 3: Switch to Excel, move right, copy cell, switch to Edge, paste, press Tab twice
    WinActivate("ahk_exe excel.exe")
    WinWaitActive("ahk_exe excel.exe",, 2)
    excel.ActiveCell.Offset(0, 1).Select  ; Move one cell to the right
    A_Clipboard := excel.ActiveCell.Text  ; Copy new cell's text
    ClipWait(1)

    WinActivate("ahk_exe msedge.exe")
    WinWaitActive("ahk_exe msedge.exe",, 2)
    Send "^v"  ; Paste
    Send "{Tab 2}"  ; Press Tab twice
    Sleep 100

  ; Step 4: Switch to Excel, move right, copy cell, switch to Edge, paste, press Enter 3 times, wait 8 seconds
WinActivate("ahk_exe excel.exe")
WinWaitActive("ahk_exe excel.exe",, 2)
excel.ActiveCell.Offset(0, 1).Select  ; Move one cell to the right
A_Clipboard := excel.ActiveCell.Text  ; Copy new cell's text
ClipWait(1)

WinActivate("ahk_exe msedge.exe")
WinWaitActive("ahk_exe msedge.exe",, 2)
Send "^v"  ; Paste
Send "{Enter 3}"  ; Press Enter three times
Sleep 3000  ; Wait 2 seconds
Send "{Enter}"
Sleep 11000

; Step 5: Repeat Step 2 - Switch to Excel, move right, copy cell, switch to Edge, paste, press Tab
WinActivate("ahk_exe excel.exe")
WinWaitActive("ahk_exe excel.exe",, 2)
excel.ActiveCell.Offset(0, 1).Select  ; Move one cell to the right
A_Clipboard := excel.ActiveCell.Text  ; Copy new cell's text
ClipWait(1)

WinActivate("ahk_exe msedge.exe")
WinWaitActive("ahk_exe msedge.exe",, 2)
Send "^v"  ; Paste
Send "{Tab}"  ; Press Tab once
Sleep 100

; Step 6: Repeat Step 2, then in Excel move down 1 cell, left 5 cells
WinActivate("ahk_exe excel.exe")
WinWaitActive("ahk_exe excel.exe",, 2)
excel.ActiveCell.Offset(0, 1).Select  ; Move one cell to the right
A_Clipboard := excel.ActiveCell.Text  ; Copy new cell's text
ClipWait(1)

WinActivate("ahk_exe msedge.exe")
WinWaitActive("ahk_exe msedge.exe",, 2)
Send "^v"  ; Paste
Send "{Tab}"  ; Press Tab once
Sleep 100

WinActivate("ahk_exe excel.exe")
WinWaitActive("ahk_exe excel.exe",, 2)
excel.ActiveCell.Offset(1, -5).Select  ; Move down 1 row, left 5 columns
 
}