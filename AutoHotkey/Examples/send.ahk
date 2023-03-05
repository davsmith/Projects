/*
    send.ahk

    Examples of using Send within an application from AutoHotkey.

    Based on the AutoHotkey notes and tutorials in OneNote (https://tinyurl.com/32hd4dxn)

    Dave Smith
    2/25/2023

*/

#HotIf WinActive("ahk_exe OneNote.exe")
; Example workaround if Send is not reliable
F2::
{
    previous_mode := SendMode("Event")
    SetKeyDelay 100
    Send ("^k")
    Send ("!e")
    Send ("^c")
    OutputDebug("Copied: " . a_clipboard . "`n")
    SendMode(previous_mode)
}

#HotIf
; Send the key to bring up the context menu
F3::
{
    SendMode "Event"
    SetKeyDelay 100
    Send("{AppsKey}")
}

