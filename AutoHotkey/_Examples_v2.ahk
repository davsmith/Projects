/*
    Examples_v2.ahk

    Examples of the fundamental concepts of the scripting language for AutoHotkey.

    Based on the AutoHotkey notes and tutorials in OneNote (https://tinyurl.com/5n7mzhc5)

    Dave Smith
    2/21/2023

*/
#Requires AutoHotkey v2.0

/*
    Hotkeys
*/
^!o:: ; Check for an instance of OneNote.exe
{
    if WinExist("ahk_exe OneNote.exe")
    {
        OutputDebug("Found an instance of OneNote.  Activating it.`n")
        WinActivate ; Use the window found by WinExist.
    } else {
        OutputDebug("Didn't find OneNote.  Launching it.`n")
        Run "onenote.exe"
    }
}

^LButton:: ; Left mouse button uses MsgBox and OutDebug functions
{
    OutputDebug("Left mouse button pressed`n") 
    MsgBox("You pressed the left button")
}


<^b:: ; Copies selected text and wraps it in formatting tags, using key up/down
{
    OutputDebug("Left <ctrl-b>`n") 
    Send "{Ctrl down}c{Ctrl up}"
    SendInput "[b]{Ctrl down}v{Ctrl up}[/b]"
}

<!^d:: ; Dynamically define/undefine a hotkey
{
    Hotkey "^!z", printKey, "On"
    MsgBox "<ctrl><alt>z defined"
}

>!^d::
{
    Hotkey "^!z", "Off"
    MsgBox "<ctrl><alt>z undefined"
}

printKey(ThisHotkey)
{
    MsgBox "You pressed " ThisHotkey
}


/*
    Hotstrings
*/
; Replaces "btw" with "by the way" when an ending character is typed
::btw::by the way

; Replaces "idk" with "I don't know" without requiring an ending character.
:*:idk::I don't know

; Runs a code block
::mujiber::
{
    MsgBox "You typed Mujiber or mujiber"
}

; Case sensitive, and don't replace the string
:CB0:Serajoul::
{
    OutputDebug "You typed a properly capitalized name"
}

; Scope the hotkeys/hotstrings to an app
MyWindowTitle := "Basics"
#HotIf WinActive("ahk_class Notepad") or WinActive("ahk_exe OneNote.exe") or WinActive(MyWindowTitle) 
    #Space::MsgBox "You pressed Win+Spacebar in Notepad, OneNote or " MyWindowTitle

    ; Print the second section of the status bar in the current window
    ^!s::
    {
        MsgBox(StatusBarGetText(2,"A"))
    }

    ; Set the current window to 1/2 transparency (255 is opaque)
    ^!t::
    {
        WinSetTransparent(128)
    }