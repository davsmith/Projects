/*
    Basics_v2.ahk

    Examples of the fundamental concepts of the scripting language for AutoHotkey.

    Based on the AutoHotkey notes and tutorials in OneNote (https://tinyurl.com/5n7mzhc5)

    Dave Smith
    2/5/2023

*/
#Requires AutoHotkey v2.0

/*
    Comments

    Single line comments can be demarcated with a ;

    Multi-line comments start with
*/

/*
    Variables
*/
//#region
; Variable type is set on assignment
pi := "Set my value to a string"
OutputDebug type(pi) . '`n'

; Variable type is dynamic
pi := 3.14
OutputDebug type(pi) . '`n'

; Variable names are not case sensitive
OutputDebug "PI " . PI . " is the same as pi " . pi . "`n"
//#endregion

/*
    Hotkeys
*/
^!o::
{
    if WinExist("ahk_class OneNote.exe")
    {
        WinActivate ; Use the window found by WinExist.
    } else {
        Run "onenote.exe"
    }
}

^LButton::
{
    MsgBox("You pressed the left button")
    OutputDebug("Hello`n") 
}

; Run Notepad CTRL+ALT+N is pressed
^!n::
{
    Run "notepad.exe"
    return
}

; Copies currently selected text and wraps it in formatting tags
<^b::
{
    Send "{Ctrl down}c{Ctrl up}"
    SendInput "[b]{Ctrl down}v{Ctrl up}[/b]"
    return
}

; Dynamically define/undefine a hotkey
<^!d::
{
    Hotkey "^!z", MyFunc, "On"
    MsgBox "Hotkey defined"
}

>^!d::
{
    Hotkey "^!z", "Off"
    MsgBox "Hotkey undefined"
}

MyFunc(ThisHotkey)
{
    MsgBox "You pressed " ThisHotkey
}

/*
    Hotstrings
*/
; Replaces "btw" with "by the way" when ending character is typed
::btw::by the way

; Replaces "idk" with "I don't know" without requiring an ending character.
:*:idk::I don't know

; Runs a code block
::mujiber::
{
    MsgBox "You typed Mujiber"
}

; Case sensitive, and don't replace the string
:CB0:Serajoul::
{
    OutputDebug "You typed a proper name"
}

; Scope the hotkeys/hotstrings to an app
MyWindowTitle := "Basics"
#HotIf WinActive("ahk_class Notepad") or WinActive(MyWindowTitle) or WinActive("ahk_exe OneNote.exe")
    #Space::MsgBox "You pressed Win+Spacebar in Notepad, OneNote or " MyWindowTitle

    ^!s::
    {
        MsgBox(StatusBarGetText(2,"A"))
    }

    ^!t::
    {
        WinSetTransparent(128)
    }


/*
*/