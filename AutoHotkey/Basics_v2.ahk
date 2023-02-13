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
; Variable type is set on assignment
pi := "Set my value to a string"
OutputDebug type(pi) . '`n'

; Variable type is dynamic
pi := 3.14
OutputDebug type(pi) . '`n'

; Variable names are not case sensitive
OutputDebug "PI " . PI . " is the same as pi " . pi . "`n"


/*
    Hotkeys
*/
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

; Copies currently selected text
^b::
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

/*
*/