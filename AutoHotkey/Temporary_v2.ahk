#Requires AutoHotkey v2.0
#SingleInstance
; #include "c:\temp\"
; #include "templib.ahk"
; #include "templib2.ahk"

::_r::
{
    Send "^s"
    Reload
}

; ::btw::Hello There
::btw:: Hello There
; MsgBox (A_ScriptDir)
; MsgBox (A_MyDocuments)
str := "Hello there"
; MsgBox (str)

#HotIf WinActive("ahk_exe ONENOTE.EXE")
^!c::
{
    Send "!hff"
    Send "consolas{enter}"
    Send "!hfs"
    Send "10.5{enter}"
}

MButton::
{
    ; Click 2
    ; Send "{AppsKey}g"


}

^!n::
{
    SetKeyDelay(200)
    SendMode "Event"
    Send "!hl"
    Send "{end}"
    Send "{enter}"
}




#HotIf WinActive("ahk_exe Code.EXE")
^!e::
{
    person_name := "Fred"
    morning_greeting := "Good morning, " person_name ". How are you?"
    MsgBox morning_greeting
}

^!t::
{
    str1 := "Hello"
    str1 .= " There"
    MsgBox str1
}

^!v::
{
    MsgBox A_AhkVersion
}


#HotIf WinActive("ahk_exe powerpnt.exe")
^!t::
{
    SendMode("event")
    SetKeyDelay(200)
    OutputDebug("Aligning top")
    MouseMove(1425, 115)
    Click
    Send("t")
}

^!b::
{
    ; In Office use lower case letters with Send
    OutputDebug("Starting align top`n")
    Send("!jd")
    Send("aat")
    OutputDebug("Ending align top`n")
}


#HotIf
^j::
{
    msgString := "This is " 2*3 " a test"
    MsgBox msgString
}

^!m::
{
    ; Moves the mouse and clicks at the specified coordinates
    ; relative to the client (see Window Spy)
    Click(0, 0)
}

^!r::
{
    ; Use the keyword "Relative" to make coordinates
    ; relative to the current mouse position
    Click(-20, -20, "Relative")
}

::_h2::
{
    Send "Hello`n"
}

::_tt::
{
    Loop
    {
        Sleep 100
        MouseGetPos(&xpos, &ypos, &WhichWindow, &WhichControl)
        try ControlGetPos(&x, &y, &w, &h, WhichControl, WhichWindow)
        ToolTip(WhichControl "`nX" x "`tY" y "`nW" w "`t" h)
    }
}
