#Requires AutoHotkey v2.0

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

^!f::
{
    OutputDebug("Send mode: " A_SendMode)
    SendMode "Event"
    SetKeyDelay 100
    Send ("^k")
    Send ("!e")
    Send ("^c")
    OutputDebug("Source: " . a_clipboard . "`n")
    v2_string := StrReplace(a_clipboard, "v1", "v2")
    OutputDebug("Target: " . v2_string . "`n")
    a_clipboard := v2_string
    Send ("^v")
    OutputDebug("Pasted " . a_clipboard . "`n")
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



#HotIf
^j::
{
    msgString := "This is " 2*3 " a test"
    MsgBox msgString
}

 

