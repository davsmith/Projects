#Requires AutoHotkey v2.0
::MI2::
{
    Send "MI^+=2^+= "
}

^!i::
{
    SendMode("event")
    SetKeyDelay 1000
    Send "+{F10}"
    ; Send "g"
}

