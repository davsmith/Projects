/*
    controls.ahk

    Examples of operations on controls within an application from AutoHotkey.

    Based on the AutoHotkey notes and tutorials in OneNote (https://tinyurl.com/5n7mzhc5)

    Dave Smith
    2/15/2023

*/
#Requires AutoHotkey v2.0

^!a::
{
    ControlClick "x130 y564", "- OneNote"
    OutputDebug("Clicked")
    return
}