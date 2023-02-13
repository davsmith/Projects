/*
    Basics_v2.ahk

    Examples of the fundamental concepts of the scripting language for AutoHotkey.

    Based on the AutoHotkey notes and tutorials in OneNote (https://tinyurl.com/at5dwsw3)

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
pi := "Set my value to a string"
pi := 3.14
OutputDebug "PI " . PI . " is the same as pi " . pi

^LButton::
{
    MsgBox("You pressed the left button")
    OutputDebug("Hello") 
}       