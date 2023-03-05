/*
    Basics_v2.ahk

    Examples of the fundamental concepts of the scripting language for AutoHotkey.

    Based on the AutoHotkey notes and tutorials in OneNote (https://tinyurl.com/5n7mzhc5)

    Dave Smith
    2/5/2023

*/
#Requires AutoHotkey v2.0

::_comments::
{
    /*
    Comments

    Single line comments can be demarcated with a ;

    Multi-line comments start with
    */
}

::_variables::
{
    ; Variables do not need to be declared and are dynamically typed

    ; Variable type is set on assignment
    pi := "Set my value to a string"
    OutputDebug type(pi) . '`n'

    ; Variable type is dynamic
    pi := 3.14
    OutputDebug type(pi) . '`n'

    ; Variable names are not case sensitive
    OutputDebug "PI " . PI . " is the same as pi " . pi . "`n"
}

::_hotkeys::
{
    /*
        Hotkeys definitions start with key combination (including modifiers)
        followed by ::, and a code block (enclosed in {} )

        The examples below are commented out since hotkey definitions can't be
        included within other code blocks (the _hotkeys section).  Copy them
        outside the code block to make them active
    
    ^!o::
    {
        MsgBox("<ctrl><alt>o")
    }

    ^LButton::
    {
        MsgBox("<ctrl><Left mouse button>")
    }

    <^b::
    {
        MsgBox("<left ctrl>b")
    }

    >^b::
    {
        MsgBox("<right ctrl>b")
    }

    >+r::
    {
        MsgBox("<right shift>r")
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

    */
}

::_hotstrings::
{
    /*
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
    */
}

::_conditionals::
{
    ; True and False are defined as constants for readability
    ; Variables/constants are not case sensitive
    if (true)
    {
        OutputDebug("True is defined as " . true . "`n")
    }

    if (not false)
    {
        OutputDebug("False is defined as " . false . "`n")
    }

    ; Empty strings and 0 are considered False
    ; Everything else (including objects) are True
    if (not "")
    {
        OutputDebug("An emptry string is False`n")
    }

    ; The = is used for comparison
    ; if, else, and else if can be used for multiple compares
    color := "Blue"
    ; color := "Silver"
    ; color := "Pink"

    if (color = "Blue" or color = "White")
    {
        OutputDebug color . " is one of the allowed values.`n"
        ExitApp
    }
    else if (color = "Silver")
    {
        OutputDebug "Silver is not an allowed color.`n"
        return
    }
    else
    {
        OutputDebug color . " is not recognized.`n"
        ExitApp
    }
}

::_loops::
{
    ; The loop statement is used for traditional For...Next operations
    ; Rather than providing an index variable A_Index is used
    loop 10
    {
        OutputDebug(a_index '`n')
    }

    ; As of Feb 22, 2023 I can't find a cleaner method for nested loops than...
    loop 2
    {
        i := a_index
        loop 3
        {
            j := a_index
            OutputDebug('(' i ',' j ')`n')
        }
    }

    ; For is  used to enumerate key/value pairs (foreach)
    colours := {red: 0xFF0000, blue: 0x0000FF, green: 0x00FF00}
    for k, v in colours.OwnProps()
        s .= k '=' v '`n'
    
    OutputDebug(s . "`n")
}

::_strings::
{
    ; Use `n to indicate a newline
    OutputDebug("Hello`nthere`nBob`n")

    ; .= is shorthand to append to a string
    greeting := "Hello "
    greeting .= "Bob`n"
    OutputDebug(greeting)

    ; Use the Format function to manipulate strings
    ;
    ; Full details about Format can be found at:
    ; https://www.autohotkey.com/docs/v2/lib/Format.htm
    ;
    ; Build a multiline string to display in a window
    ;
    s := ""

    ; Substitute parameters by specifying indicies (order is swapped)
    s .= Format("{2}, {1}!`r`n", "World", "Hello")

    ; Padding with spaces (no order is specified so default is used)
    s .= Format("|{:-10}|`r`n|{:10}|`r`n", "Left", "Right")

    ; Hexadecimal (leading 0x, lower-case letters; upper-case letters, and fixed width)
    s .= Format("{1:#x} {2:X} 0x{3:02x}`r`n", 3735928559, 195948557, 0)

    ; Floating-point
    s .= Format("{1:0.3f} {1:.10f}", 4*ATan(1))

    displayStringInWindow(s)
}

displayStringInWindow(str)
{

    ListVars  ; Use AutoHotkey's main window to display monospaced text.
    WinWaitActive "ahk_class AutoHotkey"
    ControlSetText(str, "Edit1")
    WinWaitClose
}