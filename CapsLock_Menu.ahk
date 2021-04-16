#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#SingleInstance Force  ; Replaces the old instance automatically.
SetBatchLines -1  ; Run the script at maximum speed.

CapsLockMenu()
{
Menu, CapsLockMenu, Add
Menu, CapsLockMenu, Delete
Menu, CapsLockMenu, Add, CAPSLock Menu, ToggleCapsLock
Menu, CapsLockMenu, Add
Menu, CapsLockMenu, Add, CapsLock O&N, CapsLockOn
Menu, CapsLockMenu, Add, CapsLock &off, CapsLockOff
Menu, CapsLockMenu, Add
Menu, CapsLockMenu, Add, Paste as &Plain Text, PastePlain
Menu, CapsLockMenu, Add
Menu, ConvertCaseMenu, Add
Menu, ConvertCaseMenu, Delete
Menu, ConvertCaseMenu, Add, &Title Case, Title
Menu, ConvertCaseMenu, Add, Ca&pital Case, Capital
Menu, ConvertCaseMenu, Add, &Sentence case, Sentence
Menu, ConvertCaseMenu, Add
Menu, ConvertCaseMenu, Add, &UPPERCASE, Upper
Menu, ConvertCaseMenu, Add, lo&wercase, Lower
Menu, ConvertCaseMenu, Add, &camelCase, camel
Menu, ConvertCaseMenu, Add, &PascalCase, Pascal
Menu, ConvertCaseMenu, Add
Menu, ConvertCaseMenu, Add, &Dot.Case, Dot
Menu, ConvertCaseMenu, Add, S&nake_Case, Snake
Menu, ConvertCaseMenu, Add, &Kebab-Case, Kebab
Menu, ConvertCaseMenu, Add
Menu, ConvertCaseMenu, Add, iN&VERT cASE, Invert
Menu, ConvertCaseMenu, Add, &RaNdoM caSe, Random
Menu, ConvertCaseMenu, Add, &aLtErNaTiNg cAsE, Alternating
Menu, CapsLockMenu, Add, Con&vert Case, :ConvertCaseMenu
Menu, CapsLockMenu, Add
Menu, CapsLockMenu, Add, &Emoji Keyboard, OpenEmojiKeyboard
Menu, InsertLineMenu, Add
Menu, InsertLineMenu, Delete
Menu, InsertLineMenu, Add, &Light, LightHorizontalBoxDrawing
Menu, InsertLineMenu, Add, &Double, DoubleHorizontalBoxDrawing
Menu, CapsLockMenu, Add, Insert &Line, :InsertLineMenu
Menu, CapsLockMenu, Default, CapsLock Menu
Menu, CapsLockMenu, Show
}

CopyClipboard()
{
    global ClipSaved := ""
    ClipSaved := ClipboardAll  ; save original clipboard contents
    Clipboard := ""  ; start off empty to allow ClipWait to detect when the text has arrived
    Send {Ctrl down}c{Ctrl up}
    Sleep 150
    ClipWait, 1.5, 1
    if ErrorLevel
    {
        MsgBox, 262208, AutoHotkey, Copy to clipboard failed.
        Clipboard := ClipSaved  ; restore the original clipboard contents
        ClipSaved := ""  ; clear the variable
        return
    }
}

PastePlain()
{
    ClipSaved := ClipboardAll  ; save original clipboard contents
    Clipboard := Clipboard  ; remove formatting
    Send ^v  ; send the Ctrl+V command
    Sleep 100  ; give some time to finish paste (before restoring clipboard)
    Clipboard := ClipSaved  ; restore the original clipboard contents
    ClipSaved := ""  ; clear the variable
}

; creating something that links the function to a specific state that can then be called by legacy commands
ToggleCapsLock(){
    return, CapsLockey(, True)
}
CapsLockOn(){
    return, CapsLockey(True)
}
CapsLockOff(){
    return, CapsLockey()
}

; actual function that does the work
CapsLockey(state := false, toggle := false)
{
    ; list the messages for the message box to report what happened
    static messages := {0:"CapsLock Status: OFF", 1:"CapsLock Status: ON", "failed":"CapsLock operation failed"}
    
    ; decide whether how to set the state
    state := toggle ? !GetKeyState("CapsLock", "T") : state
    ;          1    2       3                       4   5
    /*
        1: Condition what should happen
        2: Ternary operator. Signals that to its left is a condition (1) that should be used to decide between the two possibilitys on the right of it (3 or 5)
        3: Getting the toggle keystate of CapsLock and invert it, to allow toggle functionality
        4: Marker to signal that the true section of the ternary operation ends and the false section begins 
        5: The state handed in the function call
    */
    SetCapsLockState % state
    MsgBox, 262208, CapsLock Menu, % messages[(GetKeyState("CapsLock", "T") == state) ? state : "failed"]
}

CopyClipboardCLM()
{
    global ClipSaved
    WinGet, id, ID, A
    WinGetClass, class, ahk_id %id%
    if (class ~= "(Cabinet|Explore)WClass|Progman")
        Send {F2}
    Sleep 100
    CopyClipboard()
    if (ClipSaved != "")
        Clipboard := Clipboard
    else
        Exit
}

PasteClipboardCLM()
{
    global ClipSaved
    WinGet, id, ID, A
    WinGetClass, class, ahk_id %id%
    if (class ~= "(Cabinet|Explore)WClass|Progman")
        Send {F2}
    Send ^v
    Sleep 100
    Clipboard := ClipSaved
    ClipSaved := ""
}

Title()
{
    ExcludeList := ["a", "about", "above", "after", "an", "and", "as", "at", "before", "but", "by", "for", "from", "nor", "of", "or", "so", "the", "to", "via", "with", "within", "without", "yet"]
    ExactExcludeList := ["AutoHotkey", "iPad", "iPhone", "iPod", "OneNote", "USA"]
    CopyClipboardCLM()
    TitleCase := Format("{:T}", Clipboard)
    for _, v in ExcludeList
        TitleCase := RegexReplace(TitleCase, "i)(?<!\. |\? |\! |^)(" v ")(?!\.|\?|\!|$)\b", "$L1")
    for _, v in ExactExcludeList
        TitleCase := RegExReplace(TitleCase, "i)\b" v "\b", v)
    TitleCase := RegexReplace(TitleCase, "im)\b(\d+)(st|nd|rd|th)\b", "$1$L{2}")
    Clipboard := TitleCase
    PasteClipboardCLM()
}

Capital()
{
    ExactExcludeList := ["AutoHotkey", "iPad", "iPhone", "iPod", "OneNote", "USA"]
    CopyClipboardCLM()
    CapitalCase := Format("{:T}", Clipboard)
    for _, v in ExactExcludeList
        CapitalCase := RegExReplace(CapitalCase, "i)\b" v "\b", v)
    Clipboard := CapitalCase
    PasteClipboardCLM()
}

Sentence()
{
    ExactExcludeList := ["AutoHotkey", "iPad", "iPhone", "iPod", "OneNote", "USA"]
    CopyClipboardCLM()
    StringLower, Clipboard, Clipboard
    Clipboard := RegExReplace(Clipboard, "(((^|([.!?]\s+))[a-z])| i | i')", "$u1")
    for _, v in ExactExcludeList
        Clipboard := RegExReplace(Clipboard, "i)\b" v "\b", v)
    PasteClipboardCLM()
}

Upper()
{
    CopyClipboardCLM()
    StringUpper, Clipboard, Clipboard
    PasteClipboardCLM()
}

Lower()
{
    CopyClipboardCLM()
    StringLower, Clipboard, Clipboard
    PasteClipboardCLM()
}

camel()
{
    CopyClipboardCLM()
    StringUpper, Clipboard, Clipboard, T
    FirstChar := SubStr(Clipboard, 1, 1)
    StringLower, FirstChar, FirstChar
    camelCase := SubStr(Clipboard, 2)
    camelCase := StrReplace(camelCase, A_Space)
    Clipboard := FirstChar camelCase
    PasteClipboardCLM()
}

Pascal()
{
    CopyClipboardCLM()
    StringUpper, Clipboard, Clipboard, T
    Clipboard := StrReplace(Clipboard, A_Space)
    PasteClipboardCLM()
}

Dot()
{
    CopyClipboardCLM()
    if RegExMatch(Clipboard, "\-|\_|\.") != "0"
    {
        Clipboard := RegExReplace(Clipboard, "\-|\_|\.", " ")
        PasteClipboardCLM()
    }
    else
    {
        Clipboard := StrReplace(Clipboard, A_Space, ".")
        PasteClipboardCLM()
    }
}

Snake()
{
    CopyClipboardCLM()
    if RegExMatch(Clipboard, "\-|\_|\.") != "0"
    {
        Clipboard := RegExReplace(Clipboard, "\-|\_|\.", " ")
        PasteClipboardCLM()
    }
    else
    {
        Clipboard := StrReplace(Clipboard, A_Space, "_")
        PasteClipboardCLM()
    }
}

Kebab()
{
    CopyClipboardCLM()
    if RegExMatch(Clipboard, "\-|\_|\.") != "0"
    {
        Clipboard := RegExReplace(Clipboard, "\-|\_|\.", " ")
        PasteClipboardCLM()
    }
    else
    {
        Clipboard := StrReplace(Clipboard, A_Space, "-")
        PasteClipboardCLM()
    } 
}

Invert()
{
    CopyClipboardCLM()
    Inv_Char_Out := ""
    Loop % StrLen(Clipboard)
    {
        Inv_Char := SubStr(Clipboard, A_Index, 1)
        if Inv_Char is Upper
            Inv_Char_Out := Inv_Char_Out Chr(Asc(Inv_Char) + 32)
        else if Inv_Char is Lower
            Inv_Char_Out := Inv_Char_Out Chr(Asc(Inv_Char) - 32)
        else
            Inv_Char_Out := Inv_Char_Out Inv_Char
    }
    Clipboard := Inv_Char_Out
    PasteClipboardCLM()
}

Random()
{
    CopyClipboardCLM()
    RandomCase := ""
    for _, v in StrSplit(Clipboard)
    {
        Random, r, 0, 1
        RandomCase .= Format("{:" (r?"L":"U") "}", v)
    }
    Clipboard := RandomCase
    PasteClipboardCLM()
}

Alternating()
{
    CopyClipboardCLM()
    Inv_Char_Out := ""
    StringLower, Clipboard, Clipboard
    Loop, Parse, Clipboard
    {
        if (Mod(A_Index, 2) = 0)
            Inv_Char_Out .= Format("{1:U}", A_LoopField)
        else
            Inv_Char_Out .= Format("{1:L}", A_LoopField)
    }
    Clipboard := Inv_Char_Out
    PasteClipboardCLM()
}

OpenEmojiKeyboard()
{
    Send {LWin down}.
    Send {LWin up}
}

LightHorizontalBoxDrawing()
{
    InputBox, UserInput, Input, , , 150, 105
    Loop %UserInput%
        Send {U+2500}
}

DoubleHorizontalBoxDrawing()
{
    InputBox, UserInput, Input, , , 150, 105
    Loop %UserInput%
        Send {U+2550}
}


$CapsLock::
    KeyWait CapsLock, T0.25
        if ErrorLevel
            CapsLockMenu()
        else
        {
            KeyWait CapsLock, D T0.25
            if ErrorLevel
                CopyClipboard()
            else
                Send ^v
        }
    KeyWait CapsLock
return