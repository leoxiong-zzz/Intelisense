/*
AutoHotkey Version:	AutoHotkey_L 1.1.07.01 Unicode 64bit
Operating System:	Micrsoft Windows 7 Home Premium 64bit
Author:				Leo Xiong [NameLess-exe] <bouncer@txttext.com>
License:			http://www.txttext.com/license
					Txttext by Leo Xiong is licensed under a Creative Commons Attribution-ShareAlike 3.0 Unported License (http://creativecommons.org/licenses/by-sa/3.0/).
					Based on a work at www.txttext.com.

Script Function:
	Intellisense for most applications with customizable word list and tooltips.
Compatibility and Prerequisites:
*/
#SingleInstance, Force
#NoEnv
CoordMode, Caret, Screen
CoordMode, ToolTip, Screen
CoordMode, Mouse, Screen
SetBatchLines, -1
SetKeyDelay, 1

;[FileInstall] Extract the splash image if it doesn't already exist
If (!FileExist("Media\IntelisenseSplashImage.png")){
	FileCreateDir, Media
	FileInstall, Media\IntelisenseSplashImage.png, Media\IntelisenseSplashImage.png
}
;[FileInstall]

;[Aero] Aero effects applied to all Guis
VarSetCapacity(rect,16,-1)
DllCall("Dwmapi\DwmIsCompositionEnabled", "Int*", DwmIsCompositionEnabled)
If (DwmIsCompositionEnabled){
	VarSetCapacity(pMarInset, 16, -1)
	Loop, 3{
		Gui, %A_Index%:+LastFound +Resize
		Gui, %A_Index%:Color, % (A_Index = 3 ? "0x000000" : "0xFFFFFE")
		DllCall("dwmapi\DwmExtendFrameIntoClientArea", "UInt", WinExist(), "UInt", &pMarInset)
		WinSet, TransColor, 0xFFFFFE
	}
	WinSet, Trans, 255
}
;[Aero]

;[Gui] Open splash window
Gui, 3:+AlwaysOnTop +ToolWindow -Caption
Gui, 3:Add, Picture, x0 y0 gPictureSplashImage, Media\IntelisenseSplashImage.png
Gui, 3:Show,, Initialize - Intelisense
;[Gui]

;[FileInstall] Extract the word list if it doesn't already exist
If (!FileExist("Media\Words.txt"))
	FileInstall, Media\Words.txt, Media\Words.txt
;[FileInstall]

;[IniRead] Read configuration files
IniRead, HotKeySend, Media\Options.ini, Options, HotKeySend, Tab
IniRead, GuiSizeW, Media\Options.ini, Options, GuiSizeW, 220
IniRead, GuiSizeH, Media\Options.ini, Options, GuiSizeH, 100
IniRead, DropDownListSelect, Media\Options.ini, Options, DropDownListSelect, 3
IniRead, HotkeySuspend, Media\Options.ini, Options, HotkeySuspend, ^J
;[IniRead]

;[Gui] Create two Gui for the search results and settings
Gui, +AlwaysOnTop +Resize +ToolWindow -Caption MinSize220x100 MaxSize300x500
Gui, Add, ListView, x0 y0 gListViewResult vListViewResult Sort, Initializing
Gui, Font, s7 Bold, Segoe UI
Gui, Show, w%GuiSizeW% h%GuiSizeH% Hide, Search - Intelisense
Gui, 2:Font, cBlack s16, Segoe UI
Gui, 2:Add, Text, x0 y0, Options
Gui, 2:Font, s7 Bold
Gui, 2:Add, Text, x20 y30, Select Suggestion
Gui, 2:Add, Text, x20 y50, Toggle Intellisense
Gui, 2:Font, Normal
Gui, 2:Add, DropDownList, x120 y25 w130 AltSubmit Choose%DropDownListSelect% gDropDownListSelect vDropDownListSelect, % (DropDownListSelectList := "Enter|Space|Tab|Right|Numpad Enter|Numpad Right")
Gui, 2:Add, Hotkey, x120 y45 w130 gHotkeySuspend vHotkeySuspend, %HotkeySuspend%
Gui, 2:Add, Text, x0 y65 w260 Right gTextTxttext, % Chr(169) " Copyright Txttext, 2009 - " A_Year ". All rights reserved."
Gui, 2:Show, w260 h80 Hide, Options - Intelisense
;[Gui]

;[Menu] Tray Menu Options
Menu, Tray, NoStandard
Menu, Tray, Add, Settings, MenuHandler
Menu, Tray, Add, Refresh, MenuHandler
Menu, Tray, Add
Menu, Tray, Add, Suspend, Suspend
Menu, Tray, Add, Exit, MenuHandler
Menu, Tray, Default, Suspend
;[Menu]

;[StringSplit] Splits DropDownListSelectList into seperate strings
StringReplace, DropDownListSelectList, DropDownListSelectList, %A_Space%,, All
StringSplit, DropDownListSelectList, DropDownListSelectList, |
;[Hotkey] Create hotkeys for each key press, sending the result, and suspending the script
KeyDelimiter = Enter Space Right Tab Ctrl Alt AppsKey Esc Home End `` ~ ! @ # $ `% ^ & * ( ) - = + [ { ] } \ | `; : ' `" , < . > / ? NumpadDiv NumpadMult NumpadSub NumpadAdd NumpadDot NumpadEnter NumpadRight
KeyDown = a b c d e f g h i j k l m n o p q r s t u v w x y z 1 2 3 4 5 6 7 8 9 0 Numpad0 Numpad2 Numpad3 Numpad4 Numpad5 Numpad6 Numpad7 Numpad8 Numpad9 Numpad0 Numpad0 _
Loop, 10
{
	Hotkey, % "~" Chr(A_Index + 47), KeyDown
	Hotkey, % "~Numpad" Chr(A_Index + 47), KeyDown
}
Loop, Parse, KeyDown, %A_Space%
{
	Hotkey, % "~" A_LoopField, KeyDown
	Hotkey, % "~+" A_LoopField, KeyDown
}
Hotkey, ~BackSpace, KeyDown
Hotkey, %HotkeySuspend%, Suspend
;Hotkey, IfWinExist, Search - Intelisense ahk_class AutoHotkeyGUI
Loop, Parse, KeyDelimiter, %A_Space%
	Hotkey, $%A_LoopField%, KeyDelimiter
;[Hotkey]

;[DllCall] Retrieves the frequency
DllCall("QueryPerformanceFrequency", "Int64*", QueryPerformanceFrequency)
;[SetTimer] Starts a timer to check for when the window switches
SetTimer, SetTimer, 0

;[Gosub] Parse the word list
Gosub, ParseWords

;[Fade Splah] Fade the splash window
Sleep, 300
Loop, 765
	WinSet, Transparent, % 255 - (A_Index / 3), Initialize - Intelisense
;[PictureSplashImage] Destroy the Gui when use clicks on the image
PictureSplashImage:
Gui, 3:Destroy
;[PictureSplashImage]
;[Fade Splah]

Return
;[Settings] Save options upon editing
DropDownListSelect:
HotkeySuspend:
GuiControlGet, %A_ThisLabel%
IniWrite, % %A_ThisLabel%, Media\Options.ini, Options, %A_ThisLabel%
If (A_ThisLabel = "DropDownListSelect")
	Hotkey, IfWinExist, Search - Intelisense ahk_class AutoHotkeyGUI
Else If (A_ThisLabel = "HotkeySuspend"){
	Hotkey, IfWinExist
	Hotkey, % %A_ThisLabel%, Suspend
}
;[Settings]

Return
;[HotKeySend] Triggered when the user selects a result
HotKeySend:
LV_GetText(LV_GetText, LV_GetNext())
Gui, Hide
Send, % SubStr(LV_GetText, StrLen(KeyHistory) + 1)
KeyHistory := LV_GetText
If (RegExMatch(ToolTips, "i)\b" LV_GetText "\b\[(.*)]", RegExMatch))
	ToolTip, %RegExMatch1%, %A_CaretX%, % A_CaretY + 20
;[HotKeySend]

Return
;[GuiHide] Hides the Gui and clears memory
GuiHide:
Gui, Hide
ToolTip
KeyHistory =
;[GuiHide]

Return
;[GuiSize] Adjusts controls and saves sizes
GuiSize:
GuiControl, Move, ListViewResult, h%A_GuiHeight% w%A_GuiWidth%
IniWrite, %A_GuiWidth%, Media\Options.ini, Options, GuiSizeW
IniWrite, %A_GuiHeight%, Media\Options.ini, Options, GuiSizeH
;[GuiSize]

Return
;[KeyDelimiter] Reset the key history or send if the send key is pressed
KeyDelimiter:
If (A_ThisHotkey <> "$" DropDownListSelectList%DropDownListSelect%) Or (!WinExist("Search - Intelisense ahk_class AutoHotkeyGUI")){
	Send, % "{" SubStr(A_ThisHotkey, RegExMatch(A_ThisHotkey, "\w+")) "}"
	Goto, GuiHide
}
Else If (WinExist("Search - Intelisense ahk_class AutoHotkeyGUI"))
	Goto HotkeySend
;[KeyDelimiter]

Return
;[KeyDown] Triggered when a key is pressed
KeyDown:
WinGetActiveTitle, WinActive
ControlGetFocus, ControlGetFocus, %WinActive%
ToolTip
If (!ErrorLevel) And (WinActive <> "Search - Intelisense") And ((KeyHistory := (A_ThisHotkey = "~BackSpace") ? SubStr(KeyHistory, 1, -1) : (KeyHistory SubStr(A_ThisHotkey, 0))) <> ""){
	LV_Delete(), InStr = 1, DllCall("QueryPerformanceCounter", "Int64*", QueryPerformanceCounter)
	While, (SubStr((SubStr := SubStr(Words, (InStr := InStr(Words, "`n" KeyHistory, False, InStr + 1) + 1), InStr(Words, "`n", False, InStr + 1) - InStr)), 1, StrLen(KeyHistory)) = KeyHistory) And ((QueryPerformanceCounter2 - QueryPerformanceCounter) / QueryPerformanceFrequency < 0.05)
		LV_Add("", SubStr), DllCall("QueryPerformanceCounter", "Int64*", QueryPerformanceCounter2)
	If (LV_GetCount()){
		LV_ModifyCol(1, "Text", "[" KeyHistory "] " LV_GetCount() " results found in " (QueryPerformanceCounter2 - QueryPerformanceCounter) / QueryPerformanceFrequency * 1000 "ms")
		LV_Modify(1, "Select")
		Gui, Show, % "x" A_CaretX "y" A_CaretY + 20 "NoActivate"
	}
	Else
		Gui, Hide
}
Else
	Gui, Hide
;[KeyDown]

Return
;[ListViewResult] Send result when double clicked
ListViewResult:
If (!A_GuiEvent = "DoubleClick")
	Return
Goto, HotkeySend
;[ListViewResult]

Return
;[MenuHandler] Menu Handler
MenuHandler:
If (A_ThisMenuItemPos = 1)
	Gui, 2:Show
Else If (A_ThisMenuItemPos = 2)
	Goto, ParseWords
Else If (A_ThisMenuItemPos = 5)
	ExitApp
;[MenuHandler]

Return
;[ParseWords] Sorts the word list
ParseWords:
FileRead, Words, Media\Words.txt
FileRead, ToolTips, Media\ToolTips.txt
StringReplace, Words, Words, `r,, All
Sort, Words, D`n
;[ParseWords]

Return
;[SetTimer] Starts a timer to check for when the window switches
SetTimer:
WinGetActiveTitle, ActiveWinTitle
ControlGetFocus, ControlGet, %ActiveWinTitle%
If (ActiveWinTitle ControlGet <> SetTimer) And (ActiveWinTitle <> "Search - Intelisense"){
	SetTimer := ActiveWinTitle ControlGet
	Goto, GuiHide
}
;[SetTimer]

Return
;[Suspend] Suspends sciprt
Suspend:
Suspend
Goto, GuiHide
;[Suspend]

Return
;[TextTxttext] Open browser to http://www.txttext.com
TextTxttext:
Run, http://www.txttext.com
;[TextTxttext]

Return
;[#IfWinExist] Hotkeys below will only trigger when the window exists
#IfWinExist Search - Intelisense ahk_class AutoHotkeyGUI
;[Navigation] Navigates through the results list
Up::
Down::
NumpadUp::
NumpadDown::
PGUP::
PGDN::
Home::
End::
ControlSend, SysListView321, {%A_ThisHotkey%}, Search - Intelisense ahk_class AutoHotkeyGUI
LV_GetText(LV_GetText, LV_GetNext())
Send, % SubStr(LV_GetText, StrLen(KeyHistory) + 1) "+{Left " StrLen(LV_GetText) - StrLen(KeyHistory) "}"
;[Navigation]

Return
;[Function Keys] Searches definition and plays pronunciation
F1::
F2::
GuiControlGet, ListBoxResult,, ListBoxResult
LV_GetText(LV_GetText, LV_GetNext())
If (A_ThisHotkey = "F1")
	Run, http://www.google.com/search?q=define %LV_GetText%&btnI
Else If (A_ThisHotkey = "F2"){
	URLDownloadToFile, http://translate.google.com/translate_tts?q=%LV_GetText%, Media\Intelisense - Pronunciation.mp3
	SoundPlay, Media\Intelisense - Pronunciation.mp3, Wait
	FileDelete, Media\Intelisense - Pronunciation.mp3
}
;[Function Keys]

;[Del] Removes the word from the word list
Del::
LV_GetText(LV_GetText, LV_GetNext())
Gosub, GuiHide
FileDelete, Media\Words.txt
FileAppend, % RegExReplace(Words, LV_GetText "`n"), Media\Words.txt
Goto, ParseWords
;[Del]

Return
;[Hide Gui] Hides the Gui
Esc::
Left::
NumpadLeft::
Goto, GuiHide
;[Hide Gui]

;[#IfWinExist] Hotkeys below will trigger anywhere
#IfWinExist
Return
;[Insert Word] Append word to the word list
Ins::
If (KeyHistory){
	FileAppend, `n%KeyHistory%, Media\Words.txt
	Gosub, GuiHide
	Goto, ParseWords
}
Else
	Send, {Ins}
;[Insert Word]