#Requires AutoHotkey v2.0
#SingleInstance Force
; Slowly edit so that code makes sense to an employer
; Alt Q closes annoying window opened with Control + \
; F5 Compiles
; Ensure Brackets are together


; Tab is Tab Enter in B2B
;#HotIf WinActive("ahk_exe ")

; F5 Compiles in VSCode							{F5}
#HotIf WinActive("ahk_exe Code.exe") ; bracket needs to touch winactive
F5::Send "^{F5}"
#HotIf

; Deleting a word works in FE 					{Control + Backspace}
; #HotIf WinActive("ahk_exe explorer.exe") ; comment out so works in virtual, uncomment if problems occur
^BS::Send("^+{Left}{Del}")

; Hotkey to type Password for metamask 			{Shift + Alt + M}
<+!m:: {
	A_Clipboard := "" ;empty 
	if (!WinActive("ahk_exe KeePassXC.exe"))
		{
			ToolTip "Win not active"
			Run "C:\Program Files\KeePassXC\KeePassXC.exe"
			WinWaitActive "ahk_exe KeePassXC.exe"
			WinActivate "ahk_exe KeePassXC.exe"
			Sleep 200
			Send "^fmeta" ; unlock database and copy meta pw
			Sleep 500 
			Send "^c"
			Sleep 200
			Send "!{Tab}"
			Sleep 500
			Send "^v{Enter}" ; paste into metamask
			ToolTip
		}
		else {
			WinWaitActive "ahk_exe KeePassXC.exe"
			WinActivate "ahk_exe KeePassXC.exe"
			Sleep 400
			Send "^fmeta" ; unlock database and copy meta pw
			Sleep 500 
			Send "^c"
			Sleep 200
			Send "!{Tab}"
			Sleep 500
			Send "^v{Enter}" ; paste into metamask
			ToolTip
		}
}

; Rename drawing 								{Shift + Alt + F}
#HotIf WinActive("ahk_exe firefox.exe") ; outside of the hotkey function?
+!f:: {
	ToolTip "Active"
	VartoSend := StrReplace(A_Clipboard, "/", "_")	
	Send VartoSend " - Drawing"
	Sleep 200 
	Send "{Enter}"
	ToolTip
}
#HotIf

!g:: { ; Highlight Green						{Alt + G ... Shift}							
	Send "!hh"
	;Sleep 200 ;was 200
	Send "{Up}{Up}{Up}"
	;Sleep 200
	Send "{Left}{Left}{Left}{Left}{Enter}"
	;Sleep 200 
	Send "{Down}"
	;Sleep 100
	Send "^c"
	ToolTip "Copied"
	If (!ClipWait(1)){
		MsgBox "Clip not filled"
		return
	}
	else 
		ToolTip A_Clipboard
	;want to add variable into if statement, replace T3
	;function needs a string right

	if (KeyWait("Shift", "D T3") = 1) { ;return 0 if timed out, 1 = true = success key pressed
		Send "^v"
		Sleep 200
		Send "{PgDn}"
	}
	ToolTip
}

!r:: { ; Highlight Red 							{Alt + R ... Shift}
	Send "!hh"
	;Sleep 200 ;was 200
	Send "{Up}{Up}{Up}"
	;Sleep 200
	Send "{Left}{Left}{Left}{Left}{Left}{Left}{Left}{Left}{Left}{Enter}"
	;Sleep 200 
	Send "{Down}"
	;Sleep 100
	Send "^c"
	ToolTip "Copied"
	If (!ClipWait(1)){
		MsgBox "Clip not filled"
		return
	}
	else 
		ToolTip A_Clipboard
	;want to add variable into if statement, replace T3
	if (KeyWait("Shift", "D T3") = 1) { ;return 0 if timed out, 1 = true = success key pressed
		Send "^v"
		Sleep 200
		Send "{PgDn}"
	}
	else {
		ToolTip "Didn't press Shift in 3 Seconds"
		Sleep 1000
		ToolTip
		return
	}
	Sleep 500
	ToolTip
}

; 			NOT WORKING
!y:: { ; Highlight Yellow 						{Alt + Y ... Shift}
	Send "!hh"
	;Sleep 200
	Send "{Up}{Up}{Up}{Up}"
	;Sleep 200
	Send "{Left}{Left}{Left}{Left}{Left}{Left}{Enter}"
	;Sleep 200 
	Send "{Down}"
	;Sleep 100
	Send "^c"
	ToolTip "Copied"
	If (!ClipWait(1)){
		MsgBox "Clip not filled"
		return
	}
	else 
		ToolTip A_Clipboard
	KeyWait "Shift", "D"
	Send "^v"
	Sleep 200
	Send "{PgDn}"
	ToolTip
}

; Exit Script 									{Shift + Escape}
<+Escape:: {
	MsgBox "Exiting Script"
	ExitApp
	Sleep 3000
	ToolTip
}

; FanSelect Rename								{Control + Alt + F ... D}
; maybe getting too compilicated, need backup to rename and enter after copying article manually
^!f::
{
	; if FanSelect isn't open then open it
	if WinExist("ahk_exe FANselect.exe") {
		ToolTip "Window Exists"
		WinActivate "ahk_exe FANselect.exe"
		oldClip := A_Clipboard ; save old clip to restore later
		A_Clipboard := "" ; empty clipboard FIXED ISSUE
		Send "^c" ; copy highlighed

		if (!ClipWait(2)){ ; if clipwait not filled
			ToolTip "Clipboard not filled in 2s"
			Sleep 400
			Tooltip
			Exit
		}
		else {
			article := A_Clipboard ; copied article -> variable
		}	

		ToolTip "Press D"
		if (KeyWait("d", "D T8") = 0) { ; d means download
			ToolTip "Exit now"
			Sleep 300
			ToolTip
			Exit
		}
		Send "+{Tab}"
		Sleep 200 ; maybe needed, jesus
		Send "{Enter}"
		; wait for win explorer to open
		Sleep 1000 ; was 800, 1200 
		Send "^c"
		ToolTip ; clear press d 

		if (ClipWait(2) = 0) { ; if clipwait times out
			MsgBox "Type key and article not filled in 2s"
			Exit
		}
		else { ; otherwize 
			fullType := A_Clipboard
			sleep 400		
			article := StrReplace(article, "/", "_") ; Replace all '/' with '_'
			fullType := StrSplit(fullType, "_")
			fullForPaste := article " - " fullType[2]
			Send fullForPaste "{Enter}"
			A_Clipboard := oldClip ; restoring old clipboard
		}
	}
	else {
		; if fanselect not open then open it 
		Run "C:\Users\Jack\Documents\Misc\Programs\FS_Portable\FANselect.exe"
		ToolTip "Opening FanSelect"
		Sleep 1500 
		ToolTip
	}	
}

; FanSelect Rename Backup						{Control + Alt + Shift + F}
^!<+f:: {
	article := A_Clipboard
	A_Clipboard := ""
	
	Send "+{Tab}"
		Sleep 200 ; maybe needed, jesus
		Send "{Enter}"
		; wait for win explorer to open
		Sleep 1000 ; was 800, 1200 
	Send "^c" ; maybe need to save old clip and set clip = ""

	if (ClipWait(2) = 0) { ; if clipwait times out
		MsgBox "Type key and article not filled in 2s"
		Exit
	}
	else { ; otherwize 
		fullType := A_Clipboard
		sleep 400		
		article := StrReplace(article, "/", "_") ; Replace all '/' with '_'
		fullType := StrSplit(fullType, "_")
		fullForPaste := article " - " fullType[2]
		Send fullForPaste "{Enter}"
	}

}

; Shortcut for m^3/s							{Control + Alt + M}
^!m::										 	
{
	Send "m^+{=}3^+{=}/s{Space}"
}

; Open Script & Startup Path 					{Control + Alt + C}
^!c:: 
{
	scriptPath := "C:\Users\Jack\Documents\Misc\Scripts"
	startupPath := "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp"
	Run "explorer.exe"
	Sleep 600 ;was 500
	Send "^l" scriptPath "{Enter}" ;Go to AHK Scripts DIR
	;Sleep 6700
	ToolTip "Press Enter"
	
	if (KeyWait("Enter", "D T3") = 0) ;if function failed
	{
		MsgBox "Not opening startfolder"
		ToolTip
		Exit
	}
	else
	{
		Run "explorer.exe"
		Sleep 600
		Send "^l" startupPath "{Enter}" ;Go to AHK Scripts DIR
		ToolTip
	}
}

; Paste swap '_' '/'							{Alt + V}	
!v::
{
	originalClip := A_Clipboard 
	if InStr(originalClip,"_") {
		replaceUnder := StrReplace(A_Clipboard, "_", "/")
			Send replaceUnder
	}
		Else {
			replaceSlash := StrReplace(A_Clipboard, "/", "_")
			Send replaceSlash
		}
}

; Open explorer and search article in VD 		{Control + Shift + E}
#HotIf WinActive("ahk_exe CDViewer.exe")
^+e::
{
	Send "^+{Esc}"
	Sleep 600 ; was 600
	Send "!f"
	Sleep 200 ; was 300
	Send "n"
	Sleep 350 ; was 300 
	Send "Explorer{Enter}"
	Sleep 1700
	Send "^l" ; Control L not Alt D to focus search, confilicted with !d for tech signature
	Sleep 200 ; was 300
	Send "\\Dekun-fsmdt.za.ziehl-abegg.de\UK\Common\Technical Department{Enter}"
	Sleep 300 ; was 300
	Send "^f"
	Sleep 400 ; was 900 600
	varToSend := StrReplace(A_Clipboard, "/", "_") ; Replace all '/' with '_'
	sleep 200
	Send varToSend
	Sleep 500 ; was 300 
	Send "{Enter}"
}
#HotIf

;							NOTWORKING
; Input Box -> clipboard						{Control + Alt + E}

^!e::
{
	A_Clipboard := InputBox("Prompt", "Title").Value
	;varToSend := StrReplace(varToSend, "/", "_") ; Replace all '/' with '_'
}

!a:: ; attach T&C's then change signature		{Alt + A}
{
	Send "!nafb" ;Open Explorer when in Outlook
	Send "^l" ;Alt D to focus search, confilicted with !d for tech signature
	Send "\\Dekun-fsmdt.za.ziehl-abegg.de\UK\Common\Technical Department\Tech Library\JACK{Enter}"
	Sleep 2600 ; was 3000
	Send "{Tab}{Tab}{Tab}"
	Sleep 1100 ;WAS 1900
	Send "{Tab}{Tab}{Tab}{Down}{Down}{Down}{Enter}"
	;Add quote signature
	Sleep 1500
	Send "!has{Down}{Down}{Enter}"
}

; Create folder with current date				{Control + Shift + Alt + N}
^+!n::
{
	timeString := FormatTime(, "dd.MM.yy") ; lowercase mm is minutes so used MM 
	Send "^+n"
	Sleep 250 ;was 200
	Send timeString "{Enter}"
	Sleep 250 ;was 300 
	Send "{Enter}"
}

^+!m:: ; 										{Control + Shift + Alt + M}
{
	timeString := FormatTime(, "dd.MM.yy") ; lowercase mm is minutes so used MM 
	Send "^+n"
	Sleep 200
	Send timeString "  - JB"
	Sleep 100
	Send "{Left}{Left}{Left}{Left}{Left}"
}

/*
											SIGNATURE SCRIPTS
*/

^!x:: ; General
{
	Send "!m{Enter}"
	Sleep 300
	Send "!has{Down}{Enter}{Up}{Up}{Up}Hi ,{Enter}{Enter}Hope you're doing well. {Up}{Up}{Left}"
}

^+x:: ; General No Hi
{
	Send "!m{Enter}"
	Sleep 300
	Send "!has{Down}{Enter}"
}

^+d::
!d:: ; Tech No Hi
{
	Send "!m{Enter}" ; QUOTED OUT DUE TO ADDING AN ENTER AT THE START IF NO TECH INBOX
	Sleep 300
	Send "!has{Enter}{Backspace}"
}

^!d:: ; Tech with Hi
{
	Send "!m{Enter}"
	Sleep 300
	Send "!has{Enter}{Up}{Up}{Up}Hi ,{Left}"
}

^+z:: ; Quote No Hi
{
	; Send "!m{Enter}"
	Sleep 300
	Send "!has{Down}{Down}{Enter}"
}

^!z:: ;Quote Signature with Hi
{
	Send "!m{Enter}"
	Sleep 300
	Send "!has{Down}{Down}{Enter}{Up}{Up}{Up}Hi ,{Enter}{Enter}Hope you're doing well. {Up}{Up}{Left}"
}
