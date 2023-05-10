#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#SingleInstance force ; https://ahkplant.tistory.com/69



; https://stackoverflow.com/questions/31839062/autohotkey-in-windows-10-hotkeys-not-working-in-some-applications
if not A_IsAdmin
{
   Run *RunAs "%A_ScriptFullPath%"  ; Requires v1.0.92.01+
   ExitApp
}


; @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
; Initializer >>>
;https://www.autohotkey.com/board/topic/78199-script-opens-folder-supposed-to-create-folder-if-not-exist/
Path2Save = % A_ScriptDir . "\" . "AutoHotkeyOutput" . "\"
If !FileExist(Path2Save) {
 FileCreateDir, %Path2Save%
 If ErrorLevel
  MsgBox, 48, Error, An error occurred when creating the directory.`n`n%Path2Save%
 Else MsgBox, 64, Success, Directory was created.`n`n%Path2Save%
} ; Else MsgBox, 64, Exists, Directory already exists.`n`n%Path2Save%

InitTrayMenu()


return ; end of autoexec?
; <<< Initializer
; @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

InitTrayMenu(){ ; https://jacks-autohotkey-blog.com/2018/01/02/recognize-running-scripts-with-system-tray-icon-techniques-autohotkey-tip/
;	Menu, Tray, NoStandard ; remove default tray menu entries
;	Menu, Tray, Add, Exit, Exit
;    Menu, Tray, Default, &Open
    Menu, Tray, Click, 1
;    Menu, Tray, Tip, Click for Defult Menu!`r(Right-click for more options.)
    Menu, Tray, Icon, shell32.dll, 66 ; https://help4windows.com/ 41+1 tree 65+1 duplicate!
}

;Ctrl & Esc:: Exit()         ; Exit()
;^#Esc:: Exit()  ; Ctrl+Win+Esc              ; Exit()
;^#!F5:: Exit()   ; Ctrl+Win+Alt+Esc          ; Exit()
^#F5:: Exit()   ; Ctrl+Win+F5              ; Exit()
Exit(){
	ExitApp
}


^#F2::          ; Ctrl+Win+F2               ; Help
MsgBox 0, Help for %A_ScriptName%.                                                                                               .,
(
/**************** Default ******************/
^#F1::          ; Ctrl+Win+F1               ; Help
^#F12::         ; Ctrl+Win+F12              ; KeyHistory
#ScrollLock::   ; win + scrolllock          ; toggle tooltip state
^#+F5:: Exit()  ; Ctrl+Win+Shift+F5         ; Exit the default ahk.

+CapsLock::     ; Shift+CapsLock            ; Change the IME to English.
^+CapsLock::    ; Ctrl+Shift+CapsLock       ; Change the IME to Korean.

^#v::           ; Ctrl+Win+v                ; Text-only paste from ClipBoard.
^#x::           ; Ctrl+Win+x                ; Text-only cut/copy to ClipBoard & Save ClipBoard to a txt file.
^#c::           ; Ctrl+Win+c                ; Text-only cut/copy to ClipBoard & Save ClipBoard to a txt file.
^#f::    		; Ctrl+Win+f 			    ; Text-only copy & filter lines & save to a txt file.

^#+v::          ; Ctrl+Win+Shift+v            ; Text-only paste from ClipBoard & remove leading spaces.
^#+c::          ; Ctrl+Win+Shift+c            ; Text-only cut/copy to ClipBoard & remove leading spaces & Save ClipBoard to a txt file.

^#+9::          ; Ctrl+Win+Shift+9{(}       ; Text-only cut/copy to ClipBoard & extract outside parentheses & Save ClipBoard to a txt file.
^#+0::          ; Ctrl+Win+Shift+0{}}       ; Text-only cut/copy to ClipBoard & extract inside parentheses & Save ClipBoard to a txt file.

^#s::           ; Ctrl+Win+s                ; Save ClipBoard to a txt file.

/**************** @ TXT    *****************/
^NumpadLeft::   ; Ctrl+Shift+Numpad4      	; Send  " <- "
^NumpadRight::  ; Ctrl+Shift+Numpad6      	; Send  " -> "
^#NumpadLeft::  ; Ctrl+Win+Shift+Numpad4    ; Send {U+2190}
^#NumpadRight:: ; Ctrl+Win+Shift+Numpad6    ; Send {U+2192}

^+x::           ; Ctrl+Shift+x              ; Send "X" in English


/*******************************************************/
/**************** @ another ahk script *****************/
^#F2::          ; Ctrl+Win+F2               ; Help for the another ahk script.
^#F5:: Exit()   ; Ctrl+Win+F5        		; Exit the another ahk script.

/**************** @ TSV @ EMR *****************/
^!Numpad4::     ; Ctrl+Alt+Numpad4          ; Add "<-" between the numbers
^!Numpad6::     ; Ctrl+Alt+Numpad6          ; Swith the numbers & Add "->" between the numbers
^#!Numpad4::    ; Ctrl+Win+Alt+Numpad4  	; Add ←(U+2190) between the numbers
^#!Numpad6::    ; Ctrl+Win+Alt+Numpad6  	; Swith the numbers & Add → (U+2192) between the numbers

^#!c::    		; Ctrl+Win+Alt+c 			; copy & process labs & save to a text file
^#!v::    		; Ctrl+Win+Alt+v 			; (from the copied clipboard) process labs & paste? (void)

^#d::    		; Ctrl+Win+d 	 			; copy & process date & save to a text file (pending)
^#t::    		; Ctrl+Win+t 	 			; copy & replace tab to left arrow & save to a text file
^#/::    		; Ctrl+Win+/ 	 			; copy & join multi-line text (/) & save to a text file

^#l::    		; Ctrl+Win+l 	 			; copy & process labs & save to a text file
^#b::    		; Ctrl+Win+b 	 			; copy & process BP / Bwt / BMI & save to a text file (pending)
^#o::    		; Ctrl+Win+o 	 			; copy & process procedure orders & save to a text file
^#m::    		; Ctrl+Win+m 	 			; copy & process medication orders & save to a text file

/**************** @ function *****************/
functionPasteViaClipboardThenRestore(Text2Paste)
functionAutoHotkeySave(A_ScriptDir . "\" . "AutoHotkeyOutput")
functionCopy2Clipboard2Txt(Text2CopyPaste)

functionEMR_TSV_AddLeftArrow(EMR_TSV)
functionEMR_TSV_AddLeftArrowUnicode(EMR_TSV)
functionEMR_TSV_SwitchValue_AddRightArrow(EMR_TSV)
functionEMR_TSV_SwitchValue_AddRightArrowUnicode(EMR_TSV)

functionFilterLines(TextInput, RegEx2Filter)
functionProcessText_lab(TextInput)
functionProcessText_EMR_medication(TextInput) ; ^#m::    		; Ctrl+Win+m 	 			; copy & process medications & save to a text file
functionProcessText_join_MultiLine(TextInput) ; ^#/::    		; Ctrl+Win+/ 	 			; copy & join multi-line text (/) & save to a text file
)
return


/*************************************************************************************************************************
*/


functionPasteViaClipboardThenRestore(Text2Paste)
{
   ;-----------------------------------
   global ClipBoardAll0
   ClipBoardAll0 = %ClipBoardAll%           ; Calling %ClipBoardAll% retrieves all data (including formats) from ClipBoard?!
   Sleep 100
   ;-----------------------------------

   ClipBoard := Text2Paste               ; copy the processed output into the ClipBoard
   Sleep 100
   ClipWait 2                    ; Waits until the clipboard contains data, up to 2s
   Send ^v                       ; For best compatibility: SendPlay

   ;-----------------------------------
   Sleep 100                      ; Don't change clipboard while it is pasted! (Sleep > 0)
   ClipBoard = %ClipBoardAll0%           ; Restore original ClipBoard
   Sleep 100
   ClipWait 2                    ; Waits until the clipboard contains data, up to 2s
   VarSetCapacity(ClipBoardAll0, 0)      ; Free memory
   ;-----------------------------------

   return Text2Paste
}

functionAutoHotkeySave(Path2Save)   ; functionAutoHotkeySave(A_ScriptDir . "\" . "AutoHotkeyOutput")
{
   ; https://www.autohotkey.com/board/topic/30021-send-clipboard-to-notepad-and-save/
   ; to find out where the script left off last time, loop through the existing
   ; saved text files and look for the highest number there.
   numb := 0 ; https://www.autohotkey.com/boards/viewtopic.php?t=57472
   Loop, %Path2Save%\AutoHotkeySave*.txt
      If ( SubStr( A_LoopFileName, StrLen("AutoHotkeySave")+1, 4 ) + 1 > numb )
         numb := SubStr( A_LoopFileName, StrLen("AutoHotkeySave")+1, 4 )

   ;#g:: ; increment variable 'numb' or zero it and prepend zeros up to length 4
   numb := SubStr("0000" (numb+1), -3)
   StringReplace, Clipboard_n, Clipboard, `r`n, `n, all ; drop CrNl for just newline
   FileAppend, %Clipboard_n%, %Path2Save%\AutoHotkeySave%numb%.txt

   Sleep 1000
   Run Notepad.exe `"%Path2Save%\AutoHotkeySave%numb%.txt`"  ; D:\OneDrive - SNU\문서
   VarSetCapacity(Clipboard_n, 0)      ; Free memory
}

functionCopy2Clipboard2Txt(Text2CopyPaste)
{
   ClipBoard := Text2CopyPaste
   Sleep 100
   ClipWait 2                    ; Waits until the clipboard contains data, up to 2s
   functionAutoHotkeySave(A_ScriptDir . "\" . "AutoHotkeyOutput")
   return Text2CopyPaste
}


/*************************************************************************************************************************
T4(free)(진검시행)	0.90	1.03
TSH(진검시행)	3.41	3.84
Calcium	9.2	8.9
Phosphorus	3.5	2.9
Glucose	95	102
Uric acid	3.3	3.0
Chol.	171	198
T. Protein	7.2	7.2
Albumin	4.5	4.4
T. Bil.	1.0	1.0
Alk. phos.	46	49
AST(GOT)	24	33
ALT(GPT)	29	47
GGT	32	37
BUN	10	13
Creatinine	0.77	0.73
eGFR(MDRD)	104.9	111.6
eGFR(CKD EPI Cr)	102.2	104.4
TG	270	179
HDL Chol.	36	44
LDL Chol.	87	113
LDL Chol.(계산식)	81	118
hs-CRP	0.03	0.06
WBC	4.71	5.27
RBC	4.79	4.91
Hb	15.3	16.2
Hct	45.5	47.0
MCV	95.0	95.7
MCH	31.9	33.0
MCHC	33.6	34.5
RDW	11.9	12.0
Platelet	164	157
PCT	0.17	0.17
MPV	10.4	10.9
PDW	12.1	12.3
ESR	2	5
Blast
Promyelo
Myelocyte
Metamyelo
Band.neut.
Seg.neut.	51.4	46.8
Lymphocyte	38.2	39.5
Monocyte	7.9	10.8
Eosinophil	2.1	2.7
Basophil	0.4	0.2
Imm.cell
Imm.mono
AtypicalLc
Normoblast
Large unstained cell
ANC	2421	2466
Hb A1c	5.2	5.3
*/





/*************************************************************************************************************************
*/

functionEMR_TSV_AddLeftArrow(EMR_TSV)
{
	Return regexreplace(EMR_TSV, "(.*)\t([0-9]+(?:\.[0-9]*)?).*\t([0-9]+(?:\.[0-9]*)?).*", "$1`t`t$2`t <- $3") . "`r`n"
}

functionEMR_TSV_SwitchValue_AddRightArrow(EMR_TSV)
{
	Return regexreplace(EMR_TSV, "(.*)\t([0-9]+(?:\.[0-9]*)?).*\t([0-9]+(?:\.[0-9]*)?).*", "$1`t`t$3`t -> $2") . "`r`n"
}

functionEMR_TSV_AddLeftArrowUnicode(EMR_TSV)
{
	Return regexreplace(EMR_TSV, "(.*)\t([0-9]+(?:\.[0-9]*)?).*\t([0-9]+(?:\.[0-9]*)?).*", "$1`t`t$2`t" " " chr("0x" . "2190") " " "$3") . "`r`n"
}

functionEMR_TSV_SwitchValue_AddRightArrowUnicode(EMR_TSV)
{
	Return regexreplace(EMR_TSV, "(.*)\t([0-9]+(?:\.[0-9]*)?).*\t([0-9]+(?:\.[0-9]*)?).*", "$1`t`t$3`t" " " chr("0x" . "2192") " " "$2") . "`r`n"
}


^!Numpad4::     ; Ctrl+Alt+Numpad4      ; Add "<-" between the numbers
   Send ^c
   Sleep 100
   ClipWait 2                    ; Waits until the clipboard contains data, up to 2s
   ClipBoard2Text = %ClipBoard%            ; Store %ClipBoard% to ClipBoard2Text - Automatically converted to text without any format?!
   Sleep 100
   ClipWait 2                    ; Waits until the clipboard contains data, up to 2s
   functionCopy2Clipboard2Txt(functionEMR_TSV_AddLeftArrow(ClipBoard2Text))
   VarSetCapacity(ClipBoard2Text, 0)      ; Free memory
Return

^!Numpad6::     ; Ctrl+Alt+Numpad6      ; Swith the numbers & Add "->" between the numbers
   Send ^c
   Sleep 100
   ClipWait 2                    ; Waits until the clipboard contains data, up to 2s
   ClipBoard2Text = %ClipBoard%            ; Store %ClipBoard% to ClipBoard2Text - Automatically converted to text without any format?!
   Sleep 100
   ClipWait 2                    ; Waits until the clipboard contains data, up to 2s
   functionCopy2Clipboard2Txt(functionEMR_TSV_SwitchValue_AddRightArrow(ClipBoard2Text))
   VarSetCapacity(ClipBoard2Text, 0)      ; Free memory
Return

^#!Numpad4::    ; Ctrl+Win+Alt+Numpad4  	; Add ←(U+2190) between the numbers
   Send ^c
   Sleep 100
   ClipWait 2                    ; Waits until the clipboard contains data, up to 2s
   ClipBoard2Text = %ClipBoard%            ; Store %ClipBoard% to ClipBoard2Text - Automatically converted to text without any format?!
   Sleep 100
   ClipWait 2                    ; Waits until the clipboard contains data, up to 2s
   functionCopy2Clipboard2Txt(functionEMR_TSV_AddLeftArrowUnicode(ClipBoard2Text))
   VarSetCapacity(ClipBoard2Text, 0)      ; Free memory
Return

^#!Numpad6::    ; Ctrl+Win+Alt+Numpad6  	; Swith the numbers & Add → (U+2192) between the numbers
   Send ^c
   Sleep 100
   ClipWait 2                    ; Waits until the clipboard contains data, up to 2s
   ClipBoard2Text = %ClipBoard%            ; Store %ClipBoard% to ClipBoard2Text - Automatically converted to text without any format?!
   Sleep 100
   ClipWait 2                    ; Waits until the clipboard contains data, up to 2s
   functionCopy2Clipboard2Txt(functionEMR_TSV_SwitchValue_AddRightArrowUnicode(ClipBoard2Text))
   VarSetCapacity(ClipBoard2Text, 0)      ; Free memory
Return


/*************************************************************************************************************************
*/

functionFilterLines(TextInput, RegEx2Filter)
{
	TextOutput := ""

	TextOutput := regexreplace(TextInput, "(.*" . RegEx2Filter . ".*)|(.)", "$1")
    TextOutput := TextOutput . "`r`n"
	TextOutput := regexreplace(TextOutput, "(\r?\n)+", "`r`n")
	TextOutput := regexreplace(TextOutput, "^(\r?\n)+", "")

	return TextOutput
}



functionProcessText_lab(TextInput)
{
	TextOutput1 := ""
	TextOutput2 := ""
	TextOutput3 := ""
	TextOutput4 := ""
	Output_functionEMR_TSV_AddLeftArrow 							:= functionEMR_TSV_AddLeftArrow(TextInput)
	Output_functionEMR_TSV_AddLeftArrowUnicode 						:= functionEMR_TSV_AddLeftArrowUnicode(TextInput)
	Output_functionEMR_TSV_SwitchValue_AddRightArrow 				:= functionEMR_TSV_SwitchValue_AddRightArrow(TextInput)
	Output_functionEMR_TSV_SwitchValue_AddRightArrowUnicode 		:= functionEMR_TSV_SwitchValue_AddRightArrowUnicode(TextInput)

    ; https://www.autohotkey.com/docs/Objects.htm#Usage_Simple_Arrays
    ;List := ["WBC", "Seg", "ANC", "Hb", "Hct", "Platelet"]
    ;for key, value in List
    SimpleArray01 := ["FBS", "Glucose", "A1c", "Chol", "TG", "HDL", "LDL",       "AST", "ALT", "Alk", "GGT", "Bil",      "BUN", "Creatinine", "MDRD",        "CPK",        "CRP", "ESR",       "WBC", "Seg", "ANC", "Hb", "Hct", "Platelet",     "TSH", "T4", "T3"]
    for index, value in SimpleArray01
    {
      var4search := value
      %var4search% := regexreplace(functionFilterLines(Output_functionEMR_TSV_SwitchValue_AddRightArrow, var4search), "(^.* \-\> )(.*)([\w\W]*)", "$2")
      TextOutput1 := TextOutput1 . functionFilterLines(Output_functionEMR_TSV_AddLeftArrow, var4search)
      TextOutput2 := TextOutput2 . functionFilterLines(Output_functionEMR_TSV_AddLeftArrowUnicode, var4search)
      TextOutput3 := TextOutput3 . functionFilterLines(Output_functionEMR_TSV_SwitchValue_AddRightArrow, var4search)
      TextOutput4 := TextOutput4 . functionFilterLines(Output_functionEMR_TSV_SwitchValue_AddRightArrowUnicode, var4search)
    }

    tmp2 := ""
    tmp2 := tmp2 . "FBS/A1c `t`t" . FBS . "/" . A1c . " `t`t" . chr("0x" . "2190") . " `t`r`n"
    tmp2 := tmp2 . "A1c/FBS `t`t" . A1c . "/" . FBS . " `t`t" . chr("0x" . "2190") . " `t`r`n"
    tmp2 := tmp2 . "Glucose/A1c `t`t" . Glucose . "/" . A1c . " `t`t" . chr("0x" . "2190") . " `t`r`n"
    tmp2 := tmp2 . "A1c/Glucose `t`t" . A1c . "/" . Glucose . " `t`t" . chr("0x" . "2190") . " `t`r`n"
    tmp2 := tmp2 . "Chol/TG/HDL/LDL`t`t" . Chol . "/" . TG . "/" . HDL . "/" . LDL . " `t`t" . chr("0x" . "2190") . " `t`r`n"
    tmp2 := tmp2 . "AST/ALT/GGT `t`t" . AST . "/" . ALT . "/" . GGT . " `t`t" . chr("0x" . "2190") . " `t`r`n"
    tmp2 := tmp2 . "TSH/T4/T3 `t`t" . TSH . "/" . T4 . "/" . T3 . " `t`t" . chr("0x" . "2190") . " `t`r`n"
    tmp2 := tmp2 . "BUN/Cr(MDRD) `t`t" . BUN . "/" . Creatinine . "(MDRD:" . MDRD . ")" . " `t`t" . chr("0x" . "2190") . " `t`r`n"
    tmp2 := tmp2 . "WBC-Hb(Hct)-Plt: `t`t" . WBC . "k-" . Hb . "(" . Hct . "%)" . "-" . Platelet . "k" . " `t`t" . chr("0x" . "2190") . " `t`r`n"
    tmp2 := tmp2 . "WBC(Seg)-Hb(Hct)-Plt: `t`t" . WBC . "k(Seg " . Seg . "%)-" . Hb . "(" . Hct . "%)" . "-" . Platelet . "k" . " `t`t" . chr("0x" . "2190") . " `t`r`n"
    tmp2 := tmp2 . "WBC(Seg,ANC)-Hb(Hct)-Plt: `t`t" . WBC . "k(Seg " . Seg . "%,ANC " . ANC . ")-" . Hb . "(" . Hct . "%)" . "-" . Platelet . "k" . " `t`t" . chr("0x" . "2190") . " `t`r`n"
    TextOutput := tmp2 . " `t`r`n======================================= `t`r`n" . TextOutput2 . " `t`r`n---------------------------------------- `t`r`n" . TextOutput4 . " `t`r`n---------------------------------------- `t`r`n" . TextOutput1 . " `t`r`n---------------------------------------- `t`r`n" . TextOutput3 . " `t`r`n#################################### `t`r`n" . TextInput

	return TextOutput
}




/*
WBC-Hb(Hct)-Plt: 		4.71k-15.3(45.5%)-164k
WBC(Seg)-Hb(Hct)-Plt: 		4.71k(Seg 51.4%)-15.3(45.5%)-164k
WBC(Seg,ANC)-Hb(Hct)-Plt: 		4.71k(Seg 51.4%,ANC 2421)-15.3(45.5%)-164k

Glucose		95	 ← 102
Hb A1c		5.2	 ← 5.3
Chol.		171	 ← 198
HDL Chol.		36	 ← 44
LDL Chol.		87	 ← 113
LDL Chol.(계산식)		81	 ← 118
TG		270	 ← 179
WBC		4.71	 ← 5.27
Seg.neut.		51.4	 ← 46.8
ANC		2421	 ← 2466
Hb		15.3	 ← 16.2
Hb A1c		5.2	 ← 5.3
Hct		45.5	 ← 47.0
Platelet		164	 ← 157

Glucose		102	 → 95
Hb A1c		5.3	 → 5.2
Chol.		198	 → 171
HDL Chol.		44	 → 36
LDL Chol.		113	 → 87
LDL Chol.(계산식)		118	 → 81
TG		179	 → 270
WBC		5.27	 → 4.71
Seg.neut.		46.8	 → 51.4
ANC		2466	 → 2421
Hb		16.2	 → 15.3
Hb A1c		5.3	 → 5.2
Hct		47.0	 → 45.5
Platelet		157	 → 164

Glucose		95	 <- 102
Hb A1c		5.2	 <- 5.3
Chol.		171	 <- 198
HDL Chol.		36	 <- 44
LDL Chol.		87	 <- 113
LDL Chol.(계산식)		81	 <- 118
TG		270	 <- 179
WBC		4.71	 <- 5.27
Seg.neut.		51.4	 <- 46.8
ANC		2421	 <- 2466
Hb		15.3	 <- 16.2
Hb A1c		5.2	 <- 5.3
Hct		45.5	 <- 47.0
Platelet		164	 <- 157
*/


/*

  (1) DicaMaxD 250mg/1000unit tab(CaCO3/Cholecalciferol) 1 tab [P.O] daily ++  X28 Weeks
  (2) Actonel 35mg tab(Risedronate) 1 tab [P.O] qw  X28 Weeks
  (3) Zolpidem 1 tab [P.O] qw  X28 days

*/

functionProcessText_EMR_medication(TextInput) ; ^#m::    		; Ctrl+Win+m 	 			; copy & process medications & save to a text file
{
	TextOutput := TextInput . "`r`n"
	TextOutput := regexreplace(TextOutput, "^ *\([0-9]+\) ", "")
	TextOutput := regexreplace(TextOutput, "\n +\([0-9]+\) ", "`n")
	TextOutput := regexreplace(TextOutput, "X[0-9]+ Days", "")
	TextOutput := regexreplace(TextOutput, "X[0-9]+ Weeks", "")
	TextOutput := regexreplace(TextOutput, "X[0-9]+ Months", "")

	return TextOutput
}


functionProcessText_join_MultiLine(TextInput) ; ^#/::    		; Ctrl+Win+/ 	 			; copy & join multi-line text (/) & save to a text file
{
	TextOutput := TextInput
	TextOutput := regexreplace(TextOutput, "^(\r?\n)+", "")
	;TextOutput := regexreplace(TextOutput, " *((\r)?\n)+ *", " / ")
	TextOutput := regexreplace(TextOutput, " *((\r)?\n)+ *", "/")
	;TextOutput := TextOutput . "`r`n"
	TextOutput := regexreplace(TextOutput, "\/$", "")
	return TextOutput
}


functionTSV_ReplaceTab2LeftArrow(TextInput) ; ^#t::    		; Ctrl+Win+t 	 			; copy & replace tab to left arrow & save to a text file
{
	TextOutput := TextInput
	TextLeftArrow := " " . chr("0x" . "2190") . " "
	TextOutput := regexreplace(TextOutput, "\t", TextLeftArrow)
	return TextOutput
}



;([^\n\r\t]*)
;\t
;.*
; https://www.autohotkey.com/board/topic/84732-first-line-of-clipboard/
functionTSV_ExtractRow1(Input_TSV)
{
	TextOutput := SubStr(Input_TSV, 1, InStr(Input_TSV, "`r")-1)
	; TextOutput := "<<" . A_ThisFunc . ">>" . " `t`r`n" . TextOutput
	Return TextOutput
}

;([^\n\r\t]*)
;\t
;.*
functionTSV_ExtractColumn1(Input_TSV)
{
	Input_TSV := Input_TSV . "`r`n"
	TextOutput := regexreplace(Input_TSV, "([^\n\r\t]*)\t.*", "$1") . "`r`n"
	; TextOutput := "<<" . A_ThisFunc . ">>" . " `t`r`n" . TextOutput
	Return TextOutput
}

;([^\n\r\t]*)
;\t
;([^\n\r\t]*)
;(\t[^\n\r\t]*)?
functionTSV_ExtractColumn2(Input_TSV)
{
	Input_TSV := Input_TSV . "`r`n"
	TextOutput := regexreplace(Input_TSV, "([^\n\r\t]*)\t([^\n\r\t]*)(\t[^\n\r\t]*)?[^\n\r]*(\r)?\n", "$2`r`n") . "`r`n"
	; TextOutput := "<<" . A_ThisFunc . ">>" . " `t`r`n" . TextOutput
	Return TextOutput
}

;([^\n\r\t]*)
;\t
;([^\n\r\t]*)
;(?:\t([^\n\r\t]*))?
functionTSV_ExtractColumn3(Input_TSV)
{
	Input_TSV := Input_TSV . "`r`n"
	TextOutput := regexreplace(Input_TSV, "([^\n\r\t]*)\t([^\n\r\t]*)(?:\t([^\n\r\t]*))?", "$3") . "`r`n"
	; TextOutput := "<<" . A_ThisFunc . ">>" . " `t`r`n" . TextOutput
	Return TextOutput
}

functionTSV_ExtractColumn4(Input_TSV)
{
	Input_TSV := Input_TSV . "`r`n"
	TextOutput := regexreplace(Input_TSV, "([^\n\r\t]*)\t([^\n\r\t]*)\t([^\n\r\t]*)(?:\t([^\n\r\t]*))?", "$4") . "`r`n"
	; TextOutput := "<<" . A_ThisFunc . ">>" . " `t`r`n" . TextOutput
	Return TextOutput
}




functionRemoveCarriageReturn(TextInput)
{
	TextOutput := TextInput

	TextOutput := regexreplace(TextOutput, "\r", "")
	TextOutput := regexreplace(TextOutput, "\n", "")
	return TextOutput
}




; https://stackoverflow.com/questions/47983063/autohotkey-extract-text-using-regex
; https://www.autohotkey.com/docs/commands/RegExMatch.htm
; https://www.autohotkey.com/board/topic/117001-regexmatch-multiple-matches/
functionProcessText_EMR_TSV4_ExtractDate(EMR_TSV)
{
	RegExMatch_FoundPos := RegExMatch(EMR_TSV, "[0-9]{4}\-[0-9]{2}\-[0-9]{2}" , RegExMatch_Output, StartingPos := 1)
	; Return RegExMatch_FoundPos
	Return RegExMatch_Output
}






; ################################################################################################################################
; ################################################################################################################################
; ################################################################################################################################
; ################################################################################################################################




functionProcessText_EMR_TSV4_lipid(Input_TSV)
{
	TextOutput3 := ""
	TextOutput4 := ""

    ; https://www.autohotkey.com/docs/Objects.htm#Usage_Simple_Arrays
    ;List := ["WBC", "Seg", "ANC", "Hb", "Hct", "Platelet"]
    ;for key, value in List
    SimpleArray01 := ["`tChol.`t", "TG", "HDL", "LDL Chol.`t"]
    for index, value in SimpleArray01
    {
      var4search := value
      TextOutput3 := TextOutput3 . functionTSV_ExtractColumn3(functionFilterLines(Input_TSV, var4search)) . "/"
      TextOutput3 := functionRemoveCarriageReturn(TextOutput3)
      TextOutput4 := TextOutput4 . functionTSV_ExtractColumn4(functionFilterLines(Input_TSV, var4search)) . "/"
      TextOutput4 := functionRemoveCarriageReturn(TextOutput4)
	  ; MsgBox % functionFilterLines(Input_TSV, var4search)
	  ; MsgBox % functionTSV_ExtractColumn2(functionFilterLines(Input_TSV, var4search))
	  ; MsgBox % TextOutput3
    }

	; TextOutput := TextOutput3 . " `t`t" . chr("0x" . "2190") . " `t`t" . TextOutput4 . " `t`r`n"
	TextOutput3 := regexreplace(TextOutput3, "/$", "")
	TextOutput4 := regexreplace(TextOutput4, "/$", "")
	; TextOutput := TextOutput3 . " `t`t" . chr("0x" . "2190") . " `t`t" . TextOutput4
	TextOutput := TextOutput3 . " " . chr("0x" . "2190") . " " . TextOutput4
	; TextOutput := regexreplace(TextOutput, "/$", "")
	; TextOutput := TextOutput . " `t`t" . chr("0x" . "2190") . " `t`t"
	; TextOutput := "TChol/TG/HDL/LDL : " . TextOutput . " " . chr("0x" . "2190") . " "
	;TextOutput := TextOutput3
	Return TextOutput
}


functionProcessText_EMR_TSV4_glucose(Input_TSV)
{
	TextOutput3 := ""
	TextOutput4 := ""

    ; https://www.autohotkey.com/docs/Objects.htm#Usage_Simple_Arrays
    ;List := ["WBC", "Seg", "ANC", "Hb", "Hct", "Platelet"]
    ;for key, value in List
    SimpleArray01 := ["Glucose", "A1c"]
    for index, value in SimpleArray01
    {
      var4search := value
      TextOutput3 := TextOutput3 . functionTSV_ExtractColumn3(functionFilterLines(Input_TSV, var4search)) . "/"
      TextOutput3 := functionRemoveCarriageReturn(TextOutput3)
      TextOutput4 := TextOutput4 . functionTSV_ExtractColumn4(functionFilterLines(Input_TSV, var4search)) . "/"
      TextOutput4 := functionRemoveCarriageReturn(TextOutput4)
	  ; MsgBox % functionFilterLines(Input_TSV, var4search)
	  ; MsgBox % functionTSV_ExtractColumn2(functionFilterLines(Input_TSV, var4search))
	  ; MsgBox % TextOutput3
    }

	; TextOutput := TextOutput3 . " `t`t" . chr("0x" . "2190") . " `t`t" . TextOutput4 . " `t`r`n"
	TextOutput3 := regexreplace(TextOutput3, "/$", "")
	TextOutput4 := regexreplace(TextOutput4, "/$", "")
	; TextOutput := TextOutput3 . " `t`t" . chr("0x" . "2190") . " `t`t" . TextOutput4
	 TextOutput := TextOutput3 . " " . chr("0x" . "2190") . " " . TextOutput4
	; TextOutput := regexreplace(TextOutput, "/$", "")
	; TextOutput := TextOutput . " `t`t" . chr("0x" . "2190") . " `t`t"
	; TextOutput := "Glc/a1c : " . TextOutput . " " . chr("0x" . "2190") . " "
	;TextOutput := TextOutput3
	Return TextOutput
}


functionProcessText_EMR_TSV4_CBC(Input_TSV)
{
	TextOutput3 := ""
	TextOutput4 := ""

    ; https://www.autohotkey.com/docs/Objects.htm#Usage_Simple_Arrays
    ;List := ["WBC", "Seg", "ANC", "Hb", "Hct", "Platelet"]
    ;for key, value in List
    ;SimpleArray01 := ["WBC", "`tHb`t", "Platelet"]

	WBC := functionRemoveCarriageReturn(functionTSV_ExtractColumn3(functionFilterLines(Input_TSV, "WBC")))
	WBC_pre := functionRemoveCarriageReturn(functionTSV_ExtractColumn4(functionFilterLines(Input_TSV, "WBC")) * 1000)
	Hb := functionRemoveCarriageReturn(functionTSV_ExtractColumn3(functionFilterLines(Input_TSV, "`tHb`t")))
	Hb_pre := functionRemoveCarriageReturn(functionTSV_ExtractColumn4(functionFilterLines(Input_TSV, "`tHb`t")))
	Platelet := functionRemoveCarriageReturn(functionTSV_ExtractColumn3(functionFilterLines(Input_TSV, "Platelet")))
	Platelet_pre := functionRemoveCarriageReturn(functionTSV_ExtractColumn4(functionFilterLines(Input_TSV, "Platelet")))

	WBC := floor(WBC * 1000)
	WBC_pre := floor(WBC_pre * 1000)
	TextOutput3 := WBC . "-" . Hb . "-" . Platelet . "k"
	Return TextOutput3
}


functionProcessText_EMR_TSV4_ANC(Input_TSV)
{
	TextOutput3 := ""
	TextOutput4 := ""

    ; https://www.autohotkey.com/docs/Objects.htm#Usage_Simple_Arrays
    ;List := ["WBC", "Seg", "ANC", "Hb", "Hct", "Platelet"]
    ;for key, value in List
    ;SimpleArray01 := ["WBC", "`tHb`t", "Platelet"]

	ANC := functionRemoveCarriageReturn(functionTSV_ExtractColumn3(functionFilterLines(Input_TSV, "ANC")))
	ANC_pre := functionRemoveCarriageReturn(functionTSV_ExtractColumn4(functionFilterLines(Input_TSV, "ANC")) * 1000)

	If (ANC = "")
	{
		;MsgBox ANC = ""
		TextOutput3 := ""
	}
	If (ANC != "")
	{
		;MsgBox ANC != ""
		TextOutput3 := " ANC " . ANC
	}

	Return TextOutput3
}


functionProcessText_EMR_TSV4_ProcessText_FM(Input_TSV)
{
	TextOutput0 := ""

	TextOutput0 := TextOutput0 . "@ LAB (" . functionProcessText_EMR_TSV4_ExtractDate(functionTSV_ExtractRow1(Input_TSV)) . ")`r`n"

	TextOutput0 := TextOutput0 . "Glc/A1c : " . functionProcessText_EMR_TSV4_glucose(Input_TSV) . "`r`n"

    TextOutput0 := TextOutput0 . "TChol/TG/HDL/LDL : " . functionProcessText_EMR_TSV4_lipid(Input_TSV) . "`r`n"

	TextOutput0 := TextOutput0 . "CBC : " . functionProcessText_EMR_TSV4_CBC(Input_TSV) . functionProcessText_EMR_TSV4_ANC(Input_TSV) . "`r`n"

	Return TextOutput0
}

; ################################################################################################################################
; ################################################################################################################################
; ################################################################################################################################
; ################################################################################################################################

^#l::    		; Ctrl+Win+l 	 			; copy & process labs & save to a text file
;^#!c::    		; Ctrl+Win+Alt+c 			; copy & process labs & save to a text file
   Send ^c
   Sleep 100
   ClipWait 2                    ; Waits until the clipboard contains data, up to 2s
   ClipBoard2Text = %ClipBoard%            ; Store %ClipBoard% to ClipBoard2Text - Automatically converted to text without any format?!
   Sleep 100
   ClipWait 2                    ; Waits until the clipboard contains data, up to 2s
   functionCopy2Clipboard2Txt(functionProcessText_lab(ClipBoard2Text))
   ClipWait 2                    ; Waits until the clipboard contains data, up to 2s
   VarSetCapacity(ClipBoard2Text, 0)      ; Free memory
Return



^#!c::    		; Ctrl+Win+Alt+c 			; copy & process labs & save to a text file
   Send ^c
   Sleep 100
   ClipWait 2                    ; Waits until the clipboard contains data, up to 2s
   ClipBoard2Text = %ClipBoard%            ; Store %ClipBoard% to ClipBoard2Text - Automatically converted to text without any format?!
   Sleep 100
   ClipWait 2                    ; Waits until the clipboard contains data, up to 2s
   functionCopy2Clipboard2Txt(functionProcessText_EMR_TSV4_ProcessText_FM(ClipBoard2Text))
   ClipWait 2                    ; Waits until the clipboard contains data, up to 2s
   VarSetCapacity(ClipBoard2Text, 0)      ; Free memory
   Sleep 500
   Send ^a
   Sleep 100
   Send ^c
   Sleep 100
Return


/*
^#!v::    		; Ctrl+Win+Alt+v 			; (from the copied clipboard) process labs & paste
   ClipBoard2Text = %ClipBoard%            ; Store %ClipBoard% to ClipBoard2Text - Automatically converted to text without any format?!
   Sleep 100
   functionPasteViaClipboardThenRestore(functionProcessText_EMR_BP_Bwt(ClipBoard2Text))
   ClipWait 2                    ; Waits until the clipboard contains data, up to 2s
   VarSetCapacity(ClipBoard2Text, 0)      ; Free memory
Return

^#d::    		; Ctrl+Win+d 	 			; copy & process date & save to a text file
   Send ^c
   Sleep 100
   ClipWait 2                    ; Waits until the clipboard contains data, up to 2s
   ClipBoard2Text = %ClipBoard%            ; Store %ClipBoard% to ClipBoard2Text - Automatically converted to text without any format?!
   Sleep 100
   ClipWait 2                    ; Waits until the clipboard contains data, up to 2s
   functionCopy2Clipboard2Txt(functionProcessText_date(ClipBoard2Text))
   ClipWait 2                    ; Waits until the clipboard contains data, up to 2s
   VarSetCapacity(ClipBoard2Text, 0)      ; Free memory
Return


^#b::    		; Ctrl+Win+b 	 			; copy & process BP / Bwt / BMI & save to a text file
   Send ^c
   Sleep 100
   ClipWait 2                    ; Waits until the clipboard contains data, up to 2s
   ClipBoard2Text = %ClipBoard%            ; Store %ClipBoard% to ClipBoard2Text - Automatically converted to text without any format?!
   Sleep 100
   ClipWait 2                    ; Waits until the clipboard contains data, up to 2s
   functionCopy2Clipboard2Txt(functionProcessText_EMR_BP_Bwt(ClipBoard2Text))
   ClipWait 2                    ; Waits until the clipboard contains data, up to 2s
   VarSetCapacity(ClipBoard2Text, 0)      ; Free memory
Return

*/

^#m::    		; Ctrl+Win+m 	 			; copy & process medications & save to a text file
   Send ^c
   Sleep 100
   ClipWait 2                    ; Waits until the clipboard contains data, up to 2s
   ClipBoard2Text = %ClipBoard%            ; Store %ClipBoard% to ClipBoard2Text - Automatically converted to text without any format?!
   Sleep 100
   ClipWait 2                    ; Waits until the clipboard contains data, up to 2s
   functionCopy2Clipboard2Txt(functionProcessText_EMR_medication(ClipBoard2Text))
   ClipWait 2                    ; Waits until the clipboard contains data, up to 2s
   VarSetCapacity(ClipBoard2Text, 0)      ; Free memory
Return



^#/::    		; Ctrl+Win+/ 	 			; copy & join multi-line text (/) & save to a text file
   Send ^c
   Sleep 100
   ClipWait 2                    ; Waits until the clipboard contains data, up to 2s
   ClipBoard2Text = %ClipBoard%            ; Store %ClipBoard% to ClipBoard2Text - Automatically converted to text without any format?!
   Sleep 100
   ClipWait 2                    ; Waits until the clipboard contains data, up to 2s
   functionCopy2Clipboard2Txt(functionProcessText_join_MultiLine(ClipBoard2Text))
   ClipWait 2                    ; Waits until the clipboard contains data, up to 2s
   VarSetCapacity(ClipBoard2Text, 0)      ; Free memory
Return



^#t::    		; Ctrl+Win+t 	 			; copy & replace tab to left arrow & save to a text file
   Send ^c
   Sleep 100
   ClipWait 2                    ; Waits until the clipboard contains data, up to 2s
   ClipBoard2Text = %ClipBoard%            ; Store %ClipBoard% to ClipBoard2Text - Automatically converted to text without any format?!
   Sleep 100
   ClipWait 2                    ; Waits until the clipboard contains data, up to 2s
   functionCopy2Clipboard2Txt(functionTSV_ReplaceTab2LeftArrow(ClipBoard2Text))
   ClipWait 2                    ; Waits until the clipboard contains data, up to 2s
   VarSetCapacity(ClipBoard2Text, 0)      ; Free memory
Return


/*************************************************************************************************************************
*/

