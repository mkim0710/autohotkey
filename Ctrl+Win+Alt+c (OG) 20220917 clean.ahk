#NoEnv
SendMode Input
SetWorkingDir %A_ScriptDir%
#SingleInstance force

if not A_IsAdmin
{
   Run *RunAs "%A_ScriptFullPath%"  ; Requires v1.0.92.01+
   ExitApp
}

Path2Save = % A_ScriptDir . "\" . "AutoHotkeyOutput" . "\"
If !FileExist(Path2Save) {
 FileCreateDir, %Path2Save%
}
return

^#F5:: Exit()
Exit(){
	ExitApp
}

functionAutoHotkeySave(Path2Save)
{
   numb := 0
   Loop, %Path2Save%\AutoHotkeySave*.txt
      If ( SubStr( A_LoopFileName, StrLen("AutoHotkeySave")+1, 4 ) + 1 > numb )
         numb := SubStr( A_LoopFileName, StrLen("AutoHotkeySave")+1, 4 )
   numb := SubStr("0000" (numb+1), -3)
   StringReplace, Clipboard_n, Clipboard, `r`n, `n, all
   FileAppend, %Clipboard_n%, %Path2Save%\AutoHotkeySave%numb%.txt
   Sleep 1000
   Run Notepad.exe `"%Path2Save%\AutoHotkeySave%numb%.txt`"
   VarSetCapacity(Clipboard_n, 0)
}

functionCopy2Clipboard2Txt(Text2CopyPaste)
{
   ClipBoard := Text2CopyPaste
   Sleep 100
   ClipWait 2
   functionAutoHotkeySave(A_ScriptDir . "\" . "AutoHotkeyOutput")
   return Text2CopyPaste
}

functionFilterLines(TextInput, RegEx2Filter)
{
	TextOutput := ""
	TextOutput := regexreplace(TextInput, "(.*" . RegEx2Filter . ".*)|(.)", "$1")
    TextOutput := TextOutput . "`r`n"
	TextOutput := regexreplace(TextOutput, "(\r?\n)+", "`r`n")
	TextOutput := regexreplace(TextOutput, "^(\r?\n)+", "")
	return TextOutput
}

functionTSV_ExtractColumn3(Input_TSV)
{
	Input_TSV := Input_TSV . "`r`n"
	TextOutput := regexreplace(Input_TSV, "([^\n\r\t]*)\t([^\n\r\t]*)(?:\t([^\n\r\t]*))?", "$3") . "`r`n"
	Return TextOutput
}

functionRemoveCarriageReturn(TextInput)
{
	TextOutput := TextInput

	TextOutput := regexreplace(TextOutput, "\r", "")
	TextOutput := regexreplace(TextOutput, "\n", "")
	return TextOutput
}

functionProcessText_EMR_TSV4_ExtractDate(EMR_TSV)
{
	RegExMatch_FoundPos := RegExMatch(EMR_TSV, "[0-9]{4}\-[0-9]{2}\-[0-9]{2}" , RegExMatch_Output, StartingPos := 1)
	Return RegExMatch_Output
}

functionProcessText_EMR_TSV4_lipid(Input_TSV)
{
	TextOutput3 := ""
    SimpleArray01 := ["`tChol.`t", "TG", "HDL", "LDL"]
    for index, value in SimpleArray01
    {
      var4search := value
      TextOutput3 := TextOutput3 . functionRemoveCarriageReturn(functionTSV_ExtractColumn3(functionFilterLines(Input_TSV, var4search))) . "/"
    }
	TextOutput := TextOutput3
	TextOutput3 := regexreplace(TextOutput3, "/$", "")
	Return TextOutput3
}

functionProcessText_EMR_TSV4_CBC(Input_TSV)
{
	TextOutput3 := ""
	WBC := functionRemoveCarriageReturn(functionTSV_ExtractColumn3(functionFilterLines(Input_TSV, "WBC")))
	Hb := functionRemoveCarriageReturn(functionTSV_ExtractColumn3(functionFilterLines(Input_TSV, "`tHb`t")))
	Platelet := functionRemoveCarriageReturn(functionTSV_ExtractColumn3(functionFilterLines(Input_TSV, "Platelet")))
	WBC := floor(WBC * 1000)
	TextOutput3 := WBC . "-" . Hb . "-" . Platelet . "k"
	Return TextOutput3
}

functionProcessText_EMR_TSV4_ANC(Input_TSV)
{
	TextOutput3 := ""
	ANC := functionRemoveCarriageReturn(functionTSV_ExtractColumn3(functionFilterLines(Input_TSV, "ANC")))
	If (ANC = "")
	{
		TextOutput3 := ""
	}
	If (ANC != "")
	{
		TextOutput3 := " ANC " . ANC
	}
	Return TextOutput3
}

functionProcessText_EMR_TSV4_AMH(Input_TSV)
{
	TextOutput3 := ""
    SimpleArray01 := ["`tAMH`t"]
    for index, value in SimpleArray01
    {
      var4search := value
      TextOutput3 := TextOutput3 . functionRemoveCarriageReturn(functionTSV_ExtractColumn3(functionFilterLines(Input_TSV, var4search))) . "/"
    }

	TextOutput := TextOutput3
	TextOutput3 := regexreplace(TextOutput3, "/$", "")
	Return TextOutput3
}

functionProcessText_EMR_TSV4_TSH_Prolactin(Input_TSV)
{
	TextOutput3 := ""
    SimpleArray01 := ["TSH", "Prolactin"]
    for index, value in SimpleArray01
    {
      var4search := value
      TextOutput3 := TextOutput3 . functionRemoveCarriageReturn(functionTSV_ExtractColumn3(functionFilterLines(Input_TSV, var4search))) . "/"
    }

	TextOutput := TextOutput3
	TextOutput3 := regexreplace(TextOutput3, "/$", "")
	Return TextOutput3
}


functionProcessText_EMR_TSV4_LH_FSH_E2(Input_TSV)
{
	TextOutput3 := ""
    SimpleArray01 := ["`tLH`t", "FSH", "Estradiol"]
    for index, value in SimpleArray01
    {
      var4search := value
      TextOutput3 := TextOutput3 . functionRemoveCarriageReturn(functionTSV_ExtractColumn3(functionFilterLines(Input_TSV, var4search))) . "/"
    }

	TextOutput := TextOutput3
	TextOutput3 := regexreplace(TextOutput3, "/$", "")
	Return TextOutput3
}

functionProcessText_EMR_TSV4_SHBG_T_17aPG_DHEAS(Input_TSV)
{
	TextOutput3 := ""
    SimpleArray01 := ["SHBG", "Testosterone", "17-A-PG", "DHEA-S"]
    for index, value in SimpleArray01
    {
      var4search := value
      TextOutput3 := TextOutput3 . functionRemoveCarriageReturn(functionTSV_ExtractColumn3(functionFilterLines(Input_TSV, var4search))) . "/"
    }

	TextOutput := TextOutput3
	TextOutput3 := regexreplace(TextOutput3, "/$", "")
	Return TextOutput3
}

functionProcessText_EMR_TSV4_ProcessText_OG(Input_TSV)
{
	TextOutput0 := ""
	TextOutput0 := TextOutput0 . functionProcessText_EMR_TSV4_ExtractDate(functionFilterLines(Input_TSV, "LDL")) . "] " . "Lipid> " . functionProcessText_EMR_TSV4_lipid(Input_TSV)
	TextOutput0 := TextOutput0 . "`r`n"
	TextOutput0 := TextOutput0 . functionProcessText_EMR_TSV4_ExtractDate(functionFilterLines(Input_TSV, "`tHb`t")) . "] " . "CBC> " . functionProcessText_EMR_TSV4_CBC(Input_TSV) . functionProcessText_EMR_TSV4_ANC(Input_TSV)
	TextOutput0 := TextOutput0 . "`r`n"
	TextOutput0 := TextOutput0 . functionProcessText_EMR_TSV4_ExtractDate(functionFilterLines(Input_TSV, "`tAMH`t")) . "] " . "AMH> " . functionProcessText_EMR_TSV4_AMH(Input_TSV)
	TextOutput0 := TextOutput0 . "`r`n"
	TextOutput0 := TextOutput0 . functionProcessText_EMR_TSV4_ExtractDate(functionFilterLines(Input_TSV, "`tLH`t")) . "] " . "LH/FSH/E2> " . functionProcessText_EMR_TSV4_LH_FSH_E2(Input_TSV)
	TextOutput0 := TextOutput0 . "`r`n"
	TextOutput0 := TextOutput0 . functionProcessText_EMR_TSV4_ExtractDate(functionFilterLines(Input_TSV, "Prolactin")) . "] " . "TSH/PRL> " . functionProcessText_EMR_TSV4_TSH_Prolactin(Input_TSV)
	TextOutput0 := TextOutput0 . "`r`n"
	TextOutput0 := TextOutput0 . functionProcessText_EMR_TSV4_ExtractDate(functionFilterLines(Input_TSV, "DHEA-S")) . "] " . "SHBG/Testosterone/17aPG/DHEAS> " . functionProcessText_EMR_TSV4_SHBG_T_17aPG_DHEAS(Input_TSV)
	TextOutput0 := TextOutput0 . "`r`n"
	Return TextOutput0
}


^#!c::
   Send ^c
   Sleep 100
   ClipWait 2
   ClipBoard2Text = %ClipBoard%
   Sleep 100
   ClipWait 2
   functionCopy2Clipboard2Txt(functionProcessText_EMR_TSV4_ProcessText_OG(ClipBoard2Text))
   ClipWait 2
   VarSetCapacity(ClipBoard2Text, 0)
   Sleep 500
   Send ^a
   Sleep 100
   Send ^c
   Sleep 100
Return

; by MHK on 2022-09-17
