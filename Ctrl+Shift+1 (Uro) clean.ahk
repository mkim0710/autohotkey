#NoEnv
SendMode Input
SetWorkingDir %A_ScriptDir%
#SingleInstance force

if not A_IsAdmin
{
   Run *RunAs "%A_ScriptFullPath%"  ; Requires v1.0.92.01+
   ExitApp
}

Path2Save = % A_ScriptDir . "\" . "Temp" . "\"
If !FileExist(Path2Save) {
 FileCreateDir, %Path2Save%
}

MyDefaultDate := fnInputBoxWithDefaultDate()

return

^#F5:: Exit()
Exit(){
	ExitApp
}


fnPasteViaClipboardThenRestore(Text2Paste)
{
   ;-----------------------------------
   global ClipBoardAll0
   ClipBoardAll0 = %ClipBoardAll%
   Sleep 100
   ClipBoard := Text2Paste
   Sleep 100
   ClipWait 2
   Send ^v
   Sleep 100
   ClipBoard = %ClipBoardAll0%
   Sleep 100
   ClipWait 2
   VarSetCapacity(ClipBoardAll0, 0)
   return Text2Paste
}

fnTemp(Path2Save)
{
   numb := 0
   Loop, %Path2Save%\Temp*.txt
      If ( SubStr( A_LoopFileName, StrLen("Temp")+1, 4 ) + 1 > numb )
         numb := SubStr( A_LoopFileName, StrLen("Temp")+1, 4 )

   numb := SubStr("0000" (numb+1), -3)
   StringReplace, Clipboard_n, Clipboard, `r`n, `n, all
   FileAppend, %Clipboard_n%, %Path2Save%\Temp%numb%.txt
   Sleep 1000
   Run Notepad.exe `"%Path2Save%\Temp%numb%.txt`"
   VarSetCapacity(Clipboard_n, 0)
}

fnCopy2Clipboard2Txt(Text2CopyPaste)
{
   ClipBoard := Text2CopyPaste
   Sleep 100
   ClipWait 2
   fnTemp(A_ScriptDir . "\" . "Temp")
   return Text2CopyPaste
}

fnInputBoxWithDefaultDate()
{
	Title03 = InputBox for fnInputBoxWithDefaultDate()
	DisplayText03 = Enter text for InputBoxText##.`n<Caution> Format in YYYYMMDD.`n<Default> Next Wednesday

	Date4Wednesday := A_Now
	Date4Wednesday += 7 - A_WDay + 3, days
	FormatTime, FormatDate4Wednesday, %Date4Wednesday%, yyyyMMdd

	DefaultText03 := FormatDate4Wednesday
	InputBox, InputBoxText03, %Title03%, %DisplayText03%,,,,,,,, %DefaultText03%

    return InputBoxText03
}

^#!d::
   MyDefaultDate := fnInputBoxWithDefaultDate()
return


/*************************************************************************************************************************
PSA(진검시행) [Serum]
U/A (Stick + Microscopy) PANEL(검사24시간가능) [소변]
Gram stain, Genitourinary specimen culture [Urine, Voided]
Cytospin (비뇨기계 검체)  [Urine]
Bladder Ultrasound Scan(대한외래) 잔뇨검사
요류및잔뇨측정술(외래)
UDS Video, moderate (대한외래) / UDS  Cath + Rectal Cath 포함
72시간 배뇨양상기능검사 - 일반형 (외래)
72시간 배뇨양상기능검사 - 도뇨형 (외래)
TRUS(A)(외래) 전립선초음파
SONO Transrectal (Prostate)
SONO Kidney [ 검사희망일 : 2022-03-23 ]: 신경인성방광 추적용 (bladder wall thickness and hydronephrosis)  [C.I.+]
Panendoscopy ( 방광요도내시경 )
*/

fnFilterLines(TextInput, RegEx2Filter)
{
	TextOutput := ""

	TextOutput := regexreplace(TextInput, "(.*" . RegEx2Filter . ".*)|(.)", "$1")
    TextOutput := TextOutput . "`r`n"
	TextOutput := regexreplace(TextOutput, "(\r?\n)+", "`r`n")
	TextOutput := regexreplace(TextOutput, "^(\r?\n)+", "")

	return TextOutput
}

fnInputBox()
{
	OutputText := fnInputBoxWithAutomaticEscape()
    return OutputText
}


fnInputBoxWithWarnings()
{
	Title01 = InputBox for fnInputBoxWithWarnings()
	DisplayText01 = Enter text for InputBoxText##.`n<Caution> This might need to escape RegEx matacharacters: [{()<>.?\|$^*+`nThis might need to escape special characters (i.e.`, \:).`n``~!@#$`%^`&*()_-=+[]\{}|;':"`,./<>?
    DefaultText01 := "`n"
	InputBox, InputBoxText01, %Title01%, %DisplayText01%,,,,,,,,%DefaultText01%
    return InputBoxText01
}

fnInputBoxWithAutomaticEscape()
{
	Title02 = InputBox for fnInputBoxWithAutomaticEscape()
	DisplayText02 = Enter text for InputBoxText##.`n<Caution> This will automatically process escape characters.`n``~!@#$`%^&*()_-=+[]\{}|;':"`,./<>?
    DefaultText02 := "`n"
	InputBox, InputBoxText02, %Title02%, %DisplayText02%,,,,,,,, %DefaultText02%

    return fnAutomaticEscape(InputBoxText02)
}


fnAutomaticEscape(InputText001)
{
    OutputText001 := RegExReplace(InputText001, "([\s\``\~\!\@\#\$\%\^\&\*\(\)\_\-\=\+\[\]\\\{\}\|\;\'\:\""\,\.\/\<\>\?])", "\$1")
	return OutputText001
}


fnProcessText01(TextInput)
{
	global MyDefaultDate
    TextOutput1 := MyDefaultDate . " "
	TextOutput2 := "20220316 "

    SimpleArray01 := []
    SimpleArray02 := []

    SimpleArray01.Push("PSA")
    SimpleArray02.Push("PSA")

    SimpleArray01.Push("U/A")
    SimpleArray02.Push("UA /")

    SimpleArray01.Push("Genitourinary specimen culture")
    SimpleArray02.Push("UCx()")

    SimpleArray01.Push("Cytospin")
    SimpleArray02.Push("Ucyto()")

    SimpleArray01.Push("MRI")
    SimpleArray02.Push("MRI HN()")

    SimpleArray01.Push("CT")
    SimpleArray02.Push("NECT stone() HN()")

    SimpleArray01.Push("SONO Kidney")
    SimpleArray02.Push("K-US HN()")

    SimpleArray01.Push("SONO Transrectal")
    SimpleArray02.Push("TRUS /")

    SimpleArray01.Push("TRUS")
    SimpleArray02.Push("TRUS /")

    SimpleArray01.Push("Bladder Ultrasound Scan")
    SimpleArray02.Push("PVR")

    SimpleArray01.Push("요류및잔뇨측정술")
    SimpleArray02.Push("UFM //")

    SimpleArray01.Push("UDS")
    SimpleArray02.Push("UDS 요역동학검사")

    SimpleArray01.Push("배뇨양상기능검사 - 일반형")
    SimpleArray02.Push("VD 일일 회/Max /야뇨 회")

    SimpleArray01.Push("배뇨양상 기능검사 \(일반형\)")
    SimpleArray02.Push("VD 일일 회/Max /야뇨 회")

    SimpleArray01.Push("배뇨양상기능검사 - 도뇨형")
    SimpleArray02.Push("VD(도뇨) 일일 회/Max /야뇨 회")

    SimpleArray01.Push("배뇨양상 기능검사 \(도뇨형\)")
    SimpleArray02.Push("VD(도뇨) 일일 회/Max /야뇨 회")

    SimpleArray01.Push("Panendoscopy")
    SimpleArray02.Push("Panendo ")

    for key, value in SimpleArray01
    {
      var4search := value
      if (RegExMatch(TextInput, var4search) > 0) {
         TextOutput1 := TextOutput1 . SimpleArray02[key] . " "
      }
    }

    TextOutput := TextOutput1

	return TextOutput
}


^+1::
   Send ^c
   Sleep 100
   ClipWait 2
   ClipBoard2Text = %ClipBoard%
   Sleep 100
   ClipWait 2
   fnCopy2Clipboard2Txt(fnProcessText01(ClipBoard2Text))
   ClipWait 2
   VarSetCapacity(ClipBoard2Text, 0)
   Sleep 500
   Send ^a
   Send ^c
Return


