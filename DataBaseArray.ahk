#SingleInstance force
#NoTrayIcon

; Base variables===============================================================================
{
varCode1 := 
varCode2 := 
varDistributorLong := 
varDistributorShort := 
varCountry := 
empty := ","
}


; Generate startup GUI=========================================================================
{
Gui, Add, Picture, x0 y0 w315, DataBaseArrayBG.png ; BACKGROUND

Gui, Add, Text, vCode1 x50 y15 w30, %varCode1%
Gui, Add, Text, vCode2 x50 y45 w30, %varCode2%
Gui, Add, Text, vDistL x90 y45 w220, %varDistributorLong%
Gui, Add, Text, vDistS x140 y15 w30, %varDistributorShort%
Gui, Add, Text, vCount x180 y15 w20, %varCountry%


Gui, Add, Button, x3 y5 w40 h30 gUpdate , Search
Gui, Add, Button, x3 y35 w40 h30 gCode2 , Raw
;~ Gui, Add, Button, x95 y30 w220 h30 gLong , Long Name
Gui, Add, Button, x90 y5 w40 h30 gShort , Short Name
;~ Gui, Add, Button, x360 y30 w42 h30 , Country

; open sheet for new brands
Gui, Add, Button, x240 y1 w30 h13 gNewBrands, new
; Drag-move panel
Gui, Add, Picture, x1 y1 w240 h5 gUImove, titlebar.png
; copy OEM to clipboard button
Gui, Add, Button, x270 y1 w30 h13 gOEM, oem
;~ Gui, Add, Button, x300 y10 w13 h13 gUpdate, 
Gui, Add, Button, x300 y1 w13 h13 gClose, x

Gui +LastFound +AlwaysOnTop -Caption +ToolWindow
Gui, Show, x880 y7 w315 h70, Search Distributor
return
}



; Button Labels========================================================================
{
Code1:
Clipboard := varCode1
return
Code2:
Clipboard = %varCode2%%empty%%varCountry%%empty%%varDistributorShort%
Run, RawDataDownloader.ahk
return
Long:
Clipboard := varDistributorLong
return
Short:
Clipboard := varDistributorShort
return

UImove:
PostMessage, 0xA1, 2,,, A
return

OEM:
Clipboard := 227637
SplashTextOn, 35, , OEM.
WinMove, OEM, , 1150, 30
Sleep, 1000
SplashTextOff
return

NewBrands:
Run, C:\Users\simeon.todorov\Desktop\Import_Feature_Values_Template.xlsm - Shortcut
return
}


; Update based on new data==============================================================
{
Update:
LookFor := Trim(Clipboard, " `t`r`n")
;~ com excel
FilePath := "W:\CC_FS\CCC\Data In\DIS Team\BrandMapping_RPG Assignment\RPG Assignment2.1.xlsm"	; path for the excel file to search in
wb := ComObjGet(FilePath)	;get access to your file (workbook)
varCode1 := wb.Worksheets("LIST_1").Range("A:A").Find(LookFor).Text				; Code1 aka LookFor's own cell value
varCode2 := wb.Worksheets("LIST_1").Range("A:A").Find(LookFor).Offset(0, 1).Text	; Code 2
varDistributorLong := wb.Worksheets("LIST_1").Range("A:A").Find(LookFor).Offset(0, 2).Text	; DistributorLong name
varDistributorShort := wb.Worksheets("LIST_1").Range("A:A").Find(LookFor).Offset(0, 3).Text	; DistributorShort name
varCountry := wb.Worksheets("LIST_1").Range("A:A").Find(LookFor).Offset(0, 4).Text	; Country short abbreviation

; Update the text fields
GuiControl, Text, Code1, %varCode1%
GuiControl, Text, Code2, %varCode2%
GuiControl, Text, DistL, %varDistributorLong%
GuiControl, Text, DistS, %varDistributorShort%
GuiControl, Text, Count, %varCountry%

Clipboard := varDistributorShort

return
}


; Button Label for app exit====================================================================
Close:
ExitApp