; RawDataDownloader
; BY SIMEON P. TODOROV
; VERSION 1.1
; UPDATED TO BE SMART AND REQUIRE LESS ENTRIES LIKE AUTOMATIC RECOGNITION OF COUNTRY AND DISTRIBUTOR CODE
; Version 2.0
; 05.09.2019 - replaced .bat WITH AHK CODE INSTEAD. Deleted old code.
;			 - renamed script, as it doesnt use .bat file anymore
; Version 2.1
; 13.09.2019 - added progress bar window
;
; Version 2.2
; 14.10.2019 - removed dependency of if/else. No hardcoded database and no need for updating when new distributor is loaded.
;			 - added fix for GB - UK edge case
; 2019


#Singleinstance, force

; progress bar globals
global WM_USER               := 0x00000400
global PBM_SETMARQUEE        := WM_USER + 10
global PBM_SETSTATE          := WM_USER + 16
global PBS_MARQUEE           := 0x00000008
global PBS_SMOOTH            := 0x00000001


; get data from clip
split := StrSplit(Clipboard, ",")
varCode2 := split[1]
cntry := split[2]
distShort := split[3]

If (cntry = "GB")
	cntry := "UK"

; OPEN UI
Week := SubStr(A_YWeek, -3)-1													; Previous week as default for the inputbox prompt.
InputBox, period , , Period for %distShort%, , 150, 150, , , , , %Week%

; CLOSE/EXIT BUTTON STOPS THE SCRIPT
If ErrorLevel
	{
	ButtonCancel:
	ExitApp
	}
else
	
Gui, Add, Text, xm ym 0x200, Downloading RawData. Please Wait...
Gui, Add, Progress, w200 h20 hwndMARQ4 -%PBS_SMOOTH% +%PBS_MARQUEE%, 50
DllCall("User32.dll\SendMessage", "Ptr", MARQ4, "Int", PBM_SETMARQUEE, "Ptr", 1, "Ptr", 50)
Gui +LastFound +AlwaysOnTop -Caption +ToolWindow
Gui, Show, AutoSize

; find and copy said files and move them to folder on desktop. Change file extension.
FileCopy, Z:\%cntry%\RDPOOL\Separation_Run\*%varCode2%_%period%*.gf*, C:\Users\simeon.todorov\Desktop\RAW data\*.csv
; Open output folder
Run, C:\Users\simeon.todorov\Desktop\RAW data
Gui, Hide

ExitApp