; MasterHub for activation of other scripts and misc window interactions and integrations
; ...
; older versions were not tracked
; ...
; !revoked for breaking the webpage - 08.07.2019 - changed "Enter" function for Brand Swap in IExplorer changed to tab2enter instead of hijack the mouse. Perfect reliability in all window sizes.
;
; 24.07.2019 - added F1 link to RPG tracker
; 25.07.2019 - changed PASS glabel to automatically enter credentials and log into the website. netscalerGateway website must be opened prior to execution.
; 03.09.2019 - TSR GUI button now toggles script on/off. Removed redundant close button. Relocated TSR button.
;			 - Removed BATCH runner because of redundancy. Function migrated to Distributor Array script.
;			 - Removed Raw folder button because of redundancy.
;			 - Cleaned code from unused legacy lines
;			 - Removed RPG tracker function from F1 because it will not be used.
;			 - Expanded and renamed NonPanel button. Moved RPG Smart and close buttons. Moved RPGtable and Citrix buttons. Moved UImove button and resized GUI. Renamed MasterHub window.
;			 - Reordered GUI buttons and gLabels code to reflect positions better.
;			 - Stopped window from showing tray icon.
;			 - RPG Smart button now toggles script on/off. Removed redundant close button. Resized RPG Smart button. Added function to display active/disabled on button. RPGsmart script now has hidden icon from the tray. 
;			 - Redone closing sequence.
; 13.09.2019 - Replaced function to Run populateRPGnew.ahk with its new version AutoRPGMover
;
; 18.10.2019 - Removed buttons for script updating and nonpanel list, because they are obsolete
;
; 23.10.2019 - added weather widget using python
;
; created by Simeon P. Todorov
; 2019


SetWorkingDir %A_ScriptDir%
#SingleInstance force
#NoTrayIcon
CoordMode Screen
state := "Off" ; RPGsmart.ahk on/off state
SetTimer, Weather, 1200000 ; update interval 20 minutes for weather widget 

; weather
gosub Weather
Gui, 2:Add, Button, x0 y0 w120 h20 gWeather vWeather, %tempr%
Gui, 2:+LastFound +AlwaysOnTop -Caption +ToolWindow
Gui, 2:Show, x1975 y1000 w120 h20, 2ndgui
hGui2 := WinExist(2ndgui)

;GUI buttons
Gui, 1:Add, Button, x0 y0 w70 h30 gDSearch, Distributor Search
Gui, 1:Add, Button, x70 y0 w70 h30 gTSR, TSR table
Gui, 1:Add, Button, x140 y0 w70 h30 gPopulate , Populate RPG

Gui, 1:Add, Button, x0 y30 w35 h30 gCloseApp, Close
Gui, 1:Add, Button, x35 y30 w35 h30 gReloadApp, Rld
Gui, 1:Add, Button, x70 y30 w70 h30 vRPGsmart gRPGsmart , 
Gui, 1:Add, Button, x140 y30 w35 h30 gRPGtable, RPG table
Gui, 1:Add, Button, x175 y30 w35 h30 gCitrix, Citrix

; Drag-move panel
Gui, 1:Add, Picture, x210 y0 w5 h60 gUImove, titlebar.png

;GUI misc
Gui, 1:+LastFound +AlwaysOnTop -Caption +ToolWindow
Gui, 1:Show, x1922 y1022 w215 h60, MasterHubGUI
hGui1 := WinExist(MasterHubGUI)


;activate RPGsmart at startup
Run, RPGsmart.ahk
state := "On"
GuiControl, Text, RPGsmart, RPG Smart %state%
return

;---------------------------
;F12 - easy copy,paste
Sendmode, input
F1::
Send, ^c   ;copy
return
F2:: ^v   ;paste
return
;---------------------------
;SoundControl
Sendmode, input
F10::SoundSet, -5
return
F11::SoundSet, +5
return
F12::
SoundGet, volume
If (volume = 0)
{
	SoundSet, 25
	ToolTip, Unmute
	sleep, 1000
	ToolTip, 
}
else
{
	SoundSet, 0
	ToolTip, Mute
	sleep, 1000
	ToolTip, 
}
return
;---------------------------
; Enter-Search for BrandSwap window in WebTas
#If (WinActive("Brand Swap - Internet Explorer") || WinActive("Brand - Internet Explorer"))
	Enter::MouseClick, left, 229, 348
return
;---------------------------

;gLabel Routines for buttons
DSearch:
Run, DataBaseArray.ahk
return
TSR:
IfWinExist, TSRtableGUI
	{
		WinClose, TSRtableGUI
	}
else
	{
		Run, TSRGUI.ahk
	}
return
Populate:
Run, AutoRPGMover.ahk
return



RPGsmart:
DetectHiddenWindows, On
SetTitleMatchMode, 2
IfWinExist, RPGsmart.ahk
{
	WinClose, RPGsmart.ahk
	state := "Off"
	GuiControl, Text, RPGsmart, RPG Smart %state%
}
else
{
	Run, RPGsmart.ahk
	state := "On"
	GuiControl, Text, RPGsmart, RPG Smart %state%
}
DetectHiddenWindows, Off
SetTitleMatchMode, 1
return
RPGtable:
Run, C:\Users\simeon.todorov\Desktop\RPG Assignment2.1 - Shortcut
return
Citrix:
Clipboard := "iop963ERT!123"
WinActivate, NetScaler Gateway - Internet Explorer
WinWait, NetScaler Gateway - Internet Explorer
Send, sntodo{Tab}
Send, ^v
Send, {Enter}
return


UImove:
PostMessage, 0xA1, 2,,, ahk_id %hGui1%
loop, 2
{
WinGetPos, x2,y2
x2 += 53
y2 -= 22
Gui, 2:Show, x%x2% y%y2% w120 h20, 2ndgui
}
return

Weather:
Run, C:\Users\simeon.todorov\Desktop\weather.pyw,,, NewPID
Process, WaitClose, %NewPID% ; wait for the python script to finish
FileRead, tempr, weather.txt
GuiControl, 2:Text, Weather, %tempr%
FileDelete, C:\Users\simeon.todorov\Desktop\weather.txt
return
ReloadApp:
Reload
return
CloseApp:
;exit rpgsmart
DetectHiddenWindows, On
SetTitleMatchMode, 2
WinClose, RPGsmart.ahk
ExitApp