;setworkingdir %A_ScriptDir%
#SingleInstance force
#NoTrayIcon

CoordMode Screen

;GUI buttons
Gui, Add, Button, x60 y150    w49 h23 gDRON, DRON
Gui, Add, Button, x60 y250    w49 h23 gPH.CON, PH.CON
Gui, Add, Button, x0 y250     w49 h23 gM.ENH, M.ENH
Gui, Add, Button, x60 y90     w49 h23 gHFONS, HFONS
Gui, Add, Button, x60 y60     w49 h23 gDMS, DMS
Gui, Add, Button, x60 y30     w49 h23 gCABL, CABL
Gui, Add, Button, x60 y0      w49 h23 gPWER, PWER
Gui, Add, Button, x60 y180    w49 h23 gGPS, GPS
Gui, Add, Button, x270 y280   w49 h23 gAO, AO
Gui, Add, Button, x0 y30      w49 h23 gHARD, HARD
Gui, Add, Button, x60 y220    w49 h23 gPHOTO, PHOTO
Gui, Add, Button, x0 y220     w49 h23 gTELE, TELE
Gui, Add, Button, x0 y180     w49 h23 gSOFT, SOFT
Gui, Add, Button, x0 y150     w49 h23 gERGO, ERGO
Gui, Add, Button, x0 y120     w49 h23 gINPUT, INPUT
Gui, Add, Button, x0 y90      w49 h23 gNET, NET
Gui, Add, Button, x0 y60      w49 h23 gMVIS, MVIS
Gui, Add, Button, x180 y150   w49 h23 gW.COR, W.COR
Gui, Add, Button, x180 y180   w49 h23 gPOST, POST
Gui, Add, Button, x180 y210   w49 h23 gARCIV, ARCIV
Gui, Add, Button, x120 y0     w49 h23 gCE.ACC, CE.ACC
Gui, Add, Button, x0 y280     w49 h23 gTARIF, TARIF
Gui, Add, Button, x60 y120    w49 h23 gREC, REC
Gui, Add, Button, x330 y280   w49 h23 gBAGS, BAGS
Gui, Add, Button, x390 y280   w49 h23 gEL.EQ, EL.EQ
Gui, Add, Button, x180 y0     w49 h23 gPAPER, PAPER
Gui, Add, Button, x180 y30    w49 h23 gOFFICE, OFFICE
Gui, Add, Button, x180 y60    w49 h23 gCTR, CTR
Gui, Add, Button, x180 y90    w49 h23 gPRES, PRES
Gui, Add, Button, x270 y0     w49 h23 gFUD.EQ, FUD.EQ
Gui, Add, Button, x120 y150   w49 h23 gTV.CC, TV.CC
Gui, Add, Button, x120 y120   w49 h23 gCAR, CAR
Gui, Add, Button, x120 y90    w49 h23 gVG, VG
Gui, Add, Button, x120 y60    w49 h23 gPORT, PORT
Gui, Add, Button, x120 y30    w49 h23 gSTAT, STAT
Gui, Add, Button, x390 y90    w49 h23 gSDA, SDA
Gui, Add, Button, x270 y120   w49 h23 gSPORT, SPORT
Gui, Add, Button, x270 y90    w49 h23 gPET, PET
Gui, Add, Button, x270 y60    w49 h23 gBOXED, BOXED
Gui, Add, Button, x270 y30    w49 h23 gDRINK, DRINK
Gui, Add, Button, x330 y0     w49 h23 gTOYS, TOYS
Gui, Add, Button, x330 y30    w49 h23 gBABY, BABY
Gui, Add, Button, x330 y60    w49 h23 gHEALT, HEALT
Gui, Add, Button, x330 y90    w49 h23 gBUTY, BUTY
Gui, Add, Button, x390 y0     w49 h23 gTOOLS, TOOLS
Gui, Add, Button, x390 y30    w49 h23 gGARD, GARD
Gui, Add, Button, x390 y60    w49 h23 gMDA, MDA
Gui, Add, Button, x270 y180   w49 h23 gFURN, FURN
Gui, Add, Button, x270 y150   w49 h23 gCLOTH, CLOTH
Gui, Add, Button, x120 y220   w49 h23 gSW.SER, SW.SER
Gui, Add, Button, x390 y180   w49 h23 gAUTO, AUTO
Gui, Add, Button, x330 y120   w49 h23 gDETER, DETER
Gui, Add, Button, x330 y150   w49 h23 gHOUSE, HOUSE
Gui, Add, Button, x330 y180   w49 h23 gSANIT, SANIT
Gui, Add, Button, x120 y250   w49 h23 gWRNTY, WRNTY
Gui, Add, Button, x120 y280   w49 h23 gMAINT, MAINT
Gui, Add, Button, x180 y280   w49 h23 gO.SER, O.SER
Gui, Add, Button, x390 y150   w49 h23 gLIGHT, LIGHT
Gui, Add, Button, x390 y120   w49 h23 gIRON, IRON
Gui, Add, Button, x0 y0       w49 h23 gCOMP, COMP

Gui, Add, Button, x390 y235  w49 h23 gHelp, Help

;GUI misc
Gui +LastFound +AlwaysOnTop -Caption +ToolWindow
Gui, Show, x1100 y668 w440 h305, TSRtableGUI
return                                 


;gLabel Routines for buttons           

DRON:
Clipboard := "53878124"
return
PH.CON:
Clipboard := "33959346"
return
M.ENH:
Clipboard := "36638648"
return
HFONS:
Clipboard := "53878203"
return
DMS:
Clipboard := "53878412"
return
CABL:
Clipboard := "53877969"
return
PWER:
Clipboard := "48570207"
return
GPS:
Clipboard := "53878295"
return
AO:
Clipboard := "29460552"
return
HARD:
Clipboard := "30684083"
return
PHOTO:
Clipboard := "30684040"
return
TELE:
Clipboard := "30684108"
return
SOFT:
Clipboard := "30684069"
return
ERGO:
Clipboard := "30684059"
return
INPUT:
Clipboard := "30684061"
return
NET:
Clipboard := "30684054"
return
MVIS:
Clipboard := "30684064"
return
W.COR:
Clipboard := "30684152"
return
POST:
Clipboard := "53881971"
return
ARCIV:
Clipboard := "53881895"
return
CE.ACC:
Clipboard := "30683741"
return
TARIF:
Clipboard := "30684259"
return
REC:
Clipboard := "30683975"
return
BAGS:
Clipboard := "34131417"
return
EL.EQ:
Clipboard := "73211754"
return
PAPER:
Clipboard := "53957875"
return
OFFICE:
Clipboard := "34232024"
return
CTR:
Clipboard := "53881821"
return
PRES:
Clipboard := "53882038"
return
FUD.EQ:
Clipboard := "48567624"
return
TV.CC:
Clipboard := "30683719"
return
CAR:
Clipboard := "30683994"
return
VG:
Clipboard := "30684003"
return
PORT:
Clipboard := "30683711"
return
STAT:
Clipboard := "30683705"
return
SDA:
Clipboard := "27607862"
return
SPORT:
Clipboard := "29668196"
return
PET:
Clipboard := "30684493"
return
BOXED:
Clipboard := "29668192"
return
DRINK:
Clipboard := "29668190"
return
TOYS:
Clipboard := "27609067"
return
BABY:
Clipboard := "33959252"
return
HEALT:
Clipboard := "34131413"
return
BUTY:
Clipboard := "27605653"
return
TOOLS:
Clipboard := "73212550"
return
GARD:
Clipboard := "30684292"
return
MDA:
Clipboard := "30684390"
return
FURN:
Clipboard := "27607934"
return
CLOTH:
Clipboard := "34174472"
return
SW.SER:
Clipboard := "34936748"
return
AUTO:
Clipboard := "34266005"
return
DETER:
Clipboard := "53800940"
return
HOUSE:
Clipboard := "27607968"
return
SANIT:
Clipboard := "73210675"
return
WRNTY:
Clipboard := "34936764"
return
MAINT:
Clipboard := "51277456"
return
O.SER:
Clipboard := "34936787"
return
LIGHT:
Clipboard := "73222282"
return
IRON:
Clipboard := "73211755"
return
COMP:
Clipboard := "30684047"
return

Help:
Run, W:\CC_FS\CCC\Data In\DIS Team\BrandMapping_RPG Assignment\S backup\GTDC_TSR_1.2.xlsm
return