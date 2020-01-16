;delivery_type, country and sector selector
#SingleInstance force
#NoTrayIcon

; Sector Codes
{
; InfoID codes
	SID1 := 30684047
	SID2 := 30684083
	SID3 := 30684064
	SID4 := 30684054
	SID5 := 30684061
	SID6 := 30684059
	SID7 := 30684069
; Telecom codes
	SID8 := 30684108
	SID9 := 36638648
; Office codes
	SID10 := 53957875
	SID11 := 34232024
	SID12 := 53881821
	SID13 := 53882038
; Consumer codes
	SID14 := 30683741
	SID15 := 30683705
	SID16 := 30683711
	SID17 := 30684003
	SID18 := 30683994
	SID19 := 30683719
; Photo codes
	SID20 := 30684040
	SID21 := 33959346
; Multifunctional codes
	SID22 := 48570207
	SID23 := 53877969
	SID24 := 53878412
	SID25 := 53878203
	SID26 := 30683975
	SID27 := 53878124
	SID28 := 53878295
; Any other codes
	SID29 := 29460552
; Services codes
	SID30 := 53799685
	SID31 := 34936748
	SID32 := 34936764
	SID33 := 51277456
	SID34 := 34936787
; Bags codes
	SID35 := 34131417
; Electrical codes
	SID36 := 73211754
; NonPanels
	SID37 := 27607934 ; furniture
	SID38 := 48567624 ; food equipment
	SID39 := 29668190 ; drinks
	SID40 := 29668192 ; boxed consumer goods
	SID41 := 30684493 ; animals
	SID42 := 29668196 ; sport - old code 82383399
	SID43 := 34174472 ; clothing
	SID44 := 27609067 ; toys
	SID45 := 33959252 ; baby care
	SID46 := 34131413 ; healthcare
	SID47 := 27605653 ; beauty
	SID48 := 73212550 ; tools
	SID49 := 30684292 ; gardening
	SID50 := 30684390 ; MDA
	SID51 := 27607862 ; SDA
	SID52 := 73211755 ; ironmongery
	SID53 := 73222282 ; lighting
	SID54 := 34266005 ; automotive
	SID55 := 53800940 ; detergents
	SID56 := 27607968 ; household
	SID57 := 73210675 ; sanitary
; Nonpanel Office
	SID58 := 30684152
	SID59 := 53881971
	SID60 := 53881895
    SID61 := 30684259
}

; COUNTRY long VARIABLES
{
ID1 := "AE"
ID2 := "AM"
ID3 := "AO"
ID4 := "AR"
ID5 := "AT"
ID6 := "AU"
ID7 := "AZ"
ID8 := "BD"
ID9 := "BE"
ID10 :=	"BG"
ID11 := "BH"
ID12 := "BO"
ID13 := "BR"
ID14 := "BW"
ID15 := "BY"
ID16 := "CA"
ID17 := "CD"
ID18 := "CH"
ID19 := "CI"
ID20 := "CL"
ID21 := "CM"
ID22 := "CN"
ID23 := "CO"
ID24 := "CR"
ID25 := "CY"
ID26 := "CZ"
ID27 := "DE"
ID28 := "DK"
ID29 := "DZ"
ID30 := "EC"
ID31 := "EE"
ID32 := "EG"
ID33 := "ES"
ID34 := "ET"
ID35 := "FI"
ID36 := "FR"
ID37 := "GB"
ID38 := "GE"
ID39 := "GH"
ID40 := "GR"
ID41 := "GT"
ID42 := "HK"
ID43 := "HN"
ID44 := "HR"
ID45 := "HU"
ID46 := "ID"
ID47 := "IE"
ID48 := "IL"
ID49 := "IN"
ID50 := "IQ"
ID51 := "IR"
ID52 := "IT"
ID53 := "JO"
ID54 := "JP"
ID55 := "KE"
ID56 := "KH"
ID57 := "KR"
ID58 := "KW"
ID59 := "KZ"
ID60 := "LA"
ID61 := "LB"
ID62 := "LK"
ID63 := "LT"
ID64 := "LU"
ID65 := "LV"
ID66 := "LY"
ID67 := "MA"
ID68 := "MM"
ID69 := "MO"
ID70 := "MX"
ID71 := "MY"
ID72 := "MZ"
ID73 := "NA"
ID74 := "NG"
ID75 := "NL"
ID76 := "NO"
ID77 := "NZ"
ID78 := "OM"
ID79 := "PA"
ID80 := "PE"
ID81 := "PH"
ID82 := "PK"
ID83 := "PL"
ID84 := "PT"
ID85 := "QA"
ID86 := "RO"
ID87 := "RS"
ID88 := "RU"
ID89 := "SA"
ID90 := "SD"
ID91 := "SE"
ID92 := "SG"
ID93 := "SI"
ID94 := "SK"
ID95 := "SN"
ID96 := "ST"
ID97 := "SV"
ID98 := "SY"
ID99 := "TH"
ID100 := "TN"
ID101 := "TR"
ID102 := "TW"
ID103 := "TZ"
ID104 := "UA"
ID105 := "UG"
ID106 := "US"
ID107 := "UY"
ID108 := "YE"
ID109 := "ZA"
}

F3::
LookFor := Trim(Clipboard, " `t`r`n")

;~ com excel
FilePath := "W:\CC_FS\CCC\Data In\DIS Team\BrandMapping_RPG Assignment\RPG Assignment2.1.xlsm"	; path for the excel file to search in
wb := ComObjGet(FilePath)	;get access to your file (workbook)

FoundC := wb.Worksheets("LIST_1").Range("D:D").Find(LookFor).Offset(0, -1).Text	; Long distributor name
FoundD := wb.Worksheets("LIST_1").Range("D:D").Find(LookFor).Text				; Short distributor name aka LookFor value's own cell
FoundE := wb.Worksheets("LIST_1").Range("D:D").Find(LookFor).Offset(0, 1).Text	; Country short abbreviation


CoordMode, Screen
MouseGetPos, xpos, ypos

; country short abbreviation selector for webtas main window
If (WinActive("GfK Retail And Technology - Internet Explorer") || WinActive("Identification - Internet Explorer"))
	{
	Sendinput, %FoundE%
	return
	}
; DeliveryType SELECTOR
else if (xpos > 412 and xpos < 690 and ypos > 162 and ypos < 900)
	{
	Sendinput, %FoundC%
	}
; COUNTRY long SELECTOR
else if (xpos > 40 and xpos < 211 and ypos > 148 and ypos < 900)
	{
	if (FoundE = ID1)	
	Sendinput, UNITED ARAB EMIRATES
	else if (FoundE =  ID2)	
	Sendinput, ARMENIA
	else if (FoundE =  ID3)	
	Sendinput, ANGOLA
	else if (FoundE =  ID4)	
	Sendinput, ARGENTINA
	else if (FoundE =  ID5)	
	Sendinput, AUSTRIA
	else if (FoundE =  ID6)	
	Sendinput, AUSTRALIA
	else if (FoundE =  ID7)	
	Sendinput, AZERBAIJAN
	else if (FoundE =  ID8)	
	MSGBOX, "MISSING"
	else if (FoundE =  ID9)	
	Sendinput, BELGIUM
	else if (FoundE =  ID10)	
	Sendinput, BULGARIA
	else if (FoundE =  ID11)	
	Sendinput, BAHRAIN
	else if (FoundE =  ID12)	
	Sendinput, BOLIVIA
	else if (FoundE =  ID13)	
	Sendinput, BRAZIL
	else if (FoundE =  ID14)	
	MSGBOX, "MISSING"
	else if (FoundE =  ID15)	
	Sendinput, BELARUS
	else if (FoundE =  ID16)	
	Sendinput, CANADA
	else if (FoundE =  ID17)	
	Sendinput, CONGO
	else if (FoundE =  ID18)	
	Sendinput, SWITZERLAND
	else if (FoundE =  ID19)	
	MSGBOX, "MISSING"
	else if (FoundE =  ID20)	
	Sendinput, CHILE
	else if (FoundE =  ID21)	
	Sendinput, CAMEROON
	else if (FoundE =  ID22)	
	Sendinput, CHINA
	else if (FoundE =  ID23)	
	Sendinput, COLOMBIA
	else if (FoundE =  ID24)	
	Sendinput, COSTA RICA
	else if (FoundE =  ID25)	
	Sendinput, CYPRUS
	else if (FoundE =  ID26)	
	Sendinput, CZECH REPLUBLIC
	else if (FoundE =  ID27)	
	Sendinput, GERMANY
	else if (FoundE =  ID28)	
	Sendinput, DENMARK
	else if (FoundE =  ID29)	
	Sendinput, ALGERIA
	else if (FoundE =  ID30)	
	Sendinput, EQUADOR
	else if (FoundE =  ID31)	
	Sendinput, ESTONIA
	else if (FoundE =  ID32)	
	Sendinput, EGYPT
	else if (FoundE =  ID33)	
	Sendinput, SPAIN
	else if (FoundE =  ID34)	
	MSGBOX, "MISSING"
	else if (FoundE =  ID35)	
	Sendinput, FINLAND
	else if (FoundE =  ID36)	
	Sendinput, FRANCE
	else if (FoundE =  ID37)	
	Sendinput, GREAT BRITAIN
	else if (FoundE =  ID38)	
	Sendinput, GEORGIA
	else if (FoundE =  ID39)	
	Sendinput, GHANA
	else if (FoundE =  ID40)	
	Sendinput, GREECE
	else if (FoundE =  ID41)	
	Sendinput, GUATEMALA
	else if (FoundE =  ID42)	
	Sendinput, HONG KONG
	else if (FoundE =  ID43)	
	Sendinput, HUNDURAS
	else if (FoundE =  ID44)	
	Sendinput, CROATIA
	else if (FoundE =  ID45)	
	Sendinput, HUNGARY
	else if (FoundE =  ID46)	
	Sendinput, INDONESIA
	else if (FoundE =  ID47)	
	Sendinput, IRELAND
	else if (FoundE =  ID48)	
	Sendinput, ISRAEL
	else if (FoundE =  ID49)	
	Sendinput, INDIA
	else if (FoundE =  ID50)	
	Sendinput, IRAQ
	else if (FoundE =  ID51)	
	Sendinput, IRAN
	else if (FoundE =  ID52)	
	Sendinput, ITALY
	else if (FoundE =  ID53)	
	Sendinput, JORDAN
	else if (FoundE =  ID54)	
	Sendinput, JAPAN
	else if (FoundE =  ID55)	
	Sendinput, KENYA
	else if (FoundE =  ID56)	
	Sendinput, CAMBODIA
	else if (FoundE =  ID57)	
	Sendinput, KOREA
	else if (FoundE =  ID58)	
	Sendinput, KUWAIT
	else if (FoundE =  ID59)	
	Sendinput, KAZAKHSTAN
	else if (FoundE =  ID60)	
	MSGBOX, "MISSING"
	else if (FoundE =  ID61)	
	Sendinput, LEBANON
	else if (FoundE =  ID62)	
	Sendinput, SRI LANKA
	else if (FoundE =  ID63)	
	Sendinput, LITHUANIA
	else if (FoundE =  ID64)	
	Sendinput, LUXEMBOURG
	else if (FoundE =  ID65)	
	Sendinput, LITVA
	else if (FoundE = ID66)	
	Sendinput, LIBYA
	else if (FoundE =  ID67)	
	MSGBOX, "MISSING"
	else if (FoundE =  ID68)	
	Sendinput, MYANMAR
	else if (FoundE =  ID69)	
	Sendinput, MACAU
	else if (FoundE =  ID70)	
	Sendinput, MEXICO
	else if (FoundE =  ID71)	
	Sendinput, MALAYSIA
	else if (FoundE =  ID72)	
	MSGBOX, "MISSING"
	else if (FoundE =  ID73)	
	Sendinput, NAMIBIA
	else if (FoundE =  ID74)	
	MSGBOX, "MISSING"
	else if (FoundE =  ID75)	
	Sendinput, NETHERLANDS
	else if (FoundE =  ID76)	
	Sendinput, NORWAY
	else if (FoundE =  ID77)	
	Sendinput, NEW ZEALAND
	else if (FoundE =  ID78)	
	Sendinput, OMAN
	else if (FoundE =  ID79)	
	Sendinput, PANAMA
	else if (FoundE =  ID80)	
	Sendinput, PERU
	else if (FoundE =  ID81)	
	Sendinput, PHILIPPINES
	else if (FoundE =  ID82)	
	Sendinput, PAKISTAN
	else if (FoundE =  ID83)	
	Sendinput, POLAND
	else if (FoundE =  ID84)	
	Sendinput, PORTUGAL
	else if (FoundE =  ID85)	
	Sendinput, QATAR
	else if (FoundE =  ID86)	
	Sendinput, ROMANIA
	else if (FoundE =  ID87)	
	Sendinput, SERBIA
	else if (FoundE =  ID88)	
	Sendinput, RUSSIA
	else if (FoundE =  ID89)	
	Sendinput, SAUDI ARABIA
	else if (FoundE =  ID90)	
	MSGBOX, "MISSING"
	else if (FoundE =  ID91)	
	Sendinput, SWEDEN
	else if (FoundE =  ID92)	
	Sendinput, SINGAPORE
	else if (FoundE =  ID93)	
	Sendinput, SLOVENIA
	else if (FoundE =  ID94)	
	Sendinput, SLOVAKIA
	else if (FoundE =  ID95)	
	Sendinput, SENEGAL
	else if (FoundE =  ID96)	
	Sendinput, SAO TOME
	else if (FoundE =  ID97)	
	Sendinput, EL SALVADOR
	else if (FoundE =  ID98)	
	MSGBOX, "MISSING"
	else if (FoundE =  ID99)	
	Sendinput, THAILAND
	else if (FoundE =  ID100)	
	Sendinput, TUNISIA
	else if (FoundE =  ID101)	
	Sendinput, TURKEY
	else if (FoundE =  ID102)	
	Sendinput, TAIWAN
	else if (FoundE =  ID103)	
	Sendinput, TANZANIA
	else if (FoundE =  ID104)	
	Sendinput, UKRAINE
	else if (FoundE =  ID105)	
	Sendinput, UGANDA
	else if (FoundE =  ID106)	
	Sendinput, UNITED STATES
	else if (FoundE =  ID107)	
	Sendinput, URUGUAY
	else if (FoundE =  ID108)	
	Sendinput, YEMEN
	else if (FoundE =  ID109)	
	Sendinput, SOUTH AFRICA
	}
; Sector selector
else if (xpos > 1366 and xpos < 1598 and ypos > 236 and ypos < 965)
	{
	if (LookFor = SID1)
		Send, inform
	else if (LookFor = SID2)
		Send, inform
	else if (LookFor = SID3)
		Send, inform
	else if (LookFor = SID4)
		Send, inform
	else if (LookFor = SID5)
		Send, inform
	else if (LookFor = SID6)
		Send, inform
	else if (LookFor = SID7)
		Send, inform
	else if (LookFor = SID8)
		Send, telecom
	else if (LookFor = SID9)
		Send, telecom
	else if (LookFor = SID10)
		Send, off
	else if (LookFor = SID11)
		Send, off
	else if (LookFor = SID12)
		Send, off
	else if (LookFor = SID13)
		Send, off
	else if (LookFor = SID14)
		Send, co
	else if (LookFor = SID15)
		Send, co
	else if (LookFor = SID16)
		Send, co
	else if (LookFor = SID17)
		Send, co
	else if (LookFor = SID18)
		Send, co
	else if (LookFor = SID19)
		Send, co
	else if (LookFor = SID20)
		Send, pho
	else if (LookFor = SID21)
		Send, pho
	else if (LookFor = SID22)
		Send, mul
	else if (LookFor = SID23)
		Send, mul
	else if (LookFor = SID24)
		Send, mul
	else if (LookFor = SID25)
		Send, mul
	else if (LookFor = SID26)
		Send, mul
	else if (LookFor = SID27)
		Send, mul
	else if (LookFor = SID28)
		Send, mul
	else if (LookFor = SID29)
		Send, aa
	else if (LookFor = SID30)
		Send, ser
	else if (LookFor = SID31)
		Send, ser
	else if (LookFor = SID32)
		Send, ser
	else if (LookFor = SID33)
		Send, ser
	else if (LookFor = SID34)
		Send, ser
	else if (LookFor = SID35)
		Send, per
	else if (LookFor = SID36)
		Send, ele
	else if (LookFor = SID37)
		Send, furniture
	else if (LookFor = SID38)
		Send, food
	else if (LookFor = SID39)
		Send, drink
	else if (LookFor = SID40)
		Send, box
	else if (LookFor = SID41)
		Send, pet
	else if (LookFor = SID42)
		Send, sport
	else if (LookFor = SID43)
		Send, clo
	else if (LookFor = SID44)
		Send, toys
	else if (LookFor = SID45)
		Send, baby
	else if (LookFor = SID46)
		Send, health
	else if (LookFor = SID47)
		Send, beauty
	else if (LookFor = SID48)
		Send, tool
	else if (LookFor = SID49)
		Send, gard
	else if (LookFor = SID50)
		Send, maj
	else if (LookFor = SID51)
		Send, small
	else if (LookFor = SID52)
		Send, iron
	else if (LookFor = SID53)
		Send, light
	else if (LookFor = SID54)
		Send, auto
	else if (LookFor = SID55)
		Send, det
	else if (LookFor = SID56)
		Send, house
	else if (LookFor = SID57)
		Send, san
	else if (LookFor = SID58)
		Send, off
	else if (LookFor = SID59)
		Send, off
	else if (LookFor = SID60)
		Send, off
	else if (LookFor = SID61)
		Send, telecom
	}

else
	{
	MsgBox % "code missing"
	return
	}
	
sleep, 500
Send {Enter}
return