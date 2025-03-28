; AutoHotkey v1.1
; This is a comment

; get main monitor dimensions (default 1920 x 1080)
monitorX = %A_ScreenWidth%
monitorY = %A_ScreenHeight%

ToolTip, %monitorX% %monitorY%
Sleep 2000
ToolTip

!^e::Edit
!^r::Reload

	; ---------------------------------------------------

Click(x, y)	; instant click at x, y coordinates
{
	MouseMove %x%, %y%, 0
	Click
}

FindColor(startX, startY, endX, endY, color, ByRef Cx, ByRef Cy)
{
	PixelSearch, Cx, Cy, %startX%, %startY%, %endX%, %endY%, %color%, 3, Fast 
	if ErrorLevel = 0 
		return 1
}

SearchMove(word, startX, startY, endX, endY, color:=0x3296FF, xOff:=0, yOff:=0, delay:=0)
{
	Send {F3}
	SendInput, %word%
	sleep %delay%
	if(FindColor(startX, startY, endX, endY, color, Cx, Cy))
	{
		Cx += xOff
		Cy += yOff
		MouseMove %Cx%, %Cy%, 0
		return 1
	}
}

SearchClick(word, startX, startY, endX, endY, color:=0x3296FF, xOff:=0, yOff:=0, delay:=100)
{
	if(SearchMove(word, startX, startY, endX, endY, color, xOff, yOff, delay)) 
	{
		Click
		return 1
	}
}

	; ---------------------------------------------------

#IfWinActive Case Reporting System.xlsx
F1::LoadDAWNPt(dawnArray)

F2::
WinActivateBottom, Case Reporting System, , , - Google Sheets
EnterDAWNCRS(dawnArray)
return
	; ---------------------------------------------------

#If WinActive("Case Reporting System")	; both Sheets & DAWN CRS
Esc::
if(FindColor(0, 0, monitorX, monitorY, 0x00BDB4, Cx, Cy))
{
	color := "gray"
	MouseMove %Cx%, %Cy%, 0
}
else
	color := "nothing"

ToolTip, Found %color% at %Cx% %Cy%
sleep 1000
ToolTip
return

!Esc::Send, {Esc}

!s::
SearchClick("add substance", 0, 0, 2000, 2000, 0x3296FF, , , 250)
Send {tab 2}
return

F1::
WinActivateBottom, Case Reporting System.xlsx
Send, {down}
LoadDAWNPt(dawnArray)	; Start from Sheet Date
WinActivateBottom, Case Reporting System, , , - Google Sheets
EnterDAWNCRS(dawnArray)
return

F2::EnterDAWNCRS(dawnArray)

LoadDAWNPt(ByRef dawnArray)
{
	Clipboard := ""
	Send {Home}+{Space}^x
	ClipWait, 0
	dawnArray := StrSplit(clipboard, "`t", "`r")
}

EnterDAWNCRS(dawnArray)
{
	date := dawnArray[1]
	time := dawnArray[2]
	hh := SubStr(time, 1, 2)
	mm := SubStr(time, 3, 2)
	age := dawnArray[5]
	zip := Trim(dawnArray[6])
	sex := dawnArray[7]
	hisp := dawnArray[8]
	raceArray := StrSplit(dawnArray[9], ",", " ")
	summary := StrReplace(dawnArray[10], "+", "{+}")  ; preserves "+"

	substances := StrReplace(dawnArray[11], """")
	Substance := StrSplit(substances, "<", "> ")
	caseType := Substance[2]

	sArray := StrSplit(Substance[1], "`n", "`r")	; how original lol
	drugArray := {}

	For row in sArray 
	{
		sTempArray := StrSplit(sArray[row] , "(", ") ")
		routes := StrSplit(sTempArray[2], ", ")

		drugs := StrSplit(sTempArray[1], ", ")
		
		For key, drug in drugs
		{
			drug := Trim(drug)	; remove space from right side
			StringLower, drug, drug
			drugArray[drug] := routes
		}
	}

; set case type to ALCOHOL ONLY if the only substance is alcohol/etoh/ethanol
	if( (caseType = "") && (drugArray.Count() = 1) && (drugArray.HasKey("etoh") || drugArray.HasKey("alcohol") || drugArray.HasKey("ethanol")) )
		caseType := "alcohol"
	
	diagnoses := StrReplace(dawnArray[12], "+", "{+}")  ; preserves "+"
	diagArray := StrSplit(diagnoses, "`n", """`r")

	disp := dawnArray[13]
	Clipboard := ""

	Send, {F5}	
	sleep 1500
	if(FindColor(0, 0, 2000, 2000, 0x00BDB4, Cx, Cy))
	{
		Click(Cx, Cy+100)
		Send {tab}
	}

; input date and time
	Send, %date%{Tab 3}%hh%{Tab}%mm%{Tab}
	if (hh > 0 and hh <= 12)
		Send, {Down 2}

; input age
	Send, {Tab 2}%age%{Tab 3}

; input zip code or living situation
	Switch zip
	{
		case "homeless": Send, {Tab 2}{Space}
		case "Jail", "institution": Send, {Tab 2}{Down}
		case "n/d": Send, {Tab 2}{Down 3}
		case "unknown": Send, {Tab 2}{Down 3}
		
		default: Send, %zip%{Tab}{Space}
			sleep 700
			if FindColor(100, 200, 300, 400, 0x9C9B9B, Cx, Cy)	
			{
				Send, {Enter}
				sleep 200
				Send, {Tab}{Down 3}
			}
			else
				Send, {Tab}
	}

; select Sex of pt
	Send, {Tab 2}{Space}	;	Male default
	Switch (sex) {
		Case "F":	Send, {Down}
		Case "O":	Send, {Down 2}
		Case "n/d":	Send, {Down 3}
	}

; select Ethnicity (Hispanic or not)
	Send, {Tab 2}{Down}	;	Not Hisp/Lat default
	Switch (hisp) {
		Case "Yes":	Send, {Up}
		Case "n/d":	Send, {Down}
	}
	Send, {Tab 9}

; select Race(s) of pt
	For index, race in raceArray	
	{
		Switch (race) {
			Case "White":	tabs = 7
			Case "Black":	tabs = 6
			Case "Asian":	tabs = 5
			Case "Alaska":	tabs = 4
			Case "Hawaii":	tabs = 3
			Case "Other":	tabs = 2
			Case "n/d":	tabs = 1
		}
		Send, +{Tab %tabs%}{Space}{Tab %tabs%}
	}
	Send {Enter}		

; input Case Summary
	SendInput, %summary%{Tab 4}
	sleep 500

	if StrLen(summary) >= 300
		sleep 500

; input Substances pt has taken & how each was taken if known
	diagTabs := 0
	For drug, routes in drugArray
	{
		if (routes[1]) {	; if routes >= 1
			for key, route in routes
			{
				diagTabs += AddSubstance(drug, route)
			}
		}
		else 
		{
			diagTabs += AddSubstance(drug)
		}
	}
	Send {Tab %diagTabs%}

; input Diagnosis
	Loop % diagArray.Length()
		Send, {Enter}
	Send {Tab}
	for index, diagnosis in diagArray
		SendInput, %diagnosis%{Tab 2}

; select Case type
	Send, {Tab}
	Switch (caseType) {
		Case "suicide", "suicide attempt": Send, {Space}
		Case "detox": Send, {Down}
		Case "alcohol": Send, {Down 2}
		Case "adverse": Send, {Down 3}
		Case "overmedication": Send, {Down 4}
		Case "malicious": Send, {Up 3}
		Case "accidental": Send, {Up 2}
		Default: Send, {Up}
	}
	
	; nalox/bup/mtd
	Send, {Tab 2}{Up}{Tab 2}{Up}{Tab 2}{Up}{Tab 2}{Space}

; input Disposition
	Switch (disp) {
		Case "5150", "Admit (psych)": Send, {Left 7}
		Case "Admit": Send, {Left 5}
		Case "AMA": Send, {Left 4}
		Case "Died", "Expired": Send, {Left 3}
		Case "Eloped", "LWBT", "LWBS": Send, {Left 2}
		; Case "LWBS": Send, {Left 2}
		Case "n/d": Send, {Left 1}
		Case "ICU": Send, {Right 9}
		Case "Jail": Send, {Down}
		Case "Detox": Send, {Down 2}
		Case "Transfer": Send, {Down 8}
		Case "Transfer (psych)": Send, {Down 6}
		Case "Transfer (subst)": Send, {Down  5}
		; Default: Send, {Space}
	}
	Send {tab}{enter}
}

AddSubstance(substance:="", route:="")
{
	if( SearchClick("add substance", 0, 0, 2000, 2000, 0x3296FF, , , 300) ) 
	{	
		Send, {Tab 2}%substance%
		sleep 1400
		MouseMove 0, +125, 0, R
		Click
		sleep 250
		Send, {Tab}

		Switch (route) {
			Case "oral": Send, {Space}
			Case "IV", "inject": Send, {Down}
			Case "snort", "inhale": Send, {Down 2}
			Case "smoke": Send, {Down 3}
			Case "derm": Send, {Down 4}
			Case "vape": Send, {Up 3}
			Case "other": Send, {Up 2}
			Default: Send, {Up}
		}
		
		return 2
	}
}
