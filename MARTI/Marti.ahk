#SingleInstance, Force
CoordMode, Mouse, Screen
CoordMode, Pixel, Screen
CoordMode, ToolTip, Screen
AutoTrim, Off
SetWorkingDir, %A_ScriptDir%
Menu, tray, Icon, %A_ScriptDir%\martilogoico.ico, 1, 1

;--------------------VARIABLES--------------------
FormatTime, date, MMDDYYYY, ShortDate
numberslist:=[2,3,4,5,6,7,8,9,1,0]
lowercaseletterslist:=["a","b","c","d","e","f","h","j","k","m","n","p","r","s","t","w","x","y","z","g","q","i","l","o","u","v"]
uppercaseletterslist:=["A","B","C","D","E","F","H","J","K","M","N","P","R","S","T","W","X","Y","Z","G","Q","I","l","O","U","V"]
notepadtoggle:=1
randomnumberlimit:=10
randomwordlimit:=26
displaypassword:="Random Password"
checkcheck:=1
pinon=%A_ScriptDir%\pinon.png
pinoff=%A_ScriptDir%\pinoff.png
settings=%A_ScriptDir%\options.png
pintoggle:=1
loadtoggle:=1
optiontoggle:=1

;--------------------GUI--------------------
Gui, 1:Add, Text, x18 y3, Name:
Gui, 1:Add, Edit, x50 y0 w125 vname,
Gui, 1:Add, Text, x35 y23, ID:
Gui, 1:Add, Edit, x50 y20 w125 vid,
Gui, 1:Add, Text, x184 y3, Number:
Gui, 1:Add, Edit, x225 y0 w125 vnumber gnumber,
Gui, 1:Add, Text, x209 y23, Alt:
Gui, 1:Add, Edit, x225 y20 w125 Disabled valtnumber,
Gui, 1:Add, Text, x371 y03, Email:
Gui, 1:Add, Edit, x400 y0 w125 vemail gemail,
Gui, 1:Add, Text, x384 y23, Alt:
Gui, 1:Add, Edit, x400 y20 w125 Disabled valtemail,
Gui, 1:Add, Text, x530 y3, Location:
Gui, 1:Add, Edit, x575 y0 w125 vlocation glocation,
Gui, 1:Add, Text, x559 y23, Alt:
Gui, 1:Add, Edit, x575 y20 w125 Disabled valtlocation,
Gui, 1:Add, Text, x704 y03, Manager:
Gui, 1:Add, Edit, x750 y0 w125 vmanager,
Gui, 1:Add, Text, x730 y23, HR:
Gui, 1:Add, Edit, x750 y20 w125 vhr,
Gui, 1:Add, Text, x1 y43, Computer:
Gui, 1:Add, Edit, x50 y40 w125 vcomputer,
Gui, 1:Add, Text, x191 y43, Printer:
Gui, 1:Add, Edit, x225 y40 w125 vprinter,
Gui, 1:Add, Text, x365 y43, Server:
Gui, 1:Add, Edit, x400 y40 w125 vserver,
Gui, 1:Add, Text, x536 y43, Scr Sht:
Gui, 1:Add, Checkbox, x575 y44 w23 vscreenshot,
Gui, 1:Add, Text, x598 y43, Access:
Gui, 1:Add, DropDownList, x637 y40 w56 vaccess gaccess, Yes|No|Unknown||
Gui, 1:Add, Text, x695 y43, When:
Gui, 1:Add, Edit, x728 y40 w55 Disabled vwhen,
Gui, 1:Add, Text, x789 y43, Users:
Gui, 1:Add, Edit, x820 y40 w25 vaffecteduser gaffecteduser, 
Gui, 1:Add, Text, x845 y43, /
Gui, 1:Add, Edit, x850 y40 w25 Disabled vaffectedusers,
Gui, 1:Add, Text, x1 y63, Error Msg:
Gui, 1:Add, Edit, x50 y60 w375 verrormessage,
Gui, 1:Add, Text, x456 y63, Affected:
Gui, 1:Add, Edit, x500 y60 w375 vaffected,
Gui, 1:Add, Text, x1 y83, Requester:
Gui, 1:Add, Edit, x50 y80 w100 grequester vrequester,
Gui, 1:Add, DropDownList, x150 y80 w75 Disabled vrelist, Requesting||Reporting
Gui, 1:Add, Edit, x225 y80 w650 vrequesting,
Gui, 1:Add, Button, x0 y101 w100 h40 vnotepadbutton gnotepadbutton, Open Notepad
Gui, 1:Add, Button, x699 y101 w100 h40 vlogclipboard glogs, Logs
Gui, 1:Add, Button, x800 y101 w100 h40 vrticlipboard grticlipboard, RTI
Gui, 1:Add, Button, x220 y107 w50 h30 gnew, New
Gui, 1:Add, Button, x275 y107 w50 h30 gsave, Save
Gui, 1:Add, Button, x330 y107 w50 h30 gload, Load
Gui, 1:Add, Picture, x877 y1 goptions, %settings%
Gui, 1:Add, Picture, x877 y21 vpin gpin, %pinon%
Gui, 1:Add, Button, x574 y101 h20 w52 gpassword, Generate
Gui, 1:Add, Edit, x626 y101 w20 h20 Limit Number vrandomlength, 10
Gui, 1:Add, Checkbox, x647 y104 Checked vsimplepassword, Simple
Gui, 1:Add, Edit, x575 y121 w122 h19 ReadOnly Center vdisplaypassword, %displaypassword%

Gui, 1:Add, text, x425 y128, %date%
Gui, 1:Add, Edit, x2 y142 w896 h356 vnotepad,

MouseGetPos, mousex, mousey
Gui, 1:show, x%mousex% y%mousey% w900 h141, MARTI
Winset, Alwaysontop, , MARTI

GuiControl, 1:, when,Unknown

Return

;--------------------TIMERS--------------------
2follow1:
Winset, Top,, Fetcher
WinGetPos, martix, martiy,,, MARTI
newx:=martix+903
newy:=martiy+25
Gui, 2:show, NoActivate x%newx% y%newy% w%gui2w% h%gui2h%, Fetcher

Return

;--------------------LABELS--------------------
new:
msgbox, 4, Confirm Choice, Are you sure (This will clear ALL fields)?
IfMsgBox, Yes 
{
	GuiControl, 1:, name, 
	GuiControl, 1:, id, 
	GuiControl, 1:, number, 
	GuiControl, 1:, altnumber, 
	GuiControl, 1:, email, 
	GuiControl, 1:, altemail, 
	GuiControl, 1:, location, 
	GuiControl, 1:, altlocation, 
	GuiControl, 1:, manager, 
	GuiControl, 1:, hr, 
	GuiControl, 1:, computer, 
	
	GuiControl, 1:, printer,
	GuiControl, 1:, server,
	GuiControl, 1:, screenshot, 0
	GuiControl, 1:, access,
	GuiControl, 1: Choose, Access, 3
	GuiControl, 1:, when, Unknown
	GuiControl, 1: Disable, when,
	GuiControl, 1:, affecteduser,
	GuiControl, 1:, errormessage, 
	GuiControl, 1:, affected,
	GuiControl, 1:, requester,
	GuiControl, 1: Choose, relist, 1
	GuiControl, 1:, requesting,
	GuiControl, 1:, notepad,
} else 
	Return
Return

load:
FileSelectFile, iniread, 1 2 8 16 32,, Load Profile, ini (*.ini)
if (iniread=""){
	Return
} else {
	IniRead, ininame, %iniread%, information, Name
	IniRead, iniid, %iniread%, information, ID
	IniRead, ininumber, %iniread%, information, Number
	IniRead, inialtnumber, %iniread%, information, Alt Number
	IniRead, iniemail, %iniread%, information, Email
	IniRead, inialtemail, %iniread%, information, Alt Email
	IniRead, inilocation, %iniread%, information, Location
	IniRead, inialtlocation, %iniread%, information, Alt Location
	IniRead, inimanager, %iniread%, information, Manager
	IniRead, inihr, %iniread%, information, HR
	IniRead, inicomputer, %iniread%, information, Computer(s)
	GuiControl, 1:, name, %ininame%
	GuiControl, 1:, id, %iniid%
	GuiControl, 1:, number, %ininumber%
	GuiControl, 1:, altnumber, %inialtnumber%
	GuiControl, 1:, email, %iniemail%
	GuiControl, 1:, altemail, %inialtemail%
	GuiControl, 1:, location, %inilocation%
	GuiControl, 1:, altlocation, %inialtlocation%
	GuiControl, 1:, manager, %inimanager%
	GuiControl, 1:, hr, %inihr%
	GuiControl, 1:, computer, %inicomputer%
}
Return

save:
GuiControlGet, name
StringLen, namelettercount, name
if (namelettercount<1){
	Return
}
GuiControlGet, id
GuiControlGet, number
GuiControlGet, altnumber
GuiControlGet, email
GuiControlGet, altemail
GuiControlGet, location
GuiControlGet, altlocation
GuiControlGet, manager
GuiControlGet, hr
GuiControlGet, computer
IniWrite, %name%, %A_ScriptDir%\profiles\%name%.ini, information, Name
IniWrite, %id%, %A_ScriptDir%\profiles\%name%.ini, information, ID
IniWrite, %number%, %A_ScriptDir%\profiles\%name%.ini, information, Number
IniWrite, %altnumber%, %A_ScriptDir%\profiles\%name%.ini, information, Alt Number
IniWrite, %email%, %A_ScriptDir%\profiles\%name%.ini, information, Email
IniWrite, %altemail%, %A_ScriptDir%\profiles\%name%.ini, information, Alt Email
IniWrite, %location%, %A_ScriptDir%\profiles\%name%.ini, information, Location
IniWrite, %altlocation%, %A_ScriptDir%\profiles\%name%.ini, information, Alt Location
IniWrite, %manager%, %A_ScriptDir%\profiles\%name%.ini, information, Manager
IniWrite, %hr%, %A_ScriptDir%\profiles\%name%.ini, information, HR
IniWrite, %computer%, %A_ScriptDir%\profiles\%name%.ini, information, Computer(s)
Return

pin:
pintoggle:=!pintoggle
if (pintoggle){
	GuiControl, 1: -Redraw, pin
	GuiControl, 1:, pin, %pinon%
	GuiControl, 1: +Redraw, pin
	Winset, Alwaysontop, , MARTI
} else {
	GuiControl, 1: -Redraw, pin
	GuiControl, 1:, pin, %pinoff%
	GuiControl, 1: +Redraw, pin
	Winset, Alwaysontop, , MARTI
}
Return

checkcycle:
checkcheck:=!checkcheck
GuiControlGet, name
StringLen, namelettercount, name
if (namelettercount<1) {
	namecheck:= "Unknown"
} else {
	namecheck:= name
}
GuiControlGet, id
StringLen, idlettercount, id
if (idlettercount<1) {
	idcheck:= "Unknown"
} else {
	idcheck:= id
}
GuiControlGet, number
StringLen, numberlettercount, number
if (numberlettercount<1){
	numbercheck:="Unknown"
} else {
	numbercheck:= number
	GuiControlGet, altnumber
	StringLen, altnumberlettercount, altnumber
	if (altnumberlettercount<1) {
		altnumbercheck:= ""
	} else {
		altnumbercheck:= ", " altnumber
	}
}
GuiControlGet, email
StringLen, emaillettercount, email
if (emaillettercount<1){
	emailcheck:="Unknown"
} else {
	emailcheck:=email
	GuiControlGet, altemail
	StringLen, altemaillettercount, altemail
	if (altemaillettercount<1) {
		altemailcheck:= ""
	} else {
		altemailcheck:= ", " altemail
	}
}
GuiControlGet, location
StringLen, locationlettercount, location
if (locationlettercount<1) {
	locationcheck:="Unknown"
} else {
	locationcheck:=location
	GuiControlGet, altlocation
	StringLen, altlocationlettercount, altlocation
	if (altlocationlettercount<1) {
		altlocationcheck:= ""
	} else {
		altlocationcheck:= ", " altlocation
	}
}
GuiControlGet, computer
StringLen, computerlettercount, computer
if (computerlettercount<1){
	computercheck:= "Blank SCCM"
} else {
	computercheck:= computer
}
GuiControlGet, printer
StringLen, printerlettercount, printer
if (printerlettercount<1){
	printercheck:= "Not Printer Related"
} else {
	printercheck:= printer
}
GuiControlGet, server
StringLen, serverlettercount, server
if (serverlettercount<1){
	servercheck:= "Not Server Related"
} else {
	servercheck:= server
}
GuiControlGet, relist
GuiControlGet, name
GuiControlGet, requesting
StringLen, requestinglettercount, requesting
if (requestinglettercount<1) {
	requestingcheck:= "Unknown Request/Incident"
} else {
	requestingcheck:= requesting
}
GuiControlGet, requester
StringLen, requesterlettercount, requester
if (requesterlettercount<1){
	requestercheck:= namecheck " - " requestingcheck
} else {
	requestercheck:= requester " " relist " " namecheck " - " requestingcheck
}
GuiControlGet, errormessage
StringLen, errormessagelettercount, errormessage
if (errormessagelettercount<1) {
	errormessagecheck:= "No Error Message(s)"
} else {
	errormessagecheck:= errormessage
}
GuiControlGet, screenshot
if (screenshot=1) {
	screenshotcheck:= "Attached"
	screenshotsummarycheck:= "-Screenshot Attached"
} else {
	screenshotcheck:= "n/a"
	screenshotsummarycheck:= ""
}
GuiControlGet, access
GuiControlGet, when
StringLen, whenlettercount, when
if (whenlettercount<1){
	whencheck:= "Unknown"
} else {
	whencheck:= when
}
GuiControlGet, affecteduser
StringLen, affecteduserlettercount, affecteduser
if (affecteduserlettercount<1){
	affectedusercheck:= "Unknown"
} else {
	affectedusercheck:= affecteduser
	GuiControlGet, affectedusers
	StringLen, affecteduserslettercount, affectedusers
	if (affecteduserslettercount<1) {
		affecteduserscheck:= ""
	} else {
		affecteduserscheck:="/ " affectedusers
	}
}
GuiControlGet, affected
StringLen, affectedlettercount, affected
if (affectedlettercount<1) {
	affectedcheck:= "Unknown"
} else {
	affectedcheck:= affected
}
GuiControlGet, hr
StringLen, hrlettercount, hr
if (hrlettercount<1) {
	hrcheck:= "Unknown"
} else {
	hrcheck:= hr
}
GuiControlGet, manager
StringLen, managerlettercount, manager
if (managerlettercount<1){
	managercheck:= "Unknown"
} else {
	managercheck:= manager
}
goto rticlipboard
Return

rticlipboard:
if (checkcheck=1) {
	goto checkcycle
}
rti := "Customer Name: " . namecheck . "`n"
     . "Login ID: " . idcheck . "`n"
     . "Alternate Contact #: " . numbercheck . " " . altnumbercheck . "`n"
     . "Alternate Email Address: " . emailcheck . " " . altemailcheck . "`n"
     . "Current Site/Location: " . locationcheck . " " . altlocationcheck . "`n"
     . "Computers # (computer name or host name): " . computercheck . "`n"
     . "Printer Name: " . printercheck . "`n"
     . "Server Name: " . servercheck . "`n"
     . "Specific Application: " . requestercheck . " " . screenshotsummarycheck . "`n"
     . "Brief Description of their issue: " . requestercheck . " " . screenshotsummarycheck . "`n"
     . "Error Message (if applicable): " . errormessagecheck . "`n"
     . "Screenshot (if applicable): " . screenshotcheck . "`n"
     . "Have you ever been able to access or use the program/software: " . access . "`n"
     . "When was the last successful time you were able to use the program/software: " whencheck . "`n"
     . "How Many Users Affected: " . affectedusercheck . " " . affecteduserscheck . "`n"
     . "What Does This Affect: " . affectedcheck . "`n"
     . "HR: " . hrcheck . "`n"
     . "Manager: " . managercheck . "`n"
	 . "--------------------"
clipboard:=rti
checkcheck:=!checkcheck
tiplettercount:= 20
tiptodisplay:= "RTI copied to clipboard"
goto tipcycle
Return

access:
GuiControlGet, access
if (access="Unknown"){
	GuiControl, 1: Disable, when
	GuiControl, 1:, when,Unknown
} else if (access="No"){
	GuiControl, 1: Disable, when
	GuiControl, 1:, when,Never
} else if (access="Yes"){
	GuiControl, 1: Enable, when
	GuiControl, 1:, when,
}
Return

number:
GuiControlGet, number
StringLen, numberlength, number
if (numberlength<1){
	GuiControl, 1: Disable, altnumber
	GuiControl, 1:, altnumber,
} else if (numberlength>0) {
	GuiControl, 1: Enable, altnumber
}
Return

email:
GuiControlGet, email
StringLen, emaillength, email
if (emaillength<1){
	GuiControl, 1: Disable, altemail
	GuiControl, 1:, altemail,
} else if (emaillength>0) {
	GuiControl, 1: Enable, altemail
}
Return

location:
GuiControlGet, location
StringLen, locationlength, location
if (locationlength<1){
	GuiControl, 1: Disable, altlocation
	GuiControl, 1:, altlocation,
} else if (locationlength>0) {
	GuiControl, 1: Enable, altlocation
}
Return

affecteduser:
GuiControlGet, affecteduser
StringLen, affecteduserlength, affecteduser
if (affecteduserlength<1){
	GuiControl, 1:, affectedusers, 
	GuiControl, 1: Disable, affectedusers
} else if (affecteduserlength>0) {
	GuiControl, 1: Enable, affectedusers
}
Return

requester:
GuiControlGet, requester
StringLen, requesterlength, requester
if (requesterlength<1){
	GuiControl, 1: Disable, relist
} else if (requesterlength>0) {
	GuiControl, 1: Enable, relist
}
Return

notepadbutton:
Gui,+LastFound
WinGetPos, gui1x, gui1y, gui1w, gui1h
notepadtoggle:=!notepadtoggle
if (notepadtoggle){
	goto gui1retract
} else {
	goto gui1expand
}
Return

logs:
GuiControlGet, notepad
clipboard:=date "`n`n" notepad
tiplettercount:= 20
tiptodisplay:= "Notepad copied to clipboard"
goto tipcycle
Return

password:
FileDelete, ranpass.txt
GuiControlGet, randomlength
GuiControlGet, simplepassword
if (!simplepassword) {
	randomnumberlimit:=10
	randomwordlimit:=26
} else {
	randomnumberlimit:=8
	randomwordlimit:=19
}
Loop %randomlength%{
	Random, randomtype, 1, 3
	if (randomtype=1){
		Random, randomnumbers10, 1, %randomnumberlimit%
		r1:=numberslist[randomnumbers10]
		FileAppend, %r1%, ranpass.txt
		Continue
	} else if (randomtype=2){
		Random, randomnumbers26a, 1, %randomwordlimit%
		r2:=lowercaseletterslist[randomnumbers26a]
		FileAppend, %r2%, ranpass.txt
		Continue
	} else if (randomtype=3){
		Random, randomnumbers26b, 1, %randomwordlimit%
		r3:=uppercaseletterslist[randomnumbers26b]
		FileAppend, %r3%, ranpass.txt
		Continue
	}
	Break
}
FileRead, randomlygeneratedpassword, ranpass.txt
GuiControl, 1:, displaypassword, %randomlygeneratedpassword%
Gui, 1:submit, NoHide
StringLen, randomlygeneratedpasswordlettercount, randomlygeneratedpassword
tiplettercount:= randomlygeneratedpasswordlettercount
FileDelete, ranpass.txt
clipboard:=randomlygeneratedpassword
tiptodisplay:=randomlygeneratedpassword
goto tipcycle
Return

tipcycle:
	if (tiplettercount<12){
		tiplettercount=12
	}
	Loop {
		tooltiptimer++
		tooltiptimerlimit:=tiplettercount*10
		if (tooltiptimer<tooltiptimerlimit){
			MouseGetPos, mousex, mousey
			ToolTip, %tiptodisplay%, (mousex+10), (mousey+10)
			Continue
		} else {
			tooltiptimer:=0
			ToolTip, 
			Break
		}
	}
Return

options:
optiontoggle:=!optiontoggle
if (optiontoggle=0) {
	Gui, 2:Destroy
	
	gui2buttonrow:=90
	gui2textrow:=5
	fetcherspeed:=2
	
	Gui, 2:Add, Checkbox, x5 y6 w55 vfetchertoggle gfetchertoggle, Fetcher
	Gui, 2:Add, Button, x60 y3 w35 gfetchersave vfetchersave, Save
	Gui, 2:Add, Button, x95 y3 w35 gfetcherload vfetcherload, Load
	
	Gui, 2:Add, Text, x%gui2textrow% y33, Speed: 
	Gui, 2:Add, Slider, x38 y30 w100 Range1-5 TickInterval vfetcherspeed, %fetcherspeed%
	
	Gui, 2:Add, Text, x%gui2textrow% y62, First Name: - - - - -  
	Gui, 2:Add, Button, x%gui2buttonrow% y60 w40 h20 vgui2firstname, Set
	
	Gui, 2:Add, Text, x%gui2textrow% y82, Last Name: - - - - -  
	Gui, 2:Add, Button, x%gui2buttonrow% y80 w40 h20 vgui2lastname, Set
	
	Gui, 2:Add, Text, x%gui2textrow% y102, ID: - - - - - - - - - - - - 
	Gui, 2:Add, Button, x%gui2buttonrow% y100 w40 h20 vgui2id, Set
	
	Gui, 2:Add, Text, x%gui2textrow% y122, Phone Number: - -
	Gui, 2:Add, Button, x%gui2buttonrow% y120 w40 h20 vgui2phonenumber, Set
	
	Gui, 2:Add, Text, x%gui2textrow% y142, Email Address: - - - 
	Gui, 2:Add, Button, x%gui2buttonrow% y140 w40 h20 vgui2email, Set
	
	Gui, 2:Add, Text, x%gui2textrow% y162, Location: - - - - - - - 
	Gui, 2:Add, Button, x%gui2buttonrow% y160 w40 h20 vgui2location, Set
	
	WinGetPos, martix, martiy,,, MARTI
	newx:=martix+903
	newy:=martiy+25
	gui2h:=136
	gui2w:=0
	gui2hgoal:=183
	gui2wgoal:=135
	Gui, 2:submit, NoHide
	Gui, 2:-Border
	goto gui2expand
} else if (optiontoggle=1){
	goto gui2retract
}
Return

;--------------------HOTKEYS--------------------
^!p::
	goto password
Return

^!l::
	goto logs
Return

^!r::
	goto rticlipboard
Return

^!1::
	GuiControlGet, name
	StringLen, namelettercount, name
	tiplettercount:= namelettercount
	if (namelettercount<1) {
		namecheck:= "Unknown"
		tiplettercount:= 7
	} else {
		namecheck:= name
		tiplettercount:= namelettercount
	}
clipboard:= namecheck
tiptodisplay:= namecheck
goto tipcycle
Return

^!2::
	GuiControlGet, id
	StringLen, idlettercount, id
	if (idlettercount<1) {
		idcheck:= "Unknown"
		tiplettercount:= 7
	} else {
		idcheck:= id
		tiplettercount:= idlettercount
	}
tiptodisplay:= idcheck
clipboard:= idcheck
goto tipcycle
Return

^!3::
GuiControlGet, number
StringLen, numberlettercount, number
if (numberlettercount<1){
	clipboard:= "Unknown"
	tiplettercount:= 7
	tiptodisplay:= "Unknown"
	goto tipcycle
} else {
	GuiControlGet, altnumber
	StringLen, altnumberlettercount, altnumber
	if (altnumberlettercount<1) {
		goto choosenumber1
	} else {
		MouseGetPos, mousex, mousey
		newx:=mousex-100
		newy:=mousey-12
		Gui, choosenumber:-Border
		Gui, choosenumber:Add, Button, x0 y0 w100 h25 gchoosenumber1, %number%
		Gui, choosenumber:Add, Button, x100 y0 w100 h25 gchoosenumber2, %altnumber%
		Winset, Alwaysontop,, Choose A Number
		WinSet, Style, -0xC00000, Choose A Number
		Gui, choosenumber:show, x%newx% y%newy% w200 h25, Choose A Number
	}
}

Return

choosenumber1:
	clipboard:= number
	Gui, choosenumber:Destroy
	tiplettercount:= numberlettercount
	tiptodisplay:= number
	goto tipcycle
Return

choosenumber2:
	clipboard:= altnumber
	Gui, choosenumber:Destroy
	tiplettercount:= altnumberlettercount
	tiptodisplay:= altnumber
	goto tipcycle
Return

^!4::
GuiControlGet, email
StringLen, emaillettercount, email
if (emaillettercount<1){
	clipboard:="Unknown"
	tiplettercount:= 7
	tiptodisplay:= "Unknown"
	goto tipcycle
} else {
	GuiControlGet, altemail
	StringLen, altemaillettercount, altemail
	if (altemaillettercount<1){
		goto chooseemail1
	} else {
		MouseGetPos, mousex, mousey
		newx:=mousex-100
		newy:=mousey-12
		Gui, chooseemail:-Border
		Gui, chooseemail:Add, Button, x0 y0 w100 h25 gchooseemail1, %email%
		Gui, chooseemail:Add, Button, x100 y0 w100 h25 gchooseemail2, %altemail%
		Winset, Alwaysontop,, Choose An Email
		WinSet, Style, -0xc00000, Choose An Email
		Gui, chooseemail:show, x%newx% y%newy% w200 h25, Choose An Email
	}
}
Return

chooseemail1:
	clipboard:= email
	Gui, chooseemail:Destroy
	tiplettercount:= emaillettercount
	tiptodisplay:= email
	goto tipcycle
Return

chooseemail2:
	clipboard:= altemail
	Gui, chooseemail:Destroy
	tiplettercount:= altemaillettercount
	tiptodisplay:= altemail
	goto tipcycle
Return

^!5::
GuiControlGet, location
StringLen, locationlettercount, location
if(locationlettercount<1){
	clipboard:="Unknown"
	tiplettercount:= 7
	tiptodisplay:= "Unknown"
	goto tipcycle
} else {
	GuiControlGet, altlocation
	StringLen, altlocationlettercount, altlocation
	if (altlocationlettercount<1){
		goto chooselocation1
	} else {
		MouseGetPos, mousex, mousey
		newx:=mousex-100
		newy:=mousey-12
		Gui, chooselocation:-Border
		Gui, chooselocation:Add, Button, x0 y0 w100 h25 gchooselocation1, %location%
		Gui, chooselocation:Add, Button, x100 y0 w100 h25 gchooselocation2, %altlocation%
		Winset, Alwaysontop,, Choose A Location
		Winset, Style, -0xc00000, Choose A Location
		Gui, chooselocation:show, x%newx% y%newy% w200 h25, Choose A Location
	}
}
Return

chooselocation1:
	Clipboard:= location
	Gui, chooselocation:Destroy
	tiplettercount:= locationlettercount
	tiptodisplay:= location
	goto tipcycle
Return

chooselocation2:
	Clipboard:= altlocation
	Gui, chooselocation:Destroy
	tiplettercount:= altlocationlettercount
	tiptodisplay:= altlocation
	goto tipcycle
Return

^!6::
	GuiControlGet, manager
	StringLen, managerlettercount, manager
	if (managerlettercount<1) {
		managercheck:= "Unknown"
		tiplettercount:= 7
	} else {
		managercheck:= manager
		tiplettercount:= managerlettercount
	}
clipboard:= managercheck
tiptodisplay:= managercheck
goto tipcycle
Return

^!7::
	GuiControlGet, hr
	StringLen, hrlettercount, hr
	if (hrlettercount<1) {
		hrcheck:= "Unknown"
		tiplettercount:= 7
	} else {
		hrcheck:= hr
		tiplettercount:= hrlettercount
	}
clipboard:= hrcheck
tiptodisplay:= hrcheck
goto tipcycle
Return

^!8::
	GuiControlGet, name
	GuiControlGet, requester
	StringLen, requesterlettercount, requester
	StringLen, namelettercount, name
	if (requesterlettercount<1) {
		requestercheck:= name
		tiplettercount:= namelettercount
		if (namelettercount<1){
			requestercheck:= "Unknown"
			tiplettercount:= 7
			tiptodisplay:= "Unknown"
		}
	} else {
		requestercheck:= requester
		tiplettercount:= requesterlettercount
	} 
clipboard:= requestercheck
tiptodisplay:= requestercheck
goto tipcycle
Return

^!9::
	GuiControlGet, computer
	StringLen, computerlettercount, computer
	if (computerlettercount<1) {
		computercheck:= "Blank SCCM"
		tiplettercount:= 10
	} else {
		computercheck:= computer
		tiplettercount:= computerlettercount
	}
clipboard:= computercheck
tiptodisplay:= computercheck
goto tipcycle
Return

^!0::
GuiControlGet, screenshot
if (screenshot=1) {
	screenshotcheck:= "Attached"
	screenshotsummarycheck:= "-Screenshot Attached"
} else {
	screenshotcheck:= "n/a"
	screenshotsummarycheck:= ""
}
GuiControlGet, relist
GuiControlGet, name
StringLen, namelettercount, name
if (namelettercount<1) {
	namecheck:= "Unknown"
} else {
	namecheck:= name
}
GuiControlGet, requesting
StringLen, requestinglettercount, requesting
if (requestinglettercount<1) {
	requestingcheck:= "Unknown Request/Incident"
	tiptodisplay:= "Unknown Request/Incident"
} else {
	requestingcheck:= requesting
}
GuiControlGet, requester
StringLen, requesterlettercount, requester
if (requesterlettercount<1){
	requestercheck:= namecheck " - " requestingcheck
} else {
	requestercheck:= requester " " relist " " namecheck " - " requestingcheck
}
Clipboard:= requestercheck . " " . screenshotsummarycheck
tiplettercount:= (namelettercount + requestinglettercount + requesterlettercount)
if (tiplettercount<20){
	tiplettercount:=20
}
if (tiplettercount>30) {
	tiplettercount:=30
}
tiptodisplay:= requestercheck . " " . screenshotsummarycheck
goto tipcycle
Return

;--------------------GUI Animations--------------------
gui2expand:
;gui2 contents is controlled by options label
SetTimer, 2follow1, on, 100
Loop {
	if (gui2w<gui2wgoal){
		gui2w+=16
		Gui, 2:show, x%newx% y%newy% w%gui2w% h%gui2h%, Fetcher
		Continue
	} else if (gui2w>gui2wgoal){
		gui2w:=gui2wgoal
	}
	if (gui2h<gui2hgoal){
		gui2h+=16
		Gui, 2:show, x%newx% y%newy% w%gui2w% h%gui2h%, Fetcher
		Continue
	} else if (gui2h>gui2hgoal){
		gui2h:=gui2hgoal
	}
	Gui, 2:show, x%newx% y%newy% w%gui2wgoal% h%gui2hgoal%, Fetcher
	Break
}
Return

fetchertoggle:
fetchertoggle:=!fetchertoggle
if (fetchertoggle=1) {
	GuiControl, 2: Enable, fetchersave
	GuiControl, 2: Enable, fetcherload
	GuiControl, 2: Enable, fetcherspeed
	GuiControl, 2: Enable, gui2firstname
	GuiControl, 2: Enable, gui2lastname
	GuiControl, 2: Enable, gui2id
	GuiControl, 2: Enable, gui2phonenumber
	GuiControl, 2: Enable, gui2email
	GuiControl, 2: Enable, gui2location
} else if (fetchertoggle=0){
	GuiControl, 2: Disable, fetchersave
	GuiControl, 2: Disable, fetcherload
	GuiControl, 2: Disable, fetcherspeed
	GuiControl, 2: Disable, gui2firstname
	GuiControl, 2: Disable, gui2lastname
	GuiControl, 2: Disable, gui2id
	GuiControl, 2: Disable, gui2phonenumber
	GuiControl, 2: Disable, gui2email
	GuiControl, 2: Disable, gui2location
}
Return

fetchersave:

Return

fetcherload:

Return

gui2retract:
SetTimer, 2follow1, Delete
Loop {
	if (gui2h>139){
		gui2h-=8
		Gui, 2:show, x%newx% y%newy% w%gui2w% h%gui2h%, MARTI2
		Continue
	} else if (gui2h<0){
		gui2h:=0
		Continue
	}
	if (gui2w>0){
		gui2w-=8
		Gui, 2:show, x%newx% y%newy% w%gui2w% h%gui2h%, MARTI2
		Continue
	}  else if (gui2w<0){
		gui2w:=0
		Continue
	}
	Gui, 2:Destroy
	Break
}
Gui, 2:Destroy
Return

gui1expand:
GuiControl, 1:, notepadbutton, Close Notepad
Gui, 1:submit, NoHide
Loop {
	if (gui1h<300){
		gui1h+=60
		Gui, 1:show, x%gui1x% y%gui1y% w900 h%gui1h%,
		Continue
	} else if (gui1h<400){
		gui1h+=30
		Gui, 1:show, x%gui1x% y%gui1y% w900 h%gui1h%,
		Continue
	} else if (gui1h<500){
		gui1h+=15
		Gui, 1:show, x%gui1x% y%gui1y% w900 h%gui1h%,
		Continue
	}
Gui, 1:show, x%gui1x% y%gui1y% w900 h500
Break
}
Return

gui1retract:
GuiControl, 1:, notepadbutton, Open Notepad
Gui, 1:submit, NoHide
Loop {
	if (gui1h>142){
	gui1h-=15
	Gui, 1:show, x%gui1x% y%gui1y% w900 h%gui1h%,
	Continue
	} else if (guih>250){
	gui1h-=30
	Gui, 1:show, x%gui1x% y%gui1y% w900 h%gui1h%,
	Continue
	} else if (guih>350) {
	gui1h-=60
	Gui, 1:show, x%gui1x% y%gui1y% w900 h%gui1h%,
	Continue
	}
Gui, 1:show, x%gui1x% y%gui1y% w900 h141
Break
}
Return

Guiclose:
	ExitApp
Return
