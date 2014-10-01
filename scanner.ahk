; Excel barcode scanner utility
; By Clive Galway - evilc@evilc.com

#singleinstance force
OnExit, ExitApp
BuildIniName()

PreambleKey := ReadIni("PreambleKey" , "Bindings" , "F8")
PostAmbleKey := ReadIni("PostambleKey" , "Bindings" , "Enter")
LastPreambleKey := ""	; The last key the preamble was bound to - used to remove old bindings
LinkSpreadsheetKey := ReadIni("LinkSpreadsheetKey" , "Bindings" , "F4")
LastLinkSpreadsheetKey := ""

InputActive := 0
ScanList := []

CreateGui()

IniName := ""
BuildIniName()

LinkedExcel := 0
LinkedWorkbook := 0
LinkedMainHwnd := 0
NotFoundSheet := 0
NotFoundCount := 0
SheetCount := 0

RegisterPreamble(PreambleKey)
RegisterLink(LinkSpreadsheetKey)

StartListener()

return


CreateGui(){
	global PreambleKey
	global PostambleKey
	global LinkSpreadsheetKey
	
	pre := ReadIni("PreambleKey" , "Bindings" , "F8")
	post := ReadIni("PostambleKey" , "Bindings" , "Enter")
	link := ReadIni("LinkSpreadsheetKey" , "Bindings" , "F4")
	
	Gui, Add, Text, x5 y10, Preamble Key: 
	Gui, Add, Edit, xp+150 yp-2 w100 vPreambleKey gOptionChanged, %pre%
	
	Gui, Add, Text, x5 yp+30, Postamble Key: 
	Gui, Add, Edit, xp+150 yp-2 w100 vPostambleKey gOptionChanged, %post%
	
	Gui, Add, Text, x5 yp+30, Link Spreadsheet Key: 
	Gui, Add, Edit, xp+150 yp-2 w100 vLinkSpreadsheetKey gOptionChanged, %link%
	
	Gui, Add, Link,x5 yp+50, Script Author: Clive Galway (<a href="mailto://evilc@evilc.com">evilc@evilc.com</a>)
	
	Gui, Add, StatusBar
	
	UpdateStatusBar()
	Gui, Show, W300 H160, Excel Barcode Scanner Utility
	
	Gui, 2:New, -Border +AlwaysOnTop
	Gui, 2:Font, S20
	Gui, 2:Add, Text, w260 R1 Center cWhite, Scanning...
	Gui, 2:Show, w300 h65
	Gui, 2:Color, Blue
	Gui, 2:Hide
}

RegisterPreamble(key){
	global LastPreambleKey
	
	if(LastPreambleKey){
		;unbind previous binding
		hotkey, %LastPreambleKey%, OFF
	}
	hotkey, %key% , PreamblePressed
	LastPreambleKey := key
}

RegisterLink(key){
	global LastLinkSpreadsheetKey
	
	if(LastLinkSpreadsheetKey){
		;unbind previous binding
		hotkey, %LastLinkSpreadsheetKey%, OFF
	}
	hotkey, %key% , LinkSpreadhseetPressed
	LastLinkSpreadsheetKey := key
}

PreamblePressed:
	PreamblePressed()
	return

PreamblePressed(){
	global ScanList
	global LinkedWorkbook
	global PostAmbleKey
	
	;Block keyboard and read keys pressed until postamble detected.
	Input, text, , {%PostAmbleKey%}
	Gui, 2:Show
	Gui, 2:Hide
	
	if (LinkedWorkbook){
		ScanList.Insert(text)
	} else {
		;LowBeep()
		TTS("Warning: Excel not linked")
	}
}

LinkSpreadhseetPressed:
	LinkSpreadhseetPressed()
	return
	
LinkSpreadhseetPressed(){
	LinkWorkbook()
}
	
StartListener(){
	global ScanList
	global LinkedExcel
	global LinkedWorkbook
	global SheetCount
	global NotFoundSheet
	global NotFoundCount
	
	Loop {
		if (LinkedExcel){
			if (ScanList.MaxIndex()){
				found := []
				res := 0
				first := 0

				Loop % SheetCount {
					s := A_Index + 0 ; ToDo: Why is +0 needed? Returning string otherwise?
					Loop {
						; Find all instances on this sheet
						if (res == 0){
							res := LinkedWorkbook.Sheets(s).Cells.Find(ScanList[1])
							first := res.Address
						} else {
							res := LinkedWorkbook.Sheets(s).Cells.FindNext(res)
							if (res.Address = first){
								; If we wrapped around, stop
								res := 0
							}
						}

						if (res.Count){
							found.Insert({sheet: s, result: res})
						} else {
							res := 0
							break
						}
					}
				}

				if (found.MaxIndex() == 0 || found.MaxIndex() == ""){
					; Not found - add item to !NotFound! sheet
					res := 0
					res := LinkedWorkbook.Sheets(NotFoundSheet).Cells.Find(ScanList[1])
					if (!res.Count){
						; Check NotFound sheet to see if item is unique
						NotFoundCount++
						LinkedWorkbook.Sheets(NotFoundSheet).Range("A" NotFoundCount).Value := ScanList[1]
					}
					LowBeep()
				} else if (found.MaxIndex() == 1){
					; Found item
					LinkedWorkbook.Sheets(found[1].sheet).Activate
					found[1].result.Select
					found[1].result.Interior.Color := RGB2Excel(0,255,0)
					;HighBeep()
					
					LinkedWorkbook.Save
					
					; Say Name / Number of sheet
					name := GetWorksheetName()
					name := StrSplit(name,"|")
					name := name[2]
					name = %name% ; Trim whitespace
					if (name == ""){
						name := found[1].sheet
					}
					TTS(name "!")
					
				} else {
					WarningBeep()

				}
				ScanList.Remove(1)
			}
		}
		Sleep 50
	}
}

RGB2Excel(R, G, B) {
	ExcelFormat :=  (R<<16) + (G<<8) + B
	return ExcelFormat
}

LinkWorkbook(){
	global LinkedWorkbook
	global LinkedMainHwnd
	global LinkedExcel
	global NotFoundSheet
	global NotFoundCount
	global SheetCount
	
	search_open := 0

	if (IsSearchboxActive()){
		CloseDialog()
		WinWaitActive, ahk_class XLMAIN
		search_dialog_was_open := 1
	}
	if (IsSpreadsheetActive()){
		LinkedExcel := Excel_Get()
		LinkedWorkbook := GetWorkbook()
		LinkedMainHwnd := GetActiveHwnd()
		
		; Check to see if "Not Found" sheet is present.
		notfound := 0
		index := 0
		Loop % LinkedWorkbook.Sheets.Count {
			if(LinkedWorkbook.Sheets(A_Index).name == "!NotFound!"){
				notfound := 1
				index := A_Index + 0
			}
		}
		if (notfound){
			if (index != LinkedWorkbook.Sheets.Count){
				msgbox "ERROR: !NotFound! worksheet is not the last tab - please move it to the end!"
				UnlinkExcel()
				return
			}
			Loop {
				; Find first empty cell in A column
				;LinkedWorkbook.Sheets(index).Cells.Find(ScanList[1])
				val := LinkedWorkbook.Sheets(index).Range("A" A_Index).Value
				if (val == ""){
					NotFoundSheet := index
					NotFoundCount := A_Index - 1
					break
				}
			}
		} else {
			; Create NotFound sheet
			sht := LinkedWorkbook.Sheets.Add(,LinkedWorkbook.Sheets(LinkedWorkbook.Sheets.Count))
			sht.Name := "!NotFound!"
			NotFoundSheet := LinkedWorkbook.Sheets.Count
			NotFoundCount := 0
		}
		SheetCount := LinkedWorkbook.Sheets.Count - 1
	} else {
		; Spreadsheet not active app
		UnlinkExcel()
	}
	
	UpdateStatusBar(1)

}

UnlinkExcel(){
	global LinkedExcel
	global LinkedWorkbook
	global LinkedMainHwnd
	global SheetCount
	
	SheetCount := LinkedExcel := LinkedWorkbook := LinkedMainHwnd := 0
}

UpdateStatusBar(mode := 0){
	global LinkedWorkbook
	global LinkSpreadsheetKey
	
	if (LinkedWorkbook){
		SB_SetText("Linked: " GetFileName())
		HighBeep()
	} else {
		SB_SetText("Linked: NONE. Go to Excel and hit " LinkSpreadsheetKey " to assign.")
		if (mode){
			LowBeep()
		}
	}
}

IsSpreadsheetActive(){
	win := GetAhkClass()
	if (win == "XLMAIN"){
		return 1
	} else {
		return 0
	}
}

IsSearchboxActive(){
	win := GetAhkClass()
	if (win == "bosa_sdm_XL9"){
		return 1
	} else {
		return 0
	}
}

GetAhkClass(){
	wingetclass, cls, A
	return cls
}

OpenSearchBox(){
	Send ^{f}
}

CloseDialog(){
	Send {Esc}
}

GetActiveHwnd(){
	Winget, hwnd, ID, A
	return hwnd
}

LowBeep(){
	soundbeep, 500, 200
}

HighBeep(){
	soundbeep, 700, 200
}

RisingBeep(){
	sleep 250
	soundbeep, 500, 100
	soundbeep, 600, 100
	soundbeep, 800, 100
}

WarningBeep(){
	sleep 250
	soundbeep, 800, 100
	soundbeep, 500, 100
	soundbeep, 800, 100
	soundbeep, 500, 100
}

; Returns workbook object for active workbook
GetWorkbook(){
	try {
		oWorkbook := Excel_Get().ActiveWorkbook ; try to access active Workbook object
	} catch {
		msgbox Excel not found
		;return ; case when Excel doesn't exist, or it exists but there is no active workbook. Just Return or Exit or ExitApp.
		return 0
	}
	return oWorkbook
}
	
; Returns an object containing current row and column info of selected cell
GetCurrentCell(){
	str := Excel_Get().ActiveCell.Address()
	str := StrSplit(str,"$")
	str := {col: str[2], row: str[3]}
	return str
}

; Returns index of current worksheet (1st / 2nd worksheet etc)
GetWorksheetID(){
	;oWorkbook := Excel_Get().ActiveWorkbook
	oWorkbook := GetWorkbook()
	return oWorkbook.ActiveSheet.Index()
}

; Returns name of current worksheet
GetWorksheetName(){
	oWorkbook := GetWorkbook()
	;oWorkbook := Excel_Get().ActiveWorkbook
	return oWorkbook.ActiveSheet.Name()
}

; Returns filename currently open
GetFileName(){
	oWorkbook := GetWorkbook()
	;oWorkbook := Excel_Get().ActiveWorkbook
	return oWorkbook.Name()
}

ReadIni(key,section,default){
	global IniName

	ini := IniName
	IniRead, out, %ini%, %section%, %key%, %default%
	return out
}

; Updates the settings file. If value is default, it deletes the setting to keep the file as tidy as possible
UpdateIni(key, section, value, default){
	global IniName

	tmp := IniName
	if (value != default){
		; Only write the value if it differs from what is already written
		if (ReadIni(key,section,-1) != value){
			IniWrite,  %value%, %tmp%, %section%, %key%
		}
	} else {
		; Only delete the value if there is already a value to delete
		if (ReadIni(key,section,-1) != -1){
			IniDelete, %tmp%, %section%, %key%
		}
	}
}

BuildIniName(){
	global IniName

	tmp := A_Scriptname
	Stringsplit, tmp, tmp,.
	IniName := ""
	last := ""
	Loop, % tmp0 {
		if (last != ""){
			if (IniName != ""){
				IniName := IniName "."
			}
			IniName := IniName last
		}
		last := tmp%A_Index%
	}
	IniName .= ".ini"
	return
}

TTS(str){
	ComObjCreate("SAPI.SpVoice").Speak(str)
}

OptionChanged:
	OptionChanged()
	return

OptionChanged(){
	global PreambleKey
	global PostambleKey
	global LinkSpreadsheetKey
	
	gui, submit, nohide
	
	UpdateIni("PreambleKey", "Bindings", PreambleKey, "F8")
	RegisterPreamble(PreambleKey)
	
	UpdateIni("PostambleKey", "Bindings", PostambleKey, "Enter")
	
	UpdateIni("LinkSpreadsheetKey", "Bindings", LinkSpreadsheetKey, "F4")
	RegisterLink(LinkSpreadsheetKey)
	
	; Text may change in status bar - eg Link Key prompt
	UpdateStatusBar()
}

; GUI was closed
GuiClose:
ExitApp:
	ExitApp
	return

; =============================================================================================================
; 3rd party functions

Excel_Get(WinTitle="ahk_class XLMAIN") {
	; by Sean and Jethrow, minor modification by Learning one
	; http://www.autohotkey.com/forum/viewtopic.php?p=492448#492448
	ControlGet, hwnd, hwnd, , Excel71, %WinTitle%
	if (!hwnd){
		return
	}
	Window := Acc_ObjectFromWindow(hwnd, -16)
	Loop {
		try {
			Application := Window.Application
		} catch {
			ControlSend, Excel71, {esc}, %WinTitle%
		}
	} Until !!Application
	return Application
}

;------------------------------------------------------------------------------
; Acc.ahk Standard Library
; by Sean
; Updated by jethrow:
; 	Modified ComObjEnwrap params from (9,pacc) --> (9,pacc,1)
; 	Changed ComObjUnwrap to ComObjValue in order to avoid AddRef (thanks fincs)
; 	Added Acc_GetRoleText & Acc_GetStateText
; 	Added additional functions - commented below
; 	Removed original Acc_Children function
;	Added Acc_Error, Acc_ChildrenByRole, & Acc_Get functions
; last updated 10/25/2012
;------------------------------------------------------------------------------

Acc_Init()
{
	Static	h
	If Not	h
		h:=DllCall("LoadLibrary","Str","oleacc","Ptr")
}
Acc_ObjectFromEvent(ByRef _idChild_, hWnd, idObject, idChild)
{
	Acc_Init()
	If	DllCall("oleacc\AccessibleObjectFromEvent", "Ptr", hWnd, "UInt", idObject, "UInt", idChild, "Ptr*", pacc, "Ptr", VarSetCapacity(varChild,8+2*A_PtrSize,0)*0+&varChild)=0
	Return	ComObjEnwrap(9,pacc,1), _idChild_:=NumGet(varChild,8,"UInt")
}

Acc_ObjectFromPoint(ByRef _idChild_ = "", x = "", y = "")
{
	Acc_Init()
	If	DllCall("oleacc\AccessibleObjectFromPoint", "Int64", x==""||y==""?0*DllCall("GetCursorPos","Int64*",pt)+pt:x&0xFFFFFFFF|y<<32, "Ptr*", pacc, "Ptr", VarSetCapacity(varChild,8+2*A_PtrSize,0)*0+&varChild)=0
	Return	ComObjEnwrap(9,pacc,1), _idChild_:=NumGet(varChild,8,"UInt")
}

Acc_ObjectFromWindow(hWnd, idObject = 0)
{
	Acc_Init()
	If	DllCall("oleacc\AccessibleObjectFromWindow", "Ptr", hWnd, "UInt", idObject&=0xFFFFFFFF, "Ptr", -VarSetCapacity(IID,16)+NumPut(idObject==0xFFFFFFF0?0x46000000000000C0:0x719B3800AA000C81,NumPut(idObject==0xFFFFFFF0?0x0000000000020400:0x11CF3C3D618736E0,IID,"Int64"),"Int64"), "Ptr*", pacc)=0
	Return	ComObjEnwrap(9,pacc,1)
}

Acc_WindowFromObject(pacc)
{
	If	DllCall("oleacc\WindowFromAccessibleObject", "Ptr", IsObject(pacc)?ComObjValue(pacc):pacc, "Ptr*", hWnd)=0
	Return	hWnd
}

Acc_GetRoleText(nRole)
{
	nSize := DllCall("oleacc\GetRoleText", "Uint", nRole, "Ptr", 0, "Uint", 0)
	VarSetCapacity(sRole, (A_IsUnicode?2:1)*nSize)
	DllCall("oleacc\GetRoleText", "Uint", nRole, "str", sRole, "Uint", nSize+1)
	Return	sRole
}

Acc_GetStateText(nState)
{
	nSize := DllCall("oleacc\GetStateText", "Uint", nState, "Ptr", 0, "Uint", 0)
	VarSetCapacity(sState, (A_IsUnicode?2:1)*nSize)
	DllCall("oleacc\GetStateText", "Uint", nState, "str", sState, "Uint", nSize+1)
	Return	sState
}

Acc_SetWinEventHook(eventMin, eventMax, pCallback)
{
	Return	DllCall("SetWinEventHook", "Uint", eventMin, "Uint", eventMax, "Uint", 0, "Ptr", pCallback, "Uint", 0, "Uint", 0, "Uint", 0)
}

Acc_UnhookWinEvent(hHook)
{
	Return	DllCall("UnhookWinEvent", "Ptr", hHook)
}
/*	Win Events:

	pCallback := RegisterCallback("WinEventProc")
	WinEventProc(hHook, event, hWnd, idObject, idChild, eventThread, eventTime)
	{
		Critical
		Acc := Acc_ObjectFromEvent(_idChild_, hWnd, idObject, idChild)
		; Code Here:

	}
*/

; Written by jethrow
Acc_Role(Acc, ChildId=0) {
	try return ComObjType(Acc,"Name")="IAccessible"?Acc_GetRoleText(Acc.accRole(ChildId)):"invalid object"
}
Acc_State(Acc, ChildId=0) {
	try return ComObjType(Acc,"Name")="IAccessible"?Acc_GetStateText(Acc.accState(ChildId)):"invalid object"
}
Acc_Location(Acc, ChildId=0, byref Position="") { ; adapted from Sean's code
	try Acc.accLocation(ComObj(0x4003,&x:=0), ComObj(0x4003,&y:=0), ComObj(0x4003,&w:=0), ComObj(0x4003,&h:=0), ChildId)
	catch
		return
	Position := "x" NumGet(x,0,"int") " y" NumGet(y,0,"int") " w" NumGet(w,0,"int") " h" NumGet(h,0,"int")
	return	{x:NumGet(x,0,"int"), y:NumGet(y,0,"int"), w:NumGet(w,0,"int"), h:NumGet(h,0,"int")}
}
Acc_Parent(Acc) { 
	try parent:=Acc.accParent
	return parent?Acc_Query(parent):
}
Acc_Child(Acc, ChildId=0) {
	try child:=Acc.accChild(ChildId)
	return child?Acc_Query(child):
}
Acc_Query(Acc) { ; thanks Lexikos - www.autohotkey.com/forum/viewtopic.php?t=81731&p=509530#509530
	try return ComObj(9, ComObjQuery(Acc,"{618736e0-3c3d-11cf-810c-00aa00389b71}"), 1)
}
Acc_Error(p="") {
	static setting:=0
	return p=""?setting:setting:=p
}
Acc_Children(Acc) {
	if ComObjType(Acc,"Name") != "IAccessible"
		ErrorLevel := "Invalid IAccessible Object"
	else {
		Acc_Init(), cChildren:=Acc.accChildCount, Children:=[]
		if DllCall("oleacc\AccessibleChildren", "Ptr",ComObjValue(Acc), "Int",0, "Int",cChildren, "Ptr",VarSetCapacity(varChildren,cChildren*(8+2*A_PtrSize),0)*0+&varChildren, "Int*",cChildren)=0 {
			Loop %cChildren%
				i:=(A_Index-1)*(A_PtrSize*2+8)+8, child:=NumGet(varChildren,i), Children.Insert(NumGet(varChildren,i-8)=9?Acc_Query(child):child), NumGet(varChildren,i-8)=9?ObjRelease(child):
			return Children.MaxIndex()?Children:
		} else
			ErrorLevel := "AccessibleChildren DllCall Failed"
	}
	if Acc_Error()
		throw Exception(ErrorLevel,-1)
}
Acc_ChildrenByRole(Acc, Role) {
	if ComObjType(Acc,"Name")!="IAccessible"
		ErrorLevel := "Invalid IAccessible Object"
	else {
		Acc_Init(), cChildren:=Acc.accChildCount, Children:=[]
		if DllCall("oleacc\AccessibleChildren", "Ptr",ComObjValue(Acc), "Int",0, "Int",cChildren, "Ptr",VarSetCapacity(varChildren,cChildren*(8+2*A_PtrSize),0)*0+&varChildren, "Int*",cChildren)=0 {
			Loop %cChildren% {
				i:=(A_Index-1)*(A_PtrSize*2+8)+8, child:=NumGet(varChildren,i)
				if NumGet(varChildren,i-8)=9
					AccChild:=Acc_Query(child), ObjRelease(child), Acc_Role(AccChild)=Role?Children.Insert(AccChild):
				else
					Acc_Role(Acc, child)=Role?Children.Insert(child):
			}
			return Children.MaxIndex()?Children:, ErrorLevel:=0
		} else
			ErrorLevel := "AccessibleChildren DllCall Failed"
	}
	if Acc_Error()
		throw Exception(ErrorLevel,-1)
}
Acc_Get(Cmd, ChildPath="", ChildID=0, WinTitle="", WinText="", ExcludeTitle="", ExcludeText="") {
	static properties := {Action:"DefaultAction", DoAction:"DoDefaultAction", Keyboard:"KeyboardShortcut"}
	AccObj :=   IsObject(WinTitle)? WinTitle
			:   Acc_ObjectFromWindow( WinExist(WinTitle, WinText, ExcludeTitle, ExcludeText), 0 )
	if ComObjType(AccObj, "Name") != "IAccessible"
		ErrorLevel := "Could not access an IAccessible Object"
	else {
		StringReplace, ChildPath, ChildPath, _, %A_Space%, All
		AccError:=Acc_Error(), Acc_Error(true)
		Loop Parse, ChildPath, ., %A_Space%
			try {
				if A_LoopField is digit
					Children:=Acc_Children(AccObj), m2:=A_LoopField ; mimic "m2" output in else-statement
				else
					RegExMatch(A_LoopField, "(\D*)(\d*)", m), Children:=Acc_ChildrenByRole(AccObj, m1), m2:=(m2?m2:1)
				if Not Children.HasKey(m2)
					throw
				AccObj := Children[m2]
			} catch {
				ErrorLevel:="Cannot access ChildPath Item #" A_Index " -> " A_LoopField, Acc_Error(AccError)
				if Acc_Error()
					throw Exception("Cannot access ChildPath Item", -1, "Item #" A_Index " -> " A_LoopField)
				return
			}
		Acc_Error(AccError)
		StringReplace, Cmd, Cmd, %A_Space%, , All
		properties.HasKey(Cmd)? Cmd:=properties[Cmd]:
		try {
			if (Cmd = "Location")
				AccObj.accLocation(ComObj(0x4003,&x:=0), ComObj(0x4003,&y:=0), ComObj(0x4003,&w:=0), ComObj(0x4003,&h:=0), ChildId)
			  , ret_val := "x" NumGet(x,0,"int") " y" NumGet(y,0,"int") " w" NumGet(w,0,"int") " h" NumGet(h,0,"int")
			else if (Cmd = "Object")
				ret_val := AccObj
			else if Cmd in Role,State
				ret_val := Acc_%Cmd%(AccObj, ChildID+0)
			else if Cmd in ChildCount,Selection,Focus
				ret_val := AccObj["acc" Cmd]
			else
				ret_val := AccObj["acc" Cmd](ChildID+0)
		} catch {
			ErrorLevel := """" Cmd """ Cmd Not Implemented"
			if Acc_Error()
				throw Exception("Cmd Not Implemented", -1, Cmd)
			return
		}
		return ret_val, ErrorLevel:=0
	}
	if Acc_Error()
		throw Exception(ErrorLevel,-1)
}