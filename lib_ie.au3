#include-once

#Region Global Include
#include <IE.au3>
#include <IEEx.au3>
#EndRegion Global Include


#cs ----------------------------------------------------------------------------
Function to work with IE
	_ie_getActibTab() - Retrieve the Window Object of the currently active IE (on top)
	_ie_getActibWindow() - Retrieve active IE window
	_ie_getPageInNewWindow() - open page in a new hidden window and return its xml
	_ie_quitAll( $bKillAll = true ) - kill all IE instances (false - kill only invisisble)
	_ie_new( $url ) - open new IE and goto $url
	_ie_getURL( $oTab ) - get current $url in activ IE
	_ie_goto( $oTab, $url ) - naavigate to $url in active IE
	_ie_saveSearchFields($oTab) - store some search fields
	_ie_restoreSearchFields($oTab) - restore some search fields
	_ie_submitSearchForm($oTab) - submit search form
	_ie_submitLoginForm() - submit login form
	_ie_getMsgArray($oTab)- get messages from RF log to array
	_ie_getLinks( $oTab ) - get all links on web page

#ce ----------------------------------------------------------------------------

; #VARIABLES# ===================================================================================================================
#Region Global Variables

; saved user/pass
Global $sUser = ""
Global $sPass = ""

; saved fields
Global $sAktor
Global $sMsgType

; active IE tab
Global $oTab

#EndRegion Global variables


Func _ie_isActive()
	return IsObj($oTab)
EndFunc

;===============================================================================
;
; Function Name:    _ie_getActiveWindow()
; Description:      Retrieve the Window Object of the currently active IE (on top)
; Parameter(s):     None
; Requirement(s):   AutoIt3 V3.2 or higher
;                   On Success  - Returns an object variable pointing to the IE Window Object
;                   On Failure  - Returns 0 (no active IE windows)
;===============================================================================
Func _ie_getActiveWindow()

	Local $aWinList
	; get all IE windows start from top. Top window is first.
	$aWinList = WinList("[REGEXPTITLE:(?i)(.*Internet Explorer.*)]")

	if $aWinList[0][0] = 0 then return 0

	return $aWinList[1][1]	; upper active windows

EndFunc

;===============================================================================
;
; Function Name:    _ie_getActiveTab()
; Description:      Retrieve the IE Window Object of the currently active tab
; Parameter(s):     None
; Requirement(s):   AutoIt3 V3.2 or higher
;                   On Success  - sets $oTab global variable pointing to the IE Window Object
;                   On Failure  - Returns 0 and sets @ERROR
;                   @ERROR      - 0 ($_IEStatus_Success) = No Error
;                               - 7 ($_IEStatus_NoMatch) = No Match
; Author(s):        Dan Pollak
;===============================================================================
;
Func _ie_getActiveTab()
Local $aWinList
	; get all IE windows start from top. Top window is first.
	$aWinList = WinList("[REGEXPTITLE:(?i)(.*Internet Explorer.*)]")

	if $aWinList[0][0] = 0 then
		return 0
	EndIf

#cs -------
	to avoid problem that there can be empty window on top which does not have IE instance
	we check if window handle exists in the list of IE instances

	windows list :
	1. 0x1404   <- skip window not linked to IE on top
	2. 0x4566   <- real top IE window as it matched IE instance window
	3. 0x3444

	IE instaces window handles
	1. 0x3444   <- IE tab in window #1
	2. 0x4566   <- IE tab in window #2
	2. 0x4566   <- IE tab in window #2

	To select the tab from top window:
	oTab = _IEAttach( "0x4566", "embedded" )

#ce ---------

	for $i=1 to UBound($aWinList, 1)-1

		Local $n=1
		while $n
			$oTab = _IEAttach( "", "instance", $n )
			If @error > 0 then ;= $_IESTATUS_NoMatch Then
				ExitLoop
			EndIf

			; if IE instance has same window handle (hwnd) then it is the right tab
			if _IEPropertyGet( $oTab, "hwnd" ) = $aWinList[$i][1] then

				;$gIEhwnd = $aWinList[$i][1]
				$oTab = _IEAttach($aWinList[$i][1],"embedded")	; get active tab in window
				return 1
			EndIf

			$n += 1

		WEnd

	Next

	Return 0

EndFunc
;===============================================================================
;
; Function Name:    _ie_getPageInNewWindow( $sLink )
; Description:      Retrieve body text from web page
; Parameter(s):     $sLink - link to web page
; Returns:
;			Text string with page contents
;           On Failure  - Returns 0 and sets @error
;				1 - TabCreate error sets @extended to Tabcreate @ERROR
;				2 - webadress not opened, likely passowrd required, @extended = actual address
;				3 - _IEBodyReadText error and sets @extended to _IERead @error
;				4 - Quit error and @extended = Quit @error
; =========================================================

Func	_ie_getPageInNewWindow( $sLink )

	DbgFile( "-->_IEGetPageInNewWindow " & $sLink )

	DbgFile( "   _IECreate" )

	Local $err = 0
	Local $oXml = _IECreate( $sLink ,0,0 )

	if not IsObj($oXml) then
	;if @error then
		return SetError(1, 0, "*** Error _IECreate " & @error & " " & @extended )
	EndIf

	DbgFile( "   _IEAttach" )
	$oXml = _IEAttach($sLink, "url")
	if @error then
		Return SetError(5, 0, "*** Error _IEAttach " & @error & " " & @extended )
	endif

	DbgFile( "   _IEBodyReadText" )
	Local $html = _IEBodyReadText( $oXml )
	if StringLen( $html ) < 1000 then
		DbgFile( "   _IEBodyReadText ERR:" & StringLen( $html) & " " & StringLeft( StringStripWS( $html, 8), 15  ))
		Sleep(1000)
		$html = _IEBodyReadText( $oXml )
	EndIf
	if @error then
		Return SetError(3, 0, "*** Error _IEBodyReadText " & @error & " " & @extended )
	endif

	DbgFile( "   _IEBodyReadText len:" & StringLen( $html) & " " & StringLeft( StringStripWS( $html, 8), 15  ))

	DbgFile( "   _IEQuit" )
	_IEQuit($oXml)
	if @error then
		Return SetError(4, 0, "*** Error _IEQuit" & @error & " " & @extended )
	endif

	DbgFile( "<--_IEGetPageInNewWindow " )

	return $html

EndFunc

;===============================================================================
;
; Function Name:    _ie_quitAll( $bKillAll = true )
; Description:      Quit from all IE instances av $type
; Parameter(s):     $type
;						true - kill visible also
;						false - kill only invisible
; Returns:
;			Number of instances killed
;           On Failure  - Returns 0 and sets @error
;				1 - no IE instances
;				2 - Quit error and @extended = Quit @error
; =========================================================

Func	_ie_quitAll( $bKillAll = true )

	Local $oTabTmp
	Local $nTab = 0

	Local $i = 1
DbgFile( "kill all")
	While 1
	   $oTabTmp = _IEAttach( "", "instance", $i )
	   If @error > 0 then ;= $_IESTATUS_NoMatch Then
		  ExitLoop
	   EndIf

		if $bKillAll or not _IEPropertyGet( $oTabTmp, "visible" ) then
			_IEQuit( $oTabTmp )
			$nTab += 1
		Else
			$i += 1
		endif

	WEnd

	return $nTab

EndFunc

Func _ie_new( $url )
	$oTab = _IECreate( $url )
	return
EndFunc

Func _ie_getURL()
	if _ie_isActive() then
	return _IEPropertyGet( $oTab, "locationurl")
	Else
	DbgFile( "getURL - no obj" )
	EndIf

EndFunc

Func _ie_goto($url )
	if _ie_isActive() then
	_IENavigate ( $oTab, $url )
	return @error
	Else
	DbgFile( "goto - no obj" )
	EndIf
EndFunc

Func _ie_saveSearchFields()
	Local $oForm = _IEFormGetObjByName($oTab, "logfilter")
	Local $oMsgType = _IEFormElementGetObjByName($oForm,  "msgType" )
	Local $oAktor = _IEFormElementGetObjByName($oForm,  "aktor" )

	$sAktor = _IEFormElementGetValue($oAktor)
	$sMsgType = _IEFormElementGetValue($oMsgType)

EndFunc

Func _ie_restoreSearchFields()
	Local $oForm = _IEFormGetObjByName($oTab, "logfilter")
	Local $oMsgType = _IEFormElementGetObjByName($oForm,  "msgType" )
	Local $oAktor = _IEFormElementGetObjByName($oForm,  "aktor" )

	_IEFormElementSetValue($oAktor, $sAktor)
	_IEFormElementSetValue($oMsgType, $sMsgType)

EndFunc

Func _ie_submitSearchForm()
	Local $oForm = _IEFormGetObjByName($oTab, "logfilter")
	Local $oDatoFra = _IEFormElementGetObjByName($oForm,  "datoFra" )
	;Local $oDatoTil = _IEFormElementGetObjByName($oForm,  "datoTil" )
	;Local $oMsgType = _IEFormElementGetObjByName($oForm,  "msgType" )
	;Local $oAktor = _IEFormElementGetObjByName($oForm,  "aktor" )
	;Local $oMsgId = _IEFormElementGetObjByName($oForm,  "msgId" )

	; Returns the current Date and Time in format YYYY/MM/DD HH:MM:SS.
	Local $sFrom = _RFtimeDiff( -GetCurInterval(True), 'n')

	_IEFormElementSetValue($oDatoFra, $sFrom)
	_IEFormElementCheckBoxSelect($oForm, "WS-R" )
	_IEFormElementCheckBoxSelect($oForm, "WS-U" )

	Local $oSubmit = _IEGetObjById($oTab, "sokeknapp")
	_IEAction($oSubmit, "click")
	_IELoadWait($oTab)

EndFunc

;===============================================================================
; Function Name:    _ie_submitLoginForm()
; Return:
;	1 - if form submitted
;	0 - if user/pass must be submitted manuelly
;===============================================================================

Func _ie_submitLoginForm()

	Local $oLoginForm = _IEFormGetObjByName($oTab, "login")
	Local $oUserId = _IEFormElementGetObjByName($oLoginForm,  "userId" )
	Local $oPass = _IEFormElementGetObjByName($oLoginForm,  "pass" )

	_IEFormElementSetValue($oUserId, $sUser)
	_IEFormElementSetValue($oPass, $sPass)
	if $sUser <> "" and $sPass <> "" then
		_IEFormSubmit($oLoginForm)
		return 1
	endif

	Return 0
EndFunc

;===============================================================================
; Function Name:    _ie_getMsgArray()
;
; Convert message table on page to array
; Return:
;	2D array with messages
;	0 - no messages found
;	1 - failed to get message table, sets @error:
;		1 - no messages in table
;		2 -
;===============================================================================

Func	_ie_getMsgArray()

	; get all links on the page
	Local $oTable = _IETableGetCollection($oTab,0)
	if @error <> 0 then
		DbgFile("*** No messages. " & @error )
		return 0
	EndIf
	;GUICtrlSetData($idLabel,"3. Got message table... " )

	Local $aTableData = _IETableWriteToArray($oTable, true)
	if @error <> 0 then
		DbgFile("*** Error: _IETableWriteToArray: " & @error)
		return 0
	EndIf
	;GUICtrlSetData($idLabel, "4. Got message array..." )

	return $aTableData

EndFunc

Func _ie_getLinks()
	Local $oLinks = _IELinkGetCollection($oTab)
	if @error <> 0 then
		DbgFile( "*** Error: _IELinkGetCollection: " & @error)
		return 0
	EndIf
	return $oLinks
EndFunc
