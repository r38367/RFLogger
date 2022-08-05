#include-once

#Region Global Include
#include <IE.au3>
#include <IEEx.au3>
#EndRegion Global Include


#cs ----------------------------------------------------------------------------
Function to work with IE
	_IEGetActibTab() - Retrieve the Window Object of the currently active IE (on top)
	_IEGetActibWindow() -
	_IEGetPageInNewWindow() - open page in a new hidden window and return its xml


26/6/21
21 - remove debug output to edit control
#ce ----------------------------------------------------------------------------

;===============================================================================
;
; Function Name:    _IEGetActiveWindow()
; Description:      Retrieve the Window Object of the currently active IE (on top)
; Parameter(s):     None
; Requirement(s):   AutoIt3 V3.2 or higher
;                   On Success  - Returns an object variable pointing to the IE Window Object
;                   On Failure  - Returns 0 (no active IE windows)
;===============================================================================
Func _IEGetActiveWindow()

	Local $aWinList
	; get all IE windows start from top. Top window is first.
	$aWinList = WinList("[REGEXPTITLE:(?i)(.*Internet Explorer.*)]")

	if $aWinList[0][0] = 0 then return 0

	return $aWinList[1][1]	; upper active windows

EndFunc

;===============================================================================
;
; Function Name:    _IEGetActiveTab()
; Description:      Retrieve the IE Window Object of the currently active tab
; Parameter(s):     None
; Requirement(s):   AutoIt3 V3.2 or higher
;                   On Success  - Returns an object variable pointing to the IE Window Object
;                   On Failure  - Returns 0 and sets @ERROR
;                   @ERROR      - 0 ($_IEStatus_Success) = No Error
;                               - 7 ($_IEStatus_NoMatch) = No Match
; Author(s):        Dan Pollak
;===============================================================================
;
Func _IEGetActiveTab()
Local $aWinList, $oTab
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

				$gIEhwnd = $aWinList[$i][1]
				Return _IEAttach($aWinList[$i][1],"embedded")	; get active tab in window

			EndIf

			$n += 1

		WEnd

	Next

	Return -1

EndFunc
;===============================================================================
;
; Function Name:    _IEGetPageInNewWindow( $sLink )
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

Func	_IEGetPageInNewWindow( $sLink )

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
; Function Name:    _IEQuitAll( $bKillAll = true )
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

Func	_IEQuitAll( $bKillAll = true )

	Local $oTab
	Local $nTab = 0

	Local $i = 1

	While 1
	   $oTab = _IEAttach( "", "instance", $i )
	   If @error > 0 then ;= $_IESTATUS_NoMatch Then
		  ExitLoop
	   EndIf

		if $bKillAll or not _IEPropertyGet( $oTab, "visible" ) then
			_IEQuit( $oTab )
			$nTab += 1
		Else
			$i += 1
		endif

	WEnd

	return $nTab

EndFunc
