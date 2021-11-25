#include-once

#Region Global Include
#include <IE.au3>
#include <IEEx.au3>
#EndRegion Global Include


#cs ----------------------------------------------------------------------------
Function to work with IE
	_IEGetActibTab() - Retrieve the Window Object of the currently active IE (on top)
	_IEGetActibWindow() -



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
; Function Name:    _IEGetBodyText( $sLink )
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
Func	_IEGetPage( $sLink )

	DbgFile( "-->_IEGetPage " & $sLink )

	Local $err = 0
	Local $oMain = _IEAttach( $gIEhwnd, "hwnd" ) ;_IEAttach( "", "instance", 1 )

	DbgFile( "   _IEAttach" )

	Local $o = _IEEx_TabCreate( $oMain, $sLink )
;Dbg("Tab create -- " & @error & " " & isobj($o) )
	if @error then
		return SetError(1, 0, "*** Error _TabCreate " & @error & " " & @extended )
	EndIf
	DbgFile( "   _IEEX_TabCreate" )

Local $tries = 1
Local $sAdr = _IEPropertyGet( $o, "locationurl" )

; check that the link is right
; we are not on that page: password?
; wait until you enter password or enter selv?
While $sAdr <> $sLink

	DbgFile( "   " & $tries & ":" & $sAdr )

	if $tries <= 10 then ; 2 sec
		; Wait until all code loaded! otherwise window.onLoad is not completed?
		Sleep( 200 )

		$tries = $tries + 1

		$sAdr = _IEPropertyGet( $o, "locationurl" )

	Else

		;Dbg("Fail page  -- " & _IEPropertyGet( $o, "locationurl" )  )
		 Return SetError(2, 0,"*** Error opened link failed: " & @CRLF & $sLink & @CR & $sAdr )
	EndIf


WEnd

DbgFile( "   _IEPropertyGet" )

	Local $html = _IEBodyReadText( $o )
	;Dbg( "Get body " & @error & " " & isobj($o) )
	if @error then
		Return SetError(3, 0, "*** Error _IEBodyReadText " & @error & " " & @extended )
	endif
DbgFile( "   _IEBodyReadText " & StringLen( $html) & " " & StringLeft( StringStripWS( $html, 8), 15  ))
;FileWrite( _ER_GetParam( $sLink, '(?s)loggeId=(.*?)-' ) & ".txt", $html)

	_IEQuit($o)
;Dbg( "Quit " & _IEQuit($o) & " " & @error & " " & isobj($o) )
	if @error then
		Return SetError(4, 0, "*** Error _IEQuit" & @error & " " & @extended )
	endif
DbgFile( "<--_IEQuit" )

	return $html

EndFunc
