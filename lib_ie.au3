#include-once

#Region Global Include
#include <IE.au3>
#include <IEEx.au3>
#EndRegion Global Include


#cs ----------------------------------------------------------------------------

Function to work with IE
	_IEGetActibTab() - Retrieve the Window Object of the currently active IE (on top)
	_IEGetActibWindow() -
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
Local $hwnd, $oW

	$hwnd = _IEGetActiveWindow()	; get top IE window
	If $hwnd = 0 Then Return -1
	$oW = _IEAttach($hwnd,"embedded")	; get active tab in window
	if @error then return @error
	Return $oW

EndFunc

;===============================================================================
;
; Function Name:    _IEGetBodyText( $sLink )
; Description:      Retrieve body text from web page
; Parameter(s):     $sLink - link to web page
; Returns:
;			Text string with page contents
;           On Failure  - Returns 0 and sets @ERROR
;           @ERROR      - 0 ($_IEStatus_Success) = No Error
;                       - 7 ($_IEStatus_NoMatch) = No Match
;===============================================================================
; Returns web page text
Func	_IEGetPage( $sLink )

	Local $oMain = _IEAttach( "", "instance", 1 )

	Local $o = _IEEx_TabCreate( $oMain, $sLink )
Dbg("Tab create -- " & @error & " " & isobj($o) )

	; check that the link is right
if _IEPropertyGet( $o, "locationurl" ) <> $sLink then
	; we are not on that page: password?
	; wait until you enter password or enter selv?
Dbg("Fail page  -- " & _IEPropertyGet( $o, "locationurl" )  )
	SetError( 2 )
	return 0
EndIf

	Local $html = _IEBodyReadText( $o )
Dbg( "Get body " & @error & " " & isobj($o) )
;MsgBox( 0, "_IEBody Text", $html )
;	FileWrite( "_IEBodyReadText.txt", $html )


Dbg( "Quit " & _IEQuit($o) & " " & @error & " " & isobj($o) )

	return $html

EndFunc

