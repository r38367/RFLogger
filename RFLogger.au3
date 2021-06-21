#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_UseX64=n
;#AutoIt3Wrapper_Res_Field=Timestamp|%date%.%time%
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#cs
================================
Simple GUI module to create xml file from messagelist in IE

Update History:
15/06/21 - initial prototype

================================
#ce

; #VARIABLES# ===================================================================================================================
#Region Global Variables
Global $idButton
Global $idEdit
Global $nLine = 0
#EndRegion Global Variables

; #CONSTANTS# ===================================================================================================================


;OnAutoItExitRegister("MyExitFunc")
#include <Date.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <EditConstants.au3>
#include <IE.au3>
#include <Array.au3>
#include <FontConstants.au3>
#include <GUIConstantsEx.au3>



Opt('MustDeclareVars', 1)

; Timestamp

Local $msg

GUI_Create()

Do

	$msg = GUIGetMsg()

	if $msg = $idButton then
		Gui_update_list()
	EndIf

Until $msg = $GUI_EVENT_CLOSE

GUIDelete()

Exit

;===============================================================================
;
; Function Name:    GUI_Create()
; Description:      Create GUI form
; Parameter(s):     None
; Returns:          None
;===============================================================================
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include "getversion.au3"

Func GUI_Create()

	; Create input

	GUICreate( "Get all active messages - ", 500,200 ) ; & GetVersion(), 500, 200)

	$idButton = GUICtrlCreateButton("Hent", 450, 10)
	$idEdit = GUICtrlCreateEdit("", 10, 50, 480, 120, $ES_READONLY + $ES_AUTOVSCROLL + $WS_VSCROLL)
	GUIctrlsetfont(-1, 9, 0, 0, "Lucida Console" )

    GUISetState(@SW_SHOW)

    ; start GUI
	GUISetState() ; will display an empty dialog box

EndFunc

#include "_ER_ie.au3"

Func	Gui_update_list()

	Local $oTab

GUICtrlSetData($idEdit, "" , 0)

	; get Activ IE window
	$oTab = _IEGetActiveTab()
	if $oTab = 0 then
		; start IE
		;_IECreate( "https://rfadmin.test1.reseptformidleren.net/RFAdmin/loglist.rfa" )
		Dbg("*** Error: No active IE window")
		return 0
	EndIf

	; check that is is logger
	Local $url = _IEPropertyGet( $oTab, "locationurl")
	if StringInStr( $url, "loglist.rfa" ) = 0 AND StringInStr( $url, "logsearch.rfa" ) = 0 then
		Dbg("*** Error: No message log on IE page" & @CRLF & $url )
		return 0
	EndIf

	; get all links on the page
	Local $oTable = _IETableGetCollection($oTab,0)
	if @error <> 0 then
		Dbg("*** Error: No message table found on page: " & @error & @CRLF & $url )
		return 0
	EndIf

	Local $aTableData = _IETableWriteToArray($oTable, true)
	if @error <> 0 then
		Dbg("*** Error: _IETableWriteToArray: " & @error & @CRLF & $url )
		return 0
	EndIf

	; Check that it is rigth table
	if $aTableData[0][0] <> "Linker" then
		Dbg("*** Error: No message table found on " & @CRLF & $url )
		return 0
	EndIf
;_ArrayDisplay($aTableData)


	Local $txt, $sTextFromTable = ""
	Local $nMsgCount = UBound($aTableData, $UBOUND_ROWS )-1
	$txt = $nMsgCount & " messages found"
	; Get times and Msgs from table
	For $i = 1 To $nMsgCount
            $txt &=  $aTableData[$i][1]
			$txt &=  " " & $aTableData[$i][2]
			$txt &=  " " & $aTableData[$i][3]
			$txt &=  " " & $aTableData[$i][4]
			$txt &=  @CRLF
    Next

#cs
;
; Get link to Hent message
; Link# is equal to Array#
;
	Local $oLinks = _IELinkGetCollection($oTab)
	if @error <> 0 then
		MsgBox( 0, "Error ", "_IELinkGetCollection: " & @error)
		return 0
	EndIf

	;Local $iNumLinks = @extended

	$txt = ""

	For $oLink In $oLinks
		If StringInStr($oLink.href , "loggeview.rfa?" ) Then
			; process link
			$txt &= Get_line_from_link( $oLink.href ) & @CRLF

		EndIf
	Next
#ce
GUICtrlSetData($idEdit, $txt & @CRLF, 0)

EndFunc
#cs
;#include "_ER_msg.au3"
; Returns one line with msg
Func	Get_line_from_link( $sLink )

	Return $sLink

	Local $oIE = _IECreate( $sLink, 0, 1 )
	if @error <> 0 then
		Dbg( "Error: _IECreate: " & @error & @CRLF & $sLink )
		return 0
	EndIf

	Sleep(1000)

	Local $html = _IEBodyReadText( $oIE )
	_IEQuit($oIE)

	Return _IESaveMessage( $html )

EndFunc
#ce

Func	Dbg( $txt )
	GUICtrlSetData($idEdit, $txt & @CRLF, 0)

EndFunc
