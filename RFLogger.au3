#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_UseX64=n
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#cs
================================
Simple GUI module to create xml file from messagelist in IE

Update History:
15/06/21
1 - initial prototype
16/6/21
4 - working version with reading from RF table (no saving/decoding)
5 - get links from linkscollection and show them on screen
6 - get html and parse it for Name, NavnFormStyrke
7 - open link in tab and save in file - works but does not get all text from msg
8 - Save file with msg but wrong MsgId
9 - Saves messages but not msgId.
10 - strip <> and <MsgId> and without MsgBox - works fine v1.0
11 - remove extra debug info
12 - remove extra logging and make resizable -
13 - uncomment print MsgType and Id - some times not working with many msgs?
21/6/21
14 - change output (swap with title) and remove double output
15 - output time format change yyymmddhhmmss -> hh:mm:ss & filename w/o time + timestamp
16 - add version
23/6/21
17 - get riktig MsgType, MsgTime from RF
24/6/21
18 - replace ClipGet with GetBodyText
25/6/21
19 - move Msg and _IE function in own files: lib_ie.au3 lib_msg.au3

================================
#ce

; #INCLUDES# ===================================================================================================================
#Region Global Include files
#include <Date.au3>
#include <Array.au3>
#include <IE.au3>

#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <EditConstants.au3>
#include <FontConstants.au3>
#EndRegion Global Include files

; #LIB# ===================================================================================================================
#Region Lib files
;OnAutoItExitRegister("MyExitFunc")
#include "lib_msg.au3"
#include "lib_ie.au3"
#EndRegion LIB files


; #VARIABLES# ===================================================================================================================
#Region Global Variables
Global $idButton
Global $idEdit
Global $idLabel
Global $nLine = 0
#EndRegion Global Variables

; #CONSTANTS# ===================================================================================================================


;OnAutoItExitRegister("MyExitFunc")
Opt('MustDeclareVars', 1)

Main()

;===============================================================================
; Function Name:  Main
;===============================================================================
Func	Main()

	Local $msg

	GUI_Create()

	Do

		$msg = GUIGetMsg()

		if $msg = $idButton then
			;Gui_update_list()
			Hent_Button_pressed()
		EndIf

	Until $msg = $GUI_EVENT_CLOSE

	GUIDelete()

	Exit
EndFunc

;===============================================================================
;
; Function Name:    GUI_Create()
; Description:      Create GUI form
; Parameter(s):     None
; Returns:          None
;===============================================================================

Func GUI_Create()

	; Create input

	GUICreate( "Get all active messages - " & GetVersion(), 600,200,-1,-1,$WS_MINIMIZEBOX+$WS_SIZEBOX ) ; & GetVersion(), 500, 200)


	$idButton = GUICtrlCreateButton("Hent", 540, 10, 50, 30)
	GUIctrlsetfont(-1, 9, 0, 0, "Lucida Console" )
	$idEdit = GUICtrlCreateEdit("", 10, 50, 580, 120, $ES_READONLY + $ES_AUTOVSCROLL + $WS_VSCROLL)
	GUIctrlsetfont(-1, 9, 0, 0, "Lucida Console" )
	$idLabel = GUICtrlCreateLabel(	"", 10, 20, 500, 30)
	GUIctrlsetfont(-1, 9, 0, 0, "Lucida Console" )

    GUISetState(@SW_SHOW)

    ; start GUI
	GUISetState() ; will display an empty dialog box

EndFunc

;===============================================================================
; Function Name:    Hent_Button_pressed()
;===============================================================================

Func	Hent_Button_pressed()

	Local $oTab

;GUICtrlSetData($idEdit, "" )
GUICtrlSetData($idLabel, "Get Active IE" )

;Get_line_from_link( "https://rfadmin.test2.reseptformidleren.net/RFAdmin/loggeview.rfa?loggeId=c7d0d0b6-4014-4ef0-be60-d9522d39045a&filename=/nfstest2/sharedFiles/log/2021/175/21/13/c7d0d0b6-4014-4ef0-be60-d9522d39045a" )
;return

	; get Activ IE window
	$oTab = _IEGetActiveTab()
	if not IsObj($oTab) then
		; start IE
		;_IECreate( "https://rfadmin.test1.reseptformidleren.net/RFAdmin/loglist.rfa" )
		Dbg("*** Error: No active IE tab " & $oTab )
		return 0
	EndIf
	GUICtrlSetData($idLabel,"1. IE found..." )

	; check that is is logger
	Local $url = _IEPropertyGet( $oTab, "locationurl")
	if StringInStr( $url, "loglist.rfa" ) = 0 AND StringInStr( $url, "logsearch.rfa" ) = 0 then
		Dbg("*** Error: No message log on IE page" & @CRLF & $url )
		return 0
	EndIf
	GUICtrlSetData($idLabel,"2. Got message page..." )

	; get all links on the page
	Local $oTable = _IETableGetCollection($oTab,0)
	if @error <> 0 then
		Dbg("*** Error: No message table found on page: " & @error & @CRLF & $url )
		return 0
	EndIf
	GUICtrlSetData($idLabel,"3. Got message table... " )

	Local $aTableData = _IETableWriteToArray($oTable, true)
	if @error <> 0 then
		Dbg("*** Error: _IETableWriteToArray: " & @error & @CRLF & $url )
		return 0
	EndIf
	GUICtrlSetData($idLabel, "4. Got message array..." )

	; Check that it is rigth table
	if $aTableData[0][0] <> "Linker" then
		Dbg("*** Error: No message table found on " & @CRLF & $url )
		return 0
	EndIf
;_ArrayDisplay($aTableData)


	Local $txt, $sTextFromTable = ""
	Local $nMsgCount = UBound($aTableData, $UBOUND_ROWS )-1

	GUICtrlSetData($idLabel, "5. Found messages: " & $nMsgCount )

;
; Get link to Hent message
; Link# is equal to Array#
;
	Local $oLinks = _IELinkGetCollection($oTab)
	if @error <> 0 then
		Dbg( "*** Error: _IELinkGetCollection: " & @error)
		return 0
	EndIf
	GUICtrlSetData($idLabel, "6. Got links..." )

	;Local $iNumLinks = @extended

	$txt = ""
	Local $i = 1
	For $oLink In $oLinks
		If StringInStr($oLink.href , "loggeview.rfa?" ) Then
			; process link

						; 12.07.2019 18:15:04.275
			Local $msgId = $aTableData[$i][1]
			Local $msgTime = $aTableData[$i][2]
			Local $msgType = $aTableData[$i][4]

			$txt =  $msgTime
			$txt &=  " " & $msgType
			$txt &=  " " & $msgId

			GUICtrlSetData($idLabel, $i & "/" & $nMsgCount & " " & $txt)

			Local $html = _IEGetPage( $oLink.href )

			$html = _ER_GetBody($html)

			Local $fname = $msgType & "_" & $msgId & ".xml"

			if _save_xml( $fname, $html ) then

				;12.07.2019 18:15:04.275
				Local $t = StringRegExpReplace( $msgTime, "(\d+).(\d+).(\d\d\d\d) (\d\d).(\d\d).(\d\d)", "$3$2$1$4$5$6")
				If FileSetTime( $fname, $t, 0) = 0 then
					Dbg("error filesettime " & $fname )
				EndIf
			EndIf

			$i += 1

			;Local $ret = $msgTime & " " & $msgType & " " & $msgId

			GUICtrlSetData($idEdit, StringMid( $msgTime, 12, 8) & " " & $msgType & " " & $msgId & @CRLF, 0)


		EndIf
	Next

	GUICtrlSetData($idLabel, $i-1 & "/" & $nMsgCount )

;GUICtrlSetData($idEdit, $txt & @CRLF, 0)
;GUICtrlSetData($idLabel, $nMsgCount & " messages found")

EndFunc




Func	_save_xml( $fname, $html )

	; get msgid fra link
	;Local $sec=_DateDiff ( "s", "2021/1/1 00:00:00", _NowCalc() )
	;Local $fname = _ER_GetMsgType( $html) & "_" & $msgId&".xml"
	Local $h = FileOpen( $fname, 2)
	if FileWrite( $h, $html ) = 0 then
		Dbg("error write file " & $fname )
		return 0
	ElseIf FileClose( $h) = 0 then
		Dbg("error close file " & $fname )
		Return 0
	EndIf

	Return 1
EndFunc


Func	Dbg( $txt )
	GUICtrlSetData($idEdit, $txt & @CRLF, 0)

EndFunc


; -----------------------------------------------------------------------------
; Function: GetVersion
;
; Return: String yyyymmddhhmm
;
; -----------------------------------------------------------------------------
#AutoIt3Wrapper_Res_Field=Timestamp|%date% %time%

Func GetVersion()

	Local $ver
	If @Compiled Then

		$ver = FileGetVersion(@ScriptFullPath, "Timestamp")

		; dd/mm/yyyy.hh.mm.ss -> dd.mm.yy.hhmm
		$ver = StringRegExpReplace( $ver, "(\d+)[/\-\.](\d+)[/\-\.]\d*(\d\d).(\d+):(\d+):\d+" , "$1.$2.$3.$4$5" )

	Else
		$ver = "Not compiled"
	EndIf

	Return $ver

EndFunc
