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
================================
#ce

; #VARIABLES# ===================================================================================================================
#Region Global Variables
Global $idButton
Global $idEdit
Global $idLabel
Global $nLine = 0
#EndRegion Global Variables

; #CONSTANTS# ===================================================================================================================


;OnAutoItExitRegister("MyExitFunc")
#include <Date.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <EditConstants.au3>
#include <IE.au3>
#include <IEEx.au3>
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

#include "_ER_ie.au3"

Func	Gui_update_list()

	Local $oTab

GUICtrlSetData($idEdit, "" )
GUICtrlSetData($idLabel, "Getting messages..." )

;Get_line_from_link( "C:\Users\ang_\Documents\github\Msgdump\ex.html" )
;return

	; get Activ IE window
	$oTab = _IEGetActiveTab()
	if not IsObj($oTab) then
		; start IE
		;_IECreate( "https://rfadmin.test1.reseptformidleren.net/RFAdmin/loglist.rfa" )
		Dbg("*** Error: No active IE window " & $oTab )
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

#cs
	; Get times and Msgs from table
	For $i = 1 To $nMsgCount
            $txt = ""
			$txt &=  $aTableData[$i][1]
			$txt &=  " " & $aTableData[$i][2]
			$txt &=  " " & $aTableData[$i][3]
			$txt &=  " " & $aTableData[$i][4]
			$txt &=  @CRLF
			GUICtrlSetData($idLabel, $i & "/" & $nMsgCount & " messages done")
			GUICtrlSetData($idEdit, $txt, 0)
    Next
#ce
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

			$txt =  $aTableData[$i][1]
			$txt &=  " " & $aTableData[$i][2]
			$txt &=  " " & $aTableData[$i][3]
			$txt &=  " " & $aTableData[$i][4]
			GUICtrlSetData($idLabel, $i & "/" & $nMsgCount & " " & $txt)

			Local $ret = Get_line_from_link( $oLink.href ) & @CRLF

			$i += 1
			GUICtrlSetData($idEdit, $ret, 0)

		EndIf
	Next

	GUICtrlSetData($idLabel, $i-1 & "/" & $nMsgCount )

;GUICtrlSetData($idEdit, $txt & @CRLF, 0)
;GUICtrlSetData($idLabel, $nMsgCount & " messages found")

EndFunc

;#include "_ER_msg.au3"

; Returns one line with msg
Func	Get_line_from_link( $sLink )

	Local $oMain = _IEAttach( "", "instance", 1 )
	Local $o = _IEEx_TabCreate( $oMain, $sLink )
#cs
	Local $o = _IECreate( $sLink ) ;, 0, 1, 1 )
	if not isObj($o) then
		Dbg( "*** Error: _IECreate: " & @error & @CRLF & $sLink )
		return 0
	EndIf
	$o = _IEAttach( $sLink, "url" )
#ce
	;MsgBox( 0, "Created", @error & @CRLF & _IEPropertyGet( $o, "locationurl") )

;Dbg("IE created -- " & @error )

	Local $oWindow = $o.document.parentWindow
	_IEAction($oWindow, "blur")

;Dbg("Blur-- " & @error )
	_IEAction( $o, "selectall" )
;Dbg("Selectall-- " )
	_IEAction( $o, "copy")
;Dbg("Copy " & @error )
	Local $html = ClipGet() ; _IEBodyReadText($o)
;Dbg("ClipGet" & @error )

	_IEQuit($o)
	;Dbg("IE closed --"  )


	$html = _ER_GetBody($html)
;MsgBox( 0, "Get from clip", StringLeft( $html, 200 ))

;MsgBox( 0, "Strip <Msg>", StringLeft( StringStripWS($html,7), 300 ))

	Local $msgId = _ER_GetMsgId( $html)
;MsgBox( 0, "Get MsgId", $msgId )
	Local $msgType = _ER_GetMsgType( $html)
;Dbg( $msgType & " " & $msgId & @CRLF )

	Local $msgTime = StringRegExpReplace( _ER_GetMsgTime( $html), "(\d{8})(\d\d)(\d\d)(\d\d)", "$2:$3:$4")
;Dbg($msgType & " " & $msgId )

	; get msgid fra link
	;Local $sec=_DateDiff ( "s", "2021/1/1 00:00:00", _NowCalc() )
	Local $fname = _ER_GetMsgType( $html) & "_" & $msgId&".xml"
	Local $h = FileOpen( $fname, 2)
	if FileWrite( $h, $html ) = 0 then
		Dbg("error write file " & $fname )
	ElseIf FileClose( $h) = 0 then
		Dbg("error close file " & $fname )
	ElseIf FileSetTime( $fname, _ER_GetMsgTime($html), 0) = 0 then
		Dbg("error filesettime " & $fname )
	EndIf

	Return  $msgTime & " " & $msgType & " " & $msgId  ;_IESaveMessage( $html )

EndFunc


Func	Dbg( $txt )
	GUICtrlSetData($idEdit, $txt & @CRLF, 0)

EndFunc

Func	_ER_GetMsgId( $html)

	Local $a = StringRegExp( $html, 'MsgId>([0-9a-fA-F\-]+?)</', 1)
	if @error then return 0
	return $a[0]

EndFunc

Func	_ER_GetMsgType( $html)

	Local $a = StringRegExp( $html, '(?s)MsgInfo>.*?<.*?Type.*?V="(.*?)"', 1)
	if @error then return 0
	return $a[0]

EndFunc

Func	_ER_GetBody( $html)

	Local $a = StringRegExp( $html, '(?s)(<.*>)', 1)
	if @error then return 0

	; strip signatur value
	;Local $s = StringRegExpReplace( $a[0], '(?s)(SignatureValue>.{5}).*?(</)', '$1...$2')
	;if @error then return $a[0]

	; strip X509 certificate
	;$s = StringRegExpReplace( $s, '(?s)(X509Certificate>.{5}).*?(</)', '$1...$2')
	;if @error then return $s

	return $a[0]

EndFunc

Func	_ER_GetMsgTime( $html)

	Local $a
	$a = StringRegExp( $html, 'GenDate>(\d+).(\d+).(\d+).(\d+):(\d+):(\d+).*?</', 1)
	if @error then return 0
	Return $a[0] & $a[1] &  $a[2] & $a[3] & $a[4] & $a[5]

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
