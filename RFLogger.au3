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
26/6/21
20 - add specific GetMsgXXX functions
21 - remove debug output to edit control
2/7/2021
22 - do not overwrite file if it exists. To protect from messages with the same ID which first failed (#11)
23 - fixed #6 add button Clear and Rename Hent to Get
26/8/21
24 - fixed #8 Can't find active tab
27/8/21
25 - Always add logg to the end of text
5/9/21
26 - fixed #5 Apprec doesn't work -> diff msg format for apprec
   - when file exists - return 2
6/9/21
27 - fixed #24 - Resize controls when resize window
28 - added #25 - Add M9.5
24/11/21
29 - added global hwnd to fix _IEAttach
   - added global logfileName
25/11/21
30 - wait until webpage is loaded
26/11/21
31 - save files in own folder
32 - change forlder name to msg timestamp
   - open msg xml in a new hidden window
   - if msg size <1000 read again
   - add M91, M93
33 - remove _IEGetPAge
34 - add version to title
12/02/22
35 - #29 Resizable GUI
36 - #26 Add M9.11/M9.12
37 - Added resept count to M912 (_ER_GetReseptCount)
38 - #32 Add M9.2
39 - #36 Fixed log file
================================
#ce
Local const $nVer = "39"

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

Global $gIEhwnd = -1
Global $gLogFolder = "."
Global $gLogfile = "log"
Global $gDebugFile = "_debug.txt"

; #LIB# ===================================================================================================================
#Region Lib files
;OnAutoItExitRegister("MyExitFunc")
#include "lib_msg.au3"
#include "lib_ie.au3"
#EndRegion LIB files


; #VARIABLES# ===================================================================================================================
#Region Global Variables
Global $idButtonGet
Global $idButtonClear
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

		if $msg = $idButtonGet then
			;Gui_update_list()
			Get_Button_pressed()
		ElseIf $msg = $idButtonClear then
			;Gui_update_list()
			Clear_Button_pressed()
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

;--- space between elements
Local const $guiMargin = 10

;--- GUI
Local const $guiWidth = 800
Local const $guiHeight = 300
Local const $guiLeft = -1
Local const $guiTop = -1

; GUI height includes windows title (23px)
; GUI elements start
#include <WinAPI.au3>
Local const $winTitleHeight = _WinAPI_GetSystemMetrics($SM_CYCAPTION)


	; Create input
	GUICreate( "Get all active messages - v." & $nVer & "." &  GetVersion(), $guiWidth, $guiHeight, $guiLeft, $guiTop, $WS_MINIMIZEBOX+$WS_SIZEBOX ) ; & GetVersion(), 500, 200)

	;--- buttons starting from right
	Local const $guiBtnWidth = 50
	Local const $guiBtnHeight = 30
	Local $guiBtnLeft = $guiWidth - $guiMargin - $guiBtnWidth
	Local $guiBtnTop = $guiMargin

	; ----- 1st from right button
	;$guiBtnLeft = $guiWidth - $guiMargin - $guiBtnWidth
	;$guiBtnTop = $guiMargin

	$idButtonGet = GUICtrlCreateButton("Get", $guiBtnLeft, $guiBtnTop, $guiBtnWidth, $guiBtnHeight)
	GUIctrlsetfont(-1, 9, 0, 0, "Lucida Console" )
	GUICtrlSetResizing(-1, $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKSIZE)

	; ----- 2nd button from right <<<--
	$guiBtnLeft = $guiBtnLeft - $guiMargin - $guiBtnWidth

	$idButtonClear = GUICtrlCreateButton("Clear", $guiBtnLeft, $guiBtnTop, $guiBtnWidth, $guiBtnHeight)
	GUIctrlsetfont(-1, 9, 0, 0, "Lucida Console" )
	GUICtrlSetResizing(-1, $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKSIZE)

	; ----- Label
	Local const $guiLabelLeft = $guiMargin
	Local const $guiLabelTop = $guiBtnTop
	Local const $guiLabelWidth = $guiBtnLeft - $guiMargin - $guiLabelLeft
	Local const $guiLabelHeight = $guiBtnHeight

	$idLabel = GUICtrlCreateLabel(	"", $guiLabelLeft, $guiLabelTop, $guiLabelWidth, $guiLabelHeight)
	GUIctrlsetfont(-1, 9, 0, 0, "Lucida Console" )
	GUICtrlSetResizing(-1, $GUI_DOCKRIGHT + $GUI_DOCKLEFT + $GUI_DOCKTOP + $GUI_DOCKHEIGHT )

	; ----- Edit
	Local const $guiEditLeft = $guiMargin
	Local const $guiEditTop = $guiLabelTop + $guiLabelHeight + $guiMargin
	Local const $guiEditHeight = $guiHeight - $winTitleHeight - $guiLabelTop - $guiLabelHeight - $guiMargin - $guiMargin
	Local const $guiEditWidth = $guiWidth - $guiMargin - $guiEditLeft

	$idEdit = GUICtrlCreateEdit("", $guiEditLeft, $guiEditTop, $guiEditWidth, $guiEditHeight, $ES_READONLY + $ES_AUTOVSCROLL + $WS_VSCROLL)
	GUIctrlsetfont(-1, 9, 0, 0, "Lucida Console" )
	GUICtrlSetResizing(-1, $GUI_DOCKBORDERS )

    GUISetState(@SW_SHOW)

    ; start GUI
	GUISetState() ; will display an empty dialog box

EndFunc

;===============================================================================
; Function Name:    Hent_Button_pressed()
;===============================================================================

Func Clear_Button_pressed()

	GUICtrlSetData($idLabel,"" )
	GUICtrlSetData($idEdit, "" )

EndFunc

;===============================================================================
; Function Name:    Hent_Button_pressed()
;===============================================================================

Func	Get_Button_pressed()

	Local $oTab
	$gLogFolder = @YEAR & "." & @MON & "." & @MDAY ; folder will correspond to timestamp
	$gLogfile = @YEAR & "." & @MON & "." & @MDAY & "_" & @HOUR & @MIN & @SEC & "_log.txt"

;GUICtrlSetData($idEdit, "" )
GUICtrlSetData($idLabel, "Get Active IE" )

DbgFile( "start " & _Now()  )
LogFile( "start " & _Now()  )

_IELoadWaitTimeout( 3000 )

;Get_line_from_link( "https://rfadmin.test2.reseptformidleren.net/RFAdmin/loggeview.rfa?loggeId=c7d0d0b6-4014-4ef0-be60-d9522d39045a&filename=/nfstest2/sharedFiles/log/2021/175/21/13/c7d0d0b6-4014-4ef0-be60-d9522d39045a" )
;return

	; move to the end of text
	Local $cEnd = StringLen( GUICtrlRead($idEdit) )
	GuiCtrlSendMsg($idEdit, $EM_SETSEL, $cEnd, $cEnd )
	;GUICtrlSetData($idEdit, _NowTime(), 0)

	; get Activ IE window
	$oTab = _IEGetActiveTab()
	if not IsObj($oTab) then
		; start IE
		_IECreate( "https://rfadmin.test1.reseptformidleren.net/RFAdmin/loglist.rfa" )
		;Dbg("*** Error: No active IE tab " & $oTab )
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
			Local $msgSystem = $aTableData[$i][3]
			Local $msgType = $aTableData[$i][4]
			Local $msgHerId = $aTableData[$i][10]

			$txt =  $msgTime
			$txt &=  " " & $msgType
			$txt &=  " " & $msgId

			GUICtrlSetData($idLabel, $i & "/" & $nMsgCount & " " & $txt)

			Local $sParam
DbgFileClear()
DbgFile( $txt )
			Local $html = _IEGetPageInNewWindow( $oLink.href )
	if StringLen( $html ) < 1000 then
		DbgFile( $html)
		$html = _IEGetPageInNewWindow( $oLink.href )
	EndIf
		if @error then
				DbgFile( $html)
				$sParam = $html
				;return 0 ;
		Else
			$html = _ER_GetBody($html)

			$sParam = _ER_GetExtraParam( $html )

			Local $fname = $msgType & StringReplace( $sParam, " ", "_" ) & "_" & StringLeft( $msgId, 9) & ".xml"

			$sParam = $sParam & " " & $msgHerId

			Local $t = StringRegExpReplace( $msgTime, "(\d+).(\d+).(\d\d\d\d) (\d\d).(\d\d).(\d\d).*", "$3$2$1$4$5$6")
			;$gLogFolder = StringLeft( $t, 8)

			switch _save_xml( $fname, $html )
			Case 1 ; ok
				;12.07.2019 18:15:04.275
				If FileSetTime( $gLogFolder & "/"& $fname, $t, 0) = 0 then
					Dbg("error filesettime '" & $t & "'->" & $fname )
				EndIf
			Case 2 ; file exists
				$sParam = $sParam & " *** file exists"
			Case 0 ; error saving file
				$sParam = $sParam & " *** not saved"
			EndSwitch
		EndIf
			$i += 1

			Local $retText = StringMid( $msgTime, 12, 8) & " " & StringLeft( $msgId, 9) & " " & $msgType & " " & $sParam

			GUICtrlSetData($idEdit, $retText & @CRLF, 0)
			LogFile( $retText )

		EndIf
	Next

	GUICtrlSetData($idLabel, $i-1 & "/" & $nMsgCount )

;GUICtrlSetData($idEdit, $txt & @CRLF, 0)
;GUICtrlSetData($idLabel, $nMsgCount & " messages found")

EndFunc




Func	_save_xml( $fname, $html )


	if FileExists( $gLogFolder & "/" & $fname ) then
		return 2 ; do not overwrite
	endif

	if FileWrite( $gLogFolder & "/" &  $fname, $html ) = 0 then
		Dbg("error write file " & $fname )
		return 0
	EndIf

	Return 1
EndFunc


Func	Dbg( $txt )
	GUICtrlSetData($idEdit, $txt & @CRLF, 0)
EndFunc

func DbgFileClear()
	;FileDelete( $gDebugFile )
EndFunc

Func	DbgFile( $txt )
	FileWriteLine( $gDebugFile, $txt )
EndFunc

Func	LogFile( $txt )

	if not FileExists( $gLogFolder ) then
			DirCreate( $gLogFolder )
	EndIf
	FileWriteLine( $gLogFolder & "/" & $gLogfile, $txt )

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
