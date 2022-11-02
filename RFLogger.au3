#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_UseX64=n
#AutoIt3Wrapper_Res_Field=Timestamp|%date% %time%
#AutoIt3Wrapper_Run_Stop_OnError=y
#AutoIt3Wrapper_Run_Before=..\pass.exe add
#AutoIt3Wrapper_Run_After=..\pass.exe remove
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
40 - #35 Add M9.4
15/02/22
41 - #41 Put links to messages in table
18/02/22
42 - #40 Add refToPar/Conv
   - #51 fix space before RefToPar in file name
19/02/22
43 - #40 Add refToPar/Conv
21/02/22
44 - #43 Add M3
   - #51 Fix annullering
   - #52 Add date to msg log
27/02/22
45 - small fixes:
	- change only start time when press new
	- count msges from 1-max (was from max to 0)
	- add egenandel in M10
16/03/22
46 - fix #56 - show all RefNr in M91
17/03/22
47	- rewrite logging - folder with actual date
	- rewrite files
	- only one log file per folder
	- add papirresept i M10
	- add rekvirentNordisk
21/03/22
48	- fix #71
	- extend width to 1000
	- set edit buffer size to 200K
	- fix #72
22/03/22
49	- fix #72
	- show only new messages when GET button pressed
	- refresh IE when NEW button pressed and show new messages
23/03/22
50
	- fix #78
29/03/22
51	- fix #81
	- fix #80
52	- fix #69
08/04/22
53	- add #86 - type legemiddel - partially
	- add #90 - add decoded M1 in M94
54	- add #89 - kill invisible IE at start
20/05/22
	- fix #90 - improved decode b64
	- add #89 - replace re-read with _IEAttach
54.2
	- fix #14 - reopen comment for re-read if file is <1000 bytes
55
24/06/22
	- fix #101 - added name for handelsvarer to M10 output
	- fix #100 - added prodGruppe for handelsvarer to M1 output
56
04/08/22
	-fix #106 - change logfile ext from log to txt
	-fix #99 - add multidosebruker i M92
05/08/22
	- fix #88 - add interval control
	- fix #109 - changed saving to file: full name for xml with folder choice for M1
06/08/22
	- fix #112 - handle AcessDenied page
	- fix #114 - change .test1. to actual env variable used in Get
	- fix #117 - make getting interruptable
57
08/08/22
	- add #25 - add M95/M96
	- add #103 - always on top on/off
25/08/22
	- fix #127 - filename for ERM10
58
28/10/22
	- fix #129 - add M25.3
	- fix #131 - print full name
59
02/11/22
	- fix #133 - add M5
	- fix #134 - correct filename for messages without text
	- fix #21 - add M27.1/2
	- fix #20 - add ERMV
	- fix #60 - Add BytteresRekvirent in M1
	- add #140 - add M25.2 & M25.1
#ce

Local const $nVer = "59"

; #INCLUDES# ===================================================================================================================
#Region Global Include files
#include <Date.au3>
#include <Array.au3>
#include <IE.au3>

#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <EditConstants.au3>
#include <FontConstants.au3>
#include <GuiEdit.au3>
#EndRegion Global Include files

Global $gIEhwnd = -1
Global $gDebugFile = "_debug.txt"

; #LIB# ===================================================================================================================
#Region Lib files
;OnAutoItExitRegister("MyExitFunc")
#include "lib_msg.au3"
#include "lib_ie.au3"
#include "lib_time.au3"
#EndRegion LIB files


; #VARIABLES# ===================================================================================================================
#Region Global Variables
Global $gui
Global $iInterval=1	; current interval, by default 5 min
Global $aIntervals[] = [1,5,10,15,30,60,60*2,60*4,60*8] ; intervals range in minutes

Global $idButtonGet
Global $idButtonClear
Global $idButtonNew
Global $idInterval


Global $idEdit
Global $idLabel
Global $cmAlwaysOnTop
Global $nLine = 0

Global $rf_test_env = "test1"

Global $_abortGet = 0 ; flag to abort execution

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
	Local $gci ; GUI cursor info

	GUI_Create()

	GUIRegisterMsg($WM_MOUSEWHEEL, "_MOUSEWHEEL")
	GUICtrlSetOnEvent( $idButtonGet, "AbortGet" ) ; used in "GUIOnEventMode"


	Do

		; to control when mouse is over interval control
		If WinActive($gui) Then
			$gci = GUIGetCursorInfo($gui)
			If $gci[4] = $idInterval Then
				;; Mouse is over control
				ToolTip("Use mouse wheel to increase/decrease interval in min" )
			Else
				;; Mouse has left control
				ToolTip("")
			EndIf
		EndIf

		$msg = GUIGetMsg()

		if $msg = $idButtonGet then
			;Gui_update_list()
			Get_Button_pressed()
		ElseIf $msg = $idButtonClear then
			;Gui_update_list()
			Clear_Button_pressed()
		ElseIf $msg = $idButtonNew then
			;Gui_update_list()
			New_Button_pressed()
		ElseIf $msg = $cmAlwaysOnTop then
			; Context menu - Always on top
			SetAlwaysOnTop()
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
Local const $guiWidth = 1000
Local const $guiHeight = 300
Local const $guiLeft = -1
Local const $guiTop = -1

; GUI height includes windows title (23px)
; GUI elements start
#include <WinAPI.au3>
Local const $winTitleHeight = _WinAPI_GetSystemMetrics($SM_CYCAPTION)
#include <StaticConstants.au3>


; Create input
	$gui = GUICreate( "Get all active messages - v." & $nVer & "." &  GetVersion(), $guiWidth, $guiHeight, $guiLeft, $guiTop, $WS_MINIMIZEBOX+$WS_SIZEBOX ) ; & GetVersion(), 500, 200)

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

; ----- 3rd button from right <<<--
	$guiBtnLeft = $guiBtnLeft - $guiMargin - $guiBtnWidth

	$idButtonNew = GUICtrlCreateButton("New", $guiBtnLeft, $guiBtnTop, $guiBtnWidth, $guiBtnHeight)
	GUIctrlsetfont(-1, 9, 0, 0, "Lucida Console" )
	GUICtrlSetResizing(-1, $GUI_DOCKRIGHT + $GUI_DOCKTOP + $GUI_DOCKSIZE)

; ----- 4rd control from right <<<--
	$guiBtnLeft = $guiBtnLeft - $guiMargin - $guiBtnWidth

	$idInterval = GUICtrlCreateLabel( GetCurInterval(), $guiBtnLeft, $guiBtnTop, $guiBtnWidth, $guiBtnHeight, $SS_CENTER+$SS_CENTERIMAGE)
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

; Label context menu
	Local $contextmenu = GUICtrlCreateContextMenu($idLabel)
	$cmAlwaysOnTop = GUICtrlCreateMenuItem("Always on top", $contextmenu)
	SetAlwaysOnTop()

	; ----- Edit
	Local const $guiEditLeft = $guiMargin
	Local const $guiEditTop = $guiLabelTop + $guiLabelHeight + $guiMargin
	Local const $guiEditHeight = $guiHeight - $winTitleHeight - $guiLabelTop - $guiLabelHeight - $guiMargin - $guiMargin
	Local const $guiEditWidth = $guiWidth - $guiMargin - $guiEditLeft

	$idEdit = GUICtrlCreateEdit("", $guiEditLeft, $guiEditTop, $guiEditWidth, $guiEditHeight, $ES_READONLY + $ES_AUTOVSCROLL + $WS_VSCROLL)
	GUIctrlsetfont(-1, 9, 0, 0, "Lucida Console" )
	GUICtrlSetResizing(-1, $GUI_DOCKBORDERS )
	_GUICtrlEdit_SetLimitText(-1, 2000*100)

    GUISetState(@SW_SHOW)

    ; start GUI
	GUISetState() ; will display an empty dialog box

EndFunc


;===============================================================================
; Function Name:    GetCurInterval(bool $inMin)
;
; Return: if $inMin
;		false - (default) text string with current interval
;		true - number of minute in interval
;===============================================================================
Func GetCurInterval($inMin = False)
	Local $i = $aIntervals[ $iInterval ]
	if $inMin then Return $i

	if $i <60 then
		return $i & "m"
	else
		return $i/60 & "h"
	EndIf

EndFunc	;==> GetCurInterval

;===============================================================================
; Function Name:    Hent_Button_pressed()
;===============================================================================
Func _MOUSEWHEEL($hWnd, $iMsg, $wParam, $lParam)

	Local $iMPos = MouseGetPos()
	Local $gci = GUIGetCursorInfo($gui)

	If $gci[4] = $idInterval Then
		;; Mouse is over control, do stuff
		$iInterval += _WinAPI_HiWord($wParam) > 0? -1:1

		if $iInterval<0 then
			$iInterval=0
		elseif $iInterval=UBound($aIntervals) then
			$iInterval-=1
		EndIf

		GUICtrlSetData( $idInterval, GetCurInterval() )

	EndIf

    Return $GUI_RUNDEFMSG

EndFunc   ;==>_MOUSEWHEEL

;===============================================================================
; Function Name:    Hent_Button_pressed()
;===============================================================================

Func Clear_Button_pressed()

	GUICtrlSetData($idLabel,"" )
	GUICtrlSetData($idEdit, "" )

EndFunc

;===============================================================================
; Function Name:    New_Button_pressed()
;===============================================================================

Func New_Button_pressed()


	; if this was first time then
	;	create IE
	; if url = login
	;	login with credentials and save them
	;	get to loglist
	; if url = loglist
	;	get old window $oTab
	;	change time to -5 min
	;	save all the fields
	;	submit form

	; clear from invisible objects
	Local $nKilled = _IEQuitAll(false)
	if $nKilled > 0 then
		LogScreen($nKilled & " hidden instances killed" )
	EndIf

	; get Activ IE window
	Local $oTab = _IEGetActiveTab()
	if not IsObj($oTab) then
		; start IE
		; get to loglist!
		_IECreate( "https://rfadmin." & $rf_test_env & ".reseptformidleren.net/RFAdmin/loglist.rfa" )
		;Dbg("*** Error: No active IE tab " & $oTab )
		$oTab = _IEGetActiveTab()
	EndIf

	GUICtrlSetData($idLabel,"1. IE found..." )

	; check that is is logger
	Local $url = _IEPropertyGet( $oTab, "locationurl")

	; if Access denied page
	if StringInStr( $url, "RFAdmin/accessDenied" ) > 0 then
		;https://rfadmin.test1.reseptformidleren.net/RFAdmin/accessDenied.jsp
		_IENavigate ( $oTab, StringRegExpReplace( $url, "/RFAdmin/.*", "/RFAdmin/login.rfa" ) )
		$url = _IEPropertyGet( $oTab, "locationurl")
	EndIf

	; if login
	if StringInStr( $url, "login.rfa" ) > 0 then
		; this is login page
		; https://rfadmin.test1.reseptformidleren.net/RFAdmin/login.rfa
		; enter name
Local $oLoginForm = _IEFormGetObjByName($oTab, "login")
Local $oUserId = _IEFormElementGetObjByName($oLoginForm,  "userId" )
Local $oPass = _IEFormElementGetObjByName($oLoginForm,  "pass" )

_IEFormElementSetValue($oUserId, "")
_IEFormElementSetValue($oPass, "")

_IEFormSubmit($oLoginForm)
		 _IENavigate ( $oTab, StringRegExpReplace( $url, "/RFAdmin/.*", "/RFAdmin/loglist.rfa" ) )
		$url = _IEPropertyGet( $oTab, "locationurl")
	EndIf

	; check that is is logger
	if StringInStr( $url, "loglist.rfa" ) = 0 AND StringInStr( $url, "logsearch.rfa" ) = 0 then
		LogScreen("*** Error: No message log on IE page" & @CRLF & $url )
		return 0
	EndIf
	GUICtrlSetData($idLabel,"2. Got message page..." )

; we are on loglist page!

Local $oForm = _IEFormGetObjByName($oTab, "logfilter")
Local $oDatoFra = _IEFormElementGetObjByName($oForm,  "datoFra" )
Local $oDatoTil = _IEFormElementGetObjByName($oForm,  "datoTil" )
Local $oMsgType = _IEFormElementGetObjByName($oForm,  "msgType" )
Local $oAktor = _IEFormElementGetObjByName($oForm,  "aktor" )
Local $oMsgId = _IEFormElementGetObjByName($oForm,  "msgId" )

; Returns the current Date and Time in format YYYY/MM/DD HH:MM:SS.
Local $sFrom = _RFtimeDiff( -GetCurInterval(True), 'n')
;Local $sTo = _RFtimeDiff(1,'h')

_IEFormElementSetValue($oDatoFra, $sFrom)
;_IEFormElementSetValue($oDatoTil, $sTo)
;_IEFormElementSetValue($oMsgType, "")
;_IEFormElementSetValue($oAktor, "")
;_IEFormElementSetValue($oMsgId, "" ) ;f09601fe-d6c5-4c56-bc2a-b55e49834343")
_IEFormElementCheckBoxSelect($oForm, "WS-R" )
_IEFormElementCheckBoxSelect($oForm, "WS-U" )

;_IEFormSubmit($oForm, 0)
;_IELoadWait($oTab)

Local $oSubmit = _IEGetObjById($oTab, "sokeknapp")
_IEAction($oSubmit, "click")
_IELoadWait($oTab)


; if we got to password page
; get thru password page
; get to loglist
; refill form
; submit
;	Get_Button_pressed()

EndFunc
;===============================================================================
; Function Name:    Hent_Button_pressed()
;===============================================================================

Func	Get_Button_pressed()

	Local $oTab

;GUICtrlSetData($idEdit, "" )
GUICtrlSetData($idLabel, "Get Active IE" )

DbgFile( "start " & _Now()  )

_IELoadWaitTimeout( 3000 )

;Get_line_from_link( "https://rfadmin.test2.reseptformidleren.net/RFAdmin/loggeview.rfa?loggeId=c7d0d0b6-4014-4ef0-be60-d9522d39045a&filename=/nfstest2/sharedFiles/log/2021/175/21/13/c7d0d0b6-4014-4ef0-be60-d9522d39045a" )
;return

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

	; store environment
	$rf_test_env = StringRegExpReplace( $url, ".*?rfadmin\.(.*?)\.reseptformidleren.net.*", "$1")

	Local $txt, $sTextFromTable = ""
	Local $nMsgCount = UBound($aTableData, $UBOUND_ROWS )-1

	GUICtrlSetData($idLabel, "5. Found messages: " & $nMsgCount )

;
; Get links to  message
;
	Local $oLinks = _IELinkGetCollection($oTab)
	if @error <> 0 then
		Dbg( "*** Error: _IELinkGetCollection: " & @error)
		return 0
	EndIf

;
; put links into $aTable
;
	Local $i = 1
	For $oLink In $oLinks
		If StringInStr($oLink.href , "loggeview.rfa?" ) Then
			$aTableData[$i][0] = $oLink.href
			$i += 1
		EndIf
	next

	GUICtrlSetData($idLabel, "6. Got links... " )

; ====== main cycle thru messages =====
	Local $buffer = GUICtrlRead($idEdit)

	Opt("GUIOnEventMode", 1)
	$_abortGet = 0
	GUICtrlSetData($idButtonGet, "Abort" )

	$txt = ""
	For $i = $nMsgCount to 1 step -1

			Local $msgId = $aTableData[$i][1]
			Local $msgTime = StringStripWS($aTableData[$i][2],3)

			if( StringInStr( $buffer, $msgTime ) ) then ContinueLoop

			Local $msgSystem = $aTableData[$i][3]
			Local $msgType = $aTableData[$i][4]
			Local $msgHerId = $aTableData[$i][10]

			$txt =  $msgTime
			$txt &=  " " & StringLeft( $msgId, 9)
			$txt &=  " " & $msgType

			GUICtrlSetData($idLabel, $nMsgCount-$i+1 & "/" & $nMsgCount & " " & $txt)

			Local $sParam = ""
DbgFileClear()
DbgFile( $txt )
			Local $html = _IEGetPageInNewWindow( $aTableData[$i][0] )
;~ 			if StringLen( $html ) < 1000 then
;~ 				DbgFile( $html)
;~ 				$html = _IEGetPageInNewWindow( $aTableData[$i][0] )
;~ 			EndIf

			if @error then
					DbgFile( $html )
					$sParam = $html
					;return 0 ;
			Else
				$html = _ER_GetBody($html)
				$sParam = _ER_GetExtraParam( $html ) ; refPar refCon Text
			EndIf

			; add line to log file
			LogFile( $txt & " " & $sParam & " " & $msgHerId )

			if $_abortGet then
				ExitLoop
			EndIf

			; Add line to screen
			LogScreen( $txt & " " & $sParam & " " & $msgHerId )
			;LogFile( $retText )

			; strip off RefTo before save xml in a file
			$sParam = StringRegExpReplace( $sParam, "(\S{9}\s+\S{9}) ?", "", 1 )

			_save_xml( $html, $sParam )


	Next ; $i

	Opt("GUIOnEventMode", 0)
	GUICtrlSetData($idButtonGet, "Get" )
	GUICtrlSetData($idLabel, $nMsgCount-$i & "/" & $nMsgCount )

;GUICtrlSetData($idEdit, $txt & @CRLF, 0)
;GUICtrlSetData($idLabel, $nMsgCount & " messages found")

EndFunc

Func	LogScreen( $text )

	_GUICtrlEdit_AppendText($idEdit, $text & @CRLF)

EndFunc



Func	_save_xml( $html, $text, $folder=Default )

Local	$fileName
Local	$fileTime

	$fileTime = _ER_GetMsgTime( $html ) ; 20220117192033
	if $fileTime = 0 then return 1

	if $folder=Default then
		$folder = StringRegExpReplace( $fileTime, "(\d\d\d\d)(\d\d)(\d\d).*", "$1-$2-$3" ) ; 2022-01-17
	EndIf

	; strip off birthdate from NIN
	$text = StringRegExpReplace( $text, "\s(\d{11})\s+(\d\d\d\d).(\d\d).(\d\d)", " $1" )
	$filename = StringReplace( StringStripWS(_ER_GetMsgType( $html ) & " " & $text, 7), " ", "_") & "_" & StringLeft( _ER_GetMsgId( $html ), 9)  & ".xml"


	if not FileExists( $folder ) then
		DirCreate( $folder )
	endif

	FileDelete( $folder & "\" &  $fileName )

	if FileWrite( $folder & "\" &  $fileName, $html ) = 0 then
		Dbg("error write file " & $fileName)
		return 3
	EndIf

	If FileSetTime( $folder & "\" &  $fileName, $fileTime, 0) = 0 then
		Dbg("error filesettime '" & $fileTime & "'->" & $fileName)
		return 4
	EndIf

	Return 0
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

Func	LogFile( $sParam )

	Local $folder = StringRegExpReplace( $sParam , "(\d+).(\d+).(\d\d\d\d) .*", "$3-$2-$1") ; 2022-01-17
	if @extended = 0 then
		return 1
	EndIf

	Local $fileName = $folder & "\" & $folder & "_" & $rf_test_env & ".txt" ; 2022-01-17\2022-01-17_rf.txt

	if not FileExists( $folder ) then
		DirCreate( $folder )
	endif

	FileWriteLine( $fileName , $sParam )

	return 0

EndFunc

; -----------------------------------------------------------------------------
; Function: GetVersion
;
; Return: String yyyymmddhhmm
;
; -----------------------------------------------------------------------------

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


Func	AbortGet()
	$_abortGet = 1
EndFunc

; -----------------------------------------------------------------------------
; Function: SetAlwaysOnTop
;
; Toggles on/off Always on top property for GUI
;
; -----------------------------------------------------------------------------

Func SetAlwaysOnTop()
	If BitAnd(GUICtrlRead($cmAlwaysOnTop),$GUI_CHECKED) = $GUI_CHECKED Then
		GUICtrlSetState($cmAlwaysOnTop,$GUI_UNCHECKED)
		WinSetOnTop($gui, "", $WINDOWS_NOONTOP)
	Else
		GUICtrlSetState($cmAlwaysOnTop,$GUI_CHECKED)
		WinSetOnTop($gui, "", $WINDOWS_ONTOP)
	EndIf
EndFunc
