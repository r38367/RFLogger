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
17/08/22
	- add #119 - separate ie-dependent functions
#ce

Local const $nVer = "56"

; #INCLUDES# ===================================================================================================================
#Region Global Include files
;#include <Date.au3>
;#include <Array.au3>

#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <EditConstants.au3>
#include <FontConstants.au3>
#include <GuiEdit.au3>
#EndRegion Global Include files


; #LIB# ===================================================================================================================
#Region Lib files
;OnAutoItExitRegister("MyExitFunc")
#include "lib_msg.au3"
#include "lib_time.au3"
#include "lib_ie.au3"
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
Global $nLine = 0

Global $rf_test_env = "test1"

Global $_abortGet = 0 ; flag to abort getting messages from IE

Global $gDebugFile = "_debug.txt"

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

EndFunc ;==> GUI_Create


;===============================================================================
; Function Name:    GetCurInterval(bool $inMin)
;
; Return: if $inMin=
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
; Function Name:    _MOUSEWHEEL()
;===============================================================================
Func _MOUSEWHEEL($hWnd, $iMsg, $wParam, $lParam)

	Local $iMPos = MouseGetPos()
	Local $gci = GUIGetCursorInfo($gui)

	If $gci[4] = $idInterval Then
		;; Mouse is over control, do stuff
		$iInterval += _WinAPI_HiWord($wParam) > 0? 1:-1

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
; Function Name:    Clear_Button_pressed()
;===============================================================================

Func Clear_Button_pressed()

	GUICtrlSetData($idLabel,"" )
	GUICtrlSetData($idEdit, "" )

EndFunc ;==> Clear_Button_pressed

;===============================================================================
; Function Name:    New_Button_pressed()
;===============================================================================
;---
;	no_ie (open login) =>
;	login (submit) =>
;	index (logging) =>
;	loglist (submit) =>
;	logsearch (submit)
;	denied (goto login) => call me again!
;
Func New_Button_pressed()

	Local $_renew=0
	Local $url

	; clear info
	LogError("")

	; === clear from invisible objects
	Local $nKilled = _ie_quitAll(false)
	if $nKilled > 0 then
		LogScreen($nKilled & " hidden instances killed" )
	EndIf

; === active IE
	if not _ie_getActiveTab() then
		; start IE and go to login!
		_ie_new( "https://rfadmin." & $rf_test_env & ".reseptformidleren.net/RFAdmin/login.rfa" )
		DbgFile( "new: " & _ie_getURL() )
	EndIf

	; === login page
	$url = _ie_getURL()
	if StringInStr( $url, "login.rfa" ) > 0 then
		DbgFile( "login page" )
		$_renew = 1
		if not _ie_submitLoginForm() then
			return
		endif
		_ie_goto( StringRegExpReplace( $url, "/RFAdmin/.*", "/RFAdmin/loglist.rfa" ) )
	EndIf

	; === index page (after manual login )
	$url = _ie_getURL()
	if StringInStr( $url, "RFAdmin/index.rfa" ) > 0 then
		DbgFile( "index page" )
		_ie_goto( StringRegExpReplace( $url, "/RFAdmin/.*", "/RFAdmin/loglist.rfa" ) )
	EndIf

	; === loglist/search page - main page
	$url = _ie_getURL()
	if  StringInStr( $url, "RFAdmin/loglist.rfa" ) >0 OR StringInStr( $url, "RFAdmin/logsearch.rfa" ) >0 then
		DbgFile( "loglist/search page" )
		if $_renew then
			DbgFile( "restore search" )
			_ie_restoreSearchFields()
			$_renew = 0
		else
			DbgFile( "save search" )
			_ie_saveSearchFields()
		EndIf
		_ie_submitSearchForm()
	EndIf

		; === access denied after timeout - exception!
	$url = _ie_getURL()
	if StringInStr( $url, "RFAdmin/accessDenied" ) > 0 then
		DbgFile( "access denied page" )
		_ie_goto( StringRegExpReplace( $url, "/RFAdmin/.*", "/RFAdmin/loglist.rfa" ) )
		New_Button_pressed()
	EndIf

	if StringInStr( $url, "RFAdmin/loglist.rfa" ) >0 OR StringInStr( $url, "RFAdmin/logsearch.rfa" ) >0 then
		; === anothe page
		Return
	else
		LogError("*** Error: Not RF admin page " & @CRLF & $url )
	EndIf

EndFunc ;==> New_Button_pressed

;===============================================================================
; Function Name:    Get_Button_pressed()
;===============================================================================
Func	Get_Button_pressed()

	; clear status
	LogError("")

DbgFile( "start " & _Now()  )

_IELoadWaitTimeout( 3000 )

;Get_line_from_link( "https://rfadmin.test2.reseptformidleren.net/RFAdmin/loggeview.rfa?loggeId=c7d0d0b6-4014-4ef0-be60-d9522d39045a&filename=/nfstest2/sharedFiles/log/2021/175/21/13/c7d0d0b6-4014-4ef0-be60-d9522d39045a" )
;return

	if not _ie_getActiveTab() then
		LogError("*** Error: No active RF Admin" & @CRLF )
		return
	EndIf

	Local $url = _ie_getURL()
	if  StringInStr( $url, "loglist.rfa" ) = 0 AND StringInStr( $url, "logsearch.rfa" ) = 0 then
		LogError("*** Error: No messages found on page" & @CRLF )
		return
	EndIf

	; save config in case manuel endring before get
	_ie_saveSearchFields() ; testenv, aktor, msgtype

	Local $aTableData = _ie_getMsgArray()
	if not IsArray($aTableData) then
		LogError( "No messages found" )
		return 0
	endif

	; Check that it is rigth table
	if $aTableData[0][0] <> "Linker" then
		LogError("*** Error: No message table found on " & @CRLF & $url )
		return 0
	EndIf
;_ArrayDisplay($aTableData)

	; store environment
	$rf_test_env = StringRegExpReplace( $url, ".*?rfadmin\.(.*?)\.reseptformidleren.net.*", "$1")

	Local $txt ;, $sTextFromTable = ""
	Local $nMsgCount = UBound($aTableData, $UBOUND_ROWS )-1

	LogError("Found messages: " & $nMsgCount )

	;
	; Get links to  message
	;
	Local $oLinks = _ie_getLinks()
	if not IsObj($oLinks) then
		LogError( "No links found" )
		return 0
	endif



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
	LogError("Got links" )

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
			Local $html = _ie_getPageInNewWindow( $aTableData[$i][0] )
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
			$sParam = StringRegExpReplace( $sParam, "(\S{9}\s+\S{9}) ", "" )

			_save_xml( $html, $sParam )


	Next ; $i

	Opt("GUIOnEventMode", 0)
	GUICtrlSetData($idButtonGet, "Get" )
	GUICtrlSetData($idLabel, $nMsgCount-$i & "/" & $nMsgCount )

;GUICtrlSetData($idEdit, $txt & @CRLF, 0)
;GUICtrlSetData($idLabel, $nMsgCount & " messages found")

EndFunc ;==> Get_Button_pressed

;===============================================================================
; Function Name:    LogScreen()
;
; Adds a line to message log on screen
;===============================================================================

Func	LogScreen( $text )

	_GUICtrlEdit_AppendText($idEdit, $text & @CRLF)

EndFunc ;==> LogScreen

;===============================================================================
; Function Name:    LogFile()
;
; Adds a line to message log in log file
;===============================================================================

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


;===============================================================================
; Function Name:    LogError()
;
; Print error message in label
;===============================================================================

Func	LogError( $text )
	GUICtrlSetData($idLabel, $text & @CRLF )
EndFunc


;===============================================================================
; Function Name:    _save_xml()
;
; Saves xml to a file in a specofoc folder.
; if folder is not defined then folder name is a message date
;===============================================================================

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
	$filename = _ER_GetMsgType( $html ) & "_" & StringReplace( StringStripWS($text, 7), " ", "_") & "_" & StringLeft( _ER_GetMsgId( $html ), 9)  & ".xml"


	if not FileExists( $folder ) then
		DirCreate( $folder )
	endif

	FileDelete( $folder & "\" &  $fileName )

	if FileWrite( $folder & "\" &  $fileName, $html ) = 0 then
		LogScreen("error write file " & $fileName)
		return 3
	EndIf

	If FileSetTime( $folder & "\" &  $fileName, $fileTime, 0) = 0 then
		LogScreen("error filesettime '" & $fileTime & "'->" & $fileName)
		return 4
	EndIf

	Return 0
EndFunc

;===============================================================================
; Function Name:    DbgFile()
;
; Print debug info to debug file
;===============================================================================

func DbgFileClear()
	;FileDelete( $gDebugFile )
EndFunc

Func	DbgFile( $txt )
	FileWriteLine( $gDebugFile, $txt )
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
