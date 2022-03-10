#Region ***** includes
#include-once
#include <GUIConstants.au3>
#include <WinAPI.au3>

#EndRegion ***** includes

#Region Lib files
;#include "lib_msg.au3"
;#include "lib_ie.au3"
;#include "lib_time.au3"
#EndRegion LIB files

#Region *** Global variables

Global $idButtonGet
Global $idButtonClear
Global $idButtonNew

Global $idEdit
Global $idLabel
#EndRegion Global Variables

;OnAutoItExitRegister("MyExitFunc")
Opt('MustDeclareVars', 1)


_gui_main()

;===============================================================================
; Function Name:  Main
;===============================================================================
Func	_gui_main()

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
Local const $guiWidth = 800
Local const $guiHeight = 300
Local const $guiLeft = -1
Local const $guiTop = -1

; GUI height includes windows title (23px)
; GUI elements start
Local const $winTitleHeight = _WinAPI_GetSystemMetrics($SM_CYCAPTION)


; Create input
	GUICreate( "RF logger - v." & GetVersion(), $guiWidth, $guiHeight, $guiLeft, $guiTop, $WS_MINIMIZEBOX+$WS_SIZEBOX ) ; & GetVersion(), 500, 200)

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
; Clears all info in GUI
;===============================================================================
Func _gui_Clear() ;Clear_Button_pressed()

	GUICtrlSetData($idLabel, "" )
	GUICtrlSetData($idEdit, "" )

EndFunc

;===============================================================================
; Print top line
;===============================================================================
Func	_gui_setStatus( $text )
	GUICtrlSetData($idLabel, $text )
EndFunc

;===============================================================================
; Add text to the end of log
;===============================================================================
Func	_gui_addLine( $text )

	; move to the end of text
	Local $cEnd = StringLen( GUICtrlRead($idEdit) )
	GuiCtrlSendMsg($idEdit, $EM_SETSEL, $cEnd, $cEnd )

	; write line to the end
	GUICtrlSetData($idEdit, $text & @CRLF, 0)

	;GUICtrlSetData($idEdit, GUICtrlRead($idEdit) & $text & @CRLF)

EndFunc

;===============================================================================
; return log file
;===============================================================================
Func	_gui_getLog()
	return GUICtrlRead($idEdit)
EndFunc


#Region ***** System testing
;#include "test/unittest.au3"
;Test("user")
;UTAssertEqual( _user_setUser( $user ), 0 )

#EndRegion ***** System testing