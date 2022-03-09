#Region ***** includes
#include-once
#include <GUIConstants.au3>
;#include <GUIConstantsEx.au3>
;#include <staticconstants.au3>
;#include <Crypt.au3>
;#include <EditConstants.au3>
;#include <StringConstants.au3>
;#include <WinAPIConv.au3>
;#include <WindowsConstants.au3>
#EndRegion ***** includes

#Region *** Global variables

Global $_user_user ; user account
Global $_user_pass  ; RF Admin password

#EndRegion Global variables

;_user_passGUI()
;ConsoleWrite( $_user_user & @CRLF )
;ConsoleWrite( $_user_pass & @CRLF )

; function spesific to credentials
; _user_getUser()
; _user_getPass()
; _user_setPass()
; _user_setUser()
; _user_gui() -> cresate gui for pass entering

Func	_user_getUser()
	return $_user_user
EndFunc

Func	_user_setUser( $user)
	$_user_user = $user
EndFunc

Func	_user_getPass()
	return $_user_pass
EndFunc

Func	_user_setPass( $pass )
	$_user_pass = $pass
EndFunc

Func	_user_gui()

$Form1 = GUICreate("Enter login to RF Admin", 283, 132, -1, -1,  BitOR($WS_CAPTION, $WS_POPUP) )
$UserLabel = GUICtrlCreateLabel("User", 9, 10, 50, 18, BitOR($SS_CENTERIMAGE, $SS_RIGHTJUST))
$UserEdit = GUICtrlCreateInput("", 60, 10, 200, 22, $ES_AUTOHSCROLL, BitOR($WS_EX_CLIENTEDGE, $WS_EX_STATICEDGE))

$PassLabel = GUICtrlCreateLabel("Password", 9, 35, 50, 18, BitOR($SS_CENTERIMAGE, $SS_RIGHTJUST))
$PasswordEdit = GUICtrlCreateInput("", 60, 35, 200, 22, BitOR($ES_PASSWORD, $ES_AUTOHSCROLL), BitOR($WS_EX_CLIENTEDGE, $WS_EX_STATICEDGE))

;$SaveLabel = GUICtrlCreateLabel("Save", 9, 60 ) ;, 50, 18, BitOR($SS_CENTERIMAGE, $SS_RIGHTJUST))
;$HashLabel = GUICtrlCreateLabel("", 60, 60, 120,18 ) ;, BitOR($SS_CENTERIMAGE, $SS_RIGHTJUST))

$SaveCheckbox = GUICtrlCreateCheckbox("Save", 220, 80 ) ;, 100, 22, BitOR($ES_PASSWORD, $ES_AUTOHSCROLL), BitOR($WS_EX_CLIENTEDGE, $WS_EX_STATICEDGE))
GUICtrlSetState($SaveCheckbox, $GUI_CHECKED)

$ButtonOk = GUICtrlCreateButton("&OK", 93, 89, 80, 27, 0)
GUISetState(@SW_SHOW)

Local $pass, $user

While 1
    $nMsg = GUIGetMsg()
    Switch $nMsg
        Case $GUI_EVENT_CLOSE
            return 0
        Case $ButtonOk
            $pass = GUICtrlRead($PasswordEdit)
			if $pass > "" and $user > "" then ExitLoop
		Case $UserEdit
            $user = GUICtrlRead($UserEdit)
			GUICtrlSetState($PasswordEdit, $GUI_FOCUS)
		Case $PasswordEdit
            $pass = GUICtrlRead($PasswordEdit)
			if $pass > "" and $user > "" then ExitLoop

		Case $SaveCheckbox
			;GUICtrlSetState($ButtonOk, $GUI_FOCUS)
			;ExitLoop
	EndSwitch

WEnd

_user_setUser($user)
_user_setPass($pass)

Return 1

EndFunc

#Region ***** Unit testing
#include "test/unittest.au3"
Local $user = "myUser@dom.com"
Local $pass ="myPass#$.1234"

Test("user")
UTAssertEqual( _user_setUser( $user ), 0 )
UTAssertEqual( _user_getUser(), $user )
UTAssertEqual( _user_setPass( $pass ), 0 )
UTAssertEqual( _user_getPass(), $pass )

#EndRegion ***** Unit testing