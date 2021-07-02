#include-once

#include <Array.au3>
#include "../ut/unittest.au3"




#cs ----------------------------------------------------------------------------

 Micro unit test framework


#ce ----------------------------------------------------------------------------

Global $iPID = 0
Global $sWinTitle = "Create new pasient"
Global $sErrTitle = ""

Func StartApp()
	$iPID = Run( "../nypasient2.exe" )
	If WinWaitActive( $sWinTitle, "", 5 ) = 0 Then
		Return False
	Else
		Return True
	EndIF
EndFunc

Func CloseApp()
	ProcessClose( $iPID )
	If ProcessWaitClose( $iPID, 3 ) = 1 Then
		Return True
	Else
		Return False
	EndIf
EndFunc

Func SendInput( $text )

	if ControlSetText( $sWinTitle, "", "Edit1", $text ) = 1 then
		sleep(100)
		ControlSend( $sWinTitle, "", "Edit1", "{ENTER}" )
		return true
	Else
		return false
	EndIf

EndFunc

Func SelectControlContextMenuItem( $text )

; select and return context menu item
; 	Proper
; 	Upper
;
; Return true if selected

;~ $hMain = _GUICtrlMenu_GetMenu($hWnd)
;~ ControlClick( $sWinTitle, "", "Pasient", "right" )
;~ sleep(1000)
;~ MouseClick( "" )

EndFunc

Func GetErrorBox( $text )
	Local $title = WinGetTitle("[active]")
	$sErrTitle = $title
	return  StringLeft( $title, StringLen($text) ) = $text
EndFunc


func CloseErrorBox( $btn = "Button1" )
	ControlClick( $sErrTitle, "", $btn )
	$sErrTitle = ""
EndFunc


Func AssertFileEqual( $file1, $file2 )
EndFunc

