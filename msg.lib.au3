#Region ***** includes
#include-once
#include "mheader.au3"
#include "m10.au3"
#EndRegion ***** includes

#Region ***** functions
; ===============================================================================================================================
;	_get_Message( $html )
;
;	_get_MsgId( $html )
;	_get_MsgType( $html )
;	_get_MsgTime( $html )
;	_get_RefToParent( $html )
;	_get_RefToConversation( $html )
;
; ===============================================================================================================================

#EndRegion ***** functions

#Region *** Global variables

#EndRegion Global Variables

;================================================================================================================================
;	Function: _ER_GetMessageParam( $html )
;	Returns:
;				String  - for known messages returns String based on message internal xml
; "ERMXX xxxxxxxxx xxxxxxxxx"
;				""	- if message is unknown
;================================================================================================================================
Func _get_Message( $html )

	Local $ref
	Local $text
	Local $msg = _get_MsgType($html) ; ERMXXX

	;ConversationRef
	$ref = StringLeft( _get_RefToParent( $html ) & '000000000', 9)
	$ref &= " " & StringLeft( _get_RefToConversation( $html )& '000000000', 9)


	; call corresponding message parser _ER_GetMXXX with param $hhml
	$text = Call( "_get_" & $msg, $html )
	if @error=0xDEAD then ;function does not exist
		$text = ""
	EndIf

	Return	$msg & " " & $ref & " " & $text

EndFunc



#Region ***** Unit testing
#include "test/unittest.au3"

Test("_get_Message")
UTAssertEqual( _get_Message( UTFileRead( 'M10 annullering.xml') ), "ERM10 ad8d8ceb- b7f8db46- ")

#EndRegion
