#Region ***** includes
#include-once
#EndRegion ***** includes

#Region ***** functions
; ===============================================================================================================================
;	_get_param( $html, $regexp )
;	_get_paramList( $html, $regexp )
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

;====================================
; generic internal function
;====================================

Func	_get_param( $html, $regexp )

	Local $a
	$a = StringRegExp( $html, $regexp, 1)
	if @error then return ""

	Return $a[0]

EndFunc

Func	_get_paramList( $html, $regexp )

	Local $a
	$a = StringRegExp( $html, $regexp, 3)
	if @error then return ""

	Return _ArrayToString( $a, " " )

EndFunc


;================================================================================================================================
;	Common functions for all messages
;================================================================================================================================

Func	_get_MsgId( $html )

	return _get_param( $html, 'MsgId>([0-9a-fA-F\-]+?)</' )

EndFunc

Func	_get_MsgType( $html )

	Local $ret = _get_param( $html, '(?s)MsgInfo>.*?<.*?Type.*?V="(.*?)"' )
	if $ret  = "" then
		return _get_param( $html, '(?s)MsgType.*?V="(.*?)"' )
	EndIf
	return $ret

;~ 	Local $a = StringRegExp( $html, '(?s)MsgInfo>.*?<.*?Type.*?V="(.*?)"', 1)
;~ 	if @error then
;~ 		; for apprec it is different format
;~ 		$a = StringRegExp( $html, '(?s)MsgType.*?V="(.*?)"', 1)
;~ 		if @error then return ""
;~ 	EndIf
;~ 	return $a[0]

EndFunc

Func	_get_MsgTime( $html )

	Local $a
	$a = StringRegExp( $html, 'GenDate>(\d+).(\d+).(\d+).(\d+):(\d+):(\d+).*?</', 1)
	if @error then return 0
	Return $a[0] & $a[1] &  $a[2] & $a[3] & $a[4] & $a[5]

EndFunc

Func	_get_RefToParent( $html )

	return _get_param( $html, '(?s)ConversationRef>.*?RefToParent>(.*?)<.*?RefToParent>' )

EndFunc

Func	_get_RefToConversation( $html )

	return _get_param( $html, '(?s)ConversationRef>.*?RefToConversation>(.*?)<.*?RefToConversation>' )

EndFunc


#Region ***** Unit testing
#include "test/unittest.au3"
Local $xml=UTFileRead("ERM10_papir.xml")

Test("_get_param")
UTAssertEqual( _get_param( $xml, "MIGversion>(.*?)</"), "v1.2 2006-05-24")
UTAssertEqual( _get_param( $xml, "Papirresept>(.*?)<"), "true")

Test("_get_MsgId")
UTAssertEqual( _get_MsgId($xml), "5e11ac3a-c6db-4a95-93ec-0cabb97e7b2b")

Test("_get_MsgType")
UTAssertEqual( _get_MsgType( $xml ), "ERM10")
UTAssertEqual( _get_MsgType( UTFileRead("APPREC_2_49.xml") ), "APPREC")

Test("_get_MsgTime")
UTAssertEqual( _get_MsgTime($xml), "20210612092032")

Test("_get_RefToParent")
UTAssertEqual( _get_RefToParent( $xml), "7330a6db-7c45-4cc9-80fb-4817fa748c11")
UTAssertEqual( _get_RefToConversation( $xml), "fc1e29a5-d5ea-4ef7-b538-ff0ad227bf2a")

Test("_get_paramList")
$xml = '          <ReservasjonRapportFastlege>false</ReservasjonRapportFastlege>' & _
'          <AnsattId>9876543</AnsattId>' & _
'          <Egenandel>' & _
'            <StartEgenandelsperiode>2021-01-29</StartEgenandelsperiode>' & _
'            <BetaltEgenandel V="159.75" U="NOK"/>' & _
'          </Egenandel>' & _
'          <Egenandel>' & _
'            <StartEgenandelsperiode>2021-04-29</StartEgenandelsperiode>' & _
'            <BetaltEgenandel V="0" U="NOK"/>' & _
'          </Egenandel>' & _
'          <Egenandel>' & _
'            <StartEgenandelsperiode>2021-07-29</StartEgenandelsperiode>' & _
'            <BetaltEgenandel V="0" U="NOK"/>' & _
'          </Egenandel>' & _
'          <Egenandel>' & _
'            <StartEgenandelsperiode>2021-10-29</StartEgenandelsperiode>' & _
'            <BetaltEgenandel V="0" U="NOK"/>' & _
'          </Egenandel>' & _
'        </Utleveringsrapport>'
UTAssertEqual( _get_paramList( $xml, '(?s)BetaltEgenandel.*?V="(.*?)"' ), "159.75 0 0 0")
UTAssertEqual( _get_paramList( $xml, 'xxx>(.*?)</text' ), "")


#EndRegion
