#include-once

#include "array.au3"

; #CHANGES# =====================================================================================================================
; 06/09/21
;	Added:
;	Fixed:
; ===============================================================================================================================

; #FUNCTIONS#===============================================================================================================
;
;	_ER_GetBody( $html)
;	_ER_GetExtraParam( $html )
;
; ===============================================================================================================================

; #INTERNAL_USE_ONLY#============================================================================================================
;
;=== message header
;   _ER_GetMsgId($html)
;	_ER_GetMsgType( $html)
;	_ER_GetMsgTime( $html)
;	_ER_GetRefToParent( $html )
;	_ER_GetRefToConversation( $html )
;	_ER_GetDateOfBirth( $html)
;	_ER_GetFnr( $html)
;	_ER_GetPatient($html, $type=0)
;== Apprec
;	_ER_GetApprec( $html)
;	_ER_GetApprecRef( $html)
;	_ER_GetApprecType( $html)
;	_ER_GetApprecStatus( $html)
;	_ER_GetApprecError( $html)
;=== M1
;	_ER_GetM1($html)
;	_ER_isV24($html)
;	_ER_GetNavnFormStyrke($html)
;	_ER_GetRefKode($html)
;	_ER_GetRefHjemmel($html)
;=== M10
;	_ER_GetM10($html)
;	_ER_GetAnnullering( $html)
;	_ER_GetReseptId( $html)
;	_ER_GetKansellering( $html)
;=== M91-6
;	_ER_GetM91($html)
;	_ER_GetM92($html)
;	_ER_GetM93($html)
;	_ER_GetM94($html)
;	_ER_GetM95($html)
;	_ER_GetM96($html)
;=== M3-M15
;	_ER_GetM3($html)
;	_ER_GetM15($html)
;=== M5
;	_ER_GetM5($html)
;=== M911-M912
;	_ER_GetM911($html)
;	_ER_GetM912($html)
;=== M27
;	_ER_GetM271($html)
;	_ER_GetM272($html)
;=== MV
;	_ER_GetMV($html)
;=== general
;	_ER_GetParam( $html, $regexp )
;
; ===============================================================================================================================

; #Container Globals# ===========================================================================================================
Global $__gb_IEExEvalCheckRecursion = False
; ===============================================================================================================================

;================================================================================================================================
;	Function: _ER_GetBody()
;	Returns: all text within "<...>"
;================================================================================================================================

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

;================================================================================================================================
;	Function: _ER_GetExtraParam()
;	Returns:
;				String  - for known messages returns String based on message internal xml
;				""	- if message is unknown
;================================================================================================================================
Func _ER_GetExtraParam( $html )

	Local $ret = ""
	Local $ref = ""

	;ConversationRef
	$ref = StringLeft( _ER_GetRefToParent( $html ) & '000000000', 9) & " " & StringLeft( _ER_GetRefToConversation( $html )& '000000000', 9)

	Switch _ER_GetMsgType( $html)
		case "ERM1"
			$ret = _ER_GetM1($html)
		case "ERM10"
			$ret = _ER_GetM10($html)
		case "ERM91"
			$ret = _ER_GetM91($html)
		case "ERM92"
			$ret = _ER_GetM92($html)
		case "ERM93"
			$ret = _ER_GetM93($html)
		case "ERM94"
			$ret = _ER_GetM94($html)
		case "ERM95"
			$ret = _ER_GetM95($html)
		case "ERM96"
			$ret = _ER_GetM96($html)
		case "ERM3"
			$ret = _ER_GetM3($html)
		case "APPREC"
			$ret = _ER_GetApprec($html)
		case "ERM911"
			$ret = _ER_GetM911($html)
		case "ERM912"
			$ret = _ER_GetM912($html)
		case "ERM251"
			$ret = _ER_GetM251($html)
		case "ERM252"
			$ret = _ER_GetM252($html)
		case "ERM253"
			$ret = _ER_GetM253($html)
		case "ERM5"
			$ret = _ER_GetM5($html)
		case "ERM271"
			$ret = _ER_GetM271($html)
		case "ERM272"
			$ret = _ER_GetM272($html)
		case "ERMV"
			$ret = _ER_GetMV($html)


		;case Else
		;	$ret = $msgType & "_" & _ER_GetMsgId($html)
	EndSwitch

	Return	$ref & $ret ; return file name without Time Type MsgId

EndFunc

;================================================================================================================================
;	HEADER functions
;================================================================================================================================

Func	_ER_GetMsgId( $html)

	Local $a = StringRegExp( $html, 'MsgId>([0-9a-fA-F\-]+?)</', 1)
	if @error then
		; for apprec it is different format
		;<Id>0d1fb989-7bfa-4e2d-9d05-5c64cbd4f390</Id>
		$a = StringRegExp( $html, '(?s)GenDate>.*?Id>(.*?)<', 1)
		if @error then return 0
	endif
	return $a[0]

EndFunc

Func	_ER_GetMsgType( $html)

	Local $a = StringRegExp( $html, '(?s)MsgInfo>.*?<.*?Type.*?V="(.*?)"', 1)
	if @error then
		; for apprec it is different format
		$a = StringRegExp( $html, '(?s)MsgType.*?V="(.*?)"', 1)
		if @error then return 0
	EndIf
	return $a[0]

EndFunc

Func	_ER_GetMsgTime( $html)

	Local $a
	$a = StringRegExp( $html, 'GenDate>(\d+).(\d+).(\d+).(\d+):(\d+):(\d+).*?</', 1)
	if @error then return 0
	Return $a[0] & $a[1] &  $a[2] & $a[3] & $a[4] & $a[5]

EndFunc

Func	_ER_GetRefToParent( $html )

	return _ER_GetParam( $html, '(?s)ConversationRef>.*?RefToParent>(.*?)<.*?RefToParent>' )

EndFunc

Func	_ER_GetRefToConversation( $html )

	return _ER_GetParam( $html, '(?s)ConversationRef>.*?RefToConversation>(.*?)<.*?RefToConversation>' )

EndFunc

Func	_ER_GetDateOfBirth( $html)

	Local $a, $fnr
	$a = StringRegExp(  $html, '(?s)Patient>.*?DateOfBirth>(.*?)<.*?DateOfBirth>', 1)
	if @error then Return 0

	$fnr = $a[0] ; StringRegExp( $a[0], '(?s)Id>(.*)<.*?Id>', 1)
	return $fnr

EndFunc ;-> _ER_GetDateOfBirth

Func	_ER_GetFnr( $html)

	Local $a, $fnr
	$a = StringRegExp( $html, '(?s)Patient>.*?Id>([0-9]{11})<.*?Id>', 1)
	if @error then return 0
	$fnr = $a[0] ; StringRegExp( $a[0], '(?s)Id>(.*)<.*?Id>', 1)
	return $fnr

EndFunc ;-> _ER_GetFnr
 #cs
 <h:Patient>
      <h:FamilyName>Sortland</h:FamilyName>
      <h:GivenName>Herman</h:GivenName>
      <h:Sex DN="Kvinne" V="2" />
      <h:Ident>
        <h:Id>22109345931</h:Id>
        <h:TypeId DN="Fødselsnummer" V="FNR" S="2.16.578.1.12.4.1.1.8116" OT="Norsk fødselsnummer" />
      </h:Ident>
    </h:Patient>
#ce
Func	_ER_GetPatient($html, $type=1)

	Local $a, $name
	$a = StringRegExp( $html, '(?s)Patient>.*?GivenName>(.*?)<', 1)
	if @error then return 0
	$name = $a[0]

	if $type > 0 then
		$a = StringRegExp( $html, '(?s)Patient>.*?FamilyName>(.*?)<', 1)
		if @error then return 0
		$name &= " " & $a[0]
	EndIf

	return $name

EndFunc ;-> _ER_Patient


;================================================================================================================================
;	APPREC functions
;================================================================================================================================

Func	_ER_GetApprec( $html)

	Local $text = ""

	if _ER_GetApprecType( $html) then $text = " " & _ER_GetApprecType( $html)
	if _ER_GetApprecStatus($html) then $text &= " " & _ER_GetApprecStatus($html)
	if _ER_GetApprecError($html) <> 0 then $text &= " " & _ER_GetApprecError($html)
	if _ER_GetApprecRef($html) then $text &= " " & StringLeft( _ER_GetApprecRef($html), 9)
	return $text

EndFunc
#cs
<OriginalMsgId>
    <MsgType V="ERM921" DN="M9_21" />
    <IssueDate>2021-06-26T15:00:00.108+02:00</IssueDate>
    <Id>0920917c-ae57-43f9-a9e9-db309b302b47</Id>
  </OriginalMsgId>

#ce
Func	_ER_GetApprecId( $html)
	return _ER_GetParam( $html, '(?s)GenDate>.*?Id>(.*?)<' )
EndFunc

Func	_ER_GetApprecRef( $html)
	return _ER_GetParam( $html, '(?s)OriginalMsgId>.*?Id>(.*?)<' )
EndFunc

Func	_ER_GetApprecType( $html)
	return _ER_GetParam( $html, '(?s)OriginalMsgId>.*?V="(.*?)"' )
EndFunc


Func	_ER_GetApprecStatus( $html)
	return _ER_GetParam( $html, '(?s)Status.*?DN="(.*?)".*?>' )
	;<Status V="2" DN="Avvist" />
EndFunc

Func	_ER_GetApprecError( $html)
	return _ER_GetParam( $html, '(?s)Error.*?V="(.*?)".*?>' )
	; <Error V="360"
EndFunc


;================================================================================================================================
;	M1 functions
;================================================================================================================================

Func _ER_GetM1($html)

Local $text = ""

	if _ER_isV24($html) then $text = " " & _ER_isV24($html)
	if _ER_GetRefHjemmel($html) then $text &= " " & _ER_GetRefHjemmel($html)
	if _ER_isMagistrell($html) then
		$text &= " Magistrell " & StringLeft(_ER_GetMagistrellNavn($html), 40)
		$text &=  " " & _ER_GetTypeLegemiddel($html)
	ElseIf _ER_isHandelsvare($html) then
		$text &= " " & _ER_GetHandelsvareProdGruppe($html)
	else
		if _ER_GetNavnFormStyrke($html) then
			$text &= " " & _ER_GetNavnFormStyrke($html)
			$text &=  " " & _ER_GetTypeLegemiddel($html)
		EndIf
	EndIf


	if _ER_GetPatient($html) then $text &= " " & _ER_GetPatient($html)
	if _ER_GetFnr($html) then $text &= " " & _ER_GetFnr($html)
	if _ER_GetDateOfBirth($html) then $text &= " " & _ER_GetDateOfBirth($html)
	if _ER_GetParam( $html, '(?s)RefNr>(.*?)<') then $text &= " RefNr_" & _ER_GetParam( $html, '(?s)RefNr>(.*?)<' )


;<BytteresRekvirent V="1" DN="Ja" />
	$text &= _ER_GetParam( $html, '(?s)BytteresRekvirent.*?V="(.*?)" ') = 1? " BytteresRekvirent":""

	Return	$text

EndFunc

Func	_ER_isV24($html)
	if StringInStr( $html, 'xmlns="http://www.kith.no/xmlstds/eresept/m1/2010-05-01"' ) then return "v24"
	Return 0
EndFunc

Func	_ER_isMagistrell($html)
	return _ER_GetParam( $html, '(?s)Legemiddelblanding>')
EndFunc

Func _ER_GetMagistrellNavn($html)
	Return _ER_GetParam( $html, '(?s)Navn>(.*?)<' );
EndFunc

Func	_ER_isHandelsvare($html)
	return _ER_GetParam( $html, '(?s)ReseptDokHandelsvare')
EndFunc

Func _ER_GetHandelsvareNavn($html)
	Return _ER_GetParam( $html, '(?s)Navn>(.*?)<' );
EndFunc

;<ProdGruppe DN="Belter til kompresjon" S="2.16.578.1.12.4.1.1.7403" V="5050901"/>
Func _ER_GetHandelsvareProdGruppe($html)

	Local $a = StringRegExp( $html, '(?s)ProdGruppe.*?V="(.*?)"', 1)
	if @error then return 0
	return $a[0]

EndFunc


Func	_ER_GetNavnFormStyrke($html)

	Local $a = StringRegExp( $html, 'NavnFormStyrke>(.*?)</', 1)
	if @error then return 0

	; strip "/"
	if StringInStr( $a[0], "/" ) then
		$a[0] = StringLeft( $a[0], StringInStr( $a[0], "/" )-1 )
	EndIf

	return $a[0]

EndFunc

Func	_ER_GetRefKode($html)

	Local $a = StringRegExp( $html, 'RefKode.*?V="(.*?)"', 1)
	if @error then return 0
	return $a[0]

EndFunc

Func	_ER_GetRefHjemmel($html)

	Local $a = StringRegExp( $html, 'RefHjemmel.*?V="(.*?)"', 1)
	if @error then return 0
	Switch $a[0]
		Case 200
			return "$2"
		Case 300
			return "$3"
		Case 400
			return "$4"
		Case 950
			return "$H"
		Case 800
			return "$Y"
		Case 301
			return "$3a"
		Case 302
			return "$3b"

	EndSwitch

	return "$" & $a[0]

EndFunc

;================================================================================================================================
;	M10 functions
;================================================================================================================================

Func _ER_GetM10($html)

	Local $text = ""

	if _ER_GetKansellering( $html) then $text &= " " & _ER_GetKansellering($html)
	if _ER_GetAnnullering( $html) then $text &= " " & _ER_GetAnnullering($html)
	if _ER_GetRefHjemmel($html) then $text &= " " & _ER_GetRefHjemmel($html)
	if _ER_isMagistrell($html) then
		$text &= " Magistrell " & StringLeft(_ER_GetMagistrellNavn($html), 40)
	ElseIf _ER_isHandelsvare($html) then
		$text &= " " & StringLeft(_ER_GetHandelsvareNavn($html), 20)
	else
		if _ER_GetNavnFormStyrke($html) then $text &= " " & _ER_GetNavnFormStyrke($html)
	EndIf
	if _ER_GetPatient($html) then $text &= " " & _ER_GetPatient($html)
	if _ER_GetFnr($html) then $text &= " " & _ER_GetFnr($html)
	if _ER_GetDateOfBirth($html) then $text &= " " & _ER_GetDateOfBirth($html)
	if _ER_GetReseptId( $html) then $text &= " " & _ER_GetReseptId($html)
	$text &= _ER_GetEgenandel( $html )
	$text &= _ER_GetParam( $html, '(?s)Papirresept>true<')? " papir":""
	$text &= _ER_GetParam( $html, '(?s)RekvirentNordisk>true<')? " RekvirentNordisk":""
	$text &= _ER_GetParam( $html, '(?s)ByttereservasjonKunde>true<')? " Kundereservasjon":""
	$text &= _ER_GetParam( $html, '(?s)ReservasjonRapportFastlege>true<')? " ReservasjonRapportFastlege":""
	Return	$text

EndFunc

Func _ER_GetAnnullering( $html)

;<Annullering>false</Annullering>

Local $ret = _ER_GetParam( $html, '(?s)Annullering>true<' );
Local $antall = _ER_GetParam( $html, '(?s)Antall>(.*?)<' );

return $ret? "Annullering("&$antall&")": ""

EndFunc

Func _ER_GetReseptId( $html)
;<Utlevering xmlns="http://www.kith.no/xmlstds/eresept/utlevering/2013-10-08">
;<ReseptId>
	return StringLeft(_ER_GetParam( $html, '(?s)ReseptId>(.*?)<' ),9)
EndFunc

Func _ER_GetKansellering( $html)

;<Kanselleringskode V="1" DN="Ikke ønsket vare"/>
	Local $ret = _ER_GetParam( $html, '(?s)Kanselleringskode.*?DN="(.*?)"' );
	return $ret ? "Kansellering " & $ret: 0
EndFunc

;================================================================================================================================
;	M9.5 functions
;================================================================================================================================

Func _ER_GetM95($html)
Local $text = ""

	if _ER_GetParam( $html, '(?s)Fnr>(.*?)<' ) then $text &= " " & _ER_GetParam( $html, '(?s)Fnr>(.*?)<' )
	if _ER_GetParam( $html, '(?s)Fornavn>(.*?)<' ) then $text &= " " & _ER_GetParam( $html, '(?s)Fornavn>(.*?)<' )
	;if _ER_GetParam( $html, '(?s)Etternavn>(.*?)<' ) then $text &= " " & _ER_GetParam( $html, '(?s)Etternavn>(.*?)<' )
	if _ER_GetParam( $html, '(?s)Fdato>(.*?)<' ) then $text &= " " & _ER_GetParam( $html, '(?s)Fdato>(.*?)<' )
	if _ER_GetParam( $html, '(?s)RefNr>(.*?)<') then $text &= " RefNr_" & _ER_GetParamX( $html, '(?s)RefNr>(.*?)</RefNr' )
	$text &= " " & _ER_GetParam( $html, '(?s)AlleResepter.*?DN="(.*?)"' )

	Return	$text

EndFunc

;================================================================================================================================
;	M9.6 functions
;================================================================================================================================

Func _ER_GetM96($html)
Local $text = ""

	if _ER_GetParam( $html, 'Listeelement>') then

		$text &= " " & _ER_GetReseptCount( $html )

	Else
		; we did not have resepts <StatusSok V="4" DN="Ingen resept på dette søk" />
		$text &= " " & _ER_GetParam( $html, '(?s)StatusSok.*?DN="(.*?)".*?>' )

	EndIf

	Return	$text

EndFunc

;================================================================================================================================
;	M9.1 functions
;================================================================================================================================

Func _ER_GetM91($html)

Local $text = ""

	if _ER_GetParam( $html, '(?s)Fnr>(.*?)<' ) then $text &= " " & _ER_GetParam( $html, '(?s)Fnr>(.*?)<' )
	if _ER_GetParam( $html, '(?s)Fornavn>(.*?)<' ) then $text &= " " & _ER_GetParam( $html, '(?s)Fornavn>(.*?)<' )
	;if _ER_GetParam( $html, '(?s)Etternavn>(.*?)<' ) then $text &= " " & _ER_GetParam( $html, '(?s)Etternavn>(.*?)<' )
	if _ER_GetParam( $html, '(?s)Fdato>(.*?)<' ) then $text &= " " & _ER_GetParam( $html, '(?s)Fdato>(.*?)<' )
	if _ER_GetParam( $html, '(?s)Arsak DN="(.*?)"' ) then $text &= " " & _ER_GetParam( $html, '(?s)Arsak DN="(.*?)"' )
	if _ER_GetParam( $html, '(?s)RefNr>(.*?)<') then $text &= " RefNr_" & _ER_GetParamX( $html, '(?s)RefNr>(.*?)</RefNr' )
	$text &= " " & _ER_GetParam( $html, '(?s)AlleResepter DN="(.*?)"' )
	$text &= " " & _ER_GetParam( $html, '(?s)InkluderVergeinnsynsreservasjon DN="(.*?)"' )

	Return	$text

EndFunc

;================================================================================================================================
;	M9.2 functions
;================================================================================================================================

Func _ER_GetM92($html)
	Local $text = ""

	; first check if vi got resepter
	if _ER_GetParam( $html, 'Reseptinfo>') then

		$text &= " " & _ER_GetParam( $html, '(?s)Fornavn>(.*?)<' )
		$text &= " " & _ER_GetParam( $html, '(?s)Etternavn>(.*?)<' )
		$text &= " " & _ER_GetReseptCount( $html )

	Else
		; we did not have resepts
		if _ER_GetParam( $html, '(?s)Status.*?DN="(.*?)".*?>' ) then $text &= " " & _ER_GetParam( $html, '(?s)Status.*?DN="(.*?)".*?>' )

	EndIf

	; check if multidose bruker
	if _ER_GetParam( $html, '(?s)Multidosebruker>' ) then
		$text &= " multidosebruker"

;~ 		$text &= " " & _ER_GetParam( $html, '(?s)Multidosebruker>.*?Fnr>(.*?)<' )
;~ 		if _ER_GetParam( $html, '(?s)Multidoselege.*?Navn>(.*?)<' ) then $text &= " " & _ER_GetParam( $html, '(?s)Multidoselege.*?Navn>(.*?)<' )
;~ 		if _ER_GetParam( $html, '(?s)Multidoseapotek.*?Navn>(.*?)<' ) then $text &= " " & _ER_GetParam( $html, '(?s)Multidoseapotek.*?Navn>(.*?)<' )
	EndIf

	Return	$text

EndFunc

;================================================================================================================================
;	M9.3 functions
;================================================================================================================================

Func _ER_GetM93($html)
;<Fnr>02048735722</Fnr>
	Local $text = ""
	;<Kansellering DN="Forespurt resept finnes ikke i RF" V="7" />
	if _ER_GetParam( $html, '(?s)Kansellering.*?DN="(.*?)"' ) then $text = " Kansellering " & _ER_GetParam( $html, '(?s)Kansellering.*?DN="(.*?)"' )
	if _ER_GetParam( $html, '(?s)M93.*?>.*?ReseptId>(.*?-)' ) then $text = " " & _ER_GetParam( $html, '(?s)M93.*?>.*?ReseptId>(.*?-)' )
	if _ER_GetParam( $html, '(?s)RefNr>(.*?)<') then $text &= " RefNr_" & _ER_GetParam( $html, '(?s)RefNr>(.*?)<' )

	Return	$text

EndFunc

;================================================================================================================================
;	M9.4 functions
;================================================================================================================================

Func _ER_GetM94($html)
	Local $text = ""

	if _ER_GetParam( $html, '(?s)Status.*?DN="(.*?)".*?>' ) then $text &= " " & _ER_GetParam( $html, '(?s)Status.*?DN="(.*?)".*?>' )
	if _ER_GetParam( $html, '(?s)StatusSoknadSlv.*?DN="(.*?)".*?>' ) then $text &= " " & _ER_GetParam( $html, '(?s)StatusSoknadSlv.*?DN="(.*?)".*?>' )

	; decode b64 and save M1 as xml
	Local $m1 = _ER_GetM1b64( $html)
	Local $param =  _ER_GetM1( $m1 )
	; save M1 as xml in current folder
	_save_xml( $m1, $param, "ERM1")

	;(OR)
	; save to same folder as m94
	; _save_xml( $m1, $param, StringRegExpReplace( _ER_MsgGetTime( $html), "(\d\d\d\d)(\d\d)(\d\d).*", "$1-$2-$3" ) )

	; add to text
	$text &= $param ;$param alleredy has leading space
	$text &= _ER_GetEgenandel( $html )


	Return	$text

EndFunc

Func	_ER_GetM1b64( $html)

	Local $b64
	; get base64 - b64 can be more than 32K, therefor use RegExReplace, as RegExp can not handle long patterns
	; strip all before Base64
	$b64 = StringRegExpReplace( $html, "(?s).*?Base64Container(.*)", "$1",1 )
	; strip all after base64 and return only inside >...<
	$b64 = StringRegExpReplace( $b64, "(?s).*?>(.*?)</.*", "$1",1 )

	return _Base64Decode( $b64 )

EndFunc

Func	_ER_GetEgenandel( $html )

	Local $a

	$a = StringRegExp( $html, '(?s)BetaltEgenandel.*?V="(.*?)"', 3)
	if @error then return ""

	Return " ("&_ArrayToString( $a, " " )&")"

EndFunc

;================================================================================================================================
;	M3,M14,M15 functions
;================================================================================================================================
Func _ER_GetM3($html)
	Return " " & _ER_GetReseptId($html)
EndFunc

;================================================================================================================================
;	M5 functions
;================================================================================================================================
Func _ER_GetM5($html)

	Local $ret = ""
	;$ret &= " " & _ER_GetPatient( $html )
	;$ret &= " " & _ER_GetFnr( $html )

	$ret &= " " & _ER_GetParam( $html, '(?s)Arsak.*?DN="(.*?)"' );
	$ret &= " " & _ER_GetReseptId($html)
	if _ER_GetParam( $html, '(?s)NyReseptId>(.*?)<' ) then $ret &= " " & StringLeft(_ER_GetParam( $html, '(?s)NyReseptId>(.*?)<' ),9)
	return $ret
EndFunc


;================================================================================================================================
;	M9.11 functions
;================================================================================================================================
Func _ER_GetM911($html)

	Local $text = ""

	if _ER_GetParam( $html, '(?s)M911.*?Fnr>(.*?)<' ) then $text &= " " & _ER_GetParam( $html, '(?s)M911.*?Fnr>(.*?)<' )

	Return	$text

EndFunc


;================================================================================================================================
;	M9.12 functions
;================================================================================================================================
Func _ER_GetM912($html)

	Local $text = ""

	if _ER_GetParam( $html, '(?s)asient>.*?Fnr>(.*?)<' ) then $text &= " " & _ER_GetParam( $html, '(?s)asient>.*?Fnr>(.*?)<' ) & " "
	if _ER_GetParam( $html, '(?s)Multidoselege.*?Navn>(.*?)<' ) then $text &= "L" ;& _ER_GetParam( $html, '(?s)Multidoselege.*?Navn>(.*?)<' )
	if _ER_GetParam( $html, '(?s)Multidoseapotek.*?Navn>(.*?)<' ) then $text &= "A" ;& _ER_GetParam( $html, '(?s)Multidoseapotek.*?Navn>(.*?)<' )
	$text &= " " & _ER_GetReseptCount( $html )

	; if M25 exists
	; get all <VarerIBrukB64> and decode M25
	if StringInStr( $html, "VarerIBrukB64" ) then
		$text &= " " & _ER_GetM25b64( $html)
	endif

Return	$text

EndFunc

;
; there can be several base64
;
Func	_ER_GetM25b64( $html)

	Local $b64
	Local $ret = ""

	; get base64 - b64 can be more than 32K, therefor use RegExReplace, as RegExp can not handle long patterns
	; strip all before and after base 64 and return only inside >...<
	$b64 = StringRegExpReplace( $html, "(?s).*?VarerIBrukB64(.*VarerIBrukB64.*?>).*", "$1",1 )

	Local $b64array = StringRegExp( $b64, "(?s).*?>(.*?)</.*?>", 3 )
	if @error=0 then
		for $m in $b64array
			Local $xml = _Base64Decode( $m )
			$ret &= " " & StringRight(_ER_GetMsgType( $xml ),4)& "(" & _ER_GetReseptCountM252( $xml ) & ")"
		Next
	endif

	return $ret

EndFunc

;================================================================================================================================
;	Get resept counnt in M9.12 and presents count with type OID=7408
;	Returns:
;		f.eks. E12-U2-R1-T2-F1-X1
;================================================================================================================================

Func	_ER_GetReseptCount( $html )

	Local $a
	Local $CountTypes[6]
	Local $ret=""

	Local $ReseptType = "EURTF" ;Volven 7408 = https://volven.no/produkt.asp?id=469436&catID=3&subID=8

	; get all resepter
	$a = StringRegExp( $html, 'Status .*?V="(.)"', 3)
	if @error then return ""


	; count all types
	for $r in $a
		$CountTypes[ StringInStr( $ReseptType, $r) ] += 1
	Next

	for $i=1 to StringLen($ReseptType)
		if $CountTypes[$i] > 0 then $ret &= StringMid( $ReseptType, $i, 1)  & $CountTypes[$i] & "-"
	Next

	; if we got unknown type - show as X
	if $CountTypes[0] > 0 then $ret &= "X" & $CountTypes[0]

	; remove unnecessary - at the end
	if StringRight( $ret, 1) = "-" then $ret = StringTrimRight( $ret, 1)

	Return  $ret

EndFunc

;================================================================================================================================
;	M27 functions
;================================================================================================================================
Func _ER_GetM271($html)
	Local $text = ""

	if _ER_GetPatient($html) then $text &= " " & _ER_GetPatient($html)
	if _ER_GetFnr($html) then $text &= " " & _ER_GetFnr($html)

	$text &= " " &  _ER_GetParam( $html, '(?s)M271 .*?Status.*?DN="(.*?)"' )

	Return	$text

EndFunc

Func _ER_GetM272($html)
	Local $text = ""

	$text &= " " &  _ER_GetParam( $html, '(?s)M272 .*?RegistrertEndring>(.*?)<' )
;          <RegistrertEndring>true</RegistrertEndring>

	Return	$text

EndFunc

;================================================================================================================================
;	M25 functions
;================================================================================================================================
Func _ER_GetM251($html)

	Local $text = ""

	if _ER_GetPatient($html) then $text &= " " & _ER_GetPatient($html)
	if _ER_GetFnr($html) then $text &= " " & _ER_GetFnr($html)

	$text &= " " & _ER_GetReseptCountM252( $html )

	Return	$text

EndFunc

Func _ER_GetM252($html)

	Local $text = ""

	if _ER_GetPatient($html) then $text &= " " & _ER_GetPatient($html)
	if _ER_GetFnr($html) then $text &= " " & _ER_GetFnr($html)

	$text &= " " & _ER_GetReseptCountM252( $html )

	Return	$text

EndFunc

Func _ER_GetReseptCountM252( $html )

	Local $a
	Local $CountTypes[6]
	Local $ret=""

	; first resept type
	Local $ReseptType = "EPU"
	;<Type DN="Eresept" V="E" />  Volven 7491 = Type resept https://volven.no/produkt.asp?id=469436&catID=3&subID=8

	; get all resepter
	$a = StringRegExp( $html, '(?s)EnkeltoppforingLIB>.*?Type .*?V="(.)"', 3)
	if @error then return ""

	; count all types
	for $r in $a
		$CountTypes[ StringInStr( $ReseptType, $r) ] += 1
	Next

	for $i=1 to 3
		if $CountTypes[$i] > 0 then $ret &= StringMid( $ReseptType, $i, 1)  & $CountTypes[$i] & "-"
	Next

	; if we got unknown type - show as X
	if $CountTypes[0] > 0 then $ret &= "X" & $CountTypes[0]

	; remove unnecessary - at the end
	if StringRight( $ret, 1) = "-" then $ret = StringTrimRight( $ret, 1)

	Return  $ret

EndFunc


Func _ER_GetM253($html)

	Local $text = ""

	if _ER_GetPatient($html) then $text &= " " & _ER_GetPatient($html)
	if _ER_GetFnr($html) then $text &= " " & _ER_GetFnr($html)

	$text &= " " & _ER_GetReseptCountM252( $html )
	$text &= " " & _ER_GetMultidoseCountM253( $html )

	Return	$text

EndFunc

;====================================
; Decode base64 in M912 and M25 messages within tags
;	- VarerIBrukB64
;	- ReseptDokLegemiddelB64
;====================================

Func	_decodeB64( $html)

		; Get first VarerIBrukB64
		Local $tag = "(?:VarerIBrukB64|ReseptDokLegemiddelB64)"
		Local $cnt = 0
while 1
		Local $b64code = StringRegExpReplace( $html, "(?s).*?<[^/]*?"&$tag&"[^>]*?>([a-zA-Z\d/=+]+?)</.*", "$1", 1 )
		if @error then
			;ConsoleWrite( "#Error " & @extended & @CRLF )
			return SetError( 1, @extended, $html)
		Else
			if @extended = 0 then return SetError( 0, $cnt, $html) ; no more pattern
			$cnt += @extended
			Local $m = _Base64Decode( $b64code )

			;ConsoleWrite( $tag & " " & StringLeft( $b64code, 50) & " " & StringLeft( $m, 50) & @CRLF )

			$html = StringReplace( $html, $b64code, @CRLF & $m & @CRLF)
		EndIf
Wend
		;return $html
EndFunc



;================================================================================================================================
;	Get resept counnt in M25.3 which included in multidose
;
;		f.eks.(8-1)
;================================================================================================================================

Func _ER_GetMultidoseCountM253( $html )

	Local $a
	Local $CountTypes[3]
	Local $ret=""

	; first resept type
	Local $ReseptType = "12"
	;<InngarMultidose DN="Nei" V="2" />  Volven 1101 = Ja, Nei InngarMultidose

	; get all resepter
	$a = StringRegExp( $html, '(?s)EnkeltoppforingLIB>.*?InngarMultidose .*?V="(.)"', 3)
	if @error then return ""

	; count all types
	for $r in $a
		$CountTypes[ StringInStr( $ReseptType, $r) ] += 1
	Next

	; Add inngarMultidose
	$ret &= " ("
	if $CountTypes[1] > 0 then $ret &= $CountTypes[1]
	if $CountTypes[2] > 0 then $ret &= "-" & $CountTypes[2]
	if $CountTypes[0] > 0 then $ret &= "x" & $CountTypes[0]
	$ret &= ")"

	Return  $ret

EndFunc


;================================================================================================================================
;	Get type of legemiddel in M1
;	Returns:
;		f.eks. for simple reciept
;		(P) - pakning
;		(V) - virkestoff
;		(M) - merkevare
;		for magistrelle
;		(V1-L2-A1) - legemiddelblanding
;================================================================================================================================

Func	_ER_GetTypeLegemiddel( $html )

	Local $a
	Local $CountTypes[6]
	Local $ret=""

	Local $ReseptType = "LXVA" ; added extra symbols for antivirus

	; get bestanddeler for magistrelle
	$a = StringRegExp( $html, "(?i)<[^/]*Bestanddel(Legemiddel|Virkestoff|Annet)>", 3)
	if not @error then

		; count all types
		for $r in $a
			$CountTypes[ StringInStr( $ReseptType, StringLeft( $r, 1) ) ] += 1
		Next

		for $i=1 to StringLen($ReseptType)
			if $CountTypes[$i] > 0 then $ret &= StringMid( $ReseptType, $i, 1)  & $CountTypes[$i] & "-"
		Next

		; remove unnecessary - at the end
		if StringRight( $ret, 1) = "-" then $ret = StringTrimRight( $ret, 1)

	Else

		; get legemiddel type for ordinary receipt - can only be 1
		if _ER_GetParam( $html, "(?i)Legemiddelpakning>" ) then $ret &= "P"
		if _ER_GetParam( $html, "(?i)Legemiddelmerkevare>" ) then $ret &= "M"
		if _ER_GetParam( $html, "(?i)Legemiddelvirkestoff>" ) then $ret &= "V"
	EndIf

	Return  "("&$ret&")"

EndFunc

;================================================================================================================================
;	MV functions
;================================================================================================================================
Func _ER_GetMV($html)

	Local $text = ""
#cs
<SystemInfo>
            <SystemCode>FM</SystemCode>
            <SystemName>Forskrivningsmodul</SystemName>
            <Version>4.9.3.18743</Version>
<OperationSupplierInfo>
            <ServiceVendorName>Boots Norge</ServiceVendorName>
#ce

	; (?:) - non-inclusive group to choose one of three but not return it as a match
	; we get all three in any order as many as number of groups
	;
	$text &= " " & _ER_GetParamX( $html, '(?s)SystemInfo>.*?(?:SystemCode|SystemName|Version)>(.*?)</..*?(?:SystemCode|SystemName|Version)>(.*?)</..*?(?:SystemCode|SystemName|Version)>(.*?)</' )
	;$text &= " " & _ER_GetParamX( $html, '(?s)SystemInfo>.*?SystemName>(.*?)<' )
	;$text &= " " & _ER_GetParamX( $html, '(?s)SystemInfo>.*?Version>(.*?)<' )
	$text &= " (" & _ER_GetParam( $html, '(?s)OperationSupplierInfo>.*?ServiceVendorName>(.*?)<' ) & ")"

	Return	$text
EndFunc
;====================================
; generic internal function
;====================================

Func	_ER_GetParam( $html, $regexp )

	Local $a
	$a = StringRegExp( $html, $regexp, 1)
	if @error then return 0

	Return $a[0]

EndFunc

Func	_ER_GetParamX( $html, $regexp )

	Local $a
	$a = StringRegExp( $html, $regexp, 3)
	if @error then return ""

	Return _ArrayToString( $a, " " )

EndFunc

; #FUNCTION# ;===============================================================================
;
; Name...........: _Base64Decode
; Description ...: Returns the strinng decoded from the provided Base64 string.
; Syntax.........: _Base64Decode($sData)
; Parameters ....: $sData
; Return values .: Success - String decoded from Base64.
;                  Failure - Returns 0 and Sets @Error:
;                  |0 - No error.
;                  |1 - Could not create DOMDocument
;                  |2 - Could not create Element
;                  |3 - No string to return
; Author ........: turbov21
; Modified.......:
; Remarks .......:
; Related .......: _Base64Encode
; Link ..........;
; Example .......; Yes
;
; ;==========================================================================================
Func _Base64Decode($sData)
    Local $oXml = ObjCreate("Msxml2.DOMDocument")
    If Not IsObj($oXml) Then
        SetError(1, 1, 0)
    EndIf

    Local $oElement = $oXml.createElement("b64")
    If Not IsObj($oElement) Then
        SetError(2, 2, 0)
    EndIf

    $oElement.dataType = "bin.base64"
    $oElement.Text = $sData
    Local $sReturn = BinaryToString($oElement.nodeTypedValue, 4)

    If StringLen($sReturn) = 0 Then
        SetError(3, 3, 0)
    EndIf

    Return $sReturn
EndFunc   ;==>_Base64Decode

