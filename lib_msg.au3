#include-once

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
;=== M91
;	_ER_GetM10($html)
;=== M93
;	_ER_GetM10($html)
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

	Switch _ER_GetMsgType( $html)
		case "ERM1"
			$ret = _ER_GetM1($html)
		case "ERM10"
			$ret = _ER_GetM10($html)
		case "ERM95"
			$ret = _ER_GetM95($html)
		case "ERM91"
			$ret = _ER_GetM91($html)
		case "ERM93"
			$ret = _ER_GetM93($html)
		case "APPREC"
			$ret = _ER_GetApprec($html)
		case "ERM911"
			$ret = _ER_GetM911($html)
		case "ERM912"
			$ret = _ER_GetM912($html)

		;case Else
		;	$ret = $msgType & "_" & _ER_GetMsgId($html)
	EndSwitch

	Return	$ret ; return file name without Time Type MsgId

EndFunc

;================================================================================================================================
;	HEADER functions
;================================================================================================================================

Func	_ER_GetMsgId( $html)

	Local $a = StringRegExp( $html, 'MsgId>([0-9a-fA-F\-]+?)</', 1)
	if @error then return 0
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
Func	_ER_GetPatient($html, $type=0)

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
	return _ER_GetParam( $html, '(?s)Error.*?V="(.*?)".*?>' ) & ")"
	; <Error V="360"
EndFunc


;================================================================================================================================
;	M1 functions
;================================================================================================================================

Func _ER_GetM1($html)

Local $text = ""

	if _ER_isV24($html) then $text = " " & _ER_isV24($html)
	if _ER_GetRefHjemmel($html) then $text &= " " & _ER_GetRefHjemmel($html)
	if _ER_GetNavnFormStyrke($html) then $text &= " " & _ER_GetNavnFormStyrke($html)
	if _ER_GetPatient($html) then $text &= " " & _ER_GetPatient($html)
	if _ER_GetFnr($html) then $text &= " " & _ER_GetFnr($html)
	if _ER_GetDateOfBirth($html) then $text &= " " & _ER_GetDateOfBirth($html)

	Return	$text

EndFunc

Func	_ER_isV24($html)
	if StringInStr( $html, 'xmlns="http://www.kith.no/xmlstds/eresept/m1/2010-05-01"' ) then return "v24"
	Return 0
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

	return $a[0]

EndFunc

;================================================================================================================================
;	M10 functions
;================================================================================================================================

Func _ER_GetM10($html)

	Local $text = ""

	if _ER_GetKansellering( $html) then $text &= " " & _ER_GetKansellering($html)
	if _ER_GetAnnullering( $html) then $text &= " " & _ER_GetAnnullering($html)
	if _ER_GetRefHjemmel($html) then $text &= " " & _ER_GetRefHjemmel($html)
	if _ER_GetNavnFormStyrke($html) then $text &= " " & _ER_GetNavnFormStyrke($html)
	if _ER_GetPatient($html) then $text &= " " & _ER_GetPatient($html)
	if _ER_GetFnr($html) then $text &= " " & _ER_GetFnr($html)
	if _ER_GetDateOfBirth($html) then $text &= " " & _ER_GetDateOfBirth($html)
	if _ER_GetReseptId( $html) then $text &= " " & _ER_GetReseptId($html)

	Return	$text

EndFunc

Func _ER_GetAnnullering( $html)

;<Annullering>false</Annullering>
return _ER_GetParam( $html, '(?s)Annullering>true<' ) = 0 ? 0: "Annullering"

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
;<Fnr>02048735722</Fnr>
	Return	" " & _ER_GetParam( $html, '(?s)Fnr>(.*?)<' )

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
	if _ER_GetParam( $html, '(?s)RefNr>(.*?)<') then $text &= " RefNr_" & _ER_GetParam( $html, '(?s)RefNr>(.*?)<' )
	$text &= " " & _ER_GetParam( $html, '(?s)AlleResepter DN="(.*?)"' )
	$text &= " " & _ER_GetParam( $html, '(?s)InkluderVergeinnsynsreservasjon DN="(.*?)"' )

	Return	$text

EndFunc
;================================================================================================================================
;	M9.3 functions
;================================================================================================================================

Func _ER_GetM93($html)
;<Fnr>02048735722</Fnr>
	Local $text = ""
	;<Kansellering DN="Forespurt resept finnes ikke i RF" V="7" />
	if _ER_GetParam( $html, '(?s)Kansellering.*?DN="(.*?)"' ) then $text = " " & _ER_GetParam( $html, '(?s)Kansellering.*?DN="(.*?)"' )
	if _ER_GetParam( $html, '(?s)M93.*?>.*?ReseptId>(.*?-)' ) then $text = " " & _ER_GetParam( $html, '(?s)M93.*?>.*?ReseptId>(.*?-)' )

	Return	$text

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

	if _ER_GetParam( $html, '(?s)Multidosepasient>.*?Fnr>(.*?)<' ) then $text &= " " & _ER_GetParam( $html, '(?s)Multidosepasient>.*?Fnr>(.*?)<' )
	if _ER_GetParam( $html, '(?s)Multidoselege.*?Navn>(.*?)<' ) then $text &= " " & _ER_GetParam( $html, '(?s)Multidoselege.*?Navn>(.*?)<' )
	if _ER_GetParam( $html, '(?s)Multidoseapotek.*?Navn>(.*?)<' ) then $text &= " " & _ER_GetParam( $html, '(?s)Multidoseapotek.*?Navn>(.*?)<' )
	$text &= " " & _ER_GetReseptCount( $html )

Return	$text

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

;====================================
; generic internal function
;====================================

Func	_ER_GetParam( $html, $regexp )

	Local $a
	$a = StringRegExp( $html, $regexp, 1)
	if @error then return 0

	Return $a[0]

EndFunc

