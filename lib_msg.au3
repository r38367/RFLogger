#include-once

;------------------------------
; Get e/resept body
;
;------------------------------

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

EndFunc ;-> _ER_GetBody


;------------------------------
; Generic function:
;
; Return first RegEx match from string
;
;------------------------------
Func	_ER_GetMatch( $html, $regexp )

	Local $a
	$a = StringRegExp( $html, $regexp, 1)
	if @error then return 0

	Return $a[0]

EndFunc	;-> _ER_GetParam



;------------------------------
; Get msg extra info depending om msg type -> text line
;------------------------------
Func _ER_GetExtraParam( $html )

	Local $ret = ""

	Switch _ER_GetMsgType( $html)
		case "ERM1"
			$ret = _ER_GetM1($html)
		case "ERM10"
			$ret = _ER_GetM10($html)
		case "APPREC"
			$ret = _ER_GetApprec($html)
		;case Else
		;	$ret = $msgType & "_" & _ER_GetMsgId($html)
	EndSwitch


	Return	$ret ; return file name without Time Type MsgId

EndFunc


;------------------------------
; Msg Header functions
;------------------------------

Func	_ER_GetMsgId( $html)

	Local $a = StringRegExp( $html, 'MsgId>([0-9a-fA-F\-]+?)</', 1)
	if @error then return 0
	return $a[0]

EndFunc

Func	_ER_GetMsgType( $html)

	Local $a = StringRegExp( $html, '(?s)MsgInfo>.*?<.*?Type.*?V="(.*?)"', 1)
	if @error then return 0
	return $a[0]

EndFunc

Func	_ER_GetMsgTime( $html)

	Local $a
	$a = StringRegExp( $html, 'GenDate>(\d+).(\d+).(\d+).(\d+):(\d+):(\d+).*?</', 1)
	if @error then return 0
	Return $a[0] & $a[1] &  $a[2] & $a[3] & $a[4] & $a[5]

EndFunc
;

Func	_ER_GetRefToParent( $html )
	return _ER_GetMatch( $html, '(?s)ConversationRef>.*?RefToParent>(.*?)<.*?RefToParent>' )
EndFunc	;->_ER_GetRefToParent

;
Func	_ER_GetRefToConversation( $html )
	return _ER_GetMatch( $html, '(?s)ConversationRef>.*?RefToConversation>(.*?)<.*?RefToConversation>' )
EndFunc	;->_ER_GetRefToConversation



;------------------------------
; M1 functions
;------------------------------
Func _ER_GetM1($html)

Local $text = ""

	if _ER_isV24($html) then $text = " " & _ER_isV24($html)
	if _ER_GetRefHjemmel($html) then $text &= " " & _ER_GetRefHjemmel($html)
	if _ER_GetNavnFormStyrke($html) then $text &= " " & _ER_GetNavnFormStyrke($html)
	if _ER_GetPatient($html) then $text &= " " & _ER_GetPatient($html)
	if _ER_GetFnr($html) then $text &= " " & _ER_GetFnr($html)
	if _ER_GetDateOfBirth($html) then $text &= " " & _ER_GetDateOfBirth($html)

	Return	StringStripWS( $text, 1)

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
			return "§2"
		Case 300
			return "§3"
		Case 400
			return "§4"
		Case 950
			return "§H"
		Case 800
			return "§Y"
		Case 301
			return "§3a"
		Case 302
			return "§3b"

	EndSwitch

	return $a[0]

EndFunc

;------------------------------
; Patient functions
;------------------------------


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

;------------------------------
; M10 functions
;------------------------------

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

	Return	StringStripWS( $text, 1)

EndFunc



Func _ER_GetAnnullering( $html)

;Annullering>false</Annullering>
;Annullering>true</Annullering>
return _ER_GetMatch( $html, '(?s)(Annullering)>true<' )

EndFunc

Func _ER_GetReseptId( $html)
;<Utlevering xmlns="http://www.kith.no/xmlstds/eresept/utlevering/2013-10-08">
;<ReseptId>
	return StringLeft(_ER_GetMatch( $html, '(?s)ReseptId>(.*?)<' ),9)
EndFunc

Func _ER_UtleveringId( $html)
;<AnnulleringsId>ad8d8ceb-
	return StringLeft(_ER_GetMatch( $html, '(?s)AnnulleringsId>(.*?)<' ),9)
EndFunc

Func _ER_GetKansellering( $html)

;<Kanselleringskode V="1" DN="Ikke ønsket vare"/>
Local $ret = _ER_GetMatch( $html, '(?s)Kanselleringskode.*?DN="(.*?)"' );
	return $ret ? "Kansellering " & $ret: 0
EndFunc



;------------------------------
; Apprec functions
;------------------------------
Func	_ER_GetApprec( $html)
	Local $text = ""

	Return _ER_GetApprecType( $html) & " " & _ER_GetApprecStatus($html) & " " & _ER_GetApprecError( $html )
	#cs
	if _ER_GetApprecType( $html) then $text = " " & _ER_GetApprecType( $html)
	if _ER_GetApprecStatus($html) then $text &= " " & _ER_GetApprecStatus($html)
	if _ER_GetApprecRef($html) then $text &= " " & _ER_GetApprecRef($html)
	return $text
#ce
EndFunc

#ifdef

#cs
<OriginalMsgId>
    <MsgType V="ERM921" DN="M9_21" />
    <IssueDate>2021-06-26T15:00:00.108+02:00</IssueDate>
    <Id>0920917c-ae57-43f9-a9e9-db309b302b47</Id>
  </OriginalMsgId>

#ce
Func	_ER_GetApprecRef( $html)
	return _ER_GetMatch( $html, '(?s)OriginalMsgId>.*?Id>(.*?)<' )
EndFunc

Func	_ER_GetApprecType( $html)
	return _ER_GetMatch( $html, '(?s)OriginalMsgId>.*?V="(.*?)"' )
EndFunc


Func	_ER_GetApprecStatus( $html)
	return _ER_GetMatch( $html, '(?s)Status.*?V="(.*?)".*?>' )
	;<Status V="2" DN="Avvist" />
EndFunc

Func	_ER_GetApprecError( $html)
	return _ER_GetMatch( $html, '(?s)Error.*?V="(.*?)".*?>' )
	; <Error V="360"
EndFunc


;------------------------------
; Apprec functions
;------------------------------
