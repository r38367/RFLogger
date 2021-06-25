#include-once

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

Func	_ER_GetMsgTime( $html)

	Local $a
	$a = StringRegExp( $html, 'GenDate>(\d+).(\d+).(\d+).(\d+):(\d+):(\d+).*?</', 1)
	if @error then return 0
	Return $a[0] & $a[1] &  $a[2] & $a[3] & $a[4] & $a[5]

EndFunc
