#include "..\..\lib_msg.au3"
#include "..\lib_ut.au3"

Test( "_ER_GetApprecError" )
Func	_ER_GetApprec( $html)



Test( "_ER_GetApprec" )
UTAssertEqual( FileRead( "Apprec_1.xml ),  )


#cs
Func	_ER_GetApprec( $html)
	Local $text = ""

	if _ER_GetApprecType( $html) then $text = " " & _ER_GetApprecType( $html)
	if _ER_GetApprecStatus($html) then $text &= " " & _ER_GetApprecStatus($html)
	if _ER_GetApprecRef($html) then $text &= " " & _ER_GetApprecRef($html)
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
#ce
