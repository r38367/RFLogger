#Region ***** includes
#include-once
#include "mheader.au3"

#EndRegion ***** includes

#Region ***** functions

#EndRegion ***** functions

#Region *** Global variables

#EndRegion Global Variables

;=================================================
; _ER_GetM10($html) - return text with parameters
;=================================================

Func _get_ERM10($html)

	Local $text =""

	$text &= " " & _get_ERM10_kansellering( $html )
	$text &= " " & _get_ERM10_annullering( $html )
	;if _ER_GetRefHjemmel($html) then $text &= " " & _ER_GetRefHjemmel($html)
	;if _ER_GetNavnFormStyrke($html) then $text &= " " & _ER_GetNavnFormStyrke($html)
	;if _ER_GetPatient($html) then $text &= " " & _ER_GetPatient($html)
	;if _ER_GetFnr($html) then $text &= " " & _ER_GetFnr($html)
	;if _ER_GetDateOfBirth($html) then $text &= " " & _ER_GetDateOfBirth($html)
	;if _ER_GetReseptId( $html) then $text &= " " & _ER_GetReseptId($html)
	$text &= " " & _get_ERM10_papirresept( $html )

	Return	StringStripWS( $text, 7)

EndFunc

Func _get_ERM10_kansellering( $html)

;<Kanselleringskode V="1" DN="Ikke ønsket vare"/>
	Local $ret = _get_param( $html, '(?s)Kanselleringskode.*?DN="(.*?)"' );
	return $ret? "Kansellering(" & $ret & ")": ""
EndFunc

Func _get_ERM10_annullering( $html)

	Local $ret = _get_param( $html, '(?s)Annullering>true<' );
	Local $antall = _get_param( $html, '(?s)Antall>(.*?)<' );

	return $ret? "Annullering("&$antall&")": ""
EndFunc

Func	_get_ERM10_papirresept( $html )
	Local $ret = _get_param( $html, '(?s)Papirresept>true<' );
	return $ret ? "papir": ""

EndFunc


#Region ***** Unit testing
#include "test/unittest.au3"

Test("M10")
UTAssertEqual( _get_ERM10_papirresept( UTFileRead("ERM10_papir.xml") ), "papir")
UTAssertEqual( _get_ERM10_kansellering( UTFileRead("ERM10_kansellering.xml") ), "Kansellering(Uavhentet vare)")
UTAssertEqual( _get_ERM10_annullering( UTFileRead("ERM10_annullering.xml") ), "Annullering(0)")

UTAssertEqual( _get_ERM10( UTFileRead("ERM10_papir.xml") ), "papir")

#EndRegion
