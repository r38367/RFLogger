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

	Return	StringStripWS( $text, 7)

EndFunc


#Region ***** Unit testing
#include "test/unittest.au3"

Test("M10")
UTAssertEqual( _get_ERM10( UTFileRead("APPREC_2_49.xml") ), "")

#EndRegion
