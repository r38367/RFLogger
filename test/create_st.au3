#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here

; $file = au3 file with functions
; for each $function
;	create $ut_file with name ut_$function.au3
;		#include-once
;		#include "..\ut_lib.au3"
;		#include "..\..\$file"
;
;		Test("$file")
; 		UTAssertEqual( $function(0,0), 1)
;
;	if $ut_file exists in $ut_suite.au3 then
;		add #include "$ut_file" to $ut_suite.au3
; next
;