#Region ***** includes
#include-once
#EndRegion ***** includes

#Region ***** functions
;=================================================
; functions spesific to code debugging
; _dbg_file( $path ) - set debug file name
; _dbg_write( $text ) - write text to debug file
; _dbg_clear() - clear debug file
;=================================================
#EndRegion ***** functions

#Region *** Global variables
Global $debugFile = "_debug.log"

#EndRegion Global Variables


Func	_dbg_file( $fname )
	$debugFile = $fname
EndFunc

Func	_dbg_write( $txt )
	FileWriteLine( $debugFile, $txt )
EndFunc

Func _dbg_clear()
	FileDelete( $debugFile )
EndFunc



#Region ***** Unit testing
#include "test/unittest.au3"

Local $file ="unittest.debug.log"
Local $text = "english åøæ русский" & @CRLF

Test("debug")
UTAssertEqual( _dbg_file( $file), 0 )
UTAssertEqual( _dbg_write( $text ), 0)
UTAssertEqual( FileExists( $file ), 1 )
UTAssertEqual( FileRead( $file), $text )
UTAssertEqual( _dbg_clear(), 0)
UTAssertEqual( FileExists( $file ), 0 )


#EndRegion ***** Unit testing