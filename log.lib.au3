#Region ***** includes
#include-once
#EndRegion ***** includes

#Region ***** functions
;=================================================
; functions spesific to logging
; _log_file( $path ) - set log file name
; _log_write( $text ) - write text to log file
; _log_clear() - clear debug file
;=================================================
#EndRegion ***** functions

#Region *** Global variables
Global $logFile = "message.log"

#EndRegion Global Variables


Func	_log_file( $fname )
	$logFile = $fname
EndFunc

Func	_log_write( $txt )
	Return FileWriteLine( $logFile, $txt )
EndFunc

Func _log_clear()
	Return FileDelete( $logFile )
EndFunc



#Region ***** Unit testing
#include "test/unittest.au3"

Local $file ="unittest.log.log"
Local $text = "english åøæ русский" & @CRLF

Test("debug")
UTAssertEqual( _log_file( $file), 0 )
UTAssertEqual( _log_write( $text ), 1)
UTAssertEqual( FileExists( $file ), 1 )
UTAssertEqual( FileRead( $file), $text )
UTAssertEqual( _log_clear(), 1)
UTAssertEqual( FileExists( $file ), 0 )


#EndRegion ***** Unit testing