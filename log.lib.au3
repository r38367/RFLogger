#Region ***** includes
#include-once
#EndRegion ***** includes

#Region ***** functions
;=================================================
; functions spesific to logging
; _log_file( $path ) - set log file name - [path\]file.log
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

#include <FileConstants.au3>
Func	_log_write( $txt )
	if not FileExists( $logFile ) then
		Local $hf = FileOpen( $logFile ,$FO_OVERWRITE + $FO_CREATEPATH )
		if $hf = -1 then return 2
		FileClose($hf)
	EndIf
	Return FileWriteLine( $logFile, $txt )
EndFunc

Func _log_clear()
	Return FileDelete( $logFile )
EndFunc



#Region ***** Unit testing
#include "test/unittest.au3"

Local $file ="test1.log"
Local $text = "english åøæ русский" & @CRLF

Test("log file in current folder")
UTAssertEqual( _log_file( $file), 0 )
UTAssertEqual( _log_write( $text ), 1)
UTAssertEqual( FileExists( $file ), 1 )
UTAssertEqual( FileRead( $file), $text )
UTAssertEqual( _log_clear(), 1)
UTAssertEqual( FileExists( $file ), 0 )

Test("log file in existed folder")
$file ="test\test2.log"
UTAssertEqual( _log_file( $file), 0 )
UTAssertEqual( _log_write( $text ), 1)
UTAssertEqual( FileExists( $file ), 1 )
UTAssertEqual( FileRead( $file), $text )
UTAssertEqual( _log_clear(), 1)
UTAssertEqual( FileExists( $file ), 0 )

Test("log file in non-existed folder")
$dir="non-ex"
$file=$dir & "\test3.log"

UTAssertEqual( _log_file( $file), 0 )
UTAssertEqual( _log_write( $text ), 1)
UTAssertEqual( FileExists( $file ), 1 )
UTAssertEqual( FileRead( $file), $text )
UTAssertEqual( _log_clear(), 1)
UTAssertEqual( FileExists( $file ), 0 )
DirRemove($dir)

#EndRegion ***** Unit testing