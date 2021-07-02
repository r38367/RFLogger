#include-once

#include <Array.au3>

OnAutoItExitRegister("FlushTestResults")


#cs ----------------------------------------------------------------------------

 Micro unit test framework

	Test( $testName = @ScriptName )
	UTAssertTrue(Const $bool, Const $msg = "Assert Failure", Const $erl = @ScriptLineNumber)
	UTAssertFalse(Const $bool, Const $msg = "Assert Failure", Const $erl = @ScriptLineNumber)
	UTAssertEqual(Const $a, Const $b, Const $msg = "Assert Failure", Const $erl = @ScriptLineNumber)
	UTAssertFileEqual( Const $file1, Const $file2, Const $ignorewhitespace = true, Const $erl = @ScriptLineNumber )
	UTFileRead( Const $file1 )

#ce ----------------------------------------------------------------------------

Global $iTestCount = 0
Global $iTotalAssertions =0
Global $iTotalPass = 0
Global $iTotalFail = 0
Global $aResults[1]
Global $TimeEnd
Global $iTimeBegin = TimerInit()
Global $bUseSetup = false
Global $sCurrentTestName = ""
Global $errFileNotFound = 0

Func Test( $testName = @ScriptName )
	; Reset Global Variables
	$iTestCount += 1

	If $iTestCount > 1 Then
		;Call ("TestTearDown" )
	EndIf

	$sCurrentTestName = $testName
dbg( $iTestCount & ";" & $testName )

	;Call( "TestSetup" )
EndFunc


Func UTAssertTrue(Const $bool, Const $msg = "Assert Failure", Const $erl = @ScriptLineNumber)
    $iTotalAssertions += 1
	If NOT $bool Then
        DoTestFail($msg, $erl)
    Else
		DoTestPass()
	EndIf

EndFunc

Func UTAssertFalse(Const $bool, Const $msg = "Assert Failure", Const $erl = @ScriptLineNumber)
    Return UTAssertTrue( not $bool, $msg, $erl )
EndFunc

Func UTAssertEqual(Const $a, Const $b, $msg = "Assert Failure", Const $erl = @ScriptLineNumber)
   	$iTotalAssertions += 1
	If $a == $b Then
		DoTestPass()
	Else
		if IsString($errFileNotFound) then
			$msg = "File " & $errFileNotFound & " not found"
			$errFileNotFound = 0
		endif
		DoTestFail( StringFormat( $msg & " ( was=[%s], expected=[%s] )", $a, $b ), $erl)
	EndIf

EndFunc

Func UTAssertFileEqual( Const $file1, Const $file2, Const $ignorewhitespace = true, Const $erl = @ScriptLineNumber )
	Local $str1, $str2
	$iTotalAssertions += 1

	$str1 = FileRead( $file1 )
	if @error = 1 then
		DoTestFail( "Can't open " & $file1, $erl )
		return
	Else
		DoTestPass()
	EndIf

	; strip parts that can be changing
	; strip GUID
	$str1 = StringRegExpReplace( $str1, "(?i)(?s)(<ident>\s*<id>)([^/]*?)(</id>[^/]*?XXX.*?</ident>)", "$1$3" )

	if $ignorewhitespace then $str1 = StringStripWS( $str1, 8 )

	$iTotalAssertions += 1
	$str2 = FileRead( $file2 )
	if @error = 1 then
		DoTestFail( "Can't open " & $file2, $erl  )
		return
	Else
		DoTestPass()
	EndIf

	if $ignorewhitespace then $str2= StringStripWS( $str2, 8 )



	return UTAssertEqual( $str1, $str2, "Files differ", $erl )
;~ 	if $str1 == $str2 then
;~ 		DoTestPass()
;~ 	Else
;~ 		DoTestFail( StringFormat( "Files differ: file1=[%s], file2=[%s]", $file1, $file2 ), $erl)
;~ 	EndIf

;~ 	return 3

EndFunc

Func DoTestPass()
	$iTotalPass += 1
	;ConsoleWrite( "." )
EndFunc

Func DoTestFail( $msg, $erl )
	$iTotalFail += 1
	Local $message = @ScriptName & " [" & $sCurrentTestName & "] (" & $erl & ") := " & $msg & @LF
	;StringFormat( "%-15s %s", $iTotalAssertions, "Error in " & $sCurrentTestName & " -> " & $msg  )
	_ArrayAdd( $aResults, $message )
	ConsoleWrite( $message )
EndFunc


Func FlushTestResults()
	;Call ("TestTearDown" )

	ConsoleWrite( StringFormat( "Tests:%d, Total:%d, Pass:%d, Fail:%d ", _
		$iTestCount, $iTotalAssertions, $iTotalPass, $iTotalFail ) & @CRLF)

	;$TimeEnd = Round(TimerDiff($iTimeBegin)/1000 , 3)
	;ConsoleWrite( "Total Time: " & $TimeEnd & " seconds" & @CRLF )

EndFunc

Func dbg(Const $msg )
	if @Compiled then
		MsgBox( 0, "Debug", $msg )
	Else
		ConsoleWrite( $msg & @LF)
	EndIf
	return 0
EndFunc   ;==>dbg

Func err(Const $msg, Const $error = @error, Const $extended = @extended, Const $erl = @ScriptLineNumber)
	ConsoleWrite("(" & $erl & ") : = (" & $error & ")(" & $extended & ") " & $msg & @LF)

	If $error <> 0 Then SetError($error, $extended, $error)

	Return $error
EndFunc   ;==>err

; Read from demo file located in .\test\files\..
; $path2files = StringRegExReplace( @WorkingDir, "(.*\\test\\).*", "$1files")
;

Func 	UTFileRead( Const $file )
	Local $f ;, $aPaths = [ "", ".\files\", "..\files\" ]
	$f = StringRegExpReplace( @WorkingDir, "(.*\\test\\).*", "$1files\\" & $file )
ConsoleWrite( @WorkingDir& @CRLF & $f & @CRLF )

;	for $sPath in $aPaths

	;$f = $sPath & $file

		if FileExists( $f ) then Return FileRead( $f )

;	Next

	$errFileNotFound = $file

	Return 0 ; "File not found"

EndFunc ;-> UTReadFile

Func 	UTAssertFileNameEqual( $func, $file, Const $erl = @ScriptLineNumber )

	$s = UTFileRead( $file)
	if IsNumber($s) then
		DoTestFail( $file & " not found XXX", $erl )
		return 0
	endif

	; Remove XML ->
	$file = StringRegExpReplace( $file, "(.*?)[_ ](.*)\.(.*)", "$2" )
	$file = StringReplace( $file, "_", " " )

	UTAssertEqual( $func( $s ), $file )

EndFunc
