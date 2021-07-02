#Region	Include

#include <FileConstants.au3>
#include <MsgBoxConstants.au3>
#include <Array.au3>

#EndRegion Include
#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         myName

 Script Function:
	Create unit test structure

	.\run_ut.au3 -> run all unit tests
	.\lib_ut.au3
	.\run_ut_$file.au3 -> run unit tests for all functions in file
	.\ut\$file\
	.\ut\$file\ut_$function.au3 -> run unit test for particular function
	.\files - demo files

	 For each function in the selected file creates a folder with file name with unit test for each function.



#ce ----------------------------------------------------------------------------

; select the file with functions you want to create unittests for
Local $sFile = GetFileForUT()
if @error then
	MsgBox($MB_SYSTEMMODAL, $sFile, "No file(s) were selected.")
	Exit
endif

; strip fname from full path
$filename = StringRegExpReplace( $sFile, "(.*\\)?(.*)", "$2")

; get list of functions from file into array
;	[function name, param_list]
;	param_list = { "", "param", "param,param,..." }
$aFuncList = StringRegExp( FileRead( $sFile ), "(?m)^Func\s+(.*?)\(\s?(.*?)\s?\)", 3 )
;ConsoleWrite( @error & " " & @extended & @CRLF )
;_ArrayDisplay( $aFuncList )

; for each function from array
for $i=0 to UBound($aFuncList)-1 step 2
	Local $functionName = $aFuncList[$i]
	Local $functionParam = $aFuncList[$i+1]

ConsoleWrite( $i/2+1 & " " )
	; unit test filename
	Local $ut_file = "ut\ut_" & $functionName & ".au3" ; ut file name
ConsoleWrite( $ut_file & " " )

	; check if file exists
	If FileExists( $ut_file ) then
		ConsoleWrite( "exists" & @CRLF )
	Else

		; create ut folder
		if not FileExists("ut") then DirCreate( "ut" )
	;ConsoleWrite( "Dir cre " & @error & @CRLF )

		; write template file
		Local $text = '#include-once' & @CRLF _
			& '#include "..\lib_ut.au3"' & @CRLF _
			& '#include "..\..\' & $filename & '"' & @CRLF _
			& @CRLF _
			& 'Test("' & $functionName & '")' & @CRLF _
			& 'UTAssertEqual( ' & $functionName & '(' & $functionParam & '), 0)' & @CRLF


		FileWrite( $ut_file, $text )
	;	FileWrite( $hfile, $text )
		;FileClose( $hfile)
		ConsoleWrite( "created " & @error & @CRLF )
	endif


	; add unit test to testsuite if it's not there yet
	$ut_suite = "run_ut.au3"

	if not FileExists( $ut_suite) then
		; create suite
		$text = '#include "lib_ut.au3"' & @CRLF & @CRLF
		FileWrite( $ut_suite, $text)
		ConsoleWrite( "Created " & $ut_suite & @CRLF )

	endif

	; chech if ut already exists
	if StringInStr( FileRead( $ut_suite), $ut_file , 0 ) then ;'#include\s+"' & $ut_file & '"', 0 ) then
		; exists
		ConsoleWrite( $ut_file & " existed in " & $ut_suite & @CRLF )
		ContinueLoop
	endif

	$text = '#include "' & $ut_file & '"' & @CRLF
	FileWrite( $ut_suite, $text)
	ConsoleWrite( $ut_file & " added to " & $ut_suite & @CRLF )

Next

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



Func	GetFileForUT()

 ; Create a constant variable in Local scope of the message to display in FileOpenDialog.
    Local Const $sMessage = "Hold down Ctrl or Shift to choose multiple files."

    ; Display an open dialog to select a list of file(s).
    Local $sFileOpenDialog = FileOpenDialog($sMessage,  "..\", "au3 (*.au3)", $FD_FILEMUSTEXIST) ;BitOR($FD_FILEMUSTEXIST, $FD_MULTISELECT))
    If @error Then
        ; Display the error message.
        MsgBox($MB_SYSTEMMODAL, @error, "No file(s) were selected." & @error)

        ; Change the working directory (@WorkingDir) back to the location of the script directory as FileOpenDialog sets it to the last accessed folder.
        FileChangeDir(@ScriptDir)
		return SetError(1)
    Else
        ; Change the working directory (@WorkingDir) back to the location of the script directory as FileOpenDialog sets it to the last accessed folder.
        FileChangeDir(@ScriptDir)

        ; Replace instances of "|" with @CRLF in the string returned by FileOpenDialog.
        $sFileOpenDialog = StringReplace($sFileOpenDialog, "|", @CRLF)

        ; Display the list of selected files.
        return $sFileOpenDialog
    EndIf
EndFunc   ;==>Example
