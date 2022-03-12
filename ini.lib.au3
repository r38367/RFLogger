#Region ***** includes
#include-once
;#include <array.au3>
#EndRegion ***** includes

#cs
================================
Function for work with .ini file:

_ini_setFilename( $path ) - set fille path and name
_ini_getFilename() - return current ini file path and name
_ini_read( $par ) - retuun $par value from ini file
_ini_save( $par, $val ) - saves $par in ini file
_ini_delete() - deletes current ini file

================================

#ce

#Region *** Global variables

Global $_ini_filename = @ScriptFullPath & ".ini" ; ini file name
Global $_ini_ar  ; array with parameters

#EndRegion Global variables

Func	_ini_setFilename( $path )

	$_ini_filename = $path

EndFunc

Func	_ini_getFilename()

	return $_ini_filename

EndFunc

Func	_ini_read( $par )

	Local $ar = IniReadSection( $_ini_filename, "RFAdmin" )

	if IsArray( $ar ) then
		for $i=1 to $ar[0][0]
			if $ar[$i][0] = $par then
				return $ar[$i][1]
			EndIf
		Next
		return 2 ; $par not found

	else
		return 1 ; ini file not found
	EndIf

EndFunc

Func	_ini_save( $par, $val )

	IniWrite($_ini_filename, "RFAdmin", $par, $val)

EndFunc

Func	_ini_delete()

	Return FileDelete($_ini_filename)

EndFunc


#Region ***** Unit testing
#include "test/unittest.au3"
Global $path

Test("default filename")
$path = @ScriptFullPath &".ini"
UTAssertEqual( _ini_getFilename(), $path )
UTAssertEqual( _ini_save( 'Aktor', "foggia" ), 0)
UTAssertEqual( _ini_save( 'User', "Anton@nhn.no" ), 0)
UTAssertEqual( FileExists($path), 1)
UTAssertEqual( _ini_read( 'Aktor' ), "foggia" )
UTAssertEqual( _ini_read( 'User' ), "Anton@nhn.no" )
UTAssertEqual( _ini_read( 'not *** existed' ), 2 )
UTAssertEqual( _ini_delete(), 1)
UTAssertEqual( FileExists($path), 0)



Test("only filename")
$path = "x.x.ini"
UTAssertFalse( _ini_setFilename( $path ) )
UTAssertEqual( _ini_getFilename( ), $path )
UTAssertEqual( _ini_save( 'Aktor', "foggia" ), 0)
UTAssertEqual( FileExists($path), 1)
UTAssertEqual( _ini_read( 'Aktor' ), "foggia" )
UTAssertEqual( _ini_delete(), 1)
UTAssertEqual( FileExists($path), 0)


Test("path with folder" )
$path = "test/x.x.ini"
UTAssertFalse( _ini_setFilename( $path ) )
UTAssertEqual( _ini_getFilename( ), $path )
UTAssertEqual( _ini_save( 'User', 12345 ), 0)
UTAssertEqual( FileExists($path), 1)
UTAssertEqual( _ini_read( 'User' ), 12345 )
UTAssertEqual( _ini_delete(), 1)
UTAssertEqual( FileExists($path), 0)


#EndRegion
