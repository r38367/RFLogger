#include-once
#include "..\lib_ut.au3"
#include "..\..\lib_ie.au3"

Test("_IEGetPage")
UTAssertEqual( _IEGetPage($sLink), 0)
