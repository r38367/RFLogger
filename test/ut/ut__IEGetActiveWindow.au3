#include-once
#include "..\lib_ut.au3"
#include "..\..\lib_ie.au3"

Test("_IEGetActiveWindow")
UTAssertEqual( _IEGetActiveWindow(), 0)
