#include-once
#include "..\lib_ut.au3"
#include "..\..\lib_ie.au3"

Test("_IEGetActiveTab")
UTAssertEqual( _IEGetActiveTab(), 0)
