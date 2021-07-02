#include-once
#include "..\lib_ut.au3"
#include "..\..\lib_msg.au3"

Test("_ER_GetRefToParent")
UTAssertEqual( _ER_GetRefToParent($html), 0)
