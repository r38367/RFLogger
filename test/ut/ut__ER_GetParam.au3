#include-once
#include "..\lib_ut.au3"
#include "..\..\lib_msg.au3"

Test("_ER_GetParam")
UTAssertEqual( _ER_GetMatch("tag>123abc-<", "tag>([0-9a-f-]*?)<"), "123abc-" )
;UTAssertEqual( _ER_GetMatch("C:\Users\ang_\Documents\github\RFLogger\test\", "(.*\\test\\).*?"), "a" )

