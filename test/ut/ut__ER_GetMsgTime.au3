#include-once
#include "..\lib_ut.au3"
#include "..\..\lib_msg.au3"

Test("_ER_GetMsgTime")
UTAssertEqual( _ER_GetMsgTime($html), 0)
