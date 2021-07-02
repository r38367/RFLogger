#include-once
#include "..\lib_ut.au3"
#include "..\..\lib_msg.au3"

Test("_ER_GetMsgType")
UTAssertEqual( _ER_GetMsgType($html), 0)
