#include-once
#include "..\lib_ut.au3"
#include "..\..\lib_msg.au3"

Test("_ER_GetMsgId")
UTAssertEqual( _ER_GetMsgId($html), 0)
