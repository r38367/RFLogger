#include-once
#include "..\lib_ut.au3"
#include "..\..\lib_msg.au3"

Test("_ER_GetRefToConversation")
UTAssertEqual( _ER_GetRefToConversation($html), 0)
