#include-once
#include "..\lib_ut.au3"
#include "..\..\lib_msg.au3"

Test("_ER_GetApprecError")
UTAssertEqual( _ER_GetApprecError(UTFileRead("Apprec_2_49.xml")), 49)
