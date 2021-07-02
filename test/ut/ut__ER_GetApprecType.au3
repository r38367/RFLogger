#include-once
#include "..\lib_ut.au3"
#include "..\..\lib_msg.au3"

Test("_ER_GetApprecType")
UTAssertEqual( _ER_GetApprecType(UTFileRead("Apprec_1.xml")), "ERM252")
