#include-once
#include "..\lib_ut.au3"
#include "..\..\lib_msg.au3"

Test("_ER_GetApprecStatus")
UTAssertEqual( _ER_GetApprecStatus(UTFileRead("Apprec_2_49.xml")), "2")
