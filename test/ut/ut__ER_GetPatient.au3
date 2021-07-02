#include-once
#include "..\lib_ut.au3"
#include "..\..\lib_msg.au3"

Test("_ER_GetPatient")
UTAssertEqual( _ER_GetPatient($html, $type=0), 0)
