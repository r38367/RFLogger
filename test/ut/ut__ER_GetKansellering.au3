#include-once
#include "..\lib_ut.au3"
#include "..\..\lib_msg.au3"

Test("_ER_GetKansellering")
UTAssertEqual( _ER_GetKansellering( UTFileRead("M10_Kansellering.xml") ), "Kansellering Uavhentet vare")
