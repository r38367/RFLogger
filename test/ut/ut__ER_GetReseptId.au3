#include-once
#include "..\lib_ut.au3"
#include "..\..\lib_msg.au3"

Test("_ER_GetReseptId")
UTAssertEqual( _ER_GetReseptId( UTFileRead("M10 annullering lakselv.xml") ), "7925e3e8-")
