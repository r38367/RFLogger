#include-once
#include "..\lib_ut.au3"
#include "..\..\lib_msg.au3"

Test("_ER_UtleveringId")
UTAssertEqual( _ER_UtleveringId(UTFileRead("M10 annullering lakselv.xml")), "ad8d8ceb-")
