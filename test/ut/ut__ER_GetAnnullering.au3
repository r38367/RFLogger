#include-once
#include "..\lib_ut.au3"
#include "..\..\lib_msg.au3"

Test("_ER_GetAnnullering")
UTAssertEqual( _ER_GetAnnullering( UTFileRead("M10 annullering lakselv.xml") ), "Annullering") ; annullering


