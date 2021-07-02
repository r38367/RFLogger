#include-once
#include "..\lib_ut.au3"
#include "..\..\lib_msg.au3"

Test("_ER_GetBody")
UTAssertEqual( _ER_GetBody( UTFileRead("IEBodyText.html") ), UTFileRead( "IEBodyText_striped.html") )
