#include-once
#include "..\lib_ut.au3"
#include "..\..\lib_msg.au3"

Test("_ER_GetApprecRef")
UTAssertEqual( _ER_GetApprecRef( UTFileRead("Apprec_1.xml") ), "eeb97459-608a-4466-8349-2e93e32f80fa")
