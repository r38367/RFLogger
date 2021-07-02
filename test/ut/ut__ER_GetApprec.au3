#include-once
#include "..\lib_ut.au3"
#include "..\..\lib_msg.au3"

Test("_ER_GetApprec")
UTAssertEqual( _ER_GetApprec( UTFileRead("Apprec_1.xml") ), "ERM252 1 0"  )
UTAssertEqual( _ER_GetApprec( UTFileRead("Apprec_2_49.xml") ), "ERM41 2 49" )
UTAssertEqual( _ER_GetApprec( UTFileRead("Apprec_3_multierror.xml") ), "ERMV 3 320" )
