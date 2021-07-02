#include-once
#include "..\lib_ut.au3"
#include "..\..\lib_msg.au3"

Test("_ER_GetM10")
UTAssertFileNameEqual( _ER_GetM10, "ERM10_§2_Zocor Tab 40 mg_Herman_22109345931_ef82e120-.xml" )

UTAssertEqual( _ER_GetM10( UTFileRead("M10_forbruksmateriell_508.xml") ), "508 Jo 03020271700 52123ba1-")
UTAssertEqual( _ER_GetM10( UTFileRead("M10_næringsmiddel_603.xml") ), "603 Jo 03020271700 9ad72a18-")

UTAssertEqual( _ER_GetM10( UTFileRead("M10 annullering lakselv.xml") ), "Annullering Levitra Tab 10 mg Vikebe 17104714420 7925e3e8-")
UTAssertEqual( _ER_GetM10( UTFileRead("M10_kansellering.xml") ), "Kansellering Uavhentet vare 17094393213 e4738aa3-")

;UTAssertEqual( _ER_GetAnnullering(""), "Fim 1995-12-09 §3 0be321- 234b234- 06345a-" ) ; annullering Fdato
; annullering magistrell §3
; annullering §5
; annullering §2
; annullering hvit
