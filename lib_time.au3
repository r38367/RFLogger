#include <Date.au3>



Func 	_RFtimeTomorow()
Return _RFtime( StringLeft(_DateAdd ( 'D', 1, _NowCalc() ), 11 ) & "00:00:00" )
EndFunc

Func 	_RFtimeNow()
Return _RFtime( _NowCalc() )
EndFunc

;====================================================
;~ Time interval to be used:
;~ D - Add/subtract days to/from the specified date
;~ M - Add/subtract months to/from the specified date
;~ Y - Add/subtract years to/from the specified date
;~ w - Add/subtract Weeks to/from the specified date
;~ h - Add/subtract hours to/from the specified date
;~ n - Add/subtract minutes to/from the specified date
;~ s - Add/subtract seconds to/from the specified date
;=====================================================
Func 	_RFtimeDiff( $sNum, $sType)
Return	_RFtime(  _DateAdd ( $sType, $sNum, _NowCalc() ) )
EndFunc

Func _RFtime( $sDate )
	Return StringRegExpReplace( $sDate, "(\d\d\d\d).(\d\d).(\d\d) (\d\d:\d\d:\d\d)", '$3.$2.$1 $4')
EndFunc
