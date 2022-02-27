#include <Date.au3>

#cs
;====================================================
Functions to work with time in RFAdmin

_RFtime( $sDate ) - convert date to dd.mm.yyyy hh:mm:ss
_RFtimeNow() - return current RF time in format dd.mm.yyyy hh:mm:ss
_RFtimeTomorow() - return tomorrow modnight time
_RFtimeDiff( $sNum, $sType, $sDate) - return time +/- timediff
_RFtimeLast( $sRfDate ) - set/get last RF message time

;====================================================
#ce

#Region Global Variables
;====================================================
Global $sLastMsgTime = 0
;====================================================
#EndRegion Global Variables

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
Func 	_RFtimeDiff( $sNum, $sType, $sDate=Default)
Return	_RFtime(  _DateAdd ( $sType, $sNum, $sDate=Default? _NowCalc(): $sDate ) )
EndFunc

Func _RFtime( $sDate )
	Return StringRegExpReplace( $sDate, "(\d\d\d\d).(\d\d).(\d\d) (\d\d:\d\d:\d\d)", '$3.$2.$1 $4')
EndFunc


Func _RFtimeLast( $sRfDate=0 )

if $sRfDate = 0 then
	if $sLastMsgTime = 0 then  $sLastMsgTime = _RFtimeDiff( -10, 'n')
Else
	$sLastMsgTime = $sRfDate
EndIf

return $sLastMsgTime

EndFunc

