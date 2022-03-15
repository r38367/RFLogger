#Region ***** includes
#include-once
#include <Date.au3>

#EndRegion ***** includes

#Region ***** functions
;=================================================
; Functions spesific to RF Admin time format
; _RFtime( $sDate ) - convert time from yyyy.mm.dd hh:mm:ss -> dd.mm.yyyy hh:mm:ss
; _RFtimeTomorow - returns next midnight in dd.mm.yyyy 00:00:00
; _RFtimeNow - returns current time in RF Admin format -> dd.mm.yyyy hh:mm:ss
; _RFtimeDiff( $sNum, $sType) - return time diff in D,M,Y,w,h,n,s
;
;=================================================
#EndRegion ***** functions

#Region *** Global variables

#EndRegion Global Variables


;=================================================
; _RFtime( $sDate ) - convert time from yyyy.mm.dd hh:mm:ss -> dd.mm.yyyy hh:mm:ss
;=================================================
Func _RFtime( $sDate )
	Return StringRegExpReplace( $sDate, "(\d\d\d\d).(\d\d).(\d\d).(\d\d:\d\d:\d\d)", '$3.$2.$1 $4')
EndFunc

;=================================================
; _RFtimeTomorow - returns next midnight
;=================================================
Func 	_RFtimeTomorow()
Return _RFtime( StringLeft(_DateAdd ( 'D', 1, _NowCalc() ), 11 ) & "00:00:00" )
EndFunc

;=================================================
; _RFtimeNow - returns current time in RF Admin format
;=================================================
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
Func 	_RFtimeDiff( $sNum, $sType, $sDate = _NowCalc())
Return	_RFtime(  _DateAdd ( $sType, $sNum, $sDate ) )
EndFunc

#Region ***** Unit testing
#include "test/unittest.au3"

Test("_RFTime")
UTAssertEqual( _RFTime( "2022-02-21T11:06:39.968Z" ), "21.02.2022 11:06:39.968Z" )
UTAssertEqual( _RFTime( "2022-02-01T12:26:39.20" ), "01.02.2022 12:26:39.20" )
UTAssertEqual( _RFTime( "2022-12-31T13:46:39" ), "31.12.2022 13:46:39" )

Test("_RFTimeNow")
UTAssertEqual( _RFtimeNow(), _RFTime( _NowCalc() ) )

Test("_RFTimeTomorrow")
UTAssertEqual( _RFtimeTomorow(), _RFTime( _DateAdd( "D", 1, _NowCalcDate() ) & " 00:00:00" )  )

Test("_RFTimeDiff")
UTAssertEqual( _RFtimeDiff( 2, "M"), _RFTime( _DateAdd( "M", 2, _NowCalc() )  )   )
UTAssertEqual( _RFtimeDiff( 2, "Y", "2022-03-01 05:22:44"),  "01.03.2024 05:22:44"  )

#EndRegion ***** Unit testing