#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Common functions. If fails then other will fail
#include "ut\ut__ER_GetBody.au3"
#include "ut\ut__ER_GetParam.au3"

; Run M1

#cs
#include "ut\ut__ER_GetExtraParam.au3"
#include "ut\ut__ER_GetMsgId.au3"
#include "ut\ut__ER_GetMsgType.au3"
#include "ut\ut__ER_GetMsgTime.au3"
#include "ut\ut__ER_GetRefToParent.au3"
#include "ut\ut__ER_GetRefToConversation.au3"
#include "ut\ut__ER_GetM1.au3"
#include "ut\ut__ER_isV24.au3"
#include "ut\ut__ER_GetNavnFormStyrke.au3"
#include "ut\ut__ER_GetRefKode.au3"
#include "ut\ut__ER_GetRefHjemmel.au3"
#include "ut\ut__ER_GetPatient.au3"
#include "ut\ut__ER_GetDateOfBirth.au3"
#include "ut\ut__ER_GetFnr.au3"
#ce


; Run M10
#include "ut\ut__ER_GetAnnullering.au3"
#include "ut\ut__ER_GetKansellering.au3"
#include "ut\ut__ER_GetReseptId.au3"
#include "ut\ut__ER_UtleveringId.au3"
#include "ut\ut__ER_GetM10.au3"

; Run Apprec
#include "ut\ut__ER_GetApprecRef.au3"
#include "ut\ut__ER_GetApprecType.au3"
#include "ut\ut__ER_GetApprecStatus.au3"
#include "ut\ut__ER_GetApprecError.au3"
#include "ut\ut__ER_GetApprec.au3"

