#COMPILE EXE
#DIM ALL
'
'%SupressSQLErrors=-1 'No error reporting from MSSQL on errors. Remove this constant to display errors.
'%DEBUG=-1
#INCLUDE "tsh_MSSQL.INC"
' The name of the service
$SERVICE_NAME = "LogFeeder"
$SERVICE_DISPLAY_NAME = "LogFeeder "
$SERVICE_DESCRIPTION  = "Citera LogFeeder"
'
'- REM OUT this line to compile as a console application
%COMPILE_AS_SERVICE = 1
'
#IF %DEF(%COMPILE_AS_SERVICE)
  #INCLUDE "pb_srvc.inc"
#ENDIF


FUNCTION PBMAIN () AS LONG
  LOCAL lngResult AS LONG
  LOCAL lngWaitThreadID AS LONG

   '- This is the thread in the program
   '  that actually does the work
   '
   THREAD CREATE waitThread(0) TO lngWaitThreadID
   '
    '- Start the service
   #IF %DEF(%COMPILE_AS_SERVICE)
     pbsInit 0, $SERVICE_NAME, $SERVICE_DISPLAY_NAME, _
                $SERVICE_DESCRIPTION

     '- Run in a console
   #ELSE

     CON.STDOUT "Press ESC to shutdown server properly"
     DO
       IF INKEY$ = $ESC THEN
         CON.STDOUT "CONSOLE CANCELLED "
         EXIT DO
       END IF
       SLEEP 1000
     LOOP
   #ENDIF
     '
    '- Clean-up the thread handles
   THREAD CLOSE lngWaitThreadID TO lngResult

END FUNCTION

THREAD FUNCTION waitThread ( BYVAL lngDoNothing AS LONG ) AS LONG
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'  waitThread
'
'  This thread loops and is the thread you'd
'  use for your service to do something.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    LOCAL lRecs,test AS LONG

    LOCAL lConStr,lConStr2 AS STRING

    #INCLUDE "tsh_ConStr.inc"

    DIM lResultAry() AS VARIANT

    LOCAL lId,lIdx AS LONG
    LOCAL lStatusBar AS STRING

    DO

    lRecs = TsH_MSSQL_Select(lConstr, "Select ID,StatusBar from dbo.LogEntries;",lResultAry())

    IF lRecs = -1 THEN GOTO jazz

    FOR lIdx = 0 TO lRecs
        lId                 = VAL(AfxVarToStr(lResultAry(0,lIdx)))
        lStatusBar          = "y" + AfxVarToStr(lResultAry(1,lIdx))

        TsH_MSSQL_Execute(lConStr, "Update dbo.LogEntries set StatusBar='" & lStatusBar  & "' Where ID=" & FORMAT$(lId) & ";")

   NEXT

   jazz:

   '5 minutes
   SLEEP 300000

   LOOP

END FUNCTION
