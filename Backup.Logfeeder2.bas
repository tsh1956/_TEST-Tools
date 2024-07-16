#COMPILE EXE
#DIM ALL
'%SupressSQLErrors=-1 'No error reporting from MSSQL on errors. Remove this constant to display errors.
%DEBUG=-1
#INCLUDE "tsh_MSSQL.INC"
' The name of the service
$SERVICE_NAME = "LogFeeder"
$SERVICE_DISPLAY_NAME = "LogFeeder"
$SERVICE_DESCRIPTION  = "LogFeeder Helper App"
'
'- REM OUT this line to compile as a console application
%COMPILE_AS_SERVICE = 0
'
#IF %DEF(%COMPILE_AS_SERVICE)
  #INCLUDE "pb_srvc.inc"
#ENDIF




FUNCTION PBMAIN () AS LONG

    'LOCAL Built AS IPowerTime

    'LET Built = CLASS "PowerTime"

    LOCAL lIn_1,lIn_2,lIn_3,lIn_4,lIn_5 AS STRING
    LOCAL lStatus,lStatusBar AS STRING
    RESET lStatus,lStatusBar

    LOCAL lID AS LONG
    LOCAL lLogMessage AS STRING
    lIn_1 = COMMAND$(1) 'LogMessage Unique name
    lIn_2 = COMMAND$(2) 'statusflag
    lIn_3 = COMMAND$(3) 'email recipient(s) for alerts
    lIn_4 = COMMAND$(4) '
    lIn_5 = COMMAND$(5)

    IF LEN(lIn_1) <> 0 THEN
        lId = VAL(lIn_1)
        lLogMessage = UCASE$(TRIM$(lIn_1))
    ELSE
        STDOUT "no ID":EXIT FUNCTION
    END IF

    IF LEN(lIn_2) = 1 THEN
        IF TALLY(LCASE$(lIn_2),ANY "abcdexy") > 0 THEN lStatus = LCASE$(lIn_2)
    ELSE
        lStatus = ""
        STDOUT "no status flag":EXIT FUNCTION
    END IF

    LOCAL lRecs,lCount,lRnd AS LONG

    LOCAL lConStr,lConStr2 AS STRING
    #INCLUDE "tsh_ConStr.inc"

    DIM lResultAry() AS VARIANT

    LOCAL lNumOfIterations,lSecsSinceLastUpDate,lDayNumber,lSecsDelta,lDayDelta,lSecsADay,lTick AS LONG
    LOCAL lStatusBarInternal,lBlancs AS STRING

    RESET lBlancs

    LOCAL mytimeVar AS DOUBLE
    mytimeVar = TIMER

    lRecs = TsH_MSSQL_Select(lConstr, "Select Status,StatusBar,Count,Iterations,DayNumber,SecsSinceMidNight from dbo.LogEntries where LogMessage='" & lLogMessage & "';",lResultAry())

    IF lRecs = -1 THEN STDOUT "LogMessage not found":EXIT FUNCTION

    lStatusBar          = AfxVarToStr(lResultAry(1,lRecs))
    lCount              = VAL(AfxVarToStr(lResultAry(2,lRecs)))
    lNumOfIterations    = VAL(AfxVarToStr(lResultAry(3,lRecs)))
    lSecsSinceLastUpDate = VAL(AfxVarToStr(lResultAry(5,lRecs)))

    lSecsADay = 86400
    lTick = 300
    lDayNumber = VAL(AfxVarToStr(lResultAry(4,lRecs)))
    lSecsDelta = ((FIX(MyTimeVar)-lSecsSinceLastUpDate) / lTick)-1
    LDayDelta =  AfxDay() - lDayNumber

    IF %DEBUG THEN
        'STDOUT "DB  DayNumber                       " & FORMAT$(lDayNumber)
        'STDOUT "Now DayNumber                       " & FORMAT$(AfxDay())
        STDOUT "Secsonds read from database         " & FORMAT$(lSecsSinceLastUpDate)
        STDOUT "Secsonds read now                   " & FORMAT$(FIX(MyTimeVar))
        STDOUT "Seconds since last update           " & FORMAT$(FIX(MyTimeVar) - lSecsSinceLastUpDate)
        'STDOUT "Any days elapsed since last read ?  " & FORMAT$(LDayDelta)
        STDOUT "How many ticks per block?             300"
        STDOUT "Any black blocks needed?             " & FORMAT$(lSecsDelta)
        STDOUT FORMAT$((86400*LDayDelta) + lSecsDelta)
        STDOUT FORMAT$(lSecsDelta)
    END IF

    'IF lSecsDelta > 0 THEN lBlancs=STRING$(FIX((MyTimeVar-lSecsSinceLastUpDate) / 300) ,"y")

    'truncate statusbar to prevent overflow
    IF LEN(lStatusBar) > (lNumOfIterations * 3) THEN lStatusBar = LEFT$(lStatusBar,(lNumOfIterations * 3))

    INCR lCount

    TsH_MSSQL_Execute(lConStr, "Update dbo.LogEntries set Status='" & lStatus & _
    "', StatusBar='" lStatus & MID$(lStatusbar,2) & _
    "', Count=" & FORMAT$(lCount) & _
    ",TimeStamp=CAST('" & Dateformat(5) &  "'  AS DATETIME),SecsSinceMidNight=" & _
    FORMAT$(FIX(MyTimeVar)) & _
    ",DayNumber=" & FORMAT$(AfxDay(),"00") & _
    " Where LogMessage='" & lLogMessage & "';")


    'EMailSender Monitor example
    IF lStatus = "c" AND lLogMessage = "HBJ_5" THEN
        LOCAL  lPriorityValue AS LONG
        LOCAL lhtmlBody,lTransportMessage AS STRING
        'profile_ID = 1 ' Citera no
        'subject
        'reciepients
        'priority = 1
        'htmlbody
        'Transportstatus = 0
        'dBMailSend = 0
        'CallInvite = 0
        lRecs = TsH_MSSQL_Select(lConstr2,"SELECT [email_ID],[transportMessage] fROM [EmailResources].[dbo].[Emails] where transportStatus =  -1 order by email_ID desc;",lResultAry())
        lTransportMessage = AfxVarToStr(lResultAry(1,lRecs))

        lHtmlBody = "<h4>Check the Monitor WEB page for <nobr style=""color:red;"">RED</nobr> alerts!</h4>The EMAILSENDER app failed to send one or more emails at " & Dateformat(5) & ".<br>" & _
        "Check the Monitor pages <a href=""https://watchdog.hmsvisjon.no"" target=""_blank"">here</a>.<p>" & _
        "Last error message: <I>" & lTransportMessage & "</I>.<br>"

        TsH_MSSQL_Execute(lConStr2,"Insert into dbo.Emails(Profile_ID,subject,recipients,priority,htmlbody,Transportstatus,DBmailSend,CalInvite,Type) " & _
        "VALUES(9,'LogFeeder - Alert!','tor@citera.no',1,'" & lHtmlBody & "',0,0,0,'LogFeeder');")

        lPriorityValue = (2 * 1) + 5
        TsH_MSSQL_WriteToSysLog(lConStr,lPriorityValue,"mail system",1,"Alert","MSSQLServer01","Logfeeder","System","EmailSender error. check dbo.Emails table for further details.")

    END IF

END FUNCTION
