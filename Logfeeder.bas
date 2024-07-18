#COMPILE EXE
#DIM ALL
'%SupressSQLErrors=-1 'No error reporting from MSSQL on errors. Remove this constant to display errors.
%DEBUG=-1
#INCLUDE "tsh_MSSQL.INC"

FUNCTION PBMAIN () AS LONG

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

    LOCAL lNumOfIterations,lSecsSinceLastUpDate,lDayNumber,lSecsDelta,lStatusBarLength,lVisible AS LONG
    LOCAL lStatusBarInternal,lBlancs AS STRING

    RESET lBlancs

    LOCAL lMyTimeVar AS DOUBLE
    lMyTimeVar = TIMER

    lRecs = TsH_MSSQL_Select(lConstr, "Select Status,StatusBar,Count,Iterations,DayNumber,SecsSinceMidNight,Visible from dbo.LogEntries where LogMessage='" & lLogMessage & "';",lResultAry())

    IF lRecs = -1 THEN STDOUT "LogMessage not found":EXIT FUNCTION

    lStatusBar          = AfxVarToStr(lResultAry(1,lRecs))
    lCount              = VAL(AfxVarToStr(lResultAry(2,lRecs)))
    lNumOfIterations    = VAL(AfxVarToStr(lResultAry(3,lRecs)))
    lVisible            = VAL(AfxVarToStr(lResultAry(6,lRecs)))

    'Make status visible
    IF TALLY(lStatus,ANY "bcdex") > 0 THEN lVisible = 144

    'Secure that statusbar has AT LEAST as long as lNumOfIterations, adjust if not
    lStatusBarLength = LEN(lStatusBar)
    IF lStatusBarLength < lNumOfIterations THEN
        lStatusBar += STRING$((lNumOfIterations-lStatusBarLength),"y")
    END IF

    'truncate statusbar to prevent overflow
    IF LEN(lStatusBar) > (lNumOfIterations * 3) THEN lStatusBar = LEFT$(lStatusBar,(lNumOfIterations * 3))

    INCR lCount

    TsH_MSSQL_Execute(lConStr, "Update dbo.LogEntries set Status='" & lStatus & _
    "', StatusBar='" & lStatus & MID$(lStatusbar,2) & _
    "', Count=" & FORMAT$(lCount) & _
    ",TimeStamp=CAST('" & Dateformat(5) &  "'  AS DATETIME),SecsSinceMidNight=" & _
    FORMAT$(FIX(lMyTimeVar)) & _
    ",DayNumber=" & FORMAT$(AfxDay(),"00") & _
    ",Visible=" & FORMAT$(lVisible) & _
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
