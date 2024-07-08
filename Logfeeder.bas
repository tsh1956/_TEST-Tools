#COMPILE EXE
#DIM ALL
'%SupressSQLErrors=-1 'No error reporting from MSSQL on errors. Remove this constant to display errors.
%DEBUG=0
#INCLUDE "tsh_MSSQL.INC"




FUNCTION PBMAIN () AS LONG

    'LOCAL Built AS IPowerTime

    'LET Built = CLASS "PowerTime"

    LOCAL lIn_1,lIn_2,lIn_3,lIn_4,lIn_5 AS STRING
    LOCAL lStatus,lStatusBar AS STRING
    RESET lStatus,lStatusBar

    LOCAL lID AS LONG
    lIn_1 = COMMAND$(1) 'id
    lIn_2 = COMMAND$(2) 'statusflags
    lIn_3 = COMMAND$(3) 'email addresse(s) for alerts
    lIn_4 = COMMAND$(4) '
    lIn_5 = COMMAND$(5)

    IF LEN(lIn_1) <> 0 THEN
        lId = VAL(lIn_1)
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
    #INCLUDE "tsh_ConStrs.inc"

    DIM lResultAry() AS VARIANT

'---------static begin
    LOCAL lNumOfIterations,lSecsSinceMidNight,lDayNumber,lSecsDelta,lDayDelta,lSecsADay AS LONG
    LOCAL lStatusBarInternal,lBlancs AS STRING

    RESET lBlancs

    'lNumOfIterations = 288
    'lStatusBarInternal = STRING$(lNumOfIterations,"x")
'---------static stop

    LOCAL mytimeVar AS DOUBLE
    mytimeVar = TIMER


    lRecs = TsH_MSSQL_Select(lConstr, "Select Status,StatusBar,Count,Iterations,DayNumber,SecsSinceMidNight from dbo.LogEntries where ID=" & FORMAT$(lID) & ";",lResultAry())

    IF lRecs = -1 THEN STDOUT "ID not found":EXIT FUNCTION

    lStatusBar = AfxVarToStr(lResultAry(1,lRecs))
    lCount = VAL(AfxVarToStr(lResultAry(2,lRecs)))
    lNumOfIterations = VAL(AfxVarToStr(lResultAry(3,lRecs)))

    lSecsADay = 86400

    lDayNumber = VAL(AfxVarToStr(lResultAry(4,lRecs)))
    lSecsSinceMidNight = VAL(AfxVarToStr(lResultAry(5,lRecs)))
    lSecsDelta = (FIX(MyTimeVar) - lSecsSinceMidNight) / (lSecsADay/lNumOfIterations) -1
    LDayDelta =  AfxDay() - lDayNumber

    IF %DEBUG THEN
        STDOUT "DB  DayNumber     " & FORMAT$(lDayNumber)
        STDOUT "Now DayNumber     " & FORMAT$(AfxDay())
        STDOUT "DB  Secs since MN " & FORMAT$(lSecsSinceMidNight)
        STDOUT "Now Secs since MN " & FORMAT$(FIX(MyTimeVar))
        STDOUT "Seconds elapsed   " & FORMAT$(FIX(MyTimeVar) - lSecsSinceMidNight)
        STDOUT FORMAT$(LDayDelta)
        STDOUT FORMAT$((86400*LDayDelta) + lSecsDelta)
    END IF

    IF lSecsDelta > 0 THEN lBlancs=STRING$((lSecsADay * LDayDelta) + lSecsDelta,"y")

    'truncate statusbar to prevent overflow
    IF LEN(lStatusBar) > (lNumOfIterations * 3) THEN lStatusBar = LEFT$(lStatusBar,(lNumOfIterations * 3))

    INCR lCount

    TsH_MSSQL_Execute(lConStr, "Update dbo.LogEntries set Status='" & lStatus & _
    "', StatusBar='" & LEFT$(lStatus + lBlancs + lStatusbar,lNumOfIterations) & _
    "', Count=" & FORMAT$(lCount) & _
    ",TimeStamp=CAST('" & Dateformat(5) &  "'  AS DATETIME),SecsSinceMidNight=" & _
    FORMAT$(FIX(MyTimeVar)) & _
    ",DayNumber=" & FORMAT$(AfxDay(),"00") & _
    " Where ID=" & FORMAT$(lID) & ";")

    IF lStatus = "c" AND lID=7 THEN
        LOCAL  lPriorityValue AS LONG
        LOCAL lhtmlBody AS STRING
        'profile_ID = 1 ' Citera no
        'subject
        'reciepients
        'priority = 1
        'htmlbody
        'Transportstatus = 0
        'dBMailSend = 0
        'CallInvite = 0
        lHtmlBody = "<h4>Check the Monitor WEB page for <nobr style=""color:red;"">RED</nobr> alerts!</h4>The EMAILSENDER app has failed to send one or more emails.<br>" & _
        "Detailed status can be found here:<br>The database <b>EmailResources</b> on <b>SQLServer01</b> in the table <b>dbo.Emails</b>.<br>Select the latest records with <b>transportstatus=-1</b> and check the <b>transportMessage</b>.<br>"

        TsH_MSSQL_Execute(lConStr2,"Insert into dbo.Emails(Profile_ID,subject,recipients,priority,htmlbody,Transportstatus,DBmailSend,CalInvite,Type) " & _
        "VALUES(1,'LogFeeder Red Alert!','tor@citera.no',1,'" & lHtmlBody & "',0,0,0,'LOGFEEDER');")

        lPriorityValue = (2 * 1) + 5
        TsH_MSSQL_WriteToSysLog(lConStr,lPriorityValue,"mail system",1,"Alert","MSSQLServer01","Logfeeder","System","EmailSender error. check dbo.Emails table for further details.")

    END IF

END FUNCTION
