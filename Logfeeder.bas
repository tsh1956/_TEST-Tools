#COMPILE EXE
#DIM ALL
%SupressSQLErrors=-1 'No error reporting from MSSQL on errors. Remove this constant to display errors.
#INCLUDE "tsh_MSSQL.INC"

FUNCTION PBMAIN () AS LONG

    'LOCAL Built AS IPowerTime

    'LET Built = CLASS "PowerTime"

    LOCAL lIn_1,lIn_2,lIn_3,lIn_4,lIn_5 AS STRING
    LOCAL lStatus,lStatusBar AS STRING
    RESET lStatus,lStatusBar

    LOCAL lID AS LONG
    lIn_1 = COMMAND$(1)
    lIn_2 = COMMAND$(2)
    lIn_3 = COMMAND$(3)
    lIn_4 = COMMAND$(4)
    lIn_5 = COMMAND$(5)

    IF LEN(lIn_1) <> 0 THEN
        lId = VAL(lIn_1)
    ELSE
        STDOUT "no ID":EXIT FUNCTION
    END IF

    IF LEN(lIn_2) = 1 THEN
        IF TALLY(LCASE$(lIn_2),ANY "abcdexy") > 0 THEN lStatus = lIn_2
    ELSE
        lStatus = ""
        STDOUT "no status flag":EXIT FUNCTION
    END IF

    LOCAL lConStr AS STRING
    lConStr = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=LogServices;Data Source=SQLSERVER-01\DEV01"
    'lConStr = "Provider=SQLOLEDB.1;Data Source=10.10.90.6;initial catalog=LogServices;User ID=sa;Password=Yamt55fA;Encrypt=False"
    lConStr = "Provider=SQLOLEDB.1;Data Source=ubuntu14;initial catalog=LogServices;User ID=sa;Password=Yamt55fA;Encrypt=False"
    'lConStr = "Provider=SQLOLEDB.1;Persist Security Info=True;;User ID=sa;Password=**nf_991**;Initial Catalog=LogServices;Data Source=4.180.32.85\DEV01,62022"
    DIM lStatAry(1 TO 3) AS STRING
    lStatAry(1) = "a"
    lStatAry(2) = "b"
    lStatAry(3) = "c"

    LOCAL lRecs,lCount,lRnd AS LONG

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

'    IF lStatus = "" THEN
'        RANDOMIZE TIMER
'        lRnd = RND(1,3)
'        lStatus = lStatAry(lRnd)
'    END IF

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

    STDOUT "DB  DayNumber     " & FORMAT$(lDayNumber)
    STDOUT "Now DayNumber     " & FORMAT$(AfxDay())
    STDOUT "DB  Secs since MN " & FORMAT$(lSecsSinceMidNight)
    STDOUT "Now Secs since MN " & FORMAT$(FIX(MyTimeVar))
    STDOUT "Seconds elapsed   " & FORMAT$(FIX(MyTimeVar) - lSecsSinceMidNight)
    STDOUT FORMAT$(LDayDelta)
    STDOUT FORMAT$((86400*LDayDelta) + lSecsDelta)


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

    TsH_MSSQL_WriteToSysLog(lConStr,14,"user-level messages ",6,"Informational","localhost","Logfeeder","TsH","TestMelding")

END FUNCTION
