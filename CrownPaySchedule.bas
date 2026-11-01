Attribute VB_Name = "CrownPaySchedule"
Option Explicit

' =====================================================================
' Crown Superior - the payment reports, one after another, on a timer
' ---------------------------------------------------------------------
' Tony, 1 Nov: "I'd like for the payment reports to run consecutively
' automatically for each company - that's how the old system worked -
' starting with United ... on a timer ... at 6 am each day Monday through
' Sunday ... and to be able to set it to run multiple times per day".
'
'   CrownRunAllPayments        every company ticked on PaymentSchedule, now
'   CrownStartPaymentSchedule  on: runs at the times and days on PaymentSchedule
'   CrownStopPaymentSchedule   off
'
' The PaymentSchedule sheet (made the first time any of these is run) holds
' the times, the days, and the companies in the order they run - move a
' row to change the order, No to leave one out. Everything a run would have
' said in a box goes on the PaymentLog sheet instead, so nothing waits for
' a click at 6 in the morning. A company that needs a person (a sign-in
' code, a manual sign-in) is skipped with "needs you" on the log, and the
' run carries on with the next one.
'
' Like automatic quoting, the switch is kept in the workbook, so it carries
' on after an update and when the workbook is opened again. Excel never
' runs two macros at once, so a payment run and a quote run cannot overlap:
' whichever is due second starts when the first has finished.
' =====================================================================

Public CrownPayUnattended As Boolean
Public CrownPayRunning As Boolean
Public CrownPayCompany As String
Public CrownPayLastSaid As String

Private Const SCHED_SHEET As String = "PaymentSchedule"
Private Const LOG_SHEET As String = "PaymentLog"
Private Const PAY_NAME As String = "CrownPaySchedule"
Private Const PAY_CHAIN As String = "CrownPayChain"

' Row layout of PaymentSchedule.
Private Const ROW_TIMES As Long = 4
Private Const ROW_DAYS_HEAD As Long = 7
Private Const ROW_DAYS As Long = 8
Private Const ROW_COMPANY_HEAD As Long = 11
Private Const ROW_FIRST_COMPANY As Long = 12


' ---------------------------------------------------------------------
' The three macros
' ---------------------------------------------------------------------

Public Sub CrownRunAllPayments()
    Dim summary As String

    CrownPaySheetReady

    summary = CrownPayRunAll("run by hand")

    MsgBox "Payment reports finished." & vbCrLf & vbCrLf & summary & vbCrLf & vbCrLf & _
           "The PaymentLog sheet has every line.", vbInformation, "Crown Superior"
End Sub


Public Sub CrownStartPaymentSchedule()
    Dim nextRun As Date

    CrownPaySheetReady
    CrownPayRemember True

    nextRun = CrownPayBook()

    If nextRun = 0 Then
        MsgBox "The schedule is ON, but there is no time or no day ticked on the PaymentSchedule sheet, " & _
               "so nothing is booked. Fill in a time (like 6:00 AM) and Yes under the days, then run " & _
               "CrownStartPaymentSchedule again.", vbExclamation, "Crown Superior"
        Exit Sub
    End If

    CrownPayLog "", "schedule switched on - next run " & Format$(nextRun, "ddd d mmm h:nn AM/PM")

    MsgBox "Payment reports are on a timer." & vbCrLf & vbCrLf & _
           "Next run: " & Format$(nextRun, "dddd d mmm, h:nn AM/PM") & vbCrLf & vbCrLf & _
           "Times, days and companies are on the PaymentSchedule sheet - change them there " & _
           "(no need to restart it). Leave Excel open with this workbook, and save it now " & _
           "(Ctrl+S) so the timer is remembered." & vbCrLf & vbCrLf & _
           "To stop: Alt+F8 > CrownStopPaymentSchedule > Run.", vbInformation, "Crown Superior"
End Sub


Public Sub CrownStopPaymentSchedule()
    CrownPayRemember False
    CrownPaySetChain "stopped"
    CrownPayLog "", "schedule switched off"

    MsgBox "Payment reports are no longer on a timer. CrownRunAllPayments still runs them " & _
           "all by hand.", vbInformation, "Crown Superior"
End Sub


' ---------------------------------------------------------------------
' Timer
' ---------------------------------------------------------------------

' Opening the workbook (from Auto_Open) or after an update: book the next
' run again if the schedule was on.
Public Sub CrownPayResume()
    If Not CrownPayRemembered() Then Exit Sub
    If CrownPayBook() > 0 Then CrownPayLog "", "schedule carried on (workbook opened or updated)"
End Sub


' What the timer runs - not meant to be run by hand.
Public Sub CrownPayTick(Optional ByVal chain As String = "")
    If chain <> CrownPayChainNow() Then Exit Sub
    If Not CrownPayRemembered() Then Exit Sub

    On Error Resume Next
    CrownPayRunAll "on the timer"
    On Error GoTo 0

    CrownPayBook
End Sub


' Books the next time on the sheet. Returns it, or 0 when there is none.
Private Function CrownPayBook() As Date
    Dim nextRun As Date, chain As String

    nextRun = CrownPayNextTime()
    If nextRun = 0 Then Exit Function

    Randomize
    chain = CStr(CLng(Rnd() * 1000000) + 1)
    CrownPaySetChain chain

    On Error Resume Next
    Application.OnTime nextRun, "'CrownPayTick """ & chain & """'"
    On Error GoTo 0

    CrownPayBook = nextRun
End Function


' The next run after now: the earliest listed time on a ticked day.
Private Function CrownPayNextTime() As Date
    Dim ws As Worksheet
    Dim times As Collection, t As Variant
    Dim d As Long, c As Long, when As Date, best As Date
    Dim dayOn(1 To 7) As Boolean, anyDay As Boolean

    Set ws = CrownPaySheet()
    If ws Is Nothing Then Exit Function

    ' Columns B-H = Monday-Sunday.
    For c = 1 To 7
        dayOn(c) = CrownPayYes(ws.Cells(ROW_DAYS, c + 1).Value)
        If dayOn(c) Then anyDay = True
    Next c

    If Not anyDay Then Exit Function

    Set times = New Collection

    For c = 2 To 13
        If Len(Trim$(CStr(ws.Cells(ROW_TIMES, c).Value))) > 0 Then
            t = CrownPayTimeOf(ws.Cells(ROW_TIMES, c).Value)
            If t >= 0 Then times.Add t
        End If
    Next c

    If times.Count = 0 Then Exit Function

    For d = 0 To 7
        If dayOn(Weekday(Date + d, vbMonday)) Then
            For Each t In times
                when = Date + d + CDbl(t)

                If when > Now + TimeSerial(0, 0, 20) Then
                    If best = 0 Or when < best Then best = when
                End If
            Next t
        End If
    Next d

    CrownPayNextTime = best
End Function


' A time of day from a cell: a real Excel time, or text like "6:00 AM", "18:30", "6pm".
' -1 when it cannot be read.
Private Function CrownPayTimeOf(ByVal v As Variant) As Double
    Dim s As String

    CrownPayTimeOf = -1

    If IsDate(v) Then
        CrownPayTimeOf = CDbl(TimeValue(CDate(v)))
        Exit Function
    End If

    If IsNumeric(v) Then
        If CDbl(v) >= 0 And CDbl(v) < 1 Then CrownPayTimeOf = CDbl(v)
        Exit Function
    End If

    s = UCase$(Replace(Trim$(CStr(v)), ".", ""))
    s = Replace(Replace(s, "AM", " AM"), "PM", " PM")
    If InStr(s, ":") = 0 Then s = Replace(Replace(s, " AM", ":00 AM"), " PM", ":00 PM")

    On Error Resume Next
    CrownPayTimeOf = CDbl(TimeValue(s))
    If Err.Number <> 0 Then CrownPayTimeOf = -1
    On Error GoTo 0
End Function


' ---------------------------------------------------------------------
' One run: every company ticked Yes, in the order on the sheet
' ---------------------------------------------------------------------

Private Function CrownPayRunAll(ByVal how As String) As String
    Dim ws As Worksheet
    Dim r As Long, company As String, macroName As String
    Dim started As Date, summary As String

    Set ws = CrownPaySheet()
    If ws Is Nothing Then Exit Function

    If CrownPayRunning Then Exit Function

    CrownPayRunning = True
    CrownPayUnattended = True

    CrownPayLog "", "payment reports started (" & how & ")"

    r = ROW_FIRST_COMPANY

    Do While Len(Trim$(CStr(ws.Cells(r, 1).Value))) > 0
        company = Trim$(CStr(ws.Cells(r, 1).Value))
        macroName = Trim$(CStr(ws.Cells(r, 3).Value))

        If CrownPayYes(ws.Cells(r, 2).Value) And Len(macroName) > 0 Then
            CrownPayCompany = company
            CrownPayLastSaid = ""
            started = Now
            CrownPayLog company, "started"

            Err.Clear
            On Error Resume Next
            Application.Run macroName
            If Err.Number <> 0 Then CrownPayLog company, "stopped with an error: " & Err.Description
            Err.Clear
            On Error GoTo 0

            CrownPayLog company, "finished in " & Format$((Now - started) * 1440, "0") & " min"
            summary = summary & company & ": " & IIf(Len(CrownPayLastSaid) > 0, CrownPayFirstLine(CrownPayLastSaid), "done") & vbCrLf

            ' A breath between companies, so the last Chrome window has closed.
            Application.Wait Now + TimeSerial(0, 0, 3)
        End If

        r = r + 1
    Loop

    CrownPayCompany = ""
    CrownPayUnattended = False
    CrownPayRunning = False

    CrownPayLog "", "payment reports finished"
    CrownPayRunAll = summary
End Function


' ---------------------------------------------------------------------
' Boxes: shown as usual by hand, written on PaymentLog when unattended
' ---------------------------------------------------------------------

' MsgBox for the payment macros. answerWhenAlone is what an unattended run
' answers (0 = No for a Yes/No question, Cancel for OK/Cancel, OK otherwise).
Public Function CrownPayBox(ByVal prompt As String, Optional ByVal buttons As VbMsgBoxStyle = vbOKOnly, _
                            Optional ByVal title As String = "Crown Superior", _
                            Optional ByVal answerWhenAlone As VbMsgBoxResult = 0) As VbMsgBoxResult
    Dim kind As Long

    If Not CrownPayUnattended Then
        CrownPayBox = MsgBox(prompt, buttons, title)
        Exit Function
    End If

    kind = buttons And 7

    If answerWhenAlone = 0 Then
        Select Case kind
            Case vbYesNo, vbYesNoCancel: answerWhenAlone = IIf(kind = vbYesNoCancel, vbCancel, vbNo)
            Case vbOKCancel, vbRetryCancel: answerWhenAlone = vbCancel
            Case Else: answerWhenAlone = vbOK
        End Select
    End If

    CrownPayLastSaid = prompt
    CrownPayLog CrownPayCompany, Replace(Replace(prompt, vbCrLf, " "), vbLf, " ") & _
                IIf(kind = vbOKOnly, "", "  [answered: " & CrownPayAnswerWord(answerWhenAlone) & "]")

    CrownPayBox = answerWhenAlone
End Function


' InputBox for the payment macros: nobody to type a code when unattended.
Public Function CrownPayInput(ByVal prompt As String, Optional ByVal title As String = "Crown Superior") As String
    If Not CrownPayUnattended Then
        CrownPayInput = InputBox(prompt, title)
        Exit Function
    End If

    CrownPayLog CrownPayCompany, "NEEDS YOU - asked for a code, nobody here to type it, so this company was skipped: " & _
                Left$(Replace(prompt, vbCrLf, " "), 160)
End Function


Private Function CrownPayAnswerWord(ByVal a As VbMsgBoxResult) As String
    Select Case a
        Case vbYes: CrownPayAnswerWord = "Yes"
        Case vbNo: CrownPayAnswerWord = "No"
        Case vbCancel: CrownPayAnswerWord = "Cancel - this company skipped"
        Case Else: CrownPayAnswerWord = "OK"
    End Select
End Function


Private Function CrownPayFirstLine(ByVal s As String) As String
    s = Replace(s, vbCrLf, vbLf)
    CrownPayFirstLine = Left$(Split(s & vbLf, vbLf)(0), 120)
End Function


' ---------------------------------------------------------------------
' Sheets
' ---------------------------------------------------------------------

Public Sub CrownPayLog(ByVal company As String, ByVal what As String)
    Dim ws As Worksheet, r As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(LOG_SHEET)

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = LOG_SHEET
        ws.Range("A1:C1").Value = Array("When", "Company", "What happened")
        ws.Range("A1:C1").Font.Bold = True
        ws.Columns(1).ColumnWidth = 20
        ws.Columns(2).ColumnWidth = 22
        ws.Columns(3).ColumnWidth = 120
    End If

    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ' Only the last few thousand lines matter.
    If r > 5000 Then
        ws.Rows("2:1001").Delete
        r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    End If

    ws.Cells(r, 1).Value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    ws.Cells(r, 2).Value = company
    ws.Cells(r, 3).Value = Left$(what, 1000)
    On Error GoTo 0
End Sub


Private Function CrownPaySheet() As Worksheet
    On Error Resume Next
    Set CrownPaySheet = ThisWorkbook.Worksheets(SCHED_SHEET)
    On Error GoTo 0
End Function


' Makes PaymentSchedule the first time, filled in the way Tony asked for:
' 6:00 AM, every day, United first.
Private Sub CrownPaySheetReady()
    Dim ws As Worksheet, c As Long, i As Long
    Dim companies As Variant, macros As Variant

    If Not CrownPaySheet() Is Nothing Then Exit Sub

    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = SCHED_SHEET

    ws.Range("A1").Value = "Payment reports - when they run, and which companies"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14

    ws.Range("A3").Value = "Times of day - one per box, like 6:00 AM or 6:00 PM (up to 12)"
    ws.Range("A" & ROW_TIMES).Value = "Run at:"
    ws.Range("B" & ROW_TIMES).Value = "6:00 AM"

    ws.Range("A6").Value = "Days - Yes or No"
    ws.Range("A" & ROW_DAYS_HEAD).Value = "Day:"
    ws.Range("B" & ROW_DAYS_HEAD & ":H" & ROW_DAYS_HEAD).Value = Array("Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday")
    ws.Range("A" & ROW_DAYS).Value = "Run?"

    For c = 2 To 8
        ws.Cells(ROW_DAYS, c).Value = "Yes"
    Next c

    ws.Range("A10").Value = "Companies - they run in this order, top to bottom. No = left out. (Move a row to change the order; do not change the Macro column.)"
    ws.Range("A" & ROW_COMPANY_HEAD & ":C" & ROW_COMPANY_HEAD).Value = Array("Company", "Run?", "Macro")

    companies = Array("United Auto", "Trisura (Verve)", "Assurance America", "Trexis", "Progressive", "First Acceptance", "National General")
    macros = Array("CrownUaigPaymentDue", "CrownVervePaymentDue", "CrownAssurancePaymentDue", "CrownTrexisPaymentDue", _
                   "CrownProgressivePaymentDue", "CrownFirstAcceptancePaymentDue", "CrownNatGenPaymentDue")

    For i = 0 To UBound(companies)
        ws.Cells(ROW_FIRST_COMPANY + i, 1).Value = companies(i)
        ws.Cells(ROW_FIRST_COMPANY + i, 2).Value = "Yes"
        ws.Cells(ROW_FIRST_COMPANY + i, 3).Value = macros(i)
    Next i

    ws.Range("A3,A6,A10").Font.Italic = True
    ws.Range("A" & ROW_TIMES & ",A" & ROW_DAYS_HEAD & ",A" & ROW_DAYS & ",A" & ROW_COMPANY_HEAD & ":C" & ROW_COMPANY_HEAD & ",B" & ROW_DAYS_HEAD & ":H" & ROW_DAYS_HEAD).Font.Bold = True
    ws.Columns(1).ColumnWidth = 22
    ws.Columns("B:H").ColumnWidth = 12
    ws.Columns(3).ColumnWidth = 30
End Sub


Private Function CrownPayYes(ByVal v As Variant) As Boolean
    Dim s As String

    s = UCase$(Trim$(CStr(v)))
    CrownPayYes = (s = "YES" Or s = "Y" Or s = "X" Or s = "TRUE" Or s = "1")
End Function


' The on/off switch and the booking number, kept as hidden names in the workbook.
Private Sub CrownPayRemember(ByVal isOn As Boolean)
    On Error Resume Next
    ThisWorkbook.Names(PAY_NAME).Delete
    ThisWorkbook.Names.Add Name:=PAY_NAME, RefersTo:=IIf(isOn, "=TRUE", "=FALSE"), Visible:=False
    On Error GoTo 0
End Sub

Private Function CrownPayRemembered() As Boolean
    Dim v As String

    On Error Resume Next
    v = ThisWorkbook.Names(PAY_NAME).RefersTo
    On Error GoTo 0

    CrownPayRemembered = (UCase$(v) = "=TRUE")
End Function

Private Sub CrownPaySetChain(ByVal chain As String)
    On Error Resume Next
    ThisWorkbook.Names(PAY_CHAIN).Delete
    ThisWorkbook.Names.Add Name:=PAY_CHAIN, RefersTo:="=""" & chain & """", Visible:=False
    On Error GoTo 0
End Sub

Private Function CrownPayChainNow() As String
    Dim v As String

    On Error Resume Next
    v = ThisWorkbook.Names(PAY_CHAIN).RefersTo
    On Error GoTo 0

    CrownPayChainNow = Replace(Replace(v, "=", ""), """", "")
End Function
