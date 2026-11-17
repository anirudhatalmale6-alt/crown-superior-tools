Attribute VB_Name = "CrownAuto"
Option Explicit

' =====================================================================
' Crown Superior - CrownRunQuotes on a timer
' ---------------------------------------------------------------------
' The website decides WHEN quotes may go out (the schedule on Customer
' Quotes) and WHICH ones are waiting. But something on this PC has to
' actually go and quote them, and until now that was only ever somebody
' pressing CrownRunQuotes - which is why switching rating on and setting
' the hours never quoted anything by itself.
'
' CrownStartAutoQuoting starts a timer. Every 5 minutes it asks the
' website whether anything is waiting and whether now is inside the
' hours; if both, it runs CrownRunQuotes with no boxes and writes every
' line on RunLog. CrownStopAutoQuoting stops it.
'
' Excel has to stay open, with this workbook open, and the PC must not be
' asleep.
'
' 24 Oct: "if i never stopped it ever, why does it always need turning
' on". The on/off lived only in VBA's memory, and every Reset (installing
' an update), code edit or reopen of Excel wiped it - so the timer found
' itself "off" and quietly stopped. Now the switch is kept IN THE WORKBOOK
' (a hidden name, CrownAutoQuoting), so it survives all of those: the timer
' carries on after a Reset, picks itself up again after an update, and
' starts by itself when the workbook is opened (Auto_Open) - until
' CrownStopAutoQuoting is run. Save the workbook once after starting it
' so the switch is remembered the next time Excel is opened.
' =====================================================================

' How often it looks.
Private Const AUTO_EVERY_MINUTES As Long = 5

Private CrownAutoNextRun As Date
Public CrownAutoOn As Boolean

' 17 Nov (Tony: "interrupting the quote I was doing instead of waiting and
' it had the wrong information"): while somebody is working on the tool by
' hand, automatic quoting waits. Hand work = the Input sheet changed by
' anyone but the runner, or a quote button pressed by hand. It waits until
' nothing has changed for this long - or until staff ask for a quote on the
' website (Quote again, or a quote request sent while signed in): that is
' "go now" (Tony, 17 Nov: "15 minutes is too long").
Private Const AUTO_HOLD_MINUTES As Long = 10
Private CrownAutoPauseUntil As Date
Private CrownAutoPauseFrom As Date
Private CrownAutoLastPrint As String

' The switch as kept in the workbook - survives Reset and reopening.
Private Const AUTO_NAME As String = "CrownAutoQuoting"


Public Sub CrownStartAutoQuoting()
    If CrownAutoOn Then
        MsgBox "Automatic quoting is already on. Next look at " & _
               Format$(CrownAutoNextRun, "h:nn AM/PM") & ".", vbInformation, "Crown Superior"

        Exit Sub
    End If

    CrownAutoOn = True
    CrownAutoRemember True
    Randomize
    CrownAutoNote "automatic quoting started - looking every " & AUTO_EVERY_MINUTES & " minutes"

    ' The first look straight away (in 10 seconds), not in 5 minutes.
    CrownAutoSchedule 0

    MsgBox "Automatic quoting is ON." & vbCrLf & vbCrLf & _
           "Every " & AUTO_EVERY_MINUTES & " minutes it checks the website. When quotes are " & _
           "waiting and it is inside your schedule, it quotes them through United and Verve " & _
           "and sends the answers back - no boxes to press." & vbCrLf & vbCrLf & _
           "Leave Excel open with this workbook. Everything it does is written on the " & _
           "RunLog sheet." & vbCrLf & vbCrLf & _
           "It stays on - after updates, and it starts by itself when you open this " & _
           "workbook. Save the workbook now (Ctrl+S) so it remembers." & vbCrLf & vbCrLf & _
           "To stop: Alt+F8 > CrownStopAutoQuoting > Run.", vbInformation, "Crown Superior"
End Sub


Public Sub CrownStopAutoQuoting()
    CrownAutoOn = False
    CrownAutoRemember False
    CrownAutoSetChain "stopped"     ' any booking still waiting ends itself

    CrownAutoNote "automatic quoting stopped"

    MsgBox "Automatic quoting is OFF. Nothing will be quoted until you run " & _
           "CrownRunQuotes or CrownStartAutoQuoting again.", vbInformation, "Crown Superior"
End Sub


' Stop the timer without a box - CrownInstallUpdate calls this before it
' replaces this module. The switch in the workbook is LEFT AS IT IS, and
' a look is booked for 30 seconds on: by then the new code is in, and if
' automatic quoting was on it simply carries on.
Public Sub CrownAutoQuietStop()
    CrownAutoOn = False
    CrownAutoSetChain "updating"    ' any booking still waiting ends itself

    On Error Resume Next
    Application.OnTime Now + TimeSerial(0, 0, 30), "CrownAutoResume"
    Application.OnTime Now + TimeSerial(0, 0, 40), "CrownPayResume"   ' the payment reports' timer too (3 Nov)
    On Error GoTo 0
End Sub


' Opening the workbook: back on by itself if it was on when last saved.
' The payment reports' timer (CrownPaySchedule) is booked again here too.
Public Sub Auto_Open()
    If CrownAutoRemembered() Then
        On Error Resume Next
        Application.OnTime Now + TimeSerial(0, 0, 30), "CrownAutoResume"
        On Error GoTo 0
    End If

    On Error Resume Next
    Application.OnTime Now + TimeSerial(0, 0, 40), "CrownPayResume"
    On Error GoTo 0
End Sub


' Picks the timer up again, quietly, if the switch in the workbook is on.
Public Sub CrownAutoResume()
    If Not CrownAutoRemembered() Or CrownAutoOn Then Exit Sub

    Randomize

    CrownAutoOn = True
    CrownAutoNote "automatic quoting carried on by itself (after an update, a reset or opening the workbook)"
    CrownAutoSchedule 0
End Sub


' What the timer runs. Public only because Excel's timer can only call a
' public macro - it is not meant to be run by hand.
Public Sub CrownAutoTick(Optional ByVal chain As String = "")
    ' Only the most recently booked timer may run. After a Reset an older
    ' booking can still be pending in Excel, and two would quote everything
    ' twice - so a tick whose number is not the current one just ends.
    If chain <> CrownAutoChain() Then Exit Sub

    ' After a Reset the memory says "off" but the workbook still says "on".
    If Not CrownAutoOn Then
        If Not CrownAutoRemembered() Then Exit Sub
        CrownAutoOn = True
    End If

    ' Somebody is working on the tool by hand: wait, and say so.
    If CrownAutoHandWork() Then
        On Error Resume Next
        CrownPing                               ' the website still sees the tool is on
        Err.Clear
        On Error GoTo 0

        CrownAutoNote "waiting - the tool is being used by hand (Input sheet changed or a quote button pressed); " & _
                      "carries on " & AUTO_HOLD_MINUTES & " minutes after the last change, from " & Format$(CrownAutoPauseUntil, "h:nn AM/PM")
        If CrownAutoOn Then CrownAutoSchedule AUTO_EVERY_MINUTES
        Exit Sub
    End If

    CrownAutoMode = True

    On Error Resume Next
    CrownRunQuotes
    If Err.Number <> 0 Then CrownAutoNote "the run stopped with an error: " & Err.Description
    Err.Clear
    On Error GoTo 0

    CrownAutoMode = False
    CrownUnattended = False


    ' A payment run that was due and never started is started now (3 Nov).
    On Error Resume Next
    CrownPayWatch
    Err.Clear
    On Error GoTo 0

    ' What the runs left on the Input sheet is not hand work.
    CrownAutoLastPrint = CrownInputPrint()

    ' The next one is booked after this one has finished, so a long batch
    ' can never have a second run start on top of it.
    If CrownAutoOn Then CrownAutoSchedule AUTO_EVERY_MINUTES
End Sub


' A quote button pressed by hand: automatic quoting waits AUTO_HOLD_MINUTES.
Public Sub CrownAutoHoldOff()
    If CrownAutoMode Then Exit Sub              ' the runner itself pressing it
    CrownAutoPauseFrom = Now
    CrownAutoPauseUntil = Now + AUTO_HOLD_MINUTES / 1440#
End Sub

' True while hand work is going on: the Input sheet is not as the runner
' (or the last look) left it, or a quote button was pressed by hand lately.
Private Function CrownAutoHandWork() As Boolean
    Dim fp As String

    fp = CrownInputPrint()

    If Len(CrownAutoLastPrint) = 0 Then
        CrownAutoLastPrint = fp             ' first look after starting: nothing to compare with
    ElseIf fp <> CrownAutoLastPrint Then
        CrownAutoLastPrint = fp
        CrownAutoPauseFrom = Now
        CrownAutoPauseUntil = Now + AUTO_HOLD_MINUTES / 1440#
    End If

    If Now >= CrownAutoPauseUntil Then Exit Function

    ' Staff asked for a quote on the website since the wait began: go now.
    If CrownAutoStaffAskedSince(CrownAutoPauseFrom) Then
        CrownAutoPauseUntil = 0
        CrownAutoNote "a quote was asked for on the website (Quote again / sent by staff) - not waiting any longer"
        Exit Function
    End If

    CrownAutoHandWork = True
End Function

' Did staff press Quote again, or send a quote request while signed in,
' after this moment? The website answers how many seconds ago that was.
Private Function CrownAutoStaffAskedSince(ByVal since As Date) As Boolean
    Dim answer As String, at As Long, age As Double

    On Error Resume Next
    answer = CrownPost(CrownBody("staffasked"))
    On Error GoTo 0

    at = InStr(1, answer, """age"":", vbTextCompare)
    If at = 0 Then Exit Function

    age = Val(Mid$(answer, at + 6))
    If age < 0 Then Exit Function

    CrownAutoStaffAskedSince = (age < (Now - since) * 86400# - 5)
End Function

' Everything typed on the Input sheet's working area, as one string.
Private Function CrownInputPrint() As String
    Dim v As Variant, r As Long, c As Long, out As String

    On Error Resume Next
    v = ThisWorkbook.Worksheets("Input").Range("A1:AD80").Formula   ' formulas as written, so a =TODAY() or =NOW() cell never looks like hand work
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0

    For r = 1 To UBound(v, 1)
        For c = 1 To UBound(v, 2)
            If Not IsError(v(r, c)) Then
                If Len(CStr(v(r, c))) > 0 Then out = out & r & "," & c & "=" & CStr(v(r, c)) & "|"
            End If
        Next c
    Next r

    CrownInputPrint = out
End Function


Private Sub CrownAutoSchedule(ByVal minutes As Long)
    Dim chain As String

    If minutes <= 0 Then
        CrownAutoNextRun = Now + TimeSerial(0, 0, 10)
    Else
        CrownAutoNextRun = Now + TimeSerial(0, minutes, 0)
    End If

    ' A new number for this booking; any older booking still waiting in
    ' Excel no longer matches and will end itself.
    chain = CStr(CLng(Rnd() * 1000000) + 1)
    CrownAutoSetChain chain

    Application.OnTime CrownAutoNextRun, "'CrownAutoTick """ & chain & """'"
End Sub


Private Sub CrownAutoSetChain(ByVal chain As String)
    On Error Resume Next
    ThisWorkbook.Names("CrownAutoChain").Delete
    ThisWorkbook.Names.Add Name:="CrownAutoChain", RefersTo:="=""" & chain & """", Visible:=False
    On Error GoTo 0
End Sub

Private Function CrownAutoChain() As String
    Dim v As String

    On Error Resume Next
    v = ThisWorkbook.Names("CrownAutoChain").RefersTo
    On Error GoTo 0

    CrownAutoChain = Replace(Replace(v, "=", ""), """", "")
End Function


' The switch, kept as a hidden name in this workbook.
Private Sub CrownAutoRemember(ByVal isOn As Boolean)
    On Error Resume Next
    ThisWorkbook.Names(AUTO_NAME).Delete
    ThisWorkbook.Names.Add Name:=AUTO_NAME, RefersTo:=IIf(isOn, "=TRUE", "=FALSE"), Visible:=False
    On Error GoTo 0
End Sub

Private Function CrownAutoRemembered() As Boolean
    Dim v As String

    On Error Resume Next
    v = ThisWorkbook.Names(AUTO_NAME).RefersTo
    On Error GoTo 0

    CrownAutoRemembered = (UCase$(v) = "=TRUE")
End Function
