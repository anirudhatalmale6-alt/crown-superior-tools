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
    On Error GoTo 0
End Sub


' Opening the workbook: back on by itself if it was on when last saved.
Public Sub Auto_Open()
    If CrownAutoRemembered() Then
        On Error Resume Next
        Application.OnTime Now + TimeSerial(0, 0, 30), "CrownAutoResume"
        On Error GoTo 0
    End If
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

    CrownAutoMode = True

    On Error Resume Next
    CrownRunQuotes
    If Err.Number <> 0 Then CrownAutoNote "the run stopped with an error: " & Err.Description
    Err.Clear
    On Error GoTo 0

    CrownAutoMode = False
    CrownUnattended = False

    ' The next one is booked after this one has finished, so a long batch
    ' can never have a second run start on top of it.
    If CrownAutoOn Then CrownAutoSchedule AUTO_EVERY_MINUTES
End Sub


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
