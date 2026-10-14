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
' asleep. Closing Excel stops the timer; open the workbook and start it
' again.
' =====================================================================

' How often it looks.
Private Const AUTO_EVERY_MINUTES As Long = 5

Private CrownAutoNextRun As Date
Public CrownAutoOn As Boolean


Public Sub CrownStartAutoQuoting()
    If CrownAutoOn Then
        MsgBox "Automatic quoting is already on. Next look at " & _
               Format$(CrownAutoNextRun, "h:nn AM/PM") & ".", vbInformation, "Crown Superior"

        Exit Sub
    End If

    CrownAutoOn = True
    CrownAutoNote "automatic quoting started - looking every " & AUTO_EVERY_MINUTES & " minutes"

    ' The first look straight away (in 10 seconds), not in 5 minutes.
    CrownAutoSchedule 0

    MsgBox "Automatic quoting is ON." & vbCrLf & vbCrLf & _
           "Every " & AUTO_EVERY_MINUTES & " minutes it checks the website. When quotes are " & _
           "waiting and it is inside your schedule, it quotes them through United and Verve " & _
           "and sends the answers back - no boxes to press." & vbCrLf & vbCrLf & _
           "Leave Excel open with this workbook. Everything it does is written on the " & _
           "RunLog sheet." & vbCrLf & vbCrLf & _
           "To stop: Alt+F8 > CrownStopAutoQuoting > Run.", vbInformation, "Crown Superior"
End Sub


Public Sub CrownStopAutoQuoting()
    CrownAutoOn = False

    On Error Resume Next
    Application.OnTime CrownAutoNextRun, "CrownAutoTick", , False
    On Error GoTo 0

    CrownAutoNote "automatic quoting stopped"

    MsgBox "Automatic quoting is OFF. Nothing will be quoted until you run " & _
           "CrownRunQuotes or CrownStartAutoQuoting again.", vbInformation, "Crown Superior"
End Sub


' What the timer runs. Public only because Excel's timer can only call a
' public macro - it is not meant to be run by hand.
Public Sub CrownAutoTick()
    If Not CrownAutoOn Then Exit Sub

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
    If minutes <= 0 Then
        CrownAutoNextRun = Now + TimeSerial(0, 0, 10)
    Else
        CrownAutoNextRun = Now + TimeSerial(0, minutes, 0)
    End If

    Application.OnTime CrownAutoNextRun, "CrownAutoTick"
End Sub
