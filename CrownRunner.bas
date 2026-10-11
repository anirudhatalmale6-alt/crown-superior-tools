Attribute VB_Name = "CrownRunner"
Option Explicit

' =====================================================================
' Crown Superior - quoting the queue without being asked each time
' ---------------------------------------------------------------------
' Leave the tool open and run this. It asks the website what is waiting,
' quotes each one through United the way you already do, sends the answer
' back, and writes down what happened to every single one.
'
' It does NOT decide on its own that it is allowed. The switch lives on
' the website - Get a quote, Customer Quotes, "Rating quotes through the
' carriers" - so you can turn it off from your phone and this will stop
' at the next quote rather than finishing the batch.
'
' WHERE THE FIGURES COME FROM
'
' FillandScrapeUaigQuotePayment writes a row onto the QuotePay sheet:
'
'     F  UAIG          L  first due date
'     G  first name    M  down payment
'     H  last name     N  installment
'     I  e-mail
'     J  phone
'
' So this notes how far down QuotePay has got BEFORE quoting, and reads
' the row that appears afterwards. No new row means the quote did not
' finish - and then nothing is sent, because a quote that did not happen
' must not reach a customer as one that did.
'
' WHAT IT WILL STILL STOP FOR
'
' Alerts it knows are harmless are waved through by CrownUnblock. Alerts
' it does not know are written down and that one quote is abandoned -
' unattended, there is nobody to ask.
'
' It cannot get past two things that are in the quoting module itself:
' the "Could not find payment plan details" box, and the retry question
' the error handler asks. Both are ordinary message boxes and they will
' sit there until somebody presses them. Say the word and I will give
' those the same treatment.
' =====================================================================

' Set while a batch is running. CrownUnblock reads it and stops asking
' questions nobody is there to answer.
Public CrownUnattended As Boolean

' How many in a row may fail before it gives up on the batch. A single
' bad quote is a bad quote; four in a row is something wrong with the
' carrier's site or with the sign-in, and grinding through fifty more
' helps nobody.
Private Const RUN_STOP_AFTER As Long = 4

' Most quotes in one go, so an overnight run cannot surprise anybody.
Private Const RUN_MOST As Long = 25


Public Sub CrownRunQuotes()
    Dim json As String
    Dim ids As Collection, names As Collection
    Dim rateNow As Boolean
    Dim i As Long, did As Long, failed As Long, inARow As Long
    Dim who As String, note As String, answer As String
    Dim carrier As String, down As String, monthly As String, term As String
    Dim ws As Worksheet, row As Long
    Dim before As Long, after As Long

    json = CrownQuoteWork(RUN_MOST)

    If Len(Trim$(json)) = 0 Then
        MsgBox "The website did not answer, so nothing has been done.", _
               vbExclamation, "Crown Superior"

        Exit Sub
    End If

    rateNow = (InStr(1, json, """rate_now"":true", vbTextCompare) > 0)

    If Not rateNow Then
        MsgBox "Rating through the carriers is switched OFF on the website, " & _
               "so nothing has been quoted." & vbCrLf & vbCrLf & _
               "To turn it on: Get a quote, Customer Quotes, " & _
               """Rating quotes through the carriers"".", _
               vbInformation, "Crown Superior"

        Exit Sub
    End If

    Set ids = New Collection
    Set names = New Collection
    CrownReadWorkList json, ids, names

    If ids.Count = 0 Then
        MsgBox "Nothing is waiting to be quoted.", vbInformation, "Crown Superior"

        Exit Sub
    End If

    If MsgBox(ids.Count & " quote(s) waiting." & vbCrLf & vbCrLf & _
              "This will quote each one through United and send the answer " & _
              "back to the website. You can leave it running." & vbCrLf & vbCrLf & _
              "Carry on?", vbYesNo + vbQuestion, "Crown Superior") <> vbYes Then
        Exit Sub
    End If

    Set ws = CrownRunSheet()
    row = 2

    CrownUnattended = True

    ' Bring the new quotes into the workbook first.
    '
    ' The runner asks the website what is waiting and then looks each one
    ' up in the workbook. A quote that has come in since the last fetch is
    ' not there yet, and the run logs "could not find that quote in the
    ' tool" for every single one - which is what happened to Eltiana
    ' Jones. Remembering to press another button first is not something an
    ' unattended runner should need.
    '
    ' Quietly, because a message box would stop the run until somebody
    ' pressed OK. Whatever it says goes on the RunLog instead.
    ws.Cells(row, 1).value = Now
    ws.Cells(row, 6).value = "bringing new quotes in: " & _
        Replace(Replace(CrownFetchReport(), vbCrLf, " "), vbLf, " ")
    row = row + 1

    On Error GoTo Finished

    For i = 1 To ids.Count
        who = CStr(names(i))
        note = ""
        carrier = "United Auto"
        down = ""
        monthly = ""
        term = ""

        ws.Cells(row, 1).value = Now
        ws.Cells(row, 2).value = ids(i)
        ws.Cells(row, 3).value = who

        If Not CrownLoadQuoteById(CLng(ids(i)), who) Then
            note = "could not find that quote in the tool - run CrownFetchNewQuotes first"
        Else
            before = CrownQuotePayRows()

            ' His own quoting, exactly as he runs it by hand.
            On Error Resume Next
            Err.Clear
            FillandScrapeUaigQuotePayment
            If Err.Number <> 0 Then note = "the quoting stopped: " & Err.Description
            Err.Clear
            On Error GoTo Finished

            after = CrownQuotePayRows()

            If after > before Then
                CrownReadQuotePayRow after, down, monthly, term
            ElseIf Len(note) = 0 Then
                ' Nothing was written, so nothing was quoted. Said plainly
                ' rather than sending a blank quote to a customer.
                note = "no payment plan came back - nothing sent"
            End If
        End If

        If Len(down) > 0 Or Len(monthly) > 0 Then
            answer = CrownQuoteResult(CLng(ids(i)), _
                Array("carrier", "state", "down", "monthly", "term", "who"), _
                Array(carrier, "quoted", down, monthly, term, "the quoting tool"))

            If InStr(1, answer, """ok"":true", vbTextCompare) > 0 Then
                did = did + 1
                inARow = 0
                note = Trim$(note & " sent: down " & down & ", monthly " & monthly)
            Else
                failed = failed + 1
                inARow = inARow + 1
                note = Trim$(note & " the website refused it: " & Left$(answer, 90))
            End If
        Else
            failed = failed + 1
            inARow = inARow + 1
        End If

        ws.Cells(row, 4).value = down
        ws.Cells(row, 5).value = monthly
        ws.Cells(row, 6).value = note
        row = row + 1

        DoEvents

        If inARow >= RUN_STOP_AFTER Then
            ws.Cells(row, 6).value = "stopped - " & inARow & " in a row came back with nothing"
            row = row + 1

            Exit For
        End If
    Next i

Finished:
    CrownUnattended = False

    On Error Resume Next
    ws.Activate
    On Error GoTo 0

    MsgBox ids.Count & " quote(s) were waiting." & vbCrLf & _
           did & " quoted and sent to the website." & vbCrLf & _
           failed & " came back with nothing and were left for you." & vbCrLf & vbCrLf & _
           IIf(inARow >= RUN_STOP_AFTER, _
               "It stopped early because " & inARow & " in a row came back with nothing - " & _
               "worth checking you are still signed in to United." & vbCrLf & vbCrLf, "") & _
           "The RunLog sheet has a line for every one.", _
           vbInformation, "Crown Superior"
End Sub


' How far down the QuotePay sheet has got.
'
' Column F is the one his own writer counts from, so it is the one
' counted here - a different column would disagree the moment a row is
' written with some cells blank.
Private Function CrownQuotePayRows() As Long
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("QuotePay")
    On Error GoTo 0

    If ws Is Nothing Then Exit Function

    CrownQuotePayRows = ws.Cells(ws.Rows.Count, "F").End(xlUp).row
End Function


' The figures off one QuotePay row.
'
' Term comes from the pay plan his scraper picks - the 16.67% down column,
' which is the six pay. It is written out rather than worked out from the
' instalment, because a number derived from another number is a number
' that can be wrong in a new way.
Private Sub CrownReadQuotePayRow(ByVal row As Long, ByRef down As String, _
                                 ByRef monthly As String, ByRef term As String)
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("QuotePay")
    On Error GoTo 0

    If ws Is Nothing Or row < 2 Then Exit Sub

    down = CrownPlainMoney(ws.Cells(row, "M").value)
    monthly = CrownPlainMoney(ws.Cells(row, "N").value)

    If Len(down) > 0 Or Len(monthly) > 0 Then term = "6 months"
End Sub


' An amount with the dollar sign and the commas off, or nothing at all.
Private Function CrownPlainMoney(ByVal value As Variant) As String
    Dim text As String
    Dim i As Long, ch As String, out As String

    text = Trim$(CStr(value))
    If Len(text) = 0 Then Exit Function

    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)

        If (ch >= "0" And ch <= "9") Or ch = "." Then
            out = out & ch
        ElseIf ch = "-" And Len(out) = 0 Then
            out = "-"
        End If
    Next i

    If Not IsNumeric(out) Then Exit Function

    CrownPlainMoney = out
End Function


' The request numbers and names out of what the website sent.
'
' Read by hand rather than with a JSON reader, because VBA has none and
' this answer has a known shape: every item carries "id": followed by a
' number, and "name": followed by text. Anything it cannot read is
' skipped rather than guessed at.
Public Sub CrownReadWorkList(ByVal json As String, ByRef ids As Collection, _
                             ByRef names As Collection)
    Dim at As Long, stop_ As Long
    Dim number As String, who As String

    at = InStr(1, json, """items""", vbTextCompare)
    If at = 0 Then Exit Sub

    Do
        at = InStr(at + 1, json, """id"":", vbTextCompare)
        If at = 0 Then Exit Do

        number = ""
        stop_ = at + 5

        Do While stop_ <= Len(json)
            If Mid$(json, stop_, 1) >= "0" And Mid$(json, stop_, 1) <= "9" Then
                number = number & Mid$(json, stop_, 1)
                stop_ = stop_ + 1
            Else
                Exit Do
            End If
        Loop

        If Len(number) = 0 Then Exit Do

        who = CrownJsonText(json, at, "name")

        ids.Add CLng(number)
        names.Add IIf(Len(who) = 0, "request " & number, who)
    Loop
End Sub


' The text of one field, looked for after a given point.
Private Function CrownJsonText(ByVal json As String, ByVal from_ As Long, _
                               ByVal field As String) As String
    Dim at As Long, stop_ As Long
    Dim ch As String

    at = InStr(from_, json, """" & field & """:""", vbTextCompare)
    If at = 0 Then Exit Function

    at = at + Len(field) + 4
    stop_ = at

    Do While stop_ <= Len(json)
        ch = Mid$(json, stop_, 1)

        If ch = "\" Then
            stop_ = stop_ + 2
        ElseIf ch = """" Then
            Exit Do
        Else
            stop_ = stop_ + 1
        End If
    Loop

    CrownJsonText = Replace(Mid$(json, at, stop_ - at), "\/", "/")
End Function


' The sheet every run writes to.
Private Function CrownRunSheet() As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("RunLog")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "RunLog"
    End If

    ws.Cells.ClearContents
    ws.Range("A1:F1").value = Array("When", "Request", "Customer", _
                                    "Down", "Monthly", "What happened")
    ws.Columns(6).ColumnWidth = 80

    Set CrownRunSheet = ws
End Function
