Attribute VB_Name = "testingverve"
Option Explicit

' Crown 19 Nov (Tony: "quote still did not go to benchmark ... it didnt say
' anything to report what happened to verve"): this was the old Verve quote
' on the Trisura line (Verve 1.1), which Verve has closed. The button that
' runs it now quotes on Benchmark, the same as the automatic quoting, and
' says what happened - whichever button is still pointed here.
Sub ScrapeVerveQuotePaymentttt()
    CrownVerveQuotePay
End Sub


Function SplitAddress(str As String) As Variant
Dim i As Long
Dim ar As Variant
ReDim ar(1 To 2)
str = Trim(str)
For i = 1 To Len(str)
    If AscW(Mid(str, i, 1)) < 48 Or AscW(Mid(str, i, 1)) > 57 Then
        Exit For
    Else
        ar(1) = ar(1) & Mid(str, i, 1)
    End If
Next
ar(2) = Trim(Mid(str, i, Len(str)))
SplitAddress = ar
End Function
