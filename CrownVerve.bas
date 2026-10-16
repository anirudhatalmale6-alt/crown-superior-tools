Attribute VB_Name = "CrownVerve"
Option Explicit

' The Verve leg of CrownRunQuotes.
'
' (The module is CrownVerve, not CrownVerveQuote: VBA will not call a
' procedure that has the same name as the module it lives in.)
'
' His FillVerve macro types the customer in and stops on the Coverages
' page. Everything after that he did by hand, and described it:
'
'   "on coverage place it will click next at bottom of page ... then
'    underwriting questions which will be the same then next then select
'    payment plan direct bill pay 16.67% then next once it is showing
'    quote we need it to run the mvr ... clicking next and there will be
'    a pop up which will either tell you it needs to run mvr or that it is
'    going to final screen and you cant go back. it never wants to do
'    that. ... once mvr is ran it will say ... that it was successful. it
'    now needs to scrape the rates for not the initial one selected but
'    the one circled."
'
'   "5 Pay 16.67% Down CC 30 Day Due this is the rate to scrape to quote"
'
' So this does exactly that, on the browser FillVerve left open, and
' hands back the down payment and the instalment off that one row.


' The pay plan the customer is quoted on. Found by its words, not by its
' place in the table: the table is sorted by down payment, so the row it
' sits on moves from one customer to the next.
Private Const VERVE_PLAN As String = "5 Pay 16.67% Down CC 30 Day Due"

' Words that only appear on the pages they belong to.
Private Const ON_COVERAGES As String = "Vehicle Level Coverages"
Private Const ON_QUESTIONS As String = "Are all residents of the household"
Private Const ON_BILLING As String = "Billing Information"

' How many times Next is pressed looking for the quote screen before
' giving up. Coverages, questions, billing and one spare.
Private Const VERVE_MOST_PAGES As Long = 6


' Quote this customer with Verve. The customer must already be loaded -
' the runner does that before United.
'
' Returns a note for the RunLog. down/monthly are left empty when there
' is no price, and the note says why.
Public Function CrownVerveQuote(ByRef down As String, ByRef monthly As String, _
                                ByRef term As String) As String
    Dim drv As ChromeDriver
    Dim opened As Long
    Dim note As String

    down = ""
    monthly = ""
    term = ""
    CrownLastAlert = ""

    If DrvCol Is Nothing Then Set DrvCol = New Collection
    opened = DrvCol.Count

    ' His own macro for everything up to the Coverages page.
    On Error Resume Next
    Err.Clear
    FillVerve
    If Err.Number <> 0 Then note = "FillVerve stopped: " & Err.Description
    Err.Clear
    On Error GoTo 0

    If DrvCol.Count <= opened Then
        CrownVerveQuote = Trim$("Verve: FillVerve never opened a browser. " & note & " " & CrownLastAlert)

        Exit Function
    End If

    Set drv = DrvCol(DrvCol.Count).ChrmDriver

    ' Any error from his helpers ends up here rather than leaving the
    ' window open.
    On Error Resume Next
    Err.Clear
    ErrorText = ""
    note = CrownVerveFinish(drv, down, monthly, term)
    If Err.Number <> 0 Then note = "Verve: stopped - " & IIf(Len(ErrorText) > 0, ErrorText, Err.Description)
    Err.Clear
    On Error GoTo 0

    ' Close the window so a batch does not leave one open per customer.
    On Error Resume Next
    drv.Quit
    DrvCol.Remove DrvCol.Count
    On Error GoTo 0

    CrownVerveQuote = note
End Function


' From the Coverages page to a price.
Private Function CrownVerveFinish(ByVal drv As ChromeDriver, ByRef down As String, _
                                  ByRef monthly As String, ByRef term As String) As String
    Dim page As String, was As String
    Dim pages As Long, same As Long
    Dim popup As String

    page = CrownVervePageText(drv)

    If InStr(1, page, ON_COVERAGES, vbTextCompare) = 0 Then
        ' FillVerve stopped somewhere earlier. Whatever it saw is the reason.
        CrownNotePage drv, "Verve: FillVerve did not reach the Coverages page"

        If Len(CrownLastAlert) > 0 Then
            CrownVerveFinish = "Verve: stopped before Coverages - " & Left$(CrownLastAlert, 160)
        Else
            CrownVerveFinish = "Verve: stopped before the Coverages page - what was on screen is on the QuoteLook sheet"
        End If

        Exit Function
    End If

    ' Next, Next, Next until the pay plan table is on screen.
    Do While Not CrownVerveHasPlans(drv)
        pages = pages + 1

        If pages > VERVE_MOST_PAGES Then
            CrownNotePage drv, "Verve: pressed Next " & VERVE_MOST_PAGES & " times and the quote never showed"
            CrownVerveFinish = "Verve: the quote screen never came up - see the QuoteLook sheet"

            Exit Function
        End If

        page = CrownVervePageText(drv)

        ' The underwriting questions are always answered the same way. Any
        ' left blank are filled in; any already answered are left alone.
        If InStr(1, page, ON_QUESTIONS, vbTextCompare) > 0 Then CrownVerveAnswerQuestions drv

        ' The billing page: Direct Bill, 5 Pay 16.67% Down, Insured. These
        ' are what Verve puts there itself; only an empty box is filled.
        If InStr(1, page, ON_BILLING, vbTextCompare) > 0 Then CrownVerveBilling drv

        was = page

        ClickElement drv, "btnNext", "ID"
        CrownVerveSettle drv

        popup = CrownVerveTakeAlert(drv, False)

        If Len(popup) > 0 Then
            CrownNoteBlocker "Verve said (on Next): " & Left$(popup, 200)
        End If

        ' A red error on the page means Verve refused to move on.
        If CrownVerveError(drv) <> "" Then
            CrownNotePage drv, "Verve: " & CrownVerveError(drv)
            CrownVerveFinish = "Verve: " & Left$(CrownVerveError(drv), 160)

            Exit Function
        End If

        If CrownVervePageText(drv) = was Then
            same = same + 1

            If same >= 2 Then
                CrownNotePage drv, "Verve: Next did nothing twice"
                CrownVerveFinish = "Verve: the page would not move on - see the QuoteLook sheet"

                Exit Function
            End If
        Else
            same = 0
        End If
    Loop

    ' The quote screen. Next once more asks Verve to run the MVR - or, for
    ' a customer whose MVR has already been run, warns that this is the
    ' final screen. That one is always cancelled: the quote on screen is
    ' then already the final one.
    If CrownVerveElement(drv, "btnSubmit") Then
        ClickElement drv, "btnSubmit", "ID"
    Else
        ClickElement drv, "btnNext", "ID"
    End If

    drv.Wait 1500
    popup = CrownVerveTakeAlert(drv, True)

    If InStr(1, popup, "MVR", vbTextCompare) > 0 Then
        ' Ordering it re-rates the quote. Wait for "Rating Successful" and
        ' the table to come back before reading anything.
        CrownVerveSettle drv
        CrownVerveWaitRated drv
        CrownNoteBlocker "Verve: MVR ordered"
    ElseIf Len(popup) > 0 Then
        CrownNoteBlocker "Verve said (on the quote screen): " & Left$(popup, 200)
    End If

    CrownVerveFinish = CrownVerveReadPlan(drv, down, monthly, term)
End Function


' Read the one pay plan off the table.
Private Function CrownVerveReadPlan(ByVal drv As ChromeDriver, ByRef down As String, _
                                    ByRef monthly As String, ByRef term As String) As String
    Dim rows As String
    Dim lines() As String, cells() As String
    Dim i As Long, c As Long, at As Long
    Dim messages As String

    rows = CrownVervePlanRows(drv)

    If Len(rows) = 0 Then
        CrownNotePage drv, "Verve: no pay plan table after the MVR"
        CrownVerveReadPlan = "Verve: the pay plan table was not there to read - see the QuoteLook sheet"

        Exit Function
    End If

    lines = Split(rows, vbLf)

    For i = LBound(lines) To UBound(lines)
        cells = Split(lines(i), vbTab)
        at = -1

        For c = LBound(cells) To UBound(cells)
            If StrComp(CrownVerveFlat(cells(c)), CrownVerveFlat(VERVE_PLAN), vbTextCompare) = 0 Then
                at = c

                Exit For
            End If
        Next c

        ' Description, Down Payment, # of Installments, Installment Amount,
        ' Installment Charge, Total Amount of Installment, Due Date - so
        ' counted from the description, whatever comes before it.
        If at >= 0 And UBound(cells) >= at + 5 Then
            down = CrownPlainMoney(cells(at + 1))
            monthly = CrownPlainMoney(cells(at + 5))

            If Len(down) > 0 Or Len(monthly) > 0 Then term = "6 months"

            messages = CrownVerveMessages(drv)

            CrownVerveReadPlan = "Verve: " & VERVE_PLAN & " - down " & down & ", " & _
                                 Trim$(cells(at + 2)) & " payments of " & monthly & _
                                 IIf(Len(messages) > 0, " (" & Left$(messages, 120) & ")", "")

            Exit Function
        End If
    Next i

    messages = CrownVerveMessages(drv)
    CrownNotePage drv, "Verve: no '" & VERVE_PLAN & "' row in the pay plans"
    CrownVerveReadPlan = "Verve: no """ & VERVE_PLAN & """ row in the pay plans" & _
                         IIf(Len(messages) > 0, " - Verve said: " & Left$(messages, 140), "")
End Function


' ---------------------------------------------------------- first page

' Verve's first page (name, zip, email, phone) - wait until it has gone,
' clearing whatever stands in the way.
'
' FillVerve pressed Continue, looked ONCE for the "existing client"
' pop-up, then waited for the page (QACMainDiv) to go. When that pop-up
' came a moment later than the one look, nothing pressed New Client and
' the page never moved: "QACMainDiv not found". This keeps looking for as
' long as it waits.
'
' False, with the reason in CrownLastAlert and the page on QuoteLook,
' when Verve shows an error or the page is still there after 90 seconds.
Public Function CrownVervePastFirstPage(ByVal drv As Object) As Boolean
    Dim started As Date
    Dim answer As String
    Dim clicks As Long

    started = Now

    Do
        On Error Resume Next
        answer = ""
        answer = drv.ExecuteScript( _
            "if(!document.getElementById('QACMainDiv'))return 'gone';" & _
            "var b=document.querySelectorAll('[id$=""NewClientButton_CD""],[id$=""NewClientButton""]');" & _
            "for(var i=0;i<b.length;i++){if(b[i].getClientRects().length)return 'new:'+b[i].id;}" & _
            "var e=document.querySelectorAll('.ErrorMessage'),t=[];" & _
            "for(var j=0;j<e.length;j++){if(e[j].getClientRects().length){" & _
            "var s=(e[j].innerText||'').replace(/\s+/g,' ').trim();if(s)t.push(s);}}" & _
            "if(t.length)return 'error:'+t.join(' | ');" & _
            "return 'wait';")
        If Err.Number <> 0 Then answer = "wait"
        Err.Clear
        On Error GoTo 0

        If answer = "gone" Then
            CrownVervePastFirstPage = True

            Exit Function
        End If

        If Left$(answer, 4) = "new:" Then
            clicks = clicks + 1

            If clicks > 3 Then
                CrownNotePage drv, "Verve: pressed New Client 3 times and the first page is still there"
                CrownLastAlert = "Verve: the existing-client pop-up kept coming back"

                Exit Function
            End If

            CrownNoteBlocker "Verve: existing-client pop-up - pressed New Client"

            On Error Resume Next
            drv.FindElementById(Mid$(answer, 5)).Click
            Err.Clear
            On Error GoTo 0

            drv.Wait 2500
        ElseIf Left$(answer, 6) = "error:" Then
            CrownNotePage drv, "Verve first page: " & Mid$(answer, 7)
            CrownLastAlert = "Verve: " & Left$(Mid$(answer, 7), 140)

            Exit Function
        Else
            drv.Wait 1500
        End If

        DoEvents
    Loop While (Now - started) * 86400 < 90

    CrownNotePage drv, "Verve: the first page did not move on after Continue"
    CrownLastAlert = "Verve: the first page did not move on after Continue - QuoteLook has what was on screen"
End Function


' ---------------------------------------------------------------- page

' The words on the page, squashed to one line, for telling pages apart.
Private Function CrownVervePageText(ByVal drv As ChromeDriver) As String
    On Error Resume Next
    CrownVervePageText = drv.ExecuteScript( _
        "return ((document.body&&document.body.innerText)||'').replace(/\s+/g,' ').substring(0,20000);")
End Function

Private Function CrownVerveElement(ByVal drv As ChromeDriver, ByVal id As String) As Boolean
    Dim found As Variant

    On Error Resume Next
    found = drv.ExecuteScript("var e=document.getElementById(arguments[0]);" & _
                              "return !!(e&&e.getClientRects().length);", id)
    CrownVerveElement = (found = True)
End Function

' Every row of the pay plan table, cells separated by tabs.
Private Function CrownVervePlanRows(ByVal drv As ChromeDriver) As String
    On Error Resume Next
    CrownVervePlanRows = drv.ExecuteScript( _
        "var r=document.querySelectorAll(""tr[id*='BillingPayplanPreviewGridView'][id*='DXDataRow']"");" & _
        "var out=[];for(var i=0;i<r.length;i++){var t=[];" & _
        "for(var j=0;j<r[i].cells.length;j++){t.push((r[i].cells[j].innerText||'')" & _
        ".replace(/\s+/g,' ').trim());}out.push(t.join(String.fromCharCode(9)));}" & _
        "return out.join(String.fromCharCode(10));")
End Function

Private Function CrownVerveHasPlans(ByVal drv As ChromeDriver) As Boolean
    CrownVerveHasPlans = (Len(CrownVervePlanRows(drv)) > 0)
End Function

' The MESSAGES box at the top of the quote ("Rating Successful.", "MVR
' report received", or the reason it would not rate).
Private Function CrownVerveMessages(ByVal drv As ChromeDriver) As String
    Dim text As String

    text = CrownVervePageText(drv)

    If InStr(1, text, "MESSAGES", vbBinaryCompare) = 0 Then Exit Function

    text = Mid$(text, InStr(1, text, "MESSAGES", vbBinaryCompare) + 8)

    If InStr(1, text, "General Information", vbTextCompare) > 0 Then
        text = Left$(text, InStr(1, text, "General Information", vbTextCompare) - 1)
    End If

    text = Replace(text, "Continue", "")
    CrownVerveMessages = Trim$(Left$(text, 400))
End Function

' Verve's red error text, if any is showing.
Private Function CrownVerveError(ByVal drv As ChromeDriver) As String
    On Error Resume Next
    CrownVerveError = drv.ExecuteScript( _
        "var e=document.querySelectorAll('.ErrorMessage'),t=[];" & _
        "for(var i=0;i<e.length;i++){if(e[i].getClientRects().length){" & _
        "var s=(e[i].innerText||'').replace(/\s+/g,' ').trim();if(s)t.push(s);}}" & _
        "return t.join(' | ');")
End Function

' Wait for Verve's "loading" cover to go.
Private Sub CrownVerveSettle(ByVal drv As ChromeDriver)
    Dim t0 As Double
    Dim showing As Variant

    t0 = Timer

    Do
        drv.Wait 400
        DoEvents

        On Error Resume Next
        showing = drv.ExecuteScript( _
            "var p=document.querySelectorAll(""#pnlPopup,#imgPopup"");" & _
            "for(var i=0;i<p.length;i++){if(p[i].getClientRects().length&&" & _
            "getComputedStyle(p[i]).display!=='none')return true;}return false;")
        If Err.Number <> 0 Then showing = False   ' an alert is up - let the caller take it
        Err.Clear
        On Error GoTo 0

        If showing <> True Then Exit Do
    Loop While Timer - t0 < 60 And Timer >= t0

    drv.Wait 500
End Sub

' After the MVR: wait until the page says it rated and the table is back.
Private Sub CrownVerveWaitRated(ByVal drv As ChromeDriver)
    Dim t0 As Double
    Dim text As String

    t0 = Timer

    Do
        drv.Wait 1000
        DoEvents

        CrownVerveTakeAlert drv, True
        text = CrownVervePageText(drv)

        If InStr(1, text, "Rating Successful", vbTextCompare) > 0 And CrownVerveHasPlans(drv) Then Exit Do
    Loop While Timer - t0 < 90 And Timer >= t0
End Sub


' ------------------------------------------------------------- pop-ups

' Take whatever pop-up is up and say what it said.
'
' "Ordering MVR. Click OK to continue." - OK.
' "This is the final quote screen ..." - Cancel. His words: "it never
' wants to do that".
' Anything else - Cancel too. Pressing OK on a Verve pop-up nobody has
' read could bind a policy; Cancel at worst leaves the quote unfinished.
Private Function CrownVerveTakeAlert(ByVal drv As ChromeDriver, ByVal onQuote As Boolean) As String
    Dim text As String

    On Error Resume Next
    text = drv.SwitchToAlert(500).text

    If Err.Number <> 0 Then
        Err.Clear

        Exit Function
    End If

    If InStr(1, text, "Ordering MVR", vbTextCompare) > 0 _
       Or InStr(1, text, "re-rate", vbTextCompare) > 0 Then
        drv.SwitchToAlert.Accept
    Else
        drv.SwitchToAlert.Dismiss
    End If

    Err.Clear
    On Error GoTo 0

    CrownVerveTakeAlert = Trim$(text)
End Function


' ------------------------------------------------- questions and billing

' Answer any underwriting question left blank: Yes to "are all residents
' ... listed" and "are you aware that no coverage is afforded ...", No to
' everything else - as on his screenshot.
'
' A question Verve has already answered is not touched.
Private Sub CrownVerveAnswerQuestions(ByVal drv As ChromeDriver)
    Dim done As Variant

    On Error Resume Next
    done = drv.ExecuteScript( _
        "var r=document.querySelectorAll('input[type=radio]'),g={},n=0;" & _
        "for(var i=0;i<r.length;i++){var k=r[i].name||r[i].id;(g[k]=g[k]||[]).push(r[i]);}" & _
        "for(var k in g){var b=g[k];if(b.length!==2)continue;" & _
        "if(b[0].checked||b[1].checked)continue;" & _
        "var row=b[0].closest('tr')||b[0].parentNode;" & _
        "var q=((row&&row.innerText)||'');" & _
        "var yes=/all residents of the household|no coverage is afforded/i.test(q);" & _
        "(yes?b[0]:b[1]).click();n++;}" & _
        "return n;")
    On Error GoTo 0

    If CLng(Val(CStr(done))) > 0 Then
        CrownNoteBlocker "Verve: answered " & done & " blank underwriting question(s)"
    End If
End Sub

' Leave the billing page as Verve fills it unless something is blank, and
' say so rather than guess - the pay plan read at the end is found by its
' name, whichever plan is selected here.
Private Sub CrownVerveBilling(ByVal drv As ChromeDriver)
    Dim blank As Variant

    ' Sometimes Verve leaves the Pay Plan empty and Next answers
    ' "(R-833) Please select a Pay Plan." - pick ours.
    CrownVervePickPayPlan drv

    On Error Resume Next
    blank = drv.ExecuteScript( _
        "var t=document.querySelectorAll('input[id$=""_I""]'),n=0;" & _
        "for(var i=0;i<t.length;i++){if(t[i].id.indexOf('Lookup')>=0)continue;" & _
        "if(t[i].getClientRects().length&&!t[i].value)n++;}" & _
        "return n;")
    On Error GoTo 0

    If CLng(Val(CStr(blank))) > 0 Then
        CrownNotePage drv, "Verve: the Billing page had " & blank & " empty box(es)"
    End If
End Sub


' The Pay Plan list on the billing page, when it is empty: open it the way
' FillVerve opens Verve's lists (click the box), then click the row for
' VERVE_PLAN - or, if that exact wording is not there, the 16.67% one.
Private Sub CrownVervePickPayPlan(ByVal drv As ChromeDriver)
    Dim box As String, row As String, now_ As String

    On Error Resume Next
    box = CStr(drv.ExecuteScript( _
        "var t=document.querySelectorAll('input[id$=""_I""]');" & _
        "for(var i=0;i<t.length;i++){var e=t[i];if(!/pay_?plan/i.test(e.id))continue;" & _
        "if(!e.getClientRects().length)continue;" & _
        "var v=(e.value||'').trim();if(v&&!/^(select|-+|none)$/i.test(v))return '';" & _
        "return e.id;}return '';"))
    If Err.Number <> 0 Then box = ""
    Err.Clear

    If Len(box) = 0 Then
        On Error GoTo 0

        Exit Sub
    End If

    drv.FindElementById(box).Click
    drv.Wait 700

    row = CStr(drv.ExecuteScript( _
        "function f(s){return (s||'').replace(/\s+/g,' ').trim().toLowerCase();}" & _
        "var list=document.getElementById(arguments[0].replace(/_I$/,'_DDD_L_LBT'));" & _
        "if(!list)return 'no list';var want=f(arguments[1]),best=null;" & _
        "var r=list.querySelectorAll('tr');" & _
        "for(var i=0;i<r.length;i++){var t=f(r[i].innerText);" & _
        "if(t===want){best=r[i];break;}if(!best&&t.indexOf('16.67')>=0)best=r[i];}" & _
        "if(!best)return 'no 16.67 plan';" & _
        "var c=best.querySelector('td')||best;c.click();return 'ok '+f(best.innerText);", Array(box, VERVE_PLAN)))
    If Err.Number <> 0 Then row = "error: " & Err.Description
    Err.Clear

    drv.Wait 800
    now_ = CStr(drv.FindElementById(box).Value)
    Err.Clear
    On Error GoTo 0

    CrownNoteBlocker "Verve billing: the Pay Plan was empty - " & row & " (box now: " & now_ & ")"
End Sub


' ------------------------------------------------------------- helpers

Private Function CrownVerveFlat(ByVal text As String) As String
    text = Replace(text, Chr$(160), " ")

    Do While InStr(text, "  ") > 0
        text = Replace(text, "  ", " ")
    Loop

    CrownVerveFlat = Trim$(text)
End Function

' "$1,234.56" -> "1234.56"
Private Function CrownPlainMoney(ByVal value As String) As String
    Dim i As Long, ch As String, out As String

    For i = 1 To Len(value)
        ch = Mid$(value, i, 1)

        If (ch >= "0" And ch <= "9") Or ch = "." Then out = out & ch
    Next i

    If IsNumeric(out) Then CrownPlainMoney = out
End Function
