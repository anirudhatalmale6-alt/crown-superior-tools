Attribute VB_Name = "CrownMorePayments"
Option Explicit

' =====================================================================
' Crown Superior - payment due for Assurance America and Trexis
' ---------------------------------------------------------------------
' Alt+F8 > CrownAssurancePaymentDue > Run
' Alt+F8 > CrownTrexisPaymentDue > Run    (further down)
' Progressive, First Acceptance, National General: at the end - they ask
' for a verification code when signing in.
'
' Same idea as the United and Verve runs: each policy's amount due, due
' date, cancel date and status go straight onto the website, one policy at
' a time, and the AssurancePayments sheet is the report of the run.
'
' The steps on Assurance America's site are the ones Tony's own macro
' (ScrapeAssuranceAmericapaymentdue) already uses - sign in with USER_7 /
' PASS_7 at URL_11, the producer list, the rows with the pencil and the $
' icons, open each by the insured's name, read "Next Payment Due", the
' status and the cancel date, close it. What is new is where the answers
' go: straight to the website instead of a CSV in Downloads that then had
' to be filtered and uploaded.
'
' Which website record: Assurance America writes its numbers with a
' suffix ("-02"), the website usually holds the number before it. Tried
' in this order, and a name is only used when it points at ONE policy:
'   whole number -> number before the first hyphen -> after the last
'   hyphen -> first + last name with this carrier.
' =====================================================================

Private Const AA_CARRIER As String = "Assurance America"
Private Const AA_LIST_WAIT_SECONDS As Long = 120
Private Const TX_CARRIER As String = "Trexis"
Private Const PG_CARRIER As String = "Progressive"
Private Const FA_CARRIER As String = "First Acceptance"
Private Const NG_CARRIER As String = "National General"
Private Const NG_SITE As String = "https://natgenagency.com"


Public Sub CrownAssurancePaymentDue()
    Dim drv As ChromeDriver, clsDrv As Chrm
    Dim Keys As New Selenium.Keys
    Dim wsInput As Worksheet, ws As Worksheet
    Dim byNumber As Object, byName As Object
    Dim picks As Collection, idx As Variant
    Dim answer As String, info As String, parts() As String
    Dim policyNo As String, insured As String
    Dim dueAmount As String, dueDate As String, cancelDate As String, status As String
    Dim recordId As String, matchedOn As String, note As String
    Dim row As Long, done As Long, unmatched As Long, blank As Long, failed As Long
    Dim started As Date

    Set wsInput = ThisWorkbook.Worksheets("Input")

    CrownMoreIndex AA_CARRIER, byNumber, byName

    If byNumber.Count = 0 And byName.Count = 0 Then
        MsgBox "No Assurance America policies came back from the website.", vbExclamation, "Crown Superior"
        Exit Sub
    End If

    Set ws = CrownSheet("AssurancePayments")
    ws.Cells.ClearContents
    ws.Columns(2).NumberFormat = "@"
    ws.Range("A1:J1").value = Array("Record number", "Policy at Assurance", "Insured", "Amount due", _
                                    "Due date", "Cancel date", "Status", "When", "Matched on", "What happened")
    row = 2

    Set clsDrv = New Chrm
    Set clsDrv.ChrmDriver = New ChromeDriver
    Set drv = clsDrv.ChrmDriver

    On Error Resume Next
    drv.AddArgument "--force-device-scale-factor=0.70"
    drv.Start
    On Error GoTo 0

    ' Chrome did not start (usually: Chrome updated, chromedriver is old) - say so in words.
    If Not CrownChromeOk(drv) Then Exit Sub

    drv.Get wsInput.Range("URL_11").value

    ' --- sign in (his PerformLogin, by the boxes' own names) ---------------
    started = Now

    Do
        answer = CrownMoreJs(drv, _
            "return (document.querySelector(""input[formcontrolname='username']"")&&" & _
            "document.querySelector(""input[formcontrolname='password']""))?'ready':'';")
        If answer = "ready" Then Exit Do
        drv.Wait 700
        DoEvents
    Loop While (Now - started) * 86400 < 60

    If answer <> "ready" Then
        MsgBox "Assurance America did not show its sign-in page. Nothing has been changed.", _
               vbExclamation, "Crown Superior"
        GoTo Finish
    End If

    EnterData drv, Keys, "//input[@formcontrolname='username']", wsInput.Range("USER_7").value, "XPATH"
    EnterData drv, Keys, "//input[@formcontrolname='password']", wsInput.Range("PASS_7").value, "XPATH"
    ClickElement drv, "//*[@id='login-button']", "XPATH"

    ' --- the producer list: rows with both the pencil and the $ ---------
    ' His macro stopped here with "Press Ok when data will Visible". This
    ' waits for the rows instead, and only asks if they never come.
    started = Now

    Do
        answer = CrownMoreJs(drv, CrownAaPicksJs())
        If Len(answer) > 0 And answer <> "none" Then Exit Do
        drv.Wait 1500
        DoEvents
    Loop While (Now - started) * 86400 < AA_LIST_WAIT_SECONDS

    If Len(answer) = 0 Or answer = "none" Then
        If MsgBox("The Assurance America list of policies has not shown up yet." & vbCrLf & vbCrLf & _
                  "If it is on the screen now, press OK. Press Cancel to stop.", _
                  vbOKCancel + vbQuestion, "Crown Superior") = vbCancel Then GoTo Finish

        answer = CrownMoreJs(drv, CrownAaPicksJs())
    End If

    Set picks = New Collection

    If Len(answer) > 0 And answer <> "none" Then
        For Each idx In Split(answer, ",")
            picks.Add CLng(idx)
        Next idx
    End If

    If picks.Count = 0 Then
        MsgBox "No rows with a payment due on the Assurance America list. Nothing has been changed.", _
               vbInformation, "Crown Superior"
        GoTo Finish
    End If

    ' Every policy number on the list, for the same reason as Trexis.
    Dim listed As Object
    Set listed = CreateObject("Scripting.Dictionary")

    For Each idx In picks
        info = CrownMoreJs(drv, CrownAaRowJs(), CStr(idx))

        If Left$(info, 3) = "ok|" Then
            parts = Split(info, "|")
            If Len(CrownDigits(parts(2))) > 0 Then listed(CrownDigits(parts(2))) = True
        End If
    Next idx

    ' --- each policy ----------------------------------------------------
    For Each idx In picks
        dueAmount = "": dueDate = "": cancelDate = "": status = ""
        recordId = "": matchedOn = "": note = ""

        ' Back on the list before reading the next row.
        started = Now
        Do
            info = CrownMoreJs(drv, CrownAaRowJs(), CStr(idx))
            If Left$(info, 3) = "ok|" Then Exit Do
            drv.Wait 700
            DoEvents
        Loop While (Now - started) * 86400 < 30

        If Left$(info, 3) <> "ok|" Then
            ws.Cells(row, 10).value = "row " & idx & " of the list could not be read (" & info & ") - stopped here"
            row = row + 1
            failed = failed + 1
            Exit For
        End If

        parts = Split(info, "|")
        insured = parts(1)
        policyNo = parts(2)

        ' Open it the way his macro does - by the name link in that row.
        On Error Resume Next
        ClickElement drv, ".mat-row.mat-row:nth-of-type(" & idx & ") .cdk-column-InsuredName .mat-cell-link", "CSS"
        On Error GoTo 0

        drv.Wait 1000

        started = Now
        Do
            answer = CrownMoreJs(drv, CrownAaDetailJs())
            If Left$(answer, 3) = "ok|" Then Exit Do
            drv.Wait 600
            DoEvents
        Loop While (Now - started) * 86400 < 30

        If Left$(answer, 3) = "ok|" Then
            parts = Split(answer, "|")
            dueDate = CrownMoreDate(parts(1))
            dueAmount = CrownMoreMoney(parts(2))
            status = Trim$(parts(3))
            cancelDate = CrownMoreDate(parts(4))
        Else
            note = "the policy did not open (" & answer & ")"
        End If

        ' Close it - his "x" button.
        On Error Resume Next
        ClickElement drv, "//button[.//span[text()='x']]", "XPATH"
        On Error GoTo 0
        drv.Wait 800

        recordId = CrownMoreMatch(byNumber, byName, AA_CARRIER, policyNo, insured, matchedOn, note, listed)

        ws.Cells(row, 2).value = policyNo
        ws.Cells(row, 3).value = insured
        ws.Cells(row, 4).value = dueAmount
        ws.Cells(row, 5).value = dueDate
        ws.Cells(row, 6).value = cancelDate
        ws.Cells(row, 7).value = status
        ws.Cells(row, 8).value = Now
        ws.Cells(row, 9).value = matchedOn

        If Len(recordId) = 0 Then
            If Len(note) = 0 Then note = "no policy on the website with this number or this name"
            ws.Cells(row, 10).value = note
            unmatched = unmatched + 1
        ElseIf Len(dueAmount & dueDate & cancelDate & status) = 0 Then
            ws.Cells(row, 1).value = recordId
            ws.Cells(row, 10).value = Trim$("nothing read - left as it was. " & note)
            blank = blank + 1
        Else
            ws.Cells(row, 1).value = recordId
            answer = CrownMoreWrite(recordId, dueAmount, dueDate, cancelDate, status)

            If InStr(1, answer, """ok"":true", vbTextCompare) > 0 Then
                ws.Cells(row, 10).value = Trim$("written. " & note)
                done = done + 1
            ElseIf Len(Trim$(answer)) = 0 Then
                ws.Cells(row, 10).value = "the website did not answer - run again for this one"
                failed = failed + 1
            Else
                ws.Cells(row, 10).value = "refused: " & answer
                failed = failed + 1
            End If
        End If

        row = row + 1
        DoEvents
    Next idx

Finish:
    On Error Resume Next
    drv.Quit
    On Error GoTo 0

    ws.Columns("A:J").AutoFit
    ThisWorkbook.Save
    ws.Activate

    MsgBox done & " Assurance America policies written to the website." & vbCrLf & _
           unmatched & " could not be matched to a policy on the website." & vbCrLf & _
           blank & " opened but showed nothing to read - left as they were." & vbCrLf & _
           failed & " did not go through." & vbCrLf & vbCrLf & _
           "The AssurancePayments sheet has every one, and how it was matched.", _
           vbInformation, "Crown Superior"
End Sub


' Row numbers (1-based, as his nth-of-type counts them) of the list rows
' that have both a pencil and a $ icon - his test for "payment due".
Private Function CrownAaPicksJs() As String
    CrownAaPicksJs = _
        "var rs=document.querySelectorAll('mat-row.mat-row');if(!rs.length)return '';var out=[];" & _
        "for(var i=0;i<rs.length;i++){var pen=false,usd=false,ic=rs[i].querySelectorAll('i.fa');" & _
        "for(var k=0;k<ic.length;k++){var c=ic[k].className||'';" & _
        "if(c.indexOf('fa-pencil')>=0)pen=true;if(c.indexOf('fa-usd')>=0)usd=true;}" & _
        "if(pen&&usd)out.push(i+1);}return out.length?out.join(','):'none';"
End Function

' "ok|insured|policy" for row n of the list.
Private Function CrownAaRowJs() As String
    CrownAaRowJs = _
        "var n=parseInt(arguments[0],10),rs=document.querySelectorAll('mat-row.mat-row');" & _
        "if(rs.length<n)return 'list not showing';var r=rs[n-1];" & _
        "function t(s){var e=r.querySelector(s);return e?(e.innerText||'').replace(/[|]/g,' ').replace(/\s+/g,' ').trim():'';}" & _
        "return 'ok|'+t('.cdk-column-InsuredName .mat-cell-link')+'|'+t('.cdk-column-PolicyIdentifier .mat-cell-link span');"
End Function

' "ok|due date|amount|status|cancel date" once the policy has loaded.
Private Function CrownAaDetailJs() As String
    CrownAaDetailJs = _
        "function c(s){return ((s||'')+'').replace(/[|]/g,' ').replace(/\s+/g,' ').trim();}" & _
        "if(/Loading Policy/i.test(document.body.innerText||''))return 'loading';" & _
        "var st=document.getElementById('policy-status');if(!st)return 'no policy-status';" & _
        "function x(p){var r=document.evaluate(p,document,null,XPathResult.FIRST_ORDERED_NODE_TYPE,null);return r.singleNodeValue;}" & _
        "var d=x(""//span[contains(text(),'Next Payment Due')]""),due='',amt='';" & _
        "if(d){due=c(d.innerText).replace('Next Payment Due','').replace(/[()]/g,'');" & _
        "var a=x(""//span[contains(text(),'Next Payment Due')]/following::span[1]"");if(a)amt=c(a.innerText);}" & _
        "var cx=document.getElementById('cancelDate');" & _
        "return 'ok|'+c(due)+'|'+amt+'|'+c(st.innerText)+'|'+(cx?c(cx.innerText).replace(/[()]/g,''):'');"
End Function


' =====================================================================
' Trexis
' ---------------------------------------------------------------------
' Alt+F8 > CrownTrexisPaymentDue > Run
'
' His ScrapeTrexisPolicy1234 steps: sign in (URL_8, USER_6 / PASS_6),
' Policy, Search, "Export" - Trexis hands back a CSV of its policies with
' Status, Status Date and Amount Due. For a policy still In-Force the
' date the payment is due is only on the policy's own page ("Current
' Payment Due"), so those - and only those - are opened.
'
' Written to the website the way his upload did it: Status as Trexis
' words it, Amount Due, the due date from the policy page, and the Status
' Date as the cancel date unless the policy is In-Force.
' =====================================================================

Public Sub CrownTrexisPaymentDue()
    Dim drv As ChromeDriver, clsDrv As Chrm
    Dim Keys As New Selenium.Keys
    Dim wsInput As Worksheet, ws As Worksheet
    Dim byNumber As Object, byName As Object
    Dim answer As String, csvPath As String
    Dim lines() As String, head() As String, f() As String
    Dim cPol As Long, cName As Long, cStatus As Long, cStatusDate As Long, cAmount As Long
    Dim i As Long, row As Long
    Dim policyNo As String, insured As String, status As String, statusDate As String
    Dim dueAmount As String, dueDate As String, cancelDate As String
    Dim recordId As String, matchedOn As String, note As String
    Dim done As Long, unmatched As Long, failed As Long
    Dim started As Date, exportAt As Date
    Dim listed As Object

    Set wsInput = ThisWorkbook.Worksheets("Input")

    CrownMoreIndex TX_CARRIER, byNumber, byName

    If byNumber.Count = 0 And byName.Count = 0 Then
        MsgBox "No Trexis policies came back from the website.", vbExclamation, "Crown Superior"
        Exit Sub
    End If

    Set ws = CrownSheet("TrexisPayments")
    ws.Cells.ClearContents
    ws.Columns(2).NumberFormat = "@"
    ws.Range("A1:J1").value = Array("Record number", "Policy at Trexis", "Insured", "Amount due", _
                                    "Due date", "Cancel date", "Status", "When", "Matched on", "What happened")
    row = 2

    Set clsDrv = New Chrm
    Set clsDrv.ChrmDriver = New ChromeDriver
    Set drv = clsDrv.ChrmDriver

    On Error Resume Next
    drv.AddArgument "--force-device-scale-factor=0.70"
    drv.AddArgument "--disable-extensions"
    drv.AddArgument "--disable-popup-blocking"
    drv.Start
    On Error GoTo 0

    ' Chrome did not start (usually: Chrome updated, chromedriver is old) - say so in words.
    If Not CrownChromeOk(drv) Then Exit Sub

    drv.Get wsInput.Range("URL_8").value

    ' --- sign in ---------------------------------------------------------
    If Not CrownMoreWaitJs(drv, "return (document.getElementById('loginName')&&document.getElementById('loginPassword'))?'ok':'';", 60) Then
        MsgBox "Trexis did not show its sign-in page. Nothing has been changed.", vbExclamation, "Crown Superior"
        GoTo Finish
    End If

    EnterData drv, Keys, "loginName", wsInput.Range("USER_6").value, "ID"
    EnterData drv, Keys, "loginPassword", wsInput.Range("PASS_6").value, "ID"
    drv.Wait 1000
    ClickElement drv, "//*[@id='loginSubmit']/div[1]/button", "XPATH"

    If Not CrownMoreWaitJs(drv, "return document.getElementById('agentHomePolicyButton')?'ok':'';", 90) Then
        MsgBox "Signed in to Trexis, but its home page never came up. Nothing has been changed.", _
               vbExclamation, "Crown Superior"
        GoTo Finish
    End If

    drv.Wait 3000

    ' The notice box his macro closes, when it is there.
    CrownMoreJs drv, "var b=document.evaluate(""//*[@id='modalPopupAM']/div/div[3]/button"",document,null," & _
                     "XPathResult.FIRST_ORDERED_NODE_TYPE,null).singleNodeValue;if(b&&b.offsetParent)b.click();return '';"
    drv.Wait 800

    ' --- Policy > Search > Export -----------------------------------------
    CrownMoreJs drv, "var b=document.getElementById('agentHomePolicyButton');if(b)b.click();return '';"

    If Not CrownMoreWaitJs(drv, "return document.getElementById('searchPolicyQuote')?'ok':'';", 30) Then
        ' His WaitForPageChange: the policy centre sometimes needs a reload.
        drv.Refresh
        If Not CrownMoreWaitJs(drv, "return document.getElementById('searchPolicyQuote')?'ok':'';", 30) Then
            MsgBox "Trexis did not open its policy search. Nothing has been changed.", vbExclamation, "Crown Superior"
            GoTo Finish
        End If
    End If

    ClickElement drv, "searchPolicyQuote", "ID"

    If Not CrownMoreWaitJs(drv, "return document.getElementById('exportPolicyRecordsToCVS')?'ok':'';", 60) Then
        MsgBox "Trexis did not show its Export button. Nothing has been changed.", vbExclamation, "Crown Superior"
        GoTo Finish
    End If

    exportAt = Now
    ClickElement drv, "exportPolicyRecordsToCVS", "ID"

    csvPath = CrownMoreNewDownload(exportAt, 90)

    If Len(csvPath) = 0 Then
        MsgBox "Trexis's export file never arrived in Downloads. Nothing has been changed.", _
               vbExclamation, "Crown Superior"
        GoTo Finish
    End If

    lines = Split(Replace(CrownMoreReadText(csvPath), vbCrLf, vbLf), vbLf)

    If UBound(lines) < 1 Then
        MsgBox "Trexis's export file was empty. Nothing has been changed.", vbExclamation, "Crown Superior"
        GoTo Finish
    End If

    head = CrownSplitCsvLine(lines(0))
    cPol = CrownMoreColumn(head, "Policy Number")
    cName = CrownMoreColumn(head, "Insured Name")
    cStatus = CrownMoreColumn(head, "Status")
    cStatusDate = CrownMoreColumn(head, "Status Date")
    cAmount = CrownMoreColumn(head, "Amount Due")

    If cPol < 0 Or cStatus < 0 Or cAmount < 0 Then
        MsgBox "Trexis's export file did not have the columns expected (Policy Number, Status, Amount Due)." & _
               vbCrLf & vbCrLf & "It is " & csvPath & " - please send it to me.", vbExclamation, "Crown Superior"
        GoTo Finish
    End If

    ' Every number on the export, so a name match cannot land on a record
    ' that belongs to another policy on this same list.
    Set listed = CreateObject("Scripting.Dictionary")

    For i = 1 To UBound(lines)
        f = CrownSplitCsvLine(lines(i))
        If UBound(f) >= cPol Then
            If Len(CrownDigits(f(cPol))) > 0 Then listed(CrownDigits(f(cPol))) = True
        End If
    Next i

    ' --- each policy in the export -----------------------------------------
    For i = 1 To UBound(lines)
        If Len(Trim$(lines(i))) = 0 Then GoTo NextTrexis

        f = CrownSplitCsvLine(lines(i))
        If UBound(f) < cAmount Or UBound(f) < cPol Then GoTo NextTrexis

        policyNo = Trim$(f(cPol))
        insured = ""
        If cName >= 0 And UBound(f) >= cName Then insured = Trim$(f(cName))
        status = Trim$(f(cStatus))
        statusDate = ""
        If cStatusDate >= 0 And UBound(f) >= cStatusDate Then statusDate = CrownMoreDate(f(cStatusDate))
        dueAmount = CrownMoreMoney(f(cAmount))
        dueDate = ""
        cancelDate = ""
        note = ""
        matchedOn = ""

        If StrComp(Replace(status, " ", ""), "In-Force", vbTextCompare) = 0 Then
            dueDate = CrownTrexisDueDate(drv, policyNo, note)
        Else
            cancelDate = statusDate
        End If

        recordId = CrownMoreMatch(byNumber, byName, TX_CARRIER, policyNo, insured, matchedOn, note, listed)

        ws.Cells(row, 2).value = policyNo
        ws.Cells(row, 3).value = insured
        ws.Cells(row, 4).value = dueAmount
        ws.Cells(row, 5).value = dueDate
        ws.Cells(row, 6).value = cancelDate
        ws.Cells(row, 7).value = status
        ws.Cells(row, 8).value = Now
        ws.Cells(row, 9).value = matchedOn

        If Len(recordId) = 0 Then
            If Len(note) = 0 Then note = "no policy on the website with this number or this name"
            ws.Cells(row, 10).value = note
            unmatched = unmatched + 1
        Else
            ws.Cells(row, 1).value = recordId
            answer = CrownMoreWrite(recordId, dueAmount, dueDate, cancelDate, status)

            If InStr(1, answer, """ok"":true", vbTextCompare) > 0 Then
                ws.Cells(row, 10).value = Trim$("written. " & note)
                done = done + 1
            ElseIf Len(Trim$(answer)) = 0 Then
                ws.Cells(row, 10).value = "the website did not answer - run again for this one"
                failed = failed + 1
            Else
                ws.Cells(row, 10).value = "refused: " & answer
                failed = failed + 1
            End If
        End If

        row = row + 1
NextTrexis:
        DoEvents
    Next i

Finish:
    On Error Resume Next
    drv.Quit
    On Error GoTo 0

    ws.Columns("A:J").AutoFit
    ThisWorkbook.Save
    ws.Activate

    MsgBox done & " Trexis policies written to the website." & vbCrLf & _
           unmatched & " could not be matched to a policy on the website." & vbCrLf & _
           failed & " did not go through." & vbCrLf & vbCrLf & _
           "The TrexisPayments sheet has every one, and how it was matched.", _
           vbInformation, "Crown Superior"
End Sub


' The "Current Payment Due" date off one policy's own page: find it on the
' list (25 to a page), open it, read it, go back to the list.
Private Function CrownTrexisDueDate(ByVal drv As ChromeDriver, ByVal policyNo As String, ByRef note As String) As String
    Dim pages As Long, p As Long
    Dim answer As String

    CrownMoreWaitJs drv, "return document.getElementById('policiesListTable')?'ok':'';", 20

    pages = Val(CrownMoreJs(drv, "var s=document.getElementById('policiesPage');return s?String(s.options.length):'1';"))
    If pages < 1 Then pages = 1

    For p = 1 To pages
        If p > 1 Or pages > 1 Then
            CrownMoreJs drv, "var s=document.getElementById('policiesPage');if(s&&s.options.length>=" & p & "){" & _
                             "s.selectedIndex=" & (p - 1) & ";s.dispatchEvent(new Event('change',{bubbles:true}));}return '';"
            drv.Wait 1500
        End If

        answer = CrownMoreJs(drv, _
            "var w=arguments[0],td=document.querySelectorAll('#policiesListTable td.policy-number');" & _
            "for(var i=0;i<td.length;i++){if((td[i].innerText||'').trim()===w){td[i].click();return 'clicked';}}return '';", policyNo)

        If answer = "clicked" Then Exit For
    Next p

    If answer <> "clicked" Then
        note = "not on the Trexis list to open - no due date"
        Exit Function
    End If

    If CrownMoreWaitJs(drv, "return (location.href.indexOf('#PolicyDetails')>=0&&" & _
                            "document.evaluate(""//b[normalize-space(text())='Current Payment Due']/parent::div/following-sibling::div""," & _
                            "document,null,XPathResult.FIRST_ORDERED_NODE_TYPE,null).singleNodeValue)?'ok':'';", 20) Then
        answer = CrownMoreJs(drv, _
            "var d=document.evaluate(""//b[normalize-space(text())='Current Payment Due']/parent::div/following-sibling::div""," & _
            "document,null,XPathResult.FIRST_ORDERED_NODE_TYPE,null).singleNodeValue;return d?(d.innerText||'').trim():'';")
        CrownTrexisDueDate = CrownMoreDate(answer)

        If Len(CrownTrexisDueDate) = 0 Then note = "the policy page had no Current Payment Due date"
    Else
        note = "the policy page did not open - no due date"
    End If

    ' Back to the list, as his macro does.
    CrownMoreJs drv, "var b=document.getElementById('searchPolicy');if(b)b.click();return '';"

    If Not CrownMoreWaitJs(drv, "return document.getElementById('policiesListTable')?'ok':'';", 15) Then
        CrownMoreJs drv, "var b=document.getElementById('searchPolicyQuote');if(b)b.click();return '';"
        CrownMoreWaitJs drv, "return document.getElementById('policiesListTable')?'ok':'';", 20
    End If
End Function


' The newest .csv in Downloads written after a moment, waited for.
Private Function CrownMoreNewDownload(ByVal after As Date, ByVal seconds As Long) As String
    Dim fso As Object, file As Object
    Dim folder As String, best As String, bestAt As Date
    Dim started As Date

    Set fso = CreateObject("Scripting.FileSystemObject")
    folder = Environ$("USERPROFILE") & "\Downloads"
    started = Now

    Do
        best = ""
        bestAt = after - 1 / 86400

        For Each file In fso.GetFolder(folder).Files
            If LCase$(fso.GetExtensionName(file.Name)) = "csv" And file.DateLastModified >= bestAt Then
                best = file.Path
                bestAt = file.DateLastModified
            End If
        Next file

        If Len(best) > 0 Then
            Application.Wait Now + TimeSerial(0, 0, 2)   ' let it finish writing
            CrownMoreNewDownload = best
            Exit Function
        End If

        Application.Wait Now + TimeSerial(0, 0, 1)
        DoEvents
    Loop While (Now - started) * 86400 < seconds
End Function

Private Function CrownMoreReadText(ByVal path As String) As String
    Dim stream As Object

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile path
    CrownMoreReadText = stream.ReadText
    stream.Close

    ' A byte-order mark in front of the first heading.
    If Left$(CrownMoreReadText, 1) = ChrW$(&HFEFF) Then CrownMoreReadText = Mid$(CrownMoreReadText, 2)
End Function

Private Function CrownMoreColumn(ByRef head() As String, ByVal wanted As String) As Long
    Dim i As Long

    For i = LBound(head) To UBound(head)
        If StrComp(Trim$(head(i)), wanted, vbTextCompare) = 0 Then
            CrownMoreColumn = i
            Exit Function
        End If
    Next i

    CrownMoreColumn = -1
End Function

' True once the script answers "ok", within the time.
Private Function CrownMoreWaitJs(ByVal drv As Object, ByVal script As String, ByVal seconds As Long) As Boolean
    Dim started As Date

    started = Now

    Do
        If CrownMoreJs(drv, script) = "ok" Then
            CrownMoreWaitJs = True
            Exit Function
        End If

        drv.Wait 700
        DoEvents
    Loop While (Now - started) * 86400 < seconds
End Function


' ---------------------------------------------------------------------
' Shared by the carrier runs in this module
' ---------------------------------------------------------------------

' The website's policies for one carrier, by number and by name.
Private Sub CrownMoreIndex(ByVal carrier As String, ByRef byNumber As Object, ByRef byName As Object)
    Dim rows As Collection, parts As Variant
    Dim digits As String, nameKey As String
    Dim list As Variant

    Set byNumber = CreateObject("Scripting.Dictionary")
    Set byName = CreateObject("Scripting.Dictionary")
    Set rows = CrownPolicyRows(carrier)

    ' 0 record number, 2 policy number, 3 carrier, 4 first, 5 last
    For Each parts In rows
        digits = CrownDigits(CStr(parts(2)))
        If Len(digits) > 0 And Not byNumber.Exists(digits) Then byNumber.Add digits, parts

        ' Also by the part before a suffix: the website has "RGA1234761-2"
        ' where Assurance America may show "RGA1234761-02". Only a long
        ' enough part - Trexis numbers start "14-10-".
        If InStr(CStr(parts(2)), "-") > 0 Then
            digits = CrownDigits(Left$(CStr(parts(2)), InStr(CStr(parts(2)), "-") - 1))
            If Len(digits) >= 6 And Not byNumber.Exists(digits) Then byNumber.Add digits, parts
        End If

        nameKey = CrownMoreNameKey(CStr(parts(4)), CStr(parts(5)))

        If Len(nameKey) > 1 Then
            If byName.Exists(nameKey) Then
                list = byName(nameKey)
                ReDim Preserve list(UBound(list) + 1)
                list(UBound(list)) = parts
                byName(nameKey) = list
            Else
                byName.Add nameKey, Array(parts)
            End If
        End If
    Next parts
End Sub

Private Function CrownMoreNameKey(ByVal first As String, ByVal last As String) As String
    first = UCase$(Trim$(first))
    last = UCase$(Trim$(last))

    ' First word of the first name only - "MARY ANN" and "MARY" are the
    ' same customer often enough, a middle name never decides it.
    If InStr(first, " ") > 0 Then first = Left$(first, InStr(first, " ") - 1)

    CrownMoreNameKey = first & "|" & last
End Function

' The website record for a carrier's policy, or "" with the reason in note.
Private Function CrownMoreMatch(ByVal byNumber As Object, ByVal byName As Object, ByVal carrier As String, _
                                ByVal policyNo As String, ByVal insured As String, _
                                ByRef matchedOn As String, ByRef note As String, _
                                Optional ByVal listed As Object = Nothing) As String
    Dim digits As String, first As String, last As String, nameKey As String
    Dim list As Variant, parts As Variant
    Dim j As Long, found As String, words() As String

    digits = CrownDigits(policyNo)

    If Len(digits) > 0 And byNumber.Exists(digits) Then
        parts = byNumber(digits)
        CrownMoreMatch = CStr(parts(0))
        matchedOn = "policy number"
        Exit Function
    End If

    If InStr(policyNo, "-") > 0 Then
        digits = CrownDigits(Left$(policyNo, InStr(policyNo, "-") - 1))

        If Len(digits) >= 6 And byNumber.Exists(digits) Then
            parts = byNumber(digits)
        CrownMoreMatch = CStr(parts(0))
            matchedOn = "policy number before the -"
            Exit Function
        End If

        digits = CrownDigits(Mid$(policyNo, InStrRev(policyNo, "-") + 1))

        If Len(digits) >= 5 And byNumber.Exists(digits) Then
            parts = byNumber(digits)
        CrownMoreMatch = CStr(parts(0))
            matchedOn = "policy number after the -"
            Exit Function
        End If
    End If

    ' The name: "LAST, FIRST" or "FIRST LAST".
    insured = Trim$(insured)

    If InStr(insured, ",") > 0 Then
        last = Trim$(Left$(insured, InStr(insured, ",") - 1))
        first = Trim$(Mid$(insured, InStr(insured, ",") + 1))
    ElseIf InStr(insured, " ") > 0 Then
        words = Split(Application.WorksheetFunction.Trim(insured), " ")
        first = words(0)
        last = words(UBound(words))
    End If

    nameKey = CrownMoreNameKey(first, last)

    If byName.Exists(nameKey) Then
        list = byName(nameKey)

        For j = LBound(list) To UBound(list)
            parts = list(j)

            ' 19 Oct: a customer with two Trexis policies; the website held
            ' only the older. The newer one matched him by name and wrote its
            ' figures over the older one's. A record whose own number is ALSO
            ' on the carrier's list this run belongs to that policy, not this.
            If Not listed Is Nothing Then
                If listed.Exists(CrownDigits(CStr(parts(2)))) And CrownDigits(CStr(parts(2))) <> CrownDigits(policyNo) Then
                    note = "the website record for this name (" & parts(2) & ") is a different " & carrier & _
                           " policy, also on " & carrier & "'s list - not overwritten"
                    GoTo NextCandidate
                End If
            End If

            If CrownSameCarrier(carrier, CStr(parts(3))) Then
                If Len(found) > 0 Then
                    note = "that name has more than one " & carrier & " policy on the website - not guessed"
                    Exit Function
                End If

                found = CStr(parts(0))
            End If
NextCandidate:
        Next j

        If Len(found) > 0 Then
            note = ""
            CrownMoreMatch = found
            matchedOn = "name (the number did not match)"
        End If
    End If
End Function

' The same fields the United and Verve runs write.
Private Function CrownMoreWrite(ByVal recordId As String, ByVal amount As String, ByVal dueDate As String, _
                                ByVal cancelDate As String, ByVal status As String) As String
    CrownMoreWrite = CrownApiUpdate(CROWN_FORM_POLICY, CLng(recordId), _
        Array("Web_pymnt_due", "updated_due_date", "updated_cancel_date", "Status_", _
              "Update_due_dates", "Last_Updated"), _
        Array(amount, dueDate, cancelDate, status, "Yes", Format$(Now, "mm-dd-yyyy hh:mm AM/PM")))
End Function

' "10/25/2026" or "(10/25/2026)" -> "10-25-2026". No date in it ("Not
' Found", blank) -> "" - never words where a date belongs.
Private Function CrownMoreDate(ByVal text As String) As String
    Dim m As Object

    text = Trim$(Replace(Replace(text, "(", ""), ")", ""))
    If Len(text) = 0 Then Exit Function

    Set m = CreateObject("VBScript.RegExp")
    m.Pattern = "(\d{1,2})/(\d{1,2})/(\d{2,4})"

    If m.Test(text) Then
        Dim y As String

        With m.Execute(text)(0)
            y = .SubMatches(2)
            If Len(y) = 2 Then y = "20" & y
            CrownMoreDate = Format$(DateSerial(CLng(y), CLng(.SubMatches(0)), CLng(.SubMatches(1))), "mm-dd-yyyy")
        End With
    End If
End Function

' "$1,234.50" -> "1234.50"; "(12.00)" or "-$12.00" -> "-12.00".
Private Function CrownMoreMoney(ByVal text As String) As String
    Dim neg As Boolean, kept As String, i As Long, ch As String

    text = Trim$(text)
    If Len(text) = 0 Then Exit Function

    neg = (InStr(text, "-") > 0 Or InStr(text, "(") > 0)

    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If (ch >= "0" And ch <= "9") Or ch = "." Then kept = kept & ch
    Next i

    If Len(kept) = 0 Then Exit Function

    CrownMoreMoney = Replace(Format$(Val(kept), "0.00"), ",", ".")
    If neg Then CrownMoreMoney = "-" & CrownMoreMoney
End Function

' Run a script, hand back its answer as text ("" on any error).
Private Function CrownMoreJs(ByVal drv As Object, ByVal script As String, Optional ByVal arg As String = "") As String
    Dim answer As Variant

    On Error Resume Next
    If Len(arg) > 0 Then
        answer = drv.ExecuteScript(script, arg)
    Else
        answer = drv.ExecuteScript(script)
    End If
    If Err.Number <> 0 Then answer = ""
    On Error GoTo 0

    If IsNull(answer) Or IsEmpty(answer) Then answer = ""
    CrownMoreJs = CStr(answer)
End Function


' =====================================================================
' Carriers that ask for a verification code: Progressive, First
' Acceptance, National General
' ---------------------------------------------------------------------
' Alt+F8 > CrownProgressivePaymentDue > Run
' Alt+F8 > CrownFirstAcceptancePaymentDue > Run
' Alt+F8 > CrownNatGenPaymentDue > Run
' Alt+F8 > CrownNatGenLook / CrownProgressiveLook / CrownFirstAcceptanceLook
'          (save that carrier's pages for me)
'
' Each carrier gets its own Chrome profile - a folder under
' %LOCALAPPDATA%\CrownSuperior\ - so Chrome remembers the carrier's
' "don't ask again on this device" between runs and the code is asked for
' far less often. When a code IS asked for, the run stops, a box asks for
' the code that was just texted or emailed, types it in (and ticks
' "don't ask again" when there is one) and carries on.
' =====================================================================

' A Chrome that keeps its cookies between runs, one folder per carrier.
Public Function CrownCodeDriver(ByVal carrier As String, ByRef clsDrv As Chrm) As ChromeDriver
    Dim folder As String

    folder = Environ$("LOCALAPPDATA") & "\CrownSuperior"
    On Error Resume Next
    MkDir folder
    folder = folder & "\Chrome-" & Replace(carrier, " ", "")
    MkDir folder
    On Error GoTo 0

    Dim d As ChromeDriver

    Set clsDrv = New Chrm
    Set clsDrv.ChrmDriver = New ChromeDriver
    Set d = clsDrv.ChrmDriver

    On Error Resume Next
    d.AddArgument "--user-data-dir=" & folder
    d.AddArgument "--force-device-scale-factor=0.70"
    d.Start
    On Error GoTo 0

    Set CrownCodeDriver = d
End Function

' Waits (up to the seconds given) for readyJs to answer "ok", getting past
' a verification code on the way if the carrier asks for one. False when
' it never got there, or Tony pressed Cancel.
Private Function CrownCodeWait(ByVal drv As ChromeDriver, ByVal carrier As String, _
                               ByVal readyJs As String, ByVal seconds As Long) As Boolean
    Dim started As Date, answer As String, typed As String, hint As String
    Dim tries As Long, how As VbMsgBoxResult

    started = Now

    Do
        If CrownMoreJs(drv, readyJs) = "ok" Then
            CrownCodeWait = True
            Exit Function
        End If

        answer = CrownMoreJs(drv, CrownCodePageJs())

        If answer = "choose" Then
            how = MsgBox(carrier & " wants to send you a verification code." & vbCrLf & vbCrLf & _
                         "Yes = send it by TEXT message" & vbCrLf & _
                         "No = send it by EMAIL" & vbCrLf & _
                         "Cancel = stop the run", vbYesNoCancel + vbQuestion, "Crown Superior")
            If how = vbCancel Then Exit Function

            If CrownMoreJs(drv, CrownCodeChooseJs(), IIf(how = vbYes, "text", "email")) <> "ok" Then
                MsgBox "Please click how you want the code sent in the " & carrier & " window, then press OK here.", _
                       vbInformation, "Crown Superior"
            End If

            drv.Wait 3000
            started = Now

        ElseIf Left$(answer, 5) = "code|" Then
            tries = tries + 1

            If tries > 3 Then
                MsgBox carrier & " did not take the code three times, so the run has stopped." & vbCrLf & _
                       "Nothing has been changed on the website.", vbExclamation, "Crown Superior"
                Exit Function
            End If

            hint = Left$(Mid$(answer, 6), 200)
            typed = Trim$(InputBox(carrier & " has sent you a verification code." & vbCrLf & hint & vbCrLf & vbCrLf & _
                                   IIf(tries > 1, "That code was not accepted - check it and type it again." & vbCrLf & vbCrLf, "") & _
                                   "Type the code here and press OK (Cancel stops the run):", _
                                   "Crown Superior - " & carrier & " code"))
            If Len(typed) = 0 Then Exit Function

            On Error Resume Next
            drv.FindElementByCss("[data-crown-code]").Clear
            drv.FindElementByCss("[data-crown-code]").SendKeys typed
            On Error GoTo 0

            CrownMoreJs drv, CrownCodeSubmitJs()

            ' Give the carrier time to move on before looking again, so a slow
            ' page is not taken for a wrong code.
            Dim waited As Date
            waited = Now
            Do
                drv.Wait 1000
                DoEvents
                If CrownMoreJs(drv, readyJs) = "ok" Then Exit Do
                If Left$(CrownMoreJs(drv, CrownCodePageJs()), 5) <> "code|" Then Exit Do
            Loop While (Now - waited) * 86400 < 15

            started = Now
        End If

        drv.Wait 800
        DoEvents
    Loop While (Now - started) * 86400 < seconds
End Function

' "choose" on a "how should we send the code" page; "code|<what the page
' says>" on a page asking for the code (the box gets data-crown-code);
' "" otherwise.
Private Function CrownCodePageJs() As String
    CrownCodePageJs = _
        "var vis=function(e){if(!e)return false;var r=e.getBoundingClientRect(),s=getComputedStyle(e);" & _
        "return r.width>0&&r.height>0&&s.visibility!=='hidden'&&s.display!=='none';};" & _
        "var t=(document.body?document.body.innerText:'').replace(/\s+/g,' ');" & _
        "if(/choose a multi.?factor contact method/i.test(t))return 'choose';" & _
        "var ins=[].slice.call(document.querySelectorAll('input')).filter(function(i){" & _
        "var ty=(i.getAttribute('type')||'text').toLowerCase();return vis(i)&&['text','tel','number','password'].indexOf(ty)>=0;});" & _
        "var labelOf=function(i){var l='';[].slice.call(document.querySelectorAll('label')).forEach(function(x){if(i.id&&x.htmlFor===i.id)l+=' '+x.innerText;});return l;};" & _
        "var c=null;for(var k=0;k<ins.length;k++){var i=ins[k];" & _
        "var s=[i.placeholder,i.name,i.id,i.getAttribute('aria-label'),i.getAttribute('autocomplete'),labelOf(i)].join(' ').toLowerCase();" & _
        "if(/code|passcode|otp|one-time|verification|token/.test(s)){c=i;break;}}" & _
        "if(!c&&ins.length===1&&/(verification code|one-time passcode|security code|passcode was sent|enter the code)/i.test(t))c=ins[0];" & _
        "if(!c)return '';" & _
        "[].slice.call(document.querySelectorAll('[data-crown-code]')).forEach(function(e){e.removeAttribute('data-crown-code');});" & _
        "c.setAttribute('data-crown-code','1');" & _
        "var m=t.match(/[^.]*(was sent|was just sent|sent to|text message|via email)[^.]*\./i);" & _
        "return 'code|'+(m?m[0].trim():'');"
End Function

' Clicks the "text message" or the "email" choice (arguments[0]).
Private Function CrownCodeChooseJs() As String
    CrownCodeChooseJs = _
        "var re=arguments[0]==='email'?/send code to|e-?mail/i:/text message/i;" & _
        "var best=null;[].slice.call(document.querySelectorAll('a,button,li,label,div,span,p')).forEach(function(e){" & _
        "var x=(e.innerText||'').trim();if(x&&x.length<160&&re.test(x)&&e.getBoundingClientRect().width>0){" & _
        "if(!best||x.length<=(best.innerText||'').trim().length)best=e;}});" & _
        "if(!best)return 'none';" & _
        "var go=best.closest('a,button,[role=button]')||best.querySelector('a,button,[role=button]')||best.closest('li,label')||best;" & _
        "go.click();return 'ok';"
End Function

' Ticks "don't ask again / remember this device" and presses the button.
Private Function CrownCodeSubmitJs() As String
    CrownCodeSubmitJs = _
        "var vis=function(e){var r=e.getBoundingClientRect();return r.width>0&&r.height>0;};" & _
        "[].slice.call(document.querySelectorAll('input[type=checkbox]')).forEach(function(c){" & _
        "var l=c.closest('label')?c.closest('label').innerText:'';" & _
        "[].slice.call(document.querySelectorAll('label')).forEach(function(x){if(c.id&&x.htmlFor===c.id)l+=' '+x.innerText;});" & _
        "if(/ask again|remember|trust/i.test(l)&&!c.checked)c.click();});" & _
        "var bs=[].slice.call(document.querySelectorAll('button,input[type=submit],a')).filter(vis);" & _
        "for(var k=0;k<bs.length;k++){var x=(bs[k].innerText||bs[k].value||'').trim();" & _
        "if(/^(verify|continue|submit|sign in|log in|next)$/i.test(x)){bs[k].click();return 'ok';}}" & _
        "var f=document.querySelector('[data-crown-code]');" & _
        "if(f&&f.form){if(f.form.requestSubmit)f.form.requestSubmit();else f.form.submit();return 'ok';}" & _
        "return 'nobutton';"
End Function

' A date in any of the ways these carriers write one -> "mm-dd-yyyy".
' "11/05/2026", "Nov 5, 2026", "November 5" (no year: the next one).
Private Function CrownCodeDate(ByVal text As String) As String
    Dim m As Object, mon As Long, d As Long, y As Long, dt As Date

    CrownCodeDate = CrownMoreDate(text)
    If Len(CrownCodeDate) > 0 Then Exit Function

    Set m = CreateObject("VBScript.RegExp")
    m.IgnoreCase = True
    m.Pattern = "(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?\s+(\d{1,2})(?:st|nd|rd|th)?(?:,?\s+(\d{4}))?"

    If Not m.Test(text) Then Exit Function

    With m.Execute(text)(0)
        mon = (InStr(1, "janfebmaraprmayjunjulaugsepoctnovdec", LCase$(.SubMatches(0)), vbTextCompare) + 2) \ 3
        d = CLng(.SubMatches(1))

        If Len(.SubMatches(2)) > 0 Then
            y = CLng(.SubMatches(2))
        Else
            y = Year(Date)
            If DateSerial(y, mon, d) < Date - 180 Then y = y + 1
        End If
    End With

    dt = DateSerial(y, mon, d)
    CrownCodeDate = Format$(dt, "mm-dd-yyyy")
End Function

' Shared ending of each policy for the code carriers: match, write, report.
Private Sub CrownCodeRow(ByVal ws As Object, ByVal row As Long, ByVal carrier As String, _
                         ByVal byNumber As Object, ByVal byName As Object, ByVal listed As Object, _
                         ByVal policyNo As String, ByVal insured As String, _
                         ByVal dueAmount As String, ByVal dueDate As String, ByVal cancelDate As String, _
                         ByVal status As String, ByVal said As String, ByVal note As String, _
                         ByRef done As Long, ByRef unmatched As Long, ByRef blank As Long, ByRef failed As Long)
    Dim recordId As String, matchedOn As String, answer As String

    recordId = CrownMoreMatch(byNumber, byName, carrier, policyNo, insured, matchedOn, note, listed)

    ws.Cells(row, 2).value = policyNo
    ws.Cells(row, 3).value = insured
    ws.Cells(row, 4).value = dueAmount
    ws.Cells(row, 5).value = dueDate
    ws.Cells(row, 6).value = cancelDate
    ws.Cells(row, 7).value = status
    ws.Cells(row, 8).value = Now
    ws.Cells(row, 9).value = matchedOn
    ws.Cells(row, 11).value = said

    If Len(recordId) = 0 Then
        If Len(note) = 0 Then note = "no policy on the website with this number or this name"
        ws.Cells(row, 10).value = note
        unmatched = unmatched + 1
    ElseIf Len(dueAmount & dueDate & cancelDate) = 0 Then
        ws.Cells(row, 1).value = recordId
        ws.Cells(row, 10).value = Trim$("no amount or date read - left as it was. " & note)
        blank = blank + 1
    Else
        ws.Cells(row, 1).value = recordId
        answer = CrownMoreWrite(recordId, dueAmount, dueDate, cancelDate, status)

        If InStr(1, answer, """ok"":true", vbTextCompare) > 0 Then
            ws.Cells(row, 10).value = Trim$("written. " & note)
            done = done + 1
        ElseIf Len(Trim$(answer)) = 0 Then
            ws.Cells(row, 10).value = "the website did not answer - run again for this one"
            failed = failed + 1
        Else
            ws.Cells(row, 10).value = "refused: " & answer
            failed = failed + 1
        End If
    End If
End Sub

Private Function CrownCodeSheet(ByVal sheetName As String, ByVal carrierShort As String) As Object
    Dim ws As Object

    Set ws = CrownSheet(sheetName)
    ws.Cells.ClearContents
    ws.Columns(2).NumberFormat = "@"
    ws.Range("A1:K1").value = Array("Record number", "Policy at " & carrierShort, "Insured", "Amount due", _
                                    "Due date", "Cancel date", "Status", "When", "Matched on", "What happened", _
                                    carrierShort & " said")
    Set CrownCodeSheet = ws
End Function

Private Sub CrownCodeDone(ByVal ws As Object, ByVal carrier As String, ByVal sheetName As String, _
                          ByVal done As Long, ByVal unmatched As Long, ByVal blank As Long, ByVal failed As Long)
    ws.Columns("A:J").AutoFit
    ws.Columns("K").ColumnWidth = 60
    ThisWorkbook.Save
    ws.Activate

    MsgBox done & " " & carrier & " policies written to the website." & vbCrLf & _
           unmatched & " could not be matched to a policy on the website." & vbCrLf & _
           blank & " showed no amount or date - left as they were." & vbCrLf & _
           failed & " did not go through." & vbCrLf & vbCrLf & _
           "The " & sheetName & " sheet has every one, and how it was matched.", _
           vbInformation, "Crown Superior"
End Sub


' ---------------------------------------------------------------------
' Progressive (ForAgentsOnly) - the steps of his
' ScrapeForAGentonlypaymentduedate: URL_10, USER_4 / PASS_4, "Active
' Policies", each customer's page, the "$X due DATE" line.
' ---------------------------------------------------------------------
Public Sub CrownProgressivePaymentDue()
    Dim drv As ChromeDriver, clsDrv As Chrm
    Dim Keys As New Selenium.Keys
    Dim wsInput As Worksheet, ws As Object
    Dim byNumber As Object, byName As Object, listed As Object
    Dim answer As String, parts() As String, titles() As String, says() As String
    Dim people As Collection, person As Variant, links As Boolean
    Dim pageNo As Long, i As Long, k As Long, n As Long, row As Long
    Dim insured As String, policyNo As String, said As String, note As String
    Dim dueAmount As String, dueDate As String
    Dim done As Long, unmatched As Long, blank As Long, failed As Long
    Dim started As Date

    Set wsInput = ThisWorkbook.Worksheets("Input")

    CrownMoreIndex PG_CARRIER, byNumber, byName

    If byNumber.Count = 0 And byName.Count = 0 Then
        MsgBox "No Progressive policies came back from the website.", vbExclamation, "Crown Superior"
        Exit Sub
    End If

    Set ws = CrownCodeSheet("ProgressivePayments", "Progressive")
    row = 2

    Set drv = CrownCodeDriver(PG_CARRIER, clsDrv)
    If Not CrownChromeOk(drv) Then Exit Sub

    If Not CrownPgSignIn(drv) Then GoTo Finish

    If Not CrownPgSearch(drv) Then
        MsgBox "Progressive's list of active policies did not come up, so nothing has been changed.", _
               vbExclamation, "Crown Superior"
        GoTo Finish
    End If

    ' Every customer on every page of the list, with the link to their page.
    Set people = New Collection
    links = True
    pageNo = 1

    Do
        answer = CrownMoreJs(drv, CrownPgListJs())

        If Len(answer) > 0 Then
            For Each person In Split(answer, "~~")
                parts = Split(person, "||")
                people.Add Array(parts(0), parts(1), pageNo, CLng(parts(2)))
                If Len(parts(1)) = 0 Then links = False
            Next person
        End If

        If CrownMoreJs(drv, CrownPgNextJs()) <> "ok" Then Exit Do

        pageNo = pageNo + 1
        If Not CrownMoreWaitJs(drv, "return ((document.querySelector('.paginate_button.current')||{}).innerText||'').trim()==='" & pageNo & "'?'ok':'';", 20) Then Exit Do
        drv.Wait 800
    Loop While pageNo < 200

    If people.Count = 0 Then
        MsgBox "Progressive's active policy list came up empty, so nothing has been changed.", vbInformation, "Crown Superior"
        GoTo Finish
    End If

    ' Pass 1: every policy number, so a name match never lands on a record
    ' whose own policy is also here (the Trexis lesson). Pass 2 writes.
    Dim found As Collection
    Set found = New Collection
    Set listed = CreateObject("Scripting.Dictionary")

    For Each person In people
        insured = person(0)
        note = ""

        If links Then
            drv.Get person(1)
        ElseIf Not CrownPgOpenByClick(drv, person(2), person(3)) Then
            found.Add Array(insured, "", "", "the list did not come back to open this customer")
            GoTo NextPerson
        End If

        answer = ""
        started = Now
        Do
            answer = CrownMoreJs(drv, CrownPgDetailJs())
            If answer = "none" Then Exit Do

            ' The policy is there; give the payment line a few seconds to follow.
            If Left$(answer, 3) = "ok|" Then
                If InStrRev(answer, "|") < Len(answer) Or (Now - started) * 86400 > 8 Then Exit Do
            End If
            drv.Wait 700
            DoEvents
        Loop While (Now - started) * 86400 < 30

        If Left$(answer, 3) <> "ok|" Then
            found.Add Array(insured, "", "", IIf(answer = "none", "Progressive shows no policy for this customer", "the customer's page did not open"))
            GoTo NextPerson
        End If

        parts = Split(answer, "|")
        titles = Split(parts(1), "~")
        says = Split(parts(2) & " ", "~")

        For k = 0 To UBound(titles)
            policyNo = CrownPgPolicyNo(titles(k))
            said = ""
            note = ""

            If UBound(titles) = 0 Then
                ' One policy: its payment line is the one that talks about money.
                said = CrownPgPaymentLine(says)
            ElseIf UBound(titles) = UBound(says) Then
                said = Trim$(says(k))
            ElseIf k = 0 Then
                said = CrownPgPaymentLine(says)
                note = "Progressive shows " & (UBound(titles) + 1) & " policies for this customer - check the payment belongs to this one"
            Else
                note = "Progressive shows " & (UBound(titles) + 1) & " policies for this customer - no payment line read for this one"
            End If
            found.Add Array(insured, policyNo, said, note)
            If Len(CrownDigits(policyNo)) > 0 Then listed(CrownDigits(policyNo)) = True
        Next k

NextPerson:
        If Not links Then
            drv.GoBack
            drv.Wait 1500
        End If
        DoEvents
    Next person

    For Each person In found
        insured = person(0): policyNo = person(1): said = person(2): note = person(3)
        dueAmount = "": dueDate = ""

        If Len(said) > 0 Then
            If InStr(said, "$") > 0 Then dueAmount = CrownMoreMoney(Split(Mid$(said, InStr(said, "$")), " ")(0))

            If InStr(1, said, " due", vbTextCompare) > 0 Then
                dueDate = CrownCodeDate(Mid$(said, InStr(1, said, " due", vbTextCompare) + 4))
            ElseIf InStr(1, said, "withdrawn", vbTextCompare) > 0 Then
                dueDate = CrownCodeDate(Mid$(said, InStr(1, said, "withdrawn", vbTextCompare) + 9))
            End If
            If Len(dueDate) = 0 Then dueDate = CrownCodeDate(said)
        End If

        If Len(policyNo) = 0 And Len(note) > 0 Then
            ws.Cells(row, 3).value = insured
            ws.Cells(row, 8).value = Now
            ws.Cells(row, 10).value = note
            failed = failed + 1
        Else
            CrownCodeRow ws, row, PG_CARRIER, byNumber, byName, listed, policyNo, insured, _
                         dueAmount, dueDate, "", "Active", said, note, done, unmatched, blank, failed
        End If

        row = row + 1
        DoEvents
    Next person

Finish:
    On Error Resume Next
    drv.Quit
    On Error GoTo 0

    CrownCodeDone ws, PG_CARRIER, "ProgressivePayments", done, unmatched, blank, failed
End Sub

' Signs in to Progressive (USER_4 / PASS_4 at URL_10), past the code,
' to the policy search. Says so and answers False when it cannot.
Public Function CrownPgSignIn(ByVal drv As ChromeDriver) As Boolean
    Dim Keys As New Selenium.Keys
    Dim wsInput As Worksheet
    Dim answer As String, started As Date

    Set wsInput = ThisWorkbook.Worksheets("Input")

    drv.Get wsInput.Range("URL_10").value

    ' Sign in - unless Chrome is still signed in from last time.
    started = Now
    Do
        answer = CrownMoreJs(drv, "return document.getElementById('sbp-search')?'in':(document.getElementById('user1')?'login':'');")
        If Len(answer) > 0 Then Exit Do
        If Len(CrownMoreJs(drv, CrownCodePageJs())) > 0 Then Exit Do
        drv.Wait 700
        DoEvents
    Loop While (Now - started) * 86400 < 60

    If answer = "login" Then
        EnterData drv, Keys, "user1", wsInput.Range("USER_4").value, "ID"
        EnterData drv, Keys, "password1", wsInput.Range("PASS_4").value, "ID"
        ClickElement drv, "image1", "ID"
        drv.Wait 2000
    End If

    If Not CrownCodeWait(drv, PG_CARRIER, "return document.getElementById('sbp-search')?'ok':'';", 180) Then
        MsgBox "Did not get to Progressive's policy search, so nothing has been changed." & vbCrLf & _
               "(If Progressive asked for something other than a code, tell me what it showed.)", _
               vbExclamation, "Crown Superior"
        Exit Function
    End If

    CrownPgSignIn = True
End Function

' "Active Policies" and Search, then wait for the list (or "no results").
Private Function CrownPgSearch(ByVal drv As ChromeDriver) As Boolean
    CrownMoreJs drv, _
        "var s=document.getElementById('SBP_UserSelectedStatus');if(s){for(var k=0;k<s.options.length;k++){" & _
        "if(s.options[k].text.trim()==='Active Policies'){s.selectedIndex=k;s.dispatchEvent(new Event('change',{bubbles:true}));}}}" & _
        "var b=document.getElementById('sbp-search');if(b)b.click();return 'ok';"

    drv.Wait 2000
    CrownPgSearch = CrownMoreWaitJs(drv, _
        "return document.querySelector('a[data-at=loss-run-customer-search-name-link]')?'ok':'';", 60)
End Function

' This page of the list: "name||link||index~~name||link||index...".
Private Function CrownPgListJs() As String
    CrownPgListJs = _
        "var a=document.querySelectorAll('a[data-at=loss-run-customer-search-name-link]');var out=[];" & _
        "for(var k=0;k<a.length;k++){var h=a[k].getAttribute('href')||'';" & _
        "if(!h||h==='#'||/^javascript:/i.test(h))h='';else h=a[k].href;" & _
        "out.push(a[k].innerText.replace(/\|/g,' ').replace(/~/g,' ').trim()+'||'+h+'||'+k);}return out.join('~~');"
End Function

Private Function CrownPgPageJs() As String
    CrownPgPageJs = "return ((document.querySelector('.paginate_button.current')||{}).innerText||'1').trim();"
End Function

Private Function CrownPgNextJs() As String
    CrownPgNextJs = _
        "var n=document.querySelector('.paginate_button.next:not(.disabled)');if(!n)return 'last';n.click();return 'ok';"
End Function

' "none" when Progressive says there is no policy; "ok|titles|payment
' lines" (each joined with ~) once the page is there; "" while loading.
Private Function CrownPgDetailJs() As String
    CrownPgDetailJs = _
        "if(document.querySelector('h5.noPoliciesFoundMessage'))return 'none';" & _
        "var al=[].slice.call(document.querySelectorAll('p.alert')).filter(function(p){return /previous page/i.test(p.innerText);});" & _
        "if(al.length)return 'none';" & _
        "var ti=[].slice.call(document.querySelectorAll('div.policyTitle p')).map(function(e){return e.innerText.replace(/[|~]/g,' ').trim();});" & _
        "if(!ti.length)return '';" & _
        "var cs=[].slice.call(document.querySelectorAll('div.CSPItem')).map(function(e){return e.innerText.replace(/\s+/g,' ').replace(/[|~]/g,' ').trim();});" & _
        "return 'ok|'+ti.join('~')+'|'+cs.join('~');"
End Function

' The line that says what is due ("$183.01 due ...", "will be
' withdrawn", "No payment is due"), else the first one.
Private Function CrownPgPaymentLine(ByRef says() As String) As String
    Dim k As Long

    For k = 0 To UBound(says)
        If InStr(says(k), "$") > 0 Or InStr(1, says(k), "payment is due", vbTextCompare) > 0 Or _
           InStr(1, says(k), "paid in full", vbTextCompare) > 0 Then
            CrownPgPaymentLine = Trim$(says(k))
            Exit Function
        End If
    Next k

    CrownPgPaymentLine = Trim$(says(0))
End Function

' "Policy 123456789, Auto ..." -> "123456789" (his rule: the last word
' before the first comma).
Private Function CrownPgPolicyNo(ByVal title As String) As String
    Dim w() As String

    If InStr(title, ",") > 0 Then title = Left$(title, InStr(title, ",") - 1)
    w = Split(Application.WorksheetFunction.Trim(title), " ")
    If UBound(w) >= 0 Then CrownPgPolicyNo = w(UBound(w))
End Function

' Only when the list's names are not plain links: back to that page of
' the list (searching again if Progressive forgot it) and click the name.
Private Function CrownPgOpenByClick(ByVal drv As ChromeDriver, ByVal pageNo As Long, ByVal index As Long) As Boolean
    Dim tries As Long

    If Not CrownMoreWaitJs(drv, "return document.querySelector('a[data-at=loss-run-customer-search-name-link]')?'ok':'';", 20) Then
        If Not CrownPgSearch(drv) Then Exit Function
    End If

    ' Past that page (the list keeps the last page shown): search again, which starts at page 1.
    If Val(CrownMoreJs(drv, CrownPgPageJs())) > pageNo Then
        If Not CrownPgSearch(drv) Then Exit Function
    End If

    Do While CrownMoreJs(drv, CrownPgPageJs()) <> CStr(pageNo)
        tries = tries + 1
        If tries > pageNo + 2 Then Exit Function
        If CrownMoreJs(drv, CrownPgNextJs()) <> "ok" Then Exit Function
        drv.Wait 1500
    Loop

    CrownPgOpenByClick = (CrownMoreJs(drv, _
        "var a=document.querySelectorAll('a[data-at=loss-run-customer-search-name-link]');" & _
        "if(a.length<=arguments[0])return '';a[arguments[0]].click();return 'ok';", CStr(index)) = "ok")
    drv.Wait 2000
End Function


' ---------------------------------------------------------------------
' First Acceptance - the steps of his ScrapeFirstAcceptance: URL_31,
' USER_25 / PASS_25, Policy Services, Active & Reinstatable Policies,
' Open Search, each policy's status and its Billing Details.
' ---------------------------------------------------------------------
Public Sub CrownFirstAcceptancePaymentDue()
    Dim drv As ChromeDriver, clsDrv As Chrm
    Dim Keys As New Selenium.Keys
    Dim wsInput As Worksheet, ws As Object
    Dim byNumber As Object, byName As Object, listed As Object
    Dim answer As String, parts() As String
    Dim rowsXp As String, i As Long, n As Long, row As Long
    Dim policyNo As String, insured As String, status As String, said As String, note As String
    Dim dueAmount As String, dueDate As String, cancelDate As String
    Dim done As Long, unmatched As Long, blank As Long, failed As Long
    Dim started As Date

    Set wsInput = ThisWorkbook.Worksheets("Input")

    CrownMoreIndex FA_CARRIER, byNumber, byName

    If byNumber.Count = 0 And byName.Count = 0 Then
        MsgBox "No First Acceptance policies came back from the website.", vbExclamation, "Crown Superior"
        Exit Sub
    End If

    Set ws = CrownCodeSheet("FirstAcceptancePayments", "First Acceptance")
    row = 2

    Set drv = CrownCodeDriver(FA_CARRIER, clsDrv)
    If Not CrownChromeOk(drv) Then Exit Sub

    If Not CrownFaSignIn(drv) Then GoTo Finish

    drv.Wait 2000

    ' The tile is on the page or in one of its frames (his macro assumed
    ' frame 0) - go to wherever it is.
    If Not CrownFaFindFrame(drv, "//div[contains(@class, 'aspect-square') and contains(., 'Policy Services')]", 60) Then
        MsgBox "First Acceptance's Policy Services tile did not show up, so nothing has been changed.", _
               vbExclamation, "Crown Superior"
        GoTo Finish
    End If

    CrownFaClick drv, "//div[contains(@class, 'aspect-square') and contains(., 'Policy Services')]"
    CrownFaClick drv, "//button[contains(., 'Active & Reinstatable Policies')]"
    CrownFaClick drv, "//table/tbody/tr[1]/td[2]//button"
    CrownFaClick drv, "//button[.//span[text()='Open Search']]"

    rowsXp = "//div[@data-component='x-policy-search-dialog']//table/tbody/tr"

    If Not CrownFaWaitX(drv, rowsXp & "[1]/td[2]//button", 60) Then
        MsgBox "First Acceptance's policy search list did not come up, so nothing has been changed.", _
               vbExclamation, "Crown Superior"
        GoTo Finish
    End If

    n = drv.FindElementsByXPath(rowsXp).Count

    ' Every policy number first (the Trexis lesson), then each policy.
    Set listed = CreateObject("Scripting.Dictionary")
    For i = 1 To n
        policyNo = CrownFaText(drv, rowsXp & "[" & i & "]/td[2]//button")
        If Len(CrownDigits(policyNo)) > 0 Then listed(CrownDigits(policyNo)) = True
    Next i

    For i = 1 To n
        policyNo = "": insured = "": status = "": said = "": note = ""
        dueAmount = "": dueDate = "": cancelDate = ""

        If Not CrownFaWaitX(drv, rowsXp & "[" & i & "]/td[2]//button", 30) Then
            ws.Cells(row, 10).value = "the search list did not come back for row " & i & " - stopped here"
            failed = failed + 1
            Exit For
        End If

        policyNo = CrownFaText(drv, rowsXp & "[" & i & "]/td[2]//button")
        insured = Trim$(CrownFaText(drv, rowsXp & "[" & i & "]/td[3]") & " " & CrownFaText(drv, rowsXp & "[" & i & "]/td[4]"))

        CrownFaClick drv, rowsXp & "[" & i & "]/td[2]//button"

        If CrownFaWaitX(drv, "//span[contains(text(),'Policy Status')]", 30) Then
            status = CrownFaText(drv, "//span[contains(text(),'Policy Status')]/following::div[1]")

            CrownFaClick drv, "//button[normalize-space()='Billing Details']"
            CrownFaWaitX drv, "//span[text()='Payment Information']", 20

            answer = CrownMoreJs(drv, CrownFaBillingJs())
            If Left$(answer, 3) = "ok|" Then
                parts = Split(answer, "|")
                dueDate = CrownCodeDate(parts(1))
                dueAmount = CrownMoreMoney(parts(2))
                cancelDate = CrownCodeDate(parts(3))
                said = "Due Date " & parts(1) & "; Current Balance " & parts(2) & "; Cancellation Date " & parts(3)
            Else
                note = "Billing Details did not show Payment Information"
            End If
        Else
            note = "the policy did not open"
        End If

        CrownCodeRow ws, row, FA_CARRIER, byNumber, byName, listed, policyNo, insured, _
                     dueAmount, dueDate, cancelDate, status, said, note, done, unmatched, blank, failed
        row = row + 1

        CrownFaClick drv, "//button[.//span[text()='Open Search']]"
        DoEvents
    Next i

Finish:
    On Error Resume Next
    drv.Quit
    On Error GoTo 0

    CrownCodeDone ws, FA_CARRIER, "FirstAcceptancePayments", done, unmatched, blank, failed
End Sub

' Signs in to First Acceptance (USER_25 / PASS_25 at URL_31), past the
' code, to the Action Hub. Says so and answers False when it cannot.
Public Function CrownFaSignIn(ByVal drv As ChromeDriver) As Boolean
    Dim Keys As New Selenium.Keys
    Dim wsInput As Worksheet
    Dim answer As String, started As Date

    Set wsInput = ThisWorkbook.Worksheets("Input")

    drv.Get wsInput.Range("URL_31").value

    started = Now
    Do
        answer = CrownMoreJs(drv, "return document.getElementById('v-3')?'login':(document.querySelector('iframe')?'in':'');")
        If Len(answer) > 0 Then Exit Do
        If Len(CrownMoreJs(drv, CrownCodePageJs())) > 0 Then Exit Do
        drv.Wait 700
        DoEvents
    Loop While (Now - started) * 86400 < 60

    If answer = "login" Then
        EnterData drv, Keys, "v-3", wsInput.Range("USER_25").value, "ID"
        EnterData drv, Keys, "v-7", wsInput.Range("PASS_25").value, "ID"
        drv.Keyboard.SendKeys Keys.Enter
        drv.Wait 3000
    End If

    ' Signed in = the sign-in boxes have gone and the Action Hub ("Policy
    ' Services") is showing - on the page itself or inside its frame.
    ' (21 Oct: the first version also wanted no box to type in, but the
    ' Action Hub has its own Search box, so it never counted as signed in.)
    If Not CrownCodeWait(drv, FA_CARRIER, _
            "if(document.getElementById('v-7'))return '';var hub=function(d){try{return /Policy Services|Action Hub/i.test(d.body.innerText);}catch(e){return false;}};" & _
            "if(hub(document))return 'ok';var fs=document.querySelectorAll('iframe');" & _
            "for(var k=0;k<fs.length;k++){try{if(hub(fs[k].contentDocument))return 'ok';}catch(e){}}" & _
            "return (fs.length&&/apps\/portal/i.test(location.href)&&!/verification|one-time|passcode|security code/i.test(document.body.innerText))?'ok':'';", 180) Then
        MsgBox "Did not get past First Acceptance's sign-in, so nothing has been changed." & vbCrLf & _
               "(If it asked for something other than a code, send me a screenshot of it.)", _
               vbExclamation, "Crown Superior"
        Exit Function
    End If

    CrownFaSignIn = True
End Function

' The pages a "pull into Edit Data" would read, saved for me - same as
' CrownNatGenLook. Open a customer's policy pages (drivers, vehicles,
' coverages, lienholder) and press OK on each.
' Assurance America and Trexis: sign in the way their payment runs do,
' then save the pages Tony opens (a customer's drivers, vehicles,
' coverages, lienholder) so their "pull into Edit Data" can be built.
Public Sub CrownAssuranceLook()
    Dim drv As ChromeDriver, clsDrv As Chrm
    Dim Keys As New Selenium.Keys
    Dim wsInput As Worksheet

    Set wsInput = ThisWorkbook.Worksheets("Input")
    Set drv = CrownCodeDriver(AA_CARRIER, clsDrv)
    If Not CrownChromeOk(drv) Then Exit Sub

    drv.Get wsInput.Range("URL_11").value

    If CrownMoreWaitJs(drv, "return (document.querySelector(""input[formcontrolname='username']"")&&" & _
                            "document.querySelector(""input[formcontrolname='password']""))?'ok':'';", 60) Then
        EnterData drv, Keys, "//input[@formcontrolname='username']", wsInput.Range("USER_7").value, "XPATH"
        EnterData drv, Keys, "//input[@formcontrolname='password']", wsInput.Range("PASS_7").value, "XPATH"
        ClickElement drv, "//*[@id='login-button']", "XPATH"
        drv.Wait 4000
    End If

    If CrownCodeWait(drv, AA_CARRIER, "return document.querySelector(""input[formcontrolname='password']"")?'':'ok';", 120) Then
        CrownLookPages drv, AA_CARRIER
    Else
        MsgBox "Did not get past Assurance America's sign-in. Send me a screenshot of what it shows.", vbExclamation, "Crown Superior"
    End If

    On Error Resume Next
    drv.Quit
    On Error GoTo 0
End Sub

Public Sub CrownTrexisLook()
    Dim drv As ChromeDriver, clsDrv As Chrm
    Dim Keys As New Selenium.Keys
    Dim wsInput As Worksheet

    Set wsInput = ThisWorkbook.Worksheets("Input")
    Set drv = CrownCodeDriver(TX_CARRIER, clsDrv)
    If Not CrownChromeOk(drv) Then Exit Sub

    drv.Get wsInput.Range("URL_8").value

    If CrownMoreWaitJs(drv, "return (document.getElementById('loginName')&&document.getElementById('loginPassword'))?'ok':'';", 60) Then
        EnterData drv, Keys, "loginName", wsInput.Range("USER_6").value, "ID"
        EnterData drv, Keys, "loginPassword", wsInput.Range("PASS_6").value, "ID"
        drv.Wait 1000
        ClickElement drv, "//*[@id='loginSubmit']/div[1]/button", "XPATH"
    End If

    If CrownCodeWait(drv, TX_CARRIER, "return document.getElementById('agentHomePolicyButton')?'ok':'';", 120) Then
        CrownLookPages drv, TX_CARRIER
    Else
        MsgBox "Did not get past Trexis's sign-in. Send me a screenshot of what it shows.", vbExclamation, "Crown Superior"
    End If

    On Error Resume Next
    drv.Quit
    On Error GoTo 0
End Sub

Public Sub CrownProgressiveLook()
    Dim drv As ChromeDriver, clsDrv As Chrm

    Set drv = CrownCodeDriver(PG_CARRIER, clsDrv)
    If Not CrownChromeOk(drv) Then Exit Sub

    If CrownPgSignIn(drv) Then CrownLookPages drv, PG_CARRIER

    On Error Resume Next
    drv.Quit
    On Error GoTo 0
End Sub

Public Sub CrownFirstAcceptanceLook()
    Dim drv As ChromeDriver, clsDrv As Chrm

    Set drv = CrownCodeDriver(FA_CARRIER, clsDrv)
    If Not CrownChromeOk(drv) Then Exit Sub

    If CrownFaSignIn(drv) Then CrownLookPages drv, FA_CARRIER

    On Error Resume Next
    drv.Quit
    On Error GoTo 0
End Sub

' Leaves the driver in the page or frame that has the element.
Public Function CrownFaFindFrame(ByVal drv As ChromeDriver, ByVal xp As String, ByVal seconds As Long) As Boolean
    Dim started As Date, i As Long, n As Long

    started = Now
    Do
        On Error Resume Next
        drv.SwitchToDefaultContent
        If drv.FindElementsByXPath(xp).Count > 0 Then
            If Err.Number = 0 Then
                CrownFaFindFrame = True
                Exit Function
            End If
        End If
        Err.Clear

        n = drv.FindElementsByTag("iframe").Count
        For i = 0 To n - 1
            drv.SwitchToDefaultContent
            drv.SwitchToFrame i
            If Err.Number = 0 Then
                If drv.FindElementsByXPath(xp).Count > 0 Then
                    If Err.Number = 0 Then
                        CrownFaFindFrame = True
                        Exit Function
                    End If
                End If
            End If
            Err.Clear
        Next i

        drv.SwitchToDefaultContent
        On Error GoTo 0

        drv.Wait 1000
        DoEvents
    Loop While (Now - started) * 86400 < seconds
End Function

Public Function CrownFaWaitX(ByVal drv As ChromeDriver, ByVal xp As String, ByVal seconds As Long) As Boolean
    Dim started As Date

    started = Now
    Do
        On Error Resume Next
        CrownFaWaitX = (drv.FindElementsByXPath(xp).Count > 0)
        On Error GoTo 0
        If CrownFaWaitX Then Exit Function
        drv.Wait 700
        DoEvents
    Loop While (Now - started) * 86400 < seconds
End Function

Private Function CrownFaText(ByVal drv As ChromeDriver, ByVal xp As String) As String
    On Error Resume Next
    CrownFaText = Trim$(Replace(drv.FindElementByXPath(xp).text, Chr(160), " "))
    On Error GoTo 0
End Function

' Click (when it is there) and let First Acceptance's spinner finish.
Public Sub CrownFaClick(ByVal drv As ChromeDriver, ByVal xp As String)
    If CrownFaWaitX(drv, xp, 20) Then
        On Error Resume Next
        drv.ExecuteScript "arguments[0].scrollIntoView({block:'center'});", drv.FindElementByXPath(xp)
        drv.FindElementByXPath(xp).Click
        On Error GoTo 0
    End If

    drv.Wait 800
    WaitSpinnerGone drv
End Sub

' "ok|due date|current balance|cancellation date" from Payment Information.
Private Function CrownFaBillingJs() As String
    CrownFaBillingJs = _
        "var h=[].slice.call(document.querySelectorAll('span')).filter(function(s){return s.innerText.trim()==='Payment Information';})[0];" & _
        "if(!h)return '';var card=h.closest('[class*=card-shadow]')||h.parentNode.parentNode;" & _
        "var x=function(r,p){var n=document.evaluate(p,r,null,9,null).singleNodeValue;return n?n.textContent.replace(/\u00a0/g,' ').trim():'';};" & _
        "var due='',bal='',can='';[].slice.call(card.querySelectorAll('[data-component=x-description-list-item]')).forEach(function(r){" & _
        "var l=x(r,'.//div[1]//span'),v=x(r,'.//div[2]//span').replace(/[|]/g,' ');" & _
        "if(/Due Date/i.test(l))due=v;if(/Current Balance/i.test(l))bal=v;if(/Cancellation Date/i.test(l))can=v;});" & _
        "return 'ok|'+due+'|'+bal+'|'+can;"
End Function


' ---------------------------------------------------------------------
' National General - no payment run yet. This signs in (with the code)
' and saves the pages Tony opens, so the payment run can be built from
' what National General really shows.
' ---------------------------------------------------------------------
Public Sub CrownNatGenLook()
    Dim drv As ChromeDriver, clsDrv As Chrm

    Set drv = CrownCodeDriver(NG_CARRIER, clsDrv)
    If Not CrownChromeOk(drv) Then Exit Sub

    If CrownNgSignIn(drv) Then CrownLookPages drv, NG_CARRIER

    On Error Resume Next
    drv.Quit
    On Error GoTo 0
End Sub

' Signs in to National General (his Mod9FillNationalGeneral steps:
' txtUserID > btnLogin > Password > SIGN IN), getting past the code.
Private Function CrownNgSignIn(ByVal drv As ChromeDriver) As Boolean
    Dim Keys As New Selenium.Keys
    Dim wsInput As Worksheet
    Dim answer As String, url As String
    Dim started As Date

    Set wsInput = ThisWorkbook.Worksheets("Input")

    url = Trim$(wsInput.Range("URL_30").value)
    If Len(url) = 0 Then url = NG_SITE
    drv.Get url

    started = Now
    Do
        answer = CrownMoreJs(drv, "return document.getElementById('txtUserID')?'user':(document.getElementById('Password')?'pass':" & _
                                  "(document.getElementById('ctl00_lblInsName')||document.querySelector('a[href*=PolicySearch]')?'in':''));")
        If Len(answer) > 0 Then Exit Do
        If Len(CrownMoreJs(drv, CrownCodePageJs())) > 0 Then Exit Do
        drv.Wait 700
        DoEvents
    Loop While (Now - started) * 86400 < 30

    If answer = "user" Then
        EnterData drv, Keys, "txtUserID", wsInput.Range("USER_30").value, "ID"
        ClickElement drv, "btnLogin", "ID"
        CrownMoreWaitJs drv, "return document.getElementById('Password')?'ok':'';", 30
        answer = "pass"
    End If

    If answer = "pass" Then
        EnterData drv, Keys, "Password", wsInput.Range("PASS_30").value, "ID"
        ClickElement drv, "//button[contains(text(),'SIGN IN')]", "XPATH"
        drv.Wait 2500
    ElseIf answer = "" Then
        MsgBox "Please sign in to National General in the Chrome window, then press OK here.", _
               vbInformation, "Crown Superior"
    End If

    CrownNgSignIn = CrownCodeWait(drv, NG_CARRIER, _
        "return (location.hostname.indexOf('login.')!==0&&!document.querySelector('input[type=password]')" & _
        "&&!document.getElementById('txtUserID'))?'ok':'';", 180)

    If Not CrownNgSignIn Then
        MsgBox "Did not get past National General's sign-in. Send me a screenshot of what it shows.", _
               vbExclamation, "Crown Superior"
    End If
End Function


' ---------------------------------------------------------------------
' National General payment due - built from the pages Tony saved on
' 21 Oct. National General has no list of our policies to walk through,
' so this goes the other way: every National General policy on the
' website is searched for (Search.aspx, by its number), opened, and its
' Billing page read - "Next Amount Due", "Amount Due On", Status. The
' page header must show that policy's number before anything is written.
' ---------------------------------------------------------------------
Public Sub CrownNatGenPaymentDue()
    Dim drv As ChromeDriver, clsDrv As Chrm
    Dim ws As Object, policies As Collection, parts As Variant
    Dim recordId As String, siteNo As String, digits As String, base As String, suffix As String
    Dim insured As String, answer As String, bits() As String, note As String, cancelledOn As String
    Dim dueAmount As String, dueDate As String, cancelDate As String, status As String, said As String
    Dim row As Long, done As Long, blank As Long, failed As Long, skipped As Long, notFound As Long
    Dim signIns As Long

    Set policies = CrownPolicyRows(NG_CARRIER)

    If policies.Count = 0 Then
        MsgBox "No National General policies came back from the website.", vbExclamation, "Crown Superior"
        Exit Sub
    End If

    Set ws = CrownCodeSheet("NatGenPayments", "National General")
    row = 2

    Set drv = CrownCodeDriver(NG_CARRIER, clsDrv)
    If Not CrownChromeOk(drv) Then Exit Sub

    If Not CrownNgSignIn(drv) Then GoTo Finish
    signIns = 1

    For Each parts In policies
        recordId = CStr(parts(0))
        siteNo = CStr(parts(2))
        insured = Trim$(CStr(parts(4)) & " " & CStr(parts(5)))
        dueAmount = "": dueDate = "": cancelDate = "": status = "": said = "": note = ""

        ws.Cells(row, 1).value = recordId
        ws.Cells(row, 2).value = siteNo
        ws.Cells(row, 3).value = insured
        ws.Cells(row, 8).value = Now

        ' National General numbers are 10 digits, then the term ("02").
        digits = CrownDigits(siteNo)
        base = Left$(digits, 10)
        suffix = ""
        If Len(digits) >= 12 Then suffix = Mid$(digits, 11, 2)

        If Len(digits) < 8 Then
            ws.Cells(row, 10).value = "no usable policy number on the website - skipped"
            failed = failed + 1
            GoTo NextPolicy
        End If

        If CrownNgLongCancelled(parts, cancelledOn) Then
            ws.Cells(row, 10).value = "skipped - cancelled since " & cancelledOn
            skipped = skipped + 1
            GoTo NextPolicy
        End If

        answer = CrownNgOpenBilling(drv, base, suffix)

        ' Signed out part way (National General ends sessions): once more.
        If answer = "signedout" And signIns < 4 Then
            signIns = signIns + 1
            If CrownNgSignIn(drv) Then answer = CrownNgOpenBilling(drv, base, suffix)
        End If

        If Left$(answer, 3) <> "ok|" Then
            If answer = "none" Then
                ws.Cells(row, 10).value = "National General found no policy " & base
                notFound = notFound + 1
            Else
                ws.Cells(row, 10).value = "could not open it on National General (" & answer & ") - left as it was"
                failed = failed + 1
            End If
            GoTo NextPolicy
        End If

        ' ok|next amount|due on|status|status detail|cancel date|header number
        bits = Split(answer, "|")
        dueAmount = CrownMoreMoney(bits(1))
        dueDate = CrownCodeDate(bits(2))
        status = Trim$(bits(3))
        cancelDate = CrownCodeDate(bits(5))
        said = "Next Amount Due " & bits(1) & "; Amount Due On " & bits(2) & "; Status " & bits(3) & _
               IIf(Len(bits(4)) > 0, " (" & bits(4) & ")", "") & IIf(Len(bits(5)) > 0, "; cancel " & bits(5), "") & _
               "; term " & bits(6)

        ws.Cells(row, 2).value = siteNo & "  (NG " & bits(6) & ")"
        ws.Cells(row, 4).value = dueAmount
        ws.Cells(row, 5).value = dueDate
        ws.Cells(row, 6).value = cancelDate
        ws.Cells(row, 7).value = status
        ws.Cells(row, 9).value = "policy number"
        ws.Cells(row, 11).value = said

        If Len(dueAmount & dueDate & cancelDate & status) = 0 Then
            ws.Cells(row, 10).value = "the Billing page showed nothing to read - left as it was"
            blank = blank + 1
        Else
            answer = CrownMoreWrite(recordId, dueAmount, dueDate, cancelDate, status)

            If InStr(1, answer, """ok"":true", vbTextCompare) > 0 Then
                ws.Cells(row, 10).value = "written"
                done = done + 1
            ElseIf Len(Trim$(answer)) = 0 Then
                ws.Cells(row, 10).value = "the website did not answer - run again for this one"
                failed = failed + 1
            Else
                ws.Cells(row, 10).value = "refused: " & answer
                failed = failed + 1
            End If
        End If

NextPolicy:
        row = row + 1
        DoEvents
    Next parts

Finish:
    On Error Resume Next
    drv.Quit
    On Error GoTo 0

    ws.Columns("A:J").AutoFit
    ws.Columns("K").ColumnWidth = 60
    ThisWorkbook.Save
    ws.Activate

    MsgBox done & " National General policies written to the website." & vbCrLf & _
           skipped & " skipped - cancelled more than 60 days ago." & vbCrLf & _
           notFound & " not found on National General." & vbCrLf & _
           blank & " showed nothing to read - left as they were." & vbCrLf & _
           failed & " did not go through." & vbCrLf & vbCrLf & _
           "The NatGenPayments sheet has every one.", vbInformation, "Crown Superior"
End Sub

' Search for the policy, open it (the term the website names, else the
' Active one, else the newest), go to Billing and read it.
' "ok|..." (see CrownNgBillingJs), "none", "signedout" or what went wrong.
Private Function CrownNgOpenBilling(ByVal drv As ChromeDriver, ByVal base As String, ByVal suffix As String) As String
    Dim answer As String, started As Date

    drv.Get NG_SITE & "/Search.aspx?q=Policy"

    started = Now
    Do
        answer = CrownMoreJs(drv, "return document.getElementById('ctl00_MainContent_txtSearchStr')?'ok':" & _
                                  "((document.getElementById('txtUserID')||location.hostname.indexOf('login.')===0)?'signedout':'');")
        If Len(answer) > 0 Then Exit Do
        drv.Wait 700
        DoEvents
    Loop While (Now - started) * 86400 < 30

    If answer <> "ok" Then
        CrownNgOpenBilling = IIf(answer = "signedout", "signedout", "the search page did not load")
        Exit Function
    End If

    CrownMoreJs drv, _
        "var t=document.getElementById('ctl00_MainContent_txtSearchStr');t.value=arguments[0];" & _
        "var a=document.getElementById('ctl00_MainContent_ddlAction');if(a)a.value='0';" & _
        "document.getElementById('ctl00_MainContent_btnSearch').click();return 'ok';", base
    drv.Wait 1500

    ' The list of terms, or straight onto the policy when there is one.
    started = Now
    Do
        answer = CrownMoreJs(drv, CrownNgPickJs(), suffix)
        If Len(answer) > 0 Then Exit Do
        drv.Wait 800
        DoEvents
    Loop While (Now - started) * 86400 < 40

    If answer = "none" Or answer = "" Then
        CrownNgOpenBilling = IIf(answer = "none", "none", "no search answer")
        Exit Function
    End If

    ' On the policy (answer "policy", or "clicked" a term): to Billing.
    If Not CrownMoreWaitJs(drv, "return (/\/Policy\//i.test(location.pathname)&&document.getElementById('ctl00_lblHeaderPageTitleTop'))?'ok':'';", 40) Then
        CrownNgOpenBilling = "the policy did not open"
        Exit Function
    End If

    If Not CrownMoreWaitJs(drv, "return /PolicyBillPlans/i.test(location.pathname)?'ok':'';", 1) Then
        If CrownMoreJs(drv, "var b=document.querySelector('a[href$=""PolicyBillPlans.aspx""]');if(b){b.click();return 'ok';}return '';") <> "ok" Then
            drv.Get NG_SITE & "/Policy/PolicyBillPlans.aspx"
        End If
    End If

    started = Now
    Do
        answer = CrownMoreJs(drv, CrownNgBillingJs())
        If Left$(answer, 3) = "ok|" Then Exit Do
        drv.Wait 800
        DoEvents
    Loop While (Now - started) * 86400 < 40

    If Left$(answer, 3) <> "ok|" Then
        CrownNgOpenBilling = "the Billing page did not load"
        Exit Function
    End If

    ' The header must be THIS policy - never another customer's figures.
    If InStr(CrownDigits(Split(answer, "|")(6)), base) <> 1 Then
        CrownNgOpenBilling = "National General showed policy " & Split(answer, "|")(6) & " instead"
        Exit Function
    End If

    CrownNgOpenBilling = answer
End Function

' On the search results: clicks the right term and answers "clicked";
' "policy" if National General went straight to the policy; "none" when
' it found nothing; "" while still loading. arguments[0] = term ("02") or "".
Private Function CrownNgPickJs() As String
    CrownNgPickJs = _
        "if(/\/Policy\//i.test(location.pathname)&&document.getElementById('ctl00_lblHeaderPageTitleTop'))return 'policy';" & _
        "var ls=[].slice.call(document.querySelectorAll('a[id*=policySearchResultsGrid][id$=btnPolicy]'));" & _
        "if(!ls.length){var e=(document.getElementById('ctl00_MainContent_lblError')||{}).innerText||'';" & _
        "var t=document.body?document.body.innerText:'';" & _
        "return (/no (records|results|policies)|not found|0 records/i.test(e+' '+t)&&document.getElementById('ctl00_MainContent_txtSearchStr'))?'none':'';}" & _
        "var rows=ls.map(function(a){var tr=a.closest('tr');var td=tr.querySelectorAll('td');" & _
        "return {a:a,no:a.innerText.replace(/\D/g,''),st:(td[td.length-1]||{}).innerText||''};});" & _
        "var want=arguments[0],pick=null;" & _
        "if(want)rows.forEach(function(r){if(r.no.slice(-2)===want)pick=r;});" & _
        "if(!pick)rows.forEach(function(r){if(!pick&&/active/i.test(r.st))pick=r;});" & _
        "if(!pick)rows.forEach(function(r){if(!pick||r.no>pick.no)pick=r;});" & _
        "pick.a.click();return 'clicked';"
End Function

' "ok|next amount|due on|status|status detail|cancel date|header number".
Private Function CrownNgBillingJs() As String
    CrownNgBillingJs = _
        "var g=function(id){var e=document.getElementById(id);return e?e.innerText.replace(/[|]/g,' ').trim():'';};" & _
        "if(!/PolicyBillPlans/i.test(location.pathname))return '';" & _
        "if(!document.getElementById('ctl00_MainContent_InsuredInfo1_lblStatus'))return '';" & _
        "var t=(document.body.innerText||'').replace(/\s+/g,' ');" & _
        "var m=t.match(/cancel\w*\s*(?:effective\s*)?(?:date)?\s*:?\s*(\d{1,2}\/\d{1,2}\/\d{2,4})/i);" & _
        "return 'ok|'+g('ctl00_MainContent_PaymentSummary2_lblNextAmountDue')+'|'+g('ctl00_MainContent_PaymentSummary2_lblAmountDueDate')" & _
        "+'|'+g('ctl00_MainContent_InsuredInfo1_lblStatus')+'|'+g('ctl00_MainContent_InsuredInfo1_lblStatusDetail')" & _
        "+'|'+(m?m[1]:'')+'|'+g('ctl00_lblHeaderPageTitleTop');"
End Function

' Same rule as the United and Verve runs: the website says cancelled AND
' the cancel date is more than 60 days ago.
Private Function CrownNgLongCancelled(ByVal parts As Variant, ByRef since As String) As Boolean
    Dim words As String, cancelledOn As Double

    since = ""
    If UBound(parts) < 10 Then Exit Function

    words = CStr(parts(6)) & " " & CStr(parts(7))
    If InStr(1, words, "cancel", vbTextCompare) = 0 And InStr(1, words, "laps", vbTextCompare) = 0 _
       And InStr(1, words, "cx", vbTextCompare) = 0 Then Exit Function

    cancelledOn = CrownUsDate(CStr(parts(10)))
    If cancelledOn = 0 Or cancelledOn > CDbl(Date) - 60 Then Exit Function

    since = Trim$(CStr(parts(10)))
    CrownNgLongCancelled = True
End Function

' Saves each page Tony opens (picture + page code) into one folder in
' Downloads, until he presses Cancel, then opens the folder.
Private Sub CrownLookPages(ByVal drv As ChromeDriver, ByVal carrier As String)
    Dim folder As String, k As Long, frames As String

    folder = Environ$("USERPROFILE") & "\Downloads\CrownLook-" & Replace(carrier, " ", "") & "-" & Format$(Now, "yyyymmdd-hhnn")
    On Error Resume Next
    MkDir folder
    On Error GoTo 0

    Do
        If MsgBox("Signed in to " & carrier & "." & vbCrLf & vbCrLf & _
                  "In the Chrome window, open each page I asked for, then press OK here to save it" & vbCrLf & _
                  "(one customer's policy pages: drivers, vehicles, coverages / deductibles," & vbCrLf & _
                  "lienholder, and the page with the amount due and due date)." & vbCrLf & vbCrLf & _
                  "Pages saved so far: " & k & ". Press Cancel when you are finished.", _
                  vbOKCancel + vbInformation, "Crown Superior - " & carrier) = vbCancel Then Exit Do

        k = k + 1

        On Error Resume Next
        drv.SwitchToDefaultContent
        drv.TakeScreenshot.SaveAs folder & "\page" & k & ".png"
        CrownLookSave folder & "\page" & k & ".html", "<!-- " & drv.url & " -->" & vbCrLf & drv.PageSource
        frames = CrownMoreJs(drv, _
            "return [].map.call(document.querySelectorAll('iframe'),function(f,i){try{return '<!-- frame '+i+' '+f.src+' -->'+" & _
            "f.contentDocument.documentElement.outerHTML;}catch(e){return '<!-- frame '+i+' '+f.src+' (other site) -->';}}).join('\n');")
        If Len(frames) > 0 Then CrownLookSave folder & "\page" & k & "-frames.html", frames
        On Error GoTo 0
    Loop While k < 30

    If k > 0 Then
        Shell "explorer.exe """ & folder & """", vbNormalFocus
        MsgBox k & " page(s) saved in:" & vbCrLf & folder & vbCrLf & vbCrLf & _
               "Please zip that folder (right-click > Send to > Compressed folder) and send it to me in our chat.", _
               vbInformation, "Crown Superior"
    End If
End Sub

Private Sub CrownLookSave(ByVal path As String, ByVal text As String)
    With CreateObject("ADODB.Stream")
        .Type = 2
        .Charset = "utf-8"
        .Open
        .WriteText text
        .SaveToFile path, 2
        .Close
    End With
End Sub
