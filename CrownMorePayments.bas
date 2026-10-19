Attribute VB_Name = "CrownMorePayments"
Option Explicit

' =====================================================================
' Crown Superior - payment due for Assurance America and Trexis
' ---------------------------------------------------------------------
' Alt+F8 > CrownAssurancePaymentDue > Run
' Alt+F8 > CrownTrexisPaymentDue > Run    (further down)
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

        recordId = CrownMoreMatch(byNumber, byName, AA_CARRIER, policyNo, insured, matchedOn, note)

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
    Dim cPol As Long, cName As Long, cStatus As Long, cDate As Long, cAmount As Long
    Dim i As Long, row As Long
    Dim policyNo As String, insured As String, status As String, statusDate As String
    Dim dueAmount As String, dueDate As String, cancelDate As String
    Dim recordId As String, matchedOn As String, note As String
    Dim done As Long, unmatched As Long, failed As Long
    Dim started As Date, exportAt As Date

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
    cDate = CrownMoreColumn(head, "Status Date")
    cAmount = CrownMoreColumn(head, "Amount Due")

    If cPol < 0 Or cStatus < 0 Or cAmount < 0 Then
        MsgBox "Trexis's export file did not have the columns expected (Policy Number, Status, Amount Due)." & _
               vbCrLf & vbCrLf & "It is " & csvPath & " - please send it to me.", vbExclamation, "Crown Superior"
        GoTo Finish
    End If

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
        If cDate >= 0 And UBound(f) >= cDate Then statusDate = CrownMoreDate(f(cDate))
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

        recordId = CrownMoreMatch(byNumber, byName, TX_CARRIER, policyNo, insured, matchedOn, note)

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
                                ByRef matchedOn As String, ByRef note As String) As String
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

            If CrownSameCarrier(carrier, CStr(parts(3))) Then
                If Len(found) > 0 Then
                    note = "that name has more than one " & carrier & " policy on the website - not guessed"
                    Exit Function
                End If

                found = CStr(parts(0))
            End If
        Next j

        If Len(found) > 0 Then
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
