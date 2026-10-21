Attribute VB_Name = "CrownPolicyLook"
Option Explicit

' =====================================================================
' Crown Superior - write down a carrier's policy pages (step 1 of the
' "pull the real policy into Edit Data" job)
' ---------------------------------------------------------------------
' Alt+F8 > CrownRecordPolicyPages > Run, type a policy number.
'
' Signs in to the carrier the same way the payment runs do, opens that
' one policy, and writes everything on its pages onto the PolicyLook
' sheet: every table row, every box and what is in it, every tab name.
' Nothing is typed into the carrier's site except the sign-in and the
' policy number, and nothing on the policy is changed.
'
' The sheet is what tells me exactly where the drivers, vehicles,
' deductibles and lienholders sit on each carrier's pages, so the real
' pull can read them by name instead of by guesswork.
'
'   GAF...      -> Verve  (Policy Info, Drivers/Vehicles/Coverages tabs)
'   anything else -> United (the Review Policy page)
' =====================================================================

Private Const MOST_LINES As Long = 1500

Public Sub CrownRecordPolicyPages()
    Dim policyNo As String
    Dim ws As Worksheet, wsInput As Worksheet
    Dim clsDrv As Object
    Dim drv As ChromeDriver
    Dim Keys As New Selenium.Keys
    Dim By As New Selenium.By
    Dim nextRow As Long, n As Long
    Dim tabs As Variant, t As Variant
    Dim tries As Long, opened As Boolean, answer As String

    policyNo = Trim$(InputBox("Policy number to write down (United or Verve):", "Crown Superior"))
    If Len(policyNo) = 0 Then Exit Sub

    Set wsInput = ThisWorkbook.Worksheets("Input")
    Set ws = CrownLookSheet()
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If nextRow < 2 Then nextRow = 2

    Set clsDrv = New Chrm
    Set clsDrv.ChrmDriver = New ChromeDriver
    Set drv = clsDrv.ChrmDriver

    On Error Resume Next
    drv.AddArgument "--force-device-scale-factor=0.80"
    drv.Start
    On Error GoTo 0

    ' Chrome did not start (usually: Chrome updated, chromedriver is old) - say so in words.
    If Not CrownChromeOk(drv) Then Exit Sub

    If UCase$(Left$(Replace(policyNo, " ", ""), 3)) = "GAF" Then
        ' ----------------------------------------------------- Verve
        drv.Get wsInput.Range("URL_9").value

        If Not CrownLookWait(drv, By, "LoginControl_LoginNameTextBox", 40) Then
            MsgBox "Verve did not show its sign-in page.", vbExclamation, "Crown Superior"
            Exit Sub
        End If

        EnterData drv, Keys, "LoginControl_LoginNameTextBox", wsInput.Range("USER_2").value, "ID"
        EnterData drv, Keys, "LoginControl_PasswordTextBox", wsInput.Range("PASS_2").value, "ID"
        ClickElement drv, "LoginControl_LoginLinkButton", "ID"

        If Not CrownLookWait(drv, By, "QuickPolicyLookupControl_NumberInsCombo_I", 60) Then
            MsgBox "Signed in to Verve, but the policy lookup box never came up.", vbExclamation, "Crown Superior"
            Exit Sub
        End If

        ' The first try on 17 Oct came back with the home page still on
        ' screen: the lookup went in while the home page was still loading
        ' its own lists. So: let it finish, then look the policy up the way
        ' the payment run does - wait for the number to be ON the page,
        ' open it from a list if Verve shows one, and try again if not.
        drv.Wait 3000

        For tries = 1 To 3
            CrownVerveKeepAlive drv

            EnterData drv, Keys, "QuickPolicyLookupControl_NumberInsCombo_I", Replace(policyNo, " ", ""), "ID", "", True
            ClickElement drv, "QuickPolicyLookupControl_FireLookupImageButton", "ID"

            If CrownPageShows(drv, policyNo, 8) Then
                opened = True
            Else
                answer = CrownVerveLookupAnswer(drv, policyNo)

                If answer = "notfound" Then Exit For

                If answer = "clicked" Then opened = CrownPageShows(drv, policyNo, 20)
            End If

            If opened Then Exit For

            drv.Wait 3000
        Next tries

        CrownLookSettle drv, 3000

        If Not opened Then
            n = CrownLookWrite(ws, nextRow, "Verve", policyNo, "DID NOT OPEN (" & _
                               IIf(answer = "notfound", "Verve said Number not found", "lookup tried 3 times") & _
                               ") - this is what was on screen", drv)

            ThisWorkbook.Save
            ws.Activate

            MsgBox "Verve did not open " & policyNo & _
                   IIf(answer = "notfound", " - it says Number not found.", " after 3 tries.") & vbCrLf & vbCrLf & _
                   "Please send me the PolicyLook sheet anyway - it shows what Verve had on screen.", _
                   vbExclamation, "Crown Superior"

            On Error Resume Next
            drv.Quit

            Exit Sub
        End If

        n = CrownLookWrite(ws, nextRow, "Verve", policyNo, "policy page", drv)

        tabs = Array("Policy Info", "Drivers/Vehicles/Coverages", "Drivers", "Vehicles", "Coverages", "Billing", "Future")

        For Each t In tabs
            If CrownLookClickText(drv, CStr(t)) Then
                CrownLookSettle drv, 5000
                n = n + CrownLookWrite(ws, nextRow, "Verve", policyNo, "tab: " & CStr(t), drv)
            End If
        Next t
    Else
        ' ---------------------------------------------------- United
        drv.Get wsInput.Range("URL_26").value
        EnterData drv, Keys, "tbxUserID", wsInput.Range("USER_26").value, "ID"
        EnterData drv, Keys, "tbxPassword", wsInput.Range("PASS_26").value, "ID"
        drv.Keyboard.SendKeys Keys.Enter

        If Not CrownLookWait(drv, By, "rpthref", 60) Then
            MsgBox "United did not finish signing in.", vbExclamation, "Crown Superior"
            Exit Sub
        End If

        ' Same way in as the payment run: Policy Inquiry first, then the
        ' policy's own page.
        ClickElement drv, "//a[normalize-space(text())='Work with Policies']", "XPATH"
        ClickElement drv, "//a[normalize-space(text())='Policy Inquiry']", "XPATH"
        CrownLookSettle drv, 4000

        drv.Get CrownLookUnitedUrl(drv, policyNo)
        CrownLookSettle drv, 6000

        n = CrownLookWrite(ws, nextRow, "United", policyNo, "Review Policy page", drv)

        ' Anything on that page that opens more detail - written down too.
        tabs = Array("Drivers", "Vehicles", "Coverages", "Vehicle", "Driver")

        For Each t In tabs
            If CrownLookClickText(drv, CStr(t)) Then
                CrownLookSettle drv, 5000
                n = n + CrownLookWrite(ws, nextRow, "United", policyNo, "after clicking: " & CStr(t), drv)
                drv.Get CrownLookUnitedUrl(drv, policyNo)
                CrownLookSettle drv, 5000
            End If
        Next t
    End If

    ThisWorkbook.Save

    ws.Activate

    MsgBox n & " lines written on the PolicyLook sheet for " & policyNo & "." & vbCrLf & vbCrLf & _
           "Please send me that sheet (right-click the PolicyLook tab > Move or Copy > " & _
           "new book > Create a copy > save it and attach it).", vbInformation, "Crown Superior"

    On Error Resume Next
    drv.Quit
End Sub


Private Function CrownLookSheet() As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("PolicyLook")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "PolicyLook"
        ws.Range("A1:D1").value = Array("Carrier", "Policy", "Page", "What is on it")
        ws.Columns(4).ColumnWidth = 120
    End If

    Set CrownLookSheet = ws
End Function


' Everything on the page, one line per table row / box / tab.
Private Function CrownLookWrite(ByVal ws As Worksheet, ByRef nextRow As Long, ByVal carrier As String, _
                                ByVal policyNo As String, ByVal where As String, ByVal drv As Object) As Long
    Dim text As String
    Dim lines() As String
    Dim i As Long

    On Error Resume Next
    text = CStr(drv.ExecuteScript( _
        "function c(s){return ((s||'')+'').replace(/\u00a0/g,' ').replace(/\s+/g,' ').trim();}" & _
        "function vis(e){return !!(e.getClientRects&&e.getClientRects().length);}" & _
        "var out=['URL: '+location.href,'TITLE: '+document.title];" & _
        "var rows=document.querySelectorAll('tr');" & _
        "for(var i=0;i<rows.length&&out.length<" & MOST_LINES & ";i++){var r=rows[i];" & _
        "if(r.querySelector('table')||!vis(r))continue;var cs=[];" & _
        "for(var j=0;j<r.cells.length;j++){var x=c(r.cells[j].innerText);" & _
        "var f=r.cells[j].querySelector('input:not([type=hidden]),select,textarea');" & _
        "if(f){var v=f.tagName==='SELECT'?(f.options[f.selectedIndex]?f.options[f.selectedIndex].text:''):" & _
        "((f.type==='checkbox'||f.type==='radio')?(f.checked?'[x]':'[ ]'):f.value);" & _
        "x=(x?x+' ':'')+'{'+(f.id||f.name||'box')+'='+c(v)+'}';}" & _
        "if(x)cs.push(x);}" & _
        "if(cs.length)out.push('ROW: '+cs.join(' | ').substring(0,900));}" & _
        "var fs=document.querySelectorAll('input:not([type=hidden]),select,textarea');" & _
        "for(var k=0;k<fs.length&&out.length<" & MOST_LINES & ";k++){var e=fs[k];if(!vis(e))continue;" & _
        "var v2=e.tagName==='SELECT'?(e.options[e.selectedIndex]?e.options[e.selectedIndex].text:''):" & _
        "((e.type==='checkbox'||e.type==='radio')?(e.checked?'[x]':'[ ]'):e.value);" & _
        "out.push('BOX: '+(e.id||e.name||'?')+' = '+c(v2).substring(0,200));}" & _
        "var ls=document.querySelectorAll('a,li,span[class*=tab]');var seen={};" & _
        "for(var m=0;m<ls.length&&out.length<" & MOST_LINES & ";m++){var w=c(ls[m].innerText);" & _
        "if(!w||w.length>40||seen[w]||!vis(ls[m]))continue;seen[w]=1;out.push('LINK: '+w);}" & _
        "return out.join(String.fromCharCode(10));"))
    If Err.Number <> 0 Then text = "could not read the page: " & Err.Description
    On Error GoTo 0

    lines = Split(text, vbLf)

    For i = LBound(lines) To UBound(lines)
        ws.Cells(nextRow, 1).value = carrier
        ws.Cells(nextRow, 2).value = "'" & policyNo
        ws.Cells(nextRow, 3).value = where
        ws.Cells(nextRow, 4).value = "'" & Left$(lines(i), 1000)
        nextRow = nextRow + 1
    Next i

    CrownLookWrite = UBound(lines) - LBound(lines) + 1
End Function


' Click the first visible link/tab whose own words are exactly this.
Private Function CrownLookClickText(ByVal drv As Object, ByVal words As String) As Boolean
    Dim answer As Variant

    On Error Resume Next
    answer = drv.ExecuteScript( _
        "function c(s){return ((s||'')+'').replace(/\u00a0/g,' ').replace(/\s+/g,' ').trim().toLowerCase();}" & _
        "var w=c(arguments[0]),els=document.querySelectorAll('a,li,span,td,div,button');" & _
        "for(var i=0;i<els.length;i++){var e=els[i];if(!e.getClientRects().length)continue;" & _
        "if(c(e.innerText)!==w)continue;var a=e.tagName==='A'?e:(e.querySelector('a')||e);" & _
        "a.click();return 'clicked';}return '';", words)
    If Err.Number <> 0 Then answer = ""
    On Error GoTo 0

    CrownLookClickText = (CStr(answer) = "clicked")
End Function


' Wait for the page to finish (and any loading cover to go), up to ms.
Private Sub CrownLookSettle(ByVal drv As Object, ByVal ms As Long)
    Dim started As Date
    Dim answer As Variant

    started = Now
    drv.Wait 800

    Do
        On Error Resume Next
        answer = drv.ExecuteScript("return document.readyState")
        On Error GoTo 0

        If CStr(answer) = "complete" Then Exit Do

        drv.Wait 400
        DoEvents
    Loop While (Now - started) * 86400000# < ms

    drv.Wait 1200
End Sub


Private Function CrownLookWait(ByVal drv As Object, ByVal By As Selenium.By, ByVal boxId As String, _
                               ByVal seconds As Long) As Boolean
    Dim started As Date

    started = Now

    Do
        If drv.IsElementPresent(By.ID(boxId)) Then
            CrownLookWait = True

            Exit Function
        End If

        drv.Wait 500
        DoEvents
    Loop While (Now - started) * 86400 < seconds
End Function


' United's Review Policy page for one number, on the host the browser is
' already signed in to (the agent site is per state).
Private Function CrownLookUnitedUrl(ByVal drv As Object, ByVal policyNo As String) As String
    Dim here As String, host As String
    Dim letters As String, digits As String
    Dim i As Long, at As Long
    Dim ch As String

    For i = 1 To Len(policyNo)
        ch = UCase$(Mid$(policyNo, i, 1))

        If ch >= "0" And ch <= "9" Then
            digits = digits & ch
        ElseIf ch >= "A" And ch <= "Z" And Len(digits) = 0 Then
            letters = letters & ch
        End If
    Next i

    On Error Resume Next
    here = CStr(drv.Url)
    On Error GoTo 0

    at = InStr(9, here, "/")
    If at = 0 Then Exit Function

    host = Left$(here, at - 1)

    CrownLookUnitedUrl = host & "/agents/ndmacro/cmn_Review.mac/ValidatePolicy" & _
                         "?dbxPolcyPfx=" & letters & "&tbxPolicyNo=" & digits & _
                         "&TRANSFLG=PE&strPolcntFlag=Y&strMtermFlag=Y"
End Function


' =====================================================================
' Pull a United policy into Edit Data (Tony, 21 Oct: "pull GAI 206097
' from United ... all updated info to the edit data page")
' ---------------------------------------------------------------------
' Open the customer on Edit Data first, then
' Alt+F8 > CrownPullUnitedPolicy > Run, and type (or keep) the number.
'
' Reads United's Review Policy page - the one CrownRecordPolicyPages
' wrote down on 17 Oct: the drivers (principal and the others, excluded
' or not), the vehicles, each vehicle's collision/comprehensive
' deductibles and the lienholders - and lists every place Edit Data says
' something different on the UnitedPull sheet. Nothing changes until Yes
' is pressed; then the new values go onto Edit Data and the customer is
' saved with his own Update (UpdateDataToOutput).
'
' It only fills and corrects. A driver or vehicle that is on Edit Data
' but not on United is left where it is and listed, never deleted.
' =====================================================================
Public Sub CrownPullUnitedPolicy()
    Dim wsEdit As Worksheet, wsInput As Worksheet, ws As Worksheet
    Dim clsDrv As Object, drv As ChromeDriver
    Dim Keys As New Selenium.Keys
    Dim By As New Selenium.By
    Dim policyNo As String, answer As String, lines() As String, f() As String
    Dim changes As Collection, notes As Collection, ch As Variant
    Dim drivers As Collection, vehicles As Collection, covers As Collection, liens As Collection
    Dim insured As Variant, k As Long, row As Long, shown As String, n As Long
    Dim dCols As Variant, vCols As Variant, col As String, item As Variant

    Set wsEdit = ThisWorkbook.Worksheets("Edit Data")
    Set wsInput = ThisWorkbook.Worksheets("Input")

    If Val(wsEdit.Range("C5").value) < 2 Then
        MsgBox "Open the customer on Edit Data first (their row number in C5), then run this again.", _
               vbExclamation, "Crown Superior"
        Exit Sub
    End If

    policyNo = Trim$(InputBox("United policy number to pull into Edit Data" & vbCrLf & _
                              "(customer on Edit Data now: " & Trim$(wsEdit.Range("D18").Text & " " & wsEdit.Range("D19").Text) & ")", _
                              "Crown Superior", Trim$(wsEdit.Range("C7").Text)))
    If Len(policyNo) = 0 Then Exit Sub

    Set clsDrv = New Chrm
    Set clsDrv.ChrmDriver = New ChromeDriver
    Set drv = clsDrv.ChrmDriver

    On Error Resume Next
    drv.AddArgument "--force-device-scale-factor=0.80"
    drv.Start
    On Error GoTo 0

    If Not CrownChromeOk(drv) Then Exit Sub

    ' Same way in as the payment run and the recorder.
    drv.Get wsInput.Range("URL_26").value
    EnterData drv, Keys, "tbxUserID", wsInput.Range("USER_26").value, "ID"
    EnterData drv, Keys, "tbxPassword", wsInput.Range("PASS_26").value, "ID"
    drv.Keyboard.SendKeys Keys.Enter

    If Not CrownLookWait(drv, By, "rpthref", 60) Then
        MsgBox "United did not finish signing in. Nothing has been changed.", vbExclamation, "Crown Superior"
        GoTo Finish
    End If

    ClickElement drv, "//a[normalize-space(text())='Work with Policies']", "XPATH"
    ClickElement drv, "//a[normalize-space(text())='Policy Inquiry']", "XPATH"
    CrownLookSettle drv, 4000

    drv.Get CrownLookUnitedUrl(drv, policyNo)
    CrownLookSettle drv, 6000

    On Error Resume Next
    answer = CStr(drv.ExecuteScript(CrownPullUnitedJs()))
    On Error GoTo 0

    On Error Resume Next
    drv.Quit
    On Error GoTo 0

    If InStr(answer, "DRV|") = 0 And InStr(answer, "VEH|") = 0 Then
        MsgBox "United's page for " & policyNo & " did not show its drivers or vehicles, so nothing has been changed." & vbCrLf & _
               "(Check the number; if it is right, run CrownRecordPolicyPages for it and send me the PolicyLook sheet.)", _
               vbExclamation, "Crown Superior"
        Exit Sub
    End If

    ' --- what United says -------------------------------------------------
    Set drivers = New Collection: Set vehicles = New Collection
    Set covers = New Collection: Set liens = New Collection
    insured = Array("", "", "", "", "", "", "")

    lines = Split(answer, vbLf)
    For k = 0 To UBound(lines)
        f = Split(lines(k) & "||||||||", "|")
        Select Case f(0)
            Case "INS": insured = f
            Case "DRV"
                ' only drivers United still has on the policy
                If Len(f(8)) = 0 Or InStr(1, f(8), "active", vbTextCompare) > 0 Then drivers.Add f
            Case "VEH"
                If Len(f(5)) = 0 Or InStr(1, f(5), "active", vbTextCompare) > 0 Then vehicles.Add f
            Case "COV": covers.Add f
            Case "LIEN": liens.Add f
        End Select
    Next k

    ' --- compare with Edit Data ----------------------------------------------
    Set changes = New Collection
    Set notes = New Collection
    dCols = Array("D", "F", "H", "J")
    vCols = Array("P", "R", "T", "V")

    CrownPullAdd changes, wsEdit, "C7", "Policy No", Trim$(CStr(insured(6)))
    CrownPullAdd changes, wsEdit, "H8", "Phone", CrownPullPhone(CStr(insured(3)))
    CrownPullAdd changes, wsEdit, "C8", "Email", LCase$(Trim$(CStr(insured(4))))
    If Len(insured(2)) > 0 Then notes.Add "United's mailing address: " & insured(2) & " (address not changed - check it by eye)"

    For k = 1 To drivers.Count
        If k > 4 Then
            notes.Add "United has more than 4 drivers - driver " & k & " (" & drivers(k)(1) & ") has no column on Edit Data"
            Exit For
        End If

        item = drivers(k)
        col = dCols(k - 1)

        If k = 1 And Len(item(1)) = 0 Then item(1) = insured(1)

        If k > 1 Then
            CrownPullAdd changes, wsEdit, col & "16", "Driver " & k & " - is there a driver", "Yes"
            CrownPullAdd changes, wsEdit, col & "17", "Driver " & k & " - excluded", IIf(InStr(1, item(7), "exclud", vbTextCompare) > 0, "Yes", "No")
        End If

        If Len(item(1)) > 0 Then
            CrownPullAdd changes, wsEdit, col & "18", "Driver " & k & " - first name", CrownPullFirst(CStr(item(1)))
            CrownPullAdd changes, wsEdit, col & "19", "Driver " & k & " - last name", CrownPullLast(CStr(item(1)))
        Else
            notes.Add "United did not show driver " & k & "'s name (DOB " & item(2) & ") - name left as it is"
        End If

        CrownPullAdd changes, wsEdit, col & "20", "Driver " & k & " - gender", CrownPullPick(CStr(item(3)), "Male,Female")
        CrownPullAdd changes, wsEdit, col & "21", "Driver " & k & " - marital status", IIf(InStr(1, item(4), "married", vbTextCompare) = 1, "Married", IIf(Len(item(4)) > 0, "Single", ""))
        CrownPullAdd changes, wsEdit, col & "22", "Driver " & k & " - date of birth", CStr(item(2)), True
        CrownPullAdd changes, wsEdit, col & "24", "Driver " & k & " - license no", CStr(item(5))
        CrownPullAdd changes, wsEdit, col & "27", "Driver " & k & " - license state", UCase$(CStr(item(6)))
    Next k

    For k = drivers.Count + 1 To 4
        If k >= 2 Then
            If UCase$(Trim$(wsEdit.Range(dCols(k - 1) & "16").Text)) = "YES" Then
                notes.Add "Edit Data has driver " & k & " (" & Trim$(wsEdit.Range(dCols(k - 1) & "18").Text & " " & wsEdit.Range(dCols(k - 1) & "19").Text) & _
                          ") that United does not - left as it is"
            End If
        End If
    Next k

    For k = 1 To vehicles.Count
        If k > 4 Then
            notes.Add "United has more than 4 vehicles - vehicle " & k & " has no column on Edit Data"
            Exit For
        End If

        item = vehicles(k)
        col = vCols(k - 1)

        If k > 1 Then CrownPullAdd changes, wsEdit, col & "16", "Vehicle " & k & " - is there a vehicle", "Yes"
        CrownPullAdd changes, wsEdit, col & "17", "Vehicle " & k & " - VIN", UCase$(CStr(item(1)))
        CrownPullAdd changes, wsEdit, col & "18", "Vehicle " & k & " - make", CStr(item(3))
        CrownPullAdd changes, wsEdit, col & "19", "Vehicle " & k & " - year", CStr(item(2))
        CrownPullAdd changes, wsEdit, col & "20", "Vehicle " & k & " - model", CStr(item(4))

        If k <= covers.Count Then
            answer = CrownPullDeductible(CStr(covers(k)(2)), CStr(covers(k)(3)))
            If Left$(answer, 1) = "?" Then
                notes.Add "Vehicle " & k & ": United has collision " & covers(k)(2) & " / comprehensive " & covers(k)(3) & _
                          " - not one of Edit Data's choices, deductible left as it is"
            Else
                CrownPullAdd changes, wsEdit, col & "23", "Vehicle " & k & " - deductible", answer
            End If
        End If

        If k <= liens.Count And k <= 2 Then
            CrownPullAdd changes, wsEdit, col & "30", "Vehicle " & k & " - lien holder", "Yes"
            CrownPullAdd changes, wsEdit, col & "31", "Vehicle " & k & " - lien holder detail", Trim$(liens(k)(1) & ", " & liens(k)(2))
        ElseIf k <= liens.Count Then
            CrownPullAdd changes, wsEdit, col & "30", "Vehicle " & k & " - lien holder", "Yes"
            notes.Add "Vehicle " & k & " lienholder: " & liens(k)(1) & ", " & liens(k)(2) & " (Edit Data has no detail box for vehicle " & k & ")"
        End If
    Next k

    If liens.Count > 0 And vehicles.Count > 1 Then
        notes.Add "United lists lienholders without saying which vehicle each is for - they were matched in order, check them"
    End If

    ' --- the report ---------------------------------------------------------
    Set ws = CrownPullSheet()
    ws.Range("A1:E1").value = Array("Edit Data cell", "What", "Edit Data has", "United says", "Done")
    row = 2

    For Each ch In changes
        ws.Cells(row, 1).value = ch(0)
        ws.Cells(row, 2).value = ch(1)
        ws.Cells(row, 3).value = "'" & ch(2)
        ws.Cells(row, 4).value = "'" & ch(3)
        row = row + 1
        n = n + 1
        If n <= 14 Then shown = shown & ch(1) & ": " & IIf(Len(ch(2)) = 0, "(empty)", ch(2)) & "  ->  " & ch(3) & vbCrLf
    Next ch

    If notes.Count > 0 Then
        row = row + 1
        ws.Cells(row, 1).value = "Notes"
        For Each ch In notes
            row = row + 1
            ws.Cells(row, 2).value = ch
        Next ch
    End If

    ws.Columns("A:E").AutoFit
    If ws.Columns("B").ColumnWidth > 60 Then ws.Columns("B").ColumnWidth = 60

    If changes.Count = 0 Then
        MsgBox "Edit Data already matches United for " & policyNo & "." & IIf(notes.Count > 0, vbCrLf & vbCrLf & notes.Count & " note(s) on the UnitedPull sheet.", ""), _
               vbInformation, "Crown Superior"
        Exit Sub
    End If

    If n > 14 Then shown = shown & "... and " & (n - 14) & " more (all on the UnitedPull sheet)" & vbCrLf

    If MsgBox(changes.Count & " difference(s) between Edit Data and United for " & policyNo & ":" & vbCrLf & vbCrLf & shown & vbCrLf & _
              IIf(notes.Count > 0, notes.Count & " note(s) on the UnitedPull sheet." & vbCrLf & vbCrLf, "") & _
              "Put United's values on Edit Data and save this customer now?", vbYesNo + vbQuestion, "Crown Superior") = vbNo Then
        ws.Activate
        Exit Sub
    End If

    ' --- apply, then save with his own Update ------------------------------
    Application.EnableEvents = False
    row = 2
    For Each ch In changes
        If ch(4) Then
            wsEdit.Range(ch(0)).value = CDate(ch(3))
        Else
            wsEdit.Range(ch(0)).value = ch(3)
        End If
        ws.Cells(row, 5).value = "yes"
        row = row + 1
    Next ch
    Application.EnableEvents = True

    wsEdit.Activate
    UpdateDataToOutput
    Exit Sub

Finish:
    On Error Resume Next
    drv.Quit
End Sub

' One difference: Array(cell, what, Edit Data has, United says, is a date).
Private Sub CrownPullAdd(ByVal changes As Collection, ByVal ws As Worksheet, ByVal addr As String, _
                         ByVal what As String, ByVal united As String, Optional ByVal asDate As Boolean = False)
    Dim have As String

    united = Trim$(united)
    If Len(united) = 0 Then Exit Sub

    have = Trim$(ws.Range(addr).Text)

    If asDate Then
        If Not IsDate(united) Then Exit Sub
        If IsDate(ws.Range(addr).value) Then
            If CDate(ws.Range(addr).value) = CDate(united) Then Exit Sub
        End If
        united = Format$(CDate(united), "mm/dd/yyyy")
    Else
        If StrComp(Application.WorksheetFunction.Trim(have), Application.WorksheetFunction.Trim(united), vbTextCompare) = 0 Then Exit Sub
        ' "GAI-206097" and "GAI 206097" are the same number.
        If addr = "C7" And CrownDigits(have) = CrownDigits(united) And Len(CrownDigits(have)) > 0 Then Exit Sub
        ' Same phone in another layout.
        If addr = "H8" And CrownDigits(have) = CrownDigits(united) Then Exit Sub
    End If

    changes.Add Array(addr, what, have, united, asDate)
End Sub

Private Function CrownPullFirst(ByVal fullName As String) As String
    Dim w() As String
    w = Split(Application.WorksheetFunction.Trim(fullName), " ")
    If UBound(w) >= 0 Then CrownPullFirst = w(0)
End Function

Private Function CrownPullLast(ByVal fullName As String) As String
    Dim w() As String
    w = Split(Application.WorksheetFunction.Trim(fullName), " ")
    If UBound(w) >= 1 Then CrownPullLast = w(UBound(w))
End Function

' The choice from the list that the words start with, or "".
Private Function CrownPullPick(ByVal text As String, ByVal choices As String) As String
    Dim c As Variant
    For Each c In Split(choices, ",")
        If Len(text) > 0 And StrComp(Left$(c, Len(Trim$(text))), Trim$(text), vbTextCompare) = 0 Then
            CrownPullPick = c
            Exit Function
        End If
    Next c
End Function

Private Function CrownPullPhone(ByVal text As String) As String
    Dim d As String
    d = CrownDigits(text)
    If Len(d) = 11 And Left$(d, 1) = "1" Then d = Mid$(d, 2)
    If Len(d) = 10 Then CrownPullPhone = "(" & Left$(d, 3) & ") " & Mid$(d, 4, 3) & "-" & Right$(d, 4)
End Function

' Collision + comprehensive -> one of Edit Data's deductible choices
' ("1000/1000"), "None", or "?" when Edit Data has no such choice.
Private Function CrownPullDeductible(ByVal coll As String, ByVal comp As String) As String
    coll = CStr(Val(CrownDigits(coll)))
    comp = CStr(Val(CrownDigits(comp)))

    If coll = "0" And comp = "0" Then
        CrownPullDeductible = "None"
    ElseIf coll = comp And InStr(",250,500,1000,2000,2500,", "," & coll & ",") > 0 Then
        CrownPullDeductible = coll & "/" & comp
    Else
        CrownPullDeductible = "?"
    End If
End Function

Private Function CrownPullSheet() As Worksheet
    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, "UnitedPull", vbTextCompare) = 0 Then
            ws.Cells.ClearContents
            Set CrownPullSheet = ws
            Exit Function
        End If
    Next ws

    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = "UnitedPull"
    Set CrownPullSheet = ws
End Function

' Lines, "|" between fields:
'   INS|full name|mailing address|home phone|email|policy status|policy number
'   DRV|name|DOB|gender|marital|license or ID|state|driver type|status
'   VEH|VIN|year|make|model|status
'   COV|n|collision|comprehensive      (one per vehicle, in order)
'   LIEN|name|address
Private Function CrownPullUnitedJs() As String
    CrownPullUnitedJs = _
        "var clean=function(s){return (s||'').replace(/[|\r\n]/g,' ').replace(/\s+/g,' ').trim();};" & _
        "var rows=[].slice.call(document.querySelectorAll('tr')).filter(function(t){return !t.querySelector('tr');});" & _
        "var cells=function(t){return t?[].slice.call(t.children).filter(function(c){return c.tagName==='TD'||c.tagName==='TH';}).map(function(c){return clean(c.innerText);}):[];};" & _
        "var lv=function(l){for(var i=0;i<rows.length;i++){var c=cells(rows[i]);for(var j=0;j<c.length-1;j++){if(c[j].toLowerCase()===l)return c[j+1];}}return '';};" & _
        "var out=['INS|'+[lv('insured full name'),lv('mailing address'),lv('home phone'),lv('email'),lv('policy status'),lv('policy number')].join('|')];" & _
        "var nameNear=function(t){var re=/driver\s*#?\s*\d*\s*name\s*:?\s*([A-Za-z][A-Za-z .,'\-]*)/i,m;" & _
        "var tb=t.closest('table'),k,e;for(k=rows.indexOf(t)-1;k>=0&&k>=rows.indexOf(t)-3;k--){m=cells(rows[k]).join(' ').match(re);if(m)return clean(m[1]);if(/DOB/.test(cells(rows[k]).join(' ')))break;}" & _
        "if(tb){if(tb.caption&&(m=tb.caption.innerText.match(re)))return clean(m[1]);e=tb.previousElementSibling;" & _
        "for(k=0;k<3&&e;k++,e=e.previousElementSibling){m=(e.innerText||'').match(re);if(m)return clean(m[1]);}}return '';};" & _
        "var cov=0,inCov=false,coll='',comp='',inLien=false;"
    CrownPullUnitedJs = CrownPullUnitedJs & _
        "for(var i=0;i<rows.length;i++){var c=cells(rows[i]),line=c.join(' ');" & _
        "if(c.indexOf('DOB')>=0&&c.indexOf('Gender')>=0){var d=cells(rows[i+1]);" & _
        "var g=function(k){for(var q=0;q<c.length;q++){if(c[q].toLowerCase().indexOf(k)===0)return q<d.length?d[q]:'';}return '';};" & _
        "out.push('DRV|'+[nameNear(rows[i]),g('dob'),g('gender'),g('mar'),g('license')||g('id#'),g('st'),g('driver type'),g('status')].join('|'));continue;}" & _
        "if(c.indexOf('VIN#')>=0&&c.indexOf('Year')>=0){var v=cells(rows[i+1]);" & _
        "var h=function(k){var q=c.indexOf(k);return q>=0&&q<v.length?v[q]:'';};" & _
        "out.push('VEH|'+[h('VIN#'),h('Year'),h('Make'),h('Model'),h('Status')].join('|'));continue;}" & _
        "if(c[0]==='Code'&&/Coverage Description/.test(line)){if(inCov)out.push('COV|'+cov+'|'+coll+'|'+comp);cov++;inCov=true;coll='';comp='';continue;}" & _
        "if(inCov&&(c[0]==='CLP'||c[0]==='COLL'))coll=c[2]||'';" & _
        "if(inCov&&(c[0]==='CPP'||c[0]==='COMP'))comp=c[2]||'';"
    CrownPullUnitedJs = CrownPullUnitedJs & _
        "if(inCov&&/^Total Premium for Vehicle/i.test(c[0]||'')){out.push('COV|'+cov+'|'+coll+'|'+comp);inCov=false;}" & _
        "if(c[0]==='Lienholder'&&c[1]==='Address'){inLien=true;continue;}" & _
        "if(inLien){if(c.length===2&&c[0]&&!/Coverages|Information/i.test(c[0]))out.push('LIEN|'+c[0]+'|'+c[1]);else inLien=false;}}" & _
        "if(inCov)out.push('COV|'+cov+'|'+coll+'|'+comp);" & _
        "return out.join('\n');"
End Function
