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

' True while a pull is filling a NEW Edit Data row (compare against nothing).
Private mPullBlank As Boolean

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
' something different on the PolicyPull sheet. Nothing changes until Yes
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

    ' "GAI 206097" found nothing, "206097" did (Tony, 22 Oct) - so when the
    ' number has letters in front and nothing came back, try the digits.
    If InStr(answer, "DRV|") = 0 And InStr(answer, "VEH|") = 0 And CrownDigits(policyNo) <> policyNo Then
        drv.Get CrownLookUnitedUrl(drv, CrownDigits(policyNo))
        CrownLookSettle drv, 6000

        On Error Resume Next
        answer = CStr(drv.ExecuteScript(CrownPullUnitedJs()))
        On Error GoTo 0
    End If

    On Error Resume Next
    drv.Quit
    On Error GoTo 0

    If InStr(answer, "DRV|") = 0 And InStr(answer, "VEH|") = 0 Then
        MsgBox "United's page for " & policyNo & " did not show its drivers or vehicles, so nothing has been changed." & vbCrLf & _
               "(Check the number; if it is right, run CrownRecordPolicyPages for it and send me the PolicyLook sheet.)", _
               vbExclamation, "Crown Superior"
        Exit Sub
    End If

    CrownPullApply answer, policyNo, "United", False
    Exit Sub

Finish:
    On Error Resume Next
    drv.Quit
End Sub

' Compares what a carrier's policy page said (the lines its reader
' returns) with Edit Data, lists it on the PolicyPull sheet, and on Yes
' puts it on Edit Data - into the customer open there, or a new row.
Private Sub CrownPullApply(ByVal answer As String, ByVal policyNo As String, ByVal carrier As String, ByVal newRow As Boolean)
    Dim wsEdit As Worksheet, ws As Worksheet
    Dim lines() As String, f() As String
    Dim changes As Collection, notes As Collection, ch As Variant
    Dim drivers As Collection, vehicles As Collection, covers As Collection, liens As Collection
    Dim insured As Variant, addr As Variant, k As Long, row As Long, shown As String, n As Long
    Dim dCols As Variant, vCols As Variant, col As String, item As Variant
    Dim vNames As Collection, vn As Variant, w() As String

    Set wsEdit = ThisWorkbook.Worksheets("Edit Data")
    mPullBlank = newRow
    Set vNames = New Collection

    ' --- what the carrier says -------------------------------------------------
    Set drivers = New Collection: Set vehicles = New Collection
    Set covers = New Collection: Set liens = New Collection
    insured = Array("", "", "", "", "", "", "")
    addr = Array("", "", "", "", "")

    lines = Split(answer, vbLf)
    For k = 0 To UBound(lines)
        f = Split(lines(k) & "||||||||", "|")
        Select Case f(0)
            Case "INS": insured = f
            Case "DRV"
                ' only drivers the carrier still has on the policy
                If Len(f(8)) = 0 Or InStr(1, f(8), "active", vbTextCompare) > 0 Then drivers.Add f
            Case "VEH"
                If Len(f(5)) = 0 Or InStr(1, f(5), "active", vbTextCompare) > 0 Then vehicles.Add f
            Case "COV": covers.Add f
            Case "LIEN": liens.Add f
            Case "ADDR": addr = f
            Case "VEHNAME": vNames.Add f
        End Select
    Next k

    ' --- compare with Edit Data ----------------------------------------------
    Set changes = New Collection
    Set notes = New Collection
    dCols = Array("D", "F", "H", "J")
    vCols = Array("P", "R", "T", "V")

    If newRow Then CrownPullAdd changes, wsEdit, "H7", "Insurance Company", carrier
    CrownPullAdd changes, wsEdit, "C7", "Policy No", Trim$(CStr(insured(6)))
    CrownPullAdd changes, wsEdit, "H8", "Phone", CrownPullPhone(CStr(insured(3)))
    CrownPullAdd changes, wsEdit, "C8", "Email", LCase$(Trim$(CStr(insured(4))))
    If Len(insured(2)) > 0 And Len(addr(1)) > 0 Then
        CrownPullAdd changes, wsEdit, "C9", "Address", UCase$(CStr(addr(1)))
        CrownPullAdd changes, wsEdit, "C10", "City", UCase$(CStr(addr(2)))
        CrownPullAdd changes, wsEdit, "K10", "State", UCase$(CStr(addr(3)))
        CrownPullAdd changes, wsEdit, "H10", "ZIP Code", CStr(addr(4))
    ElseIf Len(insured(2)) > 0 Then
        notes.Add carrier & "'s mailing address: " & insured(2) & " (address not changed - check it by eye)"
    End If

    For k = 1 To drivers.Count
        If k > 4 Then
            notes.Add carrier & " has more than 4 drivers - driver " & k & " (" & drivers(k)(1) & ") has no column on Edit Data"
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
            CrownPullAdd changes, wsEdit, col & "18", "Driver " & k & " - first name", UCase$(CrownPullFirst(CStr(item(1))))
            CrownPullAdd changes, wsEdit, col & "19", "Driver " & k & " - last name", UCase$(CrownPullLast(CStr(item(1))))
        Else
            notes.Add carrier & " did not show driver " & k & "'s name (DOB " & item(2) & ") - name left as it is"
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
                          ") that " & carrier & " does not - left as it is"
            End If
        End If
    Next k

    For k = 1 To vehicles.Count
        If k > 4 Then
            notes.Add carrier & " has more than 4 vehicles - vehicle " & k & " has no column on Edit Data"
            Exit For
        End If

        item = vehicles(k)
        col = vCols(k - 1)

        ' The full make where another page spells it out ("DODG" -> "DODGE").
        For Each vn In vNames
            w = Split(Application.WorksheetFunction.Trim(CStr(vn(2))), " ")
            If UBound(w) >= 2 Then
                If w(0) = CStr(item(2)) And InStr(1, CStr(vn(2)), CStr(item(4)), vbTextCompare) > 0 Then item(3) = w(1)
            End If
        Next vn

        If k > 1 Then CrownPullAdd changes, wsEdit, col & "16", "Vehicle " & k & " - is there a vehicle", "Yes"
        CrownPullAdd changes, wsEdit, col & "17", "Vehicle " & k & " - VIN", UCase$(CStr(item(1)))
        CrownPullAdd changes, wsEdit, col & "18", "Vehicle " & k & " - make", UCase$(CStr(item(3)))
        CrownPullAdd changes, wsEdit, col & "19", "Vehicle " & k & " - year", CStr(item(2))
        CrownPullAdd changes, wsEdit, col & "20", "Vehicle " & k & " - model", UCase$(CStr(item(4)))

        If k <= covers.Count Then
            answer = CrownPullDeductible(CStr(covers(k)(2)), CStr(covers(k)(3)))
            If Left$(answer, 1) = "?" Then
                notes.Add "Vehicle " & k & ": " & carrier & " has collision " & covers(k)(2) & " / comprehensive " & covers(k)(3) & _
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
        notes.Add carrier & " lists lienholders without saying which vehicle each is for - they were matched in order, check them"
    End If

    ' --- the report ---------------------------------------------------------
    Set ws = CrownPullSheet()
    ws.Range("A1:E1").value = Array("Edit Data cell", "What", "Edit Data has", carrier & " says", "Done")
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
        MsgBox "Edit Data already matches " & carrier & " for " & policyNo & "." & IIf(notes.Count > 0, vbCrLf & vbCrLf & notes.Count & " note(s) on the PolicyPull sheet.", ""), _
               vbInformation, "Crown Superior"
        mPullBlank = False
        Exit Sub
    End If

    If n > 14 Then shown = shown & "... and " & (n - 14) & " more (all on the PolicyPull sheet)" & vbCrLf

    If MsgBox(changes.Count & IIf(newRow, " value(s) from " & carrier & " for a NEW Edit Data row - ", " difference(s) between Edit Data and " & carrier & " for ") & policyNo & ":" & vbCrLf & vbCrLf & shown & vbCrLf & _
              IIf(notes.Count > 0, notes.Count & " note(s) on the PolicyPull sheet." & vbCrLf & vbCrLf, "") & _
              IIf(newRow, "Start the new row with these values and save it now?", "Put " & carrier & "'s values on Edit Data and save this customer now?"), vbYesNo + vbQuestion, "Crown Superior") = vbNo Then
        ws.Activate
        mPullBlank = False
        Exit Sub
    End If

    ' --- apply, then save with his own Update (or his Send, for a new row) --
    If newRow Then ClearInputFields
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

    If Not newRow Then
        UpdateDataToOutput
        mPullBlank = False
        Exit Sub
    End If

    SendDataToOutput
    CrownPullShowNewRow wsEdit, CrownDigits(policyNo)
    mPullBlank = False
End Sub

' After his Send: show the new row on Edit Data (its number in C5).
Private Sub CrownPullShowNewRow(ByVal wsEdit As Worksheet, ByVal digits As String)
    Dim wsOut As Worksheet, r As Long

    Set wsOut = ThisWorkbook.Worksheets("Output")

    For r = wsOut.Cells(wsOut.Rows.Count, colPolicyNo).End(xlUp).Row To 2 Step -1
        If CrownDigits(CStr(wsOut.Cells(r, colPolicyNo).value)) = digits Then
            wsEdit.Range("C5").value = r
            MsgBox "Saved as a new customer - row " & r & " of Output, now open on Edit Data.", vbInformation, "Crown Superior"
            Exit Sub
        End If
    Next r
End Sub

' One difference: Array(cell, what, Edit Data has, United says, is a date).
Private Sub CrownPullAdd(ByVal changes As Collection, ByVal ws As Worksheet, ByVal addr As String, _
                         ByVal what As String, ByVal united As String, Optional ByVal asDate As Boolean = False)
    Dim have As String

    united = Trim$(united)
    If Len(united) = 0 Then Exit Sub

    have = Trim$(ws.Range(addr).Text)
    If mPullBlank Then have = ""

    If asDate Then
        If Not IsDate(united) Then Exit Sub
        If IsDate(ws.Range(addr).value) And Not mPullBlank Then
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
        If StrComp(ws.Name, "PolicyPull", vbTextCompare) = 0 Then
            ws.Cells.ClearContents
            Set CrownPullSheet = ws
            Exit Function
        End If
    Next ws

    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = "PolicyPull"
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


' =====================================================================
' Pull a Progressive policy into Edit Data
' ---------------------------------------------------------------------
' Alt+F8 > CrownPullProgressivePolicy > Run, type the policy number.
' Progressive has one page with everything (Tony, 21 Oct): Policy and
' Coverages - the insured, address, phone, email, next amount due, the
' people, the vehicles with VINs, each vehicle's comprehensive and
' collision deductibles. Signs in like the payment run (code and all),
' opens that page, opens every "Show more" so the driver details are in
' it, and hands it to the same compare / apply as the United pull.
'
' Customers Progressive quoted straight through, who were never in Excel,
' go into a NEW Edit Data row (saved with his own Send).
' =====================================================================
Public Sub CrownPullProgressivePolicy()
    Dim wsEdit As Worksheet
    Dim drv As ChromeDriver, clsDrv As Chrm
    Dim policyNo As String, digits As String, answer As String, who As String
    Dim newRow As Boolean, how As VbMsgBoxResult
    Dim started As Date

    Set wsEdit = ThisWorkbook.Worksheets("Edit Data")

    policyNo = Trim$(InputBox("Progressive policy number to pull into Edit Data:", "Crown Superior", _
                              IIf(InStr(1, wsEdit.Range("H7").Text, "progressive", vbTextCompare) > 0, Trim$(wsEdit.Range("C7").Text), "")))
    digits = CrownDigits(policyNo)
    If Len(digits) < 6 Then Exit Sub

    Select Case CrownPullWhichRow(wsEdit, digits)
        Case 0: newRow = False
        Case 1: newRow = True
        Case Else: Exit Sub
    End Select

    Set drv = CrownCodeDriver("Progressive", clsDrv)
    If Not CrownChromeOk(drv) Then Exit Sub
    If Not CrownPgSignIn(drv) Then GoTo Finish

    drv.Get "https://policyservicing.apps.foragentsonly.com/app/policy-hub/" & digits & "/policy-and-coverages"

    ' The page builds itself after it loads - wait for the vehicles.
    started = Now
    Do
        answer = CStr(CrownPullJs(drv, "var t=(document.body&&document.body.innerText)||'';" & _
                                       "return (/Policy and Coverages/.test(t)&&t.indexOf(arguments[0])>=0&&/\bVIN\s+[A-Z0-9]{17}/.test(t))?'ok':'';", digits))
        If answer = "ok" Then Exit Do
        drv.Wait 1000
        DoEvents
    Loop While (Now - started) * 86400 < 60

    If answer <> "ok" Then
        MsgBox "Progressive's Policy and Coverages page for " & digits & " did not come up, so nothing has been changed." & vbCrLf & _
               "(Check the number. If it is right, send me a screenshot of what Chrome shows.)", vbExclamation, "Crown Superior"
        GoTo Finish
    End If

    ' "Show more" on every driver and vehicle, twice (some open late).
    CrownPullJs drv, CrownPgExpandJs()
    drv.Wait 2500
    CrownPullJs drv, CrownPgExpandJs()
    drv.Wait 1500

    answer = CStr(CrownPullJs(drv, CrownPgPullJs()))

    On Error Resume Next
    drv.Quit
    On Error GoTo 0

    If InStr(answer, "DRV|") = 0 And InStr(answer, "VEH|") = 0 Then
        MsgBox "Could not read the drivers or vehicles on Progressive's page for " & digits & ", so nothing has been changed.", _
               vbExclamation, "Crown Superior"
        Exit Sub
    End If

    CrownPullApply answer, digits, "Progressive", newRow
    Exit Sub

Finish:
    On Error Resume Next
    drv.Quit
End Sub

Private Function CrownPullJs(ByVal drv As ChromeDriver, ByVal script As String, Optional ByVal arg As String = "") As String
    Dim v As Variant

    On Error Resume Next
    If Len(arg) > 0 Then v = drv.ExecuteScript(script, arg) Else v = drv.ExecuteScript(script)
    If Err.Number <> 0 Then v = ""
    On Error GoTo 0

    If IsNull(v) Or IsEmpty(v) Then v = ""
    CrownPullJs = CStr(v)
End Function

Private Function CrownPgExpandJs() As String
    CrownPgExpandJs = _
        "var n=0;[].slice.call(document.querySelectorAll('button,a,[role=button]')).forEach(function(b){" & _
        "var t=(b.innerText||'').trim(),a=b.getAttribute('aria-label')||'';" & _
        "if((/^show more$/i.test(t)||/^show more details/i.test(a))&&b.getAttribute('aria-expanded')!=='true'){b.click();n++;}});return String(n);"
End Function

' Progressive's Policy and Coverages page -> the same lines as the United
' reader, plus ADDR|street|city|state|zip.
Private Function CrownPgPullJs() As String
    Dim s As String

    ' Read as lines of text, written so it works whether the page puts the
    ' name, the address or the car on one line or several.
    s = "var B=(document.body.innerText||'');var M=((document.querySelector('main')||document.body).innerText||'');" & _
        "var L=M.split('\n').map(function(x){return x.replace(/\s+/g,' ').trim();}).filter(function(x){return x;});" & _
        "var c=function(x){return (x||'').replace(/[|]/g,' ').replace(/\s+/g,' ').trim();};" & _
        "var at=function(re,from,to){for(var i=from||0;i<(to||L.length);i++)if(re.test(L[i]))return i;return -1;};" & _
        "var junk=/^(add a (driver|vehicle)|show (more|less)|remove|update|replace|driving record|people|vehicles|go to vehicle|help)/i;" & _
        "var ST={alabama:'AL',alaska:'AK',arizona:'AZ',arkansas:'AR',california:'CA',colorado:'CO',connecticut:'CT',delaware:'DE',florida:'FL',georgia:'GA',hawaii:'HI',idaho:'ID',illinois:'IL',indiana:'IN',iowa:'IA',kansas:'KS',kentucky:'KY',louisiana:'LA',maine:'ME',maryland:'MD',massachusetts:'MA',michigan:'MI',minnesota:'MN',mississippi:'MS',missouri:'MO',montana:'MT',nebraska:'NE',nevada:'NV','new hampshire':'NH','new jersey':'NJ','new mexico':'NM','new york':'NY','north carolina':'NC','north dakota':'ND',ohio:'OH',oklahoma:'OK',oregon:'OR',pennsylvania:'PA','rhode island':'RI','south carolina':'SC','south dakota':'SD',tennessee:'TN',texas:'TX',utah:'UT',vermont:'VT',virginia:'VA',washington:'WA','west virginia':'WV',wisconsin:'WI',wyoming:'WY','district of columbia':'DC'};" & _
        "var stOf=function(x){x=(x||'').trim();return ST[x.toLowerCase()]||(/^[A-Za-z]{2}$/.test(x)?x.toUpperCase():x);};" & _
        "var out=[];var pol=((M.match(/Auto\s+(\d{6,})/)||[])[1])||'';" & _
        "var stat=((B.match(/Primary named insured\s*([A-Za-z ]+?)\s*,/)||[])[1])||'';"
    s = s & _
        "var ci=at(/^Contact information/),pp=at(/^Paperless/,ci>0?ci:0);var cb=ci>=0?L.slice(ci+1,pp>ci?pp:ci+8).join('\n'):'';" & _
        "var email=((cb.match(/[\w.+\-]+@[\w\-]+\.[\w.\-]+/)||[])[0])||'';var phone=((cb.match(/\((?:Mobile|Home|Work|Cell)\)\s*([\d\-\s().]{10,})/i)||[])[1])||'';" & _
        "var rest=cb.replace(email,' ').replace(/\((?:Mobile|Home|Work|Cell)\)\s*[\d\-\s().]{10,}/i,' ').replace(/Update (Email|Phone Number|Address)/g,'\n');" & _
        "var street='',city='',st='',zip='';var am=rest.match(/([^\n]*?\d[^\n]*?)\s*(\n\s*)?([A-Za-z .'\-]+),\s*([A-Za-z .]+?)\s+(\d{5})/);" & _
        "if(am){street=am[1].trim();city=am[3].trim();st=stOf(am[4]);zip=am[5];" & _
        "if(!am[2]){var sm=(street+' '+city).match(/^(.*\b(?:ST|STREET|AVE|AVENUE|RD|ROAD|DR|DRIVE|LN|LANE|CT|COURT|WAY|BLVD|PKWY|CIR|CIRCLE|PL|PLACE|TRL|TRAIL|HWY|TER|LOOP|RUN|PT|XING|SQ|PATH|ROW)\b(?:\s+(?:N|S|E|W|NE|NW|SE|SW))?(?:\s+(?:APT|UNIT|STE|#)\s*\S+)?)\s+(.+)$/i);" & _
        "if(sm){street=sm[1];city=sm[2];}}}" & _
        "var pay=((M.match(/Next amount due[^\n]*(\n[^\n]*)?/)||[])[0])||'';"
    s = s & _
        "var pe=at(/^People$/),ve=at(/^Vehicles$/,pe>0?pe:0),co=at(/^Vehicle Coverages$/,ve>0?ve:0);if(co<0)co=at(/^Coverages$/,ve>0?ve:0);" & _
        "var who='';var dl=[];for(var i=(pe<0?0:pe);i<(ve>0?ve:L.length);i++){if(/^Date of birth$/i.test(L[i]))dl.push(i);}" & _
        "dl.forEach(function(d,n){var role=L[d-1]||'',nm=L[d-2]||'';if(junk.test(nm))nm='';" & _
        "if(nm&&!/\s/.test(nm)&&L[d-3]&&!junk.test(L[d-3]))nm=L[d-3]+' '+nm;" & _
        "var end=(n+1<dl.length)?dl[n+1]-3:(ve>0?ve:L.length);var g='',ms='',lic='',ls='',rel='',dst='';" & _
        "for(var k=d+2;k<end;k++){var lb=L[k].toLowerCase().replace(/:$/,''),v=L[k+1]||'';" & _
        "if(/^(gender|sex)$/.test(lb))g=v;else if(/^marital/.test(lb))ms=v;else if(/(license|licence) (number|#|no\.?)$|^driver'?s license$/.test(lb))lic=v;" & _
        "else if(/license state|state (licensed|of license)|licensed in/.test(lb))ls=stOf(v);else if(/^relationship/.test(lb))rel=v;else if(/^(driver )?status$/.test(lb))dst=v;}" & _
        "if(n===0)who=nm;" & _
        "out.push('DRV|'+[c(nm),c(L[d+1]),c(g),c(ms),c(lic),c(ls),c(role+(rel?' '+rel:'')),c(dst)].join('|'));});"
    s = s & _
        "var vEnd=co>0?co:L.length;var vl=[];for(var i=(ve<0?0:ve);i<vEnd;i++){if(/VIN\s+[A-HJ-NPR-Z0-9]{17}/i.test(L[i]))vl.push(i);}" & _
        "vl.forEach(function(v,n){var vin=L[v].match(/VIN\s+([A-HJ-NPR-Z0-9]{17})/i)[1].toUpperCase();var desc=L[v].replace(/VIN\s+\S+/i,'').trim();" & _
        "if(!/^\d{4}\s/.test(desc)){var back=[];for(var k=v-1;k>=0&&back.length<3&&!junk.test(L[k]);k--)back.unshift(L[k]);desc=(back.join(' ')+' '+desc).trim();}" & _
        "var ym=desc.match(/(\d{4})\s+(\S+)\s+(.+)$/)||['','','',''];out.push('VEH|'+[vin,ym[1],c(ym[2]),c(ym[3]),''].join('|'));" & _
        "var end=(n+1<vl.length)?vl[n+1]:vEnd;for(var k=v+1;k<end;k++){if(/lien ?holder|loss payee|lessor/i.test(L[k])&&L[k+1]){" & _
        "var adr=(L[k+2]&&/\d/.test(L[k+2])&&!junk.test(L[k+2]))?L[k+2]:'';out.push('LIEN|'+c(L[k+1])+'|'+c(adr));break;}}});" & _
        "if(co>=0){var cv=0,coll='',comp='',ded=function(k){var v=L[k+1];if(/^help$/i.test(v||''))v=L[k+2];return /no coverage|not eligible|declined/i.test(v||'')?'0':((v||'').match(/\$?([\d,]+)\s*deductible/i)||['',''])[1].replace(/,/g,'');};" & _
        "for(var i=co;i<L.length;i++){if(/VIN\s+[A-HJ-NPR-Z0-9]{17}/i.test(L[i])){if(cv)out.push('COV|'+cv+'|'+coll+'|'+comp);cv++;coll='';comp='';}" & _
        "else if(cv&&/^Collision( Help)?$/i.test(L[i]))coll=ded(i);else if(cv&&/^Comprehensive( Help)?$/i.test(L[i]))comp=ded(i);else if(/^Total policy premium/i.test(L[i]))break;}if(cv)out.push('COV|'+cv+'|'+coll+'|'+comp);}"
    s = s & _
        "out.unshift('ADDR|'+[c(street),c(city),c(st),c(zip)].join('|'));" & _
        "out.unshift('INS|'+[c(who),c(street+', '+city+' '+st+' '+zip),c(phone),c(email),c(stat),pol,''].join('|'));" & _
        "out.push('PAY|'+c(pay));return out.join('\n');"

    CrownPgPullJs = s
End Function

' 0 = the customer open on Edit Data, 1 = a new row, -1 = stop.
Private Function CrownPullWhichRow(ByVal wsEdit As Worksheet, ByVal digits As String) As Long
    Dim who As String, how As VbMsgBoxResult

    If Val(wsEdit.Range("C5").value) >= 2 And CrownDigits(wsEdit.Range("C7").Text) = digits Then
        CrownPullWhichRow = 0
        Exit Function
    End If

    who = Trim$(wsEdit.Range("D18").Text & " " & wsEdit.Range("D19").Text)
    how = MsgBox("Is " & digits & " a NEW customer for Excel?" & vbCrLf & vbCrLf & _
                 "Yes = start a new Edit Data row for it" & vbCrLf & _
                 "No = put it on the customer open on Edit Data now" & IIf(Len(who) > 0, " (" & who & ")", "") & vbCrLf & _
                 "Cancel = stop", vbYesNoCancel + vbQuestion, "Crown Superior")

    If how = vbCancel Then
        CrownPullWhichRow = -1
    ElseIf how = vbYes Then
        CrownPullWhichRow = 1
    ElseIf Val(wsEdit.Range("C5").value) < 2 Then
        MsgBox "Open the customer on Edit Data first (their row number in C5), then run this again.", vbExclamation, "Crown Superior"
        CrownPullWhichRow = -1
    Else
        CrownPullWhichRow = 0
    End If
End Function


' =====================================================================
' Pull a First Acceptance policy into Edit Data
' ---------------------------------------------------------------------
' Alt+F8 > CrownPullFirstAcceptancePolicy > Run, type the policy number
' (like IAGA-17304). Signs in like the payment run, opens the policy from
' Open Search, and reads its tabs - Policy Summary (name, phone, email,
' address), Drivers (date of birth, gender, marital status, license and
' state, excluded or not), Vehicles (VIN, lienholder) and Coverages (each
' vehicle's collision / comprehensive deductible). Then the same compare
' and apply as the other pulls.
' =====================================================================
Public Sub CrownPullFirstAcceptancePolicy()
    Dim wsEdit As Worksheet
    Dim drv As ChromeDriver, clsDrv As Chrm
    Dim Keys As New Selenium.Keys
    Dim policyNo As String, digits As String, answer As String, got As String
    Dim newRow As Boolean, tabName As Variant, xp As String
    Dim started As Date

    Set wsEdit = ThisWorkbook.Worksheets("Edit Data")

    policyNo = UCase$(Trim$(InputBox("First Acceptance policy number to pull into Edit Data (like IAGA-17304):", "Crown Superior", _
                              IIf(InStr(1, wsEdit.Range("H7").Text, "first acceptance", vbTextCompare) > 0, Trim$(wsEdit.Range("C7").Text), ""))))
    digits = CrownDigits(policyNo)
    If Len(digits) < 3 Then Exit Sub

    Select Case CrownPullWhichRow(wsEdit, digits)
        Case 0: newRow = False
        Case 1: newRow = True
        Case Else: Exit Sub
    End Select

    Set drv = CrownCodeDriver("First Acceptance", clsDrv)
    If Not CrownChromeOk(drv) Then Exit Sub
    If Not CrownFaSignIn(drv) Then GoTo Finish

    ' His payment-run way in: Policy Services > Active & Reinstatable >
    ' open the first one, which brings up Open Search.
    xp = "//div[contains(@class, 'aspect-square') and contains(., 'Policy Services')]"
    If Not CrownFaFindFrame(drv, xp, 60) Then
        MsgBox "First Acceptance's Policy Services tile did not show up, so nothing has been changed.", vbExclamation, "Crown Superior"
        GoTo Finish
    End If

    CrownFaClick drv, xp
    CrownFaClick drv, "//button[contains(., 'Active & Reinstatable Policies')]"
    CrownFaClick drv, "//table/tbody/tr[1]/td[2]//button"
    CrownFaClick drv, "//button[.//span[text()='Open Search']]"

    ' Type the number into the search box, then open that policy.
    answer = CrownPullJs(drv, _
        "var d=document.querySelector('[data-component=x-policy-search-dialog]')||document;" & _
        "var i=[].slice.call(d.querySelectorAll('input')).filter(function(x){var r=x.getBoundingClientRect();return x.type!=='hidden'&&r.width>0;})[0];" & _
        "if(!i)return 'nobox';[].slice.call(document.querySelectorAll('[data-crown-q]')).forEach(function(e){e.removeAttribute('data-crown-q');});i.setAttribute('data-crown-q','1');return 'ok';")

    If answer = "ok" Then
        On Error Resume Next
        drv.FindElementByCss("[data-crown-q]").Clear
        drv.FindElementByCss("[data-crown-q]").SendKeys policyNo
        drv.FindElementByCss("[data-crown-q]").SendKeys Keys.Enter
        On Error GoTo 0
        drv.Wait 2500
        WaitSpinnerGone drv
    End If

    xp = "//div[@data-component='x-policy-search-dialog']//table/tbody/tr/td[2]//button[translate(normalize-space(.),'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ')='" & policyNo & "']"

    If Not CrownFaWaitX(drv, xp, 25) Then
        MsgBox "First Acceptance's search did not show policy " & policyNo & ", so nothing has been changed." & vbCrLf & _
               "(Check the number. Only policies First Acceptance lists under Open Search can be pulled.)", vbExclamation, "Crown Superior"
        GoTo Finish
    End If

    CrownFaClick drv, xp

    If Not CrownFaWaitX(drv, "//span[contains(text(),'Policy Status')]", 30) Then
        MsgBox "Policy " & policyNo & " did not open on First Acceptance, so nothing has been changed.", vbExclamation, "Crown Superior"
        GoTo Finish
    End If

    ' Each tab in turn, read while it is showing.
    For Each tabName In Array("Policy Summary", "Drivers", "Vehicles", "Coverages")
        CrownPullJs drv, "var n=arguments[0];var b=[].slice.call(document.querySelectorAll('button')).filter(function(x){return (x.innerText||'').trim()===n;})[0];if(b)b.click();return b?'ok':'';", CStr(tabName)
        drv.Wait 2000
        WaitSpinnerGone drv

        If tabName = "Policy Summary" Then
            ' Open the Addresses section so the street and city are in the text.
            CrownPullJs drv, "var a=[].slice.call(document.querySelectorAll('span,button,div')).filter(function(x){return x.children.length<3&&(x.innerText||'').trim()==='Addresses';})[0];" & _
                             "if(a){var t=a.closest('[class*=disclosure-trigger],button')||a;t.click();}return 'ok';"
            drv.Wait 1200
        End If

        answer = CrownPullJs(drv, CrownFaPullJs(), CStr(tabName))
        If Len(answer) > 0 Then got = got & answer & vbLf
    Next tabName

    On Error Resume Next
    drv.Quit
    On Error GoTo 0

    If InStr(got, "DRV|") = 0 And InStr(got, "VEH|") = 0 Then
        MsgBox "Could not read the drivers or vehicles for " & policyNo & " on First Acceptance, so nothing has been changed.", _
               vbExclamation, "Crown Superior"
        Exit Sub
    End If

    CrownPullApply got, policyNo, "First Acceptance", newRow
    Exit Sub

Finish:
    On Error Resume Next
    drv.Quit
End Sub

' One tab of a First Acceptance policy (arguments[0] = the tab) -> lines.
Private Function CrownFaPullJs() As String
    Dim s As String

    s = "var tab=arguments[0];var L=(document.body.innerText||'').split('\n').map(function(x){return x.replace(/\s+/g,' ').trim();}).filter(function(x){return x;});" & _
        "var c=function(x){return (x||'').replace(/[|]/g,' ').trim();};var lab=function(x){return /:$/.test(x||'');};" & _
        "var after=function(i){var v=L[i+1]||'';return lab(v)?'':v;};var out=[];" & _
        "var pi=-1;for(var i=0;i<L.length;i++){if(/^[A-Z]{2,6}-\d{3,}$/.test(L[i])){pi=i;break;}}var pol=pi>=0?L[pi]:'',who=pi>0?L[pi-1]:'';" & _
        "var st='';for(var i=0;i<L.length;i++){if(L[i]==='Policy Status'){st=L[i+1]||'';break;}}"
    s = s & _
        "if(tab==='Policy Summary'){var ph='',em='';for(var i=0;i<L.length;i++){if(/^(Home|Mobile|Cell|Work|Primary) Phone:$/i.test(L[i])&&!ph)ph=after(i);if(/^Email( Address)?:$/i.test(L[i]))em=after(i);}" & _
        "var a0=-1;for(var i=0;i<L.length;i++)if(L[i]==='Addresses'){a0=i;break;}var street='',city='',sta='',zip='';" & _
        "if(a0>=0){for(var i=a0+1;i<Math.min(L.length,a0+14);i++){var m=L[i].match(/^(.*?)\s*,\s*([A-Z]{2})\s+(\d{5})/);if(m&&!/LAWRENCEVILLE HWY|TUCKER/i.test(L[i])){city=m[1];sta=m[2];zip=m[3];street=/\d/.test(L[i-1]||'')?L[i-1]:'';break;}}}" & _
        "out.push('INS|'+[c(who),c(street+(street?', ':'')+city+' '+sta+' '+zip),c(ph),c(em).toLowerCase(),c(st),c(pol),''].join('|'));" & _
        "if(city)out.push('ADDR|'+[c(street),c(city),c(sta),c(zip)].join('|'));}"
    s = s & _
        "if(tab==='Drivers'){for(var i=1;i<L.length;i++){if(L[i]!=='Date Of Birth:')continue;var nm=L[i-1],v={};" & _
        "for(var k=i;k<L.length&&!/^Violations/.test(L[k]);k++){if(lab(L[k]))v[L[k].replace(/:$/,'').toLowerCase()]=after(k);}" & _
        "var g=(v['gender']||'').toUpperCase(),ms=(v['marital status']||'').toUpperCase(),dl=(v['dl number']||'').trim(),ls='',lic=dl;" & _
        "var dm=dl.match(/^([A-Z]{2})\s+(\S+)$/);if(dm){ls=dm[1];lic=dm[2];}" & _
        "out.push('DRV|'+[c(nm),c(v['date of birth']),g==='F'?'Female':(g==='M'?'Male':c(v['gender'])),ms==='M'?'Married':(ms?'Single':''),c(lic),c(ls),c(v['driver type']),''].join('|'));}}"
    s = s & _
        "if(tab==='Vehicles'){for(var i=1;i<L.length;i++){var vm=L[i].match(/^VIN:\s*([A-HJ-NPR-Z0-9]{17})$/i);if(!vm)continue;" & _
        "var ym=(L[i-1]||'').match(/^(\d{4})\s+(\S+)\s+(.+)$/)||['','','',''];out.push('VEH|'+[vm[1].toUpperCase(),ym[1],c(ym[2]),c(ym[3]),''].join('|'));" & _
        "for(var k=i;k<Math.min(L.length,i+30);k++){if(L[k]==='Lienholder:'){var lh=after(k);if(lh&&!/^n\/?a$/i.test(lh))out.push('LIEN|'+c(lh)+'|');break;}if(k>i&&/^VIN:/.test(L[k]))break;}}}"
    s = s & _
        "if(tab==='Coverages'){var n=0;[].slice.call(document.querySelectorAll('table')).forEach(function(t){var h=[].slice.call(t.querySelectorAll('th')).map(function(x){return (x.innerText||'').trim();});" & _
        "var di=h.indexOf('Deductible');if(di<0||h.indexOf('Coverage Type')<0)return;n++;" & _
        "var hd='';var e=t;for(var u=0;u<8&&e&&!hd;u++){var p=e.previousElementSibling;while(p&&!hd){var tx=(p.innerText||'').trim();var mm=tx.match(/(\d{4}\s+[A-Z][A-Z0-9\-]*\s+[^\n]+)/);if(mm)hd=mm[1];p=p.previousElementSibling;}e=e.parentElement;}" & _
        "var coll='',comp='';[].slice.call(t.querySelectorAll('tbody tr')).forEach(function(r){var cs=[].slice.call(r.children).map(function(x){return (x.innerText||'').trim();});var ty=cs.join(' ').toUpperCase();" & _
        "var d=(cs[di]||'').replace(/[^0-9]/g,'');if(/COLLISION/.test(ty))coll=d||'0';else if(/COMPREHENSIVE|OTHER THAN COLL/.test(ty))comp=d||'0';});" & _
        "out.push('COV|'+n+'|'+coll+'|'+comp);if(hd)out.push('VEHNAME|'+n+'|'+c(hd));});}"
    s = s & "return out.join('\n');"

    CrownFaPullJs = s
End Function
