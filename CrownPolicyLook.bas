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

        EnterData drv, Keys, "QuickPolicyLookupControl_NumberInsCombo_I", Replace(policyNo, " ", ""), "ID", "", True
        ClickElement drv, "QuickPolicyLookupControl_FireLookupImageButton", "ID"
        CrownLookSettle drv, 6000

        n = CrownLookWrite(ws, nextRow, "Verve", policyNo, "page after the lookup", drv)

        tabs = Array("Policy Info", "Drivers/Vehicles/Coverages", "Drivers", "Vehicles", "Coverages")

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
