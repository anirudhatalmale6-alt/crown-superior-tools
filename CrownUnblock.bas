Attribute VB_Name = "CrownUnblock"
Option Explicit

' =====================================================================
' Crown Superior - getting past the things that stop a quote
' ---------------------------------------------------------------------
' His words:
'
'   "it stops just for the alert that photos are required because its a
'    full coverage policy and all the tool has to do is click ok or x it
'    out ... it now stops because united added a new excluded driver page
'    ... Sometimes it stops to select the city especially when the zip
'    code is 30349 and it just needs to select atlanta"
'
' Three different stops, and they are not the same kind of problem.
'
' The photos one is already half solved and nobody noticed. Mod12uaig's
' HandleWebsiteAlert ACCEPTS the alert and then asks you whether to carry
' on. So the alert is gone either way - the thing stopping the run is the
' question afterwards. Two known alerts are already waved through, in
' three copied-out blocks; this puts that list in one place and adds the
' photos one to it.
'
' The city one is a dropdown, and a dropdown can be handled on its own
' terms: take the city the quote form already has, and if that is not one
' of the choices take the first real one, which is what you do by hand.
'
' The excluded driver page I have never seen. It is new, it is United's,
' and I am not going to guess at it - the last time I built a page handler
' from a description it passed every test I wrote and failed every real
' policy. So this records what is actually on screen onto a QuoteLook
' sheet, and the handler gets written from that.
'
' WHAT TO CHANGE IN YOUR WORKBOOK - one line
'
'   In Mod12uaig, find this line (it is there four times):
'
'       If Not HandleWebsiteAlert(drv) Then Exit Sub
'
'   and leave it exactly as it is. Then find the function further down:
'
'       Function HandleWebsiteAlert(drv As Object) As Boolean
'
'   and put a single line at the top of it:
'
'       HandleWebsiteAlert = CrownHandleAlert(drv): Exit Function
'
' That is all. Everything below the new line stays where it is and stops
' being used, so it is still there to go back to.
' =====================================================================


' The alerts that do not need you.
'
' Each of these is United telling you something rather than asking you
' something, so the answer is always OK. Anything NOT on this list still
' stops and asks, which is the point - a new alert should get your
' attention the first time, and then be added here.
'
' Matched on a few words rather than the whole sentence, because United
' rewords these and a full-sentence match would quietly start stopping
' again.
Private Function CrownKnownAlerts() As Variant
    CrownKnownAlerts = Array( _
        "run a credit check", _
        "previous policy balance", _
        "photos", _
        "photo", _
        "pictures are required", _
        "picture", _
        "images are required")
End Function

Public Function CrownAlertIsKnown(ByVal text As String) As Boolean
    Dim known As Variant
    Dim i As Long

    known = CrownKnownAlerts()

    For i = LBound(known) To UBound(known)
        If InStr(1, text, CStr(known(i)), vbTextCompare) > 0 Then
            CrownAlertIsKnown = True

            Exit Function
        End If
    Next i
End Function


' Deal with a browser alert, and only stop for one we have not met.
'
' Returns True to carry on, False to stop this quote.
'
' The alert is accepted either way - that part was never the problem. The
' change is that a known one no longer puts a question in front of you.
Public Function CrownHandleAlert(ByVal drv As Object) As Boolean
    Dim text As String
    Dim answer As VbMsgBoxResult

    On Error Resume Next
    text = drv.SwitchToAlert.text

    If Err.Number <> 0 Then
        ' No alert at all. Nothing to do and nothing to report.
        Err.Clear
        On Error GoTo 0
        CrownHandleAlert = True

        Exit Function
    End If

    drv.SwitchToAlert.Accept
    Err.Clear
    On Error GoTo 0

    If CrownAlertIsKnown(text) Then
        CrownNoteBlocker "waved through: " & Left$(text, 120)
        CrownHandleAlert = True

        Exit Function
    End If

    ' Not one we know. Written down before asking, so that even if you
    ' press No the wording is kept and can be added to the list.
    CrownNoteBlocker "STOPPED, alert not recognised: " & Left$(text, 200)

    answer = SmartMsgBox( _
        "United said:" & vbCrLf & vbCrLf & text & vbCrLf & vbCrLf & _
        "This one is not on the list of alerts that get waved through." & vbCrLf & _
        "Press Yes to carry on, No to stop." & vbCrLf & vbCrLf & _
        "Either way it has been written on the QuoteLook sheet - send me " & _
        "that and I will add it to the list.", _
        vbYesNo + vbInformation, "Crown Superior")

    CrownHandleAlert = (answer = vbYes)
End Function


' Choose the city.
'
' United asks for it when a zip covers more than one - 30349 is Atlanta,
' College Park and Union City. The quote form already holds the city the
' customer gave, so that is what gets chosen; if it is not one of the
' choices, the first real one is taken, which is what you do by hand to
' get past it.
'
' Reports what it did rather than doing it silently - picking the wrong
' city changes the price, so it has to be something you can check
' afterwards rather than something you have to catch at the time.
Public Function CrownPickCity(ByVal drv As Object, ByVal wanted As String) As String
    Dim answer As String

    On Error Resume Next
    answer = drv.ExecuteScript( _
        "function clean(s){return ((s||'')+'').replace(/\s+/g,' ')" & _
        ".replace(/^ | $/g,'');}" & _
        "var want=" & CrownJsText(wanted) & ".toLowerCase();" & _
        "var picks=document.querySelectorAll('select');var box=null;" & _
        "for(var i=0;i<picks.length;i++){var s=picks[i];" & _
        "if(!s.getClientRects().length)continue;" & _
        "var who=((s.id||'')+' '+(s.name||'')).toLowerCase();" & _
        "if(who.indexOf('city')<0)continue;box=s;break;}" & _
        "if(!box)return 'no city box';" & _
        "if(box.options.length<2)return 'city box has nothing to choose from';" & _
        "var first=null;" & _
        "for(var j=0;j<box.options.length;j++){var o=box.options[j],t=clean(o.text);" & _
        "if(!t||!o.value)continue;" & _
        "if(first===null)first=o;" & _
        "if(t.toLowerCase()===want){box.value=o.value;" & _
        "box.dispatchEvent(new Event('change',{bubbles:true}));" & _
        "return 'chose '+t;}}" & _
        "if(first===null)return 'city box has nothing to choose from';" & _
        "box.value=first.value;" & _
        "box.dispatchEvent(new Event('change',{bubbles:true}));" & _
        "return 'no match for '+want+', took the first one: '+clean(first.text);")
    On Error GoTo 0

    If Len(answer) = 0 Then answer = "could not read the city box"

    If answer <> "no city box" Then CrownNoteBlocker "city: " & answer

    CrownPickCity = answer
End Function


' Everything on screen, written down.
'
' For the excluded driver page, which I have not seen. Run a quote, let
' it stop there, and send me the QuoteLook sheet - the page's own words
' are worth more than another description of it.
Public Function CrownLookAtPage(ByVal drv As Object) As String
    Dim answer As String

    On Error Resume Next
    answer = drv.ExecuteScript( _
        "function clean(s){return ((s||'')+'').replace(/\s+/g,' ')" & _
        ".replace(/^ | $/g,'');}" & _
        "var out=[],seen={},all=document.querySelectorAll(" & _
        "'input,select,button,a,label,h1,h2,h3,legend,td');" & _
        "for(var i=0;i<all.length;i++){var el=all[i];" & _
        "if(!el.getClientRects().length)continue;" & _
        "var what=el.tagName.toLowerCase();" & _
        "var who=(el.id||el.name||'');" & _
        "var t=clean(el.innerText||el.value||'');" & _
        "if(what==='select'){var opts=[];" & _
        "for(var k=0;k<el.options.length&&k<12;k++){opts.push(clean(el.options[k].text));}" & _
        "t=opts.join(' / ');}" & _
        "if(!who&&!t)continue;if(t.length>90)t=t.substring(0,90);" & _
        "var line=what+(who?' #'+who:'')+(t?'  = '+t:'');" & _
        "if(seen[line])continue;seen[line]=1;out.push(line);" & _
        "if(out.length>60)break;}" & _
        "return 'TITLE: '+clean(document.title)+String.fromCharCode(10)+out.join(String.fromCharCode(10));")
    On Error GoTo 0

    CrownLookAtPage = answer
End Function


' Write a line onto the QuoteLook sheet.
'
' Its own sheet rather than a message box, because the whole point is for
' the run not to stop. A note you read afterwards costs nothing; a box
' you have to close is the thing being removed.
Public Sub CrownNoteBlocker(ByVal what As String)
    Dim ws As Worksheet
    Dim row As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("QuoteLook")
    On Error GoTo 0

    If ws Is Nothing Then
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "QuoteLook"
        On Error GoTo 0

        If ws Is Nothing Then Exit Sub

        ws.Range("A1:C1").value = Array("When", "What happened", "What was on the page")
        ws.Columns(2).ColumnWidth = 70
        ws.Columns(3).ColumnWidth = 110
    End If

    row = ws.Cells(ws.Rows.Count, 1).End(xlUp).row + 1
    If row < 2 Then row = 2

    ws.Cells(row, 1).value = Now
    ws.Cells(row, 2).value = what
End Sub


' The same, with a picture of the page beside it.
Public Sub CrownNotePage(ByVal drv As Object, ByVal why As String)
    Dim ws As Worksheet
    Dim row As Long

    CrownNoteBlocker why

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("QuoteLook")
    On Error GoTo 0

    If ws Is Nothing Then Exit Sub

    row = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    If row < 2 Then Exit Sub

    ws.Cells(row, 3).value = CrownLookAtPage(drv)
End Sub


' A VBA string, as something JavaScript can read.
'
' Quotes and backslashes out of the way. A city called O'Fallon would
' otherwise end the string early and the script would not run at all -
' and a script that does not run comes back empty, which reads exactly
' like a page with no city box on it.
Private Function CrownJsText(ByVal text As String) As String
    Dim out As String

    out = Replace(text, "\", "\\")
    out = Replace(out, """", "\""")
    out = Replace(out, vbCr, " ")
    out = Replace(out, vbLf, " ")

    CrownJsText = """" & out & """"
End Function
