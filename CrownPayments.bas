Attribute VB_Name = "CrownPayments"
Option Explicit

' =====================================================================
' Crown Superior - payment information, both directions
' ---------------------------------------------------------------------
' Needs the CrownAPI module to be imported as well; the key is shared.
'
' Two things live here.
'
'   CrownUploadPaymentsDue
'       Takes the UploadPaymentdue sheet exactly as it is today and writes
'       every row straight onto the website. No browser, no logging in,
'       no searching the policy list, no form validation to argue with.
'       Works for every carrier, not just one.
'
'   CrownUaigPaymentDue
'       Works through OUR United Auto policies - the list comes from the
'       website, not from whichever UAIG report happens to be in front of
'       it - looks each policy up in Policy Inquiry, reads what is owed,
'       and writes it to the CrownPayments sheet AND back to the website.
'
' The website is the record. The Google Sheet reads the website on its
' own about once an hour, so anything written here reaches the phone
' system without anybody exporting or uploading anything.
' =====================================================================

' The columns on UploadPaymentdue, as it already is.
' A term cancelled or expired longer ago than this is of no use to anyone.
Private Const UAIG_STALE_DAYS As Long = 60

' Which copy of this file is in the workbook. Run CrownVersion to see it.
Public Const CROWN_PAYMENTS_VERSION As String = "6 Oct c - presses OK on Verve's Session Expiring box"

' The boxes and labels on Verve's own pages, named once so a change on
' their side is one line here rather than a hunt through the code.
Private Const VERVE_LOOKUP_BOX As String = "QuickPolicyLookupControl_NumberInsCombo_I"
Private Const VERVE_LOOKUP_GO As String = "QuickPolicyLookupControl_FireLookupImageButton"
Private Const VERVE_BILLING_TAB As String = "//*[@id='TabBar']/ul/li[5]/div/a"
Private Const VERVE_ACTIVITY As String = "P_L_v212w3_t4_ActivityStringInsLabel"
Private Const VERVE_DUE_DATE As String = "P_L_v212w3_t4_DueDateInsLabel"
Private Const VERVE_BALANCE As String = "P_L_v212w3_t4_CurrentBalanceInsLabel"

' The Amount Due report, which is the only place Verve says what a policy's
' status is. The way to it is his, from ScrapeVervePaymentdue.
Private Const VERVE_TOOLS_MENU As String = "AgencyToolsMenu"
Private Const VERVE_REPORTS_LINK As String = "/html/body/form/div[5]/div[2]/div/table/tbody/tr/td[1]/div[1]/div[5]/div/table[2]/tbody/tr/td[2]/a"
Private Const VERVE_AMOUNT_DUE As String = "P_L_OnlineReportsListForm_OnlineReportsdxGridView_DXCBtn13Img"
Private Const VERVE_RPT_ACTIVE As String = "ctl00_P_L_ReportDisplayForm_ReportViewer1_ctl04_ctl03_ddValue"
Private Const VERVE_RPT_START As String = "ctl00_P_L_ReportDisplayForm_ReportViewer1_ctl04_ctl05_txtValue"
Private Const VERVE_RPT_COMPANY As String = "ctl00_P_L_ReportDisplayForm_ReportViewer1_ctl04_ctl09_ddValue"
Private Const VERVE_RPT_DETAIL As String = "ctl00_P_L_ReportDisplayForm_ReportViewer1_ctl04_ctl19_rbTrue"
Private Const VERVE_RPT_RUN As String = "ctl00_P_L_ReportDisplayForm_ReportViewer1_ctl04_ctl00"
Private Const VERVE_RPT_EXPORT As String = "ctl00_P_L_ReportDisplayForm_ReportViewer1_ctl05_ctl04_ctl00_ButtonImgDown"

' How long a policy has to have been cancelled before we stop asking about
' it. His words: "Policies that have been cancelled for 60 days do not have
' to be accounted for."
Private Const CANCELLED_LONG_ENOUGH As Long = 60

' How many policies in a row may fail to open before something is done
' about it. Three is comfortably more than a run of genuinely unknown
' policy numbers; eight means the site is not answering at all.
Private Const CONSECUTIVE_BEFORE_RELOAD As Long = 3
Private Const CONSECUTIVE_BEFORE_GIVING_UP As Long = 8

' How many times to look for the Future table before believing it is not
' there, and how many pages to write down in full when it is not. Ten
' pages is plenty to see a pattern and few enough that the sheet stays
' readable and the run stays quick.
Private Const FUTURE_LOOKS As Long = 3
Private Const PAGES_TO_WRITE_DOWN As Long = 10

Private Const COL_POLICY As Long = 1     ' A  Policy Number
Private Const COL_SITE As Long = 2       ' B  Website (the carrier)
Private Const COL_FIRST As Long = 3      ' C  First Name
Private Const COL_LAST As Long = 4       ' D  Last Name
Private Const COL_CANCEL As Long = 5     ' E  Cancel Date
Private Const COL_AMOUNT As Long = 6     ' F  Amount Due
Private Const COL_DUE As Long = 7        ' G  Due Date
Private Const COL_STATUS As Long = 8     ' H  Status


' ---------------------------------------------------------------------
' Our policies, from the website
' ---------------------------------------------------------------------

' Every policy the website holds, as a lookup: the digits of the policy
' number -> the record number to write back to.
'
' Digits only, because the same policy is written "GAI 212619",
' "GAI212633" and "GAI - 210184" in different records, and the carriers
' ask for it as a number anyway.
'
' carrier is optional. "United Auto", "UAIG" and "United" all mean the
' same insurer to the website.
Public Function CrownPolicyRows(Optional ByVal carrier As String = "") As Collection
    Dim csv As String, body As String, lines() As String, parts() As String
    Dim rows As New Collection
    Dim i As Long

    ' Built with a plain If, not IIf: IIf works out BOTH of its answers
    ' before choosing one, so the carrier would be encoded even when there
    ' is no carrier to encode.
    body = CrownBody("policylist")
    If Len(carrier) > 0 Then body = body & "&carrier=" & CrownEnc(carrier)

    csv = CrownPost(body)

    If Len(csv) = 0 Then
        MsgBox "The website did not answer. Check your internet, then run CrownTestConnection.", _
               vbExclamation, "Crown Superior"
        Set CrownPolicyRows = rows
        Exit Function
    End If

    If Left$(csv, 5) = "error" Then
        MsgBox "The website answered: " & csv, vbExclamation, "Crown Superior"
        Set CrownPolicyRows = rows
        Exit Function
    End If

    lines = Split(Replace(csv, vbCrLf, vbLf), vbLf)

    For i = 1 To UBound(lines)                    ' 0 is the header
        If Len(Trim$(lines(i))) > 0 Then
            parts = CrownSplitCsvLine(lines(i))

            ' 0 record number, 2 policy number, 3 carrier, 4 first, 5 last
            If UBound(parts) >= 5 Then rows.Add parts
        End If
    Next i

    Set CrownPolicyRows = rows
End Function

' Which copies of the two modules are in this workbook.
'
' File > Import does NOT replace a module of the same name, so a file can
' be imported and the old one still be the one that runs. Two days went on
' a fault that had been fixed a week earlier because neither of us could
' answer this question without a screenshot of the code.
'
' CrownAPI is asked for by name rather than called directly, so that an old
' CrownAPI reports itself as old instead of stopping the whole project from
' compiling.
Public Sub CrownVersion()
    Dim apiStamp As String

    On Error Resume Next
    apiStamp = CStr(Application.Run("CrownApiStamp"))
    On Error GoTo 0

    If Len(Trim$(apiStamp)) = 0 Then
        apiStamp = "an OLD copy." & vbCrLf & vbCrLf & _
                   "Right-click CrownAPI in the list on the left, Remove CrownAPI, " & _
                   "say No when it offers to export it, then File, Import File and " & _
                   "pick the new one. Importing on top of it does not replace it."
    End If

    MsgBox "CrownPayments is " & CROWN_PAYMENTS_VERSION & vbCrLf & vbCrLf & _
           "CrownAPI is " & apiStamp, vbInformation, "Crown Superior"
End Sub

' The policies, by the digits in their policy number.
Public Function CrownPolicyIndex(Optional ByVal carrier As String = "") As Object
    Dim rows As Collection, parts As Variant
    Dim index As Object
    Dim digits As String

    Set index = CreateObject("Scripting.Dictionary")
    Set rows = CrownPolicyRows(carrier)

    For Each parts In rows
        digits = CrownDigits(parts(2))

        ' The list comes back newest first, so the first one seen for a
        ' number is the newest record with it.
        If Len(digits) > 0 And Not index.Exists(digits) Then
            index.Add digits, parts(0)
        End If
    Next parts

    Set CrownPolicyIndex = index
End Function

' The policy number in a cell, as it was really typed.
'
' A long one like 100332072503 is held by Excel as a number and comes
' back from CStr as "1.00332072503E+11", which has no policy number in
' it at all. Formatting it as a plain integer gets the digits back.
' Takes the cell itself, not its contents: passing a Range into a Variant
' hands over the value, and a value has no .Value to read.
Public Function CrownCellText(ByVal cell As Range) As String
    Dim value As Variant

    value = cell.value

    If IsEmpty(value) Then Exit Function

    If IsNumeric(value) And Not IsDate(value) Then
        If value = Int(value) And Abs(value) < 1E+15 Then
            CrownCellText = Format$(value, "0")
            Exit Function
        End If
    End If

    CrownCellText = Trim$(CStr(value))
End Function

' The digits in a policy number, and nothing else.
Public Function CrownDigits(ByVal text As String) As String
    Dim i As Long, ch As String, out As String

    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If ch >= "0" And ch <= "9" Then out = out & ch
    Next i

    CrownDigits = out
End Function

' The website wants MM-DD-YYYY. Anything it cannot read is sent through
' untouched rather than turned into today by accident.
Private Function CrownDate(ByVal value As Variant) As String
    If Trim$(CStr(value)) = "" Then Exit Function

    If IsDate(value) Then
        CrownDate = Format$(CDate(value), "mm-dd-yyyy")
    Else
        CrownDate = Trim$(CStr(value))
    End If
End Function


' ---------------------------------------------------------------------
' 1. The upload sheet, straight onto the website
' ---------------------------------------------------------------------

Public Sub CrownUploadPaymentsDue()
    Dim wsUpload As Worksheet, wsLog As Worksheet
    Dim rows As Collection, parts As Variant
    Dim byNumber As Object, byName As Object
    Dim lastRow As Long, i As Long, logRow As Long
    Dim policyText As String, digits As String, tail As String
    Dim recordId As String, matchedOn As String, note As String
    Dim carrier As String, nameKey As String, answer As String
    Dim done As Long, missing As Long, refused As Long
    Dim names As Variant, values As Variant
    Dim candidates As Variant, j As Long

    On Error Resume Next
    Set wsUpload = ThisWorkbook.Worksheets("UploadPaymentdue")
    On Error GoTo 0

    If wsUpload Is Nothing Then
        MsgBox "There is no UploadPaymentdue sheet in this workbook.", vbExclamation, "Crown Superior"
        Exit Sub
    End If

    Set rows = CrownPolicyRows()
    If rows.Count = 0 Then Exit Sub

    ' Two ways of finding a record: by the digits in the policy number, and
    ' by the customer's name. The numbers on this sheet are the carrier's,
    ' and the carrier issues a new one on a rewrite, so a policy we hold can
    ' easily be under a different number here.
    Set byNumber = CreateObject("Scripting.Dictionary")
    Set byName = CreateObject("Scripting.Dictionary")

    For Each parts In rows
        digits = CrownDigits(parts(2))
        If Len(digits) > 0 And Not byNumber.Exists(digits) Then byNumber.Add digits, parts

        nameKey = UCase$(Trim$(parts(4)) & "|" & Trim$(parts(5)))

        If Len(Replace(nameKey, "|", "")) > 0 Then
            If byName.Exists(nameKey) Then
                candidates = byName(nameKey)
                ReDim Preserve candidates(UBound(candidates) + 1)
                candidates(UBound(candidates)) = parts
                byName(nameKey) = candidates
            Else
                byName.Add nameKey, Array(parts)
            End If
        End If
    Next parts

    Set wsLog = CrownSheet("CrownUploadLog")
    wsLog.Cells.ClearContents
    wsLog.Columns(1).NumberFormat = "@"          ' policy numbers are text, not sums
    wsLog.Range("A1:K1").value = Array("Policy Number", "Website", "First Name", "Last Name", _
                                       "Cancel Date", "Amount Due", "Due Date", "Status", _
                                       "When", "Matched on", "What happened")
    logRow = 2

    lastRow = wsUpload.Cells(wsUpload.rows.Count, COL_POLICY).End(xlUp).row

    For i = 2 To lastRow
        policyText = CrownCellText(wsUpload.Cells(i, COL_POLICY))

        If Len(Trim$(policyText)) > 0 Then
            digits = CrownDigits(policyText)
            carrier = Trim$(CStr(wsUpload.Cells(i, COL_SITE).value))
            nameKey = UCase$(Trim$(CStr(wsUpload.Cells(i, COL_FIRST).value)) & "|" & _
                             Trim$(CStr(wsUpload.Cells(i, COL_LAST).value)))
            recordId = ""
            matchedOn = ""
            note = ""

            ' 1. the whole policy number
            If byNumber.Exists(digits) Then
                parts = byNumber(digits)
                recordId = parts(0)
                matchedOn = "policy number"
            End If

            ' 2. the part after the last hyphen, which is what the old macro
            '    searched on for United Auto
            If Len(recordId) = 0 And InStr(policyText, "-") > 0 Then
                tail = CrownDigits(Mid$(policyText, InStrRev(policyText, "-") + 1))

                If Len(tail) > 0 And byNumber.Exists(tail) Then
                    parts = byNumber(tail)
                    recordId = parts(0)
                    matchedOn = "end of the policy number"
                End If
            End If

            ' 3. the customer's name, and only when it points at one policy
            If Len(recordId) = 0 And byName.Exists(nameKey) Then
                candidates = byName(nameKey)
                note = ""

                For j = LBound(candidates) To UBound(candidates)
                    parts = candidates(j)

                    If Len(carrier) = 0 Or CrownSameCarrier(carrier, CStr(parts(3))) Then
                        If Len(recordId) = 0 Then
                            recordId = parts(0)
                            matchedOn = "name"
                            If Len(carrier) > 0 Then matchedOn = "name and carrier"
                        Else
                            ' more than one - too risky to pick
                            recordId = ""
                            matchedOn = ""
                            note = "That name has more than one policy here: "
                            Exit For
                        End If
                    End If
                Next j

                If Len(note) > 0 Then
                    For j = LBound(candidates) To UBound(candidates)
                        parts = candidates(j)
                        note = note & parts(0) & " " & parts(2) & " (" & parts(3) & ")  "
                    Next j
                End If
            End If

            wsLog.Cells(logRow, 1).value = policyText
            wsLog.Cells(logRow, 2).value = wsUpload.Cells(i, COL_SITE).value
            wsLog.Cells(logRow, 3).value = wsUpload.Cells(i, COL_FIRST).value
            wsLog.Cells(logRow, 4).value = wsUpload.Cells(i, COL_LAST).value
            wsLog.Cells(logRow, 5).value = wsUpload.Cells(i, COL_CANCEL).value
            wsLog.Cells(logRow, 6).value = wsUpload.Cells(i, COL_AMOUNT).value
            wsLog.Cells(logRow, 7).value = wsUpload.Cells(i, COL_DUE).value
            wsLog.Cells(logRow, 8).value = wsUpload.Cells(i, COL_STATUS).value
            wsLog.Cells(logRow, 9).value = Now
            wsLog.Cells(logRow, 10).value = matchedOn

            If Len(recordId) = 0 Then
                If Len(note) = 0 Then note = "No policy on the website with that number or that name"
                wsLog.Cells(logRow, 11).value = note
                missing = missing + 1
            Else
                names = Array("Web_pymnt_due", "updated_due_date", "updated_cancel_date", _
                              "Status_", "Update_due_dates", "Last_Updated")
                values = Array(Trim$(CStr(wsUpload.Cells(i, COL_AMOUNT).value)), _
                               CrownDate(wsUpload.Cells(i, COL_DUE).value), _
                               CrownDate(wsUpload.Cells(i, COL_CANCEL).value), _
                               Trim$(CStr(wsUpload.Cells(i, COL_STATUS).value)), _
                               "Yes", _
                               Format$(Now, "mm-dd-yyyy hh:mm AM/PM"))

                answer = CrownUpdate(CROWN_FORM_POLICY, CLng(recordId), names, values)

                If InStr(1, answer, """ok"":true", vbTextCompare) > 0 Then
                    wsLog.Cells(logRow, 11).value = "Written to record " & recordId
                    done = done + 1
                ElseIf Len(Trim$(answer)) = 0 Then
                    wsLog.Cells(logRow, 11).value = "The website did not answer - try this row again"
                    refused = refused + 1
                Else
                    wsLog.Cells(logRow, 11).value = "Refused: " & answer
                    refused = refused + 1
                End If
            End If

            logRow = logRow + 1
        End If
    Next i

    wsLog.Columns("A:K").AutoFit
    wsLog.Activate

    MsgBox done & " policy record(s) updated on the website." & vbCrLf & _
           missing & " could not be matched to a policy there." & vbCrLf & _
           refused & " were refused or did not go through." & vbCrLf & vbCrLf & _
           "The CrownUploadLog sheet says what happened to every row, and how each " & _
           "one was matched." & vbCrLf & vbCrLf & _
           "The Google Sheet picks the changes up within the hour on its own.", _
           vbInformation, "Crown Superior"
End Sub

' Two ways of writing the same insurer. Loose on purpose: the sheet says
' "United Au" or "UAIG" where the website says "United Auto".
Public Function CrownSameCarrier(ByVal a As String, ByVal b As String) As Boolean
    Dim x As String, y As String

    x = UCase$(Trim$(a))
    y = UCase$(Trim$(b))

    If Len(x) = 0 Or Len(y) = 0 Then
        CrownSameCarrier = True
        Exit Function
    End If

    If x = y Then
        CrownSameCarrier = True
        Exit Function
    End If

    If InStr(1, y, x, vbTextCompare) = 1 Or InStr(1, x, y, vbTextCompare) = 1 Then
        CrownSameCarrier = True
        Exit Function
    End If

    If (InStr(1, x, "UNITED", vbTextCompare) > 0 Or InStr(1, x, "UAIG", vbTextCompare) > 0) _
       And InStr(1, y, "UNITED", vbTextCompare) > 0 Then
        CrownSameCarrier = True
        Exit Function
    End If

    If (InStr(1, x, "VERVE", vbTextCompare) > 0 Or InStr(1, x, "TRISURA", vbTextCompare) > 0) _
       And (InStr(1, y, "VERVE", vbTextCompare) > 0 Or InStr(1, y, "TRISURA", vbTextCompare) > 0) Then
        CrownSameCarrier = True
        Exit Function
    End If

    CrownSameCarrier = False
End Function


' ---------------------------------------------------------------------
' 2. United Auto, policy by policy, from OUR list
' ---------------------------------------------------------------------

' ---------------------------------------------------------------------
' 2. United Auto, one policy at a time - renewals included
' ---------------------------------------------------------------------
'
' Policy Inquiry does not always go straight to a policy. When a policy
' has been renewed, United holds more than one term under it and answers
' with a list instead - "Policy Inquiry Result", a line per term, each
' with its own number, its own effective date and its own status. The
' newer term is given a new number: usually the old one with 10 in front
' of it, sometimes a number with nothing in common with it at all.
'
' That list is why four policies came back empty last time. This reads
' it, takes the term with the newest effective date - ignoring anything
' cancelled or expired more than two months ago - and opens that one.
'
' Where the newer number is not on our website at all, the policy has
' been renewed without us knowing. Those are gathered up and offered at
' the end as new policy records: a copy of the one we hold, with the new
' number, the new dates and "Renewal" as the description. Nothing is
' created without being listed and agreed to first.

' A date the way United writes it, as something that can be compared.
'
' It writes them two ways on the same screen depending on where you are:
' 02/23/2026 on one, 2026-02-23 on another. Four digits at the front is
' the year, so month and day follow it; anything else is month first, the
' American way. Zero for anything that is not a date at all, so a blank
' or an "N/A" sorts below every real date instead of above them.
Public Function CrownUsDate(ByVal text As String) As Double
    Dim parts() As String
    Dim d As Long, m As Long, y As Long

    text = Trim$(Replace(Replace(text, "-", "/"), ".", "/"))
    If Len(text) = 0 Then Exit Function

    parts = Split(text, "/")
    If UBound(parts) <> 2 Then Exit Function

    If Len(Trim$(parts(0))) = 4 Then
        y = Val(parts(0))
        m = Val(parts(1))
        d = Val(parts(2))
    Else
        m = Val(parts(0))
        d = Val(parts(1))
        y = Val(parts(2))
    End If

    If y < 100 Then y = 2000 + y
    If m < 1 Or m > 12 Then Exit Function
    If d < 1 Or d > 31 Then Exit Function
    If y < 1900 Or y > 2200 Then Exit Function

    On Error Resume Next
    CrownUsDate = CDbl(DateSerial(y, m, d))
End Function

' Which column holds what. Found by its heading rather than by counting,
' because the order of them is the carrier's business, not ours.
'
' notThis keeps "Policy Status" from answering to a search for "polic".
' A heading that has both words in it is tried only when nothing else
' matches, so a report with a Policy Status column and no policy number
' column still finds something rather than nothing.
Private Function CrownColumnOf(ByRef header() As String, ByVal wanted As String, _
                               Optional ByVal notThis As String = "") As Long
    Dim i As Long
    Dim fallback As Long

    CrownColumnOf = -1
    fallback = -1

    For i = LBound(header) To UBound(header)
        If InStr(1, header(i), wanted, vbTextCompare) > 0 Then
            If Len(notThis) = 0 Or InStr(1, header(i), notThis, vbTextCompare) = 0 Then
                CrownColumnOf = i
                Exit Function
            End If

            If fallback < 0 Then fallback = i
        End If
    Next i

    CrownColumnOf = fallback
End Function

' The result table as plain text: rows separated by ~, cells by |.
'
' The page is several tables deep inside one another and every one of
' them contains the words we are looking for, so the smallest one that
' does is the real table.
Private Function CrownUaigListScript() As String
    CrownUaigListScript = _
        "var best=null,bl=0,tabs=document.getElementsByTagName('table');" & _
        "for(var i=0;i<tabs.length;i++){var t=tabs[i],x=t.innerText||'';" & _
        "if(x.indexOf('Term Effective Date')<0)continue;" & _
        "if(best===null||x.length<bl){best=t;bl=x.length;}}" & _
        "if(!best)return '';var out=[];" & _
        "for(var r=0;r<best.rows.length;r++){var cs=best.rows[r].cells,line=[];" & _
        "for(var c=0;c<cs.length;c++){line.push((cs[c].innerText||'')" & _
        ".replace(/[|~]/g,' ').replace(/\s+/g,' ').replace(/^ /,'').replace(/ $/,''));}" & _
        "out.push(line.join('|'));}return out.join('~');"
End Function

' Open one line of that list, by clicking its policy number where it sits
' rather than searching again - searching again only brings the list back.
Private Function CrownUaigClickScript(ByVal rowIndex As Long) As String
    CrownUaigClickScript = _
        "var best=null,bl=0,tabs=document.getElementsByTagName('table');" & _
        "for(var i=0;i<tabs.length;i++){var t=tabs[i],x=t.innerText||'';" & _
        "if(x.indexOf('Term Effective Date')<0)continue;" & _
        "if(best===null||x.length<bl){best=t;bl=x.length;}}" & _
        "if(!best)return 'no table';var row=best.rows[" & rowIndex & "];" & _
        "if(!row)return 'no row';var a=row.getElementsByTagName('a')[0];" & _
        "if(!a)return 'no link';a.click();return 'clicked';"
End Function

' Are we looking at one policy, or still at the list?
Private Function CrownUaigOnPolicy(ByVal drv As ChromeDriver) As Boolean
    Dim answer As String

    On Error Resume Next
    answer = drv.ExecuteScript( _
        "var t=document.body?(document.body.innerText||''):'';" & _
        "return (document.getElementById('CURAMTDUE')" & _
        "||t.indexOf('Policy Status')>=0)?'1':'0';")
    On Error GoTo 0

    CrownUaigOnPolicy = (answer = "1")
End Function

' The line of the list to act on.
'
' "policies with a newer effective date in almost every case will be the
' policy to consider" - so the newest effective date wins, and anything
' cancelled or expired more than two months ago is passed over.
'
' Answers the row number to click, or 0 for nothing usable.
Private Function CrownUaigPickRow(ByVal listText As String, _
                                  ByRef bestNumber As String, ByRef bestEff As String, _
                                  ByRef bestExp As String, ByRef bestStatus As String, _
                                  ByRef note As String) As Long
    Dim rows() As String, header() As String, cells() As String
    Dim colPolicy As Long, colEff As Long, colExp As Long, colStatus As Long
    Dim r As Long, best As Long, skipped As Long
    Dim eff As Double, expiry As Double, bestEffOn As Double
    Dim status As String

    bestNumber = ""
    bestEff = ""
    bestExp = ""
    bestStatus = ""

    rows = Split(listText, "~")
    If UBound(rows) < 1 Then Exit Function          ' a heading and nothing under it

    header = Split(rows(0), "|")
    colPolicy = CrownColumnOf(header, "Policy No")
    colEff = CrownColumnOf(header, "Term Effective")
    colExp = CrownColumnOf(header, "Term Expiration")
    colStatus = CrownColumnOf(header, "Status")

    If colPolicy < 0 Or colEff < 0 Then
        note = "the list did not have the columns expected"
        Exit Function
    End If

    For r = 1 To UBound(rows)
        cells = Split(rows(r), "|")

        If UBound(cells) >= colPolicy And UBound(cells) >= colEff Then
            status = ""
            If colStatus >= 0 And UBound(cells) >= colStatus Then status = cells(colStatus)

            eff = CrownUsDate(cells(colEff))
            expiry = 0
            If colExp >= 0 And UBound(cells) >= colExp Then expiry = CrownUsDate(cells(colExp))

            ' Long gone. Skipped whatever its dates say about being newest.
            If (InStr(1, status, "cancel", vbTextCompare) > 0 _
                Or InStr(1, status, "expire", vbTextCompare) > 0) _
               And expiry > 0 And expiry < CDbl(Date) - UAIG_STALE_DAYS Then
                skipped = skipped + 1
            ElseIf Len(Trim$(cells(colPolicy))) > 0 And eff >= bestEffOn Then
                bestEffOn = eff
                best = r
                bestNumber = Trim$(cells(colPolicy))
                bestEff = Trim$(cells(colEff))
                bestStatus = Trim$(status)
                If colExp >= 0 And UBound(cells) >= colExp Then bestExp = Trim$(cells(colExp))
            End If
        End If
    Next r

    If best > 0 Then
        note = UBound(rows) & " terms on United; took " & bestNumber & " effective " & bestEff
        If UBound(rows) = 1 Then note = "one term on United; took " & bestNumber & " effective " & bestEff
        If skipped > 0 Then note = note & "; " & skipped & " long expired"
    ElseIf skipped > 0 Then
        note = "all " & skipped & " terms cancelled or expired more than two months ago"
    End If

    CrownUaigPickRow = best
End Function


' A policy we can stop asking about.
'
'   "Policies that have been cancelled for 60 days do not have to be
'    accounted for. maybe it could mean less work."
'
' Only where the website can prove it: a status that says cancelled AND a
' cancel date more than sixty days behind us. A cancellation with no date
' on it is still looked up, because nothing here knows how old it is - and
' a policy skipped wrongly is one nobody ever looks at again.
Private Function CrownLongCancelled(ByRef parts As Variant, ByRef since As String) As Boolean
    Dim words As String
    Dim cancelledOn As Double

    since = ""

    If UBound(parts) < 10 Then Exit Function

    words = CStr(parts(6)) & " " & CStr(parts(7))

    If InStr(1, words, "cancel", vbTextCompare) = 0 _
       And InStr(1, words, "laps", vbTextCompare) = 0 Then
        Exit Function
    End If

    cancelledOn = CrownUsDate(CStr(parts(10)))

    If cancelledOn = 0 Then Exit Function
    If cancelledOn > CDbl(Date) - CANCELLED_LONG_ENOUGH Then Exit Function

    since = Trim$(CStr(parts(10)))
    CrownLongCancelled = True
End Function


' The "Verify Phone number" dialog, out of the way.
'
' United puts it over the policy page. Everything underneath it is still
' in the page, so nothing errors - the figures simply come back as they
' were on the last policy, or as nothing at all.
'
' It only ever clicks the choice that says "Do not", never the other one.
' If that wording is not there it closes the dialog instead and says so.
' Nothing here is ever going to change a customer's telephone number
' because a radio button moved.
Private Function CrownDismissPhoneCheck(ByVal drv As ChromeDriver) As String
    Dim answer As String

    On Error Resume Next
    answer = drv.ExecuteScript( _
        "var rs=document.querySelectorAll('input[type=radio]'),pick=null;" & _
        "function labelOf(r){var t='';" & _
        "if(r.id){var l=document.querySelector('label[for=\""'+r.id+'\""]');if(l)t+=' '+(l.innerText||'');}" & _
        "var n=r.nextSibling,hops=0;" & _
        "while(n&&hops<6){if(n.nodeType===1&&n.tagName==='INPUT')break;" & _
        "t+=' '+(n.textContent||'');n=n.nextSibling;hops++;}return t;}" & _
        "for(var i=0;i<rs.length;i++){" & _
        "if(/do\s*not\s*change\s*phone/i.test(labelOf(rs[i]))){pick=rs[i];break;}}" & _
        "if(pick){pick.click();" & _
        "var f=pick.form||document," & _
        "bs=f.querySelectorAll('input[type=submit],input[type=button],button');" & _
        "for(var b=0;b<bs.length;b++){" & _
        "var v=((bs[b].value||bs[b].innerText||'')+'').replace(/^\s+|\s+$/g,'').toLowerCase();" & _
        "if(v==='submit'){bs[b].click();return 'phone check answered';}}" & _
        "return 'phone check answered, no Submit found';}" & _
        "if((document.body.innerText||'').indexOf('Verify Phone number')<0)return '';" & _
        "var xs=document.querySelectorAll('a,span,button,img,div');" & _
        "for(var x=0;x<xs.length;x++){var w=((xs[x].innerText||xs[x].title||'')+'')" & _
        ".replace(/^\s+|\s+$/g,'');" & _
        "if(w==='×'||w==='x'||w.toLowerCase()==='close'){xs[x].click();" & _
        "return 'phone check closed';}}" & _
        "return 'phone check in the way and would not close';")
    On Error GoTo 0

    CrownDismissPhoneCheck = answer
End Function

' The address of one policy on United, so it can be opened without going
' anywhere near the search box.
'
' A dialog left open on the last page can stop a click landing. It cannot
' stop a navigation - the page is replaced either way - and that is what
' keeps one stuck screen from being read as a hundred policies.
'
' The host is taken from wherever the browser already is rather than
' written in, because the agent site is per state.
Private Function CrownUaigPolicyUrl(ByVal drv As ChromeDriver, ByVal policyNo As String) As String
    Dim here As String, host As String
    Dim letters As String, digits As String
    Dim i As Long, at As Long
    Dim ch As String

    digits = CrownDigits(policyNo)
    If Len(digits) = 0 Then Exit Function

    For i = 1 To Len(policyNo)
        ch = UCase$(Mid$(policyNo, i, 1))

        If ch >= "A" And ch <= "Z" Then
            letters = letters & ch
        ElseIf ch >= "0" And ch <= "9" Then
            Exit For
        End If
    Next i

    On Error Resume Next
    here = CStr(drv.Url)
    On Error GoTo 0

    If Len(here) < 12 Then Exit Function

    at = InStr(9, here, "/")
    If at = 0 Then Exit Function

    host = Left$(here, at - 1)

    CrownUaigPolicyUrl = host & "/agents/ndmacro/cmn_Review.mac/ValidatePolicy" & _
                         "?dbxPolcyPfx=" & letters & "&tbxPolicyNo=" & digits & _
                         "&TRANSFLG=PE&strPolcntFlag=Y&strMtermFlag=Y"
End Function

' What the policy page says about itself.
'
' It carries "Previous Policy Number" and "Rewritten Policy Number" -
' United naming the term before this one and the term that replaced it -
' along with the number, dates and status of the one being looked at. All
' of that beats anything worked out from a list.
'
' Comes back as six parts separated by | :
'   number | previous | rewritten | effective | expiration | status
Private Function CrownUaigDetail(ByVal drv As ChromeDriver) As String
    Dim answer As String

    On Error Resume Next
    answer = drv.ExecuteScript( _
        "function tidy(s){return (s||'').replace(/\s+/g,' ').replace(/^ | $/g,'');}" & _
        "var cells=document.getElementsByTagName('td');" & _
        "function val(label){for(var i=0;i<cells.length;i++){" & _
        "if(tidy(cells[i].innerText)!==label)continue;" & _
        "for(var j=i+1;j<Math.min(i+3,cells.length);j++){" & _
        "var v=tidy(cells[j].innerText);if(v)return v.replace(/[|]/g,' ');}" & _
        "return '';}return '';}" & _
        "return [val('Policy Number'),val('Previous Policy Number')," & _
        "val('Rewritten Policy Number'),val('Policy Effective')," & _
        "val('Policy Expiration'),val('Policy Status')].join('|');")
    On Error GoTo 0

    CrownUaigDetail = answer
End Function

' One of those six, or nothing.
Private Function CrownPart(ByVal packed As String, ByVal which As Long) As String
    Dim parts() As String

    If Len(packed) = 0 Then Exit Function

    parts = Split(packed, "|")

    If which <= UBound(parts) Then CrownPart = Trim$(parts(which))
End Function


Public Sub CrownUaigPaymentDue()
    Dim drv As ChromeDriver, clsDrv As Chrm
    Dim By As New Selenium.By
    Dim Keys As New Selenium.Keys
    Dim wsInput As Worksheet, ws As Worksheet
    Dim policies As Collection, parts As Variant
    Dim index As Object, names As Object
    Dim renewals As Collection
    Dim policyDigits As String, recordId As String, siteNumber As String
    Dim liveNumber As String, liveDigits As String
    Dim termEff As String, termExp As String, listStatus As String
    Dim dueAmount As String, dueDate As String, cancelDate As String
    Dim policyStatus As String, detailStatus As String
    Dim listText As String, note As String, answer As String, jsScript As String
    Dim detail As String, phoneNote As String
    Dim pageDigits As String, chosenDigits As String
    Dim policyUrl As String
    Dim openedByUrl As Boolean
    Dim wrongPage As Boolean
    Dim lost As Long
    Dim row As Long, done As Long, blank As Long, pick As Long, madeCount As Long
    Dim skipped As Long
    Dim cancelledOn As String
    Dim i As Long

    Set wsInput = ThisWorkbook.Worksheets("Input")

    ' Our United Auto policies, newest first. If the website has none, or
    ' cannot be reached, stop here rather than open a browser for nothing.
    Set policies = CrownPolicyRows("United Auto")

    If policies.Count = 0 Then
        MsgBox "No United Auto policies came back from the website.", vbExclamation, "Crown Superior"
        Exit Sub
    End If

    ' The digits of every policy number we hold, so a number United gives
    ' back can be recognised as one we already have.
    Set index = CreateObject("Scripting.Dictionary")
    Set names = CreateObject("Scripting.Dictionary")

    For Each parts In policies
        policyDigits = CrownDigits(parts(2))

        If Len(policyDigits) > 0 And Not index.Exists(policyDigits) Then
            index.Add policyDigits, parts(0)
            names.Add policyDigits, Trim$(parts(4) & " " & parts(5))
        End If
    Next parts

    Set renewals = New Collection

    Set ws = CrownSheet("CrownPayments")
    ws.Cells.ClearContents
    ws.Columns(2).NumberFormat = "@"       ' a long policy number is not a sum
    ws.Columns(3).NumberFormat = "@"
    ws.Range("A1:K1").value = Array("Record number", "Policy on our site", "Policy at United", _
                                    "Amount due", "Due date", "Cancel date", "Policy status", _
                                    "Term effective", "Term expires", "When", "What happened")
    row = 2

    Set clsDrv = New Chrm
    Set clsDrv.ChrmDriver = New ChromeDriver
    Set drv = clsDrv.ChrmDriver

    On Error Resume Next
    drv.AddArgument "--force-device-scale-factor=0.70"
    drv.Start
    On Error GoTo 0

    drv.Get wsInput.Range("URL_26").value
    EnterData drv, Keys, "tbxUserID", wsInput.Range("USER_26").value, "ID"
    EnterData drv, Keys, "tbxPassword", wsInput.Range("PASS_26").value, "ID"
    drv.Keyboard.SendKeys Keys.Enter

    LoopElementUntilFoundBYXPATH drv, "//a[normalize-space(text())='Quote']"
    LoopElementUntilFound drv, "rpthref"

    ' Straight to Policy Inquiry. No reports, no downloads, no tabs - the
    ' list of what to look up came from us, so nothing here depends on
    ' which report United happens to show first.
    ClickElement drv, "//a[normalize-space(text())='Work with Policies']", "XPATH"
    ClickElement drv, "//a[normalize-space(text())='Policy Inquiry']", "XPATH"

    For i = 1 To policies.Count
        parts = policies(i)
        recordId = parts(0)
        siteNumber = parts(2)
        policyDigits = CrownDigits(siteNumber)

        If Len(policyDigits) = 0 Then GoTo NextPolicy

        ' Cancelled long enough ago that nothing about it can change.
        If CrownLongCancelled(parts, cancelledOn) Then
            ws.Cells(row, 1).value = recordId
            ws.Cells(row, 2).value = siteNumber
            ws.Cells(row, 11).value = "skipped - cancelled since " & cancelledOn
            skipped = skipped + 1
            row = row + 1
            GoTo NextPolicy
        End If

        dueAmount = ""
        dueDate = ""
        cancelDate = ""
        policyStatus = ""
        detailStatus = ""
        liveNumber = ""
        termEff = ""
        termExp = ""
        listStatus = ""
        note = ""
        phoneNote = ""
        chosenDigits = ""

        On Error Resume Next

        ' Straight to the policy's own address first. Nothing on the last
        ' page can get in the way of a navigation.
        openedByUrl = False
        policyUrl = CrownUaigPolicyUrl(drv, siteNumber)

        If Len(policyUrl) > 0 Then
            drv.Get policyUrl
            drv.Wait 1500
            phoneNote = CrownDismissPhoneCheck(drv)

            If CrownUaigOnPolicy(drv) Then openedByUrl = True
        End If

        ' The search box is the fallback, and still the way a renewed policy
        ' is found - it is the search that answers with the list of terms.
        If openedByUrl Then
            lost = 0
            GoTo ReadThePolicy
        End If

        ' Bounded, and the answer is acted on. LoopElementUntilFound gives
        ' up quietly after fifty tries, and everything that followed was
        ' inside an On Error - so a search box that never came back looked
        ' exactly like a search that worked.
        If Not CrownWaitFor(drv, By, "ID", "tbxPolicyNo", 20) Then
            lost = lost + 1
            ws.Cells(row, 1).value = recordId
            ws.Cells(row, 2).value = siteNumber
            ws.Cells(row, 11).value = "the policy search box never came back - nothing written"
            blank = blank + 1
            row = row + 1

            ' Three in a row means we are stuck on something, and another
            ' hundred of these would just be the same screen read again.
            If lost >= 3 Then
                MsgBox "United stopped answering after " & (i - 1) & " policies." & vbCrLf & vbCrLf & _
                       "Nothing has been written from the screen it is stuck on. " & _
                       "Close the browser, look at what is on it, and run this again.", _
                       vbExclamation, "Crown Superior"
                Exit For
            End If

            GoTo NextPolicy
        End If

        lost = 0
        drv.FindElementById("tbxPolicyNo").Clear
        drv.FindElementById("tbxPolicyNo").SendKeys policyDigits
        ClickElement drv, "btnSubmitPol1", "ID"
        drv.Wait 2000

        ' Anything United has put over the page, out of the way first.
        ' Held separately: the list summary is written into note below, and
        ' would otherwise throw this away.
        phoneNote = CrownDismissPhoneCheck(drv)

        ' A renewed policy answers with its terms instead of going
        ' straight to one of them.
        '
        ' Asked in this order on purpose: only look for a list when we are
        ' not already on a policy. A policy page that happened to carry the
        ' words "Term Effective Date" anywhere on it would otherwise be
        ' read as a list, and clicked, for all hundred and thirty-nine of
        ' them rather than the four that need it.
        listText = ""

        If Not CrownUaigOnPolicy(drv) Then
            listText = drv.ExecuteScript(CrownUaigListScript())

            If Len(listText) > 0 Then
                pick = CrownUaigPickRow(listText, liveNumber, termEff, termExp, listStatus, note)

                If pick > 0 Then
                    chosenDigits = CrownDigits(liveNumber)
                    drv.ExecuteScript CrownUaigClickScript(pick)
                    drv.Wait 2500
                End If
            End If
        End If

        ' Only read a policy if we are actually looking at one. Reading
        ' the list as though it were a policy is how empty answers get
        ' written over good ones.
ReadThePolicy:
        wrongPage = False

        If CrownUaigOnPolicy(drv) Then
            phoneNote = Trim$(phoneNote & " " & CrownDismissPhoneCheck(drv))

            ' What the page says about itself beats anything the list said.
            detail = CrownUaigDetail(drv)
            pageDigits = CrownDigits(CrownPart(detail, 0))

            ' And the page has to BE the policy we went looking for.
            '
            ' A renewed policy legitimately opens under its other number, so
            ' the term picked out of the list counts as a match too. Anything
            ' else is another customer's screen, and nothing on it belongs on
            ' this record.
            If Len(pageDigits) = 0 Then
                wrongPage = True
                note = Trim$(note & " the page did not say which policy it was - nothing written")
            ElseIf pageDigits <> policyDigits And pageDigits <> chosenDigits Then
                wrongPage = True
                note = Trim$(note & " United was still showing " & CrownPart(detail, 0) & _
                             " - nothing written")
            End If

            If Len(CrownPart(detail, 0)) > 0 Then liveNumber = CrownPart(detail, 0)
            If Len(CrownPart(detail, 3)) > 0 Then termEff = CrownPart(detail, 3)
            If Len(CrownPart(detail, 4)) > 0 Then termExp = CrownPart(detail, 4)

            ' United saying, in as many words, that this term was replaced.
            If Len(CrownPart(detail, 2)) > 0 Then
                note = Trim$(note & " rewritten as " & CrownPart(detail, 2) & ".")
            End If

            If drv.IsElementPresent(By.ID("CURAMTDUE")) Then
                dueAmount = drv.FindElementById("CURAMTDUE").value
            End If

            detailStatus = drv.FindElementByXPath( _
                "//td[.//font[contains(.,'Policy Status')]]/following-sibling::td//strong").text

            If Trim$(detailStatus) = "" Then
                detailStatus = drv.FindElementByXPath( _
                    "//font[contains(text(),'Policy Status')]/ancestor::td/following-sibling::td//font").text
            End If

            detailStatus = Replace(detailStatus, vbLf, " ")
            detailStatus = Application.WorksheetFunction.Trim(detailStatus)

            ' The last starred instalment line is the one that is due.
            jsScript = "var rows = document.evaluate(" & _
                       """" & "//tr[td//font[contains(.,'Installment') and contains(.,'*')]]" & """" & _
                       ", document, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null);" & _
                       "if (rows.snapshotLength > 0) {" & _
                       "  var lastRow = rows.snapshotItem(rows.snapshotLength - 1);" & _
                       "  var cells = lastRow.getElementsByTagName('td');" & _
                       "  if (cells.length >= 7) { return cells[6].innerText.trim(); }" & _
                       "} return '';"
            dueDate = drv.ExecuteScript(jsScript)

            cancelDate = drv.ExecuteScript( _
                "var c = document.querySelector('#spnShowSchedule span.pay-dt');" & _
                "return c ? c.innerText.trim() : '';")
        ElseIf Len(listText) > 0 Then
            If Len(note) = 0 Then note = "United listed its terms but would not open one"
        End If

        On Error GoTo 0

        If wrongPage Then
            dueAmount = ""
            dueDate = ""
            cancelDate = ""
            detailStatus = ""
            listStatus = ""
            liveNumber = ""
            termEff = ""
            termExp = ""
        End If

        note = Trim$(note & " " & phoneNote)

        ' The list's own wording, unless the policy page gave a better one.
        policyStatus = detailStatus
        If Len(Trim$(policyStatus)) = 0 Then policyStatus = listStatus

        If Len(liveNumber) = 0 Then liveNumber = siteNumber
        liveDigits = CrownDigits(liveNumber)

        ws.Cells(row, 1).value = recordId
        ws.Cells(row, 2).value = siteNumber
        ws.Cells(row, 3).value = liveNumber
        ws.Cells(row, 4).value = dueAmount
        ws.Cells(row, 5).value = dueDate
        ws.Cells(row, 6).value = cancelDate
        ws.Cells(row, 7).value = policyStatus
        ws.Cells(row, 8).value = termEff
        ws.Cells(row, 9).value = termExp
        ws.Cells(row, 10).value = Now

        ' Nothing at all came back: the policy is not at this carrier, or
        ' the page did not load. Do not write emptiness over what the
        ' website already holds.
        If Len(Trim$(dueAmount)) = 0 And Len(Trim$(dueDate)) = 0 _
           And Len(Trim$(cancelDate)) = 0 And Len(Trim$(policyStatus)) = 0 Then
            ws.Cells(row, 11).value = Trim$("nothing found - left alone. " & note)
            blank = blank + 1
        Else
            answer = CrownUpdate(CROWN_FORM_POLICY, CLng(recordId), _
                Array("Web_pymnt_due", "updated_due_date", "updated_cancel_date", _
                      "Status_", "Update_due_dates", "Last_Updated", "Current_policy_no"), _
                Array(dueAmount, CrownDate(dueDate), CrownDate(cancelDate), _
                      policyStatus, "Yes", Format$(Now, "mm-dd-yyyy hh:mm AM/PM"), liveNumber))

            If InStr(1, answer, """ok"":true", vbTextCompare) > 0 Then
                ws.Cells(row, 11).value = Trim$("written. " & note)
                done = done + 1
            ElseIf Len(Trim$(answer)) = 0 Then
                ws.Cells(row, 11).value = "website did not answer - run again for this one"
            Else
                ws.Cells(row, 11).value = "refused: " & answer
            End If
        End If

        ' United has moved this policy on to a number we have never seen.
        If Len(liveDigits) > 0 And liveDigits <> policyDigits Then
            If index.Exists(liveDigits) Then
                ws.Cells(row, 11).value = ws.Cells(row, 11).value & _
                    " The newer term is already on the site as record " & index(liveDigits) & "."
            Else
                renewals.Add Array(recordId, liveNumber, termEff, termExp, _
                                   dueAmount, dueDate, cancelDate, policyStatus, _
                                   names(policyDigits), siteNumber)
                index.Add liveDigits, recordId          ' so one run cannot offer it twice
                ws.Cells(row, 11).value = ws.Cells(row, 11).value & " RENEWED - new number."
            End If
        End If

        row = row + 1

NextPolicy:
        On Error Resume Next

        ' Clear the dialog BEFORE trying to navigate. It is modal - while it
        ' is up these two clicks do nothing, the search box never comes back,
        ' and the next policy gets read off this same screen. That is what
        ' put one policy's figures onto seventy records.
        CrownDismissPhoneCheck drv

        ClickElement drv, "//a[normalize-space(text())='Work with Policies']", "XPATH"
        ClickElement drv, "//a[normalize-space(text())='Policy Inquiry']", "XPATH"
        drv.Wait 800
        On Error GoTo 0
    Next i

    madeCount = CrownMakeRenewals(renewals, ws)

    ws.Activate

    MsgBox policies.Count & " United Auto policies on the website." & vbCrLf & _
           skipped & " cancelled more than " & CANCELLED_LONG_ENOUGH & " days ago and passed over." & vbCrLf & _
           done & " written back to the website." & vbCrLf & _
           blank & " had nothing to read and were left as they were." & vbCrLf & _
           renewals.Count & " had been renewed under a new number, " & madeCount & " added." & vbCrLf & vbCrLf & _
           "See the CrownPayments sheet for the detail. The Google Sheet " & _
           "picks the changes up within the hour.", vbInformation, "Crown Superior"
End Sub


' ---------------------------------------------------------------------
' 3. Renewals
' ---------------------------------------------------------------------
'
' Each one becomes a copy of the policy we already hold for that customer
' - same driver, same car, same cover - carrying the new number, the new
' term dates and "Renewal" as its description. The copy is made on the
' website, so nothing has to be typed back in.
'
' They are listed and agreed to first. A policy record is not something
' to create behind someone's back.
Private Function CrownMakeRenewals(ByVal renewals As Collection, ByVal ws As Object) As Long
    Dim renewal As Variant
    Dim listing As String, answer As String
    Dim i As Long, made As Long

    If renewals.Count = 0 Then Exit Function

    For i = 1 To renewals.Count
        renewal = renewals(i)
        listing = listing & renewal(8) & " - " & renewal(9) & " is now " & renewal(1)

        If Len(renewal(2)) > 0 Then listing = listing & ", effective " & renewal(2)

        listing = listing & vbCrLf
    Next i

    If MsgBox(renewals.Count & " of these policies have been renewed at United under a new " & _
              "number that is not on our website:" & vbCrLf & vbCrLf & listing & vbCrLf & _
              "Add each one as a new policy, copied from the one we hold, with the new " & _
              "number and dates and ""Renewal"" as the description?", _
              vbYesNo + vbQuestion, "Crown Superior") <> vbYes Then
        Exit Function
    End If

    For i = 1 To renewals.Count
        renewal = renewals(i)

        ' "March 23, 2023" is how that form writes a policy date. Sent in
        ' any other shape it goes in as text the calendar cannot read.
        answer = CrownClone(CROWN_FORM_POLICY, CLng(renewal(0)), _
            Array("Policy number", "Current_policy_no", "Description", _
                  "today_date_amin", "exp_date_amin", _
                  "Web_pymnt_due", "updated_due_date", "updated_cancel_date", _
                  "Status_", "Update_due_dates", "Last_Updated"), _
            Array(renewal(1), renewal(1), "Renewal", _
                  CrownLongDate(renewal(2)), CrownLongDate(renewal(3)), _
                  renewal(4), CrownDate(renewal(5)), CrownDate(renewal(6)), _
                  renewal(7), "Yes", Format$(Now, "mm-dd-yyyy hh:mm AM/PM")))

        If InStr(1, answer, """ok"":true", vbTextCompare) > 0 Then
            made = made + 1
            CrownNoteRenewal ws, CStr(renewal(0)), "renewal added as record " & CrownIdFrom(answer)
        ElseIf Len(Trim$(answer)) = 0 Then
            CrownNoteRenewal ws, CStr(renewal(0)), "renewal NOT added - the website did not answer"
        Else
            CrownNoteRenewal ws, CStr(renewal(0)), "renewal NOT added: " & answer
        End If
    Next i

    CrownMakeRenewals = made
End Function

' The new record number out of the website's reply.
Private Function CrownIdFrom(ByVal answer As String) As String
    Dim at As Long, i As Long, ch As String

    at = InStr(1, answer, """id"":")
    If at = 0 Then Exit Function

    For i = at + 5 To Len(answer)
        ch = Mid$(answer, i, 1)
        If ch < "0" Or ch > "9" Then Exit For
        CrownIdFrom = CrownIdFrom & ch
    Next i
End Function

' Put a word about the renewal on the line the policy is already on.
Private Sub CrownNoteRenewal(ByVal ws As Object, ByVal recordId As String, ByVal words As String)
    Dim r As Long

    For r = 2 To 5000
        If Len(Trim$(CStr(ws.Cells(r, 1).value))) = 0 Then Exit For

        If CStr(ws.Cells(r, 1).value) = recordId Then
            ws.Cells(r, 11).value = ws.Cells(r, 11).value & " " & words
            Exit Sub
        End If
    Next r
End Sub

' The policy form keeps its effective and expiration dates written out -
' "March 23, 2023". Anything that is not a date is left exactly as it is
' rather than being turned into today by accident.
Public Function CrownLongDate(ByVal text As String) As String
    Dim when As Double

    text = Trim$(text)
    If Len(text) = 0 Then Exit Function

    when = CrownUsDate(text)

    If when = 0 Then
        CrownLongDate = text
    Else
        CrownLongDate = Format$(CDate(when), "mmmm d, yyyy")
    End If
End Function


' ---------------------------------------------------------------------
' 4. Trisura (Verve), one policy at a time
' ---------------------------------------------------------------------
'
' Same shape as the United one, and for the same reason: the list of what
' to look up comes from our website, so no report has to be downloaded,
' opened, trimmed and re-saved first. The sign-in and the page reading are
' his own, taken from ScrapeVervePaymentdue - the same boxes, the same
' labels, the same wait for the grey panel to go away.
'
' What Verve says about money lives on one line of text:
'
'   "Installment For $239.17 due on 9/29/2025"
'   "... Cancel Date: 10/12/2025"
'
' The first means a live schedule; the second means a cancellation is on
' the way, and the amount and date are then read from their own labels.
'
' Verve policy numbers have letters in them - GAF20108199 - so unlike
' United they are sent exactly as we hold them.

' Waits for something, and gives up rather than spinning.
'
' Not LoopElementUntilFound: its second loop has no counter and no pause
' in it, so an element that is present but never shown turns into a busy
' loop and Excel stops answering. Counted rather than timed, because
' Timer goes back to zero at midnight and a run can cross it.
Private Function CrownWaitFor(ByVal drv As ChromeDriver, ByVal By As Selenium.By, _
                              ByVal how As String, ByVal what As String, _
                              Optional ByVal seconds As Double = 20) As Boolean
    Dim tries As Long, most As Long
    Dim there As Boolean

    most = CLng(seconds * 4)
    If most < 1 Then most = 1

    For tries = 1 To most
        there = False

        On Error Resume Next
        If UCase$(how) = "ID" Then
            there = drv.IsElementPresent(By.ID(what))
        Else
            there = drv.IsElementPresent(By.xpath(what))
        End If
        On Error GoTo 0

        If there Then
            CrownWaitFor = True
            Exit Function
        End If

        drv.Wait 250
        DoEvents
    Next tries
End Function

' Everything after a marker, or nothing at all.
'
' Written with InStr rather than Split: Split(text, marker)(1) on text
' that does not contain the marker is a subscript error, and the only
' thing holding his version together is an On Error above it.
Private Function CrownAfter(ByVal text As String, ByVal marker As String) As String
    Dim at As Long

    at = InStr(1, text, marker, vbTextCompare)
    If at = 0 Then Exit Function

    CrownAfter = Trim$(Mid$(text, at + Len(marker)))
End Function

Private Function CrownBetween(ByVal text As String, ByVal after As String, ByVal upto As String) As String
    Dim rest As String
    Dim at As Long

    rest = CrownAfter(text, after)
    If Len(rest) = 0 Then Exit Function

    at = InStr(1, rest, upto, vbTextCompare)

    If at = 0 Then
        CrownBetween = rest
    Else
        CrownBetween = Trim$(Left$(rest, at - 1))
    End If
End Function

' The first thing in a piece of text that is really a date.
'
' "due on 9/29/2025 (see schedule)" gives 9/29/2025 and not the rest of
' the sentence - a whole sentence written into a date field on the
' website is not something anyone would notice until it mattered.
Private Function CrownFirstDate(ByVal text As String) As String
    Dim i As Long
    Dim ch As String, run As String

    For i = 1 To Len(text) + 1
        If i <= Len(text) Then ch = Mid$(text, i, 1) Else ch = " "

        If (ch >= "0" And ch <= "9") Or ch = "/" Or ch = "-" Then
            run = run & ch
        Else
            If CrownUsDate(run) > 0 Then
                CrownFirstDate = run
                Exit Function
            End If

            run = ""
        End If
    Next i
End Function

' An amount with the dollar sign and the thousands commas taken off.
Private Function CrownMoney(ByVal text As String) As String
    Dim i As Long, ch As String, out As String

    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)

        If (ch >= "0" And ch <= "9") Or ch = "." Then
            out = out & ch
        ElseIf ch = "-" And Len(out) = 0 Then
            out = "-"
        End If
    Next i

    CrownMoney = out
End Function


' Is the policy we asked for actually on the screen?
'
' Punctuation out of both sides first: the website writes GAF20108199,
' GAF-20108199 and GAF 20108199 for the same policy, and a check that says
' no to all but one of those is worse than no check.
Private Function CrownPageShows(ByVal drv As ChromeDriver, ByVal policyNo As String, _
                                Optional ByVal seconds As Double = 15) As Boolean
    Dim answer As String
    Dim wanted As String
    Dim i As Long, ch As String
    Dim tries As Long, most As Long

    For i = 1 To Len(policyNo)
        ch = UCase$(Mid$(policyNo, i, 1))

        If (ch >= "0" And ch <= "9") Or (ch >= "A" And ch <= "Z") Then wanted = wanted & ch
    Next i

    ' Too short to tell anything from - so the answer is NO.
    '
    ' This said True until 6 Oct, meaning "cannot check, carry on". On the
    ' run of 5 Oct the browser got stuck on one policy's page for two
    ' hours; every real policy number was correctly refused, but records
    ' holding "999" and "NA" sailed through this line and were written
    ' with the figures of whoever was on screen. Two customers ended up
    ' holding a third customer's payment.
    '
    ' A wrong yes costs a customer's record. A wrong no costs one policy,
    ' and it is reported rather than silent.
    If Len(wanted) < 6 Then
        CrownPageShows = False
        Exit Function
    End If

    ' Waited for, not asked once. The page is still being fetched when this
    ' is first called, and a check that runs too early says no to every
    ' policy - which looks exactly like a check that is working.
    most = CLng(seconds * 4)
    If most < 1 Then most = 1

    For tries = 1 To most
        answer = ""

        On Error Resume Next
        answer = drv.ExecuteScript( _
            "var t=(document.body?(document.body.innerText||''):'').toUpperCase()" & _
            ".replace(/[^A-Z0-9]+/g,'');" & _
            "return t.indexOf('" & wanted & "')>=0?'1':'0';")
        On Error GoTo 0

        If answer = "1" Then
            CrownPageShows = True
            Exit Function
        End If

        drv.Wait 250
        DoEvents
    Next tries
End Function


' Open the Future tab underneath Billing.
'
' Matched on the word rather than on a position, because the tab strip is
' drawn by the same control that draws the one above it and an xpath into
' it would be four levels of nothing. Links are tried before anything
' else: the <li> around the link has the same text, and clicking the <li>
' does nothing at all.
Private Function CrownVerveOpenFuture(ByVal drv As ChromeDriver) As String
    Dim answer As String

    On Error Resume Next
    answer = drv.ExecuteScript( _
        "function pick(list){for(var i=0;i<list.length;i++){" & _
        "var t=((list[i].innerText||'')+'').replace(/\u00a0/g,' ')" & _
        ".replace(/^\s+|\s+$/g,'').toLowerCase();" & _
        "if(t.indexOf('future')===0&&t.length<24)return list[i];}return null;}" & _
        "var el=pick(document.getElementsByTagName('a'));" & _
        "if(!el)el=pick(document.querySelectorAll('span,td,li,div,button'));" & _
        "if(!el)return 'no Future tab';el.click();return 'opened';")
    On Error GoTo 0

    CrownVerveOpenFuture = answer
End Function

' Tick Combine Installments and Fees, if it is a checkbox we can reach.
'
' Worth doing - it is how he does it by hand - but never relied on. The
' page is built with DevExpress and half its checkboxes are a span with a
' background image rather than an input, so this is allowed to come back
' empty-handed and the adding up carries the day instead.
Private Function CrownVerveCombine(ByVal drv As ChromeDriver) As String
    Dim answer As String

    On Error Resume Next
    answer = drv.ExecuteScript( _
        "function label(b){var t='';" & _
        "if(b.id){var l=document.querySelector('label[for=\""'+b.id+'\""]');" & _
        "if(l)t=l.innerText||'';}" & _
        "if(!t){var p=b.parentNode,n=0;while(p&&n<3){var x=(p.innerText||'')+'';" & _
        "if(x.length>0&&x.length<90){t=x;break;}p=p.parentNode;n++;}}" & _
        "return (t+'').toLowerCase();}" & _
        "var bs=document.querySelectorAll('input[type=checkbox]');" & _
        "for(var i=0;i<bs.length;i++){if(label(bs[i]).indexOf('combine')<0)continue;" & _
        "if(bs[i].checked)return 'already ticked';bs[i].click();return 'ticked';}" & _
        "return 'no combine box found';")
    On Error GoTo 0

    CrownVerveCombine = answer
End Function

' The Future table as plain text: rows separated by ~, cells by |.
'
' Same approach as the United result list, and for the same reason - the
' page is tables nested inside tables and every one of them contains the
' words being looked for, so the smallest one that does is the real one.
'
' Smallest alone is not enough here, and testing it against a copy of the
' page showed why: the statement sits in the same document as the future
' list, has the same four headings, and has fewer rows - so it wins on
' size and the figure written to the website is last month's instalment,
' already paid, dated in the past. Only tables actually on the screen are
' considered, which is what tells a tab that is showing from a tab that
' is merely still loaded.
Private Function CrownVerveFutureRows(ByVal drv As ChromeDriver) As String
    Dim answer As String

    On Error Resume Next
    answer = drv.ExecuteScript( _
        "var best=null,small=0,bestTrans=false," & _
        "tabs=document.getElementsByTagName('table');" & _
        "for(var i=0;i<tabs.length;i++){var t=tabs[i],x=(t.innerText||'')+'';" & _
        "if(x.indexOf('Due Date')<0||x.indexOf('Amount')<0)continue;" & _
        "if(!t.getClientRects().length)continue;" & _
        "var trans=x.indexOf('Trans Date')>=0;" & _
        "if(bestTrans&&!trans)continue;" & _
        "if(best===null||(trans&&!bestTrans)||x.length<small){" & _
        "best=t;small=x.length;bestTrans=trans;}}" & _
        "if(!best)return '';var out=[];" & _
        "for(var r=0;r<best.rows.length;r++){var cs=best.rows[r].cells,line=[];" & _
        "for(var c=0;c<cs.length;c++){line.push(((cs[c].innerText||'')+'')" & _
        ".replace(/[|~]/g,' ').replace(/\s+/g,' ').replace(/^ | $/g,''));}" & _
        "out.push(line.join('|'));}return out.join('~');")
    On Error GoTo 0

    CrownVerveFutureRows = answer
End Function

' The next payment out of that table.
'
' The earliest line carrying an amount above zero, plus anything else
' falling on the same day. Adding them is the point: unticked, Verve
' shows the instalment and its fee as two lines, and 143.16 arrives as
' 126.48 and 16.68. Zero and credit lines are passed over - the 0.00 on
' the line above the instalment is an Invoice/Legal Cancellation Notice,
' and reading it as a payment tells a customer who owes 143.16 that they
' owe nothing.
'
' Yesterday is not a future payment. This is only ever asked when the
' policy owes nothing today, so a line dated in the past cannot be money
' owed - it is the statement being read by mistake, and it would put a
' bill the customer has already paid back on their screen.
Private Function CrownVerveNextDue(ByVal rowsText As String, ByRef dueDate As String, _
                                   ByRef amount As String, ByRef note As String) As Boolean
    Dim rows() As String, header() As String, cells() As String
    Dim colDue As Long, colAmount As Long, colWhat As Long
    Dim r As Long, lines As Long
    Dim money As String, thisDate As String, what As String
    Dim total As Double

    dueDate = ""
    amount = ""

    If Len(rowsText) = 0 Then
        note = "no Future table on the page"
        Exit Function
    End If

    rows = Split(rowsText, "~")

    If UBound(rows) < 1 Then
        note = "the Future tab had no lines on it"
        Exit Function
    End If

    header = Split(rows(0), "|")
    colDue = CrownColumnOf(header, "Due Date")
    colAmount = CrownColumnOf(header, "Amount")
    colWhat = CrownColumnOf(header, "Description")

    If colDue < 0 Or colAmount < 0 Then
        note = "the Future tab did not have the columns expected"
        Exit Function
    End If

    For r = 1 To UBound(rows)
        cells = Split(rows(r), "|")

        If UBound(cells) >= colDue And UBound(cells) >= colAmount Then
            money = CrownMoney(cells(colAmount))
            thisDate = CrownFirstDate(cells(colDue))

            If Len(money) > 0 And Len(thisDate) > 0 _
               And CrownUsDate(thisDate) >= CDbl(Date) Then
                If Val(money) > 0 Then
                    If Len(dueDate) = 0 Then
                        dueDate = thisDate

                        If colWhat >= 0 And UBound(cells) >= colWhat Then
                            what = Trim$(cells(colWhat))
                        End If
                    End If

                    ' Every line falling on the same day, wherever it sits
                    ' in the table. Walked to the end rather than stopped
                    ' at the first different date, so an out-of-order row
                    ' cannot quietly drop a fee.
                    If thisDate = dueDate Then
                        total = total + Val(money)
                        lines = lines + 1
                    End If
                End If
            End If
        End If
    Next r

    If lines = 0 Then
        note = "nothing still to come on the Future tab"
        Exit Function
    End If

    amount = Format$(total, "0.00")

    note = "next payment from the Future tab"

    If Len(what) > 0 Then note = note & " - " & what

    If lines > 1 Then note = note & ", " & lines & " lines added together"

    CrownVerveNextDue = True
End Function


' Could this be a policy number at all?
'
' Asked before anything is typed into the carrier's site, so a record
' holding "999", "NA" or a blank gets a plain note rather than a lookup
' that cannot be verified afterwards. Six characters and four digits is
' below every real number any of these carriers issues - GAF20108199 has
' ten of each - and comfortably above the junk.
Private Function CrownLooksLikePolicy(ByVal policyNo As String) As Boolean
    Dim i As Long, ch As String
    Dim kept As String, digits As Long

    For i = 1 To Len(policyNo)
        ch = UCase$(Mid$(policyNo, i, 1))

        If (ch >= "0" And ch <= "9") Then
            kept = kept & ch
            digits = digits + 1
        ElseIf ch >= "A" And ch <= "Z" Then
            kept = kept & ch
        End If
    Next i

    CrownLooksLikePolicy = (Len(kept) >= 6 And digits >= 4)
End Function

' Get back to a page that has the quick lookup on it.
'
' Verve stops answering the lookup box now and then - a bad number, a
' session that has gone stale, a dialog nobody saw. Until 6 Oct nothing
' noticed, and a run would carry on asking a dead page for two hours.
Private Function CrownVerveRecover(ByVal drv As ChromeDriver, ByVal By As Selenium.By, _
                                   ByVal wsInput As Worksheet) As Boolean
    On Error Resume Next
    drv.Get wsInput.Range("URL_9").value
    On Error GoTo 0

    CrownVerveRecover = CrownWaitFor(drv, By, "ID", VERVE_LOOKUP_BOX, 45)
End Function


' Everything on the page that could possibly be the Future tab, and the
' headings of every table on it.
'
' Not a guess at a selector - the page's own words, brought back so they
' can be read by somebody who can see them. Three goes at guessing from a
' screenshot have each looked right and each been wrong; this is cheaper
' than a fourth.
Private Function CrownVerveLook(ByVal drv As ChromeDriver) As String
    Dim answer As String

    On Error Resume Next
    answer = drv.ExecuteScript( _
        "function clean(s){return ((s||'')+'')" & _
        ".replace(/\s+/g,' ').replace(/^ | $/g,'');}" & _
        "var out=[],rest=[],seen={};" & _
        "var all=document.querySelectorAll('a,span,td,li,div,button');" & _
        "for(var i=0;i<all.length;i++){var t=clean(all[i].innerText);" & _
        "if(!t||t.length>26)continue;if(all[i].getElementsByTagName('*').length>2)continue;" & _
        "if(seen[t])continue;seen[t]=1;" & _
        "if(/future|statement|billing|installment|session/i.test(t))out.push(t);" & _
        "else if(rest.length<40)rest.push(t);}" & _
        "out=out.concat(rest);" & _
        "var tabs=document.getElementsByTagName('table'),heads=[];" & _
        "for(var j=0;j<tabs.length;j++){var r=tabs[j].rows[0];if(!r)continue;" & _
        "if(!tabs[j].getClientRects().length)continue;var cells=[];" & _
        "for(var c=0;c<r.cells.length;c++){cells.push(clean(r.cells[c].innerText));}" & _
        "var line=cells.join(' / ');if(line.replace(/[ \/]/g,'')==='')continue;" & _
        "heads.push(tabs[j].rows.length+' rows: '+line.substring(0,150));" & _
        "if(heads.length>12)break;}" & _
        "return 'WORDS ON THE PAGE: '+out.join(' | ')+String.fromCharCode(10)+" & _
        "'TABLES ON SCREEN: '+(heads.length?heads.join(String.fromCharCode(10)):'(none)');")
    On Error GoTo 0

    CrownVerveLook = answer
End Function

' The sheet those notes go on, made only if something needs writing down.
Private Function CrownLookSheet(ByRef ws As Worksheet, ByRef row As Long) As Boolean
    If Not ws Is Nothing Then
        CrownLookSheet = True
        Exit Function
    End If

    Set ws = CrownSheet("VerveLook")
    ws.Cells.ClearContents
    ws.Range("A1:C1").value = Array("Record number", "Policy number", "What the page says")
    ws.Columns(3).ColumnWidth = 120
    ws.Columns(3).WrapText = False
    row = 2

    CrownLookSheet = True
End Function


' Press OK on Verve's "Session Expiring" box.
'
' It was on all ten pages the VerveLook sheet recorded. It is a modal, so
' until it is dismissed every click lands on it rather than on the tab
' underneath - which is why the Future tab was never reached.
'
' Scoped to the box itself. There is an OK on other things on that site
' and pressing the wrong one would confirm something nobody asked to
' confirm, so the button has to be inside an element whose own text
' mentions an expiring session.
'
' "The smallest element whose text matches" is not enough, and the first
' version of this was wrong for exactly that reason: the smallest match
' is usually the LABEL - a div holding the words "Session Expiring" and
' nothing else. The element wanted is the smallest one that both says it
' AND has something pressable in it.
Private Function CrownVerveKeepAlive(ByVal drv As ChromeDriver) As String
    Dim answer As String

    On Error Resume Next
    answer = drv.ExecuteScript( _
        "function clean(s){return ((s||'')+'').replace(/\s+/g,' ')" & _
        ".replace(/^ | $/g,'');}" & _
        "function okIn(el){var hits=el.querySelectorAll('input,button,a,span,td');" & _
        "for(var j=0;j<hits.length;j++){var w=clean(hits[j].innerText||hits[j].value);" & _
        "if(!w)continue;w=w.toLowerCase();" & _
        "if(w==='ok'||w==='continue'||w==='yes'||w==='keep working')return hits[j];}" & _
        "return null;}" & _
        "var best=null,small=0,btn=null,sawWords=false;" & _
        "var all=document.querySelectorAll('div,table,form,section');" & _
        "for(var i=0;i<all.length;i++){var el=all[i],t=clean(el.innerText);" & _
        "if(!/session\s*expir/i.test(t))continue;" & _
        "if(t.length>400)continue;" & _
        "if(!el.getClientRects().length)continue;" & _
        "sawWords=true;" & _
        "var b=okIn(el);if(!b)continue;" & _
        "if(best===null||t.length<small){best=el;small=t.length;btn=b;}}" & _
        "if(btn){btn.click();" & _
        "return 'pressed '+clean(btn.innerText||btn.value).toLowerCase();}" & _
        "return sawWords?'session box with no OK on it':'no session box';")
    On Error GoTo 0

    CrownVerveKeepAlive = answer
End Function


Public Sub CrownVervePaymentDue()
    Dim drv As ChromeDriver, clsDrv As Chrm
    Dim By As New Selenium.By
    Dim Keys As New Selenium.Keys
    Dim wsInput As Worksheet, ws As Worksheet
    Dim policies As Collection, parts As Variant
    Dim recordId As String, policyNo As String
    Dim activity As String, dueAmount As String, dueDate As String
    Dim cancelDate As String, policyStatus As String
    Dim answer As String, note As String
    Dim names As Variant, values As Variant
    Dim statuses As Object
    Dim reported As String, reportNote As String, cancelledOn As String
    Dim futureNote As String, futureDate As String, futureAmount As String
    Dim missed As Long, stopped As Boolean, badNumber As Long
    Dim futureRows As String
    Dim sessionNote As String
    Dim dismissed As Long
    Dim look As Long, looked As Long
    Dim wsLook As Worksheet, lookRow As Long
    Dim row As Long, done As Long, blank As Long, skipped As Long
    Dim i As Long

    Set wsInput = ThisWorkbook.Worksheets("Input")

    Set policies = CrownPolicyRows("Verve")

    If policies.Count = 0 Then
        MsgBox "No Trisura (Verve) policies came back from the website.", vbExclamation, "Crown Superior"
        Exit Sub
    End If

    If MsgBox(policies.Count & " Trisura (Verve) policies to look up, about " & _
              Int(policies.Count / 8) + 1 & " minutes." & vbCrLf & vbCrLf & _
              "Carry on?", vbYesNo + vbQuestion, "Crown Superior") <> vbYes Then
        Exit Sub
    End If

    Set ws = CrownSheet("VervePayments")
    ws.Cells.ClearContents
    ws.Columns(2).NumberFormat = "@"
    ws.Range("A1:H1").value = Array("Record number", "Policy number", "Amount due", "Due date", _
                                    "Cancel date", "Policy status", "When", "What happened")
    row = 2

    Set clsDrv = New Chrm
    Set clsDrv.ChrmDriver = New ChromeDriver
    Set drv = clsDrv.ChrmDriver

    On Error Resume Next
    drv.AddArgument "--blink-settings=imagesEnabled=false"
    drv.AddArgument "--disable-gpu"
    drv.AddArgument "--disable-extensions"
    drv.AddArgument "--disable-popup-blocking"
    drv.AddArgument "--disable-notifications"
    drv.AddArgument "--force-device-scale-factor=0.80"
    drv.Start
    On Error GoTo 0

    drv.Get wsInput.Range("URL_9").value

    If Not CrownWaitFor(drv, By, "ID", "LoginControl_LoginNameTextBox", 40) Then
        MsgBox "Verve did not show its sign-in page. Nothing has been changed.", _
               vbExclamation, "Crown Superior"
        Exit Sub
    End If

    EnterData drv, Keys, "LoginControl_LoginNameTextBox", wsInput.Range("USER_2").value, "ID"
    EnterData drv, Keys, "LoginControl_PasswordTextBox", wsInput.Range("PASS_2").value, "ID"
    ClickElement drv, "LoginControl_LoginLinkButton", "ID"

    ' The statuses first, from the Amount Due report - the only place
    ' Verve says what a policy's status is. If any of that does not go to
    ' plan it says so and the run carries on; the status worked out from
    ' the billing line is worse than the report, not useless.
    Set statuses = CrownVerveStatuses(drv, By, Keys, wsInput, reportNote)

    ' The quick lookup box sits in the header of every page once signed
    ' in, which is what lets this ask for one policy after another
    ' without navigating back to anything.
    If Not CrownWaitFor(drv, By, "ID", VERVE_LOOKUP_BOX, 60) Then
        MsgBox "Signed in, but the quick policy lookup never appeared. " & _
               "Nothing has been changed.", vbExclamation, "Crown Superior"
        Exit Sub
    End If

    For i = 1 To policies.Count
        parts = policies(i)
        recordId = parts(0)
        policyNo = Trim$(CStr(parts(2)))

        If Len(policyNo) = 0 Then GoTo NextVerve

        ' Not a policy number, so it is not looked up. Written down rather
        ' than passed over in silence - a record holding "999" is something
        ' he wants to know about and fix at his end.
        If Not CrownLooksLikePolicy(policyNo) Then
            ws.Cells(row, 1).value = recordId
            ws.Cells(row, 2).value = policyNo
            ws.Cells(row, 8).value = "that is not a policy number - nothing looked up, nothing written"
            badNumber = badNumber + 1
            row = row + 1
            GoTo NextVerve
        End If

        If CrownLongCancelled(parts, cancelledOn) Then
            ws.Cells(row, 1).value = recordId
            ws.Cells(row, 2).value = policyNo
            ws.Cells(row, 8).value = "skipped - cancelled since " & cancelledOn
            skipped = skipped + 1
            row = row + 1
            GoTo NextVerve
        End If

        activity = ""
        dueAmount = ""
        dueDate = ""
        cancelDate = ""
        policyStatus = ""
        note = ""
        futureNote = ""
        futureDate = ""
        futureAmount = ""

        On Error Resume Next

        ' Before anything is typed. The box is modal, so a lookup typed
        ' underneath it goes nowhere and every policy after that reads the
        ' page we were already on.
        sessionNote = CrownVerveKeepAlive(drv)

        If Left$(sessionNote, 7) = "pressed" Then
            drv.Wait 600
            dismissed = dismissed + 1
        End If

        EnterData drv, Keys, VERVE_LOOKUP_BOX, policyNo, "ID", "", True
        ClickElement drv, VERVE_LOOKUP_GO, "ID"

        If Not CrownPageShows(drv, policyNo) Then
            ' Verve is still showing whatever it was showing before. Reading
            ' it would put this customer's record on another customer's
            ' figures, which is exactly what happened on United.
            note = "Verve did not open this policy - nothing written"
            missed = missed + 1

            ' A few in a row is not bad luck, it is a stuck page. Go back to
            ' the start and see whether the lookup comes back; if it does
            ' not, stop rather than spend two hours asking a dead page.
            If missed = CONSECUTIVE_BEFORE_RELOAD Then
                If CrownVerveRecover(drv, By, wsInput) Then
                    note = note & " - went back to the start to clear it"
                Else
                    note = note & " - Verve stopped answering"
                End If
            ElseIf missed >= CONSECUTIVE_BEFORE_GIVING_UP Then
                note = note & " - stopping, Verve is not answering"
                stopped = True
            End If
        ElseIf CrownWaitFor(drv, By, "XPATH", VERVE_BILLING_TAB, 25) Then
            ClickElement drv, VERVE_BILLING_TAB, "XPATH"
            drv.Wait 1000

            ' Verve draws a grey panel over the figures while it fetches
            ' them. Reading underneath it gives the last policy's numbers.
            If Not WaitForPopupToDisappear(drv, 20000) Then
                note = "Verve was still loading after twenty seconds"
            End If

            drv.Wait 400

            ' Checked again after the tab has loaded: the first check was of
            ' the page before this one.
            If Not CrownPageShows(drv, policyNo, 5) Then
                note = "the billing tab was showing another policy - nothing written"
                GoTo VerveWrite
            End If

            activity = TryGetText(drv, VERVE_ACTIVITY, "ID", 3, 300)

            If InStr(1, activity, "Installment For", vbTextCompare) > 0 Then
                dueAmount = CrownMoney(CrownBetween(activity, "Installment For $", " due on"))
                dueDate = CrownFirstDate(CrownAfter(activity, "due on"))
                policyStatus = "Active"
            ElseIf InStr(1, activity, "Cancel Date:", vbTextCompare) > 0 Then
                cancelDate = CrownFirstDate(CrownAfter(activity, "Cancel Date:"))
                dueDate = CrownFirstDate(TryGetText(drv, VERVE_DUE_DATE, "ID", 3, 300))
                dueAmount = CrownMoney(TryGetText(drv, VERVE_BALANCE, "ID", 3, 300))
                policyStatus = "Cx Notice"
            ElseIf Len(Trim$(activity)) > 0 Then
                ' Something else on that line. The two labels still hold
                ' the figures; the status is left as it is rather than
                ' guessed at from wording nobody has seen yet.
                dueDate = CrownFirstDate(TryGetText(drv, VERVE_DUE_DATE, "ID", 3, 300))
                dueAmount = CrownMoney(TryGetText(drv, VERVE_BALANCE, "ID", 3, 300))
                note = "Verve said: " & Left$(activity, 80)
            End If

            ' Nothing owed today does not mean nothing owed. The
            ' billing line is about now; the Future tab is where the next
            ' instalment is, and that is the figure the customer wants to
            ' see on the website.
            '
            ' Only consulted when today is clear. A policy in arrears owes
            ' that money now and today's figure has to win.
            If Val("0" & dueAmount) <= 0 Then
                ' And again here. The billing tab is a fresh request, and
                ' the box can come back between one and the next.
                If Left$(CrownVerveKeepAlive(drv), 7) = "pressed" Then
                    drv.Wait 600
                    dismissed = dismissed + 1
                End If

                futureNote = CrownVerveOpenFuture(drv)

                If futureNote = "opened" Then
                    WaitForPopupToDisappear drv, 15000
                    drv.Wait 600

                    ' Ticked if it can be found, which is his way of doing
                    ' it. The adding up below does not depend on it.
                    If CrownVerveCombine(drv) = "ticked" Then
                        WaitForPopupToDisappear drv, 15000
                        drv.Wait 600
                    End If

                    ' The same check again. Everything read from here on
                    ' is written against this customer's record and a tab
                    ' that quietly reloaded somebody else is how seventy
                    ' records ended up on one policy's figures.
                    If Not CrownPageShows(drv, policyNo, 5) Then
                        futureNote = "the Future tab was showing another policy - not read"
                    Else
                        ' Asked more than once. The panel is fetched in the
                        ' background, so the first look can be at a page
                        ' that has not finished arriving - which reads
                        ' exactly like a page with nothing on it.
                        futureRows = ""

                        For look = 1 To FUTURE_LOOKS
                            futureRows = CrownVerveFutureRows(drv)
                            If Len(futureRows) > 0 Then Exit For
                            drv.Wait 900
                            DoEvents
                        Next look

                        If CrownVerveNextDue(futureRows, futureDate, futureAmount, futureNote) Then
                            dueAmount = futureAmount
                            dueDate = futureDate
                        End If
                    End If
                End If

                note = Trim$(note & " " & futureNote)

                ' Nothing readable. Write down what IS on the page, for the
                ' first few only - enough to see the shape of it without
                ' turning the run into a transcript.
                If Len(futureAmount) = 0 And looked < PAGES_TO_WRITE_DOWN Then
                    If CrownLookSheet(wsLook, lookRow) Then
                        wsLook.Cells(lookRow, 1).value = recordId
                        wsLook.Cells(lookRow, 2).value = policyNo
                        wsLook.Cells(lookRow, 3).value = futureNote _
                            & vbCrLf & "session box: " & sessionNote _
                            & vbCrLf & CrownVerveLook(drv)
                        lookRow = lookRow + 1
                        looked = looked + 1
                    End If
                End If
            End If

            ' What the report says beats what the billing line implies.
            reported = ""

            If statuses.Exists(UCase$(policyNo)) Then
                reported = Trim$(CStr(statuses(UCase$(policyNo))))
            ElseIf statuses.Exists(CrownDigits(policyNo)) Then
                reported = Trim$(CStr(statuses(CrownDigits(policyNo))))
            End If

            If Len(reported) > 0 Then policyStatus = reported
        Else
            note = "no policy of that number at Verve"
        End If

VerveWrite:
        On Error GoTo 0

        ws.Cells(row, 1).value = recordId
        ws.Cells(row, 2).value = policyNo
        ws.Cells(row, 3).value = dueAmount
        ws.Cells(row, 4).value = dueDate
        ws.Cells(row, 5).value = cancelDate
        ws.Cells(row, 6).value = policyStatus
        ws.Cells(row, 7).value = Now

        ' Anything read counts as the page answering again.
        If Len(Trim$(dueAmount)) > 0 Or Len(Trim$(dueDate)) > 0 _
           Or Len(Trim$(cancelDate)) > 0 Or Len(Trim$(policyStatus)) > 0 Then
            missed = 0
        End If

        If Len(Trim$(dueAmount)) = 0 And Len(Trim$(dueDate)) = 0 _
           And Len(Trim$(cancelDate)) = 0 And Len(Trim$(policyStatus)) = 0 Then
            ws.Cells(row, 8).value = Trim$("nothing found - left alone. " & note)
            blank = blank + 1
        Else
            names = Array("Web_pymnt_due", "updated_due_date", "updated_cancel_date", _
                          "Update_due_dates", "Last_Updated")
            values = Array(dueAmount, CrownDate(dueDate), CrownDate(cancelDate), _
                           "Yes", Format$(Now, "mm-dd-yyyy hh:mm AM/PM"))

            ' The status is only sent when the page actually said one.
            ' An empty one would write over what the website holds, and a
            ' policy that quietly loses its status stops being called.
            If Len(policyStatus) > 0 Then
                names = CrownWith(names, "Status_")
                values = CrownWith(values, policyStatus)
            End If

            answer = CrownUpdate(CROWN_FORM_POLICY, CLng(recordId), names, values)

            If InStr(1, answer, """ok"":true", vbTextCompare) > 0 Then
                ws.Cells(row, 8).value = Trim$("written. " & note)
                done = done + 1
            ElseIf Len(Trim$(answer)) = 0 Then
                ws.Cells(row, 8).value = "website did not answer - run again for this one"
            Else
                ws.Cells(row, 8).value = "refused: " & answer
            End If
        End If

        row = row + 1

NextVerve:
        If stopped Then Exit For
    Next i

    ws.Activate

    MsgBox policies.Count & " Trisura (Verve) policies on the website." & vbCrLf & _
           skipped & " cancelled more than " & CANCELLED_LONG_ENOUGH & " days ago and passed over." & vbCrLf & _
           badNumber & " had something in the policy number box that is not a policy number." & vbCrLf & _
           done & " written back to the website." & vbCrLf & _
           blank & " had nothing to read and were left as they were." & vbCrLf & _
           IIf(stopped, vbCrLf & "STOPPED EARLY: Verve stopped answering after " & _
               CONSECUTIVE_BEFORE_GIVING_UP & " policies in a row. Nothing was written from " & _
               "that point on. Run it again when the site is behaving." & vbCrLf, "") & vbCrLf & _
           reportNote & vbCrLf & vbCrLf & _
           IIf(dismissed > 0, "Verve's Session Expiring box was dismissed " & dismissed & _
               " times." & vbCrLf, "") & _
           IIf(looked > 0, "The Future tab could not be read on " & looked & " policies. " & _
               "What those pages actually say is written on the VerveLook sheet - " & _
               "please send me that sheet." & vbCrLf & vbCrLf, "") & _
           "See the VervePayments sheet for the detail.", vbInformation, "Crown Superior"
End Sub

' One more on the end of an array, without caring how long it was.
Private Function CrownWith(ByVal list As Variant, ByVal extra As String) As Variant
    Dim out() As String
    Dim i As Long, n As Long

    n = UBound(list) - LBound(list) + 1
    ReDim out(0 To n)

    For i = 0 To n - 1
        out(i) = CStr(list(LBound(list) + i))
    Next i

    out(n) = extra
    CrownWith = out
End Function

' ---------------------------------------------------------------------
' 5. The Amount Due report - only for the statuses
' ---------------------------------------------------------------------
'
' Verve's policy page does not say whether a policy is in force; the
' Amount Due report does. This fetches it the way his own macro does,
' takes the two columns it needs, and leaves the file alone.
'
' His version deletes columns J:U, then F:H, then A:B, and works out what
' is left by counting. That is right until Verve adds a column. This one
' reads the heading row and looks for the words - and if the words are
' not there it says so and prints the heading row it did get, so the next
' run can be told exactly what to look for instead of guessing again.
'
' Nothing in here is allowed to stop the run. A missing report means the
' status falls back to what the billing line implies, which is what it
' was doing yesterday.

' The newest .csv put in Downloads since a moment we noted before asking
' for one.
'
' Chosen by when it arrived rather than by its name, so nothing depends on
' what the download was called - and yesterday's report can never answer
' for today's.
Private Function CrownNewestCsvSince(ByVal notBefore As Date) As String
    Dim fso As Object, folder As Object, file As Object
    Dim best As String
    Dim bestAt As Date

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(Environ$("USERPROFILE") & "\Downloads")
    On Error GoTo 0

    If folder Is Nothing Then Exit Function

    For Each file In folder.Files
        If LCase$(fso.GetExtensionName(file.Name)) = "csv" Then
            If file.DateLastModified >= notBefore And file.DateLastModified > bestAt Then
                bestAt = file.DateLastModified
                best = file.Path
            End If
        End If
    Next file

    CrownNewestCsvSince = best
End Function

' Two columns out of a csv, found by what their headings say.
'
' Opened for reading only: it is his download, in his Downloads folder,
' and nothing here has any business changing it.
'
' Each row is filed twice, under the policy number as written and under
' its digits alone, because no two systems here write a policy number the
' same way.
Private Function CrownCsvLookup(ByVal path As String, ByVal keyHeading As String, _
                                ByVal valueHeading As String, ByRef heading As String, _
                                ByRef found As Long) As Object
    Dim map As Object
    Dim parts() As String
    Dim fileNumber As Integer
    Dim line As String, key As String, digits As String, value As String
    Dim keyAt As Long, valueAt As Long
    Dim first As Boolean, byContent As Boolean

    Set map = CreateObject("Scripting.Dictionary")
    Set CrownCsvLookup = map
    heading = ""
    found = 0

    If Len(path) = 0 Then Exit Function

    keyAt = -1
    valueAt = -1
    first = True

    On Error GoTo Finished

    fileNumber = FreeFile
    Open path For Input As #fileNumber

    Do Until EOF(fileNumber)
        Line Input #fileNumber, line

        If Len(Trim$(line)) > 0 Then
            parts = CrownSplitCsvLine(line)

            If first Then
                heading = line
                keyAt = CrownColumnOf(parts, keyHeading, valueHeading)
                valueAt = CrownColumnOf(parts, valueHeading)
                first = False

                ' This report comes out of a report builder that names its
                ' columns textbox6, textbox25, textbox32 - so there is
                ' nothing to match a heading against. The columns are then
                ' found by what is IN them instead, further down.
                If keyAt < 0 Or valueAt < 0 Then
                    byContent = True

                    Exit Do
                End If
            ElseIf UBound(parts) >= keyAt And UBound(parts) >= valueAt Then
                key = UCase$(Trim$(parts(keyAt)))
                value = Trim$(parts(valueAt))

                If Len(key) > 0 And Len(value) > 0 Then
                    If Not map.Exists(key) Then
                        map.Add key, value
                        found = found + 1
                    End If

                    digits = CrownDigits(key)
                    If Len(digits) > 0 And Not map.Exists(digits) Then map.Add digits, value
                End If
            End If
        End If
    Loop

Finished:
    Close #fileNumber

    If byContent Then
        Set map = CrownCsvByContent(path, heading, found)
        Set CrownCsvLookup = map
    End If
End Function

' The same report, when its headings are no use.
'
' The policy column is the one whose values look like policy numbers -
' that is safe, because a policy number has a shape and nothing else in
' the report shares it.
'
' The status column is only accepted if every different value in it is a
' word this agency actually uses for a status. A column of unrecognised
' words is NOT taken as the status: writing a guess into the status field
' of a thousand policies is a far worse outcome than leaving them as they
' are, and the sheet says so rather than going quiet.
Private Function CrownCsvByContent(ByVal path As String, ByRef heading As String, _
                                   ByRef found As Long) As Object
    Dim map As Object
    Dim rows As Collection, parts As Variant
    Dim fileNumber As Integer
    Dim line As String, key As String, digits As String, value As String
    Dim i As Long, c As Long, widest As Long
    Dim policyAt As Long, statusAt As Long
    Dim hits() As Long, good() As Long, seen() As Long

    Set map = CreateObject("Scripting.Dictionary")
    Set CrownCsvByContent = map
    Set rows = New Collection
    found = 0

    If Len(path) = 0 Then Exit Function

    On Error GoTo Done

    fileNumber = FreeFile
    Open path For Input As #fileNumber

    Do Until EOF(fileNumber)
        Line Input #fileNumber, line

        If Len(Trim$(line)) > 0 Then
            parts = CrownSplitCsvLine(line)
            rows.Add parts

            If UBound(parts) > widest Then widest = UBound(parts)
        End If
    Loop

    Close #fileNumber

    If rows.Count < 2 Then Exit Function

    ReDim hits(0 To widest)
    ReDim good(0 To widest)
    ReDim seen(0 To widest)

    ' Row 1 is the heading row, whatever it says. Counted from row 2.
    For i = 2 To rows.Count
        parts = rows(i)

        For c = 0 To widest
            If UBound(parts) >= c Then
                value = Trim$(CStr(parts(c)))

                If Len(value) > 0 Then
                    seen(c) = seen(c) + 1

                    If CrownCouldBePolicyColumn(value) Then hits(c) = hits(c) + 1
                    If CrownIsStatusWord(value) Then good(c) = good(c) + 1
                End If
            End If
        Next c
    Next i

    policyAt = -1
    statusAt = -1

    For c = 0 To widest
        If seen(c) > 0 Then
            If hits(c) * 2 > seen(c) Then
                If policyAt < 0 Or hits(c) > hits(policyAt) Then policyAt = c
            End If

            ' Every value recognised, not merely most of them. A column
            ' that is right nine times in ten is a column that writes the
            ' wrong status onto one policy in ten.
            If good(c) = seen(c) Then
                If statusAt < 0 Or seen(c) > seen(statusAt) Then statusAt = c
            End If
        End If
    Next c

    If policyAt < 0 Or statusAt < 0 Then
        heading = heading & "  [columns found by content: policy " & policyAt & ", status " & statusAt & "]"

        Exit Function
    End If

    For i = 2 To rows.Count
        parts = rows(i)

        If UBound(parts) >= policyAt And UBound(parts) >= statusAt Then
            key = UCase$(Trim$(CStr(parts(policyAt))))
            value = Trim$(CStr(parts(statusAt)))

            If Len(key) > 0 And Len(value) > 0 Then
                If Not map.Exists(key) Then
                    map.Add key, value
                    found = found + 1
                End If

                digits = CrownDigits(key)
                If Len(digits) > 0 And Not map.Exists(digits) Then map.Add digits, value
            End If
        End If
    Next i

Done:
End Function

' The same shape test, but stricter, for working out which column is
' which.
'
' A date passes the ordinary test - 10/28/2026 is ten characters with
' eight digits in it - so a report that puts the date before the policy
' number would have its dates read as policy numbers. No carrier number
' in this book has a slash in it and every date does.
Private Function CrownCouldBePolicyColumn(ByVal text As String) As Boolean
    If InStr(text, "/") > 0 Then Exit Function

    CrownCouldBePolicyColumn = CrownLooksLikePolicy(text)
End Function

' Is this one of the words this agency uses for a policy's status?
'
' A fixed list on purpose. Anything outside it means the column is not the
' status column, and the statuses are left alone.
Private Function CrownIsStatusWord(ByVal text As String) As Boolean
    Dim word As String
    Dim known As Variant
    Dim i As Long

    word = LCase$(Trim$(text))

    known = Array("active", "cancelled", "canceled", "cancel", "cancel notice", _
                  "cx notice", "pending", "pending cancellation", "expired", "lapsed", _
                  "reinstated", "non-renewed", "nonrenewed", "non renewal", "in force", _
                  "inforce", "new", "renewed", "rewritten")

    For i = LBound(known) To UBound(known)
        If word = known(i) Then
            CrownIsStatusWord = True

            Exit Function
        End If
    Next i
End Function

Private Function CrownVerveStatuses(ByVal drv As ChromeDriver, ByVal By As Selenium.By, _
                                    ByVal Keys As Selenium.Keys, ByVal wsInput As Worksheet, _
                                    ByRef note As String) As Object
    Dim map As Object
    Dim asked As Date
    Dim path As String, heading As String
    Dim tries As Long, found As Long

    Set map = CreateObject("Scripting.Dictionary")
    Set CrownVerveStatuses = map
    note = "no statuses from the report - each one is as the billing line reads it"

    On Error Resume Next

    If Not CrownWaitFor(drv, By, "ID", VERVE_TOOLS_MENU, 30) Then
        note = "Agency Tools never appeared - statuses are as the billing line reads them"
        GoTo BackToTheSite
    End If

    ClickElement drv, VERVE_TOOLS_MENU, "ID"
    drv.Wait 2000
    ClickElement drv, VERVE_REPORTS_LINK, "XPATH"
    drv.Wait 2000
    drv.ExecuteScript "window.scrollTo(0, 168);"

    If Not CrownWaitFor(drv, By, "ID", VERVE_AMOUNT_DUE, 30) Then
        note = "the Amount Due report was not on the reports list - statuses are as the billing line reads them"
        GoTo BackToTheSite
    End If

    ClickElement drv, VERVE_AMOUNT_DUE, "ID"
    drv.Wait 2000

    If Not CrownWaitFor(drv, By, "ID", VERVE_RPT_ACTIVE, 30) Then
        note = "the report would not open - statuses are as the billing line reads them"
        GoTo BackToTheSite
    End If

    SelectOption drv, VERVE_RPT_ACTIVE, "option", "ID", "Yes"

    ' Six months back, the same window his own macro asks for.
    EnterData drv, Keys, "//input[@id='" & VERVE_RPT_START & "']", _
              Format$(DateAdd("m", -6, Date), "mm/dd/yyyy"), "XPATH", True, True

    If CrownWaitFor(drv, By, "ID", VERVE_RPT_COMPANY, 30) Then
        SelectOption drv, VERVE_RPT_COMPANY, "option", "ID", "Trisura Insurance Company"
        drv.Wait 1000
    End If

    ClickElement drv, VERVE_RPT_DETAIL, "ID"
    ClickElement drv, VERVE_RPT_RUN, "ID"

    If Not CrownWaitFor(drv, By, "ID", VERVE_RPT_EXPORT, 90) Then
        note = "the report did not finish running - statuses are as the billing line reads them"
        GoTo BackToTheSite
    End If

    ' Noted before asking, so only a file that arrives after this counts.
    asked = Now

    ClickElement drv, VERVE_RPT_EXPORT, "ID"
    drv.Wait 1000
    drv.FindElementByLinkText("CSV (comma delimited)").Click

    For tries = 1 To 60
        path = CrownNewestCsvSince(asked)
        If Len(path) > 0 Then Exit For
        drv.Wait 500
        DoEvents
    Next tries

    If Len(path) = 0 Then
        note = "the report never arrived in Downloads - statuses are as the billing line reads them"
        GoTo BackToTheSite
    End If

    Set map = CrownCsvLookup(path, "polic", "status", heading, found)
    Set CrownVerveStatuses = map

    If found = 0 Then
        note = "the report came down but no column in it could be read as a policy number" & vbCrLf & _
               "with a status beside it. Nothing was written to any policy's status." & vbCrLf & _
               "Its heading row was: " & Left$(heading, 240)
    Else
        note = found & " statuses read from the Amount Due report."
    End If

BackToTheSite:
    ' Back to a signed-in page. The report has taken the window somewhere
    ' the quick policy lookup does not exist.
    drv.ExecuteScript "window.open(arguments[0])", wsInput.Range("URL_9").value
    drv.SwitchToNextWindow

    If CrownWaitFor(drv, By, "ID", "LoginControl_LoginNameTextBox", 20) Then
        EnterData drv, Keys, "LoginControl_LoginNameTextBox", wsInput.Range("USER_2").value, "ID"
        EnterData drv, Keys, "LoginControl_PasswordTextBox", wsInput.Range("PASS_2").value, "ID"
        ClickElement drv, "LoginControl_LoginLinkButton", "ID"
    End If

    Err.Clear
End Function
