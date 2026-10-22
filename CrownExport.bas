Attribute VB_Name = "CrownExport"
Option Explicit

' Writes this workbook's code, and the layout of Edit Data and Output,
' to Downloads\CrownModules.zip - for wiring new fields through macros
' that are only in this workbook (22 Oct: bodily injury / property
' damage through Finalize Policy and Send to Quote Form).
'
' No sheet DATA goes in it: no customer rows and nothing from the Input
' sheet (the carrier logins live there). Only code, headings, labels,
' formulas and which macro each button runs.
'
' Alt+F8 > CrownExportForMe > Run.
Public Sub CrownExportForMe()
    Dim folder As String, zipPath As String, comp As Object, ext As String
    Dim n As Long

    folder = Environ$("USERPROFILE") & "\Downloads\CrownModules\"
    zipPath = Environ$("USERPROFILE") & "\Downloads\CrownModules.zip"

    On Error Resume Next
    MkDir folder
    Kill folder & "*.*"
    Kill zipPath
    On Error GoTo 0

    On Error GoTo NoAccess
    For Each comp In ThisWorkbook.VBProject.VBComponents
        Select Case comp.Type
            Case 1: ext = ".bas"
            Case 2: ext = ".cls"
            Case 3: ext = ".frm"
            Case Else: ext = ".cls"
        End Select

        If comp.CodeModule.CountOfLines > 0 Then
            comp.Export folder & comp.Name & ext
            n = n + 1
        End If
    Next comp
    On Error GoTo 0

    CrownExportLayout folder & "EditData-layout.txt", "Edit Data", False
    CrownExportLayout folder & "Output-headings.txt", "Output", True
    CrownExportLayout folder & "CrownQuoteList-headings.txt", "Crown Quote List", True

    Shell "powershell -NoProfile -Command ""Compress-Archive -Path '" & folder & "*' -DestinationPath '" & zipPath & "' -Force""", vbHide
    Application.Wait Now + TimeValue("0:00:05")

    If Len(Dir$(zipPath)) > 0 Then
        MsgBox n & " code modules exported." & vbCrLf & vbCrLf & _
               "Please send me this file in the chat:" & vbCrLf & zipPath, vbInformation, "Crown Superior"
        Shell "explorer.exe /select,""" & zipPath & """", vbNormalFocus
    Else
        MsgBox n & " code modules exported to:" & vbCrLf & folder & vbCrLf & vbCrLf & _
               "Please zip that folder (right-click it > Send to > Compressed folder) and send it to me.", _
               vbInformation, "Crown Superior"
        Shell "explorer.exe """ & folder & """", vbNormalFocus
    End If
    Exit Sub

NoAccess:
    MsgBox "Excel would not let me read the code: " & Err.Description & vbCrLf & vbCrLf & _
           "File > Options > Trust Center > Trust Center Settings > Macro Settings > " & _
           "tick 'Trust access to the VBA project object model', then run this again.", vbExclamation, "Crown Superior"
End Sub

' One sheet's shape, not its contents. headingsOnly: row 1 only (the
' column names). Otherwise every formula, every label (bold or shaded
' text), every drop-down list and every button with the macro it runs.
Private Sub CrownExportLayout(ByVal path As String, ByVal sheetName As String, ByVal headingsOnly As Boolean)
    Dim ws As Worksheet, c As Range, f As Integer, shp As Shape, what As String, dv As String

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    f = FreeFile
    Open path For Output As #f
    Print #f, "Sheet: " & sheetName & "   used range " & ws.UsedRange.Address(False, False)

    If headingsOnly Then
        For Each c In ws.Range(ws.Cells(1, 1), ws.Cells(1, ws.UsedRange.Columns.Count + ws.UsedRange.Column - 1)).Cells
            If Len(CStr(c.Value)) > 0 Then Print #f, c.Address(False, False) & vbTab & c.Column & vbTab & CStr(c.Value)
        Next c
    Else
        For Each c In ws.UsedRange.Cells
            what = ""
            dv = ""

            On Error Resume Next
            dv = c.Validation.Formula1
            On Error GoTo 0

            If c.HasFormula Then
                what = "FORMULA" & vbTab & c.Formula
            ElseIf VarType(c.Value) = vbString And Len(c.Value) > 0 Then
                If (c.Font.Bold Or c.Interior.ColorIndex <> xlNone) And c.Locked Then what = "LABEL" & vbTab & CStr(c.Value)
            End If

            If Len(dv) > 0 Then what = what & vbTab & "LIST=" & dv
            If c.MergeCells And c.Address = c.MergeArea.Cells(1, 1).Address Then what = what & vbTab & "MERGED " & c.MergeArea.Address(False, False)
            If Len(what) > 0 Then Print #f, c.Address(False, False) & vbTab & what
        Next c

        For Each shp In ws.Shapes
            what = ""
            dv = ""
            On Error Resume Next
            what = shp.OnAction
            dv = Replace(shp.TextFrame2.TextRange.Text, vbLf, " ")
            On Error GoTo 0
            Print #f, "BUTTON" & vbTab & shp.Name & vbTab & shp.TopLeftCell.Address(False, False) & vbTab & _
                      dv & vbTab & "runs: " & what
        Next shp
    End If

    Close #f
End Sub
