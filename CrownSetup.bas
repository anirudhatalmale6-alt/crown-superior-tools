Attribute VB_Name = "CrownSetup"
Option Explicit

' =====================================================================
' Crown Superior - put a whole update in with one macro (second version)
' ---------------------------------------------------------------------
' The first installer removed each module and imported the new one. Excel
' only really removes a module when the macro ends, so it renamed the old
' one first - and a long name plus "_old1234" went over Excel's 31-letter
' limit for module names (error 50132, Mod12uaigScrapeQuotePayment).
'
' This one never removes or renames anything. It keeps each module and
' swaps the code inside it, so the name can be any length. A module that
' does not exist yet (CrownUnited) is added.
'
' The modules come as .txt files in the unzipped CrownUpdate folder.
' Before changing a module it EXPORTS the one there now into a backup
' folder inside CrownUpdate, so every change can be undone.
'
' Same one-time setting as before (already done if the first one ran):
'   File > Options > Trust Center > Trust Center Settings... >
'   Macro Settings > tick "Trust access to the VBA project object model".
' =====================================================================

' Modules that are only old leftovers and are taken out completely.
Private Function CrownSetupRetired() As Variant
    CrownSetupRetired = Array("Module2")
End Function


' Where the unzipped files are: the first of these with .txt files in it.
Private Function CrownSetupFolder() As String
    Dim home As String, places As Variant, i As Long

    home = Environ$("USERPROFILE")
    places = Array(home & "\Desktop\CrownUpdate\", home & "\Desktop\CrownUpdate\CrownUpdate\", _
                   home & "\Downloads\CrownUpdate\", home & "\Downloads\CrownUpdate\CrownUpdate\", _
                   home & "\OneDrive\Desktop\CrownUpdate\", home & "\OneDrive\Desktop\CrownUpdate\CrownUpdate\")

    On Error Resume Next
    For i = LBound(places) To UBound(places)
        If Len(Dir$(places(i) & "CrownSetup.bas")) > 0 And Len(Dir$(places(i) & "*.txt")) > 0 Then
            CrownSetupFolder = places(i)

            Exit Function
        End If
    Next i
End Function


' A file's text, as Windows reads it, with the "Attribute ..." header
' lines taken off - those are for importing, and as code they are an
' error.
Private Function CrownSetupCode(ByVal path As String, ByRef moduleName As String) As String
    Dim f As Integer
    Dim bytes() As Byte
    Dim text As String
    Dim lines As Variant
    Dim i As Long, n As Long
    Dim kept() As String

    f = FreeFile
    Open path For Binary Access Read As #f
    If LOF(f) = 0 Then
        Close #f

        Exit Function
    End If
    ReDim bytes(0 To LOF(f) - 1)
    Get #f, , bytes
    Close #f

    text = StrConv(bytes, vbUnicode)
    text = Replace(text, vbCrLf, vbLf)
    text = Replace(text, vbCr, vbLf)
    lines = Split(text, vbLf)
    ReDim kept(0 To UBound(lines))

    For i = LBound(lines) To UBound(lines)
        If Left$(lines(i), 10) = "Attribute " Then
            If InStr(1, lines(i), "VB_Name = """, vbBinaryCompare) > 0 Then
                moduleName = Split(lines(i), """")(1)
            End If
        Else
            kept(n) = lines(i)
            n = n + 1
        End If
    Next i

    If n = 0 Then Exit Function

    ReDim Preserve kept(0 To n - 1)
    CrownSetupCode = Join(kept, vbCrLf)
End Function


' True once the old CrownUpdate module has been swapped for the few lines
' that send Alt+F8 > CrownInstallUpdate here.
Private Function CrownSetupIsForwarder(ByVal project As Object) As Boolean
    Dim comp As Object

    On Error Resume Next
    Set comp = project.VBComponents("CrownUpdater")
    If comp Is Nothing Then Set comp = project.VBComponents("CrownUpdate")
    If comp Is Nothing Then Exit Function

    CrownSetupIsForwarder = (InStr(1, comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines), _
                                   "CrownSetupUpdate", vbTextCompare) > 0)
End Function


Public Sub CrownSetupUpdate()
    Dim folder As String, backup As String
    Dim file As String, name As String, inFile As String, code As String
    Dim project As Object, comp As Object
    Dim done As Long, added As Long, removed As Long
    Dim report As String, problems As String
    Dim i As Long
    Dim retired As Variant

    folder = CrownSetupFolder()

    If Len(folder) = 0 Then
        MsgBox "I could not find the unzipped CrownUpdate folder." & vbCrLf & vbCrLf & _
               "Right-click the CrownUpdate.zip file I sent > Extract All > Extract. " & _
               "Leave it in Downloads or on the Desktop, then run this again.", _
               vbExclamation, "Crown Superior"

        Exit Sub
    End If

    On Error Resume Next
    Set project = ThisWorkbook.VBProject
    If Not project Is Nothing Then i = project.VBComponents.Count
    If Err.Number <> 0 Or project Is Nothing Then
        On Error GoTo 0
        MsgBox "Excel is not allowing macros to update other macros yet." & vbCrLf & vbCrLf & _
               "Switch this on once:" & vbCrLf & _
               "File > Options > Trust Center > Trust Center Settings... > Macro Settings >" & vbCrLf & _
               "tick ""Trust access to the VBA project object model"" > OK > OK." & vbCrLf & vbCrLf & _
               "Then run CrownSetupUpdate again.", vbExclamation, "Crown Superior"

        Exit Sub
    End If
    On Error GoTo 0

    ' Automatic quoting is stopped first - its timer belongs to a module
    ' about to change.
    On Error Resume Next
    Application.Run "CrownAutoQuietStop"
    On Error GoTo 0

    backup = folder & "backup " & Format$(Now, "yyyy-mm-dd hh.nn") & "\"
    On Error Resume Next
    MkDir backup
    On Error GoTo 0

    file = Dir$(folder & "*.txt")

    Do While Len(file) > 0
        name = Left$(file, Len(file) - 4)
        inFile = ""
        code = CrownSetupCode(folder & file, inFile)

        If StrComp(name, "CrownSetup", vbTextCompare) = 0 Then
            ' Itself: a running macro cannot rewrite its own module.
        ElseIf (StrComp(name, "CrownUpdate", vbTextCompare) = 0 Or StrComp(name, "CrownUpdater", vbTextCompare) = 0) _
               And CrownSetupIsForwarder(project) Then
            ' Already just the forwarder - and this may be running from it.
        ElseIf Len(inFile) = 0 Then
            problems = problems & "   " & file & " is not a module - skipped" & vbCrLf
        ElseIf Len(code) = 0 Then
            problems = problems & "   " & file & " is empty - skipped" & vbCrLf
        ElseIf StrComp(inFile, name, vbTextCompare) <> 0 Then
            problems = problems & "   " & file & " says it is " & inFile & " - skipped" & vbCrLf
        Else
            Set comp = Nothing
            On Error Resume Next
            Set comp = project.VBComponents(name)
            On Error GoTo 0

            If comp Is Nothing Then
                Set comp = project.VBComponents.Add(1)   ' a standard module
                comp.Name = name
                added = added + 1
                report = report & "   " & name & " (new)" & vbCrLf
            Else
                comp.Export backup & name & ".bas"
                report = report & "   " & name & vbCrLf
            End If

            ' Everything out - including an "Option Explicit" Excel may
            ' have put in a new module itself, which would otherwise be
            ' there twice - then the new code in.
            With comp.CodeModule
                If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
                .AddFromString code
            End With

            done = done + 1
        End If

        file = Dir$()
    Loop

    retired = CrownSetupRetired()

    For i = LBound(retired) To UBound(retired)
        Set comp = Nothing
        On Error Resume Next
        Set comp = project.VBComponents(CStr(retired(i)))
        On Error GoTo 0

        If Not comp Is Nothing Then
            comp.Export backup & CStr(retired(i)) & ".bas"
            project.VBComponents.Remove comp
            removed = removed + 1
            report = report & "   " & CStr(retired(i)) & " (old leftover, taken out)" & vbCrLf
        End If
    Next i

    ThisWorkbook.Save

    MsgBox done + removed & " module(s) updated and the workbook saved:" & vbCrLf & vbCrLf & _
           report & _
           IIf(Len(problems) > 0, vbCrLf & "Not changed:" & vbCrLf & problems, "") & vbCrLf & _
           "The old versions are kept in " & backup & vbCrLf & vbCrLf & _
           "Now run: Alt+F8 > CrownStartAutoQuoting > Run.", vbInformation, "Crown Superior"
End Sub
