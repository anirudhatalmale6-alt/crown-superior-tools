Attribute VB_Name = "CrownUpdate"
Option Explicit

' =====================================================================
' Crown Superior - put a whole update in with one macro
' ---------------------------------------------------------------------
' Every change so far meant removing and importing modules by hand, one
' at a time. This does it for you: it takes every .bas file in the
' unzipped CrownUpdate folder (Desktop or Downloads) and puts it in this workbook,
' replacing the module of the same name.
'
' Before replacing anything it EXPORTS the module that is there now into
' a backup folder inside CrownUpdate, so every change can be undone.
'
' ONE-TIME SETTING (Excel only allows this with it switched on):
'   File > Options > Trust Center > Trust Center Settings... >
'   Macro Settings > tick "Trust access to the VBA project object model"
'   > OK > OK.
' =====================================================================

' Modules that are only old leftovers and are taken out completely.
Private Function CrownRetired() As Variant
    CrownRetired = Array("Module2")
End Function


' Where the unzipped files are. Windows' Extract All puts them in a folder
' named after the zip, next to it - Desktop or Downloads - and sometimes one
' folder deeper. The first of these that has .bas files in it wins.
Private Function CrownUpdateFolder() As String
    Dim home As String, places As Variant, i As Long

    home = Environ$("USERPROFILE")
    places = Array(home & "\Desktop\CrownUpdate\", home & "\Desktop\CrownUpdate\CrownUpdate\", _
                   home & "\Downloads\CrownUpdate\", home & "\Downloads\CrownUpdate\CrownUpdate\", _
                   home & "\OneDrive\Desktop\CrownUpdate\", home & "\OneDrive\Desktop\CrownUpdate\CrownUpdate\")

    On Error Resume Next
    For i = LBound(places) To UBound(places)
        If Len(Dir$(places(i) & "*.bas")) > 0 Then
            CrownUpdateFolder = places(i)

            Exit Function
        End If
    Next i
End Function


Public Sub CrownInstallUpdate()
    Dim folder As String, backup As String
    Dim file As String, name As String
    Dim project As Object, comp As Object
    Dim done As Long, removed As Long
    Dim report As String
    Dim i As Long
    Dim retired As Variant

    folder = CrownUpdateFolder()

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
               "Then run CrownInstallUpdate again.", vbExclamation, "Crown Superior"

        Exit Sub
    End If
    On Error GoTo 0

    ' Automatic quoting is stopped first - its timer belongs to a module
    ' about to be replaced.
    On Error Resume Next
    Application.Run "CrownAutoQuietStop"
    On Error GoTo 0

    backup = folder & "backup " & Format$(Now, "yyyy-mm-dd hh.nn") & "\"
    MkDir backup

    file = Dir$(folder & "*.bas")

    Do While Len(file) > 0
        name = Left$(file, Len(file) - 4)

        If StrComp(name, "CrownUpdate", vbTextCompare) <> 0 Then
            Set comp = Nothing
            On Error Resume Next
            Set comp = project.VBComponents(name)
            On Error GoTo 0

            If Not comp Is Nothing Then
                comp.Export backup & name & ".bas"

                ' Renamed before removing: Excel only really removes a module
                ' when the macro ends, and the new one would otherwise
                ' arrive as "name1".
                comp.Name = name & "_old" & Format$(Now, "nnss")
                project.VBComponents.Remove comp
            End If

            project.VBComponents.Import folder & file
            done = done + 1
            report = report & "   " & name & vbCrLf
        End If

        file = Dir$()
    Loop

    retired = CrownRetired()

    For i = LBound(retired) To UBound(retired)
        Set comp = Nothing
        On Error Resume Next
        Set comp = project.VBComponents(CStr(retired(i)))
        On Error GoTo 0

        If Not comp Is Nothing Then
            comp.Export backup & CStr(retired(i)) & ".bas"
            comp.Name = CStr(retired(i)) & "_old" & Format$(Now, "nnss")
            project.VBComponents.Remove comp
            removed = removed + 1
            report = report & "   " & CStr(retired(i)) & " (old leftover, taken out)" & vbCrLf
        End If
    Next i

    ThisWorkbook.Save

    MsgBox done + removed & " module(s) updated and the workbook saved:" & vbCrLf & vbCrLf & _
           report & vbCrLf & _
           "The old versions are kept in " & backup & vbCrLf & vbCrLf & _
           "Now run: Alt+F8 > CrownStartAutoQuoting > Run.", vbInformation, "Crown Superior"
End Sub
