Attribute VB_Name = "CrownChrome"
Option Explicit

' =====================================================================
' Crown Superior - keep ChromeDriver matched to Chrome
' ---------------------------------------------------------------------
' Every run drives Chrome through chromedriver.exe, and Chrome updates
' itself every few weeks. The day it does, the old chromedriver.exe no
' longer fits and every run stops on its first step (drv.Get) - as the
' United payment run did on 20 Oct.
'
' Alt+F8 > CrownUpdateChromeDriver > Run
'   reads which Chrome is installed, downloads the chromedriver.exe made
'   for exactly that version (Google's own "Chrome for Testing" site) and
'   puts it where SeleniumBasic looks for it. The old one is kept as
'   chromedriver.old.exe.
'
' CrownChromeOk(drv) - called right after drv.Start by the Crown runs: if
'   Chrome did not start it says so in words, instead of stopping on a
'   yellow line with no message.
' =====================================================================

Private Const CFT_URL As String = _
    "https://googlechromelabs.github.io/chrome-for-testing/latest-patch-versions-per-build-with-downloads.json"


Public Function CrownChromeOk(ByVal drv As Object) As Boolean
    Dim why As String

    On Error Resume Next
    Err.Clear
    drv.Get "about:blank"
    why = Err.Description
    If Err.Number = 0 Then CrownChromeOk = True
    Err.Clear
    On Error GoTo 0

    If CrownChromeOk Then Exit Function

    MsgBox "Chrome did not start for the tool." & vbCrLf & vbCrLf & _
           "What it said: " & Left$(why, 400) & vbCrLf & vbCrLf & _
           "This almost always means Chrome updated itself and chromedriver is now the old one." & vbCrLf & _
           "Fix: close all Chrome windows, then Alt+F8 > CrownUpdateChromeDriver > Run, then run this again.", _
           vbExclamation, "Crown Superior"
End Function


Public Sub CrownUpdateChromeDriver()
    Dim chromeVer As String, build As String, json As String, url As String
    Dim folder As String, target As String, temp As String, zipPath As String
    Dim at As Long, stopAt As Long, oldVer As String

    chromeVer = CrownChromeVersion()

    If Len(chromeVer) = 0 Then
        MsgBox "I could not find which Chrome version is installed.", vbExclamation, "Crown Superior"
        Exit Sub
    End If

    ' "142.0.7444.59" -> build "142.0.7444"
    build = Left$(chromeVer, InStrRev(chromeVer, ".") - 1)

    folder = CrownDriverFolder()

    If Len(folder) = 0 Then
        MsgBox "I could not find the SeleniumBasic folder (where chromedriver.exe lives).", vbExclamation, "Crown Superior"
        Exit Sub
    End If

    target = folder & "chromedriver.exe"
    oldVer = CrownFileVersionText(target)

    json = CrownHttpText(CFT_URL)

    If Len(json) = 0 Then
        MsgBox "Could not reach Google's Chrome for Testing site. Check the internet and try again.", vbExclamation, "Crown Superior"
        Exit Sub
    End If

    ' The block for this build, then the win32 chromedriver link inside it.
    at = InStr(1, json, """" & build & """:", vbBinaryCompare)

    If at = 0 Then
        MsgBox "Google does not list a chromedriver for Chrome " & chromeVer & " yet. Try again tomorrow.", vbExclamation, "Crown Superior"
        Exit Sub
    End If

    stopAt = InStr(at + 5, json, "},""", vbBinaryCompare)   ' start of the next build's block
    If stopAt = 0 Then stopAt = Len(json)

    at = InStr(at, json, """chromedriver""", vbBinaryCompare)
    If at > 0 And at < stopAt Then at = InStr(at, json, "chromedriver-win32.zip", vbBinaryCompare)

    If at = 0 Or at > stopAt Then
        MsgBox "Google's list has no Windows chromedriver for Chrome " & chromeVer & ".", vbExclamation, "Crown Superior"
        Exit Sub
    End If

    url = Mid$(json, InStrRev(json, """", at) + 1, at + Len("chromedriver-win32.zip") - InStrRev(json, """", at) - 1)

    temp = Environ$("TEMP") & "\crown-chromedriver\"
    On Error Resume Next
    MkDir temp
    Kill temp & "*.*"
    On Error GoTo 0

    zipPath = temp & "chromedriver.zip"

    If Not CrownHttpSave(url, zipPath) Then
        MsgBox "The download did not work:" & vbCrLf & url, vbExclamation, "Crown Superior"
        Exit Sub
    End If

    ' Unzip with Windows itself.
    Dim shellApp As Object, items As Object
    Set shellApp = CreateObject("Shell.Application")
    Set items = shellApp.Namespace(CVar(zipPath)).Items
    shellApp.Namespace(CVar(temp)).CopyHere items, 4 + 16
    Application.Wait Now + TimeSerial(0, 0, 3)

    If Len(Dir$(temp & "chromedriver-win32\chromedriver.exe")) = 0 Then
        MsgBox "The download did not contain chromedriver.exe.", vbExclamation, "Crown Superior"
        Exit Sub
    End If

    ' A running chromedriver keeps its file locked.
    CreateObject("WScript.Shell").Run "taskkill /f /im chromedriver.exe", 0, True

    On Error Resume Next
    If Len(Dir$(folder & "chromedriver.old.exe")) > 0 Then Kill folder & "chromedriver.old.exe"
    If Len(Dir$(target)) > 0 Then Name target As folder & "chromedriver.old.exe"
    Err.Clear
    FileCopy temp & "chromedriver-win32\chromedriver.exe", target

    If Err.Number <> 0 Then
        On Error GoTo 0
        MsgBox "Windows would not let me replace " & target & "." & vbCrLf & vbCrLf & _
               "Close Excel, open it again and run CrownUpdateChromeDriver once more " & _
               "(or copy " & temp & "chromedriver-win32\chromedriver.exe there by hand).", vbExclamation, "Crown Superior"
        Exit Sub
    End If
    On Error GoTo 0

    MsgBox "Done. Chrome is " & chromeVer & " and chromedriver now matches it." & vbCrLf & vbCrLf & _
           "Replaced: " & target & IIf(Len(oldVer) > 0, " (was " & oldVer & ")", "") & vbCrLf & _
           "The old one is kept as chromedriver.old.exe." & vbCrLf & vbCrLf & _
           "You can run the tool again now.", vbInformation, "Crown Superior"
End Sub


' Installed Chrome, e.g. "142.0.7444.59".
Private Function CrownChromeVersion() As String
    Dim sh As Object, v As String, paths As Variant, i As Long

    Set sh = CreateObject("WScript.Shell")

    On Error Resume Next
    v = sh.RegRead("HKCU\Software\Google\Chrome\BLBeacon\version")
    If Len(v) = 0 Then v = sh.RegRead("HKLM\Software\Google\Chrome\BLBeacon\version")
    On Error GoTo 0

    If Len(v) = 0 Then
        paths = Array(Environ$("ProgramFiles") & "\Google\Chrome\Application\chrome.exe", _
                      Environ$("ProgramFiles(x86)") & "\Google\Chrome\Application\chrome.exe", _
                      Environ$("LOCALAPPDATA") & "\Google\Chrome\Application\chrome.exe")

        For i = LBound(paths) To UBound(paths)
            v = CrownFileVersionText(CStr(paths(i)))
            If Len(v) > 0 Then Exit For
        Next i
    End If

    CrownChromeVersion = Trim$(v)
End Function

' Where SeleniumBasic keeps chromedriver.exe (the first place that has it).
Private Function CrownDriverFolder() As String
    Dim places As Variant, i As Long

    places = Array(Environ$("LOCALAPPDATA") & "\SeleniumBasic\", _
                   Environ$("ProgramFiles") & "\SeleniumBasic\", _
                   Environ$("ProgramFiles(x86)") & "\SeleniumBasic\")

    For i = LBound(places) To UBound(places)
        If Len(Dir$(places(i) & "chromedriver.exe")) > 0 Then
            CrownDriverFolder = places(i)
            Exit Function
        End If
    Next i
End Function

Private Function CrownFileVersionText(ByVal path As String) As String
    On Error Resume Next
    If Len(Dir$(path)) = 0 Then Exit Function
    CrownFileVersionText = CreateObject("Scripting.FileSystemObject").GetFileVersion(path)
End Function

Private Function CrownHttpText(ByVal url As String) As String
    Dim http As Object

    On Error Resume Next
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", url, False
    http.setTimeouts 15000, 15000, 30000, 60000
    http.send
    If Err.Number = 0 And http.Status = 200 Then CrownHttpText = http.responseText
End Function

Private Function CrownHttpSave(ByVal url As String, ByVal path As String) As Boolean
    Dim http As Object, stream As Object

    On Error Resume Next
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", url, False
    http.setTimeouts 15000, 15000, 30000, 120000
    http.send

    If Err.Number <> 0 Or http.Status <> 200 Then Exit Function

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write http.responseBody
    stream.SaveToFile path, 2
    stream.Close

    CrownHttpSave = (Err.Number = 0 And Len(Dir$(path)) > 0)
End Function
