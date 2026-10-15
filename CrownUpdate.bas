Attribute VB_Name = "CrownUpdate"
Option Explicit

' =====================================================================
' Crown Superior - updates
' ---------------------------------------------------------------------
' The installer now lives in the CrownSetup module. This is kept so that
' Alt+F8 > CrownInstallUpdate > Run keeps working for every update.
' =====================================================================

Public Sub CrownInstallUpdate()
    CrownSetupUpdate
End Sub
