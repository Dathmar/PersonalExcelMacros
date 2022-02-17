Attribute VB_Name = "AutoUpdate_Main"
Option Explicit
Const ADDINNAME = "General_Purpose_Macros"
Const ADDINDOWNLOADURL = "https://github.com/Dathmar/GeneralPurposeMacros/raw/main/General_Purpose_Macros.xlam"
Const LASTCOMMITURL = "https://api.github.com/repos/Dathmar/GeneralPurposeMacros/commits?path=General_Purpose_Macros.xlam&page=1&per_page=1"
Dim mcUpdate As AutoUpdates

Public Declare Function InternetGetConnectedState _
                         Lib "wininet.dll" (lpdwFlags As Long, _
                                            ByVal dwReserved As Long) As Boolean
Function IsConnected() As Boolean
    Dim Stat As Long
    IsConnected = (InternetGetConnectedState(Stat, 0&) <> 0)
End Function
Sub AutoUpdate()
    CheckAndUpdate False
End Sub
Sub ManualUpdate()
    On Error Resume Next
    Application.OnTime Now, "CheckAndUpdate"
End Sub
Public Sub CheckAndUpdate(Optional bManual As Boolean = True)
    Set mcUpdate = New AutoUpdates
    If bManual Then
        Application.Cursor = xlWait
    End If
    With mcUpdate
        'Set intial values of class
        'Name of this app, probably a global variable, such as GSAPPNAME
        .AppName = ADDINNAME
        'Get rid of possible old backup copy
        .RemoveOldCopy
        .DownloadName = ADDINDOWNLOADURL
        'Started check automatically or manually?
        .Manual = bManual
        'Check once a week
        If (Now - .LastUpdate >= 7) Or bManual Then
            .DoUpdate
        End If
    End With
TidyUp:
    On Error GoTo 0
    Exit Sub
End Sub

