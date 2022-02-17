VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AutoUpdates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public WithEvents sht As Worksheet
Attribute sht.VB_VarHelpID = -1

'-------------------------------------------------------------------------
' Module : clsUpdate
' Company   : JKP Application Development Services (c)
' Author : Jan Karel Pieterse
' Created   : 19-2-2007
' Purpose   : Class to check for program updates
'-------------------------------------------------------------------------

#If VBA7 Then
Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" _
        Alias "URLDownloadToFileA" (ByVal pCaller As Long, _
                                    ByVal szURL As String, ByVal szFileName As String, _
                                    ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
#Else
Private Declare Function URLDownloadToFile Lib "urlmon" _
        Alias "URLDownloadToFileA" (ByVal pCaller As Long, _
                                    ByVal szURL As String, ByVal szFileName As String, _
                                    ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
#End If


Private mdtLastUpdate As Date
Private msAppName As String
Private msBuild As String
Private msCheckURL As String
Private msCurrentAddinName As String
Private msDownloadName As String
Private msTempAddInName As String
Private mbManual As Boolean
Private Sub Class_Terminate()
    Set sht = Nothing
End Sub
Private Sub DownloadFile(strWebFilename As String, strSaveFileName As String)
' Download the file.
    URLDownloadToFile 0, strWebFilename, strSaveFileName, 0, 0
End Sub
Public Function IsThereAnUpdate(Optional bShowMsg As Boolean = False) As Boolean
    Dim sNewBuild As String
    On Error Resume Next
    SaveSetting AppName, "Updates", "LastUpdate", CStr(Int(Now))
    If err.number <> 0 And bShowMsg Then
        MsgBox "Error retrieving update information, please try again later.", vbInformation + vbOKOnly
    End If
End Function
Public Sub DoUpdate()
    If GetUpdate Then
        Application.Cursor = xlDefault
        MsgBox "Successfully updated the addins, please restart Excel to start using the new version!", vbOKOnly + vbInformation
    Else
        Application.Cursor = xlDefault
    End If
TidyUp:
    On Error GoTo 0
    Exit Sub
End Sub
Private Sub Sht_Change(ByVal Target As Range)
    Application.Cursor = xlDefault
    If Len(Target.Value) <= 4 Then
        DoUpdate
        Application.Cursor = xlDefault
    ElseIf Manual Then
        'Query failed to refresh and was called manually
        Application.Cursor = xlDefault
        MsgBox "Unable to retrieve version information, please try again later", vbInformation + vbOKOnly
    End If
    Set sht = Nothing
TidyUp:
    On Error GoTo 0
    Exit Sub
End Sub
Public Property Get Build() As String
    Build = msBuild
End Property
Public Property Let Build(ByVal sBuild As String)
    msBuild = GetSetting(AppName:=AppName, Section:="Updates", key:="LastUpdate", Default:=0)
End Property
Public Sub RemoveOldCopy()
    CurrentAddinName = ThisWorkbook.FullName
    TempAddInName = CurrentAddinName & "(OldVersion)"
    On Error Resume Next
    Kill TempAddInName
End Sub
Public Function GetUpdate() As Boolean
    On Error Resume Next
    'If workbook has been saved readonly, we can safely delete the file!
    If ThisWorkbook.ReadOnly Then
        err.Clear
        Kill CurrentAddinName
    End If
    LastUpdate = Now
    ThisWorkbook.SaveAs TempAddInName
    DoEvents
    Kill CurrentAddinName
    On Error GoTo 0
    DownloadFile DownloadName, CurrentAddinName
    If err = 0 Then GetUpdate = True
End Function
Private Property Get CurrentAddinName() As String
    CurrentAddinName = msCurrentAddinName
End Property
Private Property Let CurrentAddinName(ByVal sCurrentAddinName As String)
    msCurrentAddinName = sCurrentAddinName
End Property
Private Property Get TempAddInName() As String
    TempAddInName = msTempAddInName
End Property
Private Property Let TempAddInName(ByVal sTempAddInName As String)
    msTempAddInName = sTempAddInName
End Property
Public Property Get DownloadName() As String
    DownloadName = msDownloadName
End Property
Public Property Let DownloadName(ByVal sDownloadName As String)
    msDownloadName = sDownloadName
End Property
Public Property Get CheckURL() As String
    CheckURL = msCheckURL
End Property
Public Property Let CheckURL(ByVal sCheckURL As String)
    msCheckURL = sCheckURL
End Property
Public Property Get LastUpdate() As Date
    Dim dtNow As Date
    dtNow = Int(Now)
    mdtLastUpdate = CDate(GetSetting(AppName, "Updates", "LastUpdate", "0"))
    If mdtLastUpdate = 0 Then
        'Never checked for an update, save today!
        SaveSetting AppName, "Updates", "LastUpdate", CStr(Int(dtNow))
    End If
    LastUpdate = mdtLastUpdate
End Property
Public Property Let LastUpdate(ByVal dtLastUpdate As Date)
    mdtLastUpdate = dtLastUpdate
    SaveSetting AppName, "Updates", "LastUpdate", CStr(Int(mdtLastUpdate))
End Property
Public Property Get AppName() As String
    AppName = msAppName
End Property
Public Property Let AppName(ByVal sAppName As String)
    msAppName = sAppName
End Property
Public Property Get Manual() As Boolean
    Manual = mbManual
End Property
Public Property Let Manual(ByVal bManual As Boolean)
    mbManual = bManual
End Property



