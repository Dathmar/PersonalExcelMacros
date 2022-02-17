Attribute VB_Name = "AutoUpdates_DownloadFile"
Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" _
        Alias "URLDownloadToFileA" (ByVal pCaller As Long, _
                                    ByVal szURL As String, ByVal szFileName As String, _
                                    ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Private Sub DownloadFile(strWebFilename As String, strSaveFileName As String)
' Download the file.
    URLDownloadToFile 0, strWebFilename, strSaveFileName, 0, 0
End Sub
Sub test_download()
Dim DownloadName As String
Dim CurrentAddinName As String
DownloadName = "https://github.com/Dathmar/GeneralPurposeMacros/raw/main/General_Purpose_Macros.xlam"
CurrentAddinName = "C:\blah\blah\gpm.xlam"
DownloadFile DownloadName, CurrentAddinName
End Sub
