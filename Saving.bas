Attribute VB_Name = "Saving"
Sub TestNameStatement()
  Dim fOK As Boolean
 
  ' Folders must exist for Source, but do not need to exist _
  ' for destination
With ActiveSheet
For n = 2 To .UsedRange.Rows.Count
  fOK = RenameFileOrDir(.Cells(n, 1).Value, _
        .Cells(n, 2).Value)
    .Cells(n, 4) = fOK
Next n
End With
Application.ScreenUpdating = True
End Sub
Function RenameFileOrDir( _
  ByVal strSource As String, _
  ByVal strTarget As String, _
  Optional fOverwriteTarget As Boolean = False) As Boolean
 
  On Error GoTo PROC_ERR
 
  Dim fRenameOK As Boolean
  Dim fRemoveTarget As Boolean
  Dim strFirstDrive As String
  Dim strSecondDrive As String
  Dim fOK As Boolean
 
  If Not ((Len(strSource) = 0) Or _
         (Len(strTarget) = 0) Or _
         (Not (FileOrDirExists(strSource)))) Then
 
    ' Check if the target exists
    If FileOrDirExists(strTarget) Then
 
      If fOverwriteTarget Then
        fRemoveTarget = True
      Else
        If vbYes = MsgBox("Do you wish to overwrite the " & _
           "target file?", vbExclamation + vbYesNo, _
           "Overwrite confirmation") Then
          fRemoveTarget = True
        End If
      End If
 
      If fRemoveTarget Then
        ' Check that it's not a directory
        If ((GetAttr(strTarget) And vbDirectory)) <> _
           vbDirectory Then
          Kill strTarget
          fRenameOK = True
        Else
          MsgBox "Cannot overwrite a directory", vbOKOnly, _
            "Cannot perform operation"
          'FUTURE CODE FOR DIRECTORIES
        End If
      End If
    Else
      ' The target does not exist
      ' Check if source is a directory
      If ((GetAttr(strSource) And vbDirectory) = _
           vbDirectory) Then
        ' Source is a directory, see if drives are the same
        strFirstDrive = Left(strSource, InStr(strSource, ":\"))
        strSecondDrive = Left(strTarget, InStr(strTarget, ":\"))
        If strFirstDrive = strSecondDrive Then
          fRenameOK = True
        Else
          MsgBox "Cannot rename directories across drives", _
            vbOKOnly, "Cannot perform operation"
          'FUTURE CODE FOR DIRECTORIES ON DIFFERENT DRIVES
        End If
      Else
        'It's a file, ok to proceed
        fRenameOK = True
      End If
    End If
 
    If fRenameOK Then
      Name strSource As strTarget
      fOK = True
    End If
  End If
 
  RenameFileOrDir = fOK
 
PROC_EXIT:
  Exit Function
 
PROC_ERR:
  MsgBox "Error: " & err.number & ". " & err.Description, , _
         "RenameFileOrDir"
  Resume PROC_EXIT
End Function
Function FileOrDirExists(strDest As String) As Boolean
  Dim intLen As Integer
  Dim fReturn As Boolean

  fReturn = False

  If strDest <> vbNullString Then
    On Error Resume Next
    intLen = Len(Dir$(strDest, vbDirectory + vbNormal))
    On Error GoTo PROC_ERR
    fReturn = (Not err And intLen > 0)
  End If

PROC_EXIT:
  FileOrDirExists = fReturn
  Exit Function

PROC_ERR:
  MsgBox "Error: " & err.number & ". " & err.Description, , _
         "FileOrDirExists"
  Resume PROC_EXIT
End Function
Sub csv_to_xls()
Dim sht As Worksheet
Dim csv As Workbook
Dim csv_sht As Worksheet
Dim FSO As FileSystemObject

Set sht = ActiveSheet

Set FSO = New FileSystemObject

For n = 2 To sht.UsedRange.Rows.Count
    Application.Workbooks.Open filename:=sht.Cells(n, 1), ReadOnly:=True, delimiter:=","
    Set csv = ActiveWorkbook
    Set csv_sht = csv.ActiveSheet
    
    If Application.WorksheetFunction.CountA(csv_sht.Rows(1)) = 1 Then
        sht.Cells(n, 4) = False
    Else
        csv_sht.Cells(1, csv_sht.UsedRange.Columns.Count + 1) = "Test Form Name"
        csv_sht.Range(csv_sht.Cells(2, csv_sht.UsedRange.Columns.Count), csv_sht.Cells(csv_sht.UsedRange.Rows.Count, csv_sht.UsedRange.Columns.Count)) = sht.Cells(n, 2)
        sht.Cells(n, 4) = True
        
        If Not FSO.FolderExists(csv.path & "\xls") Then
            MkDir csv.path & "\xls"
        End If
        
        csv.SaveAs filename:=csv.path & "\xls\" & sht.Cells(n, 2) & ".xlsx", FileFormat:=xlOpenXMLWorkbook
    End If
        
        
    csv.Close
Next n
End Sub
