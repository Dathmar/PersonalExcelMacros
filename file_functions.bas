Attribute VB_Name = "file_functions"
Sub convert_to_xlsx()
Dim wb As Workbook
xl_file_name = Application.GetOpenFilename("Excel files (*.csv),*.csv", , _
    "Browse for file to be merged", MultiSelect:=True)

If IsArray(xl_file_name) Then
    For this_file = LBound(xl_file_name) To UBound(xl_file_name)
        Application.Workbooks.Open xl_file_name(this_file)
        Set wb = ActiveWorkbook
        Run save_as_xlsx(wb, wb.path)
        wb.Close savechanges:=False
    Next this_file
End If
End Sub
Function save_as_xlsx(wb As Workbook, parent_path As String)
If Right(parent_path, 1) = "/" Or Right(parent_path, 1) = "/" Then
    parent_path = Left(parent_path, Len(parent_path) - 1)
End If

wb.SaveAs parent_path & "/" & Replace(wb.Name, get_extension(wb.Name), "") & "xlsx", XlFileFormat.xlOpenXMLWorkbook
End Function
Sub move_files()
Dim this_sht As Worksheet
Dim n As Long

Set this_sht = ActiveSheet

With this_sht
For n = 2 To .UsedRange.Rows.Count
    .Cells(n, 3) = Copy_File_to_Location(.Cells(n, 1), .Cells(n, 2), False)
Next n
End With
End Sub
Sub copy_files()
Dim this_sht As Worksheet
Dim n As Long

Set this_sht = ActiveSheet

With this_sht
For n = 2 To .UsedRange.Rows.Count
    .Cells(n, 3) = Copy_File_to_Location(.Cells(n, 1), .Cells(n, 2), True)
Next n
End With
End Sub

Function Copy_File_to_Location(from_file As String, to_file As String, Optional bCopy = True) As Boolean
Dim FSO As Object
Dim path As String
Set FSO = CreateObject("scripting.filesystemobject")

path = Replace(to_file, gpm.get_filename(to_file), "")

If Len(Dir(path, vbDirectory)) = 0 Then
    MkDir Trim(path)
End If
On Error GoTo Er
If bCopy Then
    FSO.CopyFile Source:=from_file, Destination:=to_file
Else
    FSO.MoveFile Source:=from_file, Destination:=to_file
End If
On Error GoTo 0
Set FSO = Nothing
Copy_File_to_Location = True
Exit Function
Er:
Copy_File_to_Location = False
Set FSO = Nothing
End Function

