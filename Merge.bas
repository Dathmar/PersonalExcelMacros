Attribute VB_Name = "Merge"
Sub Merge_Files_by_Headers()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       03/14/2012                                        '''
'''The purpose is to merge all files that are selected to a new sheet.                      '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim from_book As Workbook
Dim to_book As Workbook
Dim xl_file_name As Variant
Dim this_workbook As Long
Dim sht_name As String
Dim n As Long
Dim i As Long
Dim to_sht As Worksheet
Dim from_sht As Worksheet
Dim to_lr As Long
Dim wb_headers As String
Dim this_sht As Long


xl_file_name = Application.GetOpenFilename("Excel files (*.*),*.*", , _
    "Browse for file to be merged", MultiSelect:=True)
Application.DisplayAlerts = False
Application.ScreenUpdating = False

If IsArray(xl_file_name) Then
    Set to_book = ActiveWorkbook
    Set to_sht = to_book.ActiveSheet
    
    For this_workbook = LBound(xl_file_name) To UBound(xl_file_name) ' iterate through each book
        Application.Workbooks.Open filename:=xl_file_name(this_workbook), ReadOnly:=True ' open the books
        Set from_book = ActiveWorkbook
        For this_sht = 1 To from_book.Sheets.Count
            wb_headers = ""
            Set from_sht = from_book.Sheets(this_sht)
        
            to_lr = to_sht.UsedRange.Rows.Count + 1
            With from_sht
            For n = 1 To .UsedRange.Columns.Count

                'If InStr(wb_headers, .Cells(1, n)) <> 0 Then
                '    .Cells(1, n) = .Cells(1, n) & "2"
                'End If
                wb_headers = wb_headers & "," & .Cells(1, n)
                i = get_column_by_header(to_sht, .Cells(1, n), False)
                
                If i <> 0 Then
                    .Range(.Cells(1, n), .Cells(.UsedRange.Rows.Count, n)).Copy _
                        Destination:=to_sht.Cells(to_lr, i)
                End If
            Next n
            End With
            
            i = get_column_by_header(to_sht, "Sheet Name")
            If i <> 0 Then
                to_lr = to_sht.Cells(to_sht.Rows.Count, i).End(xlUp).Row + 1
                to_sht.Range(to_sht.Cells(to_lr, i), to_sht.Cells(to_sht.UsedRange.Rows.Count, i)) = from_sht.Name
            End If
            
        Next this_sht
        
        i = get_column_by_header(to_sht, "File Name")
        If i <> 0 Then
            to_lr = to_sht.Cells(to_sht.Rows.Count, i).End(xlUp).Row + 1
            to_sht.Range(to_sht.Cells(to_lr, i), to_sht.Cells(to_sht.UsedRange.Rows.Count, i)) = from_book.Name
        End If
        from_book.Close
    Next this_workbook
End If
Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub
Sub Merge_File_Headers()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       03/14/2012                                        '''
'''The purpose is to merge all files that are selected to a new sheet.                      '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim from_book As Workbook
Dim to_book As Workbook
Dim xl_file_name As Variant
Dim this_workbook As Long
Dim sht_name As String
Dim n As Long
Dim i As Long
Dim to_sht As Worksheet
Dim from_sht As Worksheet
Dim to_lr As Long
Dim include_sht_name As Boolean
Dim include_file_name As Boolean

include_file_name = True
include_sht_name = True

xl_file_name = Application.GetOpenFilename("Excel files (*.*),*.*", , _
    "Browse for file to be merged", MultiSelect:=True)
Application.DisplayAlerts = False
Application.ScreenUpdating = False
to_lr = 1
If IsArray(xl_file_name) Then
    Set to_book = ActiveWorkbook
    Set to_sht = to_book.ActiveSheet
    
    For this_workbook = LBound(xl_file_name) To UBound(xl_file_name) ' iterate through each book
        Application.Workbooks.Open filename:=xl_file_name(this_workbook), ReadOnly:=True ' open the books
        Set from_book = ActiveWorkbook
        
        For i = 1 To from_book.Sheets.Count
            Set from_sht = from_book.Sheets(i)
            from_sht.Rows(1).Copy Destination:=to_sht.Rows(to_lr)
            
            If include_sht_name Then
                to_sht.Cells(to_lr, 1).Insert Shift:=xlToRight
                to_sht.Cells(to_lr, 1) = from_sht.Name
            End If
            
            If include_file_name Then
                to_sht.Cells(to_lr, 1).Insert Shift:=xlToRight
                to_sht.Cells(to_lr, 1) = from_book.Name
            End If
            
            to_lr = to_lr + 1
        Next i
        
        from_book.Close
    Next this_workbook
End If
Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub
Sub Remove_Spaces_from_Headers()
Dim book As Workbook
Dim xl_file_name As Variant
Dim n As Long
Dim i As Long
Dim sht As Worksheet
Dim lc As Long

xl_file_name = Application.GetOpenFilename("Excel files (*.*),*.*", , _
    "Browse for file to be merged", MultiSelect:=True)
Application.DisplayAlerts = False
Application.ScreenUpdating = False
to_lr = 1
If IsArray(xl_file_name) Then
    For this_workbook = LBound(xl_file_name) To UBound(xl_file_name) ' iterate through each book
        Application.Workbooks.Open filename:=xl_file_name(this_workbook), ReadOnly:=False ' open the books
        Set book = ActiveWorkbook
        
        For i = 1 To book.Sheets.Count
            Set sht = book.Sheets(i)
            lc = sht.Cells(1, sht.Columns.Count).End(xlToLeft).Column
            For n = 1 To lc
                sht.Cells(1, n) = UCase(Replace(Trim(sht.Cells(1, n)), " ", "_"))
            Next n
        Next i
        
        from_book.Close savechanges:=True
    Next this_workbook
End If
Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub
Sub Merge_Sheet_By_Name()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       03/14/2012                                        '''
'''The purpose is to merge all files that are selected to a new sheet.                      '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim from_book As Workbook
Dim to_book As Workbook
Dim xl_file_name As Variant
Dim this_workbook As Long
Dim sht_name As String
Dim n As Long
Dim i As Long
Dim to_sht As Worksheet
Dim from_sht As Worksheet
Dim to_lr As Long

sht_name = "Forms in Package"

xl_file_name = Application.GetOpenFilename("Excel files (*.*),*.*", , _
    "Browse for file to be merged", MultiSelect:=True)
Application.DisplayAlerts = False
Application.ScreenUpdating = False
to_lr = 1
If IsArray(xl_file_name) Then
    Set to_book = ActiveWorkbook
    Set to_sht = to_book.ActiveSheet
    
    For this_workbook = LBound(xl_file_name) To UBound(xl_file_name) ' iterate through each book
        Application.Workbooks.Open filename:=xl_file_name(this_workbook), ReadOnly:=True ' open the books
        Set from_book = ActiveWorkbook
        Set from_sht = from_book.Sheets(sht_name)

        from_sht.Rows(1).Copy Destination:=to_sht.Rows(to_lr)
        to_lr = to_lr + 1
        
        from_book.Close
    Next this_workbook
End If
Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub

