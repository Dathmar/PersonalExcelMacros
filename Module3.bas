Attribute VB_Name = "Module3"
Sub add_named_ranges()
Attribute add_named_ranges.VB_ProcData.VB_Invoke_Func = " \n14"
this_val = ""

end_row = 0

name_range_col = 1
match_name_col = 4
start_row = 2

match_val = ActiveSheet.Cells(start_row, match_name_col)

For n = start_row To ActiveSheet.UsedRange.Rows.Count
    this_val = ActiveSheet.Cells(n, match_name_col)
    If this_val <> match_val Then
        end_row = n - 1
        
        ActiveWorkbook.Names.Add Name:=match_val, RefersToR1C1:= _
            "='Lookup'!R" & start_row & "C" & name_range_col & ":R" & end_row & "C" & name_range_col
        
        match_val = this_val
        start_row = n
    End If
Next n

ActiveWorkbook.Names.Add Name:=match_val, RefersToR1C1:= _
    "='Lookup'!R" & start_row & "C" & name_range_col & ":R" & ActiveSheet.UsedRange.Rows.Count & "C" & name_range_col

End Sub
Sub bugline_stuff()
Dim wb As Workbook
Dim sht As Worksheet
Dim val_blah As String
Dim test As String
Dim grade As String
Dim content_area As String
Dim subpart As String

Set wb = ActiveWorkbook

For this_sht = 1 To wb.Sheets.Count
    Set sht = wb.Sheets(this_sht)
    test = sht.Cells(2, 1)
    content_area = sht.Cells(2, 2)
    grade = sht.Cells(2, 3)
    
    For n = 2 To sht.UsedRange.Rows.Count
        subpart = sht.Cells(n, 7)
        val_blah = sht.Cells(n, 6)
        sht.Cells(n, 6) = left_before(val_blah, "*")
        
        If content_area = "ELA" Or content_area = "Social Studies" Then
            sht.Cells(n, 70) = "N"
            sht.Cells(n, 71) = "N"
        ElseIf content_area = "Science" Then
            If test <> "ALT" And (grade = "6" Or grade = "7" Or grade = "8") Then
                sht.Cells(n, 70) = "Y"
                sht.Cells(n, 71) = "Y"
            Else
                sht.Cells(n, 70) = "N"
                sht.Cells(n, 71) = "N"
            End If
        
        ElseIf content_area = "Math" Then
            If test <> "ALT" And grade = "2" Then
                sht.Cells(n, 70) = "N"
                sht.Cells(n, 71) = "Y"
            ElseIf test = "ALT" And grade = "2" Then
                sht.Cells(n, 70) = "N"
                sht.Cells(n, 71) = "N"
            ElseIf grade = 3 Then
                If subpart = "1" Then
                    sht.Cells(n, 70) = "N"
                    sht.Cells(n, 71) = "Y"
                Else
                    sht.Cells(n, 70) = "Y"
                    sht.Cells(n, 71) = "Y"
                End If
            Else
                If subpart = "1" Then
                    sht.Cells(n, 70) = "N"
                    sht.Cells(n, 71) = "N"
                Else
                    sht.Cells(n, 70) = "Y"
                    sht.Cells(n, 71) = "N"
                End If
            End If
            
        
        End If
        sht.Cells(n, 73) = right_after(val_blah, "*")
    Next n
Next this_sht
End Sub
Sub Merge_Sheet_By_Name()
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

Sub delete_accnum_col()
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
Dim lc As Long
Dim log_sht As Worksheet
Dim log_row As Long

sht_name = "Forms in Package"

xl_file_name = Application.GetOpenFilename("Excel files (*.*),*.*", , _
    "Browse for file to be merged", MultiSelect:=True)
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Set from_book = ActiveWorkbook
Set from_sht = from_book.Sheets("Export Worksheet")
Set log_sht = from_book.Sheets("Log")
log_row = 2

If IsArray(xl_file_name) Then
    For this_workbook = LBound(xl_file_name) To UBound(xl_file_name) ' iterate through each book
        Application.Workbooks.Open filename:=xl_file_name(this_workbook), ReadOnly:=False ' open the books
        Set to_book = ActiveWorkbook
        Set to_sht = to_book.ActiveSheet
        
        
        to_book.Close (True)
    Next this_workbook
End If
Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub
Sub add_accnums_to_FPs()
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
Dim lc As Long
Dim log_sht As Worksheet
Dim log_row As Long

xl_file_name = Application.GetOpenFilename("Excel files (*.*),*.*", , _
    "Browse for file to be merged", MultiSelect:=True)
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Set from_book = ActiveWorkbook
Set from_sht = from_book.Sheets("Export Worksheet")
Set log_sht = from_book.Sheets("Log")
log_row = 2

If IsArray(xl_file_name) Then
    For this_workbook = LBound(xl_file_name) To UBound(xl_file_name) ' iterate through each book
        Application.Workbooks.Open filename:=xl_file_name(this_workbook), ReadOnly:=False ' open the books
        Set to_book = ActiveWorkbook
        Set to_sht = to_book.ActiveSheet
        
        to_lr = to_sht.UsedRange.Rows.Count
        lc = to_sht.UsedRange.Columns.Count + 1
        
        acc_col = get_column_by_header(to_sht, "Accnum", False)
        
        If acc_col Then
            to_sht.Columns(acc_col).Delete
            lc = acc_col
        End If
        
        to_sht.Cells(1, lc) = "Accnum"
        to_sht.Cells(2, lc) = "=VLookup(E2,'[Extract_ItmAccNums_INC0096393.xlsx]Export Worksheet'!$C:$D,2,)"
        to_sht.Cells(2, lc).Copy Destination:=to_sht.Range(to_sht.Cells(3, lc), to_sht.Cells(to_lr, lc))
        
        For n = 2 To to_sht.UsedRange.Rows.Count
            log_sht.Cells(log_row, 1) = to_book.Name
            log_sht.Cells(log_row, 2) = to_sht.Name
            log_sht.Cells(log_row, 3) = to_sht.Cells(n, 5)
            
            to_sht.Cells(n, lc) = to_sht.Cells(n, lc).Value
            
            If IsError(to_sht.Cells(n, lc)) Then
                log_sht.Cells(log_row, 4) = ""
                log_sht.Cells(log_row, 5) = "Fail"
                to_sht.Cells(n, lc) = ""
            Else
                to_sht.Cells(n, lc) = to_sht.Cells(n, lc).Value
                log_sht.Cells(log_row, 4) = to_sht.Cells(n, lc)
                log_sht.Cells(log_row, 5) = "Pass"
            End If
            log_row = log_row + 1
        Next n
        to_book.Close (True)
    Next this_workbook
End If
Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub
