Attribute VB_Name = "OIL_Coverpages"
Option Explicit
Sub Add_Coverpages_to_OIL()
Dim at_end_of_file As Boolean
Dim oil_wb As Workbook
Dim oil_sht As Worksheet

Dim lookup_wb As Workbook
Dim lookup_sht As Worksheet

Dim rSection_header As Long
Dim rSection_end As Long
Dim rRow As Long

Set oil_wb = ActiveWorkbook
Set oil_sht = oil_wb.ActiveSheet

rRow = 1


While at_end_of_file = False
    If InStr(oil_sht.Cells(rRow, 1), "ENSURE THAT VALIDATIONS HAVE BEEN RUN FOR ALL AUS") > 0 Then
        oil_sht.Cells(rRow, 1).EntireRow.Delete
    End If
    
    If is_row_section_header(oil_sht.Cells(rRow, 1)) Then
        rSection_header = rRow
        rSection_end = get_current_section_end(rSection_header, oil_sht)
        
        rRow = rRow + 1
    Else
        rRow = rRow + 1
    End If
    If rRow = 30 Then
        Exit Sub
    End If
Wend

End Sub
Private Function get_current_section_end(start_row As Long, sht As Worksheet) As Long
Dim rRow As Long

For rRow = start_row To sht.Rows.Count
    If sht.Cells(rRow, 1) = "" Then
        get_current_section_end = rRow - 1
        Exit Function
    End If
Next rRow

End Function
Private Function is_row_section_header(row_text As String) As Boolean
If InStr(row_text, "::") > 0 Then
    is_row_section_header = True
Else
    is_row_section_header = False
End If
End Function
Private Function is_workbook_open(filepath) As Boolean
Dim wb As Workbook
For Each wb In Application.Workbooks
    If wb.path = filepath Then
        is_workbook_open = True
        Exit Function
    End If
Next wb
is_workbook_open = False
End Function
