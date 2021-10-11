Attribute VB_Name = "Monday"
Option Explicit
Option Base 0
Sub Monday_MCAP_Schedule()
Dim wb As Workbook
Dim sht_actuals As Worksheet
Dim sht_project As Worksheet
Dim project_id As String
Dim project_row As Long
Dim last_project_row
Dim col_header As String
Dim actual_start_col As Long
Dim actual_start_date As String
Dim project_start As String
Dim date_start_value As String
Dim actual_finish_col As Long
Dim actual_finish_date As String
Dim project_finish As String
Dim date_finish_value As String

Set wb = ActiveWorkbook
Set sht_actuals = wb.Sheets(1)
Set sht_project = wb.Sheets(2)

Application.Calculation = xlManual

last_project_row = sht_project.UsedRange.Rows.Count

sht_project.Cells(1, 12) = "Actual Start"
sht_project.Cells(1, 13) = "Actual Finish"

For project_row = 25 To 25 ' last_project_row
    ' Only calculate actuals for Non-summary columns
    If sht_project.Cells(project_row, 1) = "No" Then
        ' get the actual dates from the actuals tab
        'get start
        col_header = sht_project.Cells(project_row, 2) & "|Actual Start"
        actual_start_col = get_column_by_name(3, col_header, sht_actuals)
        If actual_start_col <> -1 Then
            actual_start_date = get_formatted_date_or_text(sht_actuals.Cells(4, actual_start_col))
            project_start = get_formatted_date_or_text(sht_project.Cells(project_row, 3))
            
            date_start_value = process_date_checks(actual_start_date, project_start)
            
            sht_project.Cells(project_row, 12) = date_start_value
            If InStr(date_start_value, "Check") Then
                sht_project.Cells(project_row, 12).Interior.Color = RGB(255, 128, 0)
            End If
        Else
            sht_project.Cells(project_row, 12) = "NA"
        End If
        
        ' get finish
        col_header = sht_project.Cells(project_row, 2) & "|Actual Finish"
        actual_finish_col = get_column_by_name(3, col_header, sht_actuals)
        
        If actual_finish_col <> -1 Then
            actual_finish_date = get_formatted_date_or_text(sht_actuals.Cells(4, actual_finish_col))
            project_finish = get_formatted_date_or_text(sht_project.Cells(project_row, 4))
            
            date_finish_value = process_date_checks(actual_finish_date, project_finish)
            
            sht_project.Cells(project_row, 13) = date_finish_value
            If InStr(date_finish_value, "Check") Then
                sht_project.Cells(project_row, 13).Interior.Color = RGB(255, 128, 0)
            End If
        Else
            sht_project.Cells(project_row, 13) = "NA"
        End If
        
    Else
        ' put N/A in summary rows
        sht_project.Cells(project_row, 12) = "NA"
        sht_project.Cells(project_row, 13) = "NA"
    End If
Next project_row

Application.Calculation = xlAutomatic

End Sub
Function process_date_checks(actual_date As String, project_date As String) As String
' process logic for filling in cells
If IsDate(actual_date) Then actual_date = CDate(actual_date)
If IsDate(project_date) Then project_date = CDate(project_date)

If IsDate(actual_date) Then
    If Int(Format(actual_date, "YYYY")) >= 2045 Then
        actual_date = "Inactive"
    End If
End If

If project_date <> "NA" Then
    If project_date <> actual_date Then
        process_date_checks = actual_date & "|Check"
    Else
        process_date_checks = actual_date
    End If
ElseIf actual_date <> "Inactive" Then
    If project_date = "NA" And IsDate(actual_date) Then
        process_date_checks = actual_date
    Else
        process_date_checks = "NA"
    End If
Else
    process_date_checks = "Inactive"
End If
End Function
Function get_formatted_date_or_text(date_value As String) As String
Dim date_split As Variant
If date_value = "NA" Then
    get_formatted_date_or_text = "NA"
Else
    If IsDate(date_value) Then
        get_formatted_date_or_text = Format(date_value, "MM/DD/YYYY")
    ElseIf date_value <> "" Then
        date_split = Split(date_value, " ")
        get_formatted_date_or_text = date_split(1)
    Else
        get_formatted_date_or_text = "NA"
    End If
End If
End Function
Function get_column_by_name(row_num As Long, col_name As String, sht As Worksheet) As Long
Dim col As Long

For col = 1 To sht.UsedRange.Columns.Count
    If sht.Cells(row_num, col) = col_name Then
        get_column_by_name = col
        Exit Function
    End If
Next col

get_column_by_name = -1

End Function
Sub Monday_Actuals_Tabulator()
Dim wb As Workbook
Dim this_sht As Worksheet
Dim sum_sht As Worksheet
Dim c As Long
Dim sum_col As Long
Dim this_date As Date

Set wb = ActiveWorkbook

If Not WorksheetExists("Summary", wb) Then
    wb.Sheets.Add after:=wb.Sheets(wb.Sheets.Count)
    
    Set sum_sht = wb.Sheets(wb.Sheets.Count)
    sum_sht.Name = "Summary"
Else
    Set sum_sht = wb.Sheets("Summary")
    sum_sht.Cells.Delete
End If

sum_sht.Cells(5, 1) = "Name"
sum_col = 2

For Each this_sht In wb.Sheets
    If Not this_sht.Name = sum_sht.Name Then
        For c = 2 To this_sht.UsedRange.Columns.Count
            If InStr(this_sht.Cells(5, c), "Start") <> 0 Then
                this_date = Application.WorksheetFunction.Min(this_sht.Columns(c))
            Else
                this_date = Application.WorksheetFunction.Max(this_sht.Columns(c))
            End If
            
            If this_date <> 0 Then
                sum_sht.Cells(5, sum_col) = this_sht.Cells(5, c)
                sum_sht.Cells(6, sum_col) = this_date
                sum_col = sum_col + 1
            End If
        Next c
    End If
Next this_sht

End Sub
Function WorksheetExists(shtName As String, Optional wb As Workbook) As Boolean
    Dim sht As Worksheet

    If wb Is Nothing Then Set wb = ActiveWorkbook
    On Error Resume Next
    Set sht = wb.Sheets(shtName)
    On Error GoTo 0
    WorksheetExists = Not sht Is Nothing
End Function
