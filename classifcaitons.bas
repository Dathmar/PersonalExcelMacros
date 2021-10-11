Attribute VB_Name = "classifcaitons"
Option Explicit
Sub unzip_classifcations()
Dim n As Long
Dim r As Integer
Dim c As Integer
Dim r_col As Integer
Dim c_col As Integer
Dim id_col As Integer
Dim pid_col As Integer
Dim label_col As Integer
Dim label As String








End Sub
Sub build_classifcations()
Dim bp_sht As Worksheet
Dim cl_sht As Worksheet
Dim wb As Workbook
Dim r1c1 As String
Dim id As Long
Dim lc As Integer
Dim lr As Long
Dim n As Long
Dim i As Long
Dim r As Long
Dim c As Long
Dim r1c1_seq As Long
Dim seq As Long
Dim filtr_vals() As String

Application.DisplayAlerts = False

Set wb = ActiveWorkbook
Set bp_sht = wb.ActiveSheet
If sheet_exists(wb, "CLSFN_HIER") Then
   wb.Sheets("CLSFN_HIER").Delete
End If

wb.Sheets.Add
Set cl_sht = wb.ActiveSheet
cl_sht.Name = "CLSFN_HIER"

' setup headers
cl_sht.Cells(1, 1) = "ID"
cl_sht.Cells(1, 2) = "Value"
cl_sht.Cells(1, 3) = "Parent - ID"
cl_sht.Cells(1, 4) = "Row"
cl_sht.Cells(1, 5) = "Col"
cl_sht.Cells(1, 6) = "Seq"

' build values list
' r1c1 values are
' assume all others are not unique
' assume all others are sequential columns < 8
' if more then 8 columns move to row 2
lc = bp_sht.Cells(1, bp_sht.Columns.Count).End(xlToLeft).Column
lr = bp_sht.Cells(bp_sht.Rows.Count, 1).End(xlUp).Row
r = 1
c = 1
id = 1

' need to recursivly filter until last column value
ReDim filter_vals(0 To 0)
filter_vals(0) = bp_sht.Cells(1, 2)
'Call recursive_class(bp_sht, cl_sht, n, i, filtr_vals)

Application.DisplayAlerts = True
End Sub
Function recursive_class(bp_sht As Worksheet, cl_sht As Worksheet, lRow As Long, lCol As Long, filtr_vals() As Variant, Optional parent_id As Long)





End Function
Function unique_values_in_range(rng As Range)




End Function
Function sheet_exists(wb As Workbook, sht_name As String) As Boolean
Dim n As Long
For n = 1 To wb.Sheets.Count
    If wb.Sheets(n).Name = sht_name Then
        sheet_exists = True
        Exit Function
    End If
Next n
sheet_exists = False
End Function







































