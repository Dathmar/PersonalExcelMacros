Attribute VB_Name = "PARCCStat"
Option Explicit
Sub IA_Fill()
Dim ia As Worksheet
Dim lr As Long
Dim n As Long
Dim accnums() As String
Dim outputs As Variant
Dim test_maps As Variant
Dim content As String
Dim grade As String

Set ia = ActiveSheet

' select output files
outputs = Application.GetOpenFilename("Excel files (*.xls; *.xlsx), *.xls; *.xlsx", , _
    "Browse for output files", MultiSelect:=True)
' select sas files
test_maps = Application.GetOpenFilename("Excel files (*.xls; *.xlsx), *.xls; *.xlsx", , _
    "Browse for test maps", MultiSelect:=True)
' if user didn't select outputs or sas files exit sub / clicking cancel returns False
' clicking OK returns an array of filenames

content = ia.Cells(3, 2)
grade = ia.Cells(3, 3)

' get all accnums in the IA file
lr = ia.Cells(ia.Rows.Count, 1).End(xlUp).Row
ReDim accnums(0 To lr - 3)
For n = 3 To lr
    accnums(n - 3) = ia.Cells(n, 1)
    Debug.Print accnums(n - 3)
Next n

' might break this into a function get_form_information
' loop through outputs and list all needed info by file
Dim i As Long
Dim cur_file As Workbook
Dim cur_sht As Worksheet
Dim cur_row As Long
Dim k As Long
Dim rng_rows As Long
Dim rng As Range
Dim rRow As Range
cur_row = 3

' this is super ugly :-(
' Process the output IA files
For n = LBound(outputs) To UBound(outputs)
    Application.Workbooks.Open filename:=outputs(n), ReadOnly:=True, Editable:=False, notify:=False
    Set cur_file = ActiveWorkbook
    Set cur_sht = cur_file.Sheets("All_Items")
    
    ' loop through each accnum and get
    For i = LBound(accnums) To UBound(accnums)
        ' get the item number column then filter by accnum
        cur_sht.UsedRange.AutoFilter Field:=get_col("ItemNumber", cur_sht), Criteria1:=accnums(i)
        
        ' now loop through each visible row and copy data for the current accnum
        Set rng = cur_sht.UsedRange.SpecialCells(xlCellTypeVisible)
        For k = 1 To rng.Areas.Count
            
            For Each rRow In rng.Areas(k).Rows
                If rRow.Row <> 1 Then
                    
                    ia.Cells(cur_row, 1) = accnums(i)
                    ia.Cells(cur_row, 2) = content
                    ia.Cells(cur_row, 3) = grade
                    ia.Cells(cur_row, 4) = cur_sht.Cells(rRow.Row, get_col("Form", cur_sht))
                    ia.Cells(cur_row, 11) = cur_sht.Cells(rRow.Row, get_col("N_reached", cur_sht))
                    ia.Cells(cur_row, 12) = cur_sht.Cells(rRow.Row, get_col("AIS", cur_sht))
                    ia.Cells(cur_row, 13) = cur_sht.Cells(rRow.Row, get_col("AIS_as_proportion_of_max_score", cur_sht))
                    ia.Cells(cur_row, 14) = cur_sht.Cells(rRow.Row, get_col("polyserial", cur_sht))
                    
                    cur_row = cur_row + 1
                End If
            Next rRow
        Next k
        
    Next i
    cur_file.Close savechanges:=False
Next n


' loop through test maps and filter UIN and form then fill in deets
For n = LBound(test_maps) To UBound(test_maps)
    Application.Workbooks.Open filename:=test_maps(n), ReadOnly:=True, Editable:=False, notify:=False
    Set cur_file = ActiveWorkbook
    Set cur_sht = cur_file.ActiveSheet
    
    ' loop through each accnum and get
    For i = 3 To ia.UsedRange.Rows.Count

        ' get the item number column then filter by accnum then find form column and filter by form
        cur_sht.UsedRange.AutoFilter Field:=get_col("IFF_UIN", cur_sht), Criteria1:=ia.Cells(i, 1)
        cur_sht.UsedRange.AutoFilter Field:=get_col("Form", cur_sht), Criteria1:=ia.Cells(i, 4)
        
        ' now loop through each visible row and copy data for the current accnum
        Set rng = cur_sht.UsedRange.SpecialCells(xlCellTypeVisible)
        For k = 1 To rng.Areas.Count
            
            For Each rRow In rng.Areas(k).Rows
                If rRow.Row <> 1 Then
                    
                    ia.Cells(i, 5) = cur_sht.Cells(rRow.Row, get_col("Form", cur_sht))
                    ia.Cells(i, 6) = cur_sht.Cells(rRow.Row, get_col("Mode", cur_sht))
                    If ia.Cells(i, 6) = "E" Then
                        ia.Cells(i, 6) = "CBT"
                    Else
                        ia.Cells(i, 6) = "PBT"
                    End If
                    ia.Cells(i, 7) = cur_sht.Cells(rRow.Row, get_col("PARCC_Evidence_Statement_1", cur_sht))
                    ia.Cells(i, 8) = cur_sht.Cells(rRow.Row, get_col("PARCC_Subclaim_1", cur_sht))
                    ia.Cells(i, 9) = cur_sht.Cells(rRow.Row, get_col("PARCC_Subclaim_2", cur_sht))
                    ia.Cells(i, 10) = cur_sht.Cells(rRow.Row, get_col("PARCC_Subclaim_3", cur_sht))
                End If
            Next rRow
        Next k
        
    Next i
    cur_file.Close savechanges:=False
Next n


End Sub
Function get_col(ByRef col_name As String, ByRef sht As Worksheet) As Long
Dim n As Long
For n = 1 To sht.Cells(1, sht.Columns.Count).End(xlToLeft).Column
    If sht.Cells(1, n) = col_name Then
        get_col = n
        Exit Function
    End If
Next n

End Function
Function count_rows_in_range(rng As Range) As Long
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       03/14/2012                                        '''
'''Count the rows in a continuous or non-continuous range.                                  '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim i As Long
For i = 1 To rng.Areas.Count
count_rows_in_range = count_rows_in_range + rng.Areas(i).Rows.Count
Next i
End Function
