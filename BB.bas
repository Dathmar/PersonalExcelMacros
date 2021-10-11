Attribute VB_Name = "BB"
Sub blah()
For c = 1 To ActiveSheet.UsedRange.Columns.Count
 ActiveSheet.Cells(1, c) = ActiveSheet.Columns(c).ColumnWidth
Next c
End Sub


Sub list_words()
Dim original As Worksheet
Dim words As Worksheet
Dim subjects() As String
Dim descriptions() As String

Set original = ActiveWorkbook.Sheets(1)
Set words = ActiveWorkbook.Sheets(2)

For n = 2 To original.UsedRange.Rows.Count
    
    
    
Next n

End Sub
Function get_all_words(rCell As Range) As String




End Function
Function is_unique_value(chk_val As String, chk_arr() As String) As Boolean
Dim n As Long

is_unique_value = True

For n = LBound(chk_arr) To UBound(chk_arr)
    If chk_val = chk_arr(n) Then
        is_unique_value = False
        Exit Function
    End If
Next n
End Function
