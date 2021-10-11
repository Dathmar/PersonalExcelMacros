Attribute VB_Name = "ProjectPlan"
Sub expand_project_plan()
Dim n As Long
Dim i As Long
Dim lst() As Variant
Dim sht As Worksheet
Dim wb As Workbook
Dim lc As Long
Dim lr As Long

Set wb = ActiveWorkbook
Set sht = wb.ActiveSheet

lr = sht.Cells(sht.Rows.Count, 1).End(xlUp).Row
lc = sht.Cells(1, sht.Columns.Count).End(xlToLeft).Column + 1

' init
ReDim lst(0 To 1, 0 To 0)
lst(0, 0) = Trim(sht.Cells(2, 2))
lst(1, 0) = left_pad_len(sht.Cells(2, 2))

For n = 3 To lr
    this_name = Trim(sht.Cells(n, 2))
    this_pad = left_pad_len(sht.Cells(n, 2))
    ubnd = UBound(lst, 2)
    cnt = 0
    brk = False
    
    Do While brk = False
        If this_pad > lst(1, ubnd) Then
            ubnd = ubnd + 1
            ReDim Preserve lst(0 To 1, 0 To ubnd)
            brk = True
        ElseIf this_pad = lst(1, ubnd) Then
            brk = True
        ElseIf this_pad < lst(1, ubnd) Then
            ubnd = ubnd - 1
            ReDim Preserve lst(0 To 1, 0 To ubnd)
        End If
    Loop
    
    lst(0, ubnd) = this_name
    lst(1, ubnd) = this_pad
    
    For i = 0 To ubnd
        sht.Cells(n, lc + i) = lst(0, i)
    Next i
    
Next n

End Sub
Function left_pad_len(txt As String) As Long
cnt = 0
Do While Left(txt, 1) = " " Or txt = ""
    cnt = cnt + 1
    txt = Right(txt, Len(txt) - 1)
Loop
left_pad_len = cnt
End Function

