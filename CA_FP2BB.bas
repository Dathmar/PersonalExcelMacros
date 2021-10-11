Attribute VB_Name = "CA_FP2BB"
Option Explicit
Sub lots_of_fp2bb()
Dim wb As Workbook
Dim xl_file_name As Variant
Dim save_path As String
Dim this_workbook As Long

xl_file_name = Application.GetOpenFilename("Excel files (*.*),*.*", , _
    "Browse for file to be merged", MultiSelect:=True)
If IsArray(xl_file_name) Then
    For this_workbook = LBound(xl_file_name) To UBound(xl_file_name) ' iterate through each book
        Application.Workbooks.Open filename:=xl_file_name(this_workbook), ReadOnly:=True ' open the books
        Set wb = ActiveWorkbook
        Call fp2bb
        
        save_path = Left(wb.path, Len(wb.Name)) & "\Block Builders\" & CStr(get_name_without_extension(wb.Name)) & ".xlsx"
        
        Application.DisplayAlerts = False
        wb.SaveAs filename:=save_path, FileFormat:=xlOpenXMLWorkbook
        wb.Close savechanges:=False
        Application.DisplayAlerts = True
    Next this_workbook
End If

End Sub
Sub fp2bb()
Dim r As Long
Dim c As Long
Dim start_row As Long
Dim end_row As Long
Dim pe_col As Long
Dim sep_col As Long
Dim ccc_col As Long
Dim pval_col As Long
Dim poly_col As Long
Dim A_col As Long
Dim B_col As Long
Dim its_col As Long
Dim maxpts_col As Long
Dim seg_col As Long
Dim au_col As Long
Dim seq_col As Long
Dim acc_col As Long
Dim itm_nam_col As Long
Dim dom_col As Long
Dim type_col As Long
Dim class_col As Long
Dim dok_col As Long
Dim key_col As Long

Dim block As String
Dim au_id As String

Dim wb As Workbook
Dim sht As Worksheet
Dim sum_sht As Worksheet

Dim to_sht As Worksheet

Dim rng As Range

Set wb = ActiveWorkbook
Set sht = wb.Sheets(1)

Application.ScreenUpdating = False

'-!-!-!-!-!-!-!-!-!-!-!-testing code remove-!-!-!-!-!-!-!-!-!-!-!-
'Dim backup As Worksheet
'Set backup = wb.Sheets("Backup")

'sht.Cells.EntireRow.Delete
'backup.UsedRange.Copy Destination:=sht.Cells(1, 1)
'-!-!-!-!-!-!-!-!-!-!-!-!-end testing code-!-!-!-!-!-!-!-!-!-!-!-!-

If Not WorksheetExists("Summary", wb) Then
    wb.Sheets.Add(after:=sht).Name = "Summary"
End If

Set sum_sht = wb.Sheets("Summary")
sum_sht.Cells.EntireRow.Delete

Call initialize_summary(sum_sht)

pe_col = get_column_by_header(sht, "PE Code", False)
sep_col = get_column_by_header(sht, "SEP Code", False)
ccc_col = get_column_by_header(sht, "CCC Code", False)
pval_col = get_column_by_header(sht, "Pvalue", False)
poly_col = get_column_by_header(sht, "PolySerial", False)
A_col = get_column_by_header(sht, "Aparameter", False)
B_col = get_column_by_header(sht, "Bparameter", False)
its_col = get_column_by_header(sht, "ITS Item ID", False)
maxpts_col = get_column_by_header(sht, "Max Points", False)
seg_col = get_column_by_header(sht, "Segment ID", False)
au_col = get_column_by_header(sht, "AU ID", False)
seq_col = get_column_by_header(sht, "Item Sequence", False)
acc_col = get_column_by_header(sht, "Item Accnum", False)
itm_nam_col = get_column_by_header(sht, "Client Item ID", False)
dom_col = get_column_by_header(sht, "Domain", False)
type_col = get_column_by_header(sht, "ETS Item Type", False)
class_col = get_column_by_header(sht, "Part Name", False)
dok_col = get_column_by_header(sht, "DOK", False)
key_col = get_column_by_header(sht, "Answer Key Text", False)

sht.Cells(1, seq_col) = "Sequence"
sht.Cells(1, acc_col) = "Accnum"
sht.Cells(1, its_col) = "ITS ID"
sht.Cells(1, itm_nam_col) = "Item Name"
sht.Cells(1, dom_col) = "Domain"
sht.Cells(1, pe_col) = "PE"
sht.Cells(1, sep_col) = "SEP"
sht.Cells(1, ccc_col) = "CCC"
sht.Cells(1, type_col) = "Item Type"
sht.Cells(1, class_col) = "Item Class"
sht.Cells(1, dok_col) = "DOK"
sht.Cells(1, maxpts_col) = "Points"
sht.Cells(1, key_col) = "Key"
sht.Cells(1, pval_col) = "P-value"
sht.Cells(1, poly_col) = "Rpoly"
sht.Cells(1, A_col) = "a-parameter"
sht.Cells(1, B_col) = "b-parameter"

sht.Columns(pval_col).TextToColumns
sht.Columns(poly_col).TextToColumns
sht.Columns(A_col).TextToColumns
sht.Columns(B_col).TextToColumns
sht.Columns(maxpts_col).TextToColumns


With sht.Sort
    .SortFields.Clear
    .SortFields.Add key:=sht.Columns(seg_col), Order:=xlAscending
    .SetRange sht.UsedRange
    .header = xlYes
    .Apply
End With

For r = sht.Cells(sht.Rows.Count, 1).End(xlUp).Row To 2 Step -1
    If sht.Cells(r, 1) = "Form Name" Or sht.Cells(r, 1) = "" Then
        sht.Cells(r, 1).EntireRow.Delete
    End If
Next r

Call value_spacing(seg_col, sht)

For r = sht.Cells(sht.Rows.Count, 1).End(xlUp).Row + 1 To 2 Step -1
    If sht.Cells(r, 1) = "Form Name" Then
        sht.Cells(r, 1).EntireRow.Delete
    End If
    
    If sht.Cells(r, 1) = "" And sht.Cells(r - 1, 1) <> "" Then
        end_row = r - 1
        start_row = sht.Cells(end_row, 7).End(xlUp).Row
        If start_row = 1 Then start_row = 2
        
        ' apply averages
        sht.Cells(r, pval_col).Formula = "=iferror(average(" & get_column_letter(pval_col) & start_row & ":" & get_column_letter(pval_col) & end_row & "),"""")"
        sht.Cells(r, poly_col).Formula = "=iferror(average(" & get_column_letter(poly_col) & start_row & ":" & get_column_letter(poly_col) & end_row & "),"""")"
        sht.Cells(r, A_col).Formula = "=iferror(average(" & get_column_letter(A_col) & start_row & ":" & get_column_letter(A_col) & end_row & "),"""")"
        sht.Cells(r, B_col).Formula = "=iferror(average(" & get_column_letter(B_col) & start_row & ":" & get_column_letter(B_col) & end_row & "),"""")"
        
        ' apply number format
        sht.Cells(r, pval_col).NumberFormat = "#.###"
        sht.Cells(r, poly_col).NumberFormat = "#.###"
        sht.Cells(r, A_col).NumberFormat = "#.###"
        sht.Cells(r, B_col).NumberFormat = "#.###"
        
        ' apply sums
        sht.Cells(r, dok_col) = "Total Points"
        sht.Cells(r, maxpts_col).Formula = "=iferror(sum(" & get_column_letter(maxpts_col) & start_row & ":" & get_column_letter(maxpts_col) & end_row & "),"""")"
        
        sht.Rows(r + 1).EntireRow.Insert
    ElseIf sht.Cells(r, 1) <> "" And InStr(sht.Cells(r, acc_col), "_") = 0 Then
        sht.Cells(r, its_col) = right_after(sht.Cells(r, its_col), "-")
        sht.Cells(r, pe_col) = left_before(sht.Cells(r, pe_col), " ")
        sht.Cells(r, sep_col) = left_before(sht.Cells(r, sep_col), ".")
        sht.Cells(r, ccc_col) = left_before(sht.Cells(r, ccc_col), ".")
        
        block = sht.Cells(r, seg_col)
        au_id = sht.Cells(r, au_col)
        
        If InStr(sht.Cells(r, type_col), "Bar-PicturegraphSS") = 1 Then
            sht.Cells(r, type_col) = "Bar-PicturegraphSS"
            sht.Cells(r, class_col) = "iTE"
        ElseIf InStr(sht.Cells(r, type_col), "Bar-PicturegraphMS") = 1 Then
            sht.Cells(r, type_col) = "Bar-PicturegraphMS"
            sht.Cells(r, class_col) = "iTE"
        ElseIf InStr(sht.Cells(r, type_col), "Composite") = 1 Then
            sht.Cells(r, type_col) = "Composite"
            sht.Cells(r, class_col) = "COMP"
        ElseIf InStr(sht.Cells(r, type_col), "ExtendedText") = 1 Then
            sht.Cells(r, type_col) = "ExtendedText"
            sht.Cells(r, class_col) = "CR"
        ElseIf InStr(sht.Cells(r, type_col), "GridSS") = 1 Then
            sht.Cells(r, type_col) = "GridSS"
            sht.Cells(r, class_col) = "aTE"
        ElseIf InStr(sht.Cells(r, type_col), "GridMS") = 1 Then
            sht.Cells(r, type_col) = "GridMS"
            sht.Cells(r, class_col) = "aTE"
        ElseIf InStr(sht.Cells(r, type_col), "InlineChoiceListMS") = 1 Then
            sht.Cells(r, type_col) = "InlineChoiceListMS"
            sht.Cells(r, class_col) = "aTE"
        ElseIf InStr(sht.Cells(r, type_col), "InlineChoiceListSS") = 1 Then
            sht.Cells(r, type_col) = "InlineChoiceListSS"
            sht.Cells(r, class_col) = "aTE"
        ElseIf InStr(sht.Cells(r, type_col), "Leader") = 1 Then
            sht.Cells(r, type_col) = "Leader"
            sht.Cells(r, class_col) = "~"
        ElseIf InStr(sht.Cells(r, type_col), "MatchSS") = 1 Then
            sht.Cells(r, type_col) = "MatchSS"
            sht.Cells(r, class_col) = "iTE"
        ElseIf InStr(sht.Cells(r, type_col), "MatchMS") = 1 Then
            sht.Cells(r, type_col) = "MatchMS"
            sht.Cells(r, class_col) = "iTE"
        ElseIf InStr(sht.Cells(r, type_col), "MCMS") = 1 Then
            sht.Cells(r, type_col) = "MCMS"
            sht.Cells(r, class_col) = "MC"
        ElseIf InStr(sht.Cells(r, type_col), "MCSS") = 1 Then
            sht.Cells(r, type_col) = "MCSS"
            sht.Cells(r, class_col) = "MC"
        ElseIf InStr(sht.Cells(r, type_col), "ZonesSS") = 1 Then
            sht.Cells(r, type_col) = "ZonesSS"
            sht.Cells(r, class_col) = "iTE"
        ElseIf InStr(sht.Cells(r, type_col), "ZonesMS") = 1 Then
            sht.Cells(r, type_col) = "ZonesMS"
            sht.Cells(r, class_col) = "iTE"
        End If
    ElseIf sht.Cells(r, 1) <> "" And InStr(sht.Cells(r, acc_col), "_") <> 0 Then
        sht.Rows(r).EntireRow.Delete
    Else
        sht.Cells(r, seq_col) = "Block"
        sht.Cells(r, acc_col) = block
        sht.Cells(r, its_col) = au_id
        
        sht.Rows(r + 1).EntireRow.Insert
        sht.Rows(1).Copy Destination:=sht.Rows(r + 1)
        
    End If
    
Next r

sht.Rows(1).Insert
sht.Cells(1, seq_col) = "Block"
sht.Cells(1, acc_col) = block
sht.Cells(1, its_col) = au_id

If Not WorksheetExists("BB", wb) Then
    wb.Sheets.Add(after:=sht).Name = "BB"
End If

Set to_sht = wb.Sheets("BB")
to_sht.Cells.EntireRow.Delete

sht.Columns(seq_col).Copy Destination:=to_sht.Columns(1)
sht.Columns(acc_col).Copy Destination:=to_sht.Columns(2)
sht.Columns(its_col).Copy Destination:=to_sht.Columns(3)
sht.Columns(itm_nam_col).Copy Destination:=to_sht.Columns(4)
sht.Columns(dom_col).Copy Destination:=to_sht.Columns(5)
sht.Columns(pe_col).Copy Destination:=to_sht.Columns(6)
sht.Columns(sep_col).Copy Destination:=to_sht.Columns(7)
sht.Columns(ccc_col).Copy Destination:=to_sht.Columns(8)
sht.Columns(type_col).Copy Destination:=to_sht.Columns(9)
sht.Columns(class_col).Copy Destination:=to_sht.Columns(10)
sht.Columns(dok_col).Copy Destination:=to_sht.Columns(11)
sht.Columns(maxpts_col).Copy Destination:=to_sht.Columns(12)
sht.Columns(key_col).Copy Destination:=to_sht.Columns(13)
sht.Columns(pval_col).Copy Destination:=to_sht.Columns(14)
sht.Columns(poly_col).Copy Destination:=to_sht.Columns(15)
sht.Columns(A_col).Copy Destination:=to_sht.Columns(16)
sht.Columns(B_col).Copy Destination:=to_sht.Columns(17)

to_sht.Columns(1).ColumnWidth = 10.43
to_sht.Columns(2).ColumnWidth = 11
to_sht.Columns(3).ColumnWidth = 10.86
to_sht.Columns(4).ColumnWidth = 22
to_sht.Columns(5).ColumnWidth = 11.43
to_sht.Columns(6).ColumnWidth = 12
to_sht.Columns(7).ColumnWidth = 7.29
to_sht.Columns(8).ColumnWidth = 7.29
to_sht.Columns(9).ColumnWidth = 13.86
to_sht.Columns(10).ColumnWidth = 10.71
to_sht.Columns(11).ColumnWidth = 11.57
to_sht.Columns(12).ColumnWidth = 6.57
to_sht.Columns(13).ColumnWidth = 18.14
to_sht.Columns(14).ColumnWidth = 7.71
to_sht.Columns(15).ColumnWidth = 11
to_sht.Columns(16).ColumnWidth = 12.57
to_sht.Columns(17).ColumnWidth = 12.57

Call remove_borders(to_sht.Cells)

r = 1

Do
    block = to_sht.Cells(r, 2)
    start_row = r + 2
    end_row = to_sht.Cells(r, 2).End(xlDown).Row
    
    Set rng = to_sht.Range(to_sht.Cells(start_row - 1, 1), to_sht.Cells(end_row, 17))
    Set rng = Union(rng, to_sht.Range(to_sht.Cells(start_row - 2, 1), to_sht.Cells(start_row - 2, 3)), _
                    to_sht.Range(to_sht.Cells(end_row + 1, 11), to_sht.Cells(end_row + 1, 12)), _
                    to_sht.Range(to_sht.Cells(end_row + 1, 14), to_sht.Cells(end_row + 1, 17)))
                
    Call add_borders(rng)
    
    Call write_summary(start_row, end_row, block, sum_sht)
    
    r = to_sht.Cells(end_row, 2).End(xlDown).Row
        
Loop Until r >= to_sht.Rows.Count

block = to_sht.Cells(r, 2)
start_row = r + 2
end_row = to_sht.Cells(r, 2).End(xlDown).Row
    
Call write_summary(start_row, end_row, block, sum_sht)

to_sht.Cells.Font.Name = "Arial"
to_sht.Cells.Font.Size = 10

Application.DisplayAlerts = False
sht.Delete
Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub
Sub blah()
Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub

Function add_borders(rng As Range)
rng.Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous
rng.Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous
rng.Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous
rng.Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
rng.Borders(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlContinuous
rng.Borders(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlContinuous
End Function

Function remove_borders(rng As Range)
rng.Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlLineStyleNone
rng.Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlLineStyleNone
rng.Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlLineStyleNone
rng.Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlLineStyleNone
rng.Borders(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlLineStyleNone
rng.Borders(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlLineStyleNone
End Function
Function write_summary(start_row As Long, end_row As Long, block As String, sum_sht As Worksheet)
Dim last_block As String
Dim lc As Long
Dim r As Long
Dim seg_length As Long
Dim lr As Long
Dim c As Long
Dim sum_form As String

lc = sum_sht.Cells(1, Columns.Count).End(xlToLeft).Column
lr = sum_sht.Cells(sum_sht.Rows.Count, 1).End(xlUp).Row
last_block = sum_sht.Cells(1, lc)

lc = lc + 1

If Left(last_block, 1) <> Left(block, 1) And last_block <> "Category" Then
    ' do the block summary
    sum_sht.Cells(1, lc) = "Segment " & Left(last_block, 1)
    seg_length = Int(Right(last_block, Len(last_block) - 1))
    
    If lc - seg_length <= 0 Then
        seg_length = 1
    End If
    
    For r = 2 To lr
        sum_sht.Cells(r, lc) = "=sum(" & get_column_letter(lc - seg_length) & r & ":" & get_column_letter(lc - 1) & r & ")"
    Next r
    lc = lc + 1
End If

If block <> "" Then
    sum_sht.Cells(1, lc) = block
    
    sum_sht.Cells(2, lc) = "=countif(A" & start_row & ":A" & end_row & ",""<>~"")"
    sum_sht.Cells(3, lc) = "=BB!L" & end_row + 1
    sum_sht.Cells(4, lc) = "=COUNTIF(BB!J" & start_row & ":J" & end_row & ",""MC"")"
    sum_sht.Cells(5, lc) = "=COUNTIF(BB!J" & start_row & ":J" & end_row & ",""aTE"")"
    sum_sht.Cells(6, lc) = "=COUNTIF(BB!J" & start_row & ":J" & end_row & ",""iTE"")"
    sum_sht.Cells(7, lc) = "=COUNTIF(BB!J" & start_row & ":J" & end_row & ",""CR"")"
    sum_sht.Cells(8, lc) = "=COUNTIF(BB!J" & start_row & ":J" & end_row & ",""COMP"")"
    sum_sht.Cells(9, lc) = "=COUNTIF(BB!E" & start_row & ":E" & end_row & ",""PS"")"
    sum_sht.Cells(10, lc) = "=COUNTIF(BB!E" & start_row & ":E" & end_row & ",""LS"")"
    sum_sht.Cells(11, lc) = "=COUNTIF(BB!E" & start_row & ":E" & end_row & ",""ESS"")"
    sum_sht.Cells(12, lc) = "=COUNTIF(BB!E" & start_row & ":E" & end_row & ",""ETS*"")"
    sum_sht.Cells(13, lc) = "=COUNTIF(BB!G" & start_row & ":G" & end_row & ",1)"
    sum_sht.Cells(14, lc) = "=COUNTIF(BB!G" & start_row & ":G" & end_row & ",2)"
    sum_sht.Cells(15, lc) = "=COUNTIF(BB!G" & start_row & ":G" & end_row & ",3)"
    sum_sht.Cells(16, lc) = "=COUNTIF(BB!G" & start_row & ":G" & end_row & ",4)"
    sum_sht.Cells(17, lc) = "=COUNTIF(BB!G" & start_row & ":G" & end_row & ",5)"
    sum_sht.Cells(18, lc) = "=COUNTIF(BB!G" & start_row & ":G" & end_row & ",6)"
    sum_sht.Cells(19, lc) = "=COUNTIF(BB!G" & start_row & ":G" & end_row & ",7)"
    sum_sht.Cells(20, lc) = "=COUNTIF(BB!G" & start_row & ":G" & end_row & ",8)"
    sum_sht.Cells(21, lc) = "=COUNTIF(BB!H" & start_row & ":H" & end_row & ",1)"
    sum_sht.Cells(22, lc) = "=COUNTIF(BB!H" & start_row & ":H" & end_row & ",2)"
    sum_sht.Cells(23, lc) = "=COUNTIF(BB!H" & start_row & ":H" & end_row & ",3)"
    sum_sht.Cells(24, lc) = "=COUNTIF(BB!H" & start_row & ":H" & end_row & ",4)"
    sum_sht.Cells(25, lc) = "=COUNTIF(BB!H" & start_row & ":H" & end_row & ",5)"
    sum_sht.Cells(26, lc) = "=COUNTIF(BB!H" & start_row & ":H" & end_row & ",6)"
    sum_sht.Cells(27, lc) = "=COUNTIF(BB!H" & start_row & ":H" & end_row & ",7)"
    sum_sht.Cells(28, lc) = "=COUNTIF(BB!H" & start_row & ":H" & end_row & ",""n/a"")"
    sum_sht.Cells(29, lc) = "=COUNTIF(BB!K" & start_row & ":K" & end_row & ",1)"
    sum_sht.Cells(30, lc) = "=COUNTIF(BB!K" & start_row & ":K" & end_row & ",2)"
    sum_sht.Cells(31, lc) = "=COUNTIF(BB!K" & start_row & ":K" & end_row & ",3)"
    sum_sht.Cells(32, lc) = "=COUNTIF(BB!K" & start_row & ":K" & end_row & ",4)"
End If

If block = "" Then
    sum_sht.Cells(1, lc) = "Form"
    
    sum_form = ""
    
    For c = 1 To lc
        If InStr(sum_sht.Cells(1, c), "Segment") = 1 Then
            If sum_form = "" Then
                sum_form = get_column_letter(c) & "row1234"
            Else
                sum_form = sum_form & "," & get_column_letter(c) & "row1234"
            End If
        End If
    Next c
    
    For r = 2 To lr
        sum_sht.Cells(r, lc) = "=SUM(" & Replace(sum_form, "row1234", r) & ")"
    Next r
    
End If

End Function
Function left_before(rng As Variant, before As String, Optional trim_str As Boolean = False) As String
Dim ret As String
Dim txt As String

txt = CStr(rng)
If InStr(txt, before) <> 0 Then
    ret = Left(txt, InStr(txt, before) - 1)

    If trim_str Then
        ret = Trim(ret)
    End If
    left_before = ret
Else
    left_before = txt
End If

End Function
Function right_after(rng As Variant, after As String, Optional trim_str As Boolean = False) As String
Dim ret As String
Dim txt As String

txt = CStr(rng)
If InStr(txt, after) <> 0 Then
    ret = Right(txt, Len(txt) - InStr(txt, after) - Len(after) + 1)

    If trim_str Then
        ret = Trim(ret)
    End If
    right_after = ret
Else
    right_after = txt
End If


End Function
Function mid_between(rng As Variant, first As String, second As String, Optional trim_str As Boolean = False) As String
Dim ret As String
Dim txt As String

txt = CStr(rng)
If InStr(txt, first) = 0 Or InStr(txt, second) Then
    ret = Mid(txt, InStr(txt, first) + Len(first), InStr(txt, second) - InStr(txt, first) - Len(first))

    If trim_str Then
        ret = Trim(ret)
    End If
    mid_between = ret
Else
    mid_between = txt
End If

End Function
Function get_column_by_header(ws As Worksheet, header As String, Optional check_case = True) As Long
Dim n As Long
Dim lc As Long

lc = ws.UsedRange.Columns.Count

For n = 1 To lc
    If ws.Cells(1, n) = header And check_case Then
        get_column_by_header = n
        Exit Function
    ElseIf LCase(ws.Cells(1, n)) = LCase(header) And Not check_case Then
        get_column_by_header = n
        Exit Function
    End If
Next n
get_column_by_header = 0

End Function
Function get_column_letter(number As Long) As String
get_column_letter = Split(Cells(1, number).Address, "$")(1)
End Function
Function WorksheetExists(shtName As String, Optional wb As Workbook) As Boolean
    Dim sht As Worksheet

    If wb Is Nothing Then Set wb = ThisWorkbook
    On Error Resume Next
    Set sht = wb.Sheets(shtName)
    On Error GoTo 0
    WorksheetExists = Not sht Is Nothing
End Function
Function initialize_summary(sum_tab As Worksheet)
sum_tab.Cells(1, 1) = "Category"
sum_tab.Cells(2, 1) = "Items"
sum_tab.Cells(3, 1) = "Points"
sum_tab.Cells(4, 1) = "MC"
sum_tab.Cells(5, 1) = "aTE"
sum_tab.Cells(6, 1) = "iTE"
sum_tab.Cells(7, 1) = "CR"
sum_tab.Cells(8, 1) = "COMP"
sum_tab.Cells(9, 1) = "PS"
sum_tab.Cells(10, 1) = "LS"
sum_tab.Cells(11, 1) = "ESS"
sum_tab.Cells(12, 1) = "ETS"
sum_tab.Cells(13, 1) = "'SEP 1"
sum_tab.Cells(14, 1) = "'SEP 2"
sum_tab.Cells(15, 1) = "'SEP 3"
sum_tab.Cells(16, 1) = "'SEP 4"
sum_tab.Cells(17, 1) = "'SEP 5"
sum_tab.Cells(18, 1) = "'SEP 6"
sum_tab.Cells(19, 1) = "'SEP 7"
sum_tab.Cells(20, 1) = "'SEP 8"
sum_tab.Cells(21, 1) = "CCC 1"
sum_tab.Cells(22, 1) = "CCC 2"
sum_tab.Cells(23, 1) = "CCC 3"
sum_tab.Cells(24, 1) = "CCC 4"
sum_tab.Cells(25, 1) = "CCC 5"
sum_tab.Cells(26, 1) = "CCC 6"
sum_tab.Cells(27, 1) = "CCC 7"
sum_tab.Cells(28, 1) = "No CCC"
sum_tab.Cells(29, 1) = "DOK 1"
sum_tab.Cells(30, 1) = "DOK 2"
sum_tab.Cells(31, 1) = "DOK 3"
sum_tab.Cells(32, 1) = "DOK 4"

End Function
Sub value_spacing(col As Long, sht As Worksheet)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       03/09/2012                                        '''
'''The purpose is to add line spacing between each unique value in the selected column.     '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim row_count As Long
Dim this_row As Long
Dim this_str As String
Dim last_str As String
Dim selected_column As Long
Dim space_number As String
Dim n As Long
Dim selection_address As String

space_number = 2

row_count = sht.UsedRange.Rows.Count

selected_column = col
For this_row = row_count To 2 Step -1
    this_str = sht.Cells(this_row, selected_column)
    If this_str <> last_str Then
        For n = 1 To space_number
            sht.Rows(this_row + 1).EntireRow.Insert
        Next n
        last_str = this_str
    End If
Next this_row

End Sub
Function get_file_directory(path) As String
   get_file_directory = Left(path, InStrRev(path, Application.PathSeparator))
End Function
Function get_name_without_extension(filename As String) As String
Dim extension As Variant
extension = get_extension(filename)

get_name_without_extension = Replace(filename, "." & extension, "")

End Function

Public Function get_extension(filename As String) As String
Dim name_split As Variant
name_split = Split(filename, ".")
get_extension = CStr(name_split(UBound(name_split)))
End Function
