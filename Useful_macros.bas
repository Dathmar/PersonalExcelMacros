Attribute VB_Name = "Useful_macros"
Sub multi_row_by_del()
'
' Copies a row multiple times based on demimiter in a column
' Example
' INIT
'      A     B    C
'  1 Hi-By   1    2
'
' Final del = "-", splt_col = 1
'      A     B    C
'  1   Hi    1    2
'  2   By    1    2
'

Dim n As Long
Dim i As Long
Dim frm_sht As Worksheet
Dim to_sht As Worksheet
Dim wb As Workbook
Dim del As String
Dim splt As Variant
Dim frm_lr As Long
Dim to_lr As Long
Dim splt_col As Long

Application.ScreenUpdating = False
Set wb = ActiveWorkbook

If wb.Sheets.Count < 2 Then
    wb.Sheets.Add after:=wb.Sheets(wb.Sheets.Count)
End If

Set frm_sht = wb.Sheets(1)
Set to_sht = wb.Sheets(2)

del = "|"
splt_col = 6

frm_lr = frm_sht.Cells(frm_sht.Rows.Count, splt_col).End(xlUp).Row

frm_sht.Rows(1).Copy Destination:=to_sht.Rows(1)
With to_sht
For n = 2 To frm_lr
    to_lr = .Cells(.Rows.Count, 2).End(xlUp).Row + 1
    
    splt = Split(frm_sht.Cells(n, splt_col), del)
    cnt = 0
    
    If IsArray(splt) And frm_sht.Cells(n, splt_col) <> "" Then
        For i = LBound(splt) To UBound(splt)
            frm_sht.Rows(n).Copy Destination:=.Rows(to_lr + cnt)
            .Cells(to_lr + cnt, splt_col) = Trim(splt(i))
            cnt = cnt + 1
        Next i
    Else
        frm_sht.Rows(n).Copy Destination:=.Rows(to_lr)
    End If
Next n
End With
Application.ScreenUpdating = True
End Sub
Sub pipe_to_multi()
Dim n As Long
Dim i As Long
Dim j As Long
Dim p As Long
Dim lr As Long
Dim from_cnt As Long

Dim cpy_frm As Long
Dim cpy_to As Long


Dim bk As Workbook
Dim from_sht As Worksheet
Dim to_sht As Worksheet

Dim splt As Variant

Dim del As String

Set bk = ActiveWorkbook
If bk.Sheets.Count < 2 Then
    bk.Sheets.Add after:=bk.Sheets(bk.Sheets.Count)
End If

del = "|"

Set from_sht = bk.Sheets(1)
Set to_sht = bk.Sheets(2)

cpy_frm = 1
cpy_to = 16

from_cnt = from_sht.Cells(from_sht.Rows.Count, 1).End(xlUp).Row
from_sht.Rows(1).Copy Destination:=to_sht.Rows(1) ' copy headers
For n = 2 To from_cnt ' loop through each row
lr = to_sht.Cells(to_sht.Rows.Count, 2).End(xlUp).Row + 1
    For i = cpy_frm To cpy_to
        With to_sht
            If Right(from_sht.Cells(n, i), 1) = del Then from_sht.Cells(n, i) = Left(from_sht.Cells(n, i), Len(from_sht.Cells(n, i)) - 1) ' remove trailing delimiters
            
            If from_sht.Cells(n, i) <> "" Then
                splt = Split(from_sht.Cells(n, i), del) ' split delimiter list
                
                For j = LBound(splt) To UBound(splt) ' loop through list and add
                    For p = 1 To cpy_frm - 1
                        to_sht.Cells(lr + j, p) = from_sht.Cells(n, p)
                    Next p
                    to_sht.Rows(lr).Copy Destination:=to_sht.Rows(lr + j)
                    to_sht.Cells(lr + j, i) = splt(j)
                Next j
            Else
                For p = 1 To cpy_frm - 1
                    to_sht.Cells(lr + j, p) = from_sht.Cells(n, p)
                Next p
                to_sht.Cells(lr, i) = from_sht.Cells(n, i)
            End If
        End With
    Next i
Next n

End Sub
Sub pipe_to_multi_with_header_tag()
Dim n As Long
Dim i As Long
Dim j As Long
Dim p As Long
Dim lr As Long
Dim from_cnt As Long

Dim cpy_frm As Long
Dim cpy_to As Long


Dim bk As Workbook
Dim from_sht As Worksheet
Dim to_sht As Worksheet

Dim splt As Variant

Dim del As String

Set bk = ActiveWorkbook
If bk.Sheets.Count < 2 Then
    bk.Sheets.Add after:=bk.Sheets(bk.Sheets.Count)
End If

del = "|"

Set from_sht = bk.Sheets(1)
Set to_sht = bk.Sheets(2)

cpy_frm = 3
cpy_to = 29

from_cnt = from_sht.Cells(from_sht.Rows.Count, 1).End(xlUp).Row
from_sht.Range(from_sht.Cells(1, 1), from_sht.Cells(1, cpy_frm - 1)).Copy Destination:=to_sht.Range(to_sht.Cells(1, 1), to_sht.Cells(1, cpy_frm - 1)) ' copy headers
to_sht.Cells(1, cpy_frm) = "Header"
to_sht.Cells(1, cpy_frm + 1) = "Value"

For n = 2 To from_cnt ' loop through each row

    For i = cpy_frm To cpy_to
        lr = to_sht.Cells(to_sht.Rows.Count, 2).End(xlUp).Row + 1
        With to_sht
            If Right(from_sht.Cells(n, i), 1) = del Then from_sht.Cells(n, i) = Left(from_sht.Cells(n, i), Len(from_sht.Cells(n, i)) - 1) ' remove trailing delimiters

            If from_sht.Cells(n, i) <> "" Then
                splt = Split(from_sht.Cells(n, i), del) ' split delimiter list
                
                For j = LBound(splt) To UBound(splt) ' loop through list and add
                    to_sht.Range(to_sht.Cells(lr + j, 1), to_sht.Cells(lr + j, cpy_frm - 1)).Value = from_sht.Range(from_sht.Cells(n, 1), from_sht.Cells(n, cpy_frm - 1)).Value
                    to_sht.Cells(lr + j, cpy_frm) = from_sht.Cells(1, i)
                    to_sht.Cells(lr + j, cpy_frm + 1) = Trim(splt(j))
                Next j
            Else
                to_sht.Range(to_sht.Cells(lr, 1), to_sht.Cells(lr, cpy_frm - 1)).Value = from_sht.Range(from_sht.Cells(n, 1), from_sht.Cells(n, cpy_frm - 1)).Value
                to_sht.Cells(lr, cpy_frm) = from_sht.Cells(1, i)
            End If
        End With
    Next i
Next n

End Sub
Sub unlock_screen()
Application.ScreenUpdating = True
End Sub
Sub clean_selected_cells()

Dim rCell As Range
Dim txt As String
Dim rArea As Range
Dim rng As Range
Dim sht As Worksheet
Dim bk As Workbook

Set bk = ActiveWorkbook
Set sht = bk.ActiveSheet

Set rng = Selection

For Each rCell In rng.Cells
    If Not rCell.HasFormula Or rCell <> "" Then
        txt = rCell.Value2
        txt = Replace(txt, Chr(10), " ") ' line break
        txt = Replace(txt, Chr(13), " ") ' line break
        txt = Replace(txt, Chr(145), Chr(39)) ' ‘ smart single quote
        txt = Replace(txt, Chr(146), Chr(39)) ' ’ smart single quote
        txt = Replace(txt, Chr(147), Chr(34)) ' “ smart double quote
        txt = Replace(txt, Chr(148), Chr(34)) ' ” smart double quote
        txt = Replace(txt, Chr(150), "-") ' – en dash
        txt = Replace(txt, Chr(151), "-") ' — em dash
        
        Do While InStr(txt, "  ") <> 0
            txt = Replace(txt, "  ", " ")
        Loop
        
        rCell = Trim(txt)
    End If
Next rCell

End Sub
Sub single_list_multiple_vals()
Dim from_sht As Worksheet
Dim to_sht As Worksheet
Dim lc As Long
Dim lr As Long
Dim i As Long

Set from_sht = ActiveWorkbook.Sheets(1)
Set to_sht = ActiveWorkbook.Sheets(2)

lr = 2

For n = 2 To from_sht.UsedRange.Rows.Count
    lc = from_sht.Cells(n, from_sht.Columns.Count).End(xlToLeft).Column
    
    For i = 2 To lc
        to_sht.Cells(lr, 1) = from_sht.Cells(n, 1)
        to_sht.Cells(lr, 2) = from_sht.Cells(n, i)
        
        lr = lr + 1
    Next i
Next n
End Sub
Sub Add_Quotes()
Attribute Add_Quotes.VB_ProcData.VB_Invoke_Func = "q\n14"
'
' Add_Quotes Macro
'
' Keyboard Shortcut: Ctrl+q
'
    For Each rCell In Selection
        If rCell <> "" Then
            If Left(rCell, 1) <> """" Then
                rCell.Value = """" & rCell
            End If
            If Right(rCell, 1) <> """" Then
                rCell.Value = rCell & """"
            End If
        End If
    Next rCell
End Sub
Sub paste_values()
Attribute paste_values.VB_ProcData.VB_Invoke_Func = "V\n14"
'
' paste_values Macro
'
' Keyboard Shortcut: Ctrl+Shift+V
'

Dim n As Long
Dim sht As Worksheet
Dim sel_rng As Range
Dim rng_vis As Range
Dim sub_rng As Range


Set sht = ActiveSheet
Set sel_rng = Selection

If Not (sht.AutoFilterMode And sht.FilterMode) Or Not sht.FilterMode Then
    sel_rng.Copy
    sel_rng.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
Else
    Set rng_vis = Selection.SpecialCells(xlCellTypeVisible)
    For n = 1 To rng_vis.Areas.Count
        Set sub_rng = rng_vis.Areas(n)
        
        sub_rng.Copy
        sub_rng.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    Next n
End If
If Application.CutCopyMode Then Application.CutCopyMode = False
sel_rng.Select
End Sub
Sub change_case()
Attribute change_case.VB_ProcData.VB_Invoke_Func = "C\n14"
'
' change_case Macro
'
' Keyboard Shortcut: Ctrl+Shift+C
'

Dim n As Long
Dim sht As Worksheet
Dim sel_rng As Range
Dim rCell As Range
Dim sub_rng As Range
Dim chg_rng As Range

Set sht = ActiveSheet
Set sel_rng = Selection

If Not (sht.AutoFilterMode And sht.FilterMode) Or Not sht.FilterMode Then
    Set chg_rng = sel_rng
    
Else
    Set chg_rng = Selection.SpecialCells(xlCellTypeVisible)
End If

For Each rCell In chg_rng.Cells
    If UCase(rCell) = rCell Then
        rCell = LCase(rCell)
    Else
        rCell = UCase(rCell)
    End If
    
Next rCell

sel_rng.Select
End Sub
Sub merge_listed_files()

Dim wb As Workbook
Dim list As Worksheet
Dim to_sht As Worksheet

Set wb = ActiveWorkbook
Set list = wb.Sheets(1)
wb.Sheets.Add after:=wb.Sheets(wb.Sheets.Count)

Set to_sht = wb.Sheets(wb.Sheets.Count)

last_row = 1
For this_workbook = 2 To list.UsedRange.Rows.Count ' iterate through each book
    'check to see if the workbook is already open
    ' if it isn't open then open the workbook

    Application.Workbooks.Open filename:=list.Cells(this_workbook, 1), ReadOnly:=True ' open the books
    Set from_book = ActiveWorkbook

    With from_book.Sheets(1)
    If Application.WorksheetFunction.CountA(.Cells) <> 0 Then ' only do this if the sheet is not empty
        If .FilterMode Then .ShowAllData
        
        Call gpm.delete_extraneous_blank_rows_and_columns(from_book.Sheets(1))
        from_book_rows = .UsedRange.Rows.Count
        If to_sht.UsedRange.Rows.Count + from_book_rows > to_sht.Rows.Count Then
            wb.Sheets.Add after:=to_sht
            Set to_sht = wb.Sheets(wb.Sheets.Count)
            last_row = 1
        End If
        .UsedRange.Copy Destination:=to_sht.Cells(last_row, 1)
        
        last_row = to_sht.UsedRange.Rows.Count
    End If
    End With
    
    from_book.Close savechanges:=False
    Set from_book = Nothing
Next this_workbook


End Sub

Sub comparing_listed_files()
Dim n As Long
Dim list_sht As Worksheet
Dim file_name_1 As String
Dim file_name_2 As String
Dim wb_1 As Workbook
Dim wb_2 As Workbook

list_sht = ActiveSheet


For n = 2 To list_sht.UsedRange.Rows.Count
    file_name_1 = list_sht.Cells(n, 1)
    file_name_2 = list_sht.Cells(n, 2)
    
    Application.Workbooks.Open filename:=file_name_1
    Set wb_1 = ActiveWorkbook
    Application.Workbooks.Open filename:=file_name_2
    Set wb_2 = ActiveWorkbook
    
    list_sht.Cells(n, 3) = comparing(wb_1, wb_2)
Next n
End Sub
Sub testing123()
For n = 1 To ActiveWorkbook.Sheets.Count
    Sheets(n).UsedRange.Columns.Hidden = False
    
Next n
End Sub
Function get_sht_by_name(sht_name As String, ByRef this_book As Workbook) As Worksheet
Dim n As Long

For n = 1 To this_book.Sheets.Count
    If this_book.Sheets(n).Name = sht_name Then
        Set get_sht_by_name = this_book.Sheets(n)
        Exit Function
    End If
Next n
End Function
Sub comparing_files()
Dim file_name_1 As String
Dim file_name_2 As String
Dim wb_1 As Workbook
Dim wb_2 As Workbook

file_name_1 = Application.GetOpenFilename("Excel files (*.xl*),*.xl*", , _
    "Browse for the first file to be compared", MultiSelect:=False)
file_name_2 = Application.GetOpenFilename("Excel files (*.xl*),*.xl*", , _
    "Browse for the first file to be compared", MultiSelect:=False)

Application.Workbooks.Open filename:=file_name_1
Set wb_1 = ActiveWorkbook
Application.Workbooks.Open filename:=file_name_2
Set wb_2 = ActiveWorkbook

End Sub
Function comparing(wb_1 As Workbook, wb_2 As Workbook) As Boolean

If compare_books(wb_1, wb_2) Then
    comparing = True
Else
    comparing = False
End If

End Function
Function compare_books(wb_1 As Workbook, wb_2 As Workbook) As Boolean
If wb_1.Sheets.Count <> wb_2.Sheets.Count Then
    compare_books = False
    Exit Function
End If
For n = 1 To wb_1.Sheets.Count
If Compare_Sheets_1and2(wb_1.Sheets(n), wb_2.Sheets(n)) Then
    compare_books = True
Else
    compare_books = False
End If
Next n
End Function
Function Compare_Sheets_1and2(sht_1 As Worksheet, sht_2 As Worksheet) As Boolean
Dim rCell As Range

Compare_Sheets_1and2 = False

sht_1.UsedRange.Cells.Interior.Color = xlNone
sht_2.UsedRange.Cells.Interior.Color = xlNone

For Each rCell In sht_1.UsedRange.Cells
If sht_1.Range(rCell.Address).Value <> sht_2.Range(rCell.Address).Value Then
    sht_1.Range(rCell.Address).Interior.Color = RGB(255, 0, 0)
    sht_2.Range(rCell.Address).Interior.Color = RGB(255, 0, 0)
    Compare_Sheets_1and2 = True
End If
Next rCell
End Function
Sub delete_blank_rows()
Application.ScreenUpdating = False
For n = ActiveSheet.UsedRange.Rows.Count To 2 Step -1
    If Cells(n, 1) = "" Then
        Cells(n, 1).EntireRow.Delete
    End If
    If n Mod 100 = 0 Then
    Application.StatusBar = n
    End If
Next n
Application.ScreenUpdating = True
End Sub
Function resize_comments()
Dim c  As Comment
For Each c In ActiveSheet.Comments
    c.Shape.TextFrame.AutoSize = True
    If c.Shape.Width > 300 Then
        lArea = c.Shape.Width * c.Shape.Height
        c.Shape.Width = 200
        c.Shape.Height = (lArea / 200) * 1.2
    End If
Next c
End Function
Sub delete_unwanted_headers()
xl_file_name = Application.GetOpenFilename("Excel files (*.xl*),*.xl*", , _
    "Browse for file to be merged", MultiSelect:=True)

If Not IsArray(xl_file_name) Then Exit Sub

For this_book = LBound(xl_file_name) To UBound(xl_file_name)
    Application.Workbooks.Open (xl_file_name(this_book))
    Set writing_now = ActiveWorkbook
    For this_sheet = 1 To writing_now.Sheets.Count
        With writing_now.Sheets(this_sheet)
            For i = .UsedRange.Columns.Count To 1 Step -1
                If .Cells(1, i) <> "ADMIN_1" And .Cells(1, i) <> "Set_ID" And .Cells(1, i) <> "Paired_Passage" Then
                    .Columns(i).EntireColumn.Delete
                End If
            Next i
        End With
    Next this_sheet
    writing_now.Save
    writing_now.Close
Next this_book
End Sub
Sub multi_lookup()
Dim match_sht As Worksheet
Dim lookup_sht As Worksheet
Dim n As Long
Dim match_val As String
Dim lookup_col As Long
Dim pull_col As Long
Dim match_col As Long
Dim last_col As Long
Dim next_col As Long

Set lookup_sht = Sheets(1)
Set match_sht = Sheets(2)

pull_col = 7 ' column that contains the inforamtion needed will be copied from lookup_sht to match_sht
lookup_col = 1 ' column that will be filtered in lookup_sht
match_col = 1 ' column that contains the filter criteria in match_sht

Application.ScreenUpdating = False

Call gpm.delete_extraneous_blank_rows_and_columns(match_sht)
Call gpm.delete_extraneous_blank_rows_and_columns(lookup_sht)

last_col = match_sht.UsedRange.Columns.Count + 1

' loop through ever row on the match sheet
For n = 2 To match_sht.UsedRange.Rows.Count
    next_col = last_col
    match_val = match_sht.Cells(n, match_col)
    
    If match_val <> "" Then
        lookup_sht.UsedRange.AutoFilter Field:=lookup_col, Criteria1:=match_val
        
        For Each rRow In lookup_sht.UsedRange.SpecialCells(xlCellTypeVisible).Rows
            If rRow.Row <> 1 Then
                match_sht.Cells(n, next_col) = lookup_sht.Cells(rRow.Row, pull_col)
                'match_sht.Cells(n, next_col + 1) = lookup_sht.Cells(rRow.Row, pull_col + 1)
                'next_col = next_col + 2
                next_col = next_col + 1
            End If
        Next rRow
    End If

Next n
Application.ScreenUpdating = True

End Sub
Sub sort_row()
Dim lr As Long
Dim lc As Long
Dim wb As Workbook
Dim ws As Worksheet
Dim n As Long

Set wb = ActiveWorkbook
Set ws = ActiveSheet

lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row


For n = 2 To lr
    lc = ws.Cells(n, ws.Columns.Count).End(xlToLeft).Column
    
    ws.Range(ws.Cells(n, 1), ws.Cells(n, lc)).Copy
    ws.Cells(1, lc + 2).PasteSpecial Transpose:=True
    
    ws.Columns(lc + 2).Sort key1:=ws.Columns(lc + 2), order1:=xlAscending, header:=xlNo
    
    ws.Range(ws.Cells(1, lc + 2), ws.Cells(lc, lc + 2)).Copy
    ws.Cells(n, 1).PasteSpecial Transpose:=True
    
    ws.Columns(lc + 2).EntireColumn.Delete

Next n

End Sub
Sub all_cells_to_end()
Dim wb As Workbook
Dim from_sht As Worksheet
Dim to_sht As Worksheet
Dim rCell As Range
Dim n As Long

Set wb = ActiveWorkbook
Set from_sht = wb.Sheets(1)
Set to_sht = wb.Sheets(2)
n = 1

For Each rCell In from_sht.UsedRange.Cells
    If rCell.Value2 <> "" Then
        to_sht.Cells(n, 1) = Trim(rCell.Value2)
        n = n + 1
    End If
Next rCell
End Sub
Sub NewWorkbook()
Attribute NewWorkbook.VB_ProcData.VB_Invoke_Func = "B\n14"
'
' Macro1 Macro
'
' Keyboard Shortcut: Ctrl+Shift+B
'

Application.Workbooks.Add
End Sub
Sub multi_merge()
Dim n As Long
Dim from_bk As Workbook
Dim to_bk As Workbook
Dim from_sht As Worksheet
Dim to_sht As Worksheet
Dim to_match_col As Long
Dim to_last_col As Long
Dim from_lookup_col As Long
Dim copy_rows As Long


to_match_col = 7
from_lookup_col = 8

Set from_bk = Workbooks("CAASPP Mapping Document.xlsx")
Set from_sht = from_bk.ActiveSheet
Set to_bk = Workbooks("CST_SCIENCE_SAIB_IBIS_Matched_V2.xlsx")
Set to_sht = to_bk.ActiveSheet
Application.ScreenUpdating = False

to_last_col = to_sht.Cells(1, to_sht.Columns.Count).End(xlToLeft).Column + 1

For n = to_sht.Cells(to_sht.Rows.Count, 1).End(xlUp).Row To 2 Step -1
    from_sht.UsedRange.AutoFilter Field:=from_lookup_col, Criteria1:=to_sht.Cells(n, to_match_col)
    copy_rows = gpm.count_rows_in_range(from_sht.UsedRange.SpecialCells(xlCellTypeVisible)) - 1
    
    Run gpm.copy_row_n_times(to_sht, copy_rows, n)
    from_sht.UsedRange.Copy Destination:=to_sht.Cells(n, to_last_col)
    to_sht.Rows(n).EntireRow.Delete
Next n

Application.ScreenUpdating = True
End Sub
Sub enemy_counts()
Dim n As Long
Dim lr As Long
Dim sht As Worksheet
Dim sum As Worksheet
Dim i As Long
Dim slr As Long

i = 2
Set sum = ActiveWorkbook.Sheets("summ")

For n = 1 To ActiveWorkbook.Sheets.Count - 1
    Set sht = ActiveWorkbook.Sheets(n)
    With sht
    lr = .Cells(.Rows.Count, 1).End(xlUp).Row
    .Range(.Cells(2, 3), .Cells(lr, 3)).Copy Destination:=.Cells(lr + 1, 4)
    .Columns(4).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=sum.Cells(i, 2), unique:=True
    End With
    
    With sum
        slr = .Cells(.Rows.Count, 2).End(xlUp).Row
        .Range(.Cells(i, 1), .Cells(slr, 1)) = sht.Cells(2, 2)
    End With
    i = slr + 1

Next n


End Sub
Sub Reorder_Columns_by_Name()
Dim wb As Workbook
Dim ws As Workbook
Dim xl_file_name As Variant
Dim this_workbook As Long
Dim sht As Long
Dim col_arr() As Variant
Dim col As Long

xl_file_name = Application.GetOpenFilename("Excel files (*.xl*),*.xl*", , _
    "Browse for file to be merged", MultiSelect:=True)
Application.DisplayAlerts = False
Application.ScreenUpdating = False

col_arr = Array("CLIP_UIN", "CLIP_NAME", "PAS", "TEST_CODE", "TEST_FORM", "ITEM_UIN", "1", "2", "3", "4", "7", "13")


If IsArray(xl_file_name) Then
    For this_workbook = LBound(xl_file_name) To UBound(xl_file_name) ' iterate through each book
        Application.Workbooks.Open filename:=xl_file_name(this_workbook), ReadOnly:=True ' open the books
        Set wb = ActiveWorkbook
        
        For Each ws In wb.Sheets
            For n = LBound(col_arr) To UBound(col_arr)
                col = get_column_by_header(ws, CStr(col_arr(n)))
                
                If col = 0 Then
                    ws.Columns(n).Insert Shift:=xlToRight
                    ws.Cells(1, n) = col_arr(n)
                ElseIf n <> col Then
                    ws.Columns(col).Cut
                    ws.Columns(n).Insert Shift:=xlToRight
                    Application.CutCopyMode = False
                End If
            Next n
        Next ws

    Next this_workbook
End If
Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub
Sub convert_to_xlsx()
Dim xl_file_name As Variant
Dim this_workbook As Long
Dim save_book As Workbook
Dim save_path As String

xl_file_name = Application.GetOpenFilename("Excel files (*.xls),*.xls", , _
    "Browse for file to be merged", MultiSelect:=True)
Application.DisplayAlerts = False
Application.ScreenUpdating = False

If IsArray(xl_file_name) Then
    For this_workbook = LBound(xl_file_name) To UBound(xl_file_name) ' iterate through each book
        Application.Workbooks.Open filename:=xl_file_name(this_workbook), ReadOnly:=True, Editable:=True ' open the books
        Set save_book = ActiveWorkbook
        save_path = save_book.path & "/" & Replace(save_book.Name, ".xls", ".xlsx")
        
        save_book.SaveAs filename:=save_path, FileFormat:=xlOpenXMLWorkbook
        save_book.Close
    Next this_workbook
End If

Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub
Sub merge_rows_by_delim()
Dim uni_col As Long
Dim this_val As String
Dim iRow As Long
Dim iCol As Long
Dim lc As Long
Dim wb As Workbook
Dim ws As Worksheet
Dim to_sht As Worksheet
Dim to_row As Long
Dim last_val As String

uni_col = 1

Set wb = ActiveWorkbook
Set ws = wb.Sheets(1)
Set to_sht = wb.Sheets(2)
to_row = 1

lc = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

ws.Rows(1).Copy Destination:=to_sht.Rows(1)

For iRow = 2 To lr
    this_val = ws.Cells(iRow, uni_col)
    
    If this_val <> last_val Then
        to_row = to_row + 1
        to_sht.Cells(to_row, uni_col) = this_val
    End If
    
    For iCol = 2 To lc
        If iCol <> uni_col Then
            to_sht.Cells(to_row, iCol) = append_val(to_sht.Cells(to_row, iCol), ws.Cells(iRow, iCol), , True)
        End If
    Next iCol
    last_val = this_val
Next iRow

End Sub
Function append_val(text As String, append_text As String, Optional delim = "|", Optional unique = "False") As String
If text = "" Then
    text = append_text
Else
    If text <> append_text And InStr(text, append_text & delim) = 0 And InStr(text, delim & append_text) = 0 And unique Then
        text = text & delim & append_text
    ElseIf Not unique Then
        text = text & delim & append_text
    End If
End If
append_val = text
End Function
Sub list_sheet_names()
Dim wb As Workbook
Dim sht As Worksheet
Dim report_sht As Worksheet
Dim cnt As Long

cnt = 1
Set wb = ActiveWorkbook
Set report_sht = wb.ActiveSheet

For Each sht In wb.Sheets
    If sht.Name <> report_sht.Name Then
        report_sht.Cells(cnt, 1) = sht.Name
        cnt = cnt + 1
    End If
Next sht
End Sub

