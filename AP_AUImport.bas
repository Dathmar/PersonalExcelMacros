Attribute VB_Name = "AP_AUImport"
Option Explicit
Option Base 0
Private Type Au
    type As String
    id As String
    test As String
    measure As String
    form As String
    lookup_code As String
End Type
Dim error_list As String
Sub ignore_all()
Dim wb As Workbook
Dim sht As Worksheet
Dim current_test As String
Dim r As Long

Set wb = ActiveWorkbook
Application.Calculation = xlCalculationManual

' loop through all the sheets
For Each sht In wb.Sheets
    current_test = get_current_test(sht.Name)
    If current_test <> "null" Then
        For r = 4 To sht.UsedRange.Rows.Count
            sht.Cells(r, 2) = "Ignore"
        Next r
    End If
Next sht

Application.Calculation = xlAutomatic

End Sub
Sub Create_AU_Item_Import()

Dim wb As Workbook
Dim sht As Worksheet
Dim au_details As Worksheet
Dim au_itm As Worksheet
Dim au_grid As Worksheet
Dim current_test As String
Dim r As Long
Dim c As Long
Dim item_range As Range
Dim lookup_code As String
Dim base_au_id As String
Dim import_sht As Worksheet
Dim paste_rng As Range
Dim lr As Long
Dim new_lr As Long
Dim new_au_id As String
Dim new_au As Au
Dim base_au As Au

Dim form_col As Long
Dim process_col As Long
Dim start_col As Long
Dim start_row As Long
Dim tmt_row As Long
Dim err As String

Dim mcq_count As Long
Dim frq_count As Long
Dim cnt As Long

form_col = 1
process_col = 2
start_col = 3
start_row = 4
tmt_row = 3

error_list = ""

Set wb = ActiveWorkbook
Set import_sht = wb.Sheets("Import Template By Accnum")

import_sht.Range(import_sht.Cells(2, 1), import_sht.Cells(import_sht.Rows.Count, import_sht.Columns.Count)).EntireRow.Delete

Call remove_set_leaders

Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False

import_sht.Range(import_sht.Rows(2), import_sht.Rows(import_sht.Rows.Count)).Delete
' loop through all the sheets
For Each sht In wb.Sheets
    current_test = get_current_test(sht.Name)
    If current_test <> "null" Then
        For r = start_row To get_last_row(sht, form_col) - 1
        mcq_count = 0
        frq_count = 0
        
        If Left(LCase(sht.Cells(r, process_col)), 3) <> "mcq" And LCase(sht.Cells(r, process_col)) <> "ignore" Then
        sht.Cells(r, process_col) = "In Process"
        For c = start_col To get_last_column(sht, r)
            lookup_code = process_lookup_code(sht.Cells(r, c))
            base_au = get_base_au_id_from_lookup_code(current_test, lookup_code)
            base_au.form = sht.Cells(r, 1)
            base_au.lookup_code = lookup_code
                
            new_au.type = sht.Cells(tmt_row, c)
            new_au.test = base_au.test
            new_au.measure = base_au.measure
            new_au.form = base_au.form
            new_au.id = get_new_au_id(new_au)
            new_au.lookup_code = base_au.lookup_code
            
            If base_au.id <> "" And new_au.id <> "" Then
                lr = get_last_row(import_sht)
                Set paste_rng = import_sht.Cells(lr, 1)
                Call copy_items_for_au_number(base_au.id, paste_rng)
                
                new_lr = get_last_row(import_sht)
                
                Call add_seq(lr, new_lr, import_sht, new_au.id)
                Call add_use_codes(lr, new_lr, import_sht, base_au)
            
                
                If new_lr - lr = 0 Then
                    err = "No items found in subblock AU " & base_au.id ' Base AU for subblock not found.
                    error_list = concat_string(error_list, err)
                Else
                    cnt = Application.WorksheetFunction.CountIf(import_sht.Range(import_sht.Cells(lr, 7), import_sht.Cells(new_lr, 7)), "Y")
                    If Left(new_au.type, 1) = "F" Then
                        frq_count = frq_count + cnt
                    Else
                        mcq_count = mcq_count + cnt
                    End If
                End If
                
            End If
            If base_au.id = "" Then
                err = "No subblock AU found for subblock " & sht.Cells(r, c) ' Base AU for subblock not found.
                error_list = concat_string(error_list, err)
            End If
            If new_au.id = "" Then
                err = "No form AU found for subblock " & sht.Cells(r, c) ' Base AU for subblock not found.
                error_list = concat_string(error_list, err)
            End If
        Next c
        
        error_list = concat_string(error_list, "MCQ Items = " & mcq_count)
        error_list = concat_string(error_list, "FRQ Items = " & frq_count)
        
        Application.ScreenUpdating = False
        sht.Cells(r, process_col) = error_list
        error_list = ""
        Application.ScreenUpdating = True
        End If
        Next r
    End If
Next sht
Application.ScreenUpdating = True
Application.Calculation = xlAutomatic
End Sub
Function add_use_codes(start_row As Long, end_row As Long, sht As Worksheet, this_au As Au)
Dim usecode As String
Dim wb As Workbook
Dim import_sht As Worksheet
Dim r As Long

Set wb = ActiveWorkbook
Set import_sht = wb.Sheets("Import Template By Accnum")

usecode = get_usecode(this_au)

For r = start_row To end_row - 1
    If import_sht.Cells(r, 7) = "Y" Then
        import_sht.Cells(r, 6) = usecode
    Else
        import_sht.Cells(r, 6) = "Unused Item"
    End If
Next

End Function
Function get_usecode(this_au As Au) As String
Dim wb As Workbook
Dim tab_to_test As Worksheet
Dim use_code_mapping As Worksheet
Dim form_code As String
Dim usecode_type As String
Dim rng As Range
Dim sRange As Range
Dim usecode As String

Set wb = ActiveWorkbook
Set tab_to_test = wb.Sheets("Tab to Test")
Set use_code_mapping = wb.Sheets("Use Code Mapping")

tab_to_test.AutoFilterMode = False

Set sRange = tab_to_test.UsedRange
sRange.AutoFilter Field:=2, Criteria1:="=" & this_au.test

Set rng = sRange.Offset(1, 0).SpecialCells(xlCellTypeVisible)(1)

If Not rng Is Nothing Then
    usecode_type = tab_to_test.Cells(rng.Row, 4)
End If

tab_to_test.AutoFilterMode = False


If InStr(this_au.form, "F0") <> 0 Then
    form_code = "F0"
ElseIf InStr(this_au.form, "G0") <> 0 Then
    form_code = "G0"
ElseIf usecode_type = "Eng Lang" And (InStr(this_au.form, "X3") Or InStr(this_au.form, "X17") Or InStr(this_au.form, "X21") Or InStr(this_au.form, "X26") Or InStr(this_au.form, "X27") Or InStr(this_au.form, "X30")) Then
    form_code = this_au.form
Else
    form_code = Left(this_au.form, 1)
End If

use_code_mapping.AutoFilterMode = False

Set sRange = use_code_mapping.UsedRange
sRange.AutoFilter Field:=1, Criteria1:="=" & usecode_type
sRange.AutoFilter Field:=2, Criteria1:="=" & form_code
sRange.AutoFilter Field:=3, Criteria1:="=" & Left(this_au.lookup_code, 1)

Set rng = sRange.Offset(1, 0).SpecialCells(xlCellTypeVisible)(1)

If Not rng Is Nothing Then
    usecode = use_code_mapping.Cells(rng.Row, 4)
End If

If usecode = "" Then
    usecode = "Operational"
End If
get_usecode = usecode
use_code_mapping.AutoFilterMode = False

End Function
Function concat_string(to_str As String, concat_str As String, Optional del As String = ", ") As String
If to_str = "" Then
    concat_string = concat_str
Else
    If InStr(to_str, concat_str) = 0 Then
        concat_string = to_str & del & concat_str
    End If
End If
End Function
Sub remove_set_leaders()
Dim wb As Workbook
Dim au_itm As Worksheet
Dim lr As Long
Dim r As Long
Dim delete_rng As Range

Set wb = ActiveWorkbook
Set au_itm = wb.Sheets("AU Item")
Set delete_rng = Nothing
lr = get_last_row(au_itm, 1)

For r = 2 To lr
    If au_itm.Cells(r, 4) = au_itm.Cells(r, 5) Then
        If delete_rng Is Nothing Then
            Set delete_rng = au_itm.Rows(r)
        Else
            Set delete_rng = Union(delete_rng, au_itm.Rows(r))
        End If
    End If
Next r
If Not delete_rng Is Nothing Then
    delete_rng.EntireRow.Delete
End If
End Sub
Function get_new_au_id(new_au As Au) As String
Dim wb As Workbook
Dim lookup_sht As Worksheet
Dim rng As Range
Dim sRange As Range

Set wb = ActiveWorkbook
Set lookup_sht = wb.Sheets("To AUs")

lookup_sht.AutoFilterMode = False

Set sRange = lookup_sht.UsedRange
sRange.AutoFilter Field:=1, Criteria1:="=" & new_au.test
sRange.AutoFilter Field:=4, Criteria1:="=*" & new_au.form & "*"
sRange.AutoFilter Field:=6, Criteria1:="=" & new_au.type


Set rng = sRange.Offset(1, 0).SpecialCells(xlCellTypeVisible)(1)

If Not rng Is Nothing Then
    get_new_au_id = lookup_sht.Cells(rng.Row, 2)
Else
    get_new_au_id = "null"
End If

lookup_sht.AutoFilterMode = False

End Function
Function get_new_au_type(base_au As Au) As String
Dim start_char As String

start_char = Left(base_au.type, 1)
If start_char = "M" Then
    get_new_au_type = "M"
Else
    get_new_au_type = "F"
End If

End Function
Function add_seq(start_row As Long, end_row As Long, sht As Worksheet, last_au_id As String)
' remove set leaders before sequencing

Dim cnt As Long
Dim r As Long

If sht.Cells(start_row - 1, 2) = last_au_id Then
    cnt = sht.Cells(get_last_row(sht, 3) - 1, 3) + 1
Else
    cnt = 1
End If

For r = start_row To end_row - 1
    sht.Cells(r, 2) = last_au_id
    If sht.Cells(r, 7) = "Y" Then
        sht.Cells(r, 3) = cnt
        cnt = cnt + 1
    End If
Next r

End Function
Function get_last_row(sht As Worksheet, Optional check_column As Long = 1) As Long
get_last_row = sht.Cells(sht.Rows.Count, check_column).End(xlUp).Row + 1
End Function
Function get_last_column(sht As Worksheet, Optional check_row As Long = 1) As Long
get_last_column = sht.Cells(check_row, sht.Columns.Count).End(xlToLeft).Column
End Function
Function copy_filtered_range_no_headers(sht As Worksheet, to_rng As Range) As Range
Dim rTable As Range
 
Set rTable = sht.AutoFilter.Range ' MODIF
Set rTable = rTable.Resize(rTable.Rows.Count - 1)

'Move new range down to start at the fisrt data row.
Set rTable = rTable.Offset(1)
rTable.Copy Destination:=to_rng

Application.CutCopyMode = False
End Function
Function process_lookup_code(lookup_code As String) As String
Dim u_array As Variant

If Left(lookup_code, 1) = "U" Then
    u_array = Split(lookup_code, "-")
    lookup_code = u_array(0) & u_array(1)
ElseIf Left(lookup_code, 1) <> "F" Then
    lookup_code = Replace(lookup_code, "-", "")
    lookup_code = Replace(lookup_code, " ", "")
End If

process_lookup_code = lookup_code
End Function
Function copy_items_for_au_number(au_number As String, to_rng As Range)
Dim wb As Workbook
Dim lookup_sht As Worksheet
Dim rng As Range
Dim sRange As Range

Set wb = ActiveWorkbook
Set lookup_sht = wb.Sheets("AU Item")

lookup_sht.AutoFilterMode = False

Set sRange = lookup_sht.UsedRange
sRange.AutoFilter Field:=1, Criteria1:="=" & au_number

Call copy_filtered_range_no_headers(lookup_sht, to_rng)

End Function
Function get_base_au_id_from_lookup_code(test_name As String, lookup_code As String) As Au
Dim wb As Workbook
Dim lookup_sht As Worksheet
Dim rng As Range
Dim sRange As Range
Dim au_settings As Au

Set wb = ActiveWorkbook
Set lookup_sht = wb.Sheets("Base AU Metadata")

lookup_sht.AutoFilterMode = False

Set sRange = lookup_sht.UsedRange
sRange.AutoFilter Field:=1, Criteria1:="=" & test_name
sRange.AutoFilter Field:=15, Criteria1:="=" & lookup_code

Set rng = sRange.Offset(1, 0).SpecialCells(xlCellTypeVisible)(1)

au_settings.test = test_name

If Not rng Is Nothing Then
    au_settings.id = lookup_sht.Cells(rng.Row, 2)
    au_settings.measure = lookup_sht.Cells(rng.Row, 5)
    au_settings.type = lookup_sht.Cells(rng.Row, 6)
Else
    au_settings.id = "null"
    au_settings.measure = "null"
    au_settings.type = "null"
End If
get_base_au_id_from_lookup_code = au_settings
lookup_sht.AutoFilterMode = False

End Function
Function get_current_test(sht_name As String) As String
Dim wb As Workbook
Dim test_sht As Worksheet
Dim test_row As Range

Set wb = ActiveWorkbook
Set test_sht = wb.Sheets("Tab to Test")

Set test_row = test_sht.Columns(1).Find(What:=sht_name)

If Not test_row Is Nothing Then
    If test_sht.Cells(test_row.Row, 3) = True Then
        get_current_test = test_sht.Cells(test_row.Row, 2)
        Exit Function
    End If
End If
get_current_test = "null"
End Function


