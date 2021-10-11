Attribute VB_Name = "Module1"
Sub sheet_to_csv()
Dim to_wb As Workbook

Set wb = ActiveWorkbook
Application.DisplayAlerts = False
For Each sht In wb.Sheets
    If IsNumeric(gpm.left_before(sht.Name, "-")) Then
        Application.Workbooks.Add
        Set to_wb = ActiveWorkbook
        Set del_sht = ActiveWorkbook.Sheets(1)
        
        sht.Copy before:=del_sht
        del_sht.Delete
        to_wb.SaveAs filename:="C:\split\" & sht.Name & ".csv", FileFormat:=xlCSV
        to_wb.Close
    End If
    
Next sht

Application.DisplayAlerts = True
End Sub
Sub convert_to_url()
Dim wb As Workbook
Dim sht As Worksheet
Dim lr As Long
Dim r As Long
Dim h As Hyperlink
Set wb = ActiveWorkbook
Set sht = wb.ActiveSheet

For r = 2 To sht.UsedRange.Rows.Count
    For Each h In ActiveSheet.Hyperlinks
        If Cells(r, 1).Address = h.Range.Address Then
            Cells(r, 2) = h.Address
        End If
    Next h
Next r

End Sub
Sub list_characters()
Dim wb As Workbook
Dim sht As Worksheet
Dim r As Long
Dim c As Long
Dim letter As String
Dim lst As String
Dim chk_txt As String

Set wb = ActiveWorkbook
Set sht = wb.Sheets(1)

For r = 2 To sht.UsedRange.Rows.Count
    For c = 3 To 52
        chk_txt = sht.Cells(r, c)
        brk = 0
        Do While chk_txt <> "" Or brk > 100
            letter = Left(chk_txt, 1)
            If Len(Replace(lst, letter, "")) = Len(lst) Then
                lst = lst & letter
            End If
            
            chk_txt = Right(chk_txt, Len(chk_txt) - 1)
            chk_txt = Replace(chk_txt, letter, "")
            
            brk = brk + 1
        Loop
    Next c
    sht.Cells(r, 2) = lst
    lst = ""
Next r

End Sub
Sub round_all_num()
Dim wb As Workbook
Dim ws As Worksheet
Dim r As Long
Dim c As Long

Set wb = ActiveWorkbook
Set ws = ActiveSheet

For r = 2 To ws.UsedRange.Rows.Count
For c = 7 To 15
    If IsNumeric(ws.Cells(r, c)) And ws.Cells(r, c) <> "" Then
        ws.Cells(r, c) = Application.WorksheetFunction.Round(ws.Cells(r, c), 4)
    End If
Next c
Next r
End Sub

Sub one_to_one_next_sht()
Dim wb As Workbook
Dim from_sht As Worksheet
Dim to_sht As Worksheet
Dim to_row As Long

Set wb = ActiveWorkbook
Set from_sht = wb.Sheets(1)
Set to_sht = wb.Sheets(2)

to_row = 2

For r = 2 To from_sht.UsedRange.Rows.Count
    For c = 1 To 28
        to_sht.Cells(to_row, 1) = from_sht.Cells(r, c)
        to_sht.Cells(to_row, 2) = from_sht.Cells(r, c + 28)
        
        to_row = to_row + 1
    Next c
Next r
End Sub
Sub STAAR_Tasks_from_Raw_Data()
Dim wb As Workbook
Dim ws As Worksheet
Dim tb As Workbook
Dim ts As Worksheet
Dim to_bk As Workbook
Dim to_sht As Worksheet
Dim tsk_col As Integer
Dim n As Integer

Dim lr As Long
Dim fltr_lst() As String

Set wb = ActiveWorkbook
Set ws = wb.Sheets("Raw Data")

Application.Workbooks.Open filename:="\\ETSLAN.ORG\SAO\FS_K12_DATA_03\TestDev\Content Folders\ADSC\Asher\TX\TaskListTemplate\TaskListTemplate.xlsx", _
    ReadOnly:=True
Application.Calculation = xlCalculationManual

Set tb = ActiveWorkbook

' check user name and select correct sheet
If LenB(Dir("C:\Users\destrada")) Then
    Set ts = tb.Sheets("Daniel")
Else
    Set ts = tb.Sheets("Branden")
End If

lr = ts.Cells(ts.Rows.Count, 1).End(xlUp).Row

ReDim fltr_lst(0 To lr - 2)

For n = 2 To lr
    fltr_lst(n - 2) = ts.Cells(n, 1)
Next n

tb.Close savechanges:=False

tsk_col = 0
ws.Cells.Columns.Hidden = False
ws.Cells.Rows.Hidden = False
lc = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

For n = 1 To lc
    If ws.Cells(1, n) = "Task_Name" Then tsk_col = n
Next n

If tsk_col = 0 Then
    MsgBox "Could not find Task Name column."
    Exit Sub
End If

ws.UsedRange.AutoFilter Field:=tsk_col, Criteria1:=fltr_lst, Operator:=xlFilterValues

Application.Workbooks.Add

Set to_bk = ActiveWorkbook
Set to_sht = to_bk.ActiveSheet

ws.UsedRange.SpecialCells(xlCellTypeVisible).Copy Destination:=to_sht.Cells(1, 1)
Application.Calculation = xlCalculationAutomatic

End Sub
Sub fix_dates()
Dim val As String
Dim this_sht As Worksheet

For sht = 1 To ActiveWorkbook.Sheets.Count
Set this_sht = ActiveWorkbook.Sheets(sht)
For c = 3 To 4
    If InStr("Packages in Admin", this_sht.Name) = 0 Then
    For n = 2 To this_sht.UsedRange.Rows.Count
        If this_sht.Cells(n, c) <> "" Then
            val = CDate(this_sht.Cells(n, c))
            val = Format(val, "mm/dd/yy hh:mm:ss")
            this_sht.Cells(n, c).NumberFormat = "@"
            this_sht.Cells(n, c) = val
            
        End If
    Next n
    End If
Next c
Next sht
End Sub
Function clean_blb(rng As Range) As String

Dim txt As String

txt = CStr(rng.Value)

txt = Replace(txt, "/em&gt;", "")
txt = Replace(txt, "/strong&gt;", "")
txt = Replace(txt, "strong&gt;", "")
txt = Replace(txt, "em&gt;", "")
txt = Replace(txt, "&lt;/p&gt;", "")
txt = Replace(txt, "&lt;p&gt;", "")
txt = Replace(txt, "&amp;nbsp;", " ")


txt = Replace(txt, "#39;", "'")

clean_blb = Trim(txt)


End Function
Sub EndofScore()
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
Dim to_sht As Worksheet
Dim n As Long
Dim check_passwords As Variant

xl_file_name = Application.GetOpenFilename("Excel files (*.xl*),*.xl*", , _
    "Browse for file to be merged", MultiSelect:=True)
Application.DisplayAlerts = False
Application.ScreenUpdating = False

If IsArray(xl_file_name) Then
    For this_workbook = LBound(xl_file_name) To UBound(xl_file_name) ' iterate through each book
        Application.Workbooks.Open filename:=xl_file_name(this_workbook), ReadOnly:=False, Editable:=True ' open the books
        Set from_book = ActiveWorkbook
        Set to_sht = from_book.Sheets(3)
        
        to_sht.Cells(2, 1) = "LISTSCL"
        to_sht.Cells(3, 1) = "READSCL"
        
        from_book.Close savechanges:=True
    Next this_workbook

End If
Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub

