Attribute VB_Name = "Misc"
Sub rename_headers()
Dim c As Long
Dim wb As Workbook
Dim sht As Worksheet



xl_file_name = Application.GetOpenFilename("Excel files (*.x*),*.x*", , _
    "Browse for file to be merged", MultiSelect:=True)
Application.DisplayAlerts = False
Application.ScreenUpdating = False

If IsArray(xl_file_name) Then
    For this_workbook = LBound(xl_file_name) To UBound(xl_file_name) ' iterate through each book
        Application.Workbooks.Open filename:=xl_file_name(this_workbook), ReadOnly:=False
        
        Set wb = ActiveWorkbook
        
        
        For this_sht = 1 To wb.Sheets.Count
            Set sht = wb.Sheets(this_sht)
            For c = 1 To sht.Cells(1, sht.Columns.Count).End(xlToLeft).Column
                Select Case LCase(sht.Cells(1, c))
                    Case Is = LCase("a")
                        sht.Cells(1, c) = "IRT_A"
                    Case Is = LCase("b")
                        sht.Cells(1, c) = "IRT_B"
                    Case Is = LCase("c")
                        sht.Cells(1, c) = "IRT_C"
                    Case Is = LCase("d1")
                        sht.Cells(1, c) = "IRT_D1"
                    Case Is = LCase("d2")
                        sht.Cells(1, c) = "IRT_D2"
                    Case Is = LCase("d3")
                        sht.Cells(1, c) = "IRT_D3"
                    Case Is = LCase("d4")
                        sht.Cells(1, c) = "IRT_D4"
                    Case Is = LCase("d5")
                        sht.Cells(1, c) = "IRT_D5"
                    Case Is = LCase("d6")
                        sht.Cells(1, c) = "IRT_D6"
                    Case Is = LCase("itemid")
                        sht.Cells(1, c) = "UIN"
                End Select
                        
            Next c
        Next this_sht
        
        wb.Close savechanges:=True
    Next this_workbook
End If
Application.DisplayAlerts = True
Application.ScreenUpdating = True


End Sub
Sub TX_Essay_Scores()
Dim n As Long
Dim i As Long
Dim sht As Worksheet
Dim bk As Workbook
Dim sert As Integer
Set bk = ActiveWorkbook
Set sht = bk.ActiveSheet

For n = sht.Cells(sht.Rows.Count, 1).End(xlUp).Row To 3 Step -1
    sert = 0
    For i = 3 To 5
        If sert <> 0 Then
            sht.Rows(n + sert).Insert
            sht.Rows(n).Copy Destination:=sht.Rows(n + sert)
            sht.Range(sht.Cells(n + sert, 3), sht.Cells(n + sert, 5)).ClearContents
        End If
        
        If sht.Cells(n, i) = "-1" Then
            sht.Cells(n + sert, i) = "Y"
        Else
            sht.Cells(n + sert, i) = "N"
        End If
        sert = sert + 1
    Next i
    sht.Range(sht.Cells(n, 4), sht.Cells(n, 5)).ClearContents
Next n
End Sub
Sub Switch_Numbers()
Attribute Switch_Numbers.VB_ProcData.VB_Invoke_Func = "N\n14"
    If Application.ReferenceStyle = xlA1 Then
        Application.ReferenceStyle = xlR1C1
    Else
        Application.ReferenceStyle = xlA1
    End If
End Sub
Sub split_cells_to_multiple_rows()
Dim n As Long
Dim splt_sht As Worksheet
Dim to_sht As Worksheet
Dim bk As Workbook
Dim del As String
Dim splt_col As Long
Dim splt_vals As Variant
Dim lr As Long
Dim i As Long
Dim to_col As Long

splt_col = 4
to_col = 7
del = ","

Set bk = ActiveWorkbook
Set splt_sht = bk.Sheets(1)
Set to_sht = bk.Sheets(2)

' copy header
splt_sht.Rows(1).Copy Destination:=to_sht.Rows(1)


For n = 2 To splt_sht.UsedRange.Rows.Count
    lr = to_sht.Cells(to_sht.Rows.Count, 1).End(xlUp).Row
    
    If InStr(splt_sht.Cells(n, splt_col), del) = 0 Then
        ' just copy row over
        splt_sht.Rows(n).Copy Destination:=to_sht.Rows(lr + 1)
    Else
        ' copy row n times and fill in split values
        splt_vals = Split(splt_sht.Cells(n, splt_col), del)
        
        For i = LBound(splt_vals) To UBound(splt_vals)
            splt_sht.Rows(n).Copy Destination:=to_sht.Rows(lr + 1 + i)
            to_sht.Cells(lr + 1 + i, to_col) = Trim(splt_vals(i))
        Next i
        
    End If
    
Next n
End Sub
Sub char_list()
Dim rCell As Range
Dim sChar As String
Dim sCharList As String
Dim i As Long

sCharList = ""

For Each rCell In ActiveSheet.UsedRange
If rCell.Value <> "" Then
    For i = 1 To Len(rCell.Value)
        sChar = Mid(rCell.Value, i, 1)
        If InStr(sCharList, sChar) = 0 Then sCharList = sCharList & sChar
    Next i
End If
Next rCell

Debug.Print sCharList

End Sub
Sub AAA_Open_math_ML()
Call open_ml("Math")
End Sub
Sub AAA_Open_ELA_ML()
Call open_ml("ELA")
End Sub
Function open_ml(content As String)

If content = "ELA" Then
    ml_path = "\\ETSLAN.ORG\SAO\FS_K12_Data_03$\TestDev\Work_Active\PARCC\PARCC 6\Form Changes\Masterlist Functions\Hidden_ML_ELA.xlsx"
    ml_name = "Hidden_ML_ELA.xlsx"
Else
    ml_path = "\\ETSLAN.ORG\SAO\FS_K12_DATA_03\TestDev\Work_Active\PARCC\PARCC 6\Form Changes\Masterlist Functions\Hidden_ML_Math.xlsx"
    ml_name = "Hidden_ML_Math.xlsx"
End If

' check to make sure the file can be opened without read only try 5 times waiting 3 seconds between tries.
ml_pass = "8e0e86ea7fcc2499b78fd3270b0dcbbf80ca9ad3cfc7e9b21003545599a818015093abea539deccf7e403ac98af4a012a8492af130e17a4ee0bcb80dfc198663a1ed15576e8ab06c583043a65efb5bddd94946efcb62e1795c8744aa2db8e2c5c486cd50"
Application.Workbooks.Open filename:=ml_path, ReadOnly:=False, Password:=ml_pass
ml_pass = ""

Set ml = ActiveWorkbook
Set ml = Nothing
End Function
Sub add_comment()
Set cmt_sht = ActiveWorkbook.Sheets(1)
Application.Workbooks.Open ("C:\Users\adanner\Desktop\Asher\1 - WCAP\WCAP 8 gr standards.xlsx")
Set wrk_sht = ActiveWorkbook.Sheets(1)
For n = 2 To cmt_sht.UsedRange.Rows.Count
    itm_spc = cmt_sht.Cells(n, 2)
    For i = 2 To wrk_sht.UsedRange.Rows.Count
        If wrk_sht.Cells(i, 2).Value = itm_spc Then
            If cmt_sht.Cells(n, 2).Comment Is Nothing Then cmt_sht.Cells(n, 2).AddComment
            cmt_sht.Cells(n, 2).Comment.text CStr(wrk_sht.Cells(i, 3).Value)
        End If
    Next i
Next n

Application.ScreenUpdating = True
End Sub
Sub RASE_itemlist()
Dim this_sht As Worksheet
Dim grade As String
Dim form As String
Dim n As Long
Dim dash_pass() As String
Dim i As Long

'variable setting
i = 0
Set this_sht = ActiveSheet
grade = Trim(CStr(this_sht.Cells(2, 6).Value)) 'if grade column changes change this
form = Trim(CStr(this_sht.Cells(2, 2).Value)) 'if form column changes change this
ReDim dash_pass(0 To 1)
'create array of all passages and delete row
'this is based off the seqnum being ~
For n = this_sht.UsedRange.Rows.Count To 2 Step -1
    If this_sht.Cells(n, 3) = "~" Then
        ReDim Preserve dash_pass(0 To i)
        dash_pass(i) = Trim(this_sht.Cells(n, 1))
        i = i + 1
        this_sht.Rows(n).EntireRow.Delete
    End If
Next n

'check for combined sets
For n = 2 To this_sht.UsedRange.Rows.Count
    this_pass = Trim(this_sht.Cells(n, 7))
    If this_pass <> last_pass Then
        For i = LBound(dash_pass) To UBound(dash_pass)
            If dash_pass(i) = this_pass Then
                dash_pass(i) = ""
            End If
        Next i
    End If
    last_pass = this_pass
Next n

n = 0
For i = LBound(dash_pass) To UBound(dash_pass)
    If dash_pass(i) <> "" Then
        dash_pass(n) = dash_pass(i)
        n = n + 1
    End If
Next i
If n > 0 Then
    ReDim Preserve dash_pass(0 To n - 1)
Else
    ReDim Preserve dash_pass(0 To 0)
End If
For i = LBound(dash_pass) To UBound(dash_pass)
    n = this_sht.UsedRange.Rows.Count + 1
    If dash_pass(i) <> "" Then
        this_sht.Cells(n, 1) = dash_pass(i)
        this_sht.Cells(n, 7) = "~"
    End If
Next i
'save sheet as final xls
this_sht.SaveAs filename:=this_sht.Parent.path & "\" & grade & "_" & form & ".xls", FileFormat:=56

n = 2
this_pass = ""
last_pass = ""

Do While this_sht.Cells(n, 7) <> ""
    this_pass = Trim(this_sht.Cells(n, 7))
    If this_pass <> last_pass And this_pass <> "~" Then
        this_sht.Rows(n).EntireRow.Insert
        this_sht.Cells(n, 1) = this_pass & "_drct" & "," & this_pass
    ElseIf this_pass <> "~" Then
        this_sht.Cells(n, 1) = this_sht.Cells(n, 1) & "," & this_pass
    End If
    last_pass = this_pass
    n = n + 1
Loop
this_txt = this_sht.Parent.path & "\itemlist.txt"
fnum = FreeFile()
Open this_txt For Output As fnum
For n = 2 To this_sht.UsedRange.Rows.Count
    Print #fnum, this_sht.Cells(n, 1).Value
Next n
Close #fnum
this_sht.Parent.Close savechanges:=False
Set this_sht = Nothing
End Sub
Sub Report_From_PT()
Dim result_sht As Worksheet
Dim pt_sht As Worksheet
Dim n As Long
Dim i As Long
Dim proj_array() As Variant
Dim name_array() As Variant
Dim pt_sht_rows As Long
Dim cur_name As String
Dim cur_proj As String
Dim name_row As Long
Dim proj_col As Long
'set sheets variables
Set result_sht = ActiveWorkbook.Sheets(1)
Set pt_sht = ActiveWorkbook.Sheets(2)

pt_sht_rows = pt_sht.UsedRange.Rows.Count
'build array of all projects 2x2 matrix
proj_array() = result_sht.Range(result_sht.Cells(1, 1), result_sht.Cells(1, result_sht.UsedRange.Columns.Count)).Value
'build array of all names 2x2 matrix
name_array() = result_sht.Range(result_sht.Cells(1, 3), result_sht.Cells(result_sht.UsedRange.Rows.Count, 3)).Value

'loop through all PivotTable rows and record results
For n = 2 To pt_sht_rows
    'find name row in name_array only if name changes
    If cur_name <> pt_sht.Cells(n, 1) Then
        cur_name = pt_sht.Cells(n, 1) 'set name to the current name
        name_row = -1
        'loop through name_array to find the name set name_row as that row
        For i = 2 To UBound(name_array)
            If cur_name = name_array(i, 1) Then
                name_row = i
                Exit For 'name wont show twice so exit for
            End If
        Next i
    End If ' end name finding
    If name_row = -1 Then
        'MsgBox cur_name & " not found in PivotTable.  Ensure all names have a found status."
    Else
        'find column of current project from proj_array
        cur_proj = pt_sht.Cells(n, 2)
        proj_col = -1
        For i = 9 To UBound(proj_array, 2)
            If cur_proj = proj_array(1, i) Then
                proj_col = i
                Exit For 'name wont show twice so exit for
            End If
        Next i
        If proj_col = -1 Then
            MsgBox cur_proj
            Exit Sub
        End If
        result_sht.Cells(name_row, proj_col) = pt_sht.Cells(n, 3)
    End If
Next n
End Sub
Sub CAST_Data_Card_Creator()
Const template_path = "\\ETSLAN.ORG\SAO\FS_K12_DATA_03\TestDev\Work_Active\CAASPP_2015-2018\NGSS\02_Content\04_Science\11_Meetings\Data Review 2017\1. Draft Materials\for Asher\CAST_Data_card_Crosswalk.docx"

Dim wb As Workbook
Dim ws As Worksheet
Dim lr As Long
Dim wdApp As Word.Application
Dim wdDoc As Word.Document
Dim wdDoc_tmp As Word.Document
Dim tbl As Word.Table
Dim tbl_prime As Word.Table
Dim rng As Word.Range

Set wb = ActiveWorkbook
Set ws = ActiveSheet

If ws.Cells(1, 1) <> "Grade Level" Then Exit Sub

lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

Set wdApp = CreateObject("Word.Application")
Set wdDoc_tmp = wdApp.Documents.Open(filename:=template_path, ReadOnly:=True)
Set tbl_prime = wdDoc_tmp.Tables(1)

Set wdDoc = wdApp.Documents.Add
'wdApp.Visible = True

For n = 2 To lr
    
    tbl_prime.Range.Copy
    Set rng = wdDoc.content
    rng.Collapse Direction:=wdCollapseEnd
    rng.Paste
    Set rng = wdDoc.content
    rng.Collapse Direction:=wdCollapseEnd
    rng.InsertBreak wdPageBreak
    
    Set tbl = wdDoc.Tables(wdDoc.Tables.Count)
    
    tbl.Cell(2, 2).Range.text = ws.Cells(n, 1) ' Grade level
    tbl.Cell(3, 2).Range.text = ws.Cells(n, 2) ' Item Accnum
    tbl.Cell(4, 2).Range.text = ws.Cells(n, 3) ' Client Item ID
    tbl.Cell(5, 2).Range.text = ws.Cells(n, 5) ' form
    tbl.Cell(6, 2).Range.text = ws.Cells(n, 6) ' Item Sequence
    If CDbl(ws.Cells(n, 11)) < 0.2 Then   ' Item-total Correlation
        tbl.Cell(7, 2).Range.text = Round(ws.Cells(n, 11), 2) & "*"
    Else
        tbl.Cell(7, 2).Range.text = Round(ws.Cells(n, 11), 2)
    End If
    
    If CDbl(ws.Cells(n, 12)) < 0.1 Or CDbl(ws.Cells(n, 12)) > 0.95 Then    ' Item difficulty(p - Value)
        tbl.Cell(8, 2).Range.text = Round(ws.Cells(n, 12), 2) & "*"
    Else
        tbl.Cell(8, 2).Range.text = Round(ws.Cells(n, 12), 2)
    End If
    
    tbl.Cell(10, 2).Range.text = Round(ws.Cells(n, 13), 2) * 100 & "%" ' Choice A
    tbl.Cell(11, 2).Range.text = Round(ws.Cells(n, 14), 2) * 100 & "%" ' Choice B
    tbl.Cell(12, 2).Range.text = Round(ws.Cells(n, 15), 2) * 100 & "%" ' Choice c
    tbl.Cell(13, 2).Range.text = Round(ws.Cells(n, 16), 2) * 100 & "%" ' Choice D
    
    If ws.Cells(n, 29) <> "" Then tbl.Cell(10, 5).Range.text = Round(ws.Cells(n, 29), 2) * 100 & "%" ' % scoring 0 pt
    If ws.Cells(n, 30) <> "" Then tbl.Cell(11, 5).Range.text = Round(ws.Cells(n, 30), 2) * 100 & "%" ' % scoring 1 pt
    If ws.Cells(n, 31) <> "" Then tbl.Cell(12, 5).Range.text = Round(ws.Cells(n, 31), 2) * 100 & "%" ' % scoring 2 pts

    tbl.Cell(16, 2).Range.text = ws.Cells(n, 26) ' Focal
    tbl.Cell(16, 3).Range.text = ws.Cells(n, 27) ' Reference
    tbl.Cell(17, 2).Range.text = ws.Cells(n, 28) ' Group Favored

Next n
wdDoc_tmp.Close savechanges:=False
wdApp.Visible = True
wdDoc.Activate






End Sub























