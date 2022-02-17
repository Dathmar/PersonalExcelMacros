Attribute VB_Name = "PARCC_Summary"
Sub merge_PARCC()

Set filenames = ActiveSheet

For n = 2 To 37
    xlfilename = filenames.Cells(n, 1)
    
    Application.Workbooks.Open filename:=xlfilename, ReadOnly:=False
    Set wb = ActiveWorkbook
    Set all_items = wb.Sheets(1)
    Set flagged_items = wb.Sheets(2)
    
    ' find comment columns
    For c = 1 To flagged_items.UsedRange.Columns.Count
        col = LCase(Trim(flagged_items.Cells(1, c)))
        If col = "stat comment" Or col = "stat comments" Then
            stat_com = c
        ElseIf col = LCase("AD Review Category") Then
            ad_review = c
        ElseIf col = LCase("AD Comments") Then
            ad_com = c
        ElseIf col = LCase("ItemNumber") Then
            item_num = c
        ElseIf col = LCase("Form") Then
            form = c
        End If
    Next c
    
    If stat_com = 0 Or ad_review = 0 Or ad_com = 0 Or item_num = 0 Then
        MsgBox "error"
        Exit Sub
    End If
    
    lc = all_items.UsedRange.Columns.Count
    all_items.Cells(1, lc + 1) = "item"
    all_items.Cells(1, lc + 2) = "Stat Comments"
    all_items.Cells(1, lc + 3) = "AD Review Category"
    all_items.Cells(1, lc + 4) = "AD Comments"
    
    For r = 2 To flagged_items.UsedRange.Rows.Count
        If flagged_items.Cells(r, 1) <> "" Then
            
            With all_items
            ' merge in extra data for each item number
            ' filter by grade, session, item and linking forms
                .UsedRange.AutoFilter Field:=3, Criteria1:=flagged_items.Cells(r, item_num)
                .UsedRange.AutoFilter Field:=7, Criteria1:=flagged_items.Cells(r, form)
                
                ' find the last visible row
                lr = .Cells(.Rows.Count, 1).End(xlUp).Row
                If lr <> 1 Then
                    .Cells(lr, lc + 1) = flagged_items.Cells(r, item_num)
                    .Cells(lr, lc + 2) = flagged_items.Cells(r, stat_com)
                    .Cells(lr, lc + 3) = flagged_items.Cells(r, ad_review)
                    .Cells(lr, lc + 4) = flagged_items.Cells(r, ad_com)
                    
                Else
                    .Cells(lr, lc + 1) = "error"
                End If
            End With

        End If
    Next r
    
    filenames.Cells(n, 2) = "TRUE"
    wb.Save
    wb.Close
Next n


End Sub
Sub blah_PARCC()

Set filenames = ActiveSheet

For n = 2 To 37
    xlfilename = filenames.Cells(n, 1)
    
    Application.Workbooks.Open filename:=xlfilename, ReadOnly:=False
    Set wb = ActiveWorkbook
    Set all_items = wb.Sheets(1)
    
    ' find comment columns
    For c = 1 To all_items.UsedRange.Columns.Count
        col = LCase(Trim(all_items.Cells(1, c)))
        If col = LCase("item") Then
            all_items.Cells(1, c).EntireColumn.Delete
        End If
    Next c

    all_items.ShowAllData
    
    For s = wb.Sheets.Count To 2 Step -1
        wb.Sheets(s).Delete
    Next s
    
    
    filenames.Cells(n, 3) = "TRUE"
    wb.Save
    wb.Close
Next n
End Sub
Sub headers_PARCC()
Dim hb As Workbook
Dim all_items As Worksheet
Dim hed_txt As String
Set filenames = ActiveSheet
For n = 3 To 49
    xlfilename = filenames.Cells(n, 1)
    
    Application.Workbooks.Open filename:=xlfilename, ReadOnly:=True
    Set fb = ActiveWorkbook
    Set all_items = fb.ActiveSheet
    
    Application.Workbooks.Open filename:="C:\Users\adanner\Desktop\Working Folder\PARCC Stat macro\Math Offloads\By header\header file.xlsx", ReadOnly:=True
    Set hb = ActiveWorkbook
    Set hed = hb.ActiveSheet
    
    ' loop through all headers in hed file and copy the correct column from the offload file
    For hed_col = 1 To hed.UsedRange.Columns.Count
        hed_txt = hed.Cells(1, hed_col)
        
        from_col = get_from_col(all_items, hed_txt)
        lr = all_items.UsedRange.Rows.Count
        
        If from_col <> 0 Then
            With all_items
                .Range(.Cells(2, from_col), .Cells(lr, from_col)).Copy Destination:=hed.Cells(2, hed_col)
            End With
        End If
        
        
    Next hed_col
    
    save_name = Left(fb.Name, InStr(fb.Name, "d_") + 1) & ".xlsx"
    
    hb.SaveAs filename:="C:\Users\adanner\Desktop\Working Folder\PARCC Stat macro\Math Offloads\By header\" & save_name
    hb.Close
    filenames.Cells(n, 2) = "TRUE"
    fb.Close
Next n


End Sub
Function get_from_col(ws As Worksheet, hed_txt As String) As Long
get_from_col = 0
For n = 1 To ws.UsedRange.Columns.Count
    If ws.Cells(1, n) = hed_txt Then
        get_from_col = n
        Exit Function
    End If
Next n

End Function
Sub build_aggrigated_PARCC()
If ActiveSheet.Cells(2, 2) = "MATH" Then
    Call build_aggrigated_PARCC_Math
Else
    Call build_aggrigated_PARCC_ELA
End If
End Sub
Function build_aggrigated_PARCC_ELA()
Dim ab As Workbook
Dim form_list As Worksheet
Dim sum_sht As Worksheet
Dim act_sht As Worksheet
Dim rng As Range

Application.ScreenUpdating = False

Set ab = ActiveWorkbook
If ab.Sheets.Count < 5 Then ab.Worksheets.Add after:=ab.Sheets(ab.Sheets.Count)

Set form_list = ab.Sheets(2)
Set sum_sht = ab.Sheets(3)
Set sum_pbt = ab.Sheets(4)
Set act_sht = ab.Sheets(5)

sum_sht.Name = "Summary"
form_list.Name = "Forms View"
sum_pbt.Name = "Summary_PBT"

For r = 2 To sum_sht.UsedRange.Rows.Count
    flags = ""
    act_sht.Cells.Delete
    
    'filter by item and grade to get aggregated data
    form_list.UsedRange.AutoFilter Field:=11, Criteria1:=sum_sht.Cells(r, 11) ' item
    form_list.UsedRange.SpecialCells(xlCellTypeVisible).Copy Destination:=act_sht.Cells(1, 1)

    ' sum_sht includes all values
    sum_sht.Cells(r, 16).Formula = "=SUMIF('Forms View'!K:K,Summary!K" & r & ",'Forms View'!AJ:AJ)" ' N_reached
    sum_sht.Cells(r, 17).Formula = "=SUMIF('Forms View'!K:K,Summary!K" & r & ",'Forms View'!AH:AH)" ' N_omit
    sum_sht.Cells(r, 18).Formula = "=SUMIF('Forms View'!K:K,Summary!K" & r & ",'Forms View'!AO:AO)" ' N_NotReached
    sum_sht.Cells(r, 19).Formula = "=SUMIF('Forms View'!K:K,Summary!K" & r & ",'Forms View'!AW:AW)" ' N_Total
    sum_sht.Cells(r, 20).Formula = "=SUMIF('Forms View'!K:K,Summary!K" & r & ",'Forms View'!AY:AY)/S" & r ' P_plus
    sum_sht.Cells(r, 21).Formula = "=SUMIF('Forms View'!K:K,Summary!K" & r & ",'Forms View'!BO:BO)/S" & r ' polyserial
    
    ' sum_pbt contains only data for PBT forms
    sum_pbt.Cells(r, 16).Formula = "=SUMIFS('Forms View'!AJ:AJ,'Forms View'!K:K,Summary_PBT!K" & r & ",'Forms View'!M:M," & Chr(34) & Chr(42) & "PBT" & Chr(42) & Chr(34) & ")" 'N_reached
    sum_pbt.Cells(r, 17).Formula = "=SUMIFS('Forms View'!AH:AH,'Forms View'!K:K,Summary_PBT!K" & r & ",'Forms View'!M:M," & Chr(34) & Chr(42) & "PBT" & Chr(42) & Chr(34) & ")" ' N_omit
    sum_pbt.Cells(r, 18).Formula = "=SUMIFS('Forms View'!AO:AO,'Forms View'!K:K,Summary_PBT!K" & r & ",'Forms View'!M:M," & Chr(34) & Chr(42) & "PBT" & Chr(42) & Chr(34) & ")" ' N_NotReached
    sum_pbt.Cells(r, 19).Formula = "=SUMIFS('Forms View'!AW:AW,'Forms View'!K:K,Summary_PBT!K" & r & ",'Forms View'!M:M," & Chr(34) & Chr(42) & "PBT" & Chr(42) & Chr(34) & ")" ' N_Total
    sum_pbt.Cells(r, 20).Formula = "=SUMIFS('Forms View'!AY:AY,'Forms View'!K:K,Summary_PBT!K" & r & ",'Forms View'!M:M," & Chr(34) & Chr(42) & "PBT" & Chr(42) & Chr(34) & ")/S" & r ' P_plus
    sum_pbt.Cells(r, 21).Formula = "=SUMIFS('Forms View'!BO:BO,'Forms View'!K:K,Summary_PBT!K" & r & ",'Forms View'!M:M," & Chr(34) & Chr(42) & "PBT" & Chr(42) & Chr(34) & ")/S" & r ' polyserial

    With act_sht
    For n = 2 To act_sht.UsedRange.Rows.Count
        ' find flags for current item
        sum_sht.Cells(r, 12) = get_unique_col(act_sht, 15) ' flags
        sum_sht.Cells(r, 13) = get_unique_col(act_sht, 18) ' scorecat
        sum_sht.Cells(r, 14) = get_col(act_sht, 21) ' Form
        sum_sht.Cells(r, 15) = get_col(act_sht, 23) ' seqno
        
        ' find PBT  forms and place on sum_pbt
        sum_pbt.Cells(r, 12) = get_unique_col(act_sht, 15, "CBT", 13) ' flags
        sum_pbt.Cells(r, 13) = get_unique_col(act_sht, 18, "CBT", 13) ' scorecat
        sum_pbt.Cells(r, 14) = get_col(act_sht, 21, "CBT", 13) ' Form
        sum_pbt.Cells(r, 15) = get_col(act_sht, 23, "CBT", 13) ' seqno
        
        dif_arr = Array(177, 199, 200, 222, 223, 245, 246, 268, 269, 291, 292, 314, 315, 337, 338, 360, 361, 383, 384, 406, 407, 429, 430, 452, 453, 475, 476, 498, 499, 521, 522, 544, 545, 567, 568, 590, 591, 613, 614, 636, 637, 659, 660, 682, 683, 705, 706, 728, 729, 751, 752, 774, 775, 797)
        
        For elmt = LBound(dif_arr) To UBound(dif_arr)
            sum_sht.Cells(r, 22 + elmt) = get_unique_col(act_sht, CLng(dif_arr(elmt))) ' seqno
        Next elmt
    Next n
    End With
Next r
form_list.ShowAllData
Application.DisplayAlerts = False
act_sht.Delete
Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Function
Function build_aggrigated_PARCC_Math()
Dim ab As Workbook
Dim form_list As Worksheet
Dim sum_sht As Worksheet
Dim act_sht As Worksheet
Dim rng As Range

Application.ScreenUpdating = False

Set ab = ActiveWorkbook
If ab.Sheets.Count < 5 Then ab.Worksheets.Add after:=ab.Sheets(ab.Sheets.Count)

Set form_list = ab.Sheets(2)
Set sum_sht = ab.Sheets(3)
Set sum_pbt = ab.Sheets(4)
Set act_sht = ab.Sheets(5)

sum_sht.Name = "Summary"
form_list.Name = "Forms View"
sum_pbt.Name = "Summary_PBT"

For r = 2 To sum_sht.UsedRange.Rows.Count
    flags = ""
    act_sht.Cells.Delete
    
    'filter by item and grade to get aggregated data
    form_list.UsedRange.AutoFilter Field:=11, Criteria1:=sum_sht.Cells(r, 11) ' item
    form_list.UsedRange.SpecialCells(xlCellTypeVisible).Copy Destination:=act_sht.Cells(1, 1)
    
    ' sum_sht includes all values
    sum_sht.Cells(r, 16).Formula = "=SUMIF('Forms View'!K:K,Summary!K" & r & ",'Forms View'!AJ:AJ)" ' N_reached
    sum_sht.Cells(r, 17).Formula = "=SUMIF('Forms View'!K:K,Summary!K" & r & ",'Forms View'!AH:AH)" ' N_omit
    sum_sht.Cells(r, 18).Formula = "=SUMIF('Forms View'!K:K,Summary!K" & r & ",'Forms View'!AO:AO)" ' N_NotReached
    sum_sht.Cells(r, 19).Formula = "=SUMIF('Forms View'!K:K,Summary!K" & r & ",'Forms View'!BA:BA)" ' N_Total
    sum_sht.Cells(r, 20).Formula = "=SUMIF('Forms View'!K:K,Summary!K" & r & ",'Forms View'!BC:BC)/S" & r ' P_plus
    sum_sht.Cells(r, 21).Formula = "=SUMIF('Forms View'!K:K,Summary!K" & r & ",'Forms View'!BS:BS)/S" & r ' polyserial
    
    ' sum_pbt contains only data for PBT forms
    sum_pbt.Cells(r, 16).Formula = "=SUMIFS('Forms View'!AJ:AJ,'Forms View'!K:K,Summary_PBT!K" & r & ",'Forms View'!M:M," & Chr(34) & Chr(42) & "PBT" & Chr(42) & Chr(34) & ")" ' N_reached
    sum_pbt.Cells(r, 17).Formula = "=SUMIFS('Forms View'!AH:AH,'Forms View'!K:K,Summary_PBT!K" & r & ",'Forms View'!M:M," & Chr(34) & Chr(42) & "PBT" & Chr(42) & Chr(34) & ")" ' N_omit
    sum_pbt.Cells(r, 18).Formula = "=SUMIFS('Forms View'!AO:AO,'Forms View'!K:K,Summary_PBT!K" & r & ",'Forms View'!M:M," & Chr(34) & Chr(42) & "PBT" & Chr(42) & Chr(34) & ")" ' N_NotReached
    sum_pbt.Cells(r, 19).Formula = "=SUMIFS('Forms View'!BA:BA,'Forms View'!K:K,Summary_PBT!K" & r & ",'Forms View'!M:M," & Chr(34) & Chr(42) & "PBT" & Chr(42) & Chr(34) & ")" ' N_Total
    sum_pbt.Cells(r, 20).Formula = "=SUMIFS('Forms View'!BC:BC,'Forms View'!K:K,Summary_PBT!K" & r & ",'Forms View'!M:M," & Chr(34) & Chr(42) & "PBT" & Chr(42) & Chr(34) & ")/S" & r ' P_plus
    sum_pbt.Cells(r, 21).Formula = "=SUMIFS('Forms View'!BS:BS,'Forms View'!K:K,Summary_PBT!K" & r & ",'Forms View'!M:M," & Chr(34) & Chr(42) & "PBT" & Chr(42) & Chr(34) & ")/S" & r ' polyserial
    
    With act_sht
    For n = 2 To act_sht.UsedRange.Rows.Count
        ' find flags for current item
        sum_sht.Cells(r, 12) = get_unique_col(act_sht, 15) ' flags
        sum_sht.Cells(r, 13) = get_unique_col(act_sht, 18) ' scorecat
        sum_sht.Cells(r, 14) = get_col(act_sht, 21) ' Form
        sum_sht.Cells(r, 15) = get_col(act_sht, 23) ' seqno
        
        dif_arr = Array(165, 187, 188, 210, 211, 233, 234, 256, 257, 279, 280, 302, 303, 325, 326, 348, 349, 371, 372, 394, 395, 417, 418, 440, 441, 463, 464, 486, 487, 509, 510, 532, 533, 555, 556, 578, 579, 601, 602, 624, 625, 647, 648, 670, 671, 693, 694, 716, 717, 739, 740, 762, 763, 785)
        
        For elmt = LBound(dif_arr) To UBound(dif_arr)
            sum_sht.Cells(r, 22 + elmt) = get_unique_col(act_sht, CLng(dif_arr(elmt)), "CBT", 13) ' seqno
        Next elmt
        

        ' sum_pbt includes only PBT forms
        sum_pbt.Cells(r, 12) = get_unique_col(act_sht, 15, "CBT", 13) ' flags
        sum_pbt.Cells(r, 13) = get_unique_col(act_sht, 18, "CBT", 13) ' scorecat
        sum_pbt.Cells(r, 14) = get_col(act_sht, 21, "CBT", 13) ' Form
        sum_pbt.Cells(r, 15) = get_col(act_sht, 23, "CBT", 13) ' seqno
        
    Next n
    End With
Next r
form_list.ShowAllData
Application.DisplayAlerts = False
act_sht.Delete
Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Function
Function get_col(sht As Worksheet, col As Long, Optional exclude As String = "", Optional col_ex As Long = 1) As String
Dim vals As String
vals = ""
For n = 2 To sht.UsedRange.Rows.Count
    If sht.Cells(n, col) <> "" And (exclude = "" Or InStr(sht.Cells(n, col_ex), exclude) = 0) Then
        vals = vals & " " & sht.Cells(n, col)
    End If
Next n
get_col = vals
End Function
Function get_unique_col(sht As Worksheet, col As Long, Optional exclude As String = "", Optional col_ex As Long = 1) As String
Dim vals As String
Dim cur_val As Variant
vals = ""
For n = 2 To sht.UsedRange.Rows.Count
    If sht.Cells(n, col) <> "" And (exclude = "" Or InStr(sht.Cells(n, col_ex), exclude) = 0) Then
        If vals <> "" Then
            cur_val = Split(sht.Cells(n, col), " ")
            For f = LBound(cur_val) To UBound(cur_val)
                If InStr(vals, Trim(cur_val(f))) = 0 Then
                    vals = vals & " " & Trim(cur_val(f))
                End If
            Next f
        Else
            vals = sht.Cells(n, col)
        End If
    End If
Next n
get_unique_col = vals
End Function



























