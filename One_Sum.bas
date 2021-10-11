Attribute VB_Name = "One_Sum"
Option Explicit
Sub one_report_flagging()
Dim one_bk As Workbook
Dim one_sht As Worksheet
Dim n As Long
Dim sht As Long
Dim flag As Boolean

Dim score As Long
Dim tot_cnt As Long
Dim tot_score As Long
Dim rater_ave As Double

Dim lowest As String
Dim medium As String
Dim high As String
Dim extra As String

Dim lowest_col As Long
Dim medium_col As Long
Dim high_col As Long
Dim extra_col As Long

Dim lowest_rat As Double
Dim medium_rat As Double
Dim high_rat As Double
Dim extra_rat As Double

Dim lowest_cnt As Long
Dim medium_cnt As Long
Dim high_cnt As Long
Dim extra_cnt As Long

Set one_bk = ActiveWorkbook

For sht = 1 To one_bk.Sheets.Count
    Set one_sht = one_bk.Sheets(sht)
    With one_sht
    If .Cells(3, 32) <> "04" Then
        lowest_col = 24
        medium_col = 26
        high_col = 28
        extra_col = 30
    Else
        lowest_col = 26
        medium_col = 28
        high_col = 30
        extra_col = 32
    End If
    lowest = .Cells(3, lowest_col)
    medium = .Cells(3, medium_col)
    high = .Cells(3, high_col)
    extra = .Cells(3, extra_col)
    
    For n = 5 To .Cells(.Rows.Count, 11).End(xlUp).Row
        score = CLng(.Cells(n, 12))
        If .Cells(n, 11) <> "" And lowest <> "" Then
            lowest_cnt = .Cells(n, lowest_col)
            medium_cnt = .Cells(n, medium_col)
            high_cnt = .Cells(n, high_col)
            extra_cnt = .Cells(n, extra_col)
            
            tot_cnt = lowest_cnt + medium_cnt _
                  + high_cnt + extra_cnt
            
            tot_score = lowest_cnt * CLng(lowest) + medium_cnt * CLng(medium) _
                  + high_cnt * CLng(high) + extra_cnt * CLng(extra)
            
            rater_ave = tot_score / tot_cnt
            
            lowest_rat = lowest_cnt / tot_cnt
            medium_rat = medium_cnt / tot_cnt
            high_rat = high_cnt / tot_cnt
            extra_rat = extra_cnt / tot_cnt
            
            If Abs(CDbl(score - rater_ave)) > 0.2 Then
                .Rows(n).Interior.Color = RGB(188, 0, 0)
            End If
            
            Select Case score
                Case lowest
                    If (medium_rat + high_rat + extra_rat) / 3 > 0.15 Then
                        .Rows(n).Interior.Color = RGB(188, 50, 0)
                    End If
                Case medium
                    If lowest_rat > 0.15 Or (high_rat + extra_rat) / 2 > 0.15 Then
                        .Rows(n).Interior.Color = RGB(188, 50, 0)
                    End If
                Case high
                    If (lowest_rat + medium_rat) / 2 > 0.15 Or extra_rat > 0.15 Then
                        .Rows(n).Interior.Color = RGB(188, 50, 0)
                    End If
                Case extra
                    If (lowest_rat + medium_rat + high_rat) / 3 > 0.15 Then
                        .Rows(n).Interior.Color = RGB(188, 50, 0)
                    End If
            End Select
            
        End If
    Next n
    End With
Next sht
End Sub
Sub one_report_summary()
Dim one_bk As Workbook
Dim one_sht As Worksheet
Dim sum_bk As Workbook
Dim sum_sht As Worksheet

Dim sht As Long
Dim n As Long
Dim sum_row As Long
Dim this_id As String
Dim next_id As String
Dim this_acc As String
Dim next_acc As String
Dim new_sht As Boolean

Dim lowest As String
Dim medium As String
Dim high As String
Dim extra As String

Dim lr As Long

Dim lowest_col As Long
Dim medium_col As Long
Dim high_col As Long
Dim extra_col As Long

Dim lowest_higher As Long
Dim medium_lower As Long
Dim medium_higher As Long
Dim high_lower As Long
Dim high_higher As Long
Dim extra_lower As Long

Dim lowest_cnt As Long
Dim medium_cnt As Long
Dim high_cnt As Long
Dim extra_cnt As Long

Dim lowest_num As Long
Dim medium_num As Long
Dim high_num As Long
Dim extra_num As Long

Dim lowest_rat As Double
Dim medium_rat As Double
Dim high_rat As Double
Dim extra_rat As Double

Dim lowest_tot As Long
Dim medium_tot As Long
Dim high_tot As Long
Dim extra_tot As Long

Dim lowest_val As Long
Dim medium_val As Long
Dim high_val As Long
Dim extra_val As Long

Dim lowest_higher_acc As Long
Dim medium_lower_acc As Long
Dim medium_higher_acc As Long
Dim high_lower_acc As Long
Dim high_higher_acc As Long
Dim extra_lower_acc As Long

Dim lowest_cnt_acc As Long
Dim medium_cnt_acc As Long
Dim high_cnt_acc As Long
Dim extra_cnt_acc As Long

Dim lowest_num_acc As Long
Dim medium_num_acc As Long
Dim high_num_acc As Long
Dim extra_num_acc As Long

Dim lowest_rat_acc As Double
Dim medium_rat_acc As Double
Dim high_rat_acc As Double
Dim extra_rat_acc As Double

Dim lowest_tot_acc As Long
Dim medium_tot_acc As Long
Dim high_tot_acc As Long
Dim extra_tot_acc As Long

Dim lowest_val_acc As Long
Dim medium_val_acc As Long
Dim high_val_acc As Long
Dim extra_val_acc As Long

Dim prmt_nam As String

Dim rng As Range

Set one_bk = ActiveWorkbook
Application.Workbooks.Add
Set sum_bk = ActiveWorkbook
sum_row = 1

' for each sheet in one_bk summarize by ID and Set ID
For sht = 1 To one_bk.Sheets.Count
    Set one_sht = one_bk.Sheets(sht)
    With one_sht
    
    If sht = 1 Then
        Set rng = .Range(.Cells(4, 1), .Cells(.UsedRange.Rows.Count, 52))
        rng.UnMerge
        
        .Sort.SortFields.Add key:=.Range(.Cells(4, 7), .Cells(.UsedRange.Rows.Count, 7)) _
            , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .Sort.SortFields.Add key:=.Range(.Cells(4, 11), .Cells(.UsedRange.Rows.Count, 11)) _
            , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        With .Sort
            .SetRange rng
            .header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        
        If .Cells(3, 32) <> "04" Then
            lowest_col = 24
            medium_col = 26
            high_col = 28
            extra_col = 30
        Else
            lowest_col = 26
            medium_col = 28
            high_col = 30
            extra_col = 32
        End If
        
        this_id = .Cells(5, 11)
        this_acc = .Cells(5, 7)
        prmt_nam = .Cells(5, 4)
        Set sum_sht = sum_bk.Sheets(sum_bk.Sheets.Count)
        lowest = .Cells(3, lowest_col)
        medium = .Cells(3, medium_col)
        high = .Cells(3, high_col)
        extra = .Cells(3, extra_col)
    End If
    
    lr = .Cells(.Rows.Count, 11).End(xlUp).Row
    ' loop through all rows in one_bk
    For n = 5 To lr + 1
        ' only count rows with a set_id
        If .Cells(n, 11) <> "" Then
            
            ' if the set ID changes place a summary table
            If (next_id <> this_id And next_id <> "") Or new_sht Then
            
                ' remove column B (2)
                sum_sht.Cells(sum_row, 1) = "Validity Stats by Set"
                sum_sht.Range(sum_sht.Cells(sum_row, 1), sum_sht.Cells(sum_row, 6)).Merge
                sum_sht.Cells(sum_row + 1, 1) = "Prompt Name:"
                sum_sht.Cells(sum_row + 1, 2) = prmt_nam
                sum_sht.Range(sum_sht.Cells(sum_row + 1, 2), sum_sht.Cells(sum_row + 1, 4)).Merge
                sum_sht.Cells(sum_row + 1, 5) = "Set ID:"
                sum_sht.Cells(sum_row + 1, 6) = this_id
                
                sum_sht.Cells(sum_row + 2, 1) = "Score"
                sum_sht.Cells(sum_row + 2, 2) = "Average Accuracy Rate By Samples at Score Point"
                sum_sht.Cells(sum_row + 2, 3) = "Total Ratings at Score Point"
                sum_sht.Cells(sum_row + 2, 4) = "Accuracy Rate (Based on Total Ratings at Score Point)"
                sum_sht.Cells(sum_row + 2, 5) = "% Ratings Down"
                sum_sht.Cells(sum_row + 2, 6) = "% Ratings Up"
                
                sum_sht.Cells(sum_row + 3, 1) = lowest ' good
                sum_sht.Cells(sum_row + 3, 3) = lowest_tot ' good
                If lowest_cnt <> 0 Then
                    sum_sht.Cells(sum_row + 3, 2) = lowest_rat / lowest_num ' good
                    sum_sht.Cells(sum_row + 3, 4) = lowest_cnt / lowest_tot ' good
                    sum_sht.Cells(sum_row + 3, 5) = "NA"
                    sum_sht.Cells(sum_row + 3, 6) = lowest_higher / lowest_tot
                End If
                ' % up - percentage of ratings that are at 2 or 2
                sum_sht.Cells(sum_row + 4, 1) = medium
                sum_sht.Cells(sum_row + 4, 3) = medium_tot
                If medium_cnt <> 0 Then
                    sum_sht.Cells(sum_row + 4, 2) = medium_rat / medium_num
                    sum_sht.Cells(sum_row + 4, 4) = medium_cnt / medium_tot
                    sum_sht.Cells(sum_row + 4, 5) = medium_lower / medium_tot
                    sum_sht.Cells(sum_row + 4, 6) = medium_higher / medium_tot
                End If
                
                sum_sht.Cells(sum_row + 5, 1) = high
                sum_sht.Cells(sum_row + 5, 3) = high_tot
                If high_cnt <> 0 Then
                    sum_sht.Cells(sum_row + 5, 2) = high_rat / high_num
                    sum_sht.Cells(sum_row + 5, 4) = high_cnt / high_tot
                    sum_sht.Cells(sum_row + 5, 5) = high_lower / high_tot
                    sum_sht.Cells(sum_row + 5, 6) = high_higher / high_tot
                End If
                
                sum_sht.Cells(sum_row + 6, 1) = extra
                sum_sht.Cells(sum_row + 6, 3) = extra_tot
                If extra_cnt <> 0 Then
                    sum_sht.Cells(sum_row + 6, 2) = extra_rat / extra_num
                    sum_sht.Cells(sum_row + 6, 4) = extra_cnt / extra_tot
                    sum_sht.Cells(sum_row + 6, 5) = extra_lower / extra_tot
                    sum_sht.Cells(sum_row + 6, 6) = "NA"
                End If
                
                ' totals by set
                sum_sht.Cells(sum_row + 7, 1) = "Totals by Set"
                
                Set rng = sum_sht.Range(sum_sht.Cells(sum_row + 3, 2), sum_sht.Cells(sum_row + 7, 2))
                rng.NumberFormat = "0%"
                sum_sht.Cells(sum_row + 7, 2) = Application.WorksheetFunction.Average(rng)
                
                Set rng = sum_sht.Range(sum_sht.Cells(sum_row + 3, 3), sum_sht.Cells(sum_row + 7, 3))
                sum_sht.Cells(sum_row + 7, 3) = Application.WorksheetFunction.sum(rng)
                
                Set rng = sum_sht.Range(sum_sht.Cells(sum_row + 3, 4), sum_sht.Cells(sum_row + 7, 4))
                rng.NumberFormat = "0%"
                sum_sht.Cells(sum_row + 7, 4) = Application.WorksheetFunction.Average(rng)
                
                Set rng = sum_sht.Range(sum_sht.Cells(sum_row + 3, 5), sum_sht.Cells(sum_row + 7, 5))
                rng.NumberFormat = "0%"
                sum_sht.Cells(sum_row + 7, 5) = Application.WorksheetFunction.Average(rng)
                
                Set rng = sum_sht.Range(sum_sht.Cells(sum_row + 3, 6), sum_sht.Cells(sum_row + 7, 6))
                rng.NumberFormat = "0%"
                sum_sht.Cells(sum_row + 7, 6) = Application.WorksheetFunction.Average(rng)
                
                ' set row to next table start
                sum_row = sum_row + 9
                
                
                ' add accnum stuff to set id stuff
                lowest_higher_acc = lowest_higher_acc + lowest_higher
                medium_lower_acc = medium_lower_acc + medium_lower
                medium_higher_acc = medium_higher_acc + medium_higher
                high_lower_acc = high_lower_acc + high_lower
                high_higher_acc = high_higher_acc + high_higher
                extra_lower_acc = extra_lower_acc + extra_lower
                
                lowest_cnt_acc = lowest_cnt_acc + lowest_cnt
                medium_cnt_acc = medium_cnt_acc + medium_cnt
                high_cnt_acc = high_cnt_acc + high_cnt
                extra_cnt_acc = extra_cnt_acc + extra_cnt
                
                lowest_num_acc = lowest_num_acc + lowest_num
                medium_num_acc = medium_num_acc + medium_num
                high_num_acc = high_num_acc + high_num
                extra_num_acc = extra_num_acc + extra_num
                
                lowest_rat_acc = lowest_rat_acc + lowest_rat
                medium_rat_acc = medium_rat_acc + medium_rat
                high_rat_acc = high_rat_acc + high_rat
                extra_rat_acc = extra_rat_acc + extra_rat
                
                lowest_tot_acc = lowest_tot_acc + lowest_tot
                medium_tot_acc = medium_tot_acc + medium_tot
                high_tot_acc = high_tot_acc + high_tot
                extra_tot_acc = extra_tot_acc + extra_tot
                
                lowest_val_acc = lowest_val_acc + lowest_val
                medium_val_acc = medium_val_acc + medium_val
                high_val_acc = high_val_acc + high_val
                extra_val_acc = extra_val_acc + extra_val
                
                ' reset set id stuff
                lowest_higher = 0
                medium_lower = 0
                medium_higher = 0
                high_lower = 0
                high_higher = 0
                extra_lower = 0
                
                lowest_cnt = 0
                medium_cnt = 0
                high_cnt = 0
                extra_cnt = 0
                
                lowest_num = 0
                medium_num = 0
                high_num = 0
                extra_num = 0
                
                lowest_rat = 0
                medium_rat = 0
                high_rat = 0
                extra_rat = 0
                
                lowest_tot = 0
                medium_tot = 0
                high_tot = 0
                extra_tot = 0
                
                lowest_val = 0
                medium_val = 0
                high_val = 0
                extra_val = 0
                
                this_id = next_id
            End If
            
            ' if the accnum changes place the full summary table
            If (this_acc <> next_acc And next_acc <> "") Or new_sht Then
                
                sum_sht.Cells(sum_row, 1) = "Validity Stats by Prompt (All Sets and Samples with Ratings)"
                sum_sht.Range(sum_sht.Cells(sum_row, 1), sum_sht.Cells(sum_row, 6)).Merge
                sum_sht.Cells(sum_row + 1, 1) = "Prompt Name:"
                sum_sht.Cells(sum_row + 1, 2) = prmt_nam
                sum_sht.Range(sum_sht.Cells(sum_row + 1, 2), sum_sht.Cells(sum_row + 1, 6)).Merge
                
                sum_sht.Cells(sum_row + 2, 1) = "Score"
                sum_sht.Cells(sum_row + 2, 2) = "Average Accuracy Rate By Samples at Score Point"
                sum_sht.Cells(sum_row + 2, 3) = "Total Ratings at Score Point"
                sum_sht.Cells(sum_row + 2, 4) = "Accuracy Rate (Based on Total Ratings at Score Point)"
                sum_sht.Cells(sum_row + 2, 5) = "% Ratings Down"
                sum_sht.Cells(sum_row + 2, 6) = "% Ratings Up"
                
                sum_sht.Cells(sum_row + 3, 1) = lowest ' good
                sum_sht.Cells(sum_row + 3, 3) = lowest_tot_acc ' good
                If lowest_cnt_acc <> 0 Then
                    sum_sht.Cells(sum_row + 3, 2) = lowest_rat_acc / lowest_num_acc ' good
                    sum_sht.Cells(sum_row + 3, 4) = lowest_cnt_acc / lowest_tot_acc ' good
                    sum_sht.Cells(sum_row + 3, 5) = "NA"
                    sum_sht.Cells(sum_row + 3, 6) = lowest_higher_acc / lowest_tot_acc
                End If
                
                sum_sht.Cells(sum_row + 4, 1) = medium
                sum_sht.Cells(sum_row + 4, 3) = medium_tot_acc
                If medium_cnt_acc <> 0 Then
                    sum_sht.Cells(sum_row + 4, 2) = medium_rat_acc / medium_num_acc
                    sum_sht.Cells(sum_row + 4, 4) = medium_cnt_acc / medium_tot_acc
                    sum_sht.Cells(sum_row + 4, 5) = medium_lower_acc / medium_tot_acc
                    sum_sht.Cells(sum_row + 4, 6) = medium_higher_acc / medium_tot_acc
                End If
                
                sum_sht.Cells(sum_row + 5, 1) = high
                sum_sht.Cells(sum_row + 5, 3) = high_tot_acc
                If high_cnt_acc <> 0 Then
                    sum_sht.Cells(sum_row + 5, 2) = high_rat_acc / high_num_acc
                    sum_sht.Cells(sum_row + 5, 4) = high_cnt_acc / high_tot_acc
                    sum_sht.Cells(sum_row + 5, 5) = high_lower_acc / high_tot_acc
                    sum_sht.Cells(sum_row + 5, 6) = high_higher_acc / high_tot_acc
                End If
                
                sum_sht.Cells(sum_row + 6, 1) = extra
                sum_sht.Cells(sum_row + 6, 3) = extra_tot_acc
                If extra_cnt_acc <> 0 Then
                    sum_sht.Cells(sum_row + 6, 2) = extra_rat_acc / extra_num_acc
                    sum_sht.Cells(sum_row + 6, 4) = extra_cnt_acc / extra_tot_acc
                    sum_sht.Cells(sum_row + 6, 5) = extra_lower_acc / extra_tot_acc
                    sum_sht.Cells(sum_row + 6, 6) = "NA"
                End If
                
                ' totals by set
                sum_sht.Cells(sum_row + 7, 1) = "Totals by Prompt"
                
                Set rng = sum_sht.Range(sum_sht.Cells(sum_row + 3, 2), sum_sht.Cells(sum_row + 7, 2))
                rng.NumberFormat = "0%"
                sum_sht.Cells(sum_row + 7, 2) = Application.WorksheetFunction.Average(rng)
                
                Set rng = sum_sht.Range(sum_sht.Cells(sum_row + 3, 3), sum_sht.Cells(sum_row + 7, 3))
                sum_sht.Cells(sum_row + 7, 3) = Application.WorksheetFunction.sum(rng)
                
                Set rng = sum_sht.Range(sum_sht.Cells(sum_row + 3, 4), sum_sht.Cells(sum_row + 7, 4))
                rng.NumberFormat = "0%"
                sum_sht.Cells(sum_row + 7, 4) = Application.WorksheetFunction.Average(rng)
                
                Set rng = sum_sht.Range(sum_sht.Cells(sum_row + 3, 5), sum_sht.Cells(sum_row + 7, 5))
                rng.NumberFormat = "0%"
                sum_sht.Cells(sum_row + 7, 5) = Application.WorksheetFunction.Average(rng)
                
                Set rng = sum_sht.Range(sum_sht.Cells(sum_row + 3, 6), sum_sht.Cells(sum_row + 7, 6))
                rng.NumberFormat = "0%"
                sum_sht.Cells(sum_row + 7, 6) = Application.WorksheetFunction.Average(rng)
                
                ' set row to next table start
                sum_row = sum_row + 9
                
                
                
                ' reset item stuff
                lowest_higher_acc = 0
                medium_lower_acc = 0
                medium_higher_acc = 0
                high_lower_acc = 0
                high_higher_acc = 0
                extra_lower_acc = 0
                
                lowest_cnt_acc = 0
                medium_cnt_acc = 0
                high_cnt_acc = 0
                extra_cnt_acc = 0
                
                lowest_num_acc = 0
                medium_num_acc = 0
                high_num_acc = 0
                extra_num_acc = 0
                
                lowest_rat_acc = 0
                medium_rat_acc = 0
                high_rat_acc = 0
                extra_rat_acc = 0
                
                lowest_tot_acc = 0
                medium_tot_acc = 0
                high_tot_acc = 0
                extra_tot_acc = 0
                
                this_id = next_id
                this_acc = next_acc
                prmt_nam = .Cells(n, 4)
                
                If sht <> 1 And new_sht Then
                    
                    Set rng = .Range(.Cells(4, 1), .Cells(.UsedRange.Rows.Count, 52))
                    rng.UnMerge
                    
                    .Sort.SortFields.Add key:=.Range(.Cells(4, 7), .Cells(.UsedRange.Rows.Count, 7)) _
                        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
                    .Sort.SortFields.Add key:=.Range(.Cells(4, 11), .Cells(.UsedRange.Rows.Count, 11)) _
                        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
                    With .Sort
                        .SetRange rng
                        .header = xlYes
                        .MatchCase = False
                        .Orientation = xlTopToBottom
                        .SortMethod = xlPinYin
                        .Apply
                    End With
                    
                    If .Cells(3, 32) <> "04" Then
                        lowest_col = 24
                        medium_col = 26
                        high_col = 28
                        extra_col = 30
                    Else
                        lowest_col = 26
                        medium_col = 28
                        high_col = 30
                        extra_col = 32
                    End If
                    
                    this_id = .Cells(5, 11)
                    this_acc = .Cells(5, 7)
                    prmt_nam = .Cells(5, 4)
                    sum_bk.Sheets.Add after:=sum_bk.Sheets(sum_bk.Sheets.Count)
                    Set sum_sht = sum_bk.Sheets(sum_bk.Sheets.Count)
                    sum_row = 1
                    lowest = .Cells(3, lowest_col)
                    medium = .Cells(3, medium_col)
                    high = .Cells(3, high_col)
                    extra = .Cells(3, extra_col)
                End If
            End If
            
            new_sht = False
            ' set the values for each
            lowest_val = .Cells(n, lowest_col)
            medium_val = .Cells(n, medium_col)
            high_val = .Cells(n, high_col)
            extra_val = .Cells(n, extra_col)
            
            
            ' store the row's information
            Select Case .Cells(n, 12)
                Case lowest
                    lowest_higher = lowest_higher + medium_val + high_val + extra_val
                    lowest_cnt = lowest_cnt + lowest_val
                    lowest_rat = lowest_rat + .Cells(n, lowest_col + 1)
                    lowest_tot = lowest_tot + lowest_val + medium_val + high_val + extra_val
                    lowest_num = lowest_num + 1
                Case medium
                    medium_lower = medium_lower + lowest_val
                    medium_higher = medium_higher + high_val + extra_val
                    medium_cnt = medium_cnt + medium_val
                    medium_rat = medium_rat + .Cells(n, medium_col + 1)
                    medium_tot = medium_tot + lowest_val + medium_val + high_val + extra_val
                    medium_num = medium_num + 1
                Case high
                    high_lower = high_lower + medium_val + lowest_val
                    high_higher = high_higher + extra_val
                    high_cnt = high_cnt + high_val
                    high_rat = high_rat + .Cells(n, high_col + 1)
                    high_tot = high_tot + lowest_val + medium_val + high_val + extra_val
                    high_num = high_num + 1
                Case extra
                    extra_lower = extra_lower + high_val + medium_val + lowest_val
                    extra_cnt = extra_cnt + extra_val
                    extra_rat = extra_rat + .Cells(n, extra_col + 1)
                    extra_tot = extra_tot + lowest_val + medium_val + high_val + extra_val
                    extra_num = extra_num + 1
            End Select
            
            next_id = .Cells(n, 11)
            next_acc = .Cells(n, 7)
        End If
        
    Next n
    End With
    new_sht = True
Next sht

End Sub
