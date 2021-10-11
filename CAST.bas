Attribute VB_Name = "CAST"
Sub gloss_markup_qc()

Dim wb As Workbook
Dim sht As Worksheet
Dim sum_sht As Worksheet

Dim lc As Long
Dim c As Long
Dim lr As Long
Dim r As Long
Dim sum_r As Long
Dim sht_index As Long
Dim base_term_col As Long
Dim trans_lang_col As Long
Dim audio_file_col As Long
Dim mrk_col As Long
Dim fname As String

Dim sCell As String

Set wb = ActiveWorkbook
wb.Sheets.Add after:=wb.Sheets(wb.Sheets.Count)

wb.Sheets(wb.Sheets.Count).Name = "Summary"
Set sum_sht = wb.Sheets("Summary")
sum_sht.Cells(1, 1) = "Sheet"
sum_sht.Cells(1, 2) = "Base Term"
sum_sht.Cells(1, 3) = "es-mx"
sum_sht.Cells(1, 4) = "vi"
sum_sht.Cells(1, 5) = "zh-cn"
sum_sht.Cells(1, 6) = "tl"
sum_sht.Cells(1, 7) = "ar"
sum_sht.Cells(1, 8) = "zh-yue"
sum_sht.Cells(1, 9) = "ko"
sum_sht.Cells(1, 10) = "pa"
sum_sht.Cells(1, 11) = "ru"
sum_sht.Cells(1, 12) = "hmn"
sum_sht.Cells(1, 13) = "Notes"
sum_sht.Cells(1, 14) = "Error"
sum_sht.Cells(1, 15) = "Vendor"
sum_r = 2

For sht_index = 1 To wb.Sheets.Count - 1
    Set sht = wb.Sheets(sht_index)
    
    lc = sht.Cells(1, sht.Columns.Count).End(xlToLeft).Column
    
    For c = 1 To lc
        sCell = LCase(sht.Cells(1, c))
        sCell = Replace(sCell, " ", "")
        
        Select Case sCell
            Case "baseterm"
                base_term_col = c
            Case "translatedlang"
                trans_lang_col = c
            Case "audiofile"
                audio_file_col = c
        End Select
    Next c
    
    lr = sht.Cells(sht.Rows.Count, 1).End(xlUp).Row
    
    this_term = sht.Cells(2, base_term_col)
    sum_sht.Cells(sum_r, 1) = sht.Name
    sum_sht.Cells(sum_r, 2) = this_term
    For r = 2 To lr
        ' check if the base term has changed
        If this_term <> sht.Cells(r, base_term_col) Then
            'do summary of row
            For c = 3 To 11
                If sum_sht.Cells(sum_r, c) <> "ogg/m4a" Then
                    sum_sht.Cells(sum_r, 14) = True
                    sum_sht.Cells(sum_r, c).Interior.Color = RGB(162, 0, 0)
                    sum_sht.Cells(sum_r, 13) = add_sum(sum_sht.Cells(sum_r, 13).Value, "Missing file(s) for language " & sum_sht.Cells(1, c))
                End If
            Next c
            
            ' setup next row
            this_term = sht.Cells(r, base_term_col)
            sum_r = sum_r + 1
            
            sum_sht.Cells(sum_r, 1) = sht.Name
            sum_sht.Cells(sum_r, 2) = this_term
        End If
        
        'Dim base_term_col As Long
        'Dim trans_lang_col As Long
        'Dim audio_file_col As Long
        
        Select Case sht.Cells(r, trans_lang_col)
            Case "es-mx"
                mrk_col = 3
            Case "vi"
                mrk_col = 4
            Case "zh-cn"
                mrk_col = 5
            Case "tl"
                mrk_col = 6
            Case "ar"
                mrk_col = 7
            Case "zh-yue"
                mrk_col = 8
            Case "ko"
                mrk_col = 9
            Case "pa"
                mrk_col = 10
            Case "ru"
                mrk_col = 11
            Case "hmn"
                mrk_col = 12
        End Select
            
        ' QC and correct audio file /
        If Left(sht.Cells(r, audio_file_col), 1) <> "/" Then
            sht.Cells(r, audio_file_col) = "/" & sht.Cells(r, audio_file_col)
            sum_sht.Cells(sum_r, 13) = add_sum(sum_sht.Cells(sum_r, 13).Value, "Added / to " & sht.Cells(r, audio_file_col))
            sum_sht.Cells(sum_r, 14) = True
            sum_sht.Cells(sum_r, mrk_col).Interior.Color = RGB(162, 0, 0)
        End If
        
        ' QC audio file name
        ' does audio file contain translation language
        If InStr(LCase(sht.Cells(r, audio_file_col)), LCase(sht.Cells(r, trans_lang_col))) = 0 Then
            sum_sht.Cells(sum_r, 13) = add_sum(sum_sht.Cells(sum_r, 13).Value, sht.Cells(r, audio_file_col) & " does not include translation language")
            sum_sht.Cells(sum_r, 14) = True
            sum_sht.Cells(sum_r, mrk_col).Interior.Color = RGB(162, 0, 0)
        End If
        
        ' does audio file contain base term
        fname = LCase(sht.Cells(r, base_term_col))
        fname = Replace(fname, " ", "")
        fname = Replace(fname, "-", "")
        fname = Replace(fname, "#", "")
        fname = Replace(fname, "'", "")
        
        If InStr(LCase(sht.Cells(r, audio_file_col)), fname) = 0 Then
            sum_sht.Cells(sum_r, 13) = add_sum(sum_sht.Cells(sum_r, 13).Value, sht.Cells(r, base_term_col) & " does not include base term")
            sum_sht.Cells(sum_r, 14) = True
            sum_sht.Cells(sum_r, mrk_col).Interior.Color = RGB(162, 0, 0)
        End If
            
        If sum_sht.Cells(sum_r, mrk_col) = "ogg" And LCase(Right(sht.Cells(r, audio_file_col), 3)) = "m4a" Then
            sum_sht.Cells(sum_r, mrk_col) = "ogg/m4a"
        ElseIf sum_sht.Cells(sum_r, mrk_col) = "m4a" And LCase(Right(sht.Cells(r, audio_file_col), 3)) = "ogg" Then
            sum_sht.Cells(sum_r, mrk_col) = "ogg/m4a"
        ElseIf LCase(Right(sht.Cells(r, audio_file_col), 3)) <> "ogg" And LCase(Right(sht.Cells(r, audio_file_col), 3)) <> "m4a" Then
            sum_sht.Cells(sum_r, 13) = add_sum(sum_sht.Cells(sum_r, 13).Value, sht.Cells(r, audio_file_col) & " does not have the correct file type")
            sum_sht.Cells(sum_r, 14) = True
            sum_sht.Cells(sum_r, mrk_col).Interior.Color = RGB(162, 0, 0)
        ElseIf sum_sht.Cells(sum_r, mrk_col) = "ogg/m4a" Then
            sum_sht.Cells(sum_r, 13) = add_sum(sum_sht.Cells(sum_r, 13).Value, "Extra file for " & sht.Cells(r, trans_lang_col))
            sum_sht.Cells(sum_r, 14) = True
            sum_sht.Cells(sum_r, mrk_col).Interior.Color = RGB(162, 0, 0)
        Else
            sum_sht.Cells(sum_r, mrk_col) = Right(sht.Cells(r, audio_file_col), 3)
        End If
        
        If Len(sht.Cells(r, 5)) = 6 Then
            sum_sht.Cells(sum_r, 15) = "IBIS"
            sht.Cells(r, 8) = "IBIS"
        Else
            sum_sht.Cells(sum_r, 15) = "IAT"
            sht.Cells(r, 8) = "IAT"
        End If
            
    Next r
    For c = 3 To 11
        If sum_sht.Cells(sum_r, c) <> "ogg/m4a" Then
            sum_sht.Cells(sum_r, 14) = True
            sum_sht.Cells(sum_r, c).Interior.Color = RGB(162, 0, 0)
            sum_sht.Cells(sum_r, 13) = add_sum(sum_sht.Cells(sum_r, 13).Value, "Missing file(s) for language " & sum_sht.Cells(1, c))
        End If
    Next c
    
    sum_r = sum_r + 1
Next sht_index

End Sub
Function add_sum(tSum As String, sAdd As String) As String
    If tSum = "" Then
        tSum = sAdd
    Else
        tSum = tSum & Chr(10) & sAdd
    End If

    add_sum = tSum
End Function



























