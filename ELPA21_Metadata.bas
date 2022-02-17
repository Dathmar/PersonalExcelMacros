Attribute VB_Name = "ELPA21_Metadata"
Sub ELPA_QC()
Call format_sheets
Call qc_macro
End Sub
Function format_sheets()
Application.DisplayAlerts = False
With ActiveWorkbook
If .Sheets.Count = 2 Then
    .Sheets(2).Delete
End If

.Sheets.Add after:=.Sheets(1)
.Sheets(1).Cells.Copy Destination:=.Sheets(2).Cells

endrow = ActiveSheet.UsedRange.Rows.Count
For i = 1 To 1

this_name = .Sheets(2).Cells(2, i)

For n = 2 To endrow
    If .Sheets(2).Cells(n, i) = "" Then
        .Sheets(2).Cells(n, i) = this_name
    Else
        this_name = .Sheets(2).Cells(n, i)
    End If
Next n
Next i
For n = endrow To 2 Step -1
    If .Sheets(2).Cells(n, 6) = "N" Or .Sheets(2).Cells(n, 6) = "" Then
        .Sheets(2).Rows(n).EntireRow.Delete
    End If
Next n
End With
Application.DisplayAlerts = True
End Function
Sub qc_macro()
Dim qc_sht As Worksheet
Dim col_count As Long
Dim row_count As Long
Dim n As Long
Dim err_cnt As Long
Dim pld_list() As String
Dim sl_tc As String
Dim R3C1 As String
Dim R4C2 As String
Dim sl As String
Dim R4C1 As String
sl = "first"

Set qc_sht = ActiveWorkbook.Sheets(2)
col_count = qc_sht.UsedRange.Columns.Count
row_count = qc_sht.UsedRange.Rows.Count

qc_sht.Cells(1, col_count + 1) = "QC Notes"
qc_sht.Cells(1, col_count + 2) = "R1C4 to Measure"
qc_sht.Cells(1, col_count + 3) = "PLDs to R1C3"
qc_sht.Cells(1, col_count + 4) = "PLDs to R1C4"
qc_sht.Cells(1, col_count + 5) = "R1C3 to R1C4"
qc_sht.Cells(1, col_count + 6) = "Text Complexity"
qc_sht.Cells(1, col_count + 7) = "R3C1 Yes for R4C2 audio/video"
qc_sht.Cells(1, col_count + 8) = "R3C2 depends on item type"
qc_sht.Cells(1, col_count + 9) = "R1C1 matches test"
qc_sht.Cells(1, col_count + 10) = "R2C4 types have a value in R4C4"
qc_sht.Cells(1, col_count + 11) = "R4C2 should have value if R3C1 is Yes"
qc_sht.Cells(1, col_count + 12) = "Experimental to Don't Include With"
qc_sht.Cells(1, col_count + 13) = "R1C2 matches measure"
qc_sht.Cells(1, col_count + 14) = "R4C4 checked against task types in R2C4"
qc_sht.Cells(1, col_count + 15) = "R3C5 should only be blank if R3C4 is Inaccessible"
qc_sht.Cells(1, col_count + 16) = "R4C1 should be blank if Do not include has a value"
qc_sht.Cells(1, col_count + 17) = "R4C3 should be Yes depending on R2C4 value"
qc_sht.Cells(1, col_count + 18) = "Grade - R1C1 not blank"
qc_sht.Cells(1, col_count + 19) = "Task Type - R2C4 not blank"
qc_sht.Cells(1, col_count + 20) = "Domain - R1C2 not blank"
qc_sht.Cells(1, col_count + 21) = "PLD - UserCode 9 not blank"
qc_sht.Cells(1, col_count + 22) = "Standard - R1C3 not blank"
qc_sht.Cells(1, col_count + 23) = "Sub-claim - R1C4 not blank"
qc_sht.Cells(1, col_count + 24) = "Associated Practice - R1C7 not blank"
qc_sht.Cells(1, col_count + 25) = "R4C5 not blank for members"

qc_sht.Cells(1, col_count + 1).EntireColumn.ColumnWidth = 50
sl_tc = ""
' each check will happen row by row
For n = 2 To row_count
    com_plds = ""
    R1C3 = ""
    R1C4 = ""
    err_cnt = 0
    err_list = ""
    If qc_sht.Cells(n, 14) <> "" Then
        measure = qc_sht.Cells(n, 14)
    Else
        measure = qc_sht.Cells(n, 12)
    End If
    
    ' Check 1 PLD’s (User Code 9), standards (R1C3) and sub-claims (R1C4) are consistent
    If qc_sht.Cells(n, 32) <> "" Then
        pld_list = Split(qc_sht.Cells(n, 32), ",")
        Call QSortInPlace(pld_list)
        For i = LBound(pld_list) To UBound(pld_list)
            If pld_list(i) = "" Then
                err_cnt = err_cnt + 1
                err_list = err_list & Chr(10) & err_cnt & ". PLD has trailing comma"
                qc_sht.Cells(n, col_count + 3) = "FALSE"
            ElseIf InStr(pld_list(i), ".") = 0 Then
                err_cnt = err_cnt + 1
                err_list = err_list & Chr(10) & err_cnt & ". PLD has an erronious value"
                qc_sht.Cells(n, col_count + 3) = "FALSE"
            Else
                pld_list(i) = Left(pld_list(i), InStr(pld_list(i), ".") - 1)
            End If
        Next i
        plds = unique(pld_list)
        
        For i = LBound(plds) To UBound(plds)
            com_plds = Trim(com_plds & " " & plds(i))
        Next i
    Else
        com_plds = ""
    End If
    If com_plds = "" Then com_plds = "[BLANK]"
    r1c3_list = Split(qc_sht.Cells(n, 15), "|")
    For i = LBound(r1c3_list) To UBound(r1c3_list)
        R1C3 = Trim(R1C3 & " " & Left(r1c3_list(i), InStr(r1c3_list(i), ".") - 1))
    Next i
    If R1C3 = "" Then R1C3 = "[BLANK]"
    qc_sht.Cells(n, col_count + 2) = "TRUE"
    r1c4_list = Split(qc_sht.Cells(n, 16), "|")
    For i = LBound(r1c4_list) To UBound(r1c4_list)
        If InStr(r1c4_list(i), Left(measure, 1)) = 0 Then
            err_cnt = err_cnt + 1
            err_list = err_list & Chr(10) & err_cnt & ". Incorrect R1C4 value """ & r1c4_list(i) & """ does not match measure " & measure
            qc_sht.Cells(n, col_count + 2) = "FALSE"
        Else
            R1C4 = Trim(R1C4 & " " & Left(r1c4_list(i), InStr(r1c4_list(i), Left(measure, 1)) - 1))
        End If
    Next i
    If R1C4 = "" Then R1C4 = "[BLANK]"
    
    'check PLDs to standard
    If com_plds <> R1C3 Then
        err_cnt = err_cnt + 1
        err_list = err_list & Chr(10) & err_cnt & ". PLD list of " & com_plds & " does not match R1C3 values of " & R1C3
        qc_sht.Cells(n, col_count + 3) = "FALSE"
    ElseIf qc_sht.Cells(n, col_count + 3) <> "FALSE" Then
        qc_sht.Cells(n, col_count + 3) = "TRUE"
    End If
    
    ' check PLDs to R1C4
    If com_plds <> R1C4 Then
        err_cnt = err_cnt + 1
        err_list = err_list & Chr(10) & err_cnt & ". PLD list of " & com_plds & " does not match R1C4 values of " & R1C4
        qc_sht.Cells(n, col_count + 4) = "FALSE"
    Else
        qc_sht.Cells(n, col_count + 4) = "TRUE"
    End If
    
    ' Check R1C4 to R1C3
    If R1C4 <> R1C3 Then
        err_cnt = err_cnt + 1
        err_list = err_list & Chr(10) & err_cnt & ". R1C3 list of " & R1C3 & " does not match R1C4 values of " & R1C4
        qc_sht.Cells(n, col_count + 5) = "FALSE"
    Else
        qc_sht.Cells(n, col_count + 5) = "TRUE"
    End If
    
    ' text complexity
    If sl <> qc_sht.Cells(n, 2) Then
        sl = qc_sht.Cells(n, 2)
        sl_tc = qc_sht.Cells(n, 31)
        If sl_tc = "" Then sl_tc = "[BLANK]"
    End If
    tc = qc_sht.Cells(n, 31)
    If tc = "" Then
        tc = "[BLANK]"
    Else
        tc = CStr(tc)
    End If
    
    If sl = qc_sht.Cells(n, 2) And sl_tc <> tc Then
        err_cnt = err_cnt + 1
        err_list = err_list & Chr(10) & err_cnt & ". Expected text complexity of " & sl_tc & " saw " & tc
        qc_sht.Cells(n, col_count + 6) = "FALSE"
    ElseIf sl = "" Then
        qc_sht.Cells(n, col_count + 6) = "NA"
    Else
        qc_sht.Cells(n, col_count + 6) = "TRUE"
    End If
    
    ' Tech Enhanced and Tech Enabled  match XML (R3C1 and R3C2)
    ' R3C1 = Yes if the item or associated SL contains audio or video; else No
    R3C1 = qc_sht.Cells(n, 20).Value
    R4C2 = qc_sht.Cells(n, 27)
    If R3C1 = "Yes" And (InStr(R4C2, "Audio") = 0 And InStr(R4C2, "Video") = 0) Then
        err_cnt = err_cnt + 1
        err_list = err_list & Chr(10) & err_cnt & ". Expected R3C1 to be No because R4C2 does not contain Audio or Video."
        qc_sht.Cells(n, col_count + 7) = "FALSE"
    ElseIf R3C1 = "No" And (InStr(R4C2, "Audio") <> 0 And InStr(R4C2, "Video") <> 0) Then
        err_cnt = err_cnt + 1
        err_list = err_list & Chr(10) & err_cnt & ". Expected R3C1 to be Yes because R4C2 is " & R4C2
        qc_sht.Cells(n, col_count + 7) = "FALSE"
    Else
        qc_sht.Cells(n, col_count + 7) = "TRUE"
    End If
    
    ' R3C2 = Yes if item type = Match SS, Match MS, Zones SS, Zones MS, Audio CR,
    ' InlineChoiceList MS, InlineChoiceList SS (discrete or member); else = No
    r3c2 = qc_sht.Cells(n, 21)
    item_type = qc_sht.Cells(n, 8)
    
    If InStr(item_type, "MatchSS") <> 0 Or InStr(item_type, "MatchMS") <> 0 Or InStr(item_type, "ZonesSS") <> 0 Or _
       InStr(item_type, "ZonesMS") <> 0 Or InStr(item_type, "Audio") <> 0 Or InStr(item_type, "InlineChoiceListMS") <> 0 Or _
       InStr(item_type, "InlineChoiceListSS") <> 0 Or InStr(item_type, "InlineTextChoices") <> 0 Then
        yes_type = True
    Else
        yes_type = False
    End If
    
    If yes_type And r3c2 = "No" Then
        err_cnt = err_cnt + 1
        err_list = err_list & Chr(10) & err_cnt & ". Expected R3C2 to be Yes because Item Type is " & item_type
        qc_sht.Cells(n, col_count + 8) = "FALSE"
    ElseIf yes_type = False And r3c2 = "Yes" Then
        err_cnt = err_cnt + 1
        err_list = err_list & Chr(10) & err_cnt & ". Expected R3C2 to be No because Item Type is " & item_type
        qc_sht.Cells(n, col_count + 8) = "FALSE"
    Else
        qc_sht.Cells(n, col_count + 8) = "True"
    End If
    
    ' R1C1 matches IC
    r1c1 = qc_sht.Cells(n, 13)
    test = qc_sht.Cells(n, 11)
    
    If InStr(test, r1c1) = 0 Then
        err_cnt = err_cnt + 1
        err_list = err_list & Chr(10) & err_cnt & ". Expected R1C1 to match test of " & test
        qc_sht.Cells(n, col_count + 9) = "FALSE"
    Else
        qc_sht.Cells(n, col_count + 9) = "True"
    End If
    
    ' Check that all items with ‘Listen and Match’ or ‘Sentence Builder’ in R2C4 have a value in R4C4
    r2c4 = qc_sht.Cells(n, 19)
    r4c4 = qc_sht.Cells(n, 29)
    If (InStr(r2c4, "Listen and Match") <> 0 Or InStr(r2c4, "Sentence Builder") <> 0 Or InStr(r2c4, "Read and Match") <> 0) And r4c4 = "" Then
        err_cnt = err_cnt + 1
        err_list = err_list & Chr(10) & err_cnt & ". Expected R4C4 to have a value because R2C4 is " & r2c4
        qc_sht.Cells(n, col_count + 10) = "FALSE"
    ElseIf (InStr(r2c4, "Listen and Match") = 0 And InStr(r2c4, "Sentence Builder") = 0 And InStr(r2c4, "Read and Match") = 0) And r4c4 <> "" Then
        err_cnt = err_cnt + 1
        If r2c4 = "" Then r2c4 = "[BLANK]"
        err_list = err_list & Chr(10) & err_cnt & ". Expected R4C4 to be blank because R2C4 is " & r2c4
        qc_sht.Cells(n, col_count + 10) = "FALSE"
    Else
        qc_sht.Cells(n, col_count + 10) = "True"
    End If
    
    ' Verify that all items with Yes in R3C1 have a value in R4C2
    R3C1 = qc_sht.Cells(n, 20)
    R4C2 = qc_sht.Cells(n, 27)
    If R3C1 = "Yes" And R4C2 = "" Then
        err_cnt = err_cnt + 1
        err_list = err_list & Chr(10) & err_cnt & ". Expected R4C2 to have a value because R3C1 is " & r2c4
        qc_sht.Cells(n, col_count + 11) = "FALSE"
    Else
        qc_sht.Cells(n, col_count + 11) = "True"
    End If
    
    ' Experimental types
    usercode10 = qc_sht.Cells(n, 33)
    DIW = qc_sht.Cells(n, 34)
    ' Note Application.WorksheetFunction.CountIfs(qc_sht.Cells(n, 34), qc_sht.Columns(3), qc_sht.Cells(n, 1), qc_sht.Columns(1)) <> 0
    ' returns the number of accnums that match the Don't include with accnum that are part of the IC
    If InStr(usercode10, "Experimental") <> 0 And Application.WorksheetFunction.CountIfs(qc_sht.Columns(3), qc_sht.Cells(n, 34), qc_sht.Columns(1), qc_sht.Cells(n, 1)) = 0 Then
        err_cnt = err_cnt + 1
        err_list = err_list & Chr(10) & err_cnt & ". Did not find " & DIW & " in IC " & qc_sht.Cells(n, 1)
        qc_sht.Cells(n, col_count + 12) = "FALSE"
    Else
        qc_sht.Cells(n, col_count + 12) = "True"
    End If
    
    ' R1C2 matches measure for all items
    ' For Listening, the experimental items are Speaking
    ' For Reading the experimentals are Writing
    measure = qc_sht.Cells(n, 12)
    R1C2 = qc_sht.Cells(n, 14)
    If R1C2 = "" Then R1C2 = "[BLANK]"
    If measure <> R1C2 Then
        err_cnt = err_cnt + 1
        err_list = err_list & Chr(10) & err_cnt & ". Measure of " & measure & " does not match R1C2 of " & R1C2
        qc_sht.Cells(n, col_count + 13) = "FALSE"
    ElseIf InStr(usercode10, "Experimental") <> 0 And measure = "Reading" And R1C2 <> "Writing" Then
        err_cnt = err_cnt + 1
        err_list = err_list & Chr(10) & err_cnt & ". Item is Experimental expected R1C2 to be Writing"
        qc_sht.Cells(n, col_count + 13) = "FALSE"
    ElseIf InStr(usercode10, "Experimental") <> 0 And measure = "Listening" And R1C2 <> "Speaking" Then
        err_cnt = err_cnt + 1
        err_list = err_list & Chr(10) & err_cnt & ". Item is Experimental expected R1C2 to be Speaking"
        qc_sht.Cells(n, col_count + 13) = "FALSE"
    Else
        qc_sht.Cells(n, col_count + 13) = "True"
    End If
    
    ' 11. R4C4 checked against task types in R2C4 for these grades
    '   a. Listen and Match (Listening measure) grades K-12
    '   b. Word Builder (Writing measure) grades K-5
    '   c. Read and Match (Reading measure) grades K-3
    
    If n = 272 Then
        Debug.Print "here"
    End If
    k3 = True
    If test <> "ELPA21 Grade 1" And test <> "ELPA21 Grade K" And test <> "ELPA21 Grades 2-3" Then
        k3 = False
    End If
    k5 = True
    If k3 = False And test <> "ELPA21 Grades 4-5" Then
        k5 = False
    End If
    
    r2c4 = qc_sht.Cells(n, 19)
    If k3 And r2c4 = "Read and Match" And measure <> "Reading" Then
        err_cnt = err_cnt + 1
        err_list = err_list & Chr(10) & err_cnt & ". Expected measure of Reading for K-3 items with R2C4 of Read and Match"
        qc_sht.Cells(n, col_count + 14) = "FALSE"
    Else
        qc_sht.Cells(n, col_count + 14) = "True"
    End If
    
    If k5 And r2c4 = "Word Builder" And measure <> "Writing" Then
        err_cnt = err_cnt + 1
        err_list = err_list & Chr(10) & err_cnt & ". Expected measure of Writing for K-5 items with R2C4 of Word Builder"
        qc_sht.Cells(n, col_count + 14) = "FALSE"
    Else
        qc_sht.Cells(n, col_count + 14) = "True"
    End If
    
    If r2c4 = "Listen and Match" And measure <> "Listening" Then
        err_cnt = err_cnt + 1
        err_list = err_list & Chr(10) & err_cnt & ". Expected measure of Listening for K-12 items with R2C4 of Listen and Match"
        qc_sht.Cells(n, col_count + 14) = "FALSE"
    Else
        qc_sht.Cells(n, col_count + 14) = "True"
    End If
    
    ' R3C5 should only be blank if R3C4 is Inaccessible
    R3C5 = qc_sht.Cells(n, 24)
    R3C4 = qc_sht.Cells(n, 23)
    
    If R3C5 = "" And R3C4 <> "Inaccessible" Then
        err_cnt = err_cnt + 1
        err_list = err_list & Chr(10) & err_cnt & ". Expected R3C5 to have a value because R3C4 is not ""Inaccessible"""
        qc_sht.Cells(n, col_count + 15) = "FALSE"
    Else
        qc_sht.Cells(n, col_count + 15) = "True"
    End If
    
    ' R4C1 This should be populated if the “Do not include with field” is populated with a value
    R4C1 = qc_sht.Cells(n, 26)
    If R4C1 <> "" And DIW <> "" Then
        err_cnt = err_cnt + 1
        err_list = err_list & Chr(10) & err_cnt & ". Expected R4C1 to have a value because Do Not Include With is " & DIW
        qc_sht.Cells(n, col_count + 16) = "FALSE"
    Else
        qc_sht.Cells(n, col_count + 16) = "True"
    End If
    
    ' 14. R4C3 should be Yes for ClaR2C4 values of
    '   a. Conversation
    '   b. Picture Description
    '   c. Show and Share Questions
    '   d. Show and Share Presentation
    '   e. R4C3 should be blank or No for all other R2C4 task types
    R4C3 = qc_sht.Cells(n, 28)
    r2c4 = qc_sht.Cells(n, 19)
    If r2c4 = "Conversation" Or InStr(r2c4, "Picture Description") <> 0 Or InStr(r2c4, "Show and Share Questions") <> 0 Or _
       InStr(r2c4, "Show and Share Presentation") <> 0 Then
        yes_type = True
    Else
        yes_type = False
    End If
    
    If yes_type And R4C3 = "No" Then
        err_cnt = err_cnt + 1
        err_list = err_list & Chr(10) & err_cnt & ". Expected R4C3 to be Yes because R2C4 is " & r2c4
        qc_sht.Cells(n, col_count + 17) = "FALSE"
    ElseIf yes_type = False And R4C3 = "Yes" Then
        err_cnt = err_cnt + 1
        err_list = err_list & Chr(10) & err_cnt & ". Expected R4C3 to be No because R2C4 is " & r2c4
        qc_sht.Cells(n, col_count + 17) = "FALSE"
    Else
        qc_sht.Cells(n, col_count + 17) = "True"
    End If
    
    set_status = qc_sht.Cells(n, 7)
    r1c1 = qc_sht.Cells(n, 13)
    r2c4 = qc_sht.Cells(n, 19)
    R1C2 = qc_sht.Cells(n, 14)
    usercode9 = qc_sht.Cells(n, 32)
    R1C3 = qc_sht.Cells(n, 15)
    R1C4 = qc_sht.Cells(n, 16)
    R1C7 = qc_sht.Cells(n, 35)
    R4C5 = qc_sht.Cells(n, 30)
    
    qc_sht.Cells(n, col_count + 18) = True
    qc_sht.Cells(n, col_count + 19) = True
    qc_sht.Cells(n, col_count + 20) = True
    qc_sht.Cells(n, col_count + 21) = True
    qc_sht.Cells(n, col_count + 22) = True
    qc_sht.Cells(n, col_count + 23) = True
    qc_sht.Cells(n, col_count + 24) = True
    If set_status <> "Set Leader" Then
        ' Grade - R1C1 not blank
        If r1c1 = "" Then
            err_cnt = err_cnt + 1
            err_list = err_list & Chr(10) & err_cnt & ". Expected Grade (R1C1) to have a value."
            qc_sht.Cells(n, col_count + 18) = "FALSE"
        Else
            qc_sht.Cells(n, col_count + 18) = "True"
        End If
        
        ' Task Type - R2C4 not blank
        If r2c4 = "" Then
            err_cnt = err_cnt + 1
            err_list = err_list & Chr(10) & err_cnt & ". Expected Task Type (R2C4) to have a value."
            qc_sht.Cells(n, col_count + 19) = "FALSE"
        Else
            qc_sht.Cells(n, col_count + 19) = "True"
        End If
        
        ' Domain - R1C2 not blank
        If R1C2 = "" Then
            err_cnt = err_cnt + 1
            err_list = err_list & Chr(10) & err_cnt & ". Expected Domain (R1C2) to have a value."
            qc_sht.Cells(n, col_count + 20) = "FALSE"
        Else
            qc_sht.Cells(n, col_count + 20) = "True"
        End If
        
        ' PLD - UserCode 9 not blank
        If usercode9 = "" Then
            err_cnt = err_cnt + 1
            err_list = err_list & Chr(10) & err_cnt & ". Expected PLD (UserCode 9) to have a value."
            qc_sht.Cells(n, col_count + 21) = "FALSE"
        Else
            qc_sht.Cells(n, col_count + 21) = "True"
        End If
    
        ' Standard - R1C3 not blank
        If R1C3 = "" Then
            err_cnt = err_cnt + 1
            err_list = err_list & Chr(10) & err_cnt & ". Expected Standard (R1C3) to have a value."
            qc_sht.Cells(n, col_count + 22) = "FALSE"
        Else
            qc_sht.Cells(n, col_count + 22) = "True"
        End If
        qc_sht.Cells(n, col_count + 23) = "Sub-claim (R1C4) not blank"
        qc_sht.Cells(n, col_count + 24) = "Associated Practice (R1C7) not blank"
        
        ' Sub-claim - R1C4 not blank
        If R1C4 = "" Then
            err_cnt = err_cnt + 1
            err_list = err_list & Chr(10) & err_cnt & ". Expected Sub-claim (R1C4) to have a value."
            qc_sht.Cells(n, col_count + 23) = "FALSE"
        Else
            qc_sht.Cells(n, col_count + 23) = "True"
        End If
        qc_sht.Cells(n, col_count + 24) = "Associated Practice (R1C7) not blank"
        
        ' Associated Practice - R1C7 not blank
        If R1C4 = "" Then
            err_cnt = err_cnt + 1
            err_list = err_list & Chr(10) & err_cnt & ". Expected Associated Practice (R1C7) to have a value."
            qc_sht.Cells(n, col_count + 24) = "FALSE"
        Else
            qc_sht.Cells(n, col_count + 24) = "True"
        End If
    End If
    
    ' R4C5 not blank for members
    If set_status = "Set Member" Then
        If R4C5 = "" Then
            err_cnt = err_cnt + 1
            err_list = err_list & Chr(10) & err_cnt & ". Expected R4C5 to have a value for a Set Member."
            qc_sht.Cells(n, col_count + 25) = "FALSE"
        Else
            qc_sht.Cells(n, col_count + 25) = "True"
        End If
    ElseIf R4C5 <> "" Then
        err_cnt = err_cnt + 1
        err_list = err_list & Chr(10) & err_cnt & ". Expected R4C5 to be blank for Discrete Item and Set Leaders."
        qc_sht.Cells(n, col_count + 25) = "FALSE"
    Else
        qc_sht.Cells(n, col_count + 25) = "True"
    End If

    qc_sht.Cells(n, col_count + 1) = err_list
Next n


With qc_sht
    .Range(.Columns(col_count + 1), .Columns(.UsedRange.Columns.Count)).Cut
    .Columns(1).Insert Shift:=xlToRight
End With


End Sub
Function unique(aFirstArray) As Variant
  Dim arr As New Collection, A
  Dim i As Long
  Dim uni_arr() As Variant

  On Error Resume Next
  For Each A In aFirstArray
     arr.Add A, A
  Next

  For i = 1 To arr.Count
     ReDim Preserve uni_arr(0 To q)
     uni_arr(q) = arr(i)
     q = q + 1
  Next
unique = uni_arr
End Function
Sub passMaker()
Dim main_sht As Worksheet
Dim pass1 As Worksheet
Dim pass2 As Worksheet
Dim pass3 As Worksheet
Dim subject As String
Dim measure As String
Set main_sht = ActiveSheet

Application.Workbooks.Add
Set pass1 = ActiveSheet
Application.Workbooks.Add
Set pass2 = ActiveSheet
Application.Workbooks.Add
Set pass3 = ActiveSheet

' item collection name copy to pass books
main_sht.Columns(27).Copy Destination:=pass1.Columns(1)
main_sht.Columns(27).Copy Destination:=pass2.Columns(1)
main_sht.Columns(27).Copy Destination:=pass3.Columns(1)

' item collection item sequence copy to pass books
main_sht.Columns(31).Copy Destination:=pass1.Columns(2)
main_sht.Columns(31).Copy Destination:=pass2.Columns(2)
main_sht.Columns(31).Copy Destination:=pass3.Columns(2)

' item accnum copy to pass books
main_sht.Columns(7).Copy Destination:=pass1.Columns(3)
main_sht.Columns(7).Copy Destination:=pass2.Columns(3)
main_sht.Columns(7).Copy Destination:=pass3.Columns(3)

' pass1 gets PLD, R1C3 and R1C4 need to edit R1C4 with content codes
main_sht.Columns(15).Copy Destination:=pass1.Columns(4)
main_sht.Columns(16).Copy Destination:=pass1.Columns(5)
main_sht.Columns(16).Copy Destination:=pass1.Columns(6)

' pass2 gets R1C7
main_sht.Columns(18).Copy Destination:=pass2.Columns(4)

' pass3 gets PLDs
main_sht.Columns(15).Copy Destination:=pass3.Columns(4)

pass1.Cells(1, 1) = "Item Collection Name"
pass1.Cells(1, 2) = "Item Sequence"
pass1.Cells(1, 3) = "Item Accnum"
pass1.Cells(1, 4) = "PLDs (UserCode 9)"
pass1.Cells(1, 5) = "Standards (R1C3)"
pass1.Cells(1, 6) = "SubClaims (R1C4)"
pass1.Cells(1, 7) = "unique"

pass2.Cells(1, 1) = "Item Collection Name"
pass2.Cells(1, 2) = "Item Sequence"
pass2.Cells(1, 3) = "Item Accnum"
pass2.Cells(1, 4) = "Associated Practices (R1C7)"
pass2.Cells(1, 5) = "unique"

pass3.Cells(1, 1) = "Item Collection Name"
pass3.Cells(1, 2) = "Item Sequence"
pass3.Cells(1, 3) = "Item Accnum"
pass3.Cells(1, 4) = "PLDs (UserCode 9)"
pass3.Cells(1, 5) = "unique"

For n = 2 To pass1.UsedRange.Rows.Count
    If pass1.Cells(n, 6) <> "" Then
        If InStr(pass1.Cells(n, 1), "Writing") <> 0 Then
            subject = "W"
            pass1.Cells(n, 7) = "W" & pass1.Cells(n, 6)
        ElseIf InStr(pass1.Cells(n, 1), "Speaking") <> 0 Then
            subject = "S"
            pass1.Cells(n, 7) = "S" & pass1.Cells(n, 6)
        ElseIf InStr(pass1.Cells(n, 1), "Reading") <> 0 Then
            subject = "R"
            pass1.Cells(n, 7) = "R" & pass1.Cells(n, 6)
        ElseIf InStr(pass1.Cells(n, 1), "Listening") <> 0 Then
            subject = "L"
            pass1.Cells(n, 7) = "L" & pass1.Cells(n, 6)
        End If
        pass1.Cells(n, 6) = Replace(pass1.Cells(n, 6), " ", subject & " ") & subject
    End If
Next n

For n = 2 To pass2.UsedRange.Rows.Count
    If pass2.Cells(n, 4) <> "" Then
        If InStr(pass2.Cells(n, 1), "Writing") <> 0 Then
            subject = "W"
            pass2.Cells(n, 5) = "W" & pass2.Cells(n, 4)
        ElseIf InStr(pass1.Cells(n, 1), "Speaking") <> 0 Then
            subject = "S"
            pass2.Cells(n, 5) = "S" & pass2.Cells(n, 4)
        ElseIf InStr(pass2.Cells(n, 1), "Reading") <> 0 Then
            subject = "R"
            pass2.Cells(n, 5) = "R" & pass2.Cells(n, 4)
        ElseIf InStr(pass2.Cells(n, 1), "Listening") <> 0 Then
            subject = "L"
            pass2.Cells(n, 5) = "L" & pass2.Cells(n, 4)
        End If
    End If
Next n

For n = 2 To pass3.UsedRange.Rows.Count
    If pass3.Cells(n, 4) <> "" Then
        If InStr(pass3.Cells(n, 1), "Writing") <> 0 Then
            subject = "W"
            pass3.Cells(n, 5) = "W" & pass3.Cells(n, 4)
        ElseIf InStr(pass1.Cells(n, 1), "Speaking") <> 0 Then
            subject = "S"
            pass3.Cells(n, 5) = "S" & pass3.Cells(n, 4)
        ElseIf InStr(pass3.Cells(n, 1), "Reading") <> 0 Then
            subject = "R"
            pass3.Cells(n, 5) = "R" & pass3.Cells(n, 4)
        ElseIf InStr(pass3.Cells(n, 1), "Listening") <> 0 Then
            subject = "L"
            pass3.Cells(n, 5) = "L" & pass3.Cells(n, 4)
        End If
    End If
Next n

Call Split_Unique_Values_to_Books_ELPA("Pass1-R1C3,R1C4", pass1, 7)
Call Split_Unique_Values_to_Books_ELPA("Pass2-R1C7", pass2, 5)
Call Split_Unique_Values_to_Books_ELPA("Pass3-PLDs", pass3, 5)


End Sub
Sub Split_Unique_Values_to_Books_ELPA(pass As String, sht As Worksheet, col As Long)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       03/14/2012                                        '''
'''The purpose is split all unique value sets in a selected column to new workbooks which   '''
'''are then named after the unique values.                                                  '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim this_sht As Worksheet
Dim this_array As Variant
Dim elmt As Variant
Dim file_path As String
Dim new_book As Workbook
Dim this_col As Long
Dim num As Long

Set this_sht = sht
file_path = "C:\Users\adanner\Desktop\Working Folder\ELPA Metadata"
Application.ScreenUpdating = False
this_col = col
this_array = get_unique_values(this_sht, this_col)
For Each elmt In this_array
    If elmt <> "" Then
        Application.Workbooks.Add
        Set new_book = ActiveWorkbook
        this_sht.UsedRange.AutoFilter Field:=this_col, Criteria1:=elmt
        this_sht.Rows(1).Copy
        new_book.ActiveSheet.Cells(1, 1).PasteSpecial 8
        this_sht.UsedRange.SpecialCells(xlCellTypeVisible).Copy Destination:=new_book.Sheets(1).Cells(1, 1)
        
        If new_book.ActiveSheet.Cells(2, 1) <> "" Then
            If InStr(new_book.ActiveSheet.Cells(2, 1), "Writing") <> 0 Then
                subject = "Writing"
                subW = subW + 1
                num = subW
            ElseIf InStr(new_book.ActiveSheet.Cells(2, 1), "Speaking") <> 0 Then
                subject = "Speaking"
                subS = subS + 1
                num = subS
            ElseIf InStr(new_book.ActiveSheet.Cells(2, 1), "Reading") <> 0 Then
                subject = "Reading"
                subR = subR + 1
                num = subR
            ElseIf InStr(new_book.ActiveSheet.Cells(2, 1), "Listening") <> 0 Then
                subject = "Listening"
                subL = subL + 1
                num = subL
            End If
        End If
        
        ' sort by seq
        new_book.ActiveSheet.Sort.SortFields.Clear
        new_book.ActiveSheet.Sort.SortFields.Add key:=new_book.ActiveSheet.Columns(2) _
            , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        With new_book.ActiveSheet.Sort
            .SetRange new_book.ActiveSheet.UsedRange
            .header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        
        ' col sizes
        new_book.ActiveSheet.UsedRange.Columns.AutoFit
        new_book.ActiveSheet.Columns(this_col).EntireColumn.Delete
        
        
        new_book.SaveAs filename:=file_path & "\" & "Grd_4-5_" & subject & "_IBIS_MetaDataQC-Updates_v1_D110514_" & pass & "_Set " & num & ".xlsx", FileFormat:=xlOpenXMLWorkbook
        new_book.Close
    End If
Next elmt
Application.ScreenUpdating = True
End Sub

Function get_unique_values(this_sht As Worksheet, this_col As Long) As Variant
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''                                This Macro was written by                                '''
'''                                      Asher Danner                                       '''
'''                                       03/14/2012                                        '''
'''The purpose is to return all unique values in a column as an array.                      '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim row_count As Long
Dim next_col As Long
With this_sht
    next_col = .Cells(1, .Columns.Count).End(xlToLeft).Column + 1
    
    .Columns(this_col).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=.Cells(1, next_col), unique:=True
    row_count = .Cells(.Rows.Count, next_col).End(xlUp).Row
    
    get_unique_values = WorksheetFunction.Transpose(.Range(.Cells(2, next_col), .Cells(row_count, next_col)))
    .Columns(next_col).EntireColumn.Delete
    
End With
End Function
