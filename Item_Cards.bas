Attribute VB_Name = "Item_Cards"
Sub STAAR_Item_Card_Info()
Dim to_book As Workbook
Dim wd_file_name As Variant
Dim last_row As Long
Dim t As Long
Dim table_count As Long
Dim correct_table As Boolean
Dim item_code As String
Dim item_type As String
Dim item_writer As String
Dim rc As String
Dim ks As String
Dim se As String
Dim rs As String
Dim DOK As String
Dim sit As String
Dim key As String

wd_file_name = Application.GetOpenFilename("Word files (*.do*),*.do*", , _
    "Browse for file containing table to be imported", MultiSelect:=True)

Application.ScreenUpdating = False
With ActiveWorkbook.ActiveSheet
    .Cells(1, 1) = "Item Code"
    .Cells(1, 2) = "Item Type"
    .Cells(1, 3) = "Item Writer"
    .Cells(1, 4) = "Reporting Category"
    .Cells(1, 5) = "Knowledge and Skill"
    .Cells(1, 6) = "Student Expectation"
    .Cells(1, 7) = "Readiness or Supporting"
    .Cells(1, 8) = "DOK"
    .Cells(1, 9) = "Special Item Type"
    .Cells(1, 10) = "key"
    last_row = .Cells(.Rows.Count, 1).End(xlUp).Row
End With


On Error GoTo 0
If IsArray(wd_file_name) Then
    For i = LBound(wd_file_name) To UBound(wd_file_name)
        Application.StatusBar = "Opening... " & wd_file_name(i)
        Set wd_doc = GetObject(wd_file_name(i))
        'verify correct table
        table_count = wd_doc.Tables.Count
        t = 1 ' set table to 1
        'If table_count > 1 Then ' only check for correct table if there is more than 1 table
        '    correct_table = False
        '    While correct_table = False ' loop through tables until the item info table is found
        '        Debug.Print wd_doc.tables(t).Cell(1, 3).Range.Text
        '        If InStr(wd_doc.tables(t).Cell(1, 3).Range.Text, "Program") = 0 Then
        '            t = t + 1
        '        Else
        '            correct_table = True
        '        End If
        '    Wend
        'End If
        
        ' pull information from the item
        Application.StatusBar = "Processing... " & wd_file
        item_code = Left(wd_doc.Tables(t).Cell(3, 4).Range.text, Len(wd_doc.Tables(t).Cell(3, 4).Range.text) - 1) ' item code
        item_type = Left(wd_doc.Tables(t).Cell(6, 2).Range.text, Len(wd_doc.Tables(t).Cell(6, 2).Range.text) - 1) ' Item Type
        item_writer = Left(wd_doc.Tables(t).Cell(5, 4).Range.text, Len(wd_doc.Tables(t).Cell(5, 4).Range.text) - 1) ' Item Writer
        rc = Left(wd_doc.Tables(t).Cell(11, 2).Range.text, Len(wd_doc.Tables(t).Cell(11, 2).Range.text) - 1) ' Reporting Category
        ks = Left(wd_doc.Tables(t).Cell(12, 2).Range.text, Len(wd_doc.Tables(t).Cell(12, 2).Range.text) - 1)  ' Knowledge and Skill
        se = Left(wd_doc.Tables(t).Cell(12, 4).Range.text, Len(wd_doc.Tables(t).Cell(12, 4).Range.text) - 1)  ' Student Expectation
        rs = Left(wd_doc.Tables(t).Cell(13, 2).Range.text, Len(wd_doc.Tables(t).Cell(13, 2).Range.text) - 1)  ' Readiness or Supporting
        DOK = Left(wd_doc.Tables(t).Cell(14, 4).Range.text, Len(wd_doc.Tables(t).Cell(14, 4).Range.text) - 1) ' DOK
        sit = Left(wd_doc.Tables(t).Cell(15, 2).Range.text, Len(wd_doc.Tables(t).Cell(15, 2).Range.text) - 1)  ' Special Item Type
        key = Left(wd_doc.Tables(t).Cell(22, 2).Range.text, Len(wd_doc.Tables(t).Cell(22, 2).Range.text) - 1)  ' key
        
        With ActiveWorkbook.ActiveSheet
            .Cells(last_row + i, 1) = Trim(Replace(item_code, Chr(13), ""))
            .Cells(last_row + i, 2) = Trim(Replace(item_type, Chr(13), ""))
            .Cells(last_row + i, 3) = Trim(Replace(item_writer, Chr(13), ""))
            .Cells(last_row + i, 4) = Trim(Replace(rc, Chr(13), ""))
            .Cells(last_row + i, 5) = Trim(Replace(ks, Chr(13), ""))
            .Cells(last_row + i, 6) = Trim(Replace(se, Chr(13), ""))
            .Cells(last_row + i, 7) = Trim(Replace(rs, Chr(13), ""))
            .Cells(last_row + i, 8) = Trim(Replace(DOK, Chr(13), ""))
            .Cells(last_row + i, 9) = Trim(Replace(sit, Chr(13), ""))
            .Cells(last_row + i, 10) = Trim(Replace(key, Chr(13), ""))
        End With
        Set wd_doc = Nothing
    Next i
End If
Application.StatusBar = False
Application.ScreenUpdating = True
Exit Sub

blah:
Set wd_doc = Nothing
Application.StatusBar = False
Application.ScreenUpdating = True
End Sub
Sub get_last_chr()
Cells(4, 4) = Replace(Cells(4, 4), Chr(13), "")
End Sub

Sub FL_Item_Card_Info()
Dim to_book As Workbook
Dim wd_file_name As Variant
Dim last_row As Long
Dim t As Long
Dim table_count As Long
Dim correct_table As Boolean
Dim item_code As String
Dim benchmark As String
Dim DOK As String
Dim item_type As String
Dim CCSS As String

wd_file_name = Application.GetOpenFilename("Word files (*.do*),*.do*", , _
    "Browse for file containing table to be imported", MultiSelect:=True)

Application.ScreenUpdating = False
With ActiveWorkbook.ActiveSheet
    .Cells(1, 1) = "Item Code"
    .Cells(1, 2) = "Benchmark"
    .Cells(1, 3) = "DOK"
    .Cells(1, 4) = "Item Type"
    .Cells(1, 5) = "CCSS"
    last_row = .Cells(.Rows.Count, 1).End(xlUp).Row
End With


On Error GoTo 0
If IsArray(wd_file_name) Then
    For i = LBound(wd_file_name) To UBound(wd_file_name)
        Application.StatusBar = "Opening... " & wd_file_name(i)
        Set wd_doc = GetObject(wd_file_name(i))
        'verify correct table
        table_count = wd_doc.Tables.Count
        t = 1 ' set table to 1
        If table_count > 1 Then ' only check for correct table if there is more than 1 table
            correct_table = False
            While correct_table = False ' loop through tables until the item info table is found
                If InStr(wd_doc.Tables(t).Cell(1, 1).Range.text, "Item ") = 0 Then
                    t = t + 1
                Else
                    correct_table = True
                End If
            Wend
        End If
        
        ' pull information from the item
        Application.StatusBar = "Processing... " & wd_file
        item_code = Left(wd_doc.Tables(t).Cell(17, 2).Range.text, Len(wd_doc.Tables(t).Cell(17, 2).Range.text) - 1) ' item code
        benchmark = Left(wd_doc.Tables(t).Cell(7, 2).Range.text, Len(wd_doc.Tables(t).Cell(7, 2).Range.text) - 1) ' benchmark
        DOK = Left(wd_doc.Tables(t).Cell(3, 4).Range.text, Len(wd_doc.Tables(t).Cell(3, 4).Range.text) - 1) ' DOK
        item_type = Left(wd_doc.Tables(t).Cell(2, 4).Range.text, Len(wd_doc.Tables(t).Cell(2, 4).Range.text) - 1) ' item type
        CCSS = Left(wd_doc.Tables(t).Cell(9, 2).Range.text, Len(wd_doc.Tables(t).Cell(9, 2).Range.text) - 1) ' CCSS
        
        With ActiveWorkbook.ActiveSheet
            .Cells(last_row + i, 1) = Trim(Replace(item_code, Chr(13), ""))
            .Cells(last_row + i, 2) = Trim(Replace(benchmark, Chr(13), ""))
            .Cells(last_row + i, 3) = Trim(Replace(DOK, Chr(13), ""))
            .Cells(last_row + i, 4) = Trim(Replace(item_type, Chr(13), ""))
            .Cells(last_row + i, 5) = Trim(Replace(CCSS, Chr(13), ""))
        End With
        Set wd_doc = Nothing
    Next i
End If
Application.StatusBar = False
Application.ScreenUpdating = True
Exit Sub

blah:
MsgBox "blah " & wd_file
Set wd_doc = Nothing
Application.StatusBar = False
Application.ScreenUpdating = True
End Sub

Sub VA_Item_Card_Info()
Dim to_book As Workbook
Dim wd_file_name As Variant
Dim last_row As Long
Dim t As Long
Dim table_count As Long
Dim correct_table As Boolean
Dim item_code As String
Dim sol As String
Dim key As String
Dim diff As String
Dim cog As String

wd_file_name = Application.GetOpenFilename("Word files (*.do*),*.do*", , _
"Browse for file containing table to be imported", MultiSelect:=True)

Application.ScreenUpdating = False
With ActiveWorkbook.ActiveSheet
    .Cells(1, 1) = "Item Code"
    .Cells(1, 2) = "SOL"
    .Cells(1, 3) = "Key"
    .Cells(1, 4) = "Diff"
    .Cells(1, 5) = "Cog"
    last_row = .Cells(.Rows.Count, 1).End(xlUp).Row
End With


On Error GoTo 0
If IsArray(wd_file_name) Then
    For i = LBound(wd_file_name) To UBound(wd_file_name)
        Application.StatusBar = "Opening... " & wd_file_name(i)
        Set wd_doc = GetObject(wd_file_name(i))
        'verify correct table
        table_count = wd_doc.Tables.Count
        t = 1 ' set table to 1
        If table_count > 1 Then ' only check for correct table if there is more than 1 table
            correct_table = False
            While correct_table = False ' loop through tables until the item info table is found
                If InStr(wd_doc.Tables(t).Cell(1, 1).Range.text, "Item ") = 0 Then
                    t = t + 1
                Else
                    correct_table = True
                End If
            Wend
        End If
        
        ' pull information from the item
        Application.StatusBar = "Processing... " & wd_file
        item_code = Left(wd_doc.Tables(t).Cell(11, 2).Range.text, Len(wd_doc.Tables(t).Cell(11, 2).Range.text) - 1) ' item code
        sol = Left(wd_doc.Tables(t).Cell(3, 6).Range.text, Len(wd_doc.Tables(t).Cell(3, 6).Range.text) - 1) ' benchmark
        key = Left(wd_doc.Tables(t).Cell(2, 2).Range.text, Len(wd_doc.Tables(t).Cell(2, 2).Range.text) - 1) ' DOK
        diff = Left(wd_doc.Tables(t).Cell(2, 4).Range.text, Len(wd_doc.Tables(t).Cell(2, 4).Range.text) - 1) ' item type
        cog = Left(wd_doc.Tables(t).Cell(2, 6).Range.text, Len(wd_doc.Tables(t).Cell(2, 6).Range.text) - 1) ' CCSS
        
        With ActiveWorkbook.ActiveSheet
            .Cells(last_row + i, 1) = Trim(Replace(item_code, Chr(13), ""))
            .Cells(last_row + i, 2) = Trim(Replace(sol, Chr(13), ""))
            .Cells(last_row + i, 3) = Trim(Replace(key, Chr(13), ""))
            .Cells(last_row + i, 4) = Trim(Replace(diff, Chr(13), ""))
            .Cells(last_row + i, 5) = Trim(Replace(cog, Chr(13), ""))
        End With
        Set wd_doc = Nothing
    Next i
End If
Application.StatusBar = False
Application.ScreenUpdating = True
End Sub
Sub WCAP_Item_Card_Info()
Dim to_book As Workbook
Dim wd_file_name As Variant
Dim last_row As Long
Dim wd_file As String
Dim t As Long
Dim table_count As Long
Dim correct_table As Boolean
Dim rng As Range
wd_file_name = Application.GetOpenFilename("Word files (*.do*),*.do*", , _
"Browse for file containing table to be imported", MultiSelect:=True)

Application.ScreenUpdating = False
Cells(1, 1) = "Item Code"
Cells(1, 2) = "Key"
Cells(1, 3) = "Item Spec"
Cells(1, 4) = "CC"
Cells(1, 5) = "Item Type"

last_row = ActiveSheet.UsedRange.Rows.Count

If IsArray(wd_file_name) Then
    For i = LBound(wd_file_name) To UBound(wd_file_name)
        Application.StatusBar = "Opening... " & wd_file_name(i)
        Set wd_doc = GetObject(wd_file_name(i))
        
        ' pull information from the item
        Application.StatusBar = "Processing... " & wd_file
        Set rng = wd_doc.Sections(1).Headers(1).Range
        With rng.Tables(1)
            Cells(last_row + i, 1) = Left(.Cell(1, 2).Range.text, Len(.Cell(1, 2).Range.text) - 1) ' item code
            Cells(last_row + i, 2) = Left(.Cell(1, 4).Range.text, Len(.Cell(1, 4).Range.text) - 1) ' Key
            Cells(last_row + i, 3) = Left(.Cell(2, 2).Range.text, Len(.Cell(2, 2).Range.text) - 1) ' Spec
            Cells(last_row + i, 4) = Left(.Cell(2, 3).Range.text, Len(.Cell(2, 3).Range.text) - 1) ' cc
            Cells(last_row + i, 5) = Left(.Cell(2, 5).Range.text, Len(.Cell(2, 5).Range.text) - 1) ' Item type
        End With

        Set wd_doc = Nothing
    Next i
End If
Application.StatusBar = False
Application.ScreenUpdating = True
End Sub
Sub MD_Item_Card_Info()
Dim to_book As Workbook
Dim wd_file_name As Variant
Dim last_row As Long
Dim t As Long
Dim i As Long
Dim table_count As Long
Dim correct_table As Boolean
Dim item_code As String
Dim CLG As String
Dim limits As String
Dim key As String
Dim passage_id As String
Dim passage As String
Dim response_type As String
Dim item_writer As String
Dim item_type As String
Dim rCell As Range

wd_file_name = Application.GetOpenFilename("Word files (*.do*),*.do*", , _
    "Browse for file containing table to be imported", MultiSelect:=True)

Application.ScreenUpdating = False
On Error GoTo 0
If IsArray(wd_file_name) Then
    For i = LBound(wd_file_name) To UBound(wd_file_name)
        Application.StatusBar = "Opening... " & wd_file_name(i)
        Set wd_doc = GetObject(wd_file_name(i))
        'verify correct table
        table_count = wd_doc.Tables.Count
        t = 1 ' set table to 1
        If table_count > 1 Then ' only check for correct table if there is more than 1 table
            correct_table = False
            While correct_table = False ' loop through tables until the item info table is found
                If InStr(wd_doc.Tables(t).Cell(1, 1).Range.text, "Item ") = 0 Then
                    t = t + 1
                Else
                    correct_table = True
                End If
            Wend
        End If
        
        ' pull information from the item
        Application.StatusBar = "Processing... " & wd_file
        item_code = Left(wd_doc.Tables(t).Cell(22, 2).Range.text, Len(wd_doc.Tables(t).Cell(22, 2).Range.text) - 1)
        CLG = Left(wd_doc.Tables(t).Cell(6, 2).Range.text, Len(wd_doc.Tables(t).Cell(6, 2).Range.text) - 1)
        limits = Left(wd_doc.Tables(t).Cell(10, 2).Range.text, Len(wd_doc.Tables(t).Cell(10, 2).Range.text) - 1)
        key = Left(wd_doc.Tables(t).Cell(4, 2).Range.text, Len(wd_doc.Tables(t).Cell(4, 2).Range.text) - 1)
        passage_id = Left(wd_doc.Tables(t).Cell(2, 4).Range.text, Len(wd_doc.Tables(t).Cell(2, 4).Range.text) - 1)
        passage = Left(wd_doc.Tables(t).Cell(2, 2).Range.text, Len(wd_doc.Tables(t).Cell(2, 2).Range.text) - 1)
        response_type = Left(wd_doc.Tables(t).Cell(4, 4).Range.text, Len(wd_doc.Tables(t).Cell(4, 4).Range.text) - 1)
        item_writer = Left(wd_doc.Tables(t).Cell(1, 2).Range.text, Len(wd_doc.Tables(t).Cell(1, 2).Range.text) - 1)
        
        
        With ActiveWorkbook.ActiveSheet
            last_row = .UsedRange.Rows.Count + 1
            .Cells(last_row, 2) = Trim(Replace(item_code, Chr(13), ""))
            .Cells(last_row, 3) = Trim(Replace(CLG, Chr(13), ""))
            .Cells(last_row, 4) = Trim(Replace(limits, Chr(13), ""))
            .Cells(last_row, 5) = Trim(Replace(key, Chr(13), ""))
            .Cells(last_row, 6) = Trim(Replace(passage_id, Chr(13), ""))
            .Cells(last_row, 7) = Trim(Replace(passage, Chr(13), ""))
            .Cells(last_row, 10) = Trim(Replace(response_type, Chr(13), ""))
            .Cells(last_row, 17) = Trim(Replace(item_writer, Chr(13), ""))
        End With
        Set wd_doc = Nothing
    Next i
End If
Application.StatusBar = False
Application.ScreenUpdating = True
Exit Sub

blah:
MsgBox "Not able to find correct table in " & wd_doc
Set wd_doc = Nothing
Application.StatusBar = False
Application.ScreenUpdating = True
End Sub
Sub WY_Math_Item_Card_Info()
Dim to_book As Workbook
Dim wd_file_name As Variant
Dim last_row As Long
Dim t As Long
Dim table_count As Long
Dim correct_table As Boolean
Dim item_code As String
Dim benchmark As String
Dim DOK As String
Dim item_type As String
Dim CCSS As String

wd_file_name = Application.GetOpenFilename("Word files (*.do*),*.do*", , _
    "Browse for file containing table to be imported", MultiSelect:=True)

Application.ScreenUpdating = False
With ActiveWorkbook.ActiveSheet
    .Cells(1, 2) = "Item Code"
    .Cells(1, 3) = "Item Writer"
    .Cells(1, 4) = "Test ID"
    .Cells(1, 5) = "Key"
    .Cells(1, 6) = "CCSS CLUSTER"
    .Cells(1, 7) = "CCSS"
    .Cells(1, 8) = "Mathematical Practices"
    .Cells(1, 9) = "Calculartor Use"
    .Cells(1, 10) = "DOK"
    .Cells(1, 11) = "Difficulty"
    .Cells(1, 12) = "Art"
    last_row = .Cells(.Rows.Count, 1).End(xlUp).Row
End With


On Error GoTo 0
If IsArray(wd_file_name) Then
    For i = LBound(wd_file_name) To UBound(wd_file_name)
        Application.StatusBar = "Opening... " & wd_file_name(i)
        Set wd_doc = GetObject(wd_file_name(i))
        'verify correct table
        table_count = wd_doc.Tables.Count
        t = 1 ' set table to 1
        If table_count > 1 Then ' only check for correct table if there is more than 1 table
            correct_table = False
            While correct_table = False ' loop through tables until the item info table is found
                If InStr(wd_doc.Tables(t).Cell(1, 1).Range.text, "Item ") = 0 Then
                    t = t + 1
                Else
                    correct_table = True
                End If
            Wend
        End If
        
        ' pull information from the item
        Application.StatusBar = "Processing... " & wd_file
        item_code = Left(wd_doc.Tables(t).Cell(11, 2).Range.text, Len(wd_doc.Tables(t).Cell(11, 2).Range.text) - 1) ' item code
        Writer = Left(wd_doc.Tables(t).Cell(1, 2).Range.text, Len(wd_doc.Tables(t).Cell(1, 2).Range.text) - 1) ' item writer
        testID = Left(wd_doc.Tables(t).Cell(2, 2).Range.text, Len(wd_doc.Tables(t).Cell(2, 2).Range.text) - 1) ' test id
        key = Left(wd_doc.Tables(t).Cell(2, 6).Range.text, Len(wd_doc.Tables(t).Cell(2, 6).Range.text) - 1) ' key
        CCSSCluster = Left(wd_doc.Tables(t).Cell(3, 6).Range.text, Len(wd_doc.Tables(t).Cell(3, 6).Range.text) - 1) ' CCSS Cluster
        CCSS = Left(wd_doc.Tables(t).Cell(4, 2).Range.text, Len(wd_doc.Tables(t).Cell(4, 2).Range.text) - 1) ' CCSS Cluster
        mathPrac = Left(wd_doc.Tables(t).Cell(4, 4).Range.text, Len(wd_doc.Tables(t).Cell(4, 4).Range.text) - 1) ' Mathematical Practices
        calcUse = Left(wd_doc.Tables(t).Cell(4, 6).Range.text, Len(wd_doc.Tables(t).Cell(4, 6).Range.text) - 1) ' Calculator Use
        DOK = Left(wd_doc.Tables(t).Cell(5, 2).Range.text, Len(wd_doc.Tables(t).Cell(5, 2).Range.text) - 1) ' DOK
        difficulty = Left(wd_doc.Tables(t).Cell(5, 4).Range.text, Len(wd_doc.Tables(t).Cell(5, 4).Range.text) - 1) ' Difficulty
        art = Left(wd_doc.Tables(t).Cell(5, 6).Range.text, Len(wd_doc.Tables(t).Cell(5, 6).Range.text) - 1) ' Art
        
        With ActiveWorkbook.ActiveSheet
            .Cells(last_row + i, 2) = Trim(Replace(item_code, Chr(13), ""))
            .Cells(last_row + i, 3) = Trim(Replace(Writer, Chr(13), ""))
            .Cells(last_row + i, 4) = Trim(Replace(testID, Chr(13), ""))
            .Cells(last_row + i, 5) = Trim(Replace(key, Chr(13), ""))
            .Cells(last_row + i, 6) = Trim(Replace(CCSSCluster, Chr(13), ""))
            .Cells(last_row + i, 7) = Trim(Replace(CCSS, Chr(13), ""))
            .Cells(last_row + i, 8) = Trim(Replace(mathPrac, Chr(13), ""))
            .Cells(last_row + i, 9) = Trim(Replace(calcUse, Chr(13), ""))
            .Cells(last_row + i, 10) = Trim(Replace(DOK, Chr(13), ""))
            .Cells(last_row + i, 11) = Trim(Replace(difficulty, Chr(13), ""))
            .Cells(last_row + i, 12) = Trim(Replace(art, Chr(13), ""))
        End With
        Set wd_doc = Nothing
    Next i
End If
Application.StatusBar = False
Application.ScreenUpdating = True
Exit Sub

blah:
MsgBox "I broke on file " & wd_file & vbNewLine & "Contact Asher"
Set wd_doc = Nothing
Application.StatusBar = False
Application.ScreenUpdating = True
End Sub
