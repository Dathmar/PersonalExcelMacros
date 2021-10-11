Attribute VB_Name = "TN_QC"
Sub TN_Form_Planner_QC()
'Things to QC
'1.  Item order
'2.  Session number
'3.  Sample items (this one may be problematic, we can discuss once a test with sample items is ready)
'4.  Use Code (aka Item Type)
'5.  Key? (We may want to hold on QCing this until we do the complete FPQC that will happen after QAI imports all metadata.)
'6.  Calculator
'7.  Scoretype this would relate to the ResponseType on our form planners, anything with a J should be handscored.
' Column        ETS_sht     TAO_sht
' Accnum        10          26
' Sequence      8           12
' Session       7           13
' Use Code      20          16
' Key           14          35
' Calculator    ~           51
'

Dim wb As Workbook
Dim ets_sht As Worksheet
Dim tao_sht As Worksheet
Dim rNote As Range
Dim ok As Integer
Dim r As Long
Dim lr_ets As Long
Dim lr_tao As Long

Dim tao_accnum As String
Dim ets_accnum As String
Dim tao_seq As String
Dim ets_seq As String
Dim tao_session As String
Dim ets_session As String
Dim tao_use As String
Dim ets_use As String
Dim tao_key As String
Dim ets_key As String
Dim tao_calc As String

Dim notes As String

Set wb = ActiveWorkbook

If wb.Sheets(1).Name <> "TAO Form Planner" And wb.Sheets(2).Name <> "ETS Form Planner" Then
    ok = MsgBox("Please use the Form Planner QC Workbook", vbOKOnly)
    Exit Sub
End If

Set ets_sht = wb.Sheets("ETS Form Planner")
Set tao_sht = wb.Sheets("TAO Form Planner")

' check that the number or rows match between sheets
lr_ets = ets_sht.Cells(ets_sht.Rows.Count, 1).End(xlUp).Row
lr_tao = tao_sht.Cells(tao_sht.Rows.Count, 2).End(xlUp).Row

' if row count does not match then report error in first note on the TAO sheet
If lr_ets <> lr_tao Then
    tao_sht.Cells(2, 1) = "Total row count does not match"
End If

' loop through each row in the TAO sheet and report errors
For r = 2 To lr_tao
    Set rNote = tao_sht.Cells(r, 1)
    notes = rNote.Value2
    
    tao_accnum = tao_sht.Cells(r, 26)
    ets_accnum = ets_sht.Cells(r, 10)
    tao_seq = tao_sht.Cells(r, 12)
    ets_seq = ets_sht.Cells(r, 8)
    tao_session = tao_sht.Cells(r, 13)
    ets_session = ets_sht.Cells(r, 7)
    tao_use = tao_sht.Cells(r, 16)
    ets_use = ets_sht.Cells(r, 20)
    tao_key = tao_sht.Cells(r, 35)
    ets_key = ets_sht.Cells(r, 14)
    tao_calc = tao_sht.Cells(r, 51)
    
    ' accnum check
    If tao_accnum <> ets_accnum Then
        notes = update_notes(notes, "Accnums do not match")
    End If
    
    ' sequence check
    If tao_seq <> ets_seq Then
        notes = update_notes(notes, "Sequences do not match")
    End If
    
    ' session check
    If tao_session <> ets_session Then
        notes = update_notes(notes, "Sessions do not match")
    End If
    
    ' use code check
    If ets_to_tao_use(ets_use) = "Unknown" Then
        notes = update_notes(notes, "Unsupported ETS Use Code")
    ElseIf ets_to_tao_use(ets_use) <> tao_use Then
        notes = update_notes(notes, "Uses do not match")
    End If
    
    ' key check
    If ets_key <> tao_key Then
        notes = update_notes(notes, "Keys do not match")
    End If
    
    rNote.Value2 = notes
    ' calculator check
    ' not complete
    
    ' score type check
    ' not complete
Next r
End Sub
Function update_notes(cur_notes As String, add_notes As String) As String
If cur_notes = "" Then
    update_notes = add_notes
Else
    update_notes = cur_notes & Chr(10) & add_notes
End If
End Function
Function ets_to_tao_use(ets_use As String) As String
Select Case ets_use

    Case "F"
        ets_to_tao_use = "FT"
    Case "L"
        ets_to_tao_use = "IA/OP"
    Case "O"
        ets_to_tao_use = "OP"
    Case Else
        ets_to_tao_use = "Unknown"
        
End Select


End Function
