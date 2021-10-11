Attribute VB_Name = "ibisSurveyMacros"
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
Dim val_rng As Range
Dim lab_rng As Range

Dim sht As Worksheet
Set sht = ActiveSheet

If sht.Cells(6, 1) <> "" Then
        sht.Range(sht.Rows(6), sht.Rows(24)).Insert

End If
For n = sht.UsedRange.Rows.Count To 25 Step -1
    If sht.Cells(n, 1) = "Research and Development: Assessment Development (including AD Systems & Capabilities)" Then
        If Not val_rng Is Nothing Then
            Set val_rng = Union(val_rng, sht.Cells(n, 14))
        Else
            Set val_rng = sht.Cells(n, 14)
        End If
    ElseIf InStr(sht.Cells(n, 1), "Topic ") = 1 Then
        If Not lab_rng Is Nothing Then
            Set lab_rng = Union(lab_rng, sht.Cells(n, 1))
        Else
            Set lab_rng = sht.Cells(n, 1)
        End If
    End If
Next n

sht.Shapes.AddChart2(216, xlBarClustered).Select

ActiveChart.SeriesCollection.NewSeries
ActiveChart.FullSeriesCollection(1).Name = sht.Cells(2, 1)
ActiveChart.FullSeriesCollection(1).Values = val_rng
ActiveChart.FullSeriesCollection(1).XValues = lab_rng
    '"='Question 14 - Develop Forms'!$A$102,'Question 14 - Develop Forms'!$A$91,'Question 14 - Develop Forms'!$A$80,'Question 14 - Develop Forms'!$A$69,'Question 14 - Develop Forms'!$A$58,'Question 14 - Develop Forms'!$A$47,'Question 14 - Develop Forms'!$A$36,'Question 14 - Develop Forms'!$A$25"
'ActiveChart.Legend.Delete
ActiveChart.Axes(xlValue).MinimumScale = 0
ActiveChart.Axes(xlValue).MaximumScale = 10

End Sub
