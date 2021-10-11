Attribute VB_Name = "Module2"
Option Explicit
Sub copy_row_x_times()
Dim r As Long
Dim wb As Workbook
Dim from_sht As Worksheet
Dim to_sht As Worksheet
Dim to_lr As Long
Dim cpy_cnt As Long
Dim c As Long
Dim val1 As String
Dim val2 As String

Set wb = ActiveWorkbook
Set from_sht = wb.Sheets(1)
Set to_sht = wb.Sheets(2)

to_sht.Rows(1) = from_sht.Rows(1)

to_lr = 2

For r = 2 To from_sht.UsedRange.Rows.Count
    val1 = from_sht.Cells(r, 1)
    cpy_cnt = from_sht.Cells(r, 2)
    val2 = from_sht.Cells(r, 3)
    
    For c = 1 To cpy_cnt
        to_sht.Cells(to_lr, 1) = val1
        to_sht.Cells(to_lr, 2) = cpy_cnt
        to_sht.Cells(to_lr, 3) = val2
        
        to_lr = to_lr + 1
    Next c
    
Next r
End Sub
Sub create_ITFAC_docs()

Dim cur_header As String
Dim r As Long
Dim c As Long
Dim wb As Workbook
Dim ws As Worksheet
Dim wDoc As Word.Document
Dim wApp As Word.Application

Set wApp = New Word.Application
Set wb = ActiveWorkbook
Set ws = wb.ActiveSheet

For r = 2 To 2
    wApp.Documents.Add Template:="Normal", NewTemplate:=False, DocumentType:=0
    wApp.Visible = True
    For c = 1 To 29
        cur_header = ws.Cells(1, c)
        With wApp.Selection
            .Style = wdStyleHeading1
            .TypeText text:=cur_header
            .TypeParagraph
            .Style = wdStyleNormal
            .TypeText text:=ws.Cells(r, c)
            .TypeParagraph
            
        End With
        
    Next c
    
Next r

End Sub
