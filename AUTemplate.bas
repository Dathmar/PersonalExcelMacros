Attribute VB_Name = "AUTemplate"
Sub list_au_ids()

Dim ws As Worksheet
Dim rCell As Range
Dim rng As Range
Dim lr As Long
Dim val As String
Dim au_lst As String
Dim this_au As String
au_lst = ""
Set ws = ActiveSheet
lr = ws.Cells(ws.Rows.Count, 16).End(xlUp).Row

Set rng = ws.Range(ws.Cells(1, 16), ws.Cells(lr, 16))

For Each rCell In rng
val = rCell
Do While InStr(val, "TQ") <> 0
    this_au = Mid(val, InStr(val, "TQ"), 8)
    
    If au_lst = "" Then
        au_lst = this_au
    Else
        au_lst = au_lst & "|" & this_au
    End If
    val = Replace(val, this_au, "")
Loop
Do While InStr(val, "TS") <> 0
    this_au = Mid(val, InStr(val, "TS"), 8)
    
    If au_lst = "" Then
        au_lst = this_au
    Else
        au_lst = au_lst & "|" & this_au
    End If
    val = Replace(val, this_au, "")
Loop
rCell.Offset(0, 1) = au_lst
au_lst = ""
Next rCell

End Sub
