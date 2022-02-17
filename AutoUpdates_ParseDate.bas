Attribute VB_Name = "AutoUpdates_ParseDate"
Option Explicit
Sub test()
    Dim httpObject As Object
    Dim sGetResult As String
    Dim sURL As String
    Dim sRequest As String
    Dim sItem As Variant
    
    Set httpObject = CreateObject("MSXML2.XMLHTTP")
    
    sURL = "https://api.github.com/repos/Dathmar/GeneralPurposeMacros/commits?path=General_Purpose_Macros.xlam&page=1&per_page=1"

    sRequest = sURL
    httpObject.Open "GET", sRequest, False
    httpObject.send
    sGetResult = httpObject.responseText
    sGetResult = Mid(sGetResult, 2, Len(sGetResult) - 2)
    
    Dim oJSON As Object
    Set oJSON = JSON_Tools.ParseJSON(sGetResult)
    
    Dim commit_date As String
    commit_date = oJSON("obj.commit.author.date")
    Debug.Print json_date_to_str(commit_date)
End Sub
Private Function json_date_to_str(json_date As String) As String
Dim d As Date

d = DateValue(Mid$(json_date, 1, 10)) + TimeValue(Mid(json_date, 12, 8))

json_date_to_str = d
End Function
