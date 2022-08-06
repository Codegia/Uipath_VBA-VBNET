Function GetComment(ArgSheetName, ArgCell)
Dim StrComment As String
Dim c As Comment
With Sheets(ArgSheetName).Select
End With
With Range(ArgCell)
    On Error Resume Next
    Set c = .Comment
    If c Is Nothing Then
        GetComment = "Success"
    Else
        StrComment = c.Text
        GetComment = StrComment
    End If
End With
End Function