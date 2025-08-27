Attribute VB_Name = "Module2"
Sub Alternative_Afra()
    Dim strNum As String
    strNum = InputBox("Enter the alternative number:", "Alternative Number")
    If strNum <> "" Then
        ActiveInspector.WordEditor.Application.Selection.InsertAfter "Hi Afra," & vbCrLf & vbCrLf & "Hope you're doing well. An alternative for this would be " & strNum & ". Please see the attached datasheet."
    End If
End Sub

