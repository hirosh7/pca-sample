Private Sub OK_Click()
Dim username, password As String
username = "John123"
password = "qwertyupi#@"

If UsrTxt.Text = username And pwTxt.Text = password Then
MsgBox ("Sign in sucessful")
ElseIf UsrTxt.Text <> username Or pwTxt.Text <> password Then

MsgBox ("Sign in failed")
End If
End Sub
