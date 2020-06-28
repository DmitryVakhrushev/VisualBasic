Public Class Form1

    Private Sub btnShowMessage_Click(sender As System.Object, e As System.EventArgs) Handles btnShowMessage.Click

        Dim mesg As String = txtMessage.Text 'it will take the input from the second field
        Dim title As String = Nothing

        'we have two forms: txtTitle and txtMessage
        If txtTitle.TextLength > 0 Then
            title = txtTitle.Text
            showMessage(mesg, title) 'running a sub with message and title
        Else
            showMessage(mesg) 'running a sub with message and DEFAULT Title
        End If

    End Sub

    Private Sub showMessage(ByVal message As String, Optional ByVal title As String = "Default Title")

        MessageBox.Show(message, title)

    End Sub

End Class
