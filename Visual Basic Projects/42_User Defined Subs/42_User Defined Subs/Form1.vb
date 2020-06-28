Public Class Form1

    Private Sub btnRunSub_Click(sender As System.Object, e As System.EventArgs) Handles btnRunSub.Click
        addNumbers()
    End Sub

    'Private means that this sub cannot be accessed by any other class
    'outside of the form Form1.vb
    Private Sub addNumbers()
        Dim num1 As Integer = 2341
        Dim num2 As Integer = 5233
        MessageBox.Show(num1 + num2)
    End Sub
End Class
