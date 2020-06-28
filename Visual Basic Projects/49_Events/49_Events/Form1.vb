Public Class Form1

    Private Sub btnEvents_Click(sender As System.Object, e As System.EventArgs) Handles btnEvents.Click
        MessageBox.Show(sender.ToString()) 'checking what "sender" is in the argument
        MessageBox.Show(e.ToString()) ' checking what "e" is in the argument
    End Sub

    Private Sub btnEvents_MouseHover(sender As System.Object, e As System.EventArgs) Handles btnEvents.MouseHover
        btnEvents.Text = "Text changed"
        'MessageBox.Show(e.ToString())

    End Sub
End Class
