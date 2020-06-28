Public Class btnShowMessage

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click

        MessageBox.Show("This is a bare message")

        MessageBox.Show("My message has a title", "Title Chebika")

        MessageBox.Show("With an ICON", "MYTITLE", MessageBoxButtons.OK, MessageBoxIcon.Information)

        MsgBox("Message")
    End Sub
End Class
