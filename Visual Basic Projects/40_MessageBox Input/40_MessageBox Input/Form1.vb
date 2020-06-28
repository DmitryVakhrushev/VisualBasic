Public Class Form1

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click

        If MessageBox.Show("Choose your option", "TITLE", MessageBoxButtons.AbortRetryIgnore, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Abort Then

            MessageBox.Show("Aborted")

        End If

    End Sub
End Class
