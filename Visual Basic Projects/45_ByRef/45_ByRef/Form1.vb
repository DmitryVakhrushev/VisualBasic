Public Class Form1

    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Dim num1 As Integer = 0
        MessageBox.Show(num1.ToString())
        lblVariable.Text = num1.ToString()

        incrementVariable(num1)

        MessageBox.Show(num1.ToString())
        lblVariable.Text = num1.ToString

    End Sub

    Private Sub incrementVariable(ByRef x As Integer)
        x += 5
    End Sub

End Class
