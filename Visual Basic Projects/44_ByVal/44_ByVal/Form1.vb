Public Class Form1

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles btnSubtract.Click

        Dim answer As Double = subtractNumbers(TextBox1.Text, TextBox2.Text)
        MessageBox.Show(answer)

    End Sub

    'when the function subtractNumber is called it will require two values
    Private Function subtractNumbers(ByVal num1 As Double, ByVal num2 As Double) As Double

        Return num1 - num2

    End Function

End Class
