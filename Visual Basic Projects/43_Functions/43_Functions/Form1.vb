Public Class Form1

    Private Sub btnRunFunction_Click(sender As System.Object, e As System.EventArgs) Handles btnRunFunction.Click
        Dim answer As Double = solveMath() 'assigning the returning value to the variable
        MessageBox.Show(answer)
    End Sub

    Private Function solveMath() As Double
        Dim num1 As Integer = 23
        Dim num2 As Integer = 5
        Return num1 / num2 'returning the value of the function

    End Function

End Class
