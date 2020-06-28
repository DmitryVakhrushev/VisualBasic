Public Class Form1

    Private Sub btnCoercion_Click(sender As System.Object, e As System.EventArgs) Handles btnCoercion.Click
        'Dim myNum As Integer = 23
        Dim myNum As Double = 23.56
        showDataType(myNum)
    End Sub

    Private Sub showDataType(ByVal num As Integer)
        MessageBox.Show("This is an integer")
    End Sub

    Private Sub showDataType(ByVal num As Double)
        'MessageBox.Show("This is a double")
        txtN.Text = "This is a double"
    End Sub

End Class
