Public Class Form1

    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Dim userName As String = Nothing

        userName = InputBox("What is your name?", "Hello", "My Default Response", 320, 250)

        lblHelloUser.Text = "Hello, " & userName

    End Sub
End Class
