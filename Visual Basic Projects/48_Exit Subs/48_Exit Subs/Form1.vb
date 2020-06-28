Public Class Form1

    Private Sub btnExitSub_Click(sender As System.Object, e As System.EventArgs) Handles btnExitSub.Click
        Dim number As Integer = 10
        count(number)
    End Sub

    Private Sub count(ByVal num As Integer)
        While True
            listNumbers.Items.Add(num)
            num += 1
            If num > 54 Then
                Exit Sub
            End If
        End While
    End Sub
End Class
