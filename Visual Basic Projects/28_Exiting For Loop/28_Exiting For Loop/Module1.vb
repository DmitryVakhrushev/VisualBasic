Module Module1

    Sub Main()
        Console.WriteLine("Exiting For Loop")
        For num1 = 1 To 30 Step 4
            If num1 > 23 Then
                Exit For
            End If
            Console.WriteLine(num1)
        Next

        Console.ReadLine()
    End Sub

End Module
