Module Module1

    Sub Main()
        Dim userName As String = Nothing
        Dim password As String = Nothing

        Console.Write("Enter your name: ")
        userName = Console.ReadLine()
        Console.Write("Enter your password: ")
        password = Console.ReadLine()

        If (userName = "Sam" Or userName = "Tim") And password = "thenewboston" Then
            Console.WriteLine("Welcome " & userName)
        Else
            Console.WriteLine("You've entered incorrect username or password")
        End If

        Console.ReadLine()

    End Sub

End Module
