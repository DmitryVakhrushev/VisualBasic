Module Module1

    Sub Main()
        Console.Write("Enter your username: ")
        Dim userName As String = Console.ReadLine()

        Console.Write("Enter your password: ")
        Dim password As String = Console.ReadLine()

        Console.WriteLine()

        If userName = "Dmitry" Then
            Console.WriteLine("Hello Dmitry!")
        End If

        If password = "thenewboston" Then
            Console.WriteLine("You have entered right password!!!")
        End If

        Console.ReadLine()

    End Sub

End Module