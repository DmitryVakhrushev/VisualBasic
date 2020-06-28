Module Module1

    Sub Main()
        Console.Write("Enter your name: ")
        Dim userName As String = Console.ReadLine()
        Console.Write("Enter your password: ")
        Dim password As String = Console.ReadLine

        If userName = "Dmitry" Then
            Console.WriteLine("Welcome Dmitry!!!")
        ElseIf userName = "Sam" Then
            Console.WriteLine("Hello Sam!")
        Else
            Console.WriteLine("You are not a valid user")
        End If

        Console.ReadLine()

    End Sub

End Module
