Module Module1

    Sub Main()
        Dim userName As String = Nothing
        Console.WriteLine("What is your user name?")
        userName = Console.ReadLine()
        If userName.Length <= 10 Then
            Console.WriteLine("You have been granted access!")
        Else
            Console.WriteLine("You user name is not right length")
        End If

        If userName.Length.Equals(10) Then
            Console.WriteLine("It's the best length")
        End If

        Console.ReadLine()

    End Sub

End Module
