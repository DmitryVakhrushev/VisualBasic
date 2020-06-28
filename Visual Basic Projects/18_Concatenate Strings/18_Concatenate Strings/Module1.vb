Module Module1

    Sub Main()
        Dim userString As String = Nothing
        Dim programString As String = " catfish"

        Console.Write("Enter what you want: ")
        userString = Console.ReadLine()

        userString = userString + programString + " additional words"

        Console.WriteLine(userString & " hello youtube")
        Console.ReadLine()

    End Sub

End Module
