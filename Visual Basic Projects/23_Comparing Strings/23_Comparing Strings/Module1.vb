Module Module1

    Sub Main()
        Dim userString As String = Nothing
        Dim compString As String = "OnliveGamer"

        Console.WriteLine("Enter your string")
        userString = Console.ReadLine()

        Console.WriteLine(String.Compare(userString, compString))
        Console.WriteLine(String.Compare(compString, userString))

        Console.WriteLine(String.Compare(compString, userString, True)) 'make it case insensitive

        Console.ReadLine()

    End Sub

End Module
