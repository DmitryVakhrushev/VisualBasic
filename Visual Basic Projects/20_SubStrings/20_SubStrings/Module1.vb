Module Module1

    Sub Main()
        Dim userString As String = Nothing
        Console.WriteLine("Enter a string")
        userString = Console.ReadLine()

        Console.WriteLine()
        Console.WriteLine(userString.Length.ToString())
        Console.WriteLine(userString.Substring(0, 3)) 'prints 1st, 2nd, 3rd symbols
        Console.WriteLine(userString.Substring(3, 3)) 'prints 4th, 5th, 6th symbols
        Console.WriteLine(userString.Substring(6)) 'print string starting from 7 till the end

        Console.ReadLine()

    End Sub

End Module
