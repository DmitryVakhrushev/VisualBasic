Module Module1

    Sub Main()

        Console.WriteLine("Enter your string")
        Dim myString As String = Console.ReadLine()

        Console.WriteLine("Enter your decimal value")
        Dim myDouble As Double = Console.ReadLine()

        Console.WriteLine()

        Console.WriteLine(String.Format("{0:n2}", myDouble))
        Console.WriteLine(myString.ToLower())
        Console.WriteLine(myString.ToUpper())

        Console.ReadLine()

    End Sub

End Module
