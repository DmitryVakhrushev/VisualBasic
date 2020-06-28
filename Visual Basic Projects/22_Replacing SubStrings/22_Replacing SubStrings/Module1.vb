Module Module1

    Sub Main()
        Dim myString As String = Nothing
        Dim finalString As String = Nothing

        Console.WriteLine("Enter your string")
        myString = Console.ReadLine()
        finalString = myString.Replace("a", "z") 'replacing a with z in the string
        Console.WriteLine(finalString)
        Console.ReadLine()

    End Sub

End Module
