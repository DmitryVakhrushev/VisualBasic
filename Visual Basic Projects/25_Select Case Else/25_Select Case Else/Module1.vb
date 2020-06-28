Module Module1

    Sub Main()
        Dim myString As String = Nothing
        Console.WriteLine("Enter you string")
        myString = Console.ReadLine()

        Select Case myString.ToLower
            Case "hello"
                Console.WriteLine("Goodbye")
            Case "fishing"
                Console.WriteLine("Boat")
            Case Else
                Console.WriteLine("No ideas")
        End Select

        Console.ReadLine()

    End Sub

End Module
