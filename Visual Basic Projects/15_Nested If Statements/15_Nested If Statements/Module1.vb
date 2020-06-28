Module Module1

    Sub Main()
        Dim Age As Integer = Nothing
        Dim hasInsurance As Boolean = Nothing
        Console.Write("Enter your age: ")
        Age = Console.ReadLine()

        Console.Write("Do you have insurance true/false: ")
        hasInsurance = Console.ReadLine()

        If Age >= 16 Then
            If hasInsurance = True Then
                Console.WriteLine("You can drive")
            Else
                Console.WriteLine("You'd not better stopped by cops")
            End If
        Else
            Console.WriteLine("You CANNOT drive")
        End If

        Console.ReadLine()

    End Sub

End Module
