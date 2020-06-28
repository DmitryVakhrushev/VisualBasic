Module Module1

    Sub Main()
        Console.Write("Enter your first number: ")
        Dim num1 As Double = Console.ReadLine()
        Console.WriteLine("First number: " & num1)

        Console.Write("Enter your second number: ")
        Dim num2 As Double = Console.ReadLine()
        Console.WriteLine("Second number: " & num2)

        Console.WriteLine()
        Console.WriteLine("First number * Second number = " & num1 * num2)
        Console.ReadLine()

    End Sub

End Module
