Module Module1

    Sub Main()
        Console.WriteLine("Normal For Loop")
        For num1 = 1 To 20
            Console.WriteLine(num1)
        Next

        Console.WriteLine()
        Console.WriteLine("Step 5 Loop")

        For num2 = 1 To 20 Step 5
            Console.WriteLine(num2)
        Next

        Console.WriteLine()
        Console.WriteLine("Step -4 Loop")

        For num3 = 20 To 1 Step -4
            Console.WriteLine(num3)
        Next

        Console.ReadLine()
    End Sub

End Module
