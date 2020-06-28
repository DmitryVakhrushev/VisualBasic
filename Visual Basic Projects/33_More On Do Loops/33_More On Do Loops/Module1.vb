Module Module1

    Sub Main()
        Dim num1 As Integer = 0
        Console.WriteLine("Normal Do Loop")
        Do Until num1 = 5
            Console.WriteLine(num1)
            num1 += 1
        Loop

        num1 = 10
        Console.WriteLine()
        Console.WriteLine("Run at least once")
        Do
            Console.WriteLine(num1)
        Loop Until num1 = 10

        Console.ReadLine()
    End Sub

End Module
