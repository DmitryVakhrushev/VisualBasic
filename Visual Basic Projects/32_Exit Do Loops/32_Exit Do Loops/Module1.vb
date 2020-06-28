Module Module1

    Sub Main()
        Dim num1 As Integer = 0
        Do Until num1 = 15
            If num1 = 5 Then Exit Do
            Console.WriteLine(num1)
            num1 += 1
        Loop

        Console.WriteLine()

        num1 = 0

        Do While num1 <= 15
            If num1 = 7 Then Exit Do
            Console.WriteLine(num1)
            num1 += 1
        Loop

        Console.ReadLine()
    End Sub

End Module
