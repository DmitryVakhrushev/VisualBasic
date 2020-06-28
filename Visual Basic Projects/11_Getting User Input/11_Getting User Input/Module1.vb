Module Module1

    Sub Main()
        Dim myName As String = Nothing 'in case a user will not input it
        Dim myAge As Integer = Nothing
        Dim mySalary As Double = Nothing

        Console.WriteLine("What is your name?")
        myName = Console.ReadLine()

        Console.WriteLine("What is your age?")
        myAge = Console.ReadLine()

        Console.WriteLine("What is your salary?")
        mySalary = Console.ReadLine()

        Console.Write("Your name is " & myName)
        Console.Write("   Your age is " & myAge)
        Console.Write("   Your salary is " & mySalary)

        Console.ReadLine()

    End Sub

End Module
