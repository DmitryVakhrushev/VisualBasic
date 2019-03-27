Imports System

Module Program
    Sub Main(args As String())
        Console.WriteLine("Hello World!")

        PrintDivider1()
        PrintDivider2("### My Second Devider ###")

        ' Call the sub that is in another module.
        Module1.Divider3() ' specify with the full path name
        Divider3() ' or use just the Sub name

        ' Wait till any key is pressed
        Console.ReadKey(True)

    End Sub

    Sub PrintDivider1()
        Console.WriteLine("---------------------------------")
    End Sub

    Sub PrintDivider2(ByVal str As String)
        Console.WriteLine(str)
    End Sub

End Module
