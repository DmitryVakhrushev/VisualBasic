Imports System

Module Program
    Sub Main(args As String())
        Console.WriteLine("Project: Split and Join")

        Dim myStr As String = "Airlines | 2 | Hello"
        Console.WriteLine(myStr) 'Prints original string
        Console.WriteLine("----------------------------")

        '################## SPLIT ##################
        '
        '
        Dim MyStrArray As String() = myStr.Split(" | ") 'Split creates an Array
        Dim s1 As String = MyStrArray(0)
        Dim s2 As String = MyStrArray(1)
        Console.WriteLine("1st Array Value: " + s1)
        Console.WriteLine("2nd Array Value: " + s2)
        Console.WriteLine("----------------------------")

        '################## JOIN ##################
        ' https//www.dotnetperls.com/join-vbnet
        ' Join requires a delimiter character and a collection of strings. It places this separator in between the strings in the result.

        Dim s3 As String = String.Join(",", MyStrArray)
        Console.WriteLine("'String.Join' returns the string: " + s3)
        Console.WriteLine("----------------------------")

        ' Example #2 with the Three-element array
        Dim array(2) As String
        array(0) = "Dog"
        array(1) = "Cat"
        array(2) = "Python"

        ' Join array.
        Dim result As String = String.Join(",", array)

        ' Display result.
        Console.WriteLine(result)
        Console.WriteLine("----------------------------")

        Console.ReadKey(True)
    End Sub
End Module
