Imports System
'Import Module1

Module Program
    Sub Main(args As String())
        Console.WriteLine("Project: IndexOf")
        'Console.Write("Press any key to exit... ")
        'Console.ReadKey(True)

        'https://www.dotnetperls.com/array-indexof-vbnet
        REM: "IndexOf" searches an array from its beginning. It returns the index Of the element With the specified value. LastIndexOf, meanwhile, searches from the end. Both methods return -1 if no match Is found.

        Dim arr(5) As Integer
        arr(0) = 1
        arr(1) = 3
        arr(2) = 5
        arr(3) = 7
        arr(4) = 8
        arr(5) = 5

        'Search array with IndexOf.
        Dim index1 As Integer = Array.IndexOf(arr, 3)
        Console.WriteLine(index1)

        'Search with LastIndexOf.
        Dim index2 As Integer = Array.LastIndexOf(arr, 5)
        Console.WriteLine(index2)

        Console.ReadKey(True)

    End Sub


End Module
