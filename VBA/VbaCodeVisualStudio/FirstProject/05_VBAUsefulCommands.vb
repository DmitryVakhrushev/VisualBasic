Option Explicit On

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' Useful VBA Commands
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'To print in Immediate Window
Worksheets("Sheet2").Visible = xlVeryHidden
Worksheets("Sheet2").Visible = -1 'Visible
Worksheets("Sheet2").Visible = 0 'Hidden
Worksheets("Sheet2").Visible = 2 'Very Hidden


'Another example is to hide the contents of a cell by making its font color the same as its fill (background) color.
Range("A1").Font.Color = Range("A1").Interior.Color


For hiding the contents of a cell: Rather than making the text color the same as the fill color,
a custom number format like “;;;” will hide the contents of a cell without the need to adjust it at any point, even if font color and fill color are tweaked. Another bonus to this type of formatting is that cells containing data points formatted this way won’t show up on a chart either, so it’s possible to use this to hide points or even series names based upon conditions.


'In Immediate Window
'A multi-line macro in the immediate window by using the colon ":" as a line separator.
For Each sh In Worksheets : Debug.Print sh.Name : Next

'In a normal macro his code would look like the following:
For Each sh in Worksheets
Debug.Print sh.Name
Next

'------------------------------------------------------------------------------
'Unhide all the sheets in the workbook.
'------------------------------------------------------------------------------

'code sample #1
For Each sh In Sheets:sh.Visible=True:Next sh

'code sample #2
Sub UnhideAllSheets()
    Dim wsSheet As Worksheet
    For Each wsSheet In ActiveWorkbook.Worksheets
        wsSheet.Visible = xlSheetVisible
    Next wsSheet
End Sub


'------------------------------------------------------------------------------
'Removes the page break lines on the sheet
'------------------------------------------------------------------------------
ActiveSheet.DisplayPageBreaks = False


'------------------------------------------------------------------------------
'Length of the Array
'------------------------------------------------------------------------------
Dim myArray(9) As Integer
Dim i As Integer
    For i = 0 To UBound(myArray)
    myArray(i) = i * 10
    Next

'Option Base 0 'Un-comment for zero-based arrays
'Option Base 1 'Un-comment for one-based arrays
Sub ShowArrayLen()
    Dim abcarray() As Object
    ReDim abcarray(3) 'i.e. some run-time array size value

    MsgBox("Range: " & LBound(abcarray) & " to " & UBound(abcarray))
    MsgBox("Length: " & UBound(abcarray) - LBound(abcarray) + 1)
    MsgBox("Ubound: " & UBound(abcarray)) 'does not = length for zero-based arrays!
End Sub

Application.Workbooks.Count
Application.ActiveWorkbook.ActiveSheet.Name


'------------------------------------------------------------------------------
'Excel built-in functions (e.g. min, max, stdev)
'------------------------------------------------------------------------------
Dim myVar As Double

myVar = WorksheetFunction.Pi
'worksheetfunction.Sum(arg1, arg2, arg3)
Dim a As Integer, b As Integer, c As Integer, total As Integer
a = 2
b = 3
c = 4
total = worksheetFunction.sum(a,b,c)




