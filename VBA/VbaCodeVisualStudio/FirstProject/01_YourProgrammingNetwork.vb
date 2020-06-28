Option Explicit On

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' Debugging
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Print in "Immediate Window"

Printing from Application Code
Debug.Print "Salary = "; Salary

Printing Values of Properties
? BackColor
? Text1.Height
? Form1.BackColor
? Form1.Text1.Height

' "?" is a shortcut for "print"
?cells(2,1)
?range("A5")
print Cells(2,1) 


'Printing from Application Code
'The Print method sends output to the Immediate window whenever you include the Debug object prefix:

Debug.Print [items][;]

'For example, the following statement prints the value of Salary to the Immediate window every time it is executed:

Debug.Print "Salary = "; Salary

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++




'==============================================================================
'Excel 2010 VBA Tutorial 1 - Creating a Macro with Visual Basic For Applications
'==============================================================================
REM https://www.youtube.com/watch?v=ABXPb0qnKUY&index=1&list=PLS7iHfqXNVhK3yzd_4XS5k4zsvnu2mkJC

Sub Lesson001_MessageBox()

    REM "Option Explicit" forces to drfine variables
    'The program print out the "Hello"message
    MsgBox("Hello world!!!")

End Sub


'==============================================================================
'Excel 2010 VBA Tutorial 2 - Referencing Ranges
'==============================================================================
REM https://www.youtube.com/watch?v=DOV4r3n3sJc&index=2&list=PLS7iHfqXNVhK3yzd_4XS5k4zsvnu2mkJC

REM "Range" allows to reference a single cell or range

Sub Lesson002_ReferencingRanges()

    ThisWorkbook.Sheets("Sheet1").Range("A1").Value = "String in the cell A1"

    '---------------------------------------
    'Assign to E1 the value of A1
    ThisWorkbook.Sheets("Sheet1").Range("E1").Value = ThisWorkbook.Sheets("Sheet1").Range("A1").Value

    '---------------------------------------
    'Using a variable
    Dim x1 As String
    Dim string1 As String
    x1 = "B2"
    string1 = "Assign a value using a variable"
    ThisWorkbook.Sheets("Sheet1").Range(x1).Value = string1

    '---------------------------------------
    'Assign a value to the range
    ThisWorkbook.Sheets("Sheet1").Range("C1:D10").Value = 17

    '---------------------------------------
    'Add a value on ANOTHER sheet
    ThisWorkbook.Sheets("ExtraSheet").Range("B4").Value = "Hello Kitty"

End Sub


'==============================================================================
'Excel 2010 VBA Tutorial 3 - Referencing with Cells
'==============================================================================
REM https://www.youtube.com/watch?v=Dh9yQRSXp9o&index=3&list=PLS7iHfqXNVhK3yzd_4XS5k4zsvnu2mkJC

Sub Lesson003_ReferencingWithCells()

    'Referencing with cells
    'Cells(r,c): where r - row, c - column

    ThisWorkbook.Sheets("Sheet1").Cells(5, 1).Value = 27

    x = 2 'using a variable for row
    ThisWorkbook.Sheets("Sheet1").Cells(x, 1).Value = 6

    'Adding value to the same cell
    ThisWorkbook.Sheets("Sheet1").Cells(6, 1).Value = ThisWorkbook.Sheets("Sheet1").Cells(6, 1).Value + 1

End Sub


'==============================================================================
'Excel 2010 VBA Tutorial 4 - Hiding Rows and Columns
'==============================================================================
REM https://www.youtube.com/watch?v=GMtM9SxXjjg&list=PLS7iHfqXNVhK3yzd_4XS5k4zsvnu2mkJC&index=4

Sub Lesson004_HideShowRowsColumns()

    Worksheets("Sheet1").Activate()
    Rows("2:5").Hidden = True
    Columns("B:C").Hidden = True


    Worksheets("Sheet1").Activate()
    'Rows("2:3").Hidden = False
    'Columns("B:C").Hidden = False
    Rows.Hidden = False 'unhide all rows
    Columns.Hidden = False 'unhide all columns

End Sub


'==============================================================================
'Excel 2010 VBA Tutorial 5 - Referencing Selections
'==============================================================================
REM https://www.youtube.com/watch?v=GMtM9SxXjjg&list=PLS7iHfqXNVhK3yzd_4XS5k4zsvnu2mkJC&index=4

Sub Lesson05_FillSelection()

    'Make operations with selections
    '(1) Select a range on the sheet, (2) Run the macro

    Selection.Value = "Filled"
    Selection.Interior.Color = RGB(0, 50, 0)
    Selection.Font.Color = RGB(255, 255, 255)
    Selection.Font.Bold = True

End Sub


'==============================================================================
'Excel 2010 VBA Tutorial 6 - Variables
'==============================================================================
REM https://www.youtube.com/watch?v=7fhVnl3AaTc&index=6&list=PLS7iHfqXNVhK3yzd_4XS5k4zsvnu2mkJC

Sub Lesson006_UserInput()

    'Declare variables
    Dim num1 As Integer
    Dim num2 As Integer
    Dim num3 As Integer

    num1 = InputBox("Input your FIRST number")
    num2 = InputBox("Input your SECOND number")

    num3 = num1 + num2

    'Use "&" to concatenate strings
    MsgBox("Sum of First and Second numbers = " & num3)

End Sub



'==============================================================================
'Excel 2010 VBA Tutorial 7 - Numerical Operations and Decimal Points
'==============================================================================
REM https://www.youtube.com/watch?v=vLCHqMp3p5E&index=7&list=PLS7iHfqXNVhK3yzd_4XS5k4zsvnu2mkJC

Sub Lesson007_UserInput2()

    'Declare variables
    Dim num1 As Double
    Dim num2 As Double
    Dim num3 As Double

    num1 = InputBox("Input your FIRST number")
    num2 = InputBox("Input your SECOND number")

    num3 = num1 / num2

    'Use "&" to concatenate strings
    MsgBox("Devision First Num / Second Num = " & num3)

End Sub


'==============================================================================
'Excel 2010 VBA Tutorial 8 - Strings
'==============================================================================
REM https://www.youtube.com/watch?v=2xycBGdh8kg&list=PLS7iHfqXNVhK3yzd_4XS5k4zsvnu2mkJC&index=8

Sub Lesson008_inputUser()

    Dim myString As String
    Dim myString2 As String

    myString = "This is a string"
    myString2 = "The SECOND string"

    myString = myString & "Middle String " & myString2 'concatenate strings

    ThisWorkbook.Sheets("Sheet1").Cells(1, 1).Value = myString

End Sub


'==============================================================================
'Excel 2010 VBA Tutorial 9 - String Methods
'==============================================================================
REM https://www.youtube.com/watch?v=2nff9b7I35U&list=PLS7iHfqXNVhK3yzd_4XS5k4zsvnu2mkJC&index=9

Sub Lesson009_StringMethods()

    Dim myString As String
    Dim lt As String
    Dim rt As String
    Dim wbName As String
    Dim strPosition As Long
    Dim leftSide As String
    Dim rightSide As String

    myString = "Hello Kitty!"

    'Left and Right methods
    lt = Left(myString, 5) 'left() method
    rt = Right(myString, 7) 'right() method

    'Get the workbook name
    wbName = ThisWorkbook.Name
    wbName = Left(wbName, Len(wbName) - 5) 'exclude ".xlsm"

    Debug.Print(wbName) 'print in immediate window

    'In-string methods
    strPosition = InStr(1, myString, "i") 'look for the first "i" in the string
    leftSide = Left(myString, strPosition - 1)
    rightSide = Right(myString, Len(myString) - strPosition)

    myString = Replace(myString, "Hello", "Bye")

    Debug.Print(leftSide)
    Debug.Print(rightSide)

    'Debug.Print (strPosition)

End Sub


'==============================================================================
'Excel 2010 VBA Tutorial 10 - Dates and Time
'==============================================================================
REM https://www.youtube.com/watch?v=Z9CqA-bMZA4&list=PLS7iHfqXNVhK3yzd_4XS5k4zsvnu2mkJC&index=10

Sub Lesson010_DatesAndTime()
    '------------------------------------------------------------------------------
    'Dates are represented as a number of days since 1900-01-01
    '--> January 3, 1900 is the 3rd day from 1900-01-01
    '--> May 13, 1977    is the 28,258th day from 1900-01-01

    Dim myDate As Date
    myDate = #5/25/2013# 'it's always in American format dd/mm/yy
    myDate = myDate + 1 'add 1 day

    myDate = myDate + (1 / 24) 'add 1 hour
    myDate = #5/25/2013 7:37:00 PM#

    'MsgBox (myDate)
    Cells(3, 1).Value = myDate
    '------------------------------------------------------------------------------
End Sub


'==============================================================================
'Excel 2010 VBA Tutorial 11 - Methods for working with Dates and Time
'==============================================================================
REM https://www.youtube.com/watch?v=2zeBkMNojqM&index=11&list=PLS7iHfqXNVhK3yzd_4XS5k4zsvnu2mkJC

Sub Lesson011_DatesAndTimeMethods()

    Dim myDate1 As Date
    Dim myDate2 As Date
    Dim dateDifference As Long

    '-------------------------------------------------
    'Date() and Now() methods
    myDate1 = #1/3/2015#  'Janury 03, 2015
    myDate1 = Date        'Current date
    myDate2 = #12/17/2014#

    dateDifference = myDate1 - myDate2
    MsgBox("It has been " & dateDifference & " days since December 17, 2104")

    myDate2 = Now() 'current date AND time

    '-------------------------------------------------
    'Combine a date from parts and add 1 to a month
    myDate1 = Date
    myDate1 = DateSerial(Year(myDate1), Month(myDate1) + 1, Day(myDate1))
    MsgBox(myDate1)

    '-------------------------------------------------
    'Find the first day of the week for provided date
    'For January 07, 2015 (Wednesday) it will return January 05, 2015 (Monday)
    myDate1 = Date - Weekday(Date, vbMonday) + 1
    MsgBox(myDate1)

End Sub


'==============================================================================
'Excel 2010 VBA Tutorial 12 - Booleans
'==============================================================================
REM https:https://www.youtube.com/watch?v=ioY7T-EQ9tU&list=PLS7iHfqXNVhK3yzd_4XS5k4zsvnu2mkJC&index=12

Sub Lesson012_Booleans()

    Dim myBool As Boolean
    myBool = False
    myBool = (5 = 6) 'evaluates to False as 5<>6
    myBool = ("Hi" = "Hi")

    MsgBox(myBool)

End Sub



'===========================================================================
Sub SelectSingleCellByPosition()

    'The most common way is to use "Range" keyword to select a single cell
    Range("A13").Select()

    'Select active cell and change it value
    ActiveCell.Value = 11

    'Using "Cells" allows to specify row and column numbers
    Cells(13, 2).Select()
    ActiveCell.Value = "The Lorax"

    'Unusual but useful [C13]
    'To insert dates we need to include them between two has tags #....#
    ' # mm / dd / yyyy #
    [c13].Select()
    ActiveCell.Value = #3/2/2012#

    'Select relative WorkBook
    Workbooks("Lesson_05_.xlsm").Activate()

    Range("A14").Value = 14
    Range("B14").Value = "Wreck it Ralph"
    Range("C14").Value = #11/2/2012#



End Sub

'select a rectangular range (option #1)
     Range("A1:C1").Select
     Selection.Interior.Color = rgbDarkBlue

     Range("A1:C1").Font.Color = rgbWhite

     [A1:C1].Font.Size = 14

'select a rectangular range (option #2)
     Range("A2", "C2").Interior.Color = rgbLightBlue

Range(Cells(2, 1), Cells(2, 3)).Font.Color = rgbDarkBlue

Sub ReferToRangeNames()

    Range("ID").Font.Italic = True
    [Title].Font.Color = rgbDarkBlue

End Sub

Sub SelectVariableCol()

    'Selecting a List from Top to Bottom
    Range("A3", Range("A1").End(xlDown)).Select()
    Selection.Font.Italic = True

    'Select a column similar to Ctrl + Arrow Down
    Range("B3", Range("B2").End(xlDown)).Select()
    Selection.Font.Color = rgbDarkBlue

    'Select a range similar to Ctrl + Arrow Down + Arrow Right
    Range("A3", Range("A1").End(xlDown).End(xlToRight)).Select()
    Selection.Interior.Color = rgbAliceBlue

End Sub

Sub CopyFilmListMethos2()

    'Copy and Paste the range
    Worksheets("Sheet1").Activate()
    Range("A1").CurrentRegion.Copy(Worksheets("Sheet3").Range("A1"))

    'Autofit columns width
    Worksheets("Sheet3").Activate()
    Columns("B:C").AutoFit()

    'Another way to autofit columns
    Range("A:B").EntireColumn.AutoFit()

End Sub


