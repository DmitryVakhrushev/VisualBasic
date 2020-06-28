'------------------------------------------
' Loop through each cell in a range of cells when given a Range object
'------------------------------------------

Sub trimCleanValues()
    'Loop through each cell in a range of cells when given a Range object
    'http://stackoverflow.com/questions/3875415/loop-through-each-cell-in-a-range-of-cells-when-given-a-range-object

    Dim rCell As Range
    Dim rRng As Range

    Range("A2").Select()
    Range(Selection, Selection.End(xlDown)).Select()

    rRng = Selection
    'Set rRng = Sheet1.Range("A1:A6")

    For Each rCell In rRng.Cells

        'Debug.Print rCell.Address, rCell.Value

        rCell.Value = Application.WorksheetFunction.Clean(rCell.Value)
        rCell.Value = Application.WorksheetFunction.Trim(rCell.Value)


    Next rCell

End Sub



'------------------------------------------
' VBA Functions
'------------------------------------------
Public Function MultNum(myNum)
	x = myNum * 10
	MultNum = x 
End Function

' To get the result we have to assign a value to a variable that has the same name as function name (MultNum)
' in a cell print = and then MultNum(insert parameter), e.g. =MultNum(A1), =MultNum(17)



'------------------------------------------
' VBA Commands
'------------------------------------------
Sub PopularCommands()
	
	'Select relative worksheet, e.g. "Sheet1"
    Worksheets("Sheet1").Activate
	
	Worksheets.Add
	
    'Add values to a Range (cells)
    Range("A1").Value = "Created by" 												' string
	Range("B1").Value = Environ("UserName") 										' current user
    Range("B2").Value = Date 														' current date
	ThisWorkbook.Sheets("Sheet1").Range("A1").Value = "Hello World"
	ThisWorkbook.Sheets("Sheet1").Range("E1").Value = ThisWorkbook.Sheets("Sheet1").Range("A1").Value ' Reference the cell A1 into E1
	
	' Using variables
    x1 = "B2"
    string1 = "Assign a value using a variable"
    ThisWorkbook.Sheets("Sheet1").Range(x1).Value = string1
	
	ThisWorkbook.Sheets("Sheet1").Range("C1:D10").Value = 17						' Assign a value to the range
	
	
	'Add values to a Range (cells)
	'Cells(x,y) where x - row, y - column
	ThisWorkbook.Sheets("Sheet1").Cells(3, 1).Value = 19
	ThisWorkbook.Sheets("Sheet1").Cells(3, 1).Value = ThisWorkbook.Sheets("Sheet1").Cells(3, 1).Value + 1  ' Adding value to the same cell
	
	'--------------------------------------------------------------------
	'Format titles
    Range("A1:A3").Font.Color = vbBlue						' Font color
    Range("A1:A3").Interior.Color = rgbPaleTurquoise		' Cell color
	

    ' Message Box
    MsgBox ("Hello world!!!")
    
	'Assign a value to Date
	'To insert dates we need to include them between two has tags #....#, e.g. # mm / dd / yyyy #
    [c13].Select
    ActiveCell.Value = #3/2/2012#
	
	
    

End Sub
