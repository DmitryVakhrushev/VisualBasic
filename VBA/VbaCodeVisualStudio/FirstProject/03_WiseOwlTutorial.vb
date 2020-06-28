Sub AssignToCell()

    'Cells(x,y) --> where x - row, y - column
    ThisWorkbook.Sheets("Sheet1").Cells(6, 1).Value = 6

    ' Adding value to the same cell
    ThisWorkbook.Sheets("Sheet1").Cells(6, 1).Value = ThisWorkbook.Sheets("Sheet1").Cells(6, 1).Value + 1

End Sub



Sub CreateAndLabelNewSheet()

    'Create a new worksheet
    REM you can ALSO start a comment with a key word "Rem"
    'Basic VBA sentence contains: thing.action or 'Object.Method()
    Worksheets.Add()

    'add titles to cells
    'Object.Property = Value
    Range("A1").Value = "Created by"
    Range("A2").Value = "Created on"
    Range("A3").Value = "Version"

    'add user values to cells
    Range("B1").Value = Environ("UserName")
    Range("B2").Value = Date
    Range("B3").Value = 1

    'format titles
    Range("A1:A3").Font.Color = vbBlue
    Range("A1:A3").Interior.Color = rgbPaleTurquoise

End Sub