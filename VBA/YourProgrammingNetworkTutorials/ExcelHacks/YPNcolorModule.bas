Attribute VB_Name = "YPNcolorModule"
'Created by Matthew Sands on behalf of www.yourprogrammingnetwork.co.uk - 2014.
'enjoy!

'Function to use in the cell which takes a range and returns the colorindex of the cell
Public Function ColorIndex(ByVal aRange As range)
    ColorIndex = aRange.Interior.ColorIndex
End Function


Public Sub colorCells()
Attribute colorCells.VB_Description = "colors the cell with colorindex based on their values"
Attribute colorCells.VB_ProcData.VB_Invoke_Func = "C\n14"
    With Application.Selection
        For Each c In .Cells
            c.Interior.ColorIndex = c.Value
        Next c
    End With
End Sub
