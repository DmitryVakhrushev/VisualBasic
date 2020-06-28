Option Explicit On

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' Useful Shortcuts
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'F4 or Ctrl + Y --> repeats the last action

'Ctrl + ' --> copies the formula from the cell above or next to it

'Ctrl + ;                           --> Date 12/16/2014	 	
'Ctrl + Shift + ;                   --> Time 10:00 PM		
'Ctrl + ; <space> Ctrl + Shift + ;  --> 'Date and time 12/16/2014 22:00	 	

'Ctrl + Enter --> to replicate the same value to all the cells in the selected range


'Double click on the "Format Painter" to copy the format into multiple cells
'COUNTIF(A:A,A1) --> count duplicates

'Ctrl + Shify + F3 -- assign named ranges based on the column names from the table

'To represent fractions, e.g. 1/3, put in the cell a zero, space, and a fraction --> 0 1/3, 0 7/8


'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' Immediate Window
' (View > Immediate Window)
' Shortcut key (Ctrl + G).


'===========================================================================
'Weekly Excel Hacks - Episode 01
'===========================================================================

'https://www.youtube.com/watch?v=2-xbU7J57n4&list=PLS7iHfqXNVhJvdFXgsZlfZeAXptzzSwET&index=1

'Function to use in the cell which takes a range and returns the colorindex of the cell
Public Function ColorIndex(ByVal aRange As range)
    ColorIndex = aRange.Interior.ColorIndex
End Function


'===========================================================================
'Weekly Excel Hacks - Episode 03
'===========================================================================
REM https://www.youtube.com/watch?v=HOo_FzV2C-0&list=PLS7iHfqXNVhJvdFXgsZlfZeAXptzzSwET&index=3

Private Sub CommandButton1_Click()

    'If an error occurs go to the another point in the code
On Error Goto Error
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    'Application.Visible = False

    With ThisWorkbook.Sheets("Sheet1").Cells(2, 2)
        .Value = 1
        Do Until .Value >= 250000
            .Value = .Value + 1
        Loop
    End With

    'Always turn those options back ON
Error:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    'Application.Visible = True

End Sub


'===========================================================================
'Weekly Excel Hacks - Episode 04
'===========================================================================
REM https://www.youtube.com/watch?v=PJCgJCMsAAQ&index=4&list=PLS7iHfqXNVhJvdFXgsZlfZeAXptzzSwET

'Excel will disappear and only the form will be on the screen.


'assign this code to a button
Private Sub commandButton1_Click()
    UserForm1.Show()
End Sub

'--- User Form Subs ---'
Private Sub UserForm_Deactivate()
    Application.Visible = True
End Sub

Private Sub UserForm_Initialize()
    Application.Visible = False
End Sub

Private Sub UserForm_Terminate()
    Application.Visible = True
End Sub


'===========================================================================
'Weekly Excel Hacks - Episode 05
'===========================================================================
REM https://www.youtube.com/watch?v=_our7MgxyQM&index=5&list=PLS7iHfqXNVhJvdFXgsZlfZeAXptzzSwET
'To make the cursor the have a "circled" shape while being in the executing mode

Sub clickMe()

    Dim x As Double
    Application.Cursor = xlWait

    'if error occurs inside of the loop the Cursor turns to its normal state anyway
    On Error GoTo myError

    With ThisWorkbook.Sheets("Sheet1")
        For x = 1 To 1000 Step 0.01
            .Cells(1, 1).Value = .Cells(1, 1).Value + 1
            .Cells(2, 1).Value = x
        Next x
    End With

myError:
    Application.Cursor = xlDefault

End Sub

'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&

Sub test()

    Dim s1 As String
    Dim s2 As String
    s1 = "Hello Kitty"
    s2 = Left(s1, 3)
    MsgBox(s2)
End Sub