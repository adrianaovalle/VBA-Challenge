Attribute VB_Name = "Module1"

Sub RunAll()
'This sub-routine runs the calculations for all the Worksheets
'Button created in one sheet to run this sub-routine in all sheets
Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call stocks
    Next
    Application.ScreenUpdating = True
Sheet1.Activate
End Sub

Sub ResetAll()
'This sub-routine resets the results for all the Worksheets
'Button created in one sheet to run this sub-routine for all sheets

Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call Reset
    Next
    Application.ScreenUpdating = True
Sheet1.Activate
End Sub




