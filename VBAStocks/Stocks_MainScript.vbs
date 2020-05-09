Attribute VB_Name = "Module1"
Sub stocks()
'This sub-routine runs the calculations for one Worksheet. It includes the first challenge.
Dim i As Long
Dim l As Long
Dim j As Integer
Dim openbig As Double
Dim closeend As Double
Dim vol As Double
Dim row As Integer


'This initializes all variables for first data row
j = 2
vol = 0
l = Cells(Rows.Count, 1).End(xlUp).row
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
openbig = Cells(2, 3).Value

'Check if the dates are in ascending order, if not then message where it is wrong
For i = 2 To (l - 1)
    If Cells(i, 1).Value = Cells(i + 1, 1).Value Then
        If Cells(i, 2).Value > Cells(i + 1, 2).Value Then
            MsgBox ("Wrong Date in Row: " & i - 1)
            Stop
        End If
    End If
Next i


'Look into all rows of the worksheet, aggregate the volumes per stock and calculate the yearly change and the percentage of change of each stock.
'It assumes that the year opening happens in the first data point of the stock and the year end closing happens in the latest datapoint of the stock.
'It also assigns colors depending on the results of the yearly change

For i = 2 To (l - 1)
    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Or i = (l - 1) Then
        Cells(j, 9).Value = Cells(i, 1).Value
        vol = vol + Cells(i, 7).Value
        closeend = Cells(i, 6).Value
        If i = l - 1 Then
            vol = vol + Cells(i + 1, 7).Value
            closeend = Cells(i + 1, 6).Value
        End If
        Cells(j, 10).Value = closeend - openbig
        If openbig <> 0 Then
            Cells(j, 11).Value = (closeend - openbig) / openbig
        Else
            Cells(j, 11).Value = 0
        End If
        Cells(j, 11).NumberFormat = "0.00%"
        Cells(j, 12).Value = vol
        Cells(j, 12).NumberFormat = "#,##0"
        If (closeend - openbig) < 0 Then
            Cells(j, 10).Interior.Color = RGB(255, 0, 0)
        Else
            Cells(j, 10).Interior.Color = RGB(51, 204, 51)
        End If
        openbig = Cells(i + 1, 3).Value
        j = j + 1
        vol = 0
    Else
    vol = vol + Cells(i, 7).Value
    End If
Next i

'Calculate the summary table for % of greatest increase, Greatest decrease and greatest total volume

Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"

l = Cells(Rows.Count, 10).End(xlUp).row
Cells(2, 17).Value = WorksheetFunction.Max(Range(Cells(2, 11), Cells(l, 11)))
row = Application.WorksheetFunction.Match(Cells(2, 17).Value, Range(Cells(1, 11), Cells(l, 11)), 0)
Cells(2, 16).Value = Cells(row, 9)
Cells(3, 17).Value = WorksheetFunction.Min(Range(Cells(2, 11), Cells(l, 11)))
row = Application.WorksheetFunction.Match(Cells(3, 17).Value, Range(Cells(1, 11), Cells(l, 11)), 0)
Cells(3, 16).Value = Cells(row, 9)
Cells(4, 17).Value = WorksheetFunction.Max(Range(Cells(2, 12), Cells(l, 12)))
row = Application.WorksheetFunction.Match(Cells(4, 17).Value, Range(Cells(1, 12), Cells(l, 12)), 0)
Cells(4, 16).Value = Cells(row, 9)
Cells(4, 17).NumberFormat = "#,##0"
Cells(2, 17).NumberFormat = "0.00%"
Cells(3, 17).NumberFormat = "0.00%"
Columns("I:Q").EntireColumn.AutoFit
End Sub
Sub Reset()
Dim i As Long

'This sub-routine resets the results of one Worksheet

i = Cells(Rows.Count, 10).End(xlUp).row
Range("I1:Q" & i).ClearContents
Range("I1:Q" & i).ClearFormats


End Sub

