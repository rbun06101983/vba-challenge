Attribute VB_Name = "Module1"
Sub Stock_RunThrough()
 'Work with study group and tutor
 'Setting Loop through data sheets
 Dim ws As Worksheet
 Dim starting_ws As Worksheet
 Set starting_ws = ActiveSheet 'remember which worksheet is active in the beginning
 For Each ws In ThisWorkbook.Worksheets
    ws.Activate
 'Dim WallStreet As Workbook
    'For Each WallStreet In ThisWorkbook.Worksheets
    'WallStreet.Activate
    'Declaring Variables
    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim Yearly_Change As Double
    Dim Ticker_Name As String
    Dim Percent_Change As Double
    Dim Volume As Double
    Dim Lastrow As Long
    Dim Row As Double
    Dim column As Integer
    'Setting Summary Data (i.e Ticker, Yearly Change, etc.)
    Cells(1, "I").Value = "Ticker"
    Cells(1, "J").Value = "Yearly Change"
    Cells(1, "K").Value = "Percent Change"
    Cells(1, "L").Value = "Total Stock Volume"
    'start Volume at 0
    Volume = 0
    'set output to second row / anchoring yourself in second row
    Row = 2
    column = 1
    Dim i As Long
    'Opening Price
    Open_Price = Cells(2, column + 2).Value
    Lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    'Looping through Tickers
    For i = 2 To Lastrow
    'are we in the same ticker, and if not...
    If Cells(i + 1, column).Value <> Cells(i, column).Value Then
     'Give Ticker Name
     Ticker_Name = Cells(i, column).Value
     Cells(Row, column + 8).Value = Ticker_Name
     'Set Closing Price
     Close_Price = Cells(i, column + 5).Value
     'Add the Yearly Change
     Yearly_Change = Close_Price - Open_Price
     Cells(Row, column + 9).Value = Yearly_Change
     'Add the Percent Change
     If (Open_Price = 0 And Close_Price = 0) Then
        Percent_Change = 0
     ElseIf (Open_Price = 0 And Close_Price <> 0) Then
        Percent_Change = 1
     Else
        Percent_Change = Yearly_Change / Open_Price
        Cells(Row, column + 10).Value = Percent_Change
        Cells(Row, column + 10).NumberFormat = "0.00%"
     End If
     'Add the Total Volume
     Volume = Volume + Cells(i, column + 6).Value
     Cells(Row, column + 11).Value = Volume
     'Add onto the summary table
     Row = Row + 1
     'Reset the Open Price
     Open_Price = Cells(i + 1, column + 2)
     'reset the Volume Total
     Volumn = 0
     'if the Cells have the same ticker
     Else
      Volume = Volume + Cells(i, column + 6).Value
     End If
     Next i
     'Determine the Last Row of the Yearly Change per WallStreet
     YCLastRow = Cells(Rows.Count, column + 8).End(xlUp).Row
     'Setting the Cell Colors
     For j = 2 To YCLastRow
        If (Cells(j, column + 9).Value > 0 Or Cells(j, column + 9).Value = 0) Then
            Cells(j, column + 9).Interior.ColorIndex = 10
        ElseIf Cells(j, column + 9).Value < 0 Then
            Cells(j, column + 9).Interior.ColorIndex = 3
        End If
    Next j
    'Set up Greatest% Incerease, Decrease and Total Volume calculation
    Cells(2, column + 14).Value = "Greatest % Increase"
    Cells(3, column + 14).Value = "Greatest % Deecrease"
    Cells(4, column + 14).Value = "Greatest Total Volume"
    Cells(1, column + 15).Value = "Ticker"
    Cells(1, column + 16).Value = "Value"
    'Look at each row and find greatest value and its Ticker
    For Z = 2 To YCLastRow
    If Cells(Z, column + 10).Value = Application.WorksheetFunction.Max(Range("K2:K" & YCLastRow)) Then
        Cells(2, column + 15).Value = Cells(Z, column + 8).Value
        Cells(2, column + 16).Value = Cells(Z, column + 10).Value
        Cells(2, column + 16).NumberFormat = "0.00%"
    ElseIf Cells(Z, column + 10).Value = Application.WorksheetFunction.Min(Range("K2:K" & YCLastRow)) Then
        Cells(3, column + 15).Value = Cells(Z, column + 8).Value
        Cells(3, column + 16).Value = Cells(Z, column + 10).Value
        Cells(3, column + 16).NumberFormat = "0.00%"
    ElseIf Cells(Z, column + 11).Value = Application.WorksheetFunction.Max(Range("L2:L" & YCLastRow)) Then
        Cells(4, column + 15).Value = Cells(Z, column + 8).Value
        Cells(4, column + 16).Value = Cells(Z, column + 11).Value
    End If
    Next Z
 ws.Cells(1, 1) = "<ticker>" 'this sets cell A1 of each sheet to "ticker"
Next
starting_ws.Activate 'activate the worksheet that was originally active
End Sub
