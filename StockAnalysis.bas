Sub StockAnalysis()

' Loop through all sheets
For Each WS In Worksheets

' Insert code here

' Determine the Last Row
LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row

' Make sure all cells are sorted before data retrieval (not necessary)
' Disabled      With WS.Range("A1:G" & LastRow)
' Disabled      .Cells.Sort Key1:=.Columns("A"), Order1:=xlAscending, _
' Disabled              Key2:=.Columns("B"), Order2:=xlAscending, _
' Disabled              Orientation:=xlTopToBottom, Header:=xlYes
' Disabled       End With



' Generate Summary_Table
' Create table elements
WS.Range("I1").Value = "Ticker"
WS.Range("J1").Value = "Yearly Change"
WS.Range("K1").Value = "Percent Change"
WS.Range("L1").Value = "Total Stock Volume"


' Set an initial variable for holding the stock ticker
Dim Stock_Ticker As String

' Set an initial variable for holding the volume per stock
Dim Stock_Volume As Double
Stock_Volume = 0

' Keep track of the location for each stock in the Summary_Table
Dim Summary_Table_Row As Integer
Summary_Table_Row = 1

' Check Open Price of the first stock
Dim Open_Price As Double
Open_Price = WS.Range("C2").Value
  
' Loop through all stocks
For I = 2 To LastRow

' Check if we are still within the same stock, if it is not...
    If WS.Range("A" & I).Value <> WS.Range("A" & I + 1).Value Then
      
        ' Add 1 to the Summary Table Row
        Summary_Table_Row = Summary_Table_Row + 1
      
        ' Set the Stock Ticker
        Stock_Ticker = WS.Range("A" & I).Value

        ' Add to the Volume Total
        Stock_Volume = Stock_Volume + WS.Range("G" & I).Value

        'Check the Close Price
        Close_Price = WS.Range("F" & I).Value
      
        ' Calculate Yearly Change and Percent Change
        Yearly_Change = Close_Price - Open_Price
        Percent_Change = Yearly_Change / Open_Price
      
        ' Print the Stock Ticker in the Summary Table
        WS.Range("I" & Summary_Table_Row).Value = Stock_Ticker

        ' Print the Yearly Change and the Percent Change in the Summary Table
        WS.Range("J" & Summary_Table_Row).Value = Yearly_Change
        WS.Range("K" & Summary_Table_Row).Value = Percent_Change
      
        ' Print the Brand Amount to the Summary Table
        WS.Range("L" & Summary_Table_Row).Value = Stock_Volume
      
      
        ' Reset the Stock Volume
        Stock_Volume = 0
     
        ' Check Open Price for the next stock
        Open_Price = WS.Range("C" & I + 1).Value
    
        ' If the cell immediately following a row is the same brand...
    Else

        ' Add to the Stock Volume
        Stock_Volume = Stock_Volume + WS.Range("G" & I).Value

    End If

  Next I



' Generate Greatest_Table
' Create table elements
WS.Range("O2").Value = "Greatest % Increase"
WS.Range("O3").Value = "Greatest % Decrease"
WS.Range("O4").Value = "Greatest Total Volume"
WS.Range("P1").Value = "Ticker"
WS.Range("Q1").Value = "Value"

' Check Greatest Increase / Decrease / Volume
Greatest_Increase = Application.WorksheetFunction.Max(WS.Range("K2:K" & Summary_Table_Row))
Greatest_Decrease = Application.WorksheetFunction.Min(WS.Range("K2:K" & Summary_Table_Row))
Greatest_Volume = Application.WorksheetFunction.Max(WS.Range("L2:L" & Summary_Table_Row))

' Print Greatest Increase / Decrease / Volume
WS.Range("Q2").Value = Greatest_Increase
WS.Range("Q3").Value = Greatest_Decrease
WS.Range("Q4").Value = Greatest_Volume

' Check and Print Ticker for Greatest Increase / Decrease / Volume
For I = 2 To Summary_Table_Row

If WS.Range("K" & I).Value = WS.Range("Q2").Value Then
WS.Range("P2").Value = WS.Range("I" & I).Value
End If

If WS.Range("K" & I).Value = WS.Range("Q3").Value Then
WS.Range("P3").Value = WS.Range("I" & I).Value
End If

If WS.Range("L" & I).Value = WS.Range("Q4").Value Then
WS.Range("P4").Value = WS.Range("I" & I).Value
End If

Next I



'Formatting

'Yearly Change
For I = 2 To Summary_Table_Row

' Positive Green
If WS.Range("J" & I).Value >= 0 Then
WS.Range("J" & I).Interior.Color = 65280

' Negative Red
ElseIf WS.Range("J" & I).Value < 0 Then
WS.Range("J" & I).Interior.Color = 255

End If

Next I

' To prevent larger numbers from displaying in scientific (exponential) notation (not necessary)
' Disabled       WS.Range("L2:L" & Summary_Table_Row).NumberFormat = "0"
' Disabled       WS.Range("Q4").NumberFormat = "0"

' Display Percentage to 2 decimal place
WS.Range("K2:K" & Summary_Table_Row).NumberFormat = "0.00%"
WS.Range("Q2").NumberFormat = "0.00%"
WS.Range("Q3").NumberFormat = "0.00%"

'AutoFit columns (not necessary)
WS.Range("A:Q").Columns.AutoFit

' Check loop (not necessary)
' Disabled       MsgBox ("Analysis for Stocks in " + WS.Name + " Complete")

Next WS

'Fixes Complete
MsgBox ("Analysis Complete")

End Sub


