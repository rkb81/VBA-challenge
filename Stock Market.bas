Attribute VB_Name = "Module1"
Sub Stock_Market()

Dim Ticker_Name As String
Dim Summary_Data As Integer
Dim Start As Long
  
Dim ws As Worksheet
Dim Last_Row As Long
Dim i As Long

Dim Total_Stock_Volume As Double
Dim Opening_Price As Double
Dim Closing_Price As Double
Dim Yearly_Change As Double
Dim Percent_Change As Double

Dim Max_Val As Double
Dim Max_Ticker As String
Dim Min_Val As Double
Dim Min_Ticker As String
Dim Greatest_Volume As Double
Dim Greatest_Volume_Ticker As String

For Each ws In Worksheets

  Start = 2
  Summary_Data = 2
  Yearly_Change = 0

  'Create Headers
  ws.Cells(1, 9).Value = "Ticker"
  ws.Cells(1, 10).Value = "Yearly Change"
  ws.Cells(1, 11).Value = "Percent Change"
  ws.Cells(1, 12).Value = "Total Stock Volume"

  Total_Stock_Volume = 0

  Last_Row = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

  For i = 2 To Last_Row
  
    Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value

    'Calculate values for each ticker
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        Ticker_Name = ws.Cells(i, 1).Value
        
        Opening_Price = ws.Cells(Start, 3).Value
        Closing_Price = ws.Cells(i, 6).Value
    
        Start = i + 1
      
        'Calculate and Yearly_Change and Percentage_Change
        Yearly_Change = Closing_Price - Opening_Price
        Percent_Change = Yearly_Change / Opening_Price

        'Format and display values and ticker names
        ws.Range("J" & Summary_Data).NumberFormat = "0.00"
        ws.Range("K" & Summary_Data).NumberFormat = "0.00%"
        ws.Range("I" & Summary_Data).Value = Ticker_Name
        ws.Range("J" & Summary_Data).Value = Yearly_Change
        ws.Range("K" & Summary_Data).Value = Percent_Change
        ws.Range("L" & Summary_Data).Value = Total_Stock_Volume

        'Color formatting of red and green
        If ws.Range("J" & Summary_Data).Value > 0 Then
            ws.Range("J" & Summary_Data).Interior.ColorIndex = 4
        ElseIf ws.Range("J" & Summary_Data).Value < 0 Then
            ws.Range("J" & Summary_Data).Interior.ColorIndex = 3
        End If

        If ws.Range("K" & Summary_Data).Value > 0 Then
            ws.Range("K" & Summary_Data).Interior.ColorIndex = 4
        ElseIf ws.Range("K" & Summary_Data).Value < 0 Then
            ws.Range("K" & Summary_Data).Interior.ColorIndex = 3
        End If

        Total_Stock_Volume = 0
        Summary_Data = Summary_Data + 1

    End If

  Next i
  
  'Create Headings
  ws.Cells(2, 15).Value = "Greatest % Increase"
  ws.Cells(3, 15).Value = "Greatest % Decrease"
  ws.Cells(4, 15).Value = "Greatest Total Volume"
  ws.Cells(1, 16).Value = "Ticker"
  ws.Cells(1, 17).Value = "Value"

  'Define values and index ticker names
  Max_Val = Application.WorksheetFunction.Max(ws.Range("K2:K" & Last_Row))
  Min_Val = Application.WorksheetFunction.Min(ws.Range("K2:K" & Last_Row))
  Greatest_Volume = Application.WorksheetFunction.Max(ws.Range("L2:L" & Last_Row))
  Max_Ticker = Application.WorksheetFunction.Match(Max_Val, ws.Range("K2:K" & Last_Row), 0)
  Min_Ticker = Application.WorksheetFunction.Match(Min_Val, ws.Range("K2:K" & Last_Row), 0)
  Greatest_Volume_Ticker = Application.WorksheetFunction.Match(Greatest_Volume, ws.Range("L2:L" & Last_Row), 0)
    
  'Format and display values and ticker names
  ws.Cells(2, 17).NumberFormat = "0.00%"
  ws.Cells(3, 17).NumberFormat = "0.00%"
  ws.Cells(2, 17).Value = Max_Val
  ws.Cells(3, 17).Value = Min_Val
  ws.Cells(4, 17).Value = Greatest_Volume
  ws.Cells(2, 16).Value = ws.Cells(Max_Ticker + 1, 9)
  ws.Cells(3, 16).Value = ws.Cells(Min_Ticker + 1, 9)
  ws.Cells(4, 16).Value = ws.Cells(Greatest_Volume_Ticker + 1, 9)

Next ws

End Sub


