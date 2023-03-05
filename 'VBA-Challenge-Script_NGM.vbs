' VBA-Challenge Final Script

' Create a script that loops through all the stocks for one year and outputs the following information:
  ' The ticker symbol.
  ' Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
  ' The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
  ' The total stock volume of the stock.
  
Sub wallstreet()

' Across Worksheets
Dim ws As Long
Dim shtCount As Long

shtCount = Sheets.Count

For ws = 1 To shtCount

' How many rows are there? Use LastRow
lastrow = Sheets(ws).Cells(Rows.Count, 1).End(xlUp).Row
    
   'enter names into appropriate cell headers
    Sheets(ws).Range("I1") = "Ticker"
    Sheets(ws).Range("J1") = "Yearly Change"
    Sheets(ws).Range("K1") = "Percent Change"
    Sheets(ws).Range("L1") = "Total Stock Volume"

  ' Set an initial variable for holding the ticker name
  Dim ticker_Symbol As String

  ' Set an initial variable for holding the total
  Volume = 0

  ' Keep track of the location for each ticker symbols in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  ' Loop through all tickers
  For i = 2 To lastrow
    
    ' Check if we are still within the same ticker symbols, if it is not...
    If Sheets(ws).Cells(i + 1, 1).Value <> Sheets(ws).Cells(i, 1).Value Then
        ticker_Symbol = Sheets(ws).Cells(i, 1).Value
        Volume = Volume + Sheets(ws).Cells(i, 7).Value
        
        ' Print the ticker symbols, vol in the Summary Table
          Sheets(ws).Range("I" & Summary_Table_Row).Value = ticker_Symbol
          Sheets(ws).Range("L" & Summary_Table_Row).Value = Volume

        ' Reset the Vol
          Volume = 0
          
                ' Change in stock and percentages needs to be calculated?
                Stock_Closed = Sheets(ws).Cells(i, 6).Value
                     
                  'Stock loop?
                  If Stock_Open = 0 Then
                  Yearly_Change = 0
                  Percent_Change = 0
                  
                  Else:
                      Yearly_Change = Stock_Closed - Stock_Open
                      Percent_Change = Yearly_Change / Stock_Open
                    End If
                    

          Sheets(ws).Range("J" & Summary_Table_Row).Value = Yearly_Change
          
          Sheets(ws).Range("K" & Summary_Table_Row).Value = Percent_Change
          Sheets(ws).Range("K" & Summary_Table_Row).Style = "Percent"
          Sheets(ws).Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
      
      
        ' Make sure to use conditional formatting that will highlight positive change in green and negative change in red.
          ' red = Interior.ColorIndex = 3
          ' green = Interior.ColorIndex = 4
          
            If Sheets(ws).Range("J" & Summary_Table_Row).Value < 0 Then
            Sheets(ws).Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            ElseIf Sheets(ws).Range("J" & Summary_Table_Row).Value > 0 Then
            Sheets(ws).Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            End If
      

     ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
    ElseIf Sheets(ws).Cells(i - 1, 1).Value <> Sheets(ws).Cells(i, 1) Then
        Stock_Open = Sheets(ws).Cells(i, 3)
         
    ' If the cell immediately following a row is the same Ticker..
    Else
    ' Add to the Vol Total
      Volume = Volume + Sheets(ws).Cells(i, 7).Value
      
    End If

  Next i



' Functionality % changes (col 11, K)
 ' Set an initial variable for holding
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestVolume As Double

  ' Set names of col and rows
    Sheets(ws).Range("P1").Value = "Ticker"
    Sheets(ws).Range("Q1").Value = "Value"
    Sheets(ws).Range("O2").Value = "Greatest % Increase"
    Sheets(ws).Range("O3").Value = "Greatest % Decrease"
    Sheets(ws).Range("O4").Value = "Greatest Total Volume"

    ' Reset to 0
    GreatestIncrease = 0
    GreatestDecrease = 0
    GreatestVolume = 0

  ' Loop through all percent changes for ticker
  For a = 2 To lastrow
    If Sheets(ws).Cells(a, 11).Value > GreatestIncrease Then
        GreatestIncrease = Sheets(ws).Cells(a, 11).Value
        Sheets(ws).Range("Q2").Value = GreatestIncrease
        Sheets(ws).Range("Q2").Style = "Percent"
        Sheets(ws).Range("Q2").NumberFormat = "0.00%"
        Sheets(ws).Range("P2").Value = Sheets(ws).Cells(a, 9).Value
          End If
    Next a

    For b = 2 To lastrow
    If Sheets(ws).Cells(b, 11).Value < GreatestDecrease Then
        GreatestDecrease = Sheets(ws).Cells(b, 11).Value
        Sheets(ws).Range("Q3").Value = GreatestDecrease
        Sheets(ws).Range("Q3").Style = "Percent"
        Sheets(ws).Range("Q3").NumberFormat = "0.00%"
        Sheets(ws).Range("P3").Value = Sheets(ws).Cells(b, 9).Value
             End If
    Next b

    For c = 2 To lastrow
    If Sheets(ws).Cells(c, 12).Value > GreatestVolume Then
        GreatestVolume = Sheets(ws).Cells(c, 12).Value
        Sheets(ws).Range("Q4").Value = GreatestVolume
        Sheets(ws).Range("P4").Value = Sheets(ws).Cells(c, 9).Value
            End If
    Next c
    

' Loop across WS
 Next ws

      End Sub


