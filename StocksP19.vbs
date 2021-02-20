Sub StocksPriceVarP19()
  
  'Add the Name per each Column of the Summary Table
  
  Range("I1").Value = "Ticker"

  Range("J1").Value = "Yearly Change"
  
  Range("K1").Value = "Percentage Change"

  Range("L1").Value = "Total Stock Volume"
  
  
 ' Set an initial variable for holding the brand name
  Dim Ticker_Name As String

  ' Set an initial variable for holding the total per credit card brand
  Dim YRL_Ticker_Chg As Double
  Dim Perc_YRL_Ticker_Chg As Double
  Dim Total_Vol
  
  YRL_Ticker_Chg = 0
  Perc_YRL_Ticker_Chg = 0
  Total_Vol = 0

  ' Keep track of the location for each Ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  
  ' Counts the number of rows
  lastrow = Cells(Rows.Count, 1).End(xlUp).Row
 
  ' Loop through each row, Loop through all Tickers Values
    For i = 2 To lastrow
         
        If Cells(i, 3).Value <> 0 Then
          
     'Check if we are still within the same ticker, if it is not...
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      'Set the Ticker name
            Ticker_Name = Cells(i, 1).Value

      'These Functions are the operations over the variables for TickerCounter
            YRL_Ticker_Chg = YRL_Ticker_Chg + ((Cells(i, 6).Value) - (Cells(i, 3).Value))
      
      
            Perc_YRL_Ticker_Chg = Round(((((Cells(i, 6).Value) - (Cells(i, 3).Value)) / (Cells(i, 3).Value)) * 100), 2)
            

            Total_Vol = Total_Vol + (Cells(i, 7).Value)
      
      ' Print the Ticket Name in the Summary Table
      Range("I" & Summary_Table_Row).Value = Ticker_Name

      ' Print the Yearly Changes on Ticket Price in the Summary Table
      Range("J" & Summary_Table_Row).Value = YRL_Ticker_Chg
      
      ' Print the Percentage of Change on Ticket Price in the Summary Table
      Range("K" & Summary_Table_Row).Value = Perc_YRL_Ticker_Chg
      
      ' Print the Total Anual Volume of Transaction per Ticket in the Summary Table
      Range("L" & Summary_Table_Row).Value = Total_Vol
            

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the three variables
      
      YRL_Ticker_Chg = 0
      
      Perc_YRL_Ticker_Chg = 0
      
      Total_Vol = 0
    
    ' If the cell immediately following a row is the same Ticker..
    Else

      ' These Functions are the operations over the variables for TickerCounter
      
      YRL_Ticker_Chg = YRL_Ticker_Chg + ((Cells(i, 6).Value) - (Cells(i, 3).Value))
      
      Perc_YRL_Ticker_Chg = (((Cells(i, 6).Value) - (Cells(i, 3).Value)) / (Cells(i, 3).Value)) * 100

      Total_Vol = Total_Vol + (Cells(i, 7).Value)

    End If
   End If

Next i

     'Color variations -> Positive Green, Negative Red
     lastrowcolored = Cells(Rows.Count, 10).End(xlUp).Row
         
         For j = 2 To lastrowcolored
            
            If Cells(j, 10) > 0 Then
                Cells(j, 10).Interior.ColorIndex = 4 ' Green
            Else
                Cells(j, 10).Interior.ColorIndex = 3 ' Red
            End If
        
        Next j
  
End Sub


