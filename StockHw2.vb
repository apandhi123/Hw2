Sub stockmarket()

  For Each ws In Worksheets
  
  ' Set an initial variable for holding the ticker name
    Dim Ticker_Name As String

  ' Set an initial variable for holding the total stock volume
    Dim Total_Volume As Double
    Total_Volume = 0
  
  ' Set the initial/opening price of the ticker
    Dim Initiaprice
    Initialprice = ws.Cells(2, 3).Value
  
  ' Set the initial variable for holding the Yearly change
    Dim Yearchange
  
  ' Set the initial variable for holding the Percent change
    Dim Percentchange
    
  ' Keep track of the location for each credit card brand in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
  
  ' Insert the Ticker name & all other headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Total Stock Volume"
    ws.Cells(1, 11).Value = "Yearly Change"
    ws.Cells(1, 12).Value = "Percent Change"
    ws.Cells(2, 15).Value = "Greatest % increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest total volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
       
     Dim LastRow
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
 
  ' Loop through all ticker names
    For I = 2 To LastRow
        
    ' Check if we are still within the same Ticker name, if it is not...
        If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then

      ' Set the ticker name
            Ticker_Name = ws.Cells(I, 1).Value

      ' Add to the stock volume
            Total_Volume = Total_Volume + ws.Cells(I, 7).Value

      ' Print the ticker name in the Summary Table
            ws.Range("I" & Summary_Table_Row).Value = Ticker_Name

      ' Print the total volume to the Summary Table
            ws.Range("J" & Summary_Table_Row).Value = Total_Volume

      ' Reset the stock volume
            Total_Volume = 0
      
      'Set the ending price
      
            Endingprice = ws.Cells(I, 6).Value
    
    ' Calculate the yearly change in stock price
    
            Yearchange = Endingprice - Initialprice
    
    
    ' Print the yearly change to the summary row table
  
            ws.Range("K" & Summary_Table_Row).Value = Yearchange
      
    ' Color the negative and positive changes
    
            If Yearchange > 0 Then
      
      ' Color it green
        
            ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
      
            Else
        ' Color it red
        
            ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
        
            End If
      
        
    ' Condition if dividing by 0
     
        If Initialprice = 0 Then
    
        Percentchange = 0
    
        Else
      
    ' Calculate the percent change in stock price
        Percentchange = ((Yearchange / Initialprice) * 100)
        
        End If
          
     ' Print the percent change to the summary row table
        ws.Range("L" & Summary_Table_Row).Value = Percentchange & "%"
        
      ' Bonus Question
      
        Dim Min
        Dim Max
        Dim Maxvol
      
        Min = ws.Cells(2, 12).Value
        Max = ws.Cells(2, 12).Value
        Maxvol = ws.Cells(2, 10).Value
      
        Dim Minname
        Dim Maxname
        Dim Totalname
      
        Minname = ws.Cells(2, 9).Value
        Maxname = ws.Cells(2, 9).Value
        Totalname = ws.Cells(2, 9).Value
        
      ' Print the greatest % decrease
      
      
        For J = 2 To Summary_Table_Row
            
            If ws.Cells(J, 12).Value < Min Then
            Min = ws.Cells(J, 12).Value
            Minname = ws.Cells(J, 9).Value
            ws.Cells(3, 17).Value = Round(Min * 100, 2)
            ws.Cells(3, 16).Value = Minname
            End If
        
        Next J
      
      ' Print the greatest % increase
      
        For K = 2 To Summary_Table_Row
            
            If ws.Cells(K, 12).Value > Max Then
            Max = ws.Cells(K, 12).Value
            Maxname = ws.Cells(K, 9).Value
            ws.Cells(2, 17).Value = Round(Max * 100, 2)
            ws.Cells(2, 16).Value = Maxname
            End If
        
        Next K
      
      ' Print the greatest volume
      
        For L = 2 To Summary_Table_Row
            
            If ws.Cells(L, 10).Value > Maxvol Then
            Maxvol = ws.Cells(L, 10).Value
            Totalname = ws.Cells(L, 9).Value
            ws.Cells(4, 16).Value = Totalname
            ws.Cells(4, 17).NumberFormat = "0"
            ws.Cells(4, 17).Value = Maxvol
            
            End If
        
        Next L

     ' Add one to the summary table row
      
        Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the initial price
    
        Initialprice = ws.Cells(I + 1, 3).Value
      
    ' If the cell immediately following a row is the same ticker...
    Else

      ' Add to the stock volume
      Total_Volume = Total_Volume + ws.Cells(I, 7).Value

    End If

    Next I
    
 
   MsgBox (ws.Name)

    Next ws

End Sub
