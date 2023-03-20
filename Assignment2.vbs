Attribute VB_Name = "Module1"
Sub Assignment2()

    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    
    For Each ws In Worksheets
    

        
        ' Create variables for place holders
        
        Dim Stock_Name As String  ' Ticker name
        Dim Stock_Total As Double ' Stock volume
        Dim Start_price As Double ' Starting price of stock at the begining of year
        Dim End_price As Double   ' Ending price of stock at the ending of year
        Dim Yearly_change As Double ' For yearly change calculation
        Dim Percent_change As Double 'For percentage change calculation
        Dim GrtIncStock As Double    ' To find stock with greatest increase
        Dim LstIncStock As Double    ' To find least increased stock
        Dim HighVolStock As Double   ' To find stock with highest volume
        Dim GISName As String        ' To find name of stock with graetest increase
        Dim LISName As String        ' To find name of greatest decreaed stock
        Dim HVSName As String        ' To find name of stock with highest volume
     
        ' Determine the lastrow
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
        ' Setting initial value of Stock Volume total
        Stock_Total = 0

        ' Keep track of the location for each Stock in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        ' Defining names of headers of each column to be calculated
        ws.Range("K1").Value = "Ticker"
        ws.Range("L1").Value = "Yearly Change"
        ws.Range("M1").Value = "Percent Change"
        ws.Range("N1").Value = "Total Stock Volume"

    ' Loop through all rows
        For i = 2 To LastRow
            Start_price = ws.Cells(2, 3).Value
    
            ' If the stock name changes
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                ' Set the Stock name
                Stock_Name = ws.Cells(i, 1).Value
                
                ' Find the price of stock at the end of year
                End_price = ws.Cells(i, 6).Value
                
                ' Calculation of Yearly change
                Yearly_change = End_price - Start_price
                
                ' Calculation of percentage change
                Percent_change = Yearly_change / Start_price
      
                ' Add to the Stock Total
                Stock_Total = Stock_Total + ws.Cells(i, 7).Value

                ' Print the Stock Name in the Summary Table
                ws.Range("K" & Summary_Table_Row).Value = Stock_Name
                
                ' Print the Yearly change in the Summary Table
                ws.Range("L" & Summary_Table_Row).Value = Yearly_change
                
                ' Conditional formatting based on yearly change in stock price
                    If Yearly_change > 0 Then
                        ws.Range("L" & Summary_Table_Row).Interior.ColorIndex = 4
                    Else
                        ws.Range("L" & Summary_Table_Row).Interior.ColorIndex = 3
                    End If
      
                ' Print the Percent change in the Summary Table
                ws.Range("M" & Summary_Table_Row).Value = FormatPercent(Percent_change, 2)
      
                ' Print the Stock Volume to the Summary Table
                ws.Range("N" & Summary_Table_Row).Value = Stock_Total

                ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
      
                ' Reset the stock Total
                Stock_Total = 0
                
                ' Reset the stock starting price
                Start_price = ws.Cells(i + 1, 3).Value
      
           

            ' If the cell immediately following a row is the same stock...
            Else

                ' Add to the Stock Total
                Stock_Total = Stock_Total + ws.Cells(i, 7).Value

            End If
    
        Next i
        
    'Bonus assignment
        ' Determine the lastrow in Summary table
        lrow = ws.Range("M1").End(xlDown).Row
              
  
        'Setting initial values to find stock with greatest % increase
        GrtIncStock = ws.Cells(2, 13).Value
        GISName = ws.Cells(2, 11).Value
        
        'Looping to compare next value and replace with higher value
        For i = 2 To lrow
    
    
            If GrtIncStock < ws.Cells(i, 13).Value Then
        
                GrtIncStock = ws.Cells(i, 13).Value
                GISName = ws.Cells(i, 11).Value
   
        
            End If
    
        Next i
          
        ' Print the stock value and Ticker name with greatest % increase
        ws.Range("R2").Value = GISName
        ws.Range("S2").Value = FormatPercent(GrtIncStock, 2)
  
        ws.Range("R1").Value = "Ticker"
        ws.Range("S1").Value = "Value"
        ws.Range("Q2").Value = "Greatest % Increase"
        
        'Setting initial values to find stock with least % increase
        LstIncStock = ws.Cells(2, 13).Value
        LISName = ws.Cells(2, 11).Value
    
        'Looping to compare next value and replace with lower value
        For i = 2 To lrow
    
    
            If LstIncStock > ws.Cells(i, 13).Value Then
        
                LstIncStock = Cells(i, 13).Value
                LISName = Cells(i, 11).Value
   
        
            End If
    
        Next i
  
        ' Print the stock value and Ticker name with least % increase
        ws.Range("R3").Value = LISName
        ws.Range("S3").Value = FormatPercent(LstIncStock, 2)
        ws.Range("Q3").Value = "Greatest % Decrease"
        
        'Setting initial values to find stock with greatest volume
        HighVolStock = ws.Cells(2, 14).Value
        HVSName = ws.Cells(2, 11).Value
        
        'Looping to compare next value and replace with higher value
        For i = 2 To lrow
           
    
            If HighVolStock < ws.Cells(i, 14).Value Then
        
                HighVolStock = ws.Cells(i, 14).Value
                HVSName = ws.Cells(i, 11).Value
   
        
            End If
    
        Next i
   
        ' Print the stock value and Ticker name with greatest volume
        ws.Range("R4").Value = HVSName
        ws.Range("S4").Value = HighVolStock
        ws.Range("Q4").Value = "Greatest Total Volume"
  
    Next ws

    
     
End Sub


