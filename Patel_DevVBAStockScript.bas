Attribute VB_Name = "Module1"
Sub tickerdata()

    For Each ws In Worksheets
    
        Dim ticker As String
        Dim volume As Double
        Dim ticker_row As Integer
        Dim openprice As Double
        Dim closeprice As Double
        Dim yearlychange As Double
        Dim percentchange As Double
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        volume = 0
        
        ticker_row = 2
        
        openprice = ws.Cells(2, 3).Value
        
        endlist = ws.Cells(Rows.Count, 1).End(xlUp).Row
       
        For i = 2 To endlist
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                ticker = ws.Cells(i, 1).Value
                
                volume = volume + ws.Cells(i, 7).Value
                
                ws.Range("I" & ticker_row).Value = ticker
                
                ws.Range("L" & ticker_row).Value = volume
                
                closeprice = ws.Cells(i, 6).Value
                
                yearlychange = (closeprice - openprice)
                
                ws.Range("J" & ticker_row).Value = yearlychange
                
                 If openprice = 0 Then
                 
                    percentchange = 0
                    
                 Else
                 
                    percentchange = yearlychange / openprice
                    
                 End If
                 
                 
                ws.Range("K" & ticker_row).Value = percentchange
                
                ticker_row = ticker_row + 1
                
                volume = 0
                
                openprice = ws.Cells(i + 1, 3)
            
            Else
            
                volume = volume + ws.Cells(i, 7).Value
                
                
            End If
            
        Next i
        
        percent_end = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        For i = 2 To percent_end
        
            If ws.Cells(i, 10).Value > 0 Then
                
                    ws.Cells(i, 10).Interior.ColorIndex = 10
                    
            Else
                
                    ws.Cells(i, 10).Interior.ColorIndex = 3
                
            End If
        
        Next i
    
    Next ws

End Sub
