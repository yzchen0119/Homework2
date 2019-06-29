Attribute VB_Name = "Module21"
Option Explicit

Sub stock_analysis()


Dim ws As Worksheet
Dim i, j As Long
Dim firstrow_nextstock As Long
Dim percent_change As Double
Dim max_increase, min_increase, max_total_volume As Double
Dim total_volume, open_price, close_price As Double

    
    For Each ws In ActiveWorkbook.Worksheets
    
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        i = 2
        j = 2
        firstrow_nextstock = 2
        total_volume = 0


        Do While ws.Cells(i, 1).Value <> ""
        
        
        
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1) Then
                    
                    total_volume = total_volume + ws.Cells(i, 7).Value
                    
                    open_price = ws.Cells(firstrow_nextstock, 3).Value
                    close_price = ws.Cells(i, 6).Value
                    
                        If open_price <> 0 Then
                    
                             percent_change = (close_price - open_price) / open_price
                             ws.Cells(j, 11).Value = FormatPercent(percent_change, 2)
                        
                        End If
                        
                            
                    
                    ws.Cells(j, 9).Value = ws.Cells(i, 1).Value
                    
                    ws.Cells(j, 10).Value = close_price - open_price
                    
                        If ws.Cells(j, 10).Value > 0 Then
                             ws.Cells(j, 10).Interior.ColorIndex = 4
                             ElseIf ws.Cells(j, 10).Value < 0 Then
                                 ws.Cells(j, 10).Interior.ColorIndex = 3
                        End If
                        
                   
                    
                    ws.Cells(j, 12).Value = total_volume
                    
                    j = j + 1
                    
                    total_volume = 0
                   
                    firstrow_nextstock = i + 1
                    
                Else
                    total_volume = total_volume + ws.Cells(i, 7).Value
                    
                            
                    
                    
            End If
        
        i = i + 1
        
        Loop
      
      
            
        max_increase = ws.Cells(2, 11).Value
        min_increase = ws.Cells(2, 11).Value
        max_total_volume = ws.Cells(2, 12).Value
        i = 2
      
        Do While ws.Cells(i, 9).Value <> ""
        
            
            If ws.Cells(i, 11).Value > max_increase Then
                
                max_increase = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(2, 17).Value = FormatPercent(max_increase, 2)
            
            End If
            
            
            If ws.Cells(i, 11).Value < min_increase Then
                
                min_increase = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(3, 17).Value = FormatPercent(min_increase, 2)
            
            End If
            
            If ws.Cells(i, 12).Value > max_total_volume Then
                
                max_total_volume = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(4, 17).Value = max_total_volume
            
            End If
         
         
         
        i = i + 1
        Loop
         

      
    
    ws.UsedRange.Columns.AutoFit
      
      
      
      
    Next ws
        
            
    
End Sub

