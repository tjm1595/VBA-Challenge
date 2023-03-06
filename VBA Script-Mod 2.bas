Attribute VB_Name = "Module1"
Sub Stocks()
    
    
    Dim ws As Worksheet
    
    For Each ws In Worksheets
    
    
    Dim Tickers As String
    Dim YearlyChange As Double
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim TotalVolume As Double
    TotalVolume = 0
    
    
    Dim SummaryTable As Integer
    SummaryTable = 2
    Dim LastRow As Long
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To LastRow
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
             
                Tickers = ws.Cells(i, 1).Value
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
                ws.Range("I" & SummaryTable).Value = Tickers
                ws.Range("L" & SummaryTable).Value = TotalVolume
                
              
                ClosePrice = ws.Cells(i, 6).Value
                ws.Range("K" & SummaryTable).Value = (ClosePrice - OpenPrice) / OpenPrice
                ws.Range("J" & SummaryTable).Value = ClosePrice - OpenPrice
                SummaryTable = SummaryTable + 1
                TotalVolume = 0
               
             ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                OpenPrice = ws.Cells(i, 3).Value
                
                
            Else
            
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
                
                
            
            End If
            
           
        
                        
            Next i

            
   Next ws
   

     
End Sub



