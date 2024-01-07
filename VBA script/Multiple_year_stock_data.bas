Attribute VB_Name = "Module1"
Dim Ws As Worksheet
Dim WsName As String
Dim i, col, num, LRowA, LRowI As Long
Dim TotalVol, PercntChange, GrtInc, GrtDec, GrtVol As Double




Sub Multiple_year_stock_data()


    'Last row in column A
    LRowA = Cells(Rows.Count, 1).End(xlUp).Row
    'MsgBox (lrow)
    
    For Each Ws In Worksheets
        
        
        WsName = Ws.Name
        'MsgBox (WsName)
        
        num = 2
        col = 2
        
        Ws.Cells(1, 9).Value = "Ticker"
        Ws.Cells(1, 10).Value = "Yearly Change"
        Ws.Cells(1, 11).Value = "Percent Change"
        Ws.Cells(1, 12).Value = "Total Stock Volume"
        Ws.Cells(1, 16).Value = "Ticker"
        Ws.Cells(1, 17).Value = "Value"
        Ws.Cells(2, 15).Value = "Greatest % Increase"
        Ws.Cells(3, 15).Value = "Greatest % Decrease"
        Ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        For i = 2 To LRowA
        
            If Ws.Cells(i + 1, 1).Value <> Ws.Cells(i, 1).Value Then
            
                Ws.Cells(num, 9).Value = Ws.Cells(i, 1).Value
                
                Ws.Cells(num, 10).Value = Ws.Cells(i, 6).Value - Ws.Cells(col, 3).Value
                    
                    'For Yearly Change calculations
                    If Ws.Cells(num, 10).Value < 0 Then
                    
                        Ws.Cells(num, 10).Interior.ColorIndex = 3
                        
                    Else
                    
                        Ws.Cells(num, 10).Interior.ColorIndex = 4
                        
                    End If
                    
                    'For Percnt Change calculations
                    If Ws.Cells(col, 3).Value <> 0 Then
                    
                        PercntChange = ((Ws.Cells(i, 6).Value - Ws.Cells(col, 3).Value) / Ws.Cells(col, 3).Value)
                    
                        Ws.Cells(num, 11).Value = Format(PercntChange, "Percent")
                    
                    Else
                    
                        Ws.Cells(num, 11).Value = Format(0, "Percent")
                    
                    End If
                    
                
                TotalVol = WorksheetFunction.Sum(Range(Ws.Cells(col, 7), Ws.Cells(i, 7)))
                
                Ws.Cells(num, 12).Value = TotalVol
                
                num = num + 1
                
                col = i + 1
                
            End If
            
        Next i
        
        'Last row in column I
        LRowI = Ws.Cells(Rows.Count, 9).End(xlUp).Row
        'MsgBox (LRowI)
        
        GrtIncr = Ws.Cells(2, 11).Value
        GrtDecr = Ws.Cells(2, 11).Value
        GrtVol = Ws.Cells(2, 12).Value
        
    
        For i = 2 To LRowI
        
            'Greatest % Increase
            If Ws.Cells(i, 11).Value > GrtIncr Then
            
                GrtIncr = Ws.Cells(i, 11).Value
                
                Ws.Cells(2, 16).Value = Ws.Cells(i, 9).Value
                
                Else
                
                GrtIncr = GrtIncr
                
            End If
            
            Ws.Cells(2, 17).Value = Format(GrtIncr, "Percent")
        
        
            'Greatest % Decrease
            If Ws.Cells(i, 11).Value < GrtDecr Then
            
                GrtDecr = Ws.Cells(i, 11).Value
                
                Ws.Cells(3, 16).Value = Ws.Cells(i, 9).Value
                
                Else
                
                GrtDecr = GrtDecr
                
            End If
            
            Ws.Cells(3, 17).Value = Format(GrtDecr, "Percent")
            
        
            'Greatest Total Volume
            If Ws.Cells(i, 12).Value > GrtVol Then
            
                GrtVol = Ws.Cells(i, 12).Value
                
                Ws.Cells(4, 16).Value = Ws.Cells(i, 9).Value
            Else
            
                GrtVol = GrtVol
                
            End If
            
            Ws.Cells(4, 17).Value = Format(GrtVol, "Scientific")
        
        Next i
        
        'Autofit columns in all worksheets
        Ws.Columns.AutoFit
        
    Next Ws

End Sub

