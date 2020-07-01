Attribute VB_Name = "Module1"
Sub VBAStocks():

    'Set worksheet parameters
    For Each ws In Worksheets

    'Count rows, designate to count to the last row of each worksheet
    Last_Row = ws.Cells(Rows.Count, "A").End(xlUp).Row

    'Designate header columns
    ws.Range("I1").Value = "Ticker"
    ws.Range("O1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Volume"
    ws.Range("P1").Value = "Value"
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"

    'Define start of counting rows.
    j = 0
    Startrow = 2
    Yearly_change = 0


    'Make a loop to summerize data
    For i = 2 To Last_Row
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
    
      If Total_Stock_Volume = 0 Then
        ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
        ws.Range("J" & 2 + j).Value = "%" & 0
        ws.Range("L" & 2 + j).Value = 0
       
       Else
        'Find first non zero start value
        If ws.Cells(Startrow, 3) = 0 Then
           For find_Value = Startrow To 1
                If ws.Cells(find_Value, 3).Value <> 0 Then
                    Start = find_Value
                    Exit For
                 End If
            Next find_Value
        End If
        
        'Calculate Yearly Change
        Yearly_change = (ws.Cells(i, 6) - ws.Cells(Startrow, 3))
        Percent_Change = Round((Yearly_change / ws.Cells(Startrow, 3) * 100), 2)
        
        'Start of next stock ticker
        Start = i + 1
        
        'put results in designated location
        ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
        ws.Range("J" & 2 + j).Value = Round(Yearly_change, 2)
        ws.Range("K" & 2 + j).Value = "%" & Percent_Change
        ws.Range("L" & 2 + j).Value = Total_Stock_Volume
        
        'Format cells, Color Green for % > 0, Red for % < 0
        Select Case Yearly_change
            Case Is > 0
                ws.Range("J" & 2 + j).Interior.ColorIndex = 4
            Case Is < 0
                ws.Range("J" & 2 + j).Interior.ColorIndex = 3
            Case Else
             ws.Range("J" & 2 + j).Interior.ColorIndex = 0
        End Select
      
      End If
        
        'reset variables back to 0 for next stock ticker
       Total_Stock_Volume = 0
       Yearly_change = 0
       j = j + 1
        
        Else
        Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
    
    End If
        
'Part II - Challenge

Next i

Max = 0
Min = 0

'Create a loop for the Greatest % Increase
For i = 2 To Last_Row
    If ws.Cells(i, 11).Value > Max Then
    Max = ws.Cells(i, 11).Value
    Ticker = ws.Cells(i, 9).Value
    ws.Range("O2") = Ticker
    ws.Range("P2") = Max * 100 & "%"
End If

Next i

'Create a loop for the Greatest % Decrease
For i = 2 To Last_Row
    If ws.Cells(i, 11).Value < Min Then
    Min = ws.Cells(i, 11).Value
    Ticker = ws.Cells(i, 9).Value
    ws.Range("O3") = Ticker
    ws.Range("P3") = Min * 100 & "%"
End If

Next i

'Make a loop for the Greatest Total Volume

For i = 2 To Last_Row
    If ws.Cells(i, 12).Value > Max Then
    Max = ws.Cells(i, 12).Value
    Ticker = ws.Cells(i, 9).Value
    ws.Range("O4") = Ticker
    ws.Range("P4") = Max
    
    End If
    
    Next i
    
    Next ws
        

End Sub

