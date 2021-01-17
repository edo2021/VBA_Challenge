Attribute VB_Name = "Module1"

' Prepared By EDRIS GEMTESSA


Sub STOCK_MARCKET_ANALYSIS()

    
   
Dim WS As Worksheet

    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate
        ' assign last row  to check data from all rows
        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row

        ' create Header name for the column
        
        Cells(1, "I").Value = "Ticker"
        Cells(1, "J").Value = "Yearly Change"
        Cells(1, "K").Value = "Percent Change"
        Cells(1, "L").Value = "Total Stock Volume"
        
        'create Variable the Open_Price, Close_Price,Yearly_Change,Ticker_Name,Percent_Change
        
        Dim Open_Price As Double
        Dim Close_Price As Double
        Dim Yearly_Change As Double
        Dim Ticker_Name As String
        Dim Percent_Change As Double
        
           'create variable total volume and initializating the total volume
        
        Dim Volume As Double
        Volume = 0
        
         'create varible  summary row and column and initializating
         
        Dim Summary_Row As Double
         Summary_Row = 2
         
        Dim Column As Integer
        Column = 1
        Dim i As Long
        
        'Assign  Open Price
        Open_Price = Cells(2, Column + 2).Value
        
         ' Loop through all ticker
        
        For i = 2 To LastRow
         
            If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
            
                ' assign Ticker name
                
                Ticker_Name = Cells(i, Column).Value
                
                Cells(Summary_Row, Column + 8).Value = Ticker_Name
                
                ' Assign Close Price
                
                Close_Price = Cells(i, Column + 5).Value
                
                ' calculate Yearly Change and assign to the cell
                
                Yearly_Change = Close_Price - Open_Price
                Cells(Summary_Row, Column + 9).Value = Yearly_Change
                
                ' calculate Percent Change and assign to the cell
                
                If (Open_Price = 0 And Close_Price = 0) Then
                    Percent_Change = 0
                ElseIf (Open_Price = 0 And Close_Price <> 0) Then
                    Percent_Change = 1
                Else
                    Percent_Change = Yearly_Change / Open_Price
                    Cells(Summary_Row, Column + 10).Value = Percent_Change
                    Cells(Summary_Row, Column + 10).NumberFormat = "0.00%"
                End If
                
                ' calculate Total Volume and assign to cell
                
                Volume = Volume + Cells(i, Column + 6).Value
                Cells(Summary_Row, Column + 11).Value = Volume
                
                ' iterate by the summary row
                
                Summary_Row = Summary_Row + 1
                
                Open_Price = Cells(i + 1, Column + 2)
                ' reset the total Volume
                Volume = 0
         
            Else
                Volume = Volume + Cells(i, Column + 6).Value
            End If
        Next i
        
        ' check the Last Row of Yearly Change per WS
        Yearly_Change_LastRow = WS.Cells(Rows.Count, Column + 8).End(xlUp).Row
        ' Format the cell depending on the result yearly change
        For j = 2 To Yearly_Change_LastRow
            If (Cells(j, Column + 9).Value > 0 Or Cells(j, Column + 9).Value = 0) Then
                Cells(j, Column + 9).Interior.ColorIndex = 10
            ElseIf Cells(j, Column + 9).Value < 0 Then
                Cells(j, Column + 9).Interior.ColorIndex = 3
            End If
        Next j
        
        ' assign Greatest % Increase, % Decrease, and Total Volume to the cells
        
        Cells(2, Column + 14).Value = "Greatest % Increase"
        Cells(3, Column + 14).Value = "Greatest % Decrease"
        Cells(4, Column + 14).Value = "Greatest Total Volume"
        Cells(1, Column + 15).Value = "Ticker"
        Cells(1, Column + 16).Value = "Value"
        
        ' Look through each rows to find the greatest value and its associate ticker
        
        For Z = 2 To Yearly_Change_LastRow
            If Cells(Z, Column + 10).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & Yearly_Change_LastRow)) Then
                Cells(2, Column + 15).Value = Cells(Z, Column + 8).Value
                Cells(2, Column + 16).Value = Cells(Z, Column + 10).Value
                Cells(2, Column + 16).NumberFormat = "0.00%"
            ElseIf Cells(Z, Column + 10).Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & Yearly_Change_LastRow)) Then
                Cells(3, Column + 15).Value = Cells(Z, Column + 8).Value
                Cells(3, Column + 16).Value = Cells(Z, Column + 10).Value
                Cells(3, Column + 16).NumberFormat = "0.00%"
            ElseIf Cells(Z, Column + 11).Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & Yearly_Change_LastRow)) Then
                Cells(4, Column + 15).Value = Cells(Z, Column + 8).Value
                Cells(4, Column + 16).Value = Cells(Z, Column + 11).Value
            End If
        Next Z
        
    Next WS
        
End Sub

