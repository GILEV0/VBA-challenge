Attribute VB_Name = "Module1"
Sub VBA_Challenge2()
    
'Set worksheet'
Dim ws As Worksheet
    
'Declare all variables in worksheet'
Dim Ticker As String
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Total_Stock_Volume As Double
Dim year_open As Double
Dim year_close As Double
    
    
'Define a Summary table'
Dim Summary_Table_Row As Integer
    
'Loop through each worksheet'
For Each ws In Worksheets

    'Create column headers for Summary Table'
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    'assign integer for loop to start at'
    Summary_Table_Row = 2
    previous_i = 1
    Total_Stock_Volume = 0
    
    'Find the Lastrow of worksheet'
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        'For each Ticker find the yearly change, percent change, and total stock volume'
        For i = 2 To LastRow
    
            'Check if the Ticker changes or is not equal to the previous, then identify the Ticker'
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            Ticker = ws.Cells(i, 1).Value
        
            'Next iteration'
            previous_i = previous_i + 1
    
            'Identify the Year open & Year Close'
            year_open = ws.Cells(previous_i, 3).Value
            year_close = ws.Cells(i, 6).Value
    
            'Initiate a for loop to calculate the total Stock Volume
            For j = previous_i To i
    
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(j, 7).Value
    
            Next j
    
                
            If year_open = 0 Then
    
                Percent_Change = year_close
    
            Else
                Yearly_Change = year_close - year_open
    
                Percent_Change = Yearly_Change / year_open
    
            End If
                 
            'Insert the totals into the Summary table'
            ws.Cells(Summary_Table_Row, 9).Value = Ticker
            ws.Cells(Summary_Table_Row, 10).Value = Yearly_Change
            ws.Cells(Summary_Table_Row, 11).Value = Percent_Change
            
    
            'Format Percent Change to %'
            ws.Cells(Summary_Table_Row, 11).NumberFormat = "0.00%"
                
            ws.Cells(Summary_Table_Row, 12).Value = Total_Stock_Volume
            
    
            'Move to the next row of the Summary Table'
    
            Summary_Table_Row = Summary_Table_Row + 1
    
            'Set all variables back to 0'
            Total_Stock_Volume = 0
            Yearly_Change = 0
            Percent_Change = 0
    
            'Move i number to variable previous_i
            previous_i = i
    
        End If

    Next i
        
        
    jLastRow = ws.Cells(Rows.Count, 10).End(xlUp).Row
        For j = 2 To jLastRow
            
            'onditional formatting that will highlight positive change in green and negative change in red'
            If ws.Cells(j, 10) > 0 Then
                    
                ws.Cells(j, 10).Interior.ColorIndex = 4
                
            Else
                    
                ws.Cells(j, 10).Interior.ColorIndex = 3
            
            End If
        
        Next j
        
            
Next ws


End Sub


