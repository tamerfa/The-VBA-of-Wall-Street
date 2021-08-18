Attribute VB_Name = "module1"
Sub wallstreet():

'Loop through all sheets
For Each ws In Worksheets
    'Create headers for summary table
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    'Create a variable to hold the location of the last row
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Create variables to hold ticker name, opening price, closing price, stock volume
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim stockVol As Double
    
    'Set initial value of stockVol
    stockVol = 0
    'Set initial value of openPrice
    openPrice = ws.Range("C2").Value
    
    'Create a variable to hold the row position in summary table
    Dim sRow As Integer
    'Set initial sRow value
    sRow = 2
    
    For i = 2 To lastrow
        If (ws.Range("A" & i + 1).Value <> ws.Range("A" & i).Value) Then
            ticker = ws.Range("A" & i).Value
            closePrice = ws.Range("F" & i).Value
            stockVol = stockVol + ws.Range("G" & i).Value
            '-------------------------
            'Fill in the summary table
            '-------------------------
            ws.Range("I" & sRow).Value = ticker
            ws.Range("J" & sRow).Value = (closePrice - openPrice)
            
                'Format Yearly Change cell to be filled with green for positive change and red for negative change
                If (ws.Range("J" & sRow).Value >= 0) Then
                    ws.Range("J" & sRow).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & sRow).Interior.ColorIndex = 3
                End If
            
            If openPrice = 0 Then
                ws.Range("K" & sRow).Value = 0
            Else
                ws.Range("K" & sRow).Value = ws.Range("J" & sRow) / openPrice
                
            End If
            
            'Format Percent Change cell as percentage
            ws.Range("K" & sRow).NumberFormat = "0.00%"
            
            ws.Range("L" & sRow).Value = stockVol
            
            'Increment sRow for next values
            sRow = sRow + 1
            
            'Reset opening price and stock volume
            openPrice = ws.Range("C" & i + 1).Value
            stockVol = 0
        Else
            stockVol = stockVol + ws.Range("G" & i).Value
        End If
    Next i
    'Adjust summary columns to autofit
    ws.Columns("I:L").AutoFit
    
Next ws
            
    
End Sub

