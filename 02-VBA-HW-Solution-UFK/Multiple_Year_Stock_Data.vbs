Sub Alphabetical_Testing()
For Each ws In Worksheets
    'Gets the values from specified worksheet below
    'Set ch = ThisWorkbook.Sheets("2014")

    Dim lastrow As Double
    'Counts the number of rows in column. Got this code from stackoverflow
    'double will work with loop but long won't
    lastrow = ws.Range("A1", ws.Range("A1").End(xlDown)).Rows.Count


    Dim total_vol As Double
    total_vol = 0
    
    Dim i As Double
    Dim j As Double
    j = 1
    Dim k as Double

    Dim open1 As Double
    open1 = ws.Cells(2, 3).Value
    
    'close and open were illegal variables
    Dim close1 As Double
    Dim delta As Double
    Dim percentage As Double
    
    'Set labels
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Volume"
    ws.Cells(1, 11).Value = "Yearly Change"
    ws.Cells(1, 12).Value = "Percentage Change"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1,17).Value = "Value"
    ws.Cells(4,15).Value = "Greatest Volume"
    ws.Cells(3,15).Value = "Greatest % Decrease"
    ws.Cells(2, 15).Value = "Greatest % Increase"

    

'loop from first populated row to last row
For i = 2 To lastrow
        'Whenever there is a change in the ticker value, it copies the ticker and volume
        
        If (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then
            
            j = j + 1
            'Sets ticker value
            ws.Cells(j, 9).Value = ws.Cells(i, 1).Value
            
            'Calculates total volume for that ticker, copies it and resets
            total_vol = total_vol + ws.Cells(i, 7).Value
            ws.Cells(j, 10).Value = total_vol
            total_vol = 0
            'MsgBox() for debugging
            
            
            'Get the value of the current cell as the closing value of the current ticker
            close1 = ws.Cells(i, 6).Value
            
            'calculate difference between the two and sets value
            delta = close1 - open1
            ws.Cells(j, 11).Value = delta
            
            'calculate percentage difference and calculate value
            ws.Cells(j, 12).Value = delta / open1
            
            'Set style
            ws.Range("L2:L" & j).Style = "Percent"
            
            'pick the value of the next cell as the opening value of the first cell
            'If the value is zero then pick 0.0001 to prevent overflow
            If (ws.Cells(i + 1, 3).Value = 0) Then
                open1 = 0.0001
            Else
                open1 = ws.Cells(i + 1, 3).Value
            End If
                
                'color change
                If (ws.Cells(j, 12).Value > 0) Then
                    ws.Cells(j, 12).Interior.ColorIndex = 4
                Else
                    ws.Cells(j, 12).Interior.ColorIndex = 3
                End If
            
            Else
            total_vol = total_vol + ws.Cells(i, 7).Value
        End If

    Next i

    'Set starting value for KPIs
    'Greatest Volume
    ws.Cells(4, 16).Value = ws.Cells(2,9).Value
    ws.Cells(4, 17).Value = ws.Cells(2,10).Value
    
    'Greatest Percent Decrease
    ws.Cells(3, 16).Value = ws.Cells(2,9).Value
    ws.Cells(3, 17).Value = ws.Cells(2,12).Value

    'Greatest Percent Increase
    ws.Cells(2, 16).Value = ws.Cells(2,9).Value
    ws.Cells(2, 17).Value = ws.Cells(2,12).Value

    'loop through j since you have the final value for the current worksheet
    For k = 2 to j

        'Compares Greatest Volume of row 2 to row k to find max
        if (ws.Cells(k, 10).Value > ws.Cells(4, 17).Value) then
            ws.Cells(4, 17).Value = ws.Cells(k, 10).Value
			ws.Cells(4, 16).Value = ws.Cells(k,9).Value
        End If 

        'Compares Percent change of row 2 to row k to find min
        if (ws.Cells(k, 12).Value < ws.Cells(3, 17).Value) then
            ws.Cells(3, 17).Value = ws.Cells(k, 12).Value
			ws.Cells(3, 16).Value = ws.Cells(k,9).Value
        End If 

        'Compares Percent change of row 2 to row k to find max
        if (ws.Cells(k, 12).Value > ws.Cells(2, 17).Value) then
            ws.Cells(2, 17).Value = ws.Cells(k, 12).Value
			ws.Cells(2, 16).Value = ws.Cells(k,9).Value
        End If 

        'Set Style
        ws.Cells(2, 17).Style = "Percent"
        ws.Cells(3, 17).Style = "Percent"


    Next k 

Next ws


'MsgBox () for




End Sub


