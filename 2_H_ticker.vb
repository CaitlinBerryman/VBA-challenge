Sub ticker()

'big tick energy

'dim variables. determine which loops counters need to be in/out of, remember to reset them as needed

'last row of a worksheet
Dim lastrow As Long

'counter creating a new row for each ticker
Dim ticker As Long

'row counter to fetch open value of the first entry of a ticker
Dim openrow As Long

'first open value of a ticker
Dim openval As Double

'final close value of a ticker
Dim closeval As Double

'stock volume
Dim vol As LongLong


'loop through sheets

For a = 1 To 3

    'make sure the active cell location updates when sheet switches
    Sheets(a).Select
    Cells(1, 1).Select
    
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'prints combined data for each ticker group on a new row
    ticker = 1
    
    vol = 0
    
    'this will fetch the open value of the first row of the ticker group
    openrow = 0
    
    

    'enter ticker/yearchange/%change/volume on each ws
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"

    'loop through rows
    For i = 2 To lastrow
    
        'if ticker name matches the name in the following row, add counters
        If Cells(i + 1, 1).Value = Cells(i, 1).Value Then
            vol = vol + Cells(i, 7).Value
            'half of PLNT's 2015 values are zero, this is to skip them so it only starts the row count from the first non-zero open value
            If Cells(i, 3).Value > 0 Then
                openrow = openrow + 1
            End If
            
        'if they are different (end of group)
        Else
            'counters and math
            ticker = ticker + 1
            vol = vol + Cells(i, 7).Value
            openval = Cells(i - openrow, 3).Value
            closeval = Cells(i, 6).Value
            
            'name
            Cells(ticker, 9).Value = Cells(i, 1).Value
            
            'yearly change formula, using open and close values. conditional format colour. 3 = red, 4 = green
            Cells(ticker, 10).Value = closeval - openval
            If Cells(ticker, 10).Value > 0 Then
                Cells(ticker, 10).Interior.ColorIndex = 4
            ElseIf Cells(ticker, 10).Value < 0 Then
                Cells(ticker, 10).Interior.ColorIndex = 3
            End If
            
            'those changes as % using the first open value
            'PLNT being a pain again, all 2014 values are zero so this removes the divide by zero error
            If openval > 0 Then
                Cells(ticker, 11).Value = (closeval - openval) / openval
            Else
                Cells(ticker, 11).Value = 0
            End If
            Cells(ticker, 11).NumberFormat = "0.00%"
            
            'total stock volume
            Cells(ticker, 12).Value = vol
            
            'reset vol and openrow counters (not ticker!)
            vol = 0
            openrow = 0
    
        End If

    Next i

'bonus round!

'set up table
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"
Columns("O:O").EntireColumn.AutoFit

'maths
Cells(2, 17).Formula = "=MAX(K:K)"
Cells(2, 17).NumberFormat = "0.00%"
Cells(2, 16).Formula = "=INDEX(I:I,MATCH(Q2,K:K,0))"
Cells(3, 17).Formula = "=MIN(K:K)"
Cells(3, 17).NumberFormat = "0.00%"
Cells(3, 16).Formula = "=INDEX(I:I,MATCH(Q3,K:K,0))"
Cells(4, 17).Formula = "=MAX(L:L)"
Cells(4, 16).Formula = "=INDEX(I:I,MATCH(Q4,L:L,0))"
Columns("Q:Q").EntireColumn.AutoFit

Next a



End Sub
