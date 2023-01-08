Attribute VB_Name = "Module2"
Sub Module2_Challenge()

'Loop through worksheets
For Each ws In Worksheets

'Declaring variables
Dim TickerName As String

Dim ConsolodatedTicker As Integer

Dim ConsolodatedOpenPrice As Long

Dim OpenPrice As Double

Dim ClosePrice As Double

Dim TotalYearlyChange As Double

Dim PercentageChange As Double

Dim TotalStockVolume As Double

Dim i As Long

Dim j As Long

Dim tickerLastRow As Long

Dim ConsolodatedTickerLastRow As Long

Dim YearlyChangeLastRow As Long

Dim PercentageChangeLastRow As Long

Dim TotalStockValueLastRow As Long

Dim GreatestPercentIncrease As Double
Dim GPITicker As String

Dim GreatestPercentDecrease As Double
Dim GPDTicker As String

Dim GreatestTotalVolume As Double
Dim GTVTicker As String



''For Summary Table One

'Create New Column Headers for Summary Table One
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percentage Change"
    ws.Range("L1") = "Total Stock Volume"
    
    'Keep track of the location for each Ticker Name and Open Price
    ConsolodatedTicker = 2
    ConsolodatedOpenPrice = 2
    
    'Set Total Stock Volume to zero
    TotalStockVolume = 0
    
    'Define tickerLastRow, <ticker> to end of column
    tickerLastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row '

''Create Summary Table One
    'Loop through <ticker><open><close><volume> to produce Summary One Table
    For i = 2 To tickerLastRow
    OpenPrice = ws.Range("C" & ConsolodatedOpenPrice).Value
    ClosePrice = ws.Range("F" & i).Value

        'Check If we are still within the same Ticker Name
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            'Set the Ticker Name in Summary Table One
            TickerName = ws.Cells(i, 1).Value
                    
            'Print the Ticker Name in column I in Summary Table One
            ws.Range("I" & ConsolodatedTicker).Value = TickerName
            
            'Add to the Total Stock Volume
            TotalStockVolume = TotalStockVolume + Cells(i, 7).Value
            
            'Print the Total Stock Volume in column L in Summary Table One
            ws.Range("L" & ConsolodatedTicker).Value = TotalStockVolume
            
            'Add to the Yearly Change in Summary Table One
            TotalYearlyChange = ClosePrice - OpenPrice
            
            'Print the Yearly Change in column J in Summary Table One
            ws.Range("J" & ConsolodatedTicker).Value = TotalYearlyChange
                            
            'Add to the Percentage Change in Summary Table One
            If OpenPrice = 0 Then
                PercentageChange = 0
                
            Else
                PercentageChange = (TotalYearlyChange / OpenPrice)
                
            End If
            
                'Highlight Yearly Change If greater than 0 then change cell to green, If less than 0 then change cell to red
                If ws.Range("J" & ConsolodatedTicker).Value > 0 Then
                    ws.Range("J" & ConsolodatedTicker).Interior.ColorIndex = 4 'green

                ElseIf ws.Range("J" & ConsolodatedTicker).Value < 0 Then
                    ws.Range("J" & ConsolodatedTicker).Interior.ColorIndex = 3 'red

                End If
            
           'Format the Percentage Change in column K in Summary Table One
             ws.Range("K" & ConsolodatedTicker).NumberFormat = "0.00%"
             
            'Print the Percentage Change in column K in Summary Table One
            ws.Range("K" & ConsolodatedTicker).Value = PercentageChange
                   
            'Add one to Summary Table One rows
            ConsolodatedTicker = ConsolodatedTicker + 1
            ConsolodatedOpenPrice = i + 1
            
            'Reset the TotalYearlyChange, TotalStockVolume
            TotalStockVolume = 0
        
        'If the cell immediately following a row is the same ticker name
        Else
        
            'Add to the Total Stock Volume
            TotalStockVolume = TotalStockVolume + ws.Range("G" & i).Value
            
        End If
        
    Next i
    
''Create Summary Table Two
'Create New Column/Row Headers for Summary Table Two
    ws.Range("O2") = "Greatest % Increase"
    ws.Range("O3") = "Greatest % Decrease"
    ws.Range("O4") = "Greatest Total Volume"
    ws.Range("P1") = "Ticker"
    ws.Range("Q1") = "Value"

'Format the Q2 and Q3 for Percentage
ws.Range("Q2:Q3").NumberFormat = "0.00%"

    'Initialize values
    GreatestPercentIncrease = ws.Range("K2").Value
    GreatestPercentDecrease = ws.Range("K2").Value
    GreatestTotalVolume = ws.Range("L2").Value
    
    'Define ConsolodatedTickerLastRow
    ConsolodatedTickerLastRow = ws.Cells(Rows.Count, "I").End(xlUp).Row
    
    'This would also work and I just want to remember this, so I'm keeping it in.
    'ws.Range("I2").End(xlDown).Select
    'ConsolodatedTickerLastRow = Selection.Row

    'Loop through to find Greatest Percent Increase, Greatest Percent Increase, and Greatest Total Volume
    For j = 2 To ConsolodatedTickerLastRow

        'Find Greatest Percent Increase Value
        If ws.Cells(j, 11).Value > GreatestPercentIncrease Then
            GreatestPercentIncrease = ws.Cells(j, 11).Value
            
            GPITicker = ws.Cells(j, 9).Value

        End If

        'Find Greatest Percent Decrease Value
        If ws.Cells(j, 11).Value < GreatestPercentDecrease Then
            GreatestPercentDecrease = ws.Cells(j, 11).Value
            
            GPDTicker = ws.Cells(j, 9).Value

        End If

        'Find Greatest Total Volume
        If ws.Cells(j, 12).Value > GreatestTotalVolume Then
            GreatestTotalVolume = ws.Cells(j, 12).Value
            
            GTVTicker = ws.Cells(j, 9).Value

        End If
            'Print GPI Value in Summary Table Two
            ws.Cells(2, 17).Value = GreatestPercentIncrease

            'Print GPI Ticker in Summary Table Two
            ws.Cells(2, 16).Value = GPITicker
            'Print GPD Value in Summary Table Two
            ws.Cells(3, 17).Value = GreatestPercentDecrease
            
            'Print GPD Ticker in Summary Table Two
            ws.Cells(3, 16).Value = GPDTicker
            
            'Print GTV Value in Summary Table Two
            ws.Cells(4, 17).Value = GreatestTotalVolume
            'Print GTV Ticker in Summary Table Two
            ws.Cells(4, 16).Value = GTVTicker
    
    Next j

        
'Format columns so everything can be read easily
ws.Columns("A:Q").AutoFit
                
Next ws

End Sub





