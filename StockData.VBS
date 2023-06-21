

Sub Stock_Market()
    ' --------------------------------------------
    ' Run on every worksheet (that is, every year) at once.
    ' --------------------------------------------
    For Each ws In Worksheets

        ' --------------------------------------------
        ' Loop through one year stock data and reads/stores:
        ' ticker symbol, volume of stock, open price, close price
        ' --------------------------------------------
        
        Dim LastRow As Long
        Dim NewTableRow As Integer
        Dim Ticker As String
        Dim OpenTotal As Double
        Dim CloseTotal As Double
        Dim Change As Double
        Dim YearChange As Double
        Dim Count As Integer
        Dim PercentChange As Double
        Dim VolumeTotal As Double
        Dim GreatestIncrease As Double
        Dim GrestestDecrease As Double
        Dim GreatestTotalVolume As Double
        Dim GITicker As String
        Dim GDTicker As String
        Dim GTVTicker As String

        'Initialize count, NewTableRow for output and row count, Greatest Values
        NewTableRow = 2
        Count = 1
        GreatestIncrease = 0
        GreatestDecrease = 0
        GreatestTotalVolume = 0

        ' Determine the Last Row of provided data
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'MsgBox (LastRow)

        

        'Write Headers for Results Columns
        ws.Cells(1,9).Value = "Ticker"
        ws.Cells(1,10).Value = "Yearly Change"
        ws.Cells(1,11).Value = "Percent Change"
        ws.Cells(1,12).Value = "Total Stock Volume"
        ws.Cells(1,17).Value = "Ticker"
        ws.Cells(1,18).Value = "Value"
        'Write Row Labels For Greatest results on Increase, Decrease, and Total Volume
        ws.Cells(2,16).Value = "Greatest % Increase"
        ws.Cells(3,16).Value = "Greatest % Decrease"
        ws.Cells(4,16).Value = "Greatest Total Volume"
       
       '---------------------------------------------
       'For each unique stock diplay the Ticker symbol, it's change per year, percent change, and total volume
       '---------------------------------------------

        'Loop through rows 2 to the last row
        For i = 2 To LastRow

            'If the next ticker value is different perform calculations and display results in same sheet
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                
                'write Ticker 
                ws.Cells(NewTableRow, 9).Value = ws.Cells(i, 1).Value

                'calculate last row for Totals, change for year and pecentage change from opening price.
                OpenTotal = OpenTotal + ws.Cells(i, 3).Value
                CloseTotal = CloseTotal + ws.Cells(i, 6).Value
                VolumeTotal = VolumeTotal + ws.Cells(i,7).Value 
                YearChange = CloseTotal - OpenTotal
                Percentchange = (YearChange / (OpenTotal/Count))

                'write YearChange
                ws.Cells(NewTableRow, 10).Value = YearChange                
                'Format and write Percentage
                 ws.Cells(NewTableRow, 11).Value = FormatPercent(PercentChange)
                'Write Volume Total
                ws.Cells(NewTableRow, 12).Value = VolumeTotal 


                

                'Format yearly change output cells red for negative green for positive including 0
                If YearChange < 0 Then
                    ws.Cells(NewTableRow, 10).Interior.ColorIndex = 3
                
                Else
                    ws.Cells(NewTableRow, 10).Interior.ColorIndex = 4

                End If
                
                'Increment NewTableRow for next stock
                NewTableRow = NewTableRow + 1

                'reset OpenTotal, Close Total, YearChange,VoulmeTotal, and Count for next stock
                OpenTotal = 0
                CloseTotal = 0
                YearChange = 0
                VolumeTotal = 0
                Count = 1
                
            Else 

                'Calculate Open and Close Total Values, Volume Total, and count number of rows per stock
                OpenTotal = OpenTotal + ws.Cells(i, 3).Value
                CloseTotal = CloseTotal + ws.Cells(i, 6).Value
                VolumeTotal = VolumeTotal + ws.Cells(i,7).Value
                Count = Count + 1
                
                
            End If
            
        Next i

        '-------------------------------------------------------------------
        'Calculate and display Greatest results on Increase, Decrease, and Total Volume
        '--------------------------------------------------------------------

        For i = 2 To NewTableRow
            'local variables to compare to Greatest Values
            Percentage = ws.Cells(i, 11).Value
            Volume = ws.Cells(i,12).Value
            
            'Check if current value is larger than last highest value if so save as new highest value with Ticker name
            If Percentage > GreatestIncrease Then
                GreatestIncrease = Percentage
                GITicker = ws.Cells(i,9).Value
           
            'Check if current value is lower than last lowest value if so save as new lowest value with Ticker name
            ElseIf Percentage < GreatestDecrease Then
                GreatestDecrease = Percentage
                GDTicker = ws.Cells(i,9).Value 

            Else      
                'do nothing and go to next stock

            End If
           
            'Check if current value is larger than last highest value if so save as new highest value with Ticker name
            If Volume > GreatestTotalVolume Then
              GreatestTotalVolume = Volume
              GTVTicker = ws.Cells(i, 9).Value
            End If  

        Next i 

        'Write Ticker symbols and Values for Greatest Percent Increase and Decrease, and Greatest Total Volume
        ws.Cells(2,17).Value = GITicker
        ws.Cells(2,18).Value = FormatPercent(GreatestIncrease)
        ws.Cells(3,17).Value = GDTicker
        ws.Cells(3,18).Value = FormatPercent(GreatestDecrease)
        ws.Cells(4,17).Value = GTVTicker
        ws.Cells(4,18).Value = GreatestTotalVolume

    Next ws
End Sub



