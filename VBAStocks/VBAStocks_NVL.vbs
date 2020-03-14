Sub GetYearlyData()

    Dim Ticker As String
    Dim YearOpen As Double
    Dim YearClose As Double
    Dim TotalVolume As Double
    Dim LastRow As Long
    Dim SummaryRow As Integer
    Dim i As Long
    

    For Each ws In Worksheets
	'Set Up Summary Table Header and First Row
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        SummaryRow = 2
        
	'Get Last Row of current Worksheet
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

	'Get First Data Row on the sheet  
        Ticker = ws.Cells(2, 1).Value
        YearOpen = ws.Cells(2, 3).Value
        TotalVolume = ws.Cells(2, 7).Value
        
        For i = 3 To LastRow + 1
	    'Check if Ticker Symbol changed 
            If ws.Cells(i, 1).Value = Ticker Then
		'No change Add To Total
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            Else
		'Changed Write Summary Data
                ws.Cells(SummaryRow, 9).Value = Ticker
                ws.Cells(SummaryRow, 10).Value = ws.Cells(i - 1, 6).Value - YearOpen
                If ws.Cells(SummaryRow, 10).Value >= 0 Then
                    ws.Cells(SummaryRow, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(SummaryRow, 10).Interior.ColorIndex = 3
                End If
                If YearOpen = 0 Then
                    ws.Cells(SummaryRow, 11).Value = 0
                Else
                    ws.Cells(SummaryRow, 11).Value = ws.Cells(SummaryRow, 10) / YearOpen
                End If
                ws.Cells(SummaryRow, 11).NumberFormat = "0.00%"
                ws.Cells(SummaryRow, 12).Value = TotalVolume

		'Set Values for New Ticker Symbol
                Ticker = ws.Cells(i, 1).Value
                YearOpen = ws.Cells(i, 3).Value
                TotalVolume = ws.Cells(i, 7).Value

		'Increment Summary Table Row
                SummaryRow = SummaryRow + 1
            End If
        Next i
        
    Next ws

    'Run Sub routine to get Maximum values
    Call FindMaxValues
End Sub

Sub FindMaxValues()

    Dim LastRow As Long
    Dim MatchIndex As Long
    Dim MaxPercent As Double
    Dim MaxVolume As Double
    
    For Each ws In Worksheets
        'Set Maximum Table Headers
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"

        'Get Last Row for the current worksheet
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'Get Max Percent Gain for the K Column
        MaxPercent = WorksheetFunction.Max(ws.Range("K2:K" & LastRow))
        ws.Range("Q2").Value = MaxPercent
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        MatchIndex = WorksheetFunction.Match(MaxPercent, ws.Range("K2:K" & LastRow), 0)
        ws.Range("P2").Value = ws.Cells(MatchIndex + 1, 9)
 
        'Get Max Percent Loss (Min) for the K Column
        MaxPercent = WorksheetFunction.Min(ws.Range("K2:K" & LastRow))
        ws.Range("Q3").Value = MaxPercent
        MatchIndex = WorksheetFunction.Match(MaxPercent, ws.Range("K2:K" & LastRow), 0)
        ws.Range("P3").Value = ws.Cells(MatchIndex + 1, 9)
       
        'Get Max Volume for L Column	
        MaxVolume = WorksheetFunction.Max(ws.Range("L2:L" & LastRow))
        ws.Range("Q4").Value = MaxVolume
        MatchIndex = WorksheetFunction.Match(MaxVolume, ws.Range("L2:L" & LastRow), 0)
        ws.Range("P4").Value = ws.Cells(MatchIndex + 1, 9)
 
        'Adjust Cell Width
	ws.Columns("A:Q").AutoFit      
    Next ws

End Sub
