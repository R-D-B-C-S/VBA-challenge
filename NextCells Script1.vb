Sub NextCells():
Dim Row As Long
Dim i As Long
Dim r As Long
Dim ri As Long
Dim TempSymbol As String
Dim TempVolume As Double
Dim FirstPrice As Double
Dim LastPrice As Double
Dim SummaryTableRow As Integer
Dim YearlyChange As Double
Dim PercentChange As Double

'Constants for Summary Table Columns

Const TickerColumn As Integer = 9
Const YearlyChangeColumn As Integer = 10
Const VolumeColumn As Integer = 11
Const PercentColumn As Integer = 12

For Worksheet = 1 To ActiveWorkbook.Worksheets.Count
Worksheets(Worksheet).Activate


Summary_Table_Row = 2
TempVolume = 0
TempSymbol = ""
LastPrice = 0
'Declare first value for price
FirstPrice = Cells(2, 3).Value




    For Row = 2 To 753001
        If Cells(Row + 1, 1).Value <> Cells(Row, 1).Value And Cells(Row + 1, 1).Value <> 0 Then
            
            'temporary variables
            
            
            TempSymbol = Cells(Row, 1).Value
            LastPrice = Cells(Row, 6).Value
            
            'Update Variables
            
            'TempSymbol = Cells(Row, 1).Value
            'TempVolume = Cells(Row, 7).Value
            YearlyChange = LastPrice - FirstPrice
            PercentChange = YearlyChange / FirstPrice
            
            'Update Summary Sheet
            Cells(Summary_Table_Row, TickerColumn).Value = Cells(Row, 1).Value
            Cells(Summary_Table_Row, YearlyChangeColumn).Value = YearlyChange
            Cells(Summary_Table_Row, VolumeColumn).Value = Cells(Row, 7).Value + TempVolume
            Cells(Summary_Table_Row, PercentColumn).Value = PercentChange
            FirstPrice = Cells(Row + 1, 3).Value
            TempVolume = 0
            'Go down new row for summary table
            Summary_Table_Row = Summary_Table_Row + 1
        Else
            'Store accumulative Values
            
            TempVolume = TempVolume + Cells(Row, 7).Value
            'TempSymbol = Cells(Row, 1).Value
            'TempPrice
            'Update Summary Table

            'Cells(Summary_Table_Row, VolumeColumn).Value = TempVolume
            'Here would update percent change, might be better in if section than else section since its percent difference between first price and last price
            
        End If
        
    Next Row

    For i = 2 To 2977
        If Cells(i, YearlyChangeColumn) > 0 Then
            Cells(i, YearlyChangeColumn).Interior.ColorIndex = 4
        Else
            Cells(i, YearlyChangeColumn).Interior.ColorIndex = 3
        End If
    Next i
    'Finding max, min and largest volume
    Cells(2, 16).Value = WorksheetFunction.Max(Range("l2:l3001"))
    Cells(3, 16).Value = WorksheetFunction.Min(Range("l2:l3001"))
    Cells(4, 16).Value = WorksheetFunction.Max(Range("k2:k3001"))
    
    'Assigning the ticker based on if the row it's in matches the value
    For i = 2 To 3001
        If Cells(i, PercentColumn).Value = Cells(2, 15).Value And Cells(i, PercentColumn).Value <> 0 Then
            Cells(2, 15).Value = Cells(i, 9).Value
        'Else
            'Cells(2, 15).Value = "condition failed"
        End If
        If Cells(i, PercentColumn).Value = Cells(3, 15).Value And Cells(i, PercentColumn).Value <> 0 Then
            Cells(3, 15).Value = Cells(i, 9).Value
        'Else
            'Cells(3, 15).Value = "condition failed"
        End If
        If Cells(i, VolumeColumn).Value = Cells(4, 15).Value And Cells(i, VolumeColumn).Value <> 0 Then
            Cells(4, 15).Value = Cells(i, 9).Value
        'Else
            'Cells(4, 15).Value = "condition failed"
        End If
    Next i
    

Next Worksheet
End Sub
