Sub Ticker():
'Set ws as WorkSheet
Dim ws As Worksheet
'Loop Worksheets
For Each ws In Worksheets

    'Colum title
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "YearOpen"
    ws.Range("K1").Value = "YearClose"
    ws.Range("L1").Value = "Yearly Change"
    ws.Range("M1").Value = "Percent of Yearly Change"
    ws.Range("N1").Value = "Total Volume"
    
    'Colum Width for Summary Table
    ws.Range("I:N").ColumnWidth = 14
    
    'Variables and Values
    Dim Ticker As String
    Dim NameTicker As String
    Dim SummaryTableRow As Integer
    Summary_Table_Row = 2

    Dim YearStart As Double
    YearStart = 0
    Dim YearClose As Double
    Dim Yearly_Change As Double
    Dim PercentChange As Double
    Dim Volume As Double
    
    
    
        
    Dim LastRow As Double
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    YearStart = 0
    Volume = 0



    'loop ticker
    For i = 2 To LastRow
    
        'Comparar el nombre de la empresa con la celda siguiente'
        If IsNull(NameTicker) = True Then
            NameTicker = ws.Cells(i, 1).Value
        End If
        
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        
        'A selecionar el primer valor year, lo guardas en una variable and compare to end'
        
            NameTicker = ws.Cells(i, 1).Value
            
            
            'print in range
            ws.Range("I" & Summary_Table_Row).Value = NameTicker
            
            YearClose = ws.Cells(i, 6).Value
            
            ws.Range("J" & Summary_Table_Row).Value = YearStart
            
            ws.Range("K" & Summary_Table_Row).Value = YearClose
            
            Yearly_Change = YearClose - YearStart

            ws.Range("L" & Summary_Table_Row).Value = Yearly_Change
            
            'Si YearStart = 100% Yearly Change is x %
            
            If YearStart <> "0" Then
                PercentChange = Yearly_Change / YearStart
            End If
            
            ws.Range("M" & Summary_Table_Row).Value = PercentChange
            ws.Range("N" & Summary_Table_Row).Value = Volume
            
            'format
            
            If PercentChange > "0" Then
                ws.Range("M" & Summary_Table_Row).Interior.ColorIndex = 4
            ElseIf PercentChange < "0" Then
                ws.Range("M" & Summary_Table_Row).Interior.ColorIndex = 3
            End If
            
            ws.Range("M" & Summary_Table_Row).NumberFormat = "0.00%"
            
            
           
            ws.Range("N" & Summary_Table_Row).NumberFormat = "#,##0"
            
            
            
            Summary_Table_Row = Summary_Table_Row + 1
            YearStart = 0
            Volume = 0
        
        Else
            'Volume
            If YearStart = "0" Then
                YearStart = ws.Cells(i, 3).Value
            End If
            
            Volume = Volume + ws.Cells(i, 7).Value
        End If
        
        
            
        
                
    
    Next i
Next


End Sub
