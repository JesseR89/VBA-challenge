Sub stockVolume()

Dim ws As Worksheet

For Each ws In Worksheets
    Dim lastRow As Integer
   lastRow = ws.Range("A1").End(xlDown).Row
    Dim cell As Range
    Dim tickers As Collection
    Dim tickervalue As Variant
    Dim tickercell As Range
    Dim I As Long
    Dim greatestIncrease As Double
    


    Application.ScreenUpdating = False


    ticker = ws.Cells(2, 1).Value
    tickertotal = 2
    open_price = ws.Cells(2, 3).Value
    TotalStockVolumn = Cells(2, 7)

    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volumn"
    
    For tickerrow = 2 To Range("A" & Rows.Count).End(xlUp).Row
        
        If ticker <> ws.Cells(tickerrow + 1, 1).Value Then
        Cells(tickertotal, 9).Value = ticker
        ticker = Cells(tickerrow + 1, 1).Value
        
        YearlyChange = Cells(tickerrow, 6).Value - open_price
        ws.Cells(tickertotal, 10).Value = YearlyChange
        ws.Cells(tickertotal, 11).Value = YearlyChange / open_price
        ws.Cells(tickertotal, 12).Value = TotalStockVolumn + Cells(tickerrow, 7).Value
        TotalStockVolumn = 0
        tickertotal = tickertotal + 1
        open_price = ws.Cells(tickerrow + 1, 3).Value
        
    

        Else
  
            TotalStockVolumn = TotalStockVolumn + Cells(tickerrow, 7).Value
        
        End If
    
    
    Set tickers = New Collection
    
    For Each cell In ws.Range("A2:A" & lastRow)
        On Error Resume Next
        tickers.Add cell.Value, CStr(cell.Value)
        On Error GoTo 0
        Next ws.cell
    
    On Error Resume Next
    Set tickercell = ws.Rows(1).Find(What:="Ticker", LookIn:=x1Values, LookAt:=(x1Whole))
    On Error GoTo 0
    
    If tickercell Is Nothing Then
        Set tickercell = ws.Cells(1, 1)
        tickercell.Value = "ticker"
    End If
    
    lastRow = Cells(Rows.Count, 1).End(x1up).Row

    greatestIncrease = Cells(2, 1).Value
   Set greatestIncrease = Cells(2, 1)

    For I = 2 To lastRow
         If ws.Cells(I, 1).Value > greatestIncrease Then
            greatestIncrease = ws.Cells(I, 1).Value
            Set greatestIncrease = ws.Cells(I, 1)
        End If
    Next I
Next tickerrow
Next ws
End Sub
