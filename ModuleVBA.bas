Attribute VB_Name = "Module1"
Sub TickerSymbol():

    Dim ws As Worksheet
    Dim i As Long
    Dim LastRow As Long
    Dim TableRow As Long
    Dim Ticker As String
    Dim vol As Integer
    Dim year_open As Double
    Dim year_close As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim Total_volume As Double
    Dim Workbook As Worksheets
    Dim Count As Integer
    Dim year As Long
    Dim Value As Double
    
    
    For Each ws In Worksheets
    Count = Count + 1
   
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("k1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("Q1").Value = "Ticker"
    ws.Range("R1").Value = "Value"
    ws.Range("P2").Value = "Greatest % Increase"
    ws.Range("P3").Value = "Greatest % Decrease"
    ws.Range("P4").Value = "Greatest Total Volume"
    ws.Range("L1").Value = "Total Stock Volume"
    
    
    Count = 0
    TableRow = 2
    LastRow = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row
    
    year_open = ws.Cells(2, 3).Value
     
     For i = 2 To LastRow
     
        If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
        Ticker = ws.Cells(i, 1).Value
        Total_volume = Total_volume + ws.Cells(i, 7).Value
        ElseIf ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
        Ticker = ws.Cells(i, 1).Value
        Total_volume = Total_volume + ws.Cells(i, 7).Value
        
        
        
        year_close = ws.Cells(i, 6).Value
        
        yearly_change = year_close - year_open
    
        percent_change = (year_close - year_open) / year_open
        
        
        ws.Cells(TableRow, 9).Value = Ticker
        ws.Cells(TableRow, 10).Value = yearly_change
        ws.Cells(TableRow, 11).Value = percent_change
        ws.Cells(TableRow, 12).Value = Total_volume
        
        year_open = ws.Cells(i + 1, 3).Value
        
        
        TableRow = TableRow + 1
        Total_volume = 0
        
        
       
        End If
        Next i
        
        
        ws.Columns("K").NumberFormat = "0.00%"
        ws.Cells(2, 18).NumberFormat = "0.00%"
        ws.Cells(3, 18).NumberFormat = "0.00%"
        
        ws.Columns("I:R").AutoFit
        
        
        
    
        
        
        Dim j As Long
        
        
        

  For j = 2 To LastRow
  

      If ws.Cells(j, "J") >= 0 Then

        ws.Cells(j, "J").Interior.ColorIndex = 4


      ElseIf ws.Cells(j, "J") < 0 Then

        ws.Cells(j, "J").Interior.ColorIndex = 3

      End If

    

    Next j
    
    
  Dim change As Range
    Set change = Range("K2:K753001")
    Dim total As Range
    Set total = Range("L2:L753001")
    

    ws.Cells(2, 18).Value = WorksheetFunction.Max(change)
    ws.Cells(3, 18).Value = WorksheetFunction.Min(change)
    ws.Cells(4, 18).Value = WorksheetFunction.Max(total)

 Next ws
        
End Sub


