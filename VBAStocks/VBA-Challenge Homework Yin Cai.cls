VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub ticker()

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
    
        Dim ticker_idx As Integer
        Dim yearly_open As Double
        Dim yearly_close As Double
        Dim yearly_change As Double
        Dim volume As LongLong
        Dim max_percentage As Double
        Dim min_percentage As Double
        Dim max_percentage_ticker As String
        Dim min_percentage_ticker As String
        Dim max_volumn As LongLong
        Dim max_volumn_ticker As String
        
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        last_col = ws.Cells(1, Columns.Count).End(xlToLeft).Column
        
        ws.Range("J1").Value = "Ticker"
        ws.Range("K1").Value = "Yearly_change"
        ws.Range("L1").Value = "Percentange change"
        ws.Range("M1").Value = "Total Stock Volumn"
        ws.Range("P2").Value = "Greatest % Increase"
        ws.Range("P3").Value = "Greatest % Decrease"
        ws.Range("P4").Value = "Greatest Total Value"
        ws.Range("Q1").Value = "Ticker"
        ws.Range("R1").Value = "Value"
        
        
        ticker_idx = 2
        yearly_open = ws.Cells(2, 3).Value
        volumn = 0
        max_percentage = 0
        min_percentage = 0
        max_volumn = 0
        
        
        
        
        For i = 2 To last_row
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1) Then
                ws.Cells(ticker_idx, 10).Value = ws.Cells(i, 1).Value
                
                yearly_close = ws.Cells(i, 6).Value
                yearly_change = yearly_close - yearly_open
                ws.Cells(ticker_idx, 11).Value = yearly_change
                    If yearly_change < 0 Then
                        ws.Cells(ticker_idx, 11).Interior.ColorIndex = 3
                    ElseIf yearly_change > 0 Then
                        ws.Cells(ticker_idx, 11).Interior.ColorIndex = 4
                    End If
                    
                    
                
                    If yearly_open = 0 Then
                        ws.Cells(ticker_idx, 12).Value = 0
                    Else
                        ws.Cells(ticker_idx, 12).Value = yearly_change / yearly_open
                    End If
                    
                        
                ws.Cells(ticker_idx, 12).NumberFormat = "0.00%"
                
                ws.Cells(ticker_idx, 13).Value = volumn + ws.Cells(i, 7).Value
                
                If yearly_open <> 0 Then
                If yearly_change / yearly_open > max_percentage Then
                    max_percentage = ws.Cells(ticker_idx, 12).Value
                    max_percentage_ticker = ws.Cells(ticker_idx, 10).Value
                End If
                    
                If yearly_change / yearly_open < min_percentage Then
                    min_percentage = ws.Cells(ticker_idx, 12).Value
                    min_percentage_ticker = ws.Cells(ticker_idx, 10).Value
                End If
                End If
                
                    
                If volumn > max_volumn Then
                    max_volumn = ws.Cells(ticker_idx, 13).Value
                    max_volumn_ticker = ws.Cells(ticker_idx, 10).Value
                End If
                
                ticker_idx = ticker_idx + 1
                yearly_open = ws.Cells(i + 1, 3).Value
                volumn = 0
                
            Else
                volumn = volumn + ws.Cells(i, 7).Value
                
            End If
            
            
            
          Next i
          ws.Range("Q2").Value = max_percentage_ticker
          ws.Range("R2").Value = max_percentage
          ws.Range("R2").NumberFormat = "0.00%"
          ws.Range("Q3").Value = min_percentage_ticker
          ws.Range("R3").Value = min_percentage
          ws.Range("R3").NumberFormat = "0.00%"
          ws.Range("Q4").Value = max_volumn_ticker
          ws.Range("R4").Value = max_volumn
    Next ws
    
End Sub
