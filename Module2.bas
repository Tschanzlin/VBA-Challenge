Attribute VB_Name = "Module2"

Option Explicit

Sub StockData1()

'Additional Work Notes
    'Format output cells
    'NOTE:  Alpha worksheet - PLNT ticker had no data; deleted to fix run-time error but could have added If Then statement to check for zero values

'Declare variables
    'Objects
        Dim ws As Object
    'Input variables
        Dim Ticker As String
        Dim OpenPx As Double, ClosePx As Double, VolSum As Double
    'Function variables
        Dim Cnt As Long, LastRow As Long, i As Long
    'Ouput variables
        Dim PxDolChg As Double, PxPerChg As Double
        Dim MaxPxIncr As Double, MaxPxDecr As Double, MaxVol As Double
        Dim MaxTicker As String, MinTicker As String, MaxVolTicker As String
       
'Loop through each worksheet
    
    For Each ws In Worksheets
 
    Cnt = 2
    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    'Loop through worksheet dataset; set initial Max and Min Price Change and Vol values to zero
        
        MaxPxIncr = 0
        MaxPxDecr = 0
        MaxVol = 0
                
        For i = 1 To LastRow
          
        'Reset values before each loop
            
            VolSum = 0
            Ticker = ws.Cells(Cnt, 1).Value
            OpenPx = ws.Cells(Cnt, 3).Value
        
        
        'Loop to caculate VolSum and Set Counter
        'If statement need to avoid 1004 run-time error and continue cyclying throus spreadsheets
          
            If ws.Cells(Cnt, 2).Value > 0 Then
                Do While Ticker = ws.Cells(Cnt, 1).Value
                    VolSum = VolSum + ws.Cells(Cnt, 7).Value
                    ClosePx = ws.Cells(Cnt, 6).Value
                    Cnt = Cnt + 1
                Loop
              
            'Calculate and Display Ticker Output (PxDolChg and PxPerChg) before beginning next Ticker loop
                PxDolChg = ClosePx - OpenPx
                PxPerChg = PxDolChg / OpenPx
                ws.Cells(i + 1, 9).Value = Ticker
                ws.Cells(i + 1, 10).Value = PxDolChg
                ws.Cells(i + 1, 11).Value = PxPerChg
                ws.Cells(i + 1, 12).Value = VolSum
                
            'Format PxDolChg cells
                If PxDolChg > 0 Then
                    ws.Cells(i + 1, 10).Interior.ColorIndex = 4
                ElseIf PxDolChg < 0 Then
                    ws.Cells(i + 1, 10).Interior.ColorIndex = 3
                End If
          
            'Calculate MaxPxIncr, MaxPxDecr, MaxVol and related tickers
            
                If PxPerChg > MaxPxIncr Then
                    MaxPxIncr = PxPerChg
                    MaxTicker = ws.Cells(i + 1, 9).Value
                ElseIf PxPerChg < MaxPxDecr Then
                    MaxPxDecr = PxPerChg
                    MinTicker = ws.Cells(i + 1, 9).Value
                End If
                
                If VolSum > MaxVol Then
                    MaxVol = VolSum
                    MaxVolTicker = ws.Cells(i + 1, 9).Value
                End If
              
                 
            End If
          
            'Diplay Max / Min workshet values
                ws.Cells(2, 15).Value = MaxTicker
                ws.Cells(2, 16).Value = MaxPxIncr
                ws.Cells(3, 15).Value = MinTicker
                ws.Cells(3, 16).Value = MaxPxDecr
                ws.Cells(4, 15).Value = MaxVolTicker
                ws.Cells(4, 16).Value = MaxVol
        
        Next i
        
    Next ws

End Sub
