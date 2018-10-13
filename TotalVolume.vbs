Attribute VB_Name = "Module1"
Sub TotalVolume()
'create a loop that will combine the total volume of each stock in a given year

'Lets start with just turning ticker symbol into a variable. That way we can search it later.
Dim TickerSymbol As String

'Next we're going to turn the TickerVolume into a variable we can add stuff with
Dim TickerVolume As Double
TickerVolume = 0

Dim YearSummary As Integer

'Let's start the loop.
For Each ws In Worksheets

'We're going to insert a new column at the end of each worksheet and name it "Ticker"
    YearSummary = 2
    ws.Range("I1").EntireColumn.Insert
    ws.Cells(1, 9).Value = "Ticker"

'Then do the same thing but name the new column "Total Stock Volume"

    ws.Range("J1").EntireColumn.Insert
    ws.Cells(1, 10).Value = "Total Stock Volume"

'Alright, let's start the For Loop
    For i = 2 To ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    'Then we differentiate the ticker symbols from each other
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        TickerSymbol = ws.Cells(i, 1).Value
        'Then we add the TickerVolume together
        TickerVolume = TickerVolume + ws.Cells(i, 7).Value
       'Then we print to the columns
        ws.Range("I" & YearSummary).Value = TickerSymbol
        ws.Range("J" & YearSummary).Value = TickerVolume
        YearSummary = YearSummary + 1
        TickerVolume = 0
        
        Else
        
        TickerVolume = TickerVolume + ws.Cells(i, 7).Value
        
        End If
    Next i
Next ws
    
End Sub
