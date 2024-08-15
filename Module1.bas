Attribute VB_Name = "Module1"
Sub Multi_year_stock()
    'create a variable to hold the counter
    Dim i As Long
    Dim j As Long
    Dim ws As Worksheet
    
    'Loop thru all sheets
    For Each ws In Worksheets
    
    
    
    'set value for summary table
    ws.Cells(1, 9).value = "Ticker"
    ws.Cells(1, 10).value = "Quarterly Change"
    ws.Cells(1, 11).value = "Percentage Change"
    ws.Cells(1, 12).value = "Total Stock Volume"
    
    'Set an initial variable to hold the ticker, value change quarterly, % change quarterly
    Dim ticker As String
    Dim value_change As Double
    Dim percentage_change As Double
    
    'Set  an initial variable to hold the open /end quarter
    Dim Open_Q As Double
    Dim End_Q As Double
    
    'Set an initial value to hold the Open Quarter
    Open_Q = ws.Cells(2, 3).value
    
    'Set an initial variable to hold the volume
    Dim volume As Double
    volume = 0
    
    'Keep track of the location of each ticker in the summary table
    Dim ST_Row As Double
    ST_Row = 2
    ws.Range("J" & ST_Row).value = Open_Q
    
    'determine the last row
    LastRow = ws.Cells(Rows.count, 1).End(xlUp).Row
    ' MsgBox LastRow
    
    ' Loop thru all stocker
    For i = 2 To LastRow
      
        'Check if we are still screening the same stock, if it is not
        If ws.Cells(i + 1, 1).value <> ws.Cells(i, 1).value Then
                
            'Set the ticker
            ticker = ws.Cells(i, 1).value
            
            'Add to the volume total
            volume = volume + ws.Cells(i, 7).value
            
            'Set the end quarter value
            End_Q = ws.Cells(i, 6).value
            
            'calculate the value change & % change quarterly
            value_change = End_Q - Open_Q
            percentage_change = Round(value_change * 100 / Open_Q, 2)
            
            'print the value change & % change quarterly
            ws.Range("J" & ST_Row).value = value_change
            ws.Range("K" & ST_Row).value = Str(percentage_change) + "%"
            
            'condition formatting for percentage change
            If percentage_change > 0 Then
            ws.Range("K" & ST_Row).Interior.ColorIndex = 4
            
            ElseIf ws.Range("K" & ST_Row).value < 0 Then
            ws.Range("K" & ST_Row).Interior.ColorIndex = 3
            End If
            


            'print ticker to the summary table
            ws.Range("I" & ST_Row).value = ticker
            
            'print the total volume to the summary table
            ws.Range("L" & ST_Row).value = volume
            
            ' add one row to the summary table
            ST_Row = ST_Row + 1
            
            Open_Q = ws.Cells(i + 1, 3).value
            
            
            'reset the volume
            volume = 0
              
            'if the cell immadiately following a row in the same brand
            Else
            
            'Add to the total volume
            volume = volume + ws.Cells(i, 7).value
            
        End If
        
            Next i
Next ws
End Sub


'new sub to find mix/max
Sub finding_min_max()

'set initial variable to hold mix,max change and volume, and worksheet
Dim max As Double
Dim min As Double
Dim max_total As Double
Dim ws As Worksheet



'Loop thru all sheets
For Each ws In Worksheets
' set inital value for max, min, and max total volume
max = 0
min = 0
max_total = 0


'Set Headers
ws.Cells(1, 16).value = "Ticker"
ws.Cells(1, 17).value = "Value"
ws.Cells(2, 15).value = "Greatest % increase"
ws.Cells(3, 15).value = "Greatest % decrease"
ws.Cells(4, 15).value = "Greatest Total Volume"

'set initial variable to find the last row of summary table
Dim LastRow_ST As Long
Dim j As Integer


'find the last row of summary table
LastRow_ST = ws.Range("K" & Rows.count).End(xlUp).Row

'loop from 2 to last row of summary table
'to find max increase
For j = 2 To LastRow_ST
If ws.Range("K" & j).value >= max Then
max = ws.Range("K" & j).value
ws.Cells(2, 16).value = ws.Range("I" & j).value
End If

'to find max decrease
If ws.Range("K" & j).value <= min Then
min = ws.Range("K" & j).value
ws.Cells(3, 16).value = ws.Range("I" & j).value
End If

'to find max volume
If ws.Range("L" & j).value >= max_total Then
max_total = ws.Range("L" & j).value
ws.Cells(4, 16).value = ws.Range("I" & j).value
End If

Next j
'print the max increase/decrease/volume
ws.Cells(2, 17).value = Str(max * 100) + "%"
ws.Cells(3, 17).value = Str(min * 100) + "%"
ws.Cells(4, 17).value = max_total

Next ws

End Sub

