Attribute VB_Name = "Module1"
Sub hw02_vbachallenge()

    ' holds ticker value
    Dim tickvalue As String
    
    
    ' variables to hold opening at beginning and close at end of year
    Dim open_by As Double
    Dim close_ey As Double
    Dim vol As Long
    Dim ws_number As Integer
        ws_number = ActiveWorkbook.Worksheets.Count
        MsgBox (ws_number & " worksheets")
    
    
    ' variable to dynamically set the range for each raw data table
    Dim range_raw As Long

    
    
    ' For loop to cycle through each worksheet
    For ws = 1 To ws_number
        
        ' Selects the worksheet before running analysis
        Worksheets(ws).Select
        
        ' structures the summary table forea each sheet. this must be reset for each sheet
        Dim sumtable As Integer
        sumtable = 2
        
        ' sets the range for the raw data tables
        range_raw = Cells(Rows.Count, 1).End(xlUp).Row
        MsgBox ("Raw Table " & ws & ": " & range_raw & " rows")
        
    
    
        ' For loop to search through the raw data table
        For i = 2 To range_raw
        
            ' Safety Stop to end the the loop if the end of the table is reached
            If Cells(i, 1).Value = "" Then
                MsgBox ("No more tickers")
                Exit For
            
            ' skips entry if either open or close value is 0
            ElseIf Cells(i, 3).Value = 0 Or Cells(i, 6).Value = 0 Then
                GoTo NextIteration
            
            ' identifies the beginning-of-year value for each ticker
            ElseIf Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
                tickvalue = Cells(i, 1).Value
                open_by = Cells(i, 3).Value
            
            ' identifies the end-of-year value for each ticker
            ElseIf Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                close_ey = Cells(i, 6).Value
                vol = Cells(i, 7).Value
                
                ' writes to the summary table for each ticker
                Cells(sumtable, 10).Value = tickvalue
                Cells(sumtable, 11).Value = close_ey - open_by
                Cells(sumtable, 12).Value = (close_ey - open_by) / open_by
                Cells(sumtable, 13).Value = vol

                ' adds a row for each new ticker
                sumtable = sumtable + 1
            End If
            
NextIteration:
        Next i
        
        
        
        ' variable to dynamically set the range for each summary table
        Dim range_summary As Long
            range_summary = Cells(Rows.Count, 10).End(xlUp).Row
            MsgBox ("Summary Table " & ws & ": " & range_summary & " rows")
        
        'variables to store the greatest inc, dec, & vol for the given range
        Dim g_inc As Double
        Dim g_dec As Double
        Dim g_vol As Long
        Dim g_inc_tick As String
        Dim g_dec_tick As String
        Dim g_vol_tick As String
        
        g_inc = Cells(2, 12).Value
        g_dec = Cells(2, 12).Value
        g_vol = Cells(2, 13).Value
        

        ' for loop to assign conditional formatting to the change in $ column
        For j = 2 To range_summary
            
            'conditional formatting
            If Cells(j, 10).Value = "" Then
                MsgBox ("conditional format complete")
                Exit For
            ElseIf Cells(j, 11).Value > 0 Then
                Cells(j, 11).Interior.ColorIndex = 4
            Else
                Cells(j, 11).Interior.ColorIndex = 3
            End If
            
            ' storing the greatest increase to g_inc
            If Cells(j, 12).Value > g_inc Then
                g_inc = Cells(j, 12).Value
                g_inc_tick = Cells(j, 10).Value
            End If
            
            ' storing the greatest decrease to g_dec
            If Cells(j, 12).Value < g_dec Then
                g_dec = Cells(j, 12).Value
                g_dec_tick = Cells(j, 10).Value
            End If
            
            ' storing the greatest vol to g_inc
            If Cells(j, 13).Value > g_vol Then
                g_vol = Cells(j, 13).Value
                g_vol_tick = Cells(j, 10).Value
            End If
            
        Next j
        
        
        
        ' writes the greatest inc, dec, and vol to below cells
        Range("o2").Value = "Greatest % Increase"
        Range("o3").Value = "Greatest % Decrease"
        Range("o4").Value = "Greatest Volume"
        Range("p2").Value = g_inc_tick
        Range("p3").Value = g_dec_tick
        Range("p4").Value = g_vol_tick
        Range("q2").Value = g_inc
        Range("q3").Value = g_dec
        Range("q4").Value = g_vol
        
        
        
    Next ws



End Sub

Sub resetbutton()
    
    Dim ws_number As Integer
       ws_number = ActiveWorkbook.Worksheets.Count
       MsgBox (ws_number & " worksheets")
       
    For ws = 1 To ws_number
        Sheets(ws).Range("j2:r5000").Clear
    Next ws
    
End Sub


