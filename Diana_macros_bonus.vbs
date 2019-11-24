Attribute VB_Name = "Module2"
Sub bonus()

    'create variables needed to fill the table
    Dim g_increase_ticker As String
    Dim g_increase_value As Double
    Dim g_decrease_ticker As String
    Dim g_decrease_value As Double
    Dim g_volume_ticker As String
    Dim g_volume_value As LongLong
    Dim num_rows_summary As Long

    For Each ws In Worksheets
    
        ws.Select
        
        'checks number of cells
        num_rows_summary = Cells(Rows.Count, 9).End(xlUp).Row

        'creates the summary table
        Range("O2").Value = "Greatest % increase"
        Range("O3").Value = "Greatest % decrease"
        Range("O4").Value = "Greatest Total Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        
        'formats the cells of the summary table
        Range("Q2:Q3").NumberFormat = "0.00%"
        
        'set the first row of data as the temp "winner"
        g_increase_ticker = Range("I2")
        g_increase_value = Range("K2")
        g_decrease_ticker = Range("I2")
        g_decrease_value = Range("K2")
        g_volume_ticker = Range("I2")
        g_volume_value = Range("L2")
        
        For i = 3 To num_rows_summary
            
            'Check greatest increase
            If Cells(i, 11) > g_increase_value Then
                g_increase_ticker = Cells(i, 9)
                g_increase_value = Cells(i, 11)
            End If
            
            'Check greatest decrease
            If Cells(i, 11) < g_decrease_value Then
                g_decrease_ticker = Cells(i, 9)
                g_decrease_value = Cells(i, 11)
            End If
            
            'Check greatest volume
            If Cells(i, 12) > g_volume_value Then
                g_volume_ticker = Cells(i, 9)
                g_volume_value = Cells(i, 12)
            End If
        
        Next i
    
        'write down in the summary table
        Range("P2").Value = g_increase_ticker
        Range("Q2").Value = g_increase_value
        Range("P3").Value = g_decrease_ticker
        Range("Q3").Value = g_decrease_value
        Range("P4").Value = g_volume_ticker
        Range("Q4").Value = g_volume_value
        
    Next ws

End Sub


