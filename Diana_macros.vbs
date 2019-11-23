Attribute VB_Name = "Module1"

Sub solution()
    
    
    'create variables needed to fill the table
    Dim active_ticker As String
    Dim first_open As Double
    Dim last_close As Double
    Dim sum_volume As LongLong
    Dim num_rows_sheet As Long
    Dim active_summary_row As Integer

    ' for loop to go through all the worksheets
    
    For Each ws In Worksheets
    
        ws.Select
        
        'checks number of cells
        num_rows_sheet = Cells(Rows.Count, 1).End(xlUp).Row

        
        'creates the summary table
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        
        'formats the cells of the summary table
        Range("J2:j" & num_rows_sheet).NumberFormat = "0.00"
        Range("K2:K" & num_rows_sheet).NumberFormat = "0.00%"
    
        sum_volume = 0
        active_summary_row = 2
        
        For i = 2 To num_rows_sheet
            
           'set values of variables for the first row of each sheet
            If i = 2 Then
                first_open = Cells(i, 3).Value
                active_ticker = Cells(i, 1).Value
                
            End If
                
            'check if next line has the same ticker
            
            If Cells(i, 1) = Cells(i + 1, 1) Then
                sum_volume = sum_volume + Cells(i, 7).Value
            
            Else
                'finish setting values for the variables of current ticker
                sum_volume = sum_volume + Cells(i, 7).Value
                last_close = Cells(i, 6).Value
            
                'save all the variables values in the summary table
                Range("I" & active_summary_row).Value = active_ticker
                
                Range("J" & active_summary_row).Value = last_close - first_open
                    'format the cells based on the value
                    If Range("j" & active_summary_row).Value >= 0 Then
                        Range("j" & active_summary_row).Interior.ColorIndex = 4
                    Else
                        Range("j" & active_summary_row).Interior.ColorIndex = 3
                    End If
                    
                'avoid the division by 0 case
                If first_open = 0 Then
                    Range("K" & active_summary_row).Value = Null
                    Else
                        Range("K" & active_summary_row).Value = (last_close - first_open) / first_open
                End If
                
                Range("L" & active_summary_row).Value = sum_volume
                
                'update all the variables for the next ticker
                active_summary_row = active_summary_row + 1
                sum_volume = 0
                first_open = Cells(i + 1, 3).Value
                active_ticker = Cells(i + 1, 1).Value
                last_close = 0
            
            End If
        
            Next i
            
        Next ws
        
End Sub


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

