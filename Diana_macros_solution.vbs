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


