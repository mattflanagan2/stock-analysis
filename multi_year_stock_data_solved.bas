Attribute VB_Name = "Module1"
Sub stocks()
    Dim volume As Double
    Dim ticker As String
    Dim open_price As Double
    Dim close_price As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim summary_table_row As Double
    Dim ws As Worksheet
    
'======================================================
    For Each ws In Worksheets
    ws.Activate
    
    
    
        summary_table_row = 2
        volume = 0
        previous_i = 1
     
        EndRow = Cells(Rows.Count, "A").End(xlUp).Row
    
        Cells(1, 9).Value = "Ticker"             'Creating headers for the summary table
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Volume"
    
    
        For i = 2 To EndRow                              'starting the loop to read each ticker
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                ticker = Cells(i, 1).Value
            
                previous_i = previous_i + 1 'why does this make it work
            
                volume = volume + Cells(i, 7).Value 'finding volume
            
                year_open = Cells(previous_i, 3).Value 'defining our open year value
                year_close = Cells(i, 6).Value
            
            
                yearly_change = year_close - year_open
            
                If year_open = 0 Then
                
                    percent_change = year_close
                
                Else
    
                    percent_change = yearly_change / year_open
                
                End If
            
                Range("I" & summary_table_row).Value = ticker 'outputting current ticker into summary table
                Range("L" & summary_table_row).Value = volume 'outputting current volume into summary table
                Range("J" & summary_table_row).Value = yearly_change
                Range("K" & summary_table_row).Value = percent_change
                Range("K" & summary_table_row).NumberFormat = "0.00%" 'changing K column to % format
            
                summary_table_row = summary_table_row + 1 'adding a value of 1 to summary table row so that it creates a next line
            
                volume = 0 'resetting everthing before it moves to the next i
                yearly_change = 0
                percent_change = 0
                previous_i = i
            Else
                volume = volume + Cells(i, 7).Value   'do I need these???
                yearly_change = year_close - year_open
        
            End If
        Next i


'======================================================

        jEndRow = Cells(Rows.Count, "J").End(xlUp).Row 'counting the J column end row
        
            For j = 2 To jEndRow        'loop to determine color
                If Cells(j, 10) > 0 Then    'if the cells is greater than 0....
                    Cells(j, 10).Interior.ColorIndex = 4 'it will make the cell green
                Else                                                    'if not.....
                    Cells(j, 10).Interior.ColorIndex = 3 'it will make the cell red
                End If
            Next j

'=======================================================


        Cells(2, 15).Value = "Greatest Percent Increase"
        Cells(3, 15).Value = "Greatest Percent Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        Range("O1").EntireColumn.AutoFit
        Range("Q1").EntireColumn.AutoFit
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"

        kEndRow = Cells(Rows.Count, "K").End(xlUp).Row
    
        increase = 0
        decrease = 0
        greatest = 0
        
            For k = 3 To kEndRow
                last_k = k - 1
            
                current_k = Cells(k, 11).Value
                previous_k = Cells(last_k, 11).Value
            
            
                current_volume = Cells(k, 12).Value
                previous_volume = Cells(last_k, 12).Value
            
            
            
                If increase > current_k And increase > previous_k Then
                    increase = increase
                
                ElseIf current_k > increase And current_k > previous_k Then
                    increase = current_k
                    increase_name = Cells(k, 9).Value
                ElseIf previous_k > increase And previous_k > current_k Then
                    increase = previous_k
                    increase_name = Cells(last_k, 9).Value
                End If
            
                If decrease < current_k And decrease < previous_k Then
                    decrease = decrease
                
                ElseIf current_k < increase And current_k < previous_k Then
                    decrease = current_k
                    decrease_name = Cells(k, 9).Value
                ElseIf previous_k < increase And previous_k < current_k Then
                    decrease = previous_k
                    decrease_name = Cells(last_k, 9).Value
                End If
                
                If greatest > current_volume And greatest > previous_volume Then
                    greatest = greatest
                ElseIf current_volume > greatest And current_volume > previou_volume Then
                    greatest = current_volume
                    greatest_name = Cells(k, 9).Value
                ElseIf previous_volume > greatest And previous_volume > current_volume Then
                    greatest = previous_volume
                    greatest_name = Cells(last_k, 9).Value
                End If
                
                
            Next k
            
        Cells(2, 16).Value = increase_name
        Cells(2, 17).Value = increase
        Cells(2, 17).NumberFormat = "0.00%"
        
        Cells(3, 16).Value = decrease_name
        Cells(3, 17).Value = decrease
        Cells(3, 17).NumberFormat = "0.00%"
        
        Cells(4, 16).Value = greatest_name
        Cells(4, 17).Value = greatest
        Range("L1").EntireColumn.AutoFit
        Range("L1").EntireColumn.NumberFormat = "0,000"
        Cells(4, 17).NumberFormat = "0,000"
        
        
    Next ws
    
End Sub







