Attribute VB_Name = "Module1"
Sub Updtd_Stock_Data()
    
        Dim Ticker_A As String
    
        Dim Total_Stock_Volume As Double
        Total_Stock_Volume = 0

        Dim Summary_Row As Integer
        Summary_Row = 2
        
        Dim Opening_Year As Double
        Opening_Year = Cells(2, 3).Value
        
        Dim Closing_Year As Double
        Dim Yearly_Change As Double
        Dim Percent_Change As Double

        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"

        lastrow = Cells(Rows.Count, 1).End(xlUp).Row

        For i = 2 To lastrow

            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        

              Ticker_A = Cells(i, 1).Value
              Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
              Range("I" & Summary_Row).Value = Ticker_A
              Range("L" & Summary_Row).Value = Total_Stock_Volume

              Closing_Year = Cells(i, 6).Value

              Yearly_Change = (Closing_Year - Opening_Year)
              
              Range("J" & Summary_Row).Value = Yearly_Change

                If (Opening_Year = 0) Then

                    Percent_Change = 0

                Else
                    
                    Percent_Change = Yearly_Change / Opening_Year
                
                End If

              Range("K" & Summary_Row).Value = Percent_Change
              Range("K" & Summary_Row).NumberFormat = "0.00%"
   
              Summary_Row = Summary_Row + 1

              Total_Stock_Volume = 0

              Opening_Year = Cells(i + 1, 3)
            
            Else
              
              Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value

            End If
        
        Next i

    lastrow_summary_table = Cells(Rows.Count, 9).End(xlUp).Row
     
    For i = 2 To lastrow_summary_table
            If Cells(i, 10).Value > 0 Then
                Cells(i, 10).Interior.ColorIndex = 4
            Else
                Cells(i, 10).Interior.ColorIndex = 3
            End If
    Next i

End Sub
