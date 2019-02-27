Sub Stock_Data_Hard()
    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    
    For Each ws In Worksheets
    ' Activate each worksheet in the Spreadsheet
    ws.Activate
    
    
    ' --------------------------------------------
    ' WORKING STORAGE SECTION
    ' --------------------------------------------
    
    
    ' Determine the Last Row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Set an initial variable for holding the Stock Name
    Dim Stock_Name As String
    
    ' Set an initial variable for holding the Yearly Change on Stock
    Dim Yearly_Change As Double
        Yearly_Change = 0
    
    ' Set an initial variable for holding the Percentage Change on Stock
    Dim Percentage_Change
    Dim Hold_Percentage_Change As Double
        
    ' Set an initial variable for holding the Total Volume of Stock trad
    Dim Stock_Total As Double
        Stock_Total = 0
    
    ' Set an initial variable for holding the Opening Price on Stock
    Dim Opening_Price As Double
        Opening_Price = Cells(2, 3).Value
        
    ' Set an initial variable for holding the Closing Price on Stock
    Dim Closing_Price As Double
        Closing_Price = 0
        
    ' Set an initial variable for holding the Stock Ticker with Greatest % increase
    Dim Greatest_Percent_Increase As Double
        Greatest_Percent_Increase = 0
   
    ' Set an initial variable for holding the Stock Ticker with Greatest % Decrease
    Dim Greatest_Percent_Decrease As Double
        Greatest_Percent_Decrease = 0
    
    ' Set an initial variable for holding the Stock Ticker with Greatest total volume
    Dim Greatest_Total_Volume As Double
        Greatest_Total_Volume = 0
        
    ' Hold Stock names to be displyed on the Summary Table
    Dim Greatest_Percent_Increase_Stock As String
    Dim Greatest_Percent_Decrease_Stock As String
    Dim Greatest_Total_Volume_Stock As String
    Dim Greatest_Table_Row As Integer
        Greatest_Table_Row = 2
   
    
    ' Keep track of the location for each Stock in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    '----------------------------------------------
    ' PROCEDURE DIVISION
    '----------------------------------------------
    
    ' Loop through all stock Ticker for the year
    For i = 2 To LastRow
    
    
    ' Check if we are still within the same Stock, if it is not...
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            ' Set the Stock name
            Stock_Name = Cells(i, 1).Value
            
            ' Set the Stock Closing Price
            Closing_Price = Cells(i, 6).Value
            
            Yearly_Change = Closing_Price - Opening_Price
            
            'Calculate Percentage of Change in Stock Price
            If Opening_Price <> 0 Then
               Hold_Percentage_Change = Round((Yearly_Change / Opening_Price), 4)
               Percentage_Change = Hold_Percentage_Change
            Else
               Percentage_Change = 0
            End If

            ' Add to the Stock Total
            Stock_Total = Stock_Total + Cells(i, 7).Value

           
            ' Print the Stock Name in the Summary Table
            Range("i" & Summary_Table_Row).Value = Stock_Name
            
            ' Print the Yearly Change in the Summary Table
            Range("j" & Summary_Table_Row).Value = Yearly_Change
            
            If Yearly_Change > 0 Then
                Range("j" & Summary_Table_Row).Interior.ColorIndex = 3
            Else
                Range("j" & Summary_Table_Row).Interior.ColorIndex = 7
            End If
            
            'Print the Percentage of Change in Stock Price
            Range("k" & Summary_Table_Row).Value = FormatPercent(Percentage_Change)
'            Range("k" & Summary_Table_Row).Value = Percentage_Change
            
            ' Print the total stock volume to the Summary Table
            Range("l" & Summary_Table_Row).Value = Stock_Total

            '------------------------------------------------
            ' PRINT SUMMARY TABLE
            '------------------------------------------------
            
            If Summary_Table_Row = 2 Then
                
                Greatest_Percent_Increase = Hold_Percentage_Change
                Greatest_Percent_Increase_Stock = Stock_Name
                
                Greatest_Percent_Decrease = Hold_Percentage_Change
                Greatest_Percent_Decrease_Stock = Stock_Name
                
                Greatest_Total_Volume = Stock_Total
                Greatest_Total_Volume_Stock = Stock_Name
                
               'MsgBox ("Summary_Table_Row :" + Str(Summary_Table_Row) + " Greatest_Percent_Increase :" + Str(Greatest_Percent_Increase) + " Hold_Percentage_Change :" + Str(Hold_Percentage_Change) + " Greatest_Percent_Decrease :" + Str(Greatest_Percent_Decrease) + "Greatest_Total_Volume_Stock : " + Greatest_Total_Volume_Stock)
                
            Else
                If Greatest_Total_Volume_Stock <> Stock_Name Then
                'MsgBox ("before " + "Summary_Table_Row :" + Str(Summary_Table_Row) + " Old Great % increase :" + Str(Greatest_Percent_Increase) + " Great % increase :" + Str(Hold_Percentage_Change) + " Gret % Decrease :" + Str(Greatest_Percent_Decrease) + "Summary_Table_Row" + Greatest_Total_Volume_Stock)
 
                    If Round(Hold_Percentage_Change, 4) > Round(Greatest_Percent_Increase, 4) Then
                       'MsgBox ("Increase :" + "Summary_Table_Row :" + Str(Summary_Table_Row) + " Old Great % increase :" + Str(Greatest_Percent_Increase) + " New % change :" + Str(Hold_Percentage_Change) + "Stock Name" + Greatest_Total_Volume_Stock)
                        Greatest_Percent_Increase = Hold_Percentage_Change
                        Greatest_Percent_Increase_Stock = Stock_Name
                   
                    End If
                    If Round(Hold_Percentage_Change, 4) < Round(Greatest_Percent_Decrease, 4) Then
                       'MsgBox (" Decrease :" + "Summary_Table_Row :" + Str(Summary_Table_Row) + " Old Great % Decrease :" + Str(Greatest_Percent_Decrease) + " Great % change :" + Str(Hold_Percentage_Change) + " Stock Name :" + Str(Greatest_Percent_Decrease) + "Stock Name :" + Greatest_Total_Volume_Stock)
                        Greatest_Percent_Decrease = Hold_Percentage_Change
                        Greatest_Percent_Decrease_Stock = Stock_Name

                    
                    End If
                    If Greatest_Total_Volume < Stock_Total Then
                        Greatest_Total_Volume = Stock_Total
                        Greatest_Total_Volume_Stock = Stock_Name
                       'MsgBox ("stock total :" + "Summary_Table_Row :" + Str(Summary_Table_Row) + " Old Great % increase :" + Str(Greatest_Percent_Increase) + " Great % increase :" + Str(Hold_Percentage_Change) + " Gret % Decrease :" + Str(Greatest_Percent_Decrease) + "Summary_Table_Row" + Greatest_Total_Volume_Stock)
                    
                    End If
                End If
            End If
            

            '---------------------------------------------
            ' RESET VARIABLES AFTER PRINTING SUMMARY TABLE
            '---------------------------------------------

            ' Reset the Stock Total
            Stock_Total = 0
            
            ' Reset the Stock Opening Price
            Opening_Price = Cells(i + 1, 3).Value
            
            ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
  

        ' If the cell immediately following a row is the same stock...
        Else

            ' Add to the Stock Total
            Stock_Total = Stock_Total + Cells(i, 7).Value
            
 
        End If
   
    Next i
    
            '-----------------------------
            ' PRINT SECOND SUMMARY TABLE
            '-----------------------------
            
             ' Print the total Greatest Percent Increase to the second Summary Table
            Range("p" & Greatest_Table_Row).Value = "Greatest Percent Increase"
            Range("q" & Greatest_Table_Row).Value = Greatest_Percent_Increase_Stock
            Range("r" & Greatest_Table_Row).Value = FormatPercent(Greatest_Percent_Increase)
            
            
            ' Print the total Greatest Percent Decrease to the second Summary Table
            Greatest_Table_Row = Greatest_Table_Row + 1
            Range("p" & Greatest_Table_Row).Value = "Greatest Percent Decrease"
            Range("q" & Greatest_Table_Row).Value = Greatest_Percent_Decrease_Stock
            Range("r" & Greatest_Table_Row).Value = FormatPercent(Greatest_Percent_Decrease)
            
            ' Print the total Gretatest Total Volume to the second Summary Table
              Greatest_Table_Row = Greatest_Table_Row + 1

            Range("p" & Greatest_Table_Row).Value = "Greatest Total Volume"
            Range("q" & Greatest_Table_Row).Value = Greatest_Total_Volume_Stock
            Range("r" & Greatest_Table_Row).Value = Greatest_Total_Volume


            Range("i" & 1).Value = "Ticker"
            Range("j" & 1).Value = "Yearly Change"
            Range("k" & 1).Value = "Percent Change"
            Range("l" & 1).Value = "Total Stock Volume"
            Range("q" & 1).Value = "Ticker"
            Range("r" & 1).Value = "Value"
    
    Next ws
    MsgBox ("Fixes Complete")
   
 
End Sub









