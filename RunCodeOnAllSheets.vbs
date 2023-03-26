Attribute VB_Name = "RunCodeOnAllSheets"
Sub RunCodeOnAllTabs()

    'Declare Worksheet as a variable
    Dim Current As Worksheet
    
    'Loop and run for each worksheet tab
    For Each Current In Worksheets
    'Activate new worksheet in workbook
    Current.Activate

        ' Define variables
        Dim Ticker As String 'stores current Ticker value from column A
        Dim row_Variable As Long 'Stores current row index
        Dim counter_Variable As Long 'Stores current index, separate from row
        Dim OpenValue_Variable As Double 'stores current open from cloumn C
        Dim CloseValue_Variable As Double 'stores current closevalue from column F
        Dim YearlyChange As Double
        Dim totalStockVolume_Variable As Double 'stores the current <vol> value summation from column G for the same Ticker value
        Dim lastRow_Variable As Long 'find last row
        Dim i As Long
        Dim rng As Range
        
    'Initialize Variables
    row_Variable = 2 ' [row_Variable] is initialized with the value of 2.
    counter_Variable = 2 '[counter_Variable] is initialized with the value of 2.
    lastRow_Variable = Range("A" & Rows.Count).End(xlUp).Row ' Last row with data from column A
    
        ' Output Range
        Set rng = Range("H1:L" & lastRow_Variable)
        ' Clears celss in output range
        rng.ClearContents
    
        ' Creating the header for columns
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
            
        '-------------------
        ' This [For] Loop iterates from row 2 to row [lastRow_Variable]
         For row_Variable = 2 To lastRow_Variable
         
                ' Ticker is assigned the value of the cell in column A at row [row_Variable]
                Ticker = Range("A" & row_Variable).Value
                
                ' [openValue_Variable] is assigned the value of the cell in column C at row [row_Variable]
                OpenValue_Variable = Range("C" & row_Variable).Value
                
                ' Assigns the Ticker value from column A at row to column I at row [counter_Variable]
                Range("I" & counter_Variable).Value = Range("A" & row_Variable).Value
                
                ' Total Stock is initialized with the value of the cell in column G and row [row_Variable]
                totalStockVolume_Variable = Range("G" & row_Variable).Value
                
                ' This [While] loop iterates as long as the value of the cell in the next row is the same as the current row
                While Range("A" & row_Variable).Value = Range("A" & row_Variable + 1).Value
                    
                    ' Increases the value of [row_Variable] by 1
                    row_Variable = row_Variable + 1
           
                    ' Adds the value of the current cell in column G to the value of [totalStockVolume_Variable]
                    totalStockVolume_Variable = totalStockVolume_Variable + Range("G" & row_Variable).Value
                    
            Wend 'close while loop
           '-------------------
           
                ' [openValue_Variable] is assigned the value of Range(C & [row_Variable]).Value
                CloseValue_Variable = Range("F" & row_Variable).Value
       
                Range("J" & counter_Variable).Value = CloseValue_Variable - OpenValue_Variable
                
                ' Calculate Percent change
                Range("K" & counter_Variable).Value = FormatPercent((CloseValue_Variable - OpenValue_Variable) / OpenValue_Variable)
           
           '-----------------------
           ' Color coding, positive green and negative red
           
           ' Checks to see if the percent change is negative, condition to red
           If Range("K" & counter_Variable).Value < 0 Then
                    Range("J" & counter_Variable).Interior.ColorIndex = 3
       
                    ' Checks to see if the percent change is positive, conditon to green
                    ElseIf Range("K" & counter_Variable).Value > 0 Then

                        Range("J" & counter_Variable).Interior.ColorIndex = 4

                    Else ' Runs the following code if the percent change is zero
               
                End If
                
                 Range("L" & counter_Variable).Value = totalStockVolume_Variable
                
                ' Advances the [counter_Variable] by 1. This ensures the value in the next iteration is recorded in the row index = [counter_Variable]
                counter_Variable = counter_Variable + 1
               
            Next row_Variable ' Ends the [For] Loop, increasing the value of [row_Variable] by 1
           
           
           ' Autofits column widths to data
            rng.EntireColumn.AutoFit
         
    Next Current
           
End Sub
