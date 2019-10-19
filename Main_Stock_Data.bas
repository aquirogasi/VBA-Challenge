Attribute VB_Name = "Main_Stock_Data"
Sub Stock_Data()

Dim i As Long
Dim Summary_Table_2 As Integer
Dim Ticker As String
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Total_Stock_Volume As LongLong
Dim Open_Stock As Double
Dim End_Stock As Double

Dim AddressOfMax_Percentage As Range
Dim AddressOfMin_Percentage As Range
Dim AddressOfMax_Volume As Range

Dim Max_Percent_Increase As Double
Dim Max_Percent_Decrease As Double
Dim Max_Total_Volume As LongLong

Dim Max_Percent_Increase_Ticker As String
Dim Max_Percent_Decrease_Ticker As String
Dim Max_Total_Volume_Ticker As String


'Loop through all Sheets

For Each ws In Worksheets
    
    'Variable Initializations Inside For ws Loop / Outside For i Loop
    
    Total_Stock_Volume = 0

    Open_Stock = ws.Range("C2").Value
    
    Summary_Table_2 = 2
    
    Last_Row = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    'Loop through all the Stocks without including header until last non-blank Row
    
    For i = 2 To Last_Row
    
    
        ' Check if we are still within the same stock, if it is not...
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        'Set End_Stock Value
    
        End_Stock = ws.Cells(i, 6).Value
    
        'Calculate Yearly Change: closing price at end of the year - opening price at beginning of the year
    
        Yearly_Change = End_Stock - Open_Stock
    
            'Calculate Percent Change when Open Stock Value is greatest than zero
        
            If Open_Stock > 0 Then
    
                Percent_Change = Yearly_Change / Open_Stock
                
            'Calculate Percent Change when Open Stock Value is equal to zero / Avoid division by zero error
            
            ElseIf Open_Stock = 0 Then
            
                Percent_Change = Yearly_Change
                
            End If
                
    
        'Calculate Total Stock Volume
    
        Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
    
        'Set Ticket Symbol
    
        Ticker = ws.Cells(i, 1).Value
    
        'Insert Summary #1 Headers
    
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
    
    
    
        'Conditional Color based on Yearly Change's values
    
            If Yearly_Change < 0 Then
                ws.Cells(Summary_Table_2, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(Summary_Table_2, 10).Interior.ColorIndex = 4
            End If
        
        
        'Insert calculated values into their corresponding positions
    
        ws.Cells(Summary_Table_2, 9).Value = Ticker
        ws.Cells(Summary_Table_2, 10).Value = Yearly_Change
        ws.Cells(Summary_Table_2, 11).Value = FormatPercent(Percent_Change, 2)
        ws.Cells(Summary_Table_2, 12).Value = Total_Stock_Volume
    
        'Increasing the summary table counter by 1 in order to move through the rows
    
        Summary_Table_2 = Summary_Table_2 + 1
    
        'Updating the Open Stock value
    
        Open_Stock = ws.Cells(i + 1, 3).Value
    
        ' Reset the Total Stock Volume
        
        Total_Stock_Volume = 0
        
        
        Else
    
        ' Add to the Total Stock Volume
        
        Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
    
        End If
    

    'End For i Loop
    Next i


'Set Adress of Greatest Percentage Increase by calling AddressofMax function defined in Module: AddressOfMax_Function

Set AddressOfMax_Percentage = AddressOfMax(ws.Range("K2", ws.Range("K2").End(xlDown).Offset(-1, 0)))

'Set Adress of Greatest Percentage Decrease by calling AddressofMin function defined in Module: AddressOfMin_Function

Set AddressOfMin_Percentage = AddressOfMin(ws.Range("K2", ws.Range("K2").End(xlDown).Offset(-1, 0)))


'Set Adress of Greatest Percentage Increase by calling AddressofMax function defined in Module: AddressOfMax_Function

Set AddressOfMax_Volume = AddressOfMax(ws.Range("L2", ws.Range("L2").End(xlDown).Offset(-1, 0)))



'Find the value of Greatest Percentage Increase and the corresponding Ticker

Max_Percent_Increase = AddressOfMax_Percentage.Value

Max_Percent_Increase_Ticker = AddressOfMax_Percentage.Offset(0, -2).Value



'Find the value of Greatest Percentage Decrease and the corresponding Ticker

Max_Percent_Decrease = AddressOfMin_Percentage.Value

Max_Percent_Decrease_Ticker = AddressOfMin_Percentage.Offset(0, -2).Value


'Find the value of Greatest Total Volume and the corresponding Ticker

Max_Total_Volume = AddressOfMax_Volume.Value

Max_Total_Volume_Ticker = AddressOfMax_Volume.Offset(0, -3)




'Insert Summary #2 Headers

ws.Cells(1, 16) = "Ticker"
ws.Cells(1, 17) = "Value"

ws.Cells(2, 15) = "Greatest % Increase"
ws.Cells(3, 15) = "Greatest % Decrease"
ws.Cells(4, 15) = "Greatest Total Volume"

'Insert Calculated Max and Min Values

ws.Cells(2, 17) = FormatPercent(Max_Percent_Increase, 2)
ws.Cells(3, 17) = FormatPercent(Max_Percent_Decrease, 2)
ws.Cells(4, 17) = Max_Total_Volume

ws.Cells(2, 16) = Max_Percent_Increase_Ticker
ws.Cells(3, 16) = Max_Percent_Decrease_Ticker
ws.Cells(4, 16) = Max_Total_Volume_Ticker

'Autofit Columns to Display data correctly

ws.Columns("I:O").AutoFit

'Format Summary Table # 2 with Colors and Borders

ws.Range("P1:Q1").Font.Bold = True
ws.Range("O2:O4").Font.Bold = True
ws.Range("P1:Q1").Interior.ColorIndex = 22
ws.Range("O2:Q2").Interior.ColorIndex = 37
ws.Range("O3:Q3").Interior.ColorIndex = 34
ws.Range("O4:Q4").Interior.ColorIndex = 37
ws.Range("O2:Q4").Borders.LineStyle = xlContinuous
ws.Range("P1:Q1").Borders.LineStyle = xlContinuous


'End For ws Loop
Next ws

End Sub

