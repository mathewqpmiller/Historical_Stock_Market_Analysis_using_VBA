Attribute VB_Name = "Mathew_M_VBA_Module"
Sub stock_market_analysis()

    'Define Summary Table Variables
    Dim i, j As Integer
    Dim Ticker_Symbol As String
    Dim Stock_Volume As Double
    Dim Price_at_Open As Double
    Dim Price_at_Close As Double
    Dim Price_Change As Double
    Dim Percentage_Price_Change As Double
    'Set initial values for Summary Table variables
    Ticker_Symbol = " "
    Stock_Volume = 0
    Price_at_Open = 0
    Price_at_Close = 0
    Price_Change = 0
    Percentage_Price_Change = 0
    'Define Advanced Summary Table Variables
    Dim Top_Percentage_Increase_Ticker As String
    Dim Top_Percentage_Decrease_Ticker As String
    Dim Top_Total_Volume_Ticker As String
    Top_Percentage_Increase_Ticker = " "
    Top_Percentage_Decrease_Ticker = " "
    Top_Total_Volume_Ticker = " "
    Dim Top_Percentage_Increase As Double
    Dim Top_Percentage_Decrease As Double
    Dim Top_Total_Volume As Double
    Top_Percentage_Increase = 0
    Top_Percentage_Decrease = 0
    Top_Total_Volume = 0
    
    'Define Stock Analysis Summary Table and Headers
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    Range("J1").Value = "Stock-Ticker"
    Range("K1").Value = "Annual Price Change"
    Range("L1").Value = "Annual Percent Change"
    Range("M1").Value = "Total Stock Volume"
    Range("J1:M1").Columns.AutoFit
    'Define Stock Analysis Advanced Summary Table and Headers/Titles
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Stock-Ticker"
    Range("Q1").Value = "Value"
    Range("O1:O4").Columns.AutoFit
    Range("P1:P4").Columns.AutoFit
    Range("Q1:Q4").Columns.AutoFit
    
    'Establish Price Open starting point befor beginning loop
    Price_at_Open = Cells(2, 3).Value
    
    'Define For loop of rows and columns through to the last row with data
    Lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To Lastrow
        
        'If statement sorting and grouping like Ticker Symbols for Summary Table
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            'Where ticker symbol is located on spreadsheet
            Ticker_Symbol = Cells(i, 1).Value
             
            'Where the Close Price is located in spreadsheet
            'Referenced website: https://www.reddit.com/r/vba/comments/9ksy0f/need_help_with_looping_through_stock_ticker_data/
            Price_at_Close = Cells(i, 6).Value
        
            'How to calculate the Price Change value
            Price_Change = Price_at_Close - Price_at_Open
            
                'How to calculate the Price Change Percentage value
                'Referenced website:https://freesoft.dev/program/163047389
                If Price_at_Open <> 0 Then
                Percentage_Price_Change = (Price_Change / Price_at_Open) * 100
                End If
            
            'How to calculate the Total Stock Volume value
            Stock_Volume = Stock_Volume + Cells(i, 7).Value

            'Where to place the grouped Ticker Symbol string in Summary Table
            Range("J" & Summary_Table_Row).Value = Ticker_Symbol
            
            'Where to place the Price Change value in Summary Table
            Range("K" & Summary_Table_Row).Value = Price_Change
            
                'Cell formatting of Price Change column based on stock gain or loss
                If (Price_Change > 0) Then
                    'Fill with green for increase in price
                    Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
                ElseIf (Price_Change <= 0) Then
                    'Fill with red for decrease in price
                    Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
                End If
            
            'Where to place Percentage Change Value in summary table
            'Referenced website: https://freesoft.dev/program/163047389
            Range("L" & Summary_Table_Row).Value = (CStr(Percentage_Price_Change) & "%")
        
            'Where to place Total Stock Volume in Summary Table
            Range("M" & Summary_Table_Row).Value = Stock_Volume
        
            'Where to move once all grouped Ticker Symbol information has been filled into Summary Table
            Summary_Table_Row = Summary_Table_Row + 1
            
            Price_Change = 0
            
            Price_at_Close = 0
            
            Price_at_Open = Cells(i + 1, 3).Value
            
                'Calculate and Parse the Summary Table strings to the Advanced Summary Table strings
                'Referenced website: https://freesoft.dev/program/163047389
                If (Percentage_Price_Change > Top_Percentage_Increase) Then
                   Top_Percentage_Increase = Percentage_Price_Change
                    Top_Percentage_Increase_Ticker = Ticker_Symbol
                ElseIf (Percentage_Price_Change < Top_Percentage_Decrease) Then
                    Top_Percentage_Decrease = Percentage_Price_Change
                    Top_Percentage_Decrease_Ticker = Ticker_Symbol
                End If
                       
                If (Stock_Volume > Top_Total_Volume) Then
                    Top_Total_Volume = Stock_Volume
                    Top_Total_Volume_Ticker = Ticker_Symbol
                End If
                
                Range("P2").Value = Top_Percentage_Increase_Ticker
                Range("Q2").Value = (CStr(Top_Percentage_Increase) & "%")
                Range("P3").Value = Top_Percentage_Decrease_Ticker
                Range("Q3").Value = (CStr(Top_Percentage_Decrease) & "%")
                Range("P4").Value = Top_Total_Volume_Ticker
                Range("Q4").Value = Top_Total_Volume
                
                'Reset Advanced Summary Table counters
                Percentage_Price_Change = 0
            'Reset Summary Table counters
            Stock_Volume = 0
          
            
         Else
            
            'If no change then continue to add information to running total
            Stock_Volume = Stock_Volume + Cells(i, 7).Value
            
        End If
       
    Next i
        
End Sub

