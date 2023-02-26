Attribute VB_Name = "Module1"
Sub Loop_Through_All_Years()

' This macro will run the Loop_Through_Stock_One_Year sub for all years available in the file

    ' Step 1: Define Variables

        Dim year As Integer
        year = 2018
        
    ' Step 2: Call macro for each year
        
        For year = 2018 To 2020
        
            Sheets("" & year & "").Select
            Call Loop_Through_Stock_One_Year
            
        Next year
        
End Sub


Sub Loop_Through_Stock_One_Year()

' This macro will run through all the stocks for a given year, and return a number of parameters

    ' Step 1: Create Headers
    
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
    
    ' Step 2: Define Variables
        
        Dim TickerSymbol As String
        Dim OpeningPrice As Double
        Dim ClosingPrice As Double
        Dim Volume As Double
        Dim i As Long
        Dim NumberTicker As Integer
        i = 2
        NumberTicker = 1
        
    ' Step 3.1: Loop through all tickers
         
        Do While Not (IsEmpty(Cells(i, 1)))
                
            TickerSymbol = Cells(i, 1).Value
            OpeningPrice = Cells(i, 3).Value
            ClosingPrice = 0
            Volume = 0
                
            ' Step 3.2: Loop through all values for one given ticker
                
            Do While Cells(i, 1).Value = TickerSymbol
                
                ClosingPrice = Cells(i, 6).Value
                Volume = Volume + Cells(i, 7).Value
                                     
                i = i + 1
                    
            Loop
                
            
            ' Step 3.3: Save Values for selected ticker
                
            NumberTicker = NumberTicker + 1
        
            Cells(NumberTicker, 9).Value = TickerSymbol
            Cells(NumberTicker, 10).Value = ClosingPrice - OpeningPrice
            Cells(NumberTicker, 11).Value = (ClosingPrice - OpeningPrice) / OpeningPrice
            Cells(NumberTicker, 12).Value = Volume
                
                
            ' Step 3.4: Format Values
            
            Cells(NumberTicker, 10).Select
            Selection.NumberFormat = "[$USD] #,##0.00"
            Cells(NumberTicker, 11).Select
            Selection.NumberFormat = "0.00%"
                
            ' Step 3.5: Include Conditional Formatting
                
                ' Step 3.5.1: Conditional Formatting for Yearly Change Column
                    
                        Cells(NumberTicker, 10).Select
                        Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
                            Formula1:="=0"
                        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
                        With Selection.FormatConditions(1).Interior
                            .PatternColorIndex = xlAutomatic
                            .Color = 255
                            .TintAndShade = 0
                        End With
                        Selection.FormatConditions(1).StopIfTrue = False
                        Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual _
                            , Formula1:="=0"
                        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
                        With Selection.FormatConditions(1).Interior
                            .PatternColorIndex = xlAutomatic
                            .Color = 5287936
                            .TintAndShade = 0
                        End With
                        Selection.FormatConditions(1).StopIfTrue = False
                    
                ' Step 3.5.2: Conditional Formatting for Percent Change column
                        
                        Cells(NumberTicker, 11).Select
                        Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
                            Formula1:="=0"
                        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
                        With Selection.FormatConditions(1).Interior
                            .PatternColorIndex = xlAutomatic
                            .Color = 255
                            .TintAndShade = 0
                        End With
                        Selection.FormatConditions(1).StopIfTrue = False
                        Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual _
                            , Formula1:="=0"
                        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
                        With Selection.FormatConditions(1).Interior
                            .PatternColorIndex = xlAutomatic
                            .Color = 5287936
                            .TintAndShade = 0
                        End With
                        Selection.FormatConditions(1).StopIfTrue = False
            
            Loop
            
            
    ' Step 4: Run sub Retrieve_Largest_Change
    
        Call Retrieve_Largest_Change
        
        
End Sub

Sub Retrieve_Largest_Change()

' This macro will run through the outcome of a given year, and retrieve the tickers with the greatest Percent Change, greatest Percent Decrease and greatest Total Volume

    ' Step 1: Create Headers
    
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"


    ' Step 2: Define Variables
        
        Dim TickerSymbolIncrease As String
        Dim TickerSymbolDecrease As String
        Dim TickerSymbolVolume As String
        Dim GreatestIncrease As Double
        Dim GreatestDecrease As Double
        Dim PercentChange As Double
        Dim GreatestVolume As Double
        Dim TotalVolume As Double
        Dim i As Integer
        
        i = 2
        TickerSymbolIncrease = Cells(i, 9).Value
        TickerSymbolDecrease = Cells(i, 9).Value
        TickerSymbolVolume = Cells(i, 9).Value
        GreatestIncrease = Cells(i, 11).Value
        GreatestDecrease = Cells(i, 11).Value
        GreatestVolume = Cells(i, 12).Value
        
    ' Step 3.1: Loop through result
    
        Do While Not (IsEmpty(Cells(i, 9)))
        
            PercentChange = Cells(i, 11).Value
            TotalVolume = Cells(i, 12).Value
            
            If PercentChange > GreatestIncrease Then
            
                TickerSymbolIncrease = Cells(i, 9).Value
                GreatestIncrease = Cells(i, 11).Value
            
            End If
            
            If PercentChange < GreatestDecrease Then
            
                TickerSymbolDecrease = Cells(i, 9).Value
                GreatestDecrease = Cells(i, 11).Value
            
            End If
            
            If TotalVolume > GreatestVolume Then
            
                TickerSymbolVolume = Cells(i, 9).Value
                GreatestVolume = Cells(i, 12).Value
            
            End If
            
            i = i + 1
            
        Loop

    ' Step XX: Input Values
    
        Range("P2").Value = TickerSymbolIncrease
        Range("Q2").Value = GreatestIncrease
        Range("P3").Value = TickerSymbolDecrease
        Range("Q3").Value = GreatestDecrease
        Range("P4").Value = TickerSymbolVolume
        Range("Q4").Value = GreatestVolume

    ' Step 3.4: Format Values
            
        Range("Q2").Select
        Selection.NumberFormat = "0.00%"
            
        Range("Q3").Select
        Selection.NumberFormat = "0.00%"


End Sub

