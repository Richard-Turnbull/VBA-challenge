Sub StockData()

'--------------------------------------------------------------------------
' SET DIMENSIONS
'--------------------------------------------------------------------------
   
    'set dimensions for ws to loop through worksheets
    Dim ws As Worksheet
   
    'set dimension for number of sheets in workbook
    Dim n_sheets As Integer
   
    'set dimensions for ticker summary table
    Dim LastRowA, LastRowI, LastRowL As Long
    Dim OpenPrice, ClosePrice, TotalVolumne As Double
    Dim Ticker As String
          
    'set dimensions for greatest % and Volume summary table
    Dim MaxIncrease, MaxDecrease, MaxVolume As Double
    Dim MaxIncreaseTicker, MaxDecreaseTicker, MaxVolumeTicker As String

'--------------------------------------------------------------------------
' SET DYNAMIC ARRAY TO ENABLE LOOPING THROUGH NUMBER OF SHEETS IN WORKBOOK
'--------------------------------------------------------------------------

    For Each ws In Worksheets
    
        'set number of worksheets in variable n_sheets
        n_sheets = Application.Sheets.Count
        
        'for loop to loop through number of sheets
        For a = 0 To n_sheets
     
'--------------------------------------------------------------------------
' SET HEADERS FOR SUMMARY TABLES IN ALL SHEETS
'--------------------------------------------------------------------------
    
            'set values for column headers on ticker summary table
            ws.Cells(1, 9).Value = "Ticker"
            ws.Cells(1, 10).Value = "Yearly Change"
            ws.Cells(1, 11).Value = "Percentage Change"
            ws.Cells(1, 12).Value = "Total Stock Volume"
          
            'set values for column headers on greatest % and Volume summary table
            ws.Cells(1, 16).Value = "Ticker"
            ws.Cells(1, 17).Value = "Value"
          
            'set values for row headers on greatest % and Volume summary table
            ws.Cells(2, 15).Value = "Greatest % Increase"
            ws.Cells(3, 15).Value = "Greatest % Decrease"
            ws.Cells(4, 15).Value = "Greatest Total Volume"
             
'--------------------------------------------------------------------------
' POPULATE TOTALVOLUME IN COLUMN L IN ALL SHEETS
'--------------------------------------------------------------------------
        
            'Set last rows for columns A and L
            LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
'            LastRowA = ws.Range("A" & Rows.Count).End(xlUp).Row
            LastRowL = 2
               
            'Set Total Stock Volume Value
            TotalVolumne = 0
                        
            'loop through all rows of data set to populate column L of ticker summary table
            For i = 2 To LastRowA
               
                'check if ticker matches row below (to identify if volume should be added to total volume for the ticker)
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                    'if ticker is different, add volume to TotalVolume variable
                    TotalVolumne = TotalVolume + ws.Cells(i, 7).Value
                                            
                    'populate TotalVolume in next cell down in column L
                    ws.Range("L" & LastRowL).Value = TotalVolume
                
                    'add 1 to LastRowL count
                    LastRowL = LastRowL + 1
                
                    'reset TotalVolume to 0 for next ticker
                    TotalVolume = 0
                                                                                       
                Else

                'if the ticker remains the same in next row, add the volume to TotalVolume variable
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            
                End If
            Next i
        Next a
    Next
 
'--------------------------------------------------------------------------
' ACTIVATE SHEET "2016"
'--------------------------------------------------------------------------
        
    Sheets("2016").Activate

'--------------------------------------------------------------------------
' CREATE THE TICKER SUMMARY TABLE
'--------------------------------------------------------------------------
          
    'Set last rows for columns A and L
    LastRowA = Range("A" & Rows.Count).End(xlUp).Row
    LastRowL = 2
        
    'sort data (A:G) by Ticker (column A Ascending) then Date (column B descending)
    Range("A1:G" & LastRowA).Columns.Sort key1:=Columns("A"), Order1:=xlAscending, Key2:=Columns("B"), Order2:=xlDescending, Header:=xlYes
    
        'loop through all rows of data set to populate columns I:K of ticker summary table
        For i = 2 To LastRowA

           'check if ticker matches row below
            If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
        
                'if does not match, copy cell
                Cells(i, 1).Copy
                    
                'save column I row count in variable "LastRowI"
                LastRowI = Range("I" & Rows.Count).End(xlUp).Row
        
                'paste ticker into next available row
                Range("I" & LastRowI + 1).PasteSpecial
                                
                'Set ClosePrice variables
                ClosePrice = Cells(i, 6).Value
            
            End If
                                    
            'check if ticker matches row below
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                                
                'Set OpenPrice variable
                OpenPrice = Cells(i, 3).Value
                               
                'Calculate yearly change and populate in next row down in column "J"
                Range("J" & LastRowI + 1).Value = ClosePrice - OpenPrice
                    
                    'Set color formatting if yearly change (>0 = Green, <0 = red)
                    If Range("J" & LastRowI + 1).Value > 0 Then
                        Range("J" & LastRowI + 1).Interior.ColorIndex = 4
                    ElseIf Range("J" & LastRowI + 1).Value < 0 Then
                        Range("J" & LastRowI + 1).Interior.ColorIndex = 3
                    End If
            End If

            'if OpenPrice or ClosePrice = 0 then populate column K as 0, if not, calculate the percentage change.
            If OpenPrice = 0 Or ClosePrice = 0 Then
                    Range("K" & LastRowI + 1).Value = 0
                Else
                    Range("K" & LastRowI + 1).Value = (ClosePrice - OpenPrice) / ClosePrice
            End If

        Next i
    
'--------------------------------------------------------------------------
' CREATE THE GREATEST % & TOTAL VALUE SUMMARY TABLE
'--------------------------------------------------------------------------
    LastRowSummary = Range("K" & Rows.Count).End(xlUp).Row

    MaxIncrease = Cells(2, 11).Value
    MaxDecrease = Cells(2, 11).Value
    MaxVolume = Cells(2, 12).Value

    For j = 1 To LastRowSummary - 1
               
        If Cells(j + 1, 11).Value > MaxIncrease Then
            MaxIncrease = Cells(j + 1, 11).Value
            MaxIncreaseTicker = Cells(j + 1, 9).Value
            Range("Q2").Value = MaxIncrease
            Range("P2").Value = MaxIncreaseTicker
        End If

    Next j

    For k = 1 To LastRowSummary - 1
               
        If Cells(k + 1, 11).Value < MaxDecrease Then
            MaxDecrease = Cells(k + 1, 11).Value
            MaxDecreaseTicker = Cells(k + 1, 9).Value
            Range("Q3").Value = MaxDecrease
            Range("P3").Value = MaxDecreaseTicker
        End If

    Next k

    For m = 1 To LastRowSummary - 1
               
        If Cells(m + 1, 12).Value > MaxVolume Then
            MaxVolume = Cells(m + 1, 12).Value
            MaxVolumeTicker = Cells(m + 1, 9).Value
            Range("Q4").Value = MaxVolume
            Range("P4").Value = MaxVolumeTicker
        End If

    Next m

'--------------------------------------------------------------------------
' ACTIVATE SHEET "2015"
'--------------------------------------------------------------------------
        
    Sheets("2015").Activate

'--------------------------------------------------------------------------
' CREATE THE TICKER SUMMARY TABLE
'--------------------------------------------------------------------------
          
    'Set last rows for columns A and L
    LastRowA = Range("A" & Rows.Count).End(xlUp).Row
    LastRowL = 2
        
    'Set Total Stock Volume Value
    TotalVolumne = 0
        
    'sort data (A:G) by Ticker (column A Ascending) then Date (column B descending)
    Range("A1:G" & LastRowA).Columns.Sort key1:=Columns("A"), Order1:=xlAscending, Key2:=Columns("B"), Order2:=xlDescending, Header:=xlYes
    
    
        'loop through all rows of data set to populate columns I:K of ticker summary table
        For i = 2 To LastRowA

           'check if ticker matches row below (to identify last day in year results for the ticker)
            If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
        
                'if does not match, copy cell
                Cells(i, 1).Copy
                    
                'save column I row count in variable "LastRowI"
                LastRowI = Range("I" & Rows.Count).End(xlUp).Row
        
                'paste ticker into next available row
                Range("I" & LastRowI + 1).PasteSpecial
                                
                'Set ClosePrice variables
                ClosePrice = Cells(i, 6).Value
            
            End If
                                    
            'check if ticker matches row below (to identify first day in year results for the ticker)
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                                
                'Set OpenPrice variables
                OpenPrice = Cells(i, 3).Value
                               
                'Calculate yearly change and populate in next row down in column "J"
                Range("J" & LastRowI + 1).Value = ClosePrice - OpenPrice
                    
                    'Set color formatting if yearly change (>0 = Green, <0 = red)
                    If Range("J" & LastRowI + 1).Value > 0 Then
                        Range("J" & LastRowI + 1).Interior.ColorIndex = 4
                    ElseIf Range("J" & LastRowI + 1).Value < 0 Then
                        Range("J" & LastRowI + 1).Interior.ColorIndex = 3
                    End If
            End If

            'if OpenPrice or ClosePrice = 0 then populate column K as 0, if not, calculate the percentage change.
            If OpenPrice = 0 Or ClosePrice = 0 Then
                    Range("K" & LastRowI + 1).Value = 0
                Else
                    Range("K" & LastRowI + 1).Value = (ClosePrice - OpenPrice) / ClosePrice
            End If

        Next i
    
'--------------------------------------------------------------------------
' CREATE THE GREATEST % & TOTAL VALUE SUMMARY TABLE
'--------------------------------------------------------------------------
    LastRowSummary = Range("K" & Rows.Count).End(xlUp).Row

    MaxIncrease = Cells(2, 11).Value
    MaxDecrease = Cells(2, 11).Value
    MaxVolume = Cells(2, 12).Value

    For j = 1 To LastRowSummary - 1
               
        If Cells(j + 1, 11).Value > MaxIncrease Then
            MaxIncrease = Cells(j + 1, 11).Value
            MaxIncreaseTicker = Cells(j + 1, 9).Value
            Range("Q2").Value = MaxIncrease
            Range("P2").Value = MaxIncreaseTicker
        End If

    Next j

    For k = 1 To LastRowSummary - 1
               
        If Cells(k + 1, 11).Value < MaxDecrease Then
            MaxDecrease = Cells(k + 1, 11).Value
            MaxDecreaseTicker = Cells(k + 1, 9).Value
            Range("Q3").Value = MaxDecrease
            Range("P3").Value = MaxDecreaseTicker
        End If

    Next k

    For m = 1 To LastRowSummary - 1
               
        If Cells(m + 1, 12).Value > MaxVolume Then
            MaxVolume = Cells(m + 1, 12).Value
            MaxVolumeTicker = Cells(m + 1, 9).Value
            Range("Q4").Value = MaxVolume
            Range("P4").Value = MaxVolumeTicker
        End If

    Next m
 
'--------------------------------------------------------------------------
' ACTIVATE SHEET "2014"
'--------------------------------------------------------------------------
        
    Sheets("2014").Activate

'--------------------------------------------------------------------------
' CREATE THE TICKER SUMMARY TABLE
'--------------------------------------------------------------------------
          
    'Set last rows for columns A and L
    LastRowA = Range("A" & Rows.Count).End(xlUp).Row
    LastRowL = 2
        
    'Set Total Stock Volume Value
    TotalVolumne = 0
        
    'sort data (A:G) by Ticker (column A Ascending) then Date (column B descending)
    Range("A1:G" & LastRowA).Columns.Sort key1:=Columns("A"), Order1:=xlAscending, Key2:=Columns("B"), Order2:=xlDescending, Header:=xlYes
    
    
        'loop through all rows of data set to populate columns I:K of ticker summary table
        For i = 2 To LastRowA

           'check if ticker matches row below (to identify last day in year results for the ticker)
            If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
        
                'if does not match, copy cell
                Cells(i, 1).Copy
                    
                'save column I row count in variable "LastRowI"
                LastRowI = Range("I" & Rows.Count).End(xlUp).Row
        
                'paste ticker into next available row
                Range("I" & LastRowI + 1).PasteSpecial
                                
                'Set ClosePrice variables
                ClosePrice = Cells(i, 6).Value
            
            End If
                                    
            'check if ticker matches row below (to identify first day in year results for the ticker)
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                                
                'Set OpenPrice variables
                OpenPrice = Cells(i, 3).Value
                               
                'Calculate yearly change and populate in next row down in column "J"
                Range("J" & LastRowI + 1).Value = ClosePrice - OpenPrice
                    
                    'Set color formatting if yearly change (>0 = Green, <0 = red)
                    If Range("J" & LastRowI + 1).Value > 0 Then
                        Range("J" & LastRowI + 1).Interior.ColorIndex = 4
                    ElseIf Range("J" & LastRowI + 1).Value < 0 Then
                        Range("J" & LastRowI + 1).Interior.ColorIndex = 3
                    End If
            End If
            
            'if OpenPrice or ClosePrice = 0 then populate column K as 0, if not, calculate the percentage change.
            If OpenPrice = 0 Or ClosePrice = 0 Then
                    Range("K" & LastRowI + 1).Value = 0
                Else
                    Range("K" & LastRowI + 1).Value = (ClosePrice - OpenPrice) / ClosePrice
            End If
                
        Next i
               
'--------------------------------------------------------------------------
' CREATE THE GREATEST % & TOTAL VALUE SUMMARY TABLE
'--------------------------------------------------------------------------
    LastRowSummary = Range("K" & Rows.Count).End(xlUp).Row

    MaxIncrease = Cells(2, 11).Value
    MaxDecrease = Cells(2, 11).Value
    MaxVolume = Cells(2, 12).Value

    For j = 1 To LastRowSummary - 1
               
        If Cells(j + 1, 11).Value > MaxIncrease Then
            MaxIncrease = Cells(j + 1, 11).Value
            MaxIncreaseTicker = Cells(j + 1, 9).Value
            Range("Q2").Value = MaxIncrease
            Range("P2").Value = MaxIncreaseTicker
        End If

    Next j

    For k = 1 To LastRowSummary - 1
               
        If Cells(k + 1, 11).Value < MaxDecrease Then
            MaxDecrease = Cells(k + 1, 11).Value
            MaxDecreaseTicker = Cells(k + 1, 9).Value
            Range("Q3").Value = MaxDecrease
            Range("P3").Value = MaxDecreaseTicker
        End If

    Next k

    For m = 1 To LastRowSummary - 1
               
        If Cells(m + 1, 12).Value > MaxVolume Then
            MaxVolume = Cells(m + 1, 12).Value
            MaxVolumeTicker = Cells(m + 1, 9).Value
            Range("Q4").Value = MaxVolume
            Range("P4").Value = MaxVolumeTicker
        End If

    Next m
    
End Sub