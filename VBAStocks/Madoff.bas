Attribute VB_Name = "Madoff"
Sub Madoff()

    Application.ScreenUpdating = False
   
    Dim Sheetcount As Integer, LastRow As Long, SummRow As Integer, StockName As String, YearEnd As Double, _
    YearStart As Double, YearChange As Double, PercentChange As Double, Vol As Double
    
    Vol = 0
    Sheetcount = ActiveWorkbook.Worksheets.Count
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
  
          For i = 1 To Sheetcount
    
     
        Worksheets(i).Activate
        
      
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        Cells(1, 14).Value = "Year End"
        Cells(1, 15).Value = "Year Start"
        
        
        SummRow = 2
        
      
        
        If i = 1 Then
        
          
 For k = 2 To LastRow
            
  If Cells(k, 1).Value <> Cells(k + 1, 1).Value Then
      StockName = Cells(k, 1).Value
     Range("I" & SummRow).Value = StockName

  YearEnd = Cells(k, 6).Value
    YearStart = Cells(k - 261, 3).Value
      YearChange = YearEnd - YearStart
      Range("N" & SummRow).Value = YearEnd
           Range("O" & SummRow).Value = YearStart
            Range("J" & SummRow).Value = YearChange
                    
                
 If YearStart = 0 Then
        PercentChange = 0
                    Else
           PercentChange = YearChange / YearStart
          End If
                    Range("K" & SummRow).Value = PercentChange
                    
                    
                    If PercentChange >= 0 Then
                        Range("K" & SummRow).Interior.Color = RGB(0, 255, 0)
                    ElseIf PercentChange < 0 Then
                        Range("K" & SummRow).Interior.Color = RGB(255, 0, 0)
                    Else

                    End If
                  
                    Vol = Vol + Cells(k, 7).Value
                    Range("L" & SummRow).Value = Vol
                    
                    SummRow = SummRow + 1
                    Vol = 0
                    
                Else
                    Vol = Vol + Cells(k, 7).Value
                
                End If
            Next k
        
        Else
        
          
            For j = 2 To LastRow
            
                If Cells(j, 1).Value <> Cells(j + 1, 1).Value Then
            
                    StockName = Cells(j, 1).Value
                    Range("I" & SummRow).Value = StockName
                    
                    
                    YearEnd = Cells(j, 6).Value
                    YearStart = Cells(j - 260, 3).Value
                    YearChange = YearEnd - YearStart
                    Range("N" & SummRow).Value = YearEnd
                    Range("O" & SummRow).Value = YearStart
                    Range("J" & SummRow).Value = YearChange
                    
                If YearStart = 0 Then
                        PercentChange = 0
                    Else
                        PercentChange = YearChange / YearStart
                    End If

                    Range("K" & SummRow).Value = PercentChange
                    
                   
                    If PercentChange >= 0 Then
                        Range("K" & SummRow).Interior.Color = RGB(0, 255, 0)
                    ElseIf PercentChange < 0 Then
                        Range("K" & SummRow).Interior.Color = RGB(255, 0, 0)
                    Else

                    End If
         
                    Vol = Vol + Cells(j, 7).Value
                    Range("L" & SummRow).Value = Vol
                    
                    SummRow = SummRow + 1
                    Vol = 0
                Else
                    Vol = Vol + Cells(j, 7).Value
                    
                End If
            Next j
        End If
    Next i

  
    Worksheets("2016").Activate

Application.ScreenUpdating = True
End Sub

