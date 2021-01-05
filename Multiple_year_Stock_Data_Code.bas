Attribute VB_Name = "Module1"
Sub Wall_Street_Ticker_w_Bonus()

  Dim j As Integer
  Dim ws_num As Integer

  Dim starting_ws As Worksheet
  Set starting_ws = ActiveSheet 'remember which worksheet is active in the beginning
  ws_num = ThisWorkbook.Worksheets.Count

    For j = 1 To ws_num
      ThisWorkbook.Worksheets(j).Activate
      ThisWorkbook.Worksheets(j).Cells.EntireColumn.AutoFit
      ThisWorkbook.Worksheets(j).Cells.EntireRow.AutoFit
      
        'Title Cells
        Cells(1, 9).Value = "Ticker"
  
        Cells(1, 10).Value = "Yearly Change"
  
        Cells(1, 11).Value = "Percent Change"
  
        Cells(1, 12).Value = "Total Stock Volume"
  
        Cells(2, 16).Value = "Greatest % Increase"
  
        Cells(3, 16).Value = "Greatest % Decrease"
  
        Cells(4, 16).Value = "Greatest Total Volume"
  
        Cells(1, 17).Value = "Ticker"
  
        Cells(1, 18).Value = "Value"
  
  
        'Set an initial variable for holding ticker symbol
        Dim ticker_symbol As String
  
        'Set an initial variable for holding Yearly Change
        Dim Yearly_Change As Double
  
        'Set an inital variable for holding Percent Change
        Dim Percent_Change As Double
  
        'Set an inital variable for holding Total Stock Volume
        Dim Stk_Vol As LongLong
        Stk_Vol = 0
  

        Dim lastrow As LongLong
   
        'Define Bonus variables
        Dim GrPIn As Double
    
        Dim GrPDec As Double
    
        Dim GrTotVol As LongLong
   
        Dim GrTicker As String
    
        Dim i As LongLong
  
        'BONUS Keep track of the location for each ticker in the Greatest table
        Dim Greatest_Summary_Table_Row As Integer
        Greatest_Summary_Table_Row = 2
    
        'Keep track of the location for each ticker in the summary table
        Dim Summary_Table_Row As LongLong
        Summary_Table_Row = 2
  
        'Set lastrow variable
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
  
        'Get opening price of first ticker
        Opening_price = Cells(2, 3).Value
    
    
        'Get ticker symbol of first ticker
        ticker_symbol = Cells(2, 1).Value
  
 
        'Loop through stocks for 2 to lastrow
        For i = 2 To lastrow
        
            If Opening_price = 0 Then
               
               Percent_Change = 0
               
               Closing_price = Cells(i, 6).Value
      
               Yearly_Change = Closing_price - Opening_price
                
               Stk_Vol = Cells(i, 7).Value + Stk_Vol
               
             
            
            Else
    
                Closing_price = Cells(i, 6).Value
      
                Yearly_Change = Closing_price - Opening_price
        
                Percent_Change = (Closing_price - Opening_price) / Opening_price
        
                Stk_Vol = Cells(i, 7).Value + Stk_Vol
                
            End If
            
    
          If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        
        
            'Output ticker in the summary table
            Cells(Summary_Table_Row, 9).Value = ticker_symbol
        
            'Output Stock Volume in summary table
            Cells(Summary_Table_Row, 12).Value = Stk_Vol
        
            'Output Yearly Change in summary table
            Cells(Summary_Table_Row, 10).Value = Yearly_Change
        
            'Output Percent Change in summary table
            Cells(Summary_Table_Row, 11).Value = FormatPercent(Percent_Change)
        
        
            ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
    
            'Reset the Stock Volume
            Stk_Vol = 0
        
    
            'Get the opening price of next ticker(i+1 row..)
            Opening_price = Cells(i + 1, 3).Value
        
            'Also get the ticker symbol for the next ticker & store it
            ticker_symbol = Cells(i + 1, 1).Value
        
        
          End If
          
          If Cells(Summary_Table_Row, 10).Value > 0 Then
    
            Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
        
          ElseIf Cells(Summary_Table_Row, 10).Value < 0 Then
        
            Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
        
          ElseIf Cells(Summary_Table_Row, 10).Value = 0 Then
    
            Cells(Summary_Table_Row, 10).Interior.ColorIndex = 15
            
          Else
            
             Cells(Summary_Table_Row, 10).Interior.ColorIndex = 2
        
          End If
    
      
        
       Next i
       
    Call FindGreatest
    

  
  Next j

    starting_ws.Activate 'activate the worksheet that was originally active

End Sub



Sub FindGreatest()

  GrPIn = 0
  GrPDec = 0
  GrTotVol = 0
      
      For i = 2 To Cells(Rows.Count, "I").End(xlUp).Row
      'For i = 2 To lastrow
        
        
        
          If Cells(i, 11).Value > GrPIn Then
            GrPIn = Cells(i, 11).Value
            
            'Set Ticker to column "q"
            Range("Q2").Value = Cells(i, 9).Value
            
            'Set Value to column "r"
            Range("R2").Value = FormatPercent(GrPIn)
            
            
            ElseIf Cells(i, 11).Value < GrPDec Then
          
                GrPDec = Cells(i, 11).Value
                Range("Q3").Value = Cells(i, 9).Value
                Range("R3").Value = FormatPercent(GrPDec)
                
            End If
                
                
            If Cells(i, 12).Value > GrTotVol Then
            
                GrTotVol = Cells(i, 12).Value
                
                'Set GrTotV Ticker to column "q"
                Range("Q4").Value = Cells(i, 9).Value
            
                'SetGrTotV Value to column "r"
                Range("R4").Value = GrTotVol
            
            
            End If
        
       Next i
       
       
End Sub


