VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Stock_Ticker()
  
  Dim Stock_Ticker As String
  
  Dim Volume_Total As Double
  Volume_Total = 0

  lastrow = Cells(Rows.Count, 1).End(xlUp).Row
  
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  
  For i = 2 To lastrow
   
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
   
      Stock_Ticker = Cells(i, 1).Value
     
      Volume_Total = Volume_Total + Cells(i, 7).Value
      
      Range("M" & Summary_Table_Row).Value = Stock_Ticker

      Range("P" & Summary_Table_Row).Value = Volume_Total
      
      Summary_Table_Row = Summary_Table_Row + 1
            
      Volume_Total = 0
    
    Else

      Volume_Total = Volume_Total + Cells(i, 7).Value

    End If

  Next i

End Sub

Sub Stock_Price_Change()
    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim Start_Row As Long
    Dim End_Row As Long
    Dim Stock_Ticker As String
    Dim Change_Total As Double

    Dim Summary_Table_Row As Long
    Summary_Table_Row = 2
  

    Start_Row = 2
    End_Row = Cells(Rows.Count, 1).End(xlUp).Row

        
  
    For i = Start_Row To End_Row
        If Stock_Ticker <> Cells(i, 1).Value Then
            If Stock_Ticker <> "" Then
                Close_Price = Cells(i - 1, 6).Value
                Change_Total = (Close_Price - Open_Price)
                Range("N" & Summary_Table_Row).Value = Change_Total
                Summary_Table_Row = Summary_Table_Row + 1
            End If
            Stock_Ticker = Cells(i, 1).Value
            Open_Price = Cells(i, 3).Value
            Change_Total = 0
            End If
    
         
        If i = End_Row Then
        Close_Price = Cells(i, 6).Value
        Change_Total = (Close_Price - Open_Price)
        Range("N" & Summary_Table_Row).Value = Change_Total
        End If
        
        
      Next i
        
       

    Close_Price = Cells(End_Row, 6).Value
    Change_Total = (Close_Price - Open_Price)
    Cells(End_Row, 6).Value = Change_Total

End Sub

Sub Percent_Stock_Change()

    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim Start_Row As Long
    Dim End_Row As Long
    Dim Stock_Ticker As String
    Dim Percent_Change As Double
    Dim rng As Range
    
    
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
  
    
    
    Start_Row = 2
    End_Row = Cells(Rows.Count, 1).End(xlUp).Row

    For i = Start_Row To End_Row
        If Stock_Ticker <> Cells(i, 1).Value Then
            If Stock_Ticker <> "" Then
                Close_Price = Cells(i - 1, 6).Value
                Percent_Change = (Close_Price - Open_Price) / (Open_Price)
                Range("O" & Summary_Table_Row).Value = Percent_Change
                Summary_Table_Row = Summary_Table_Row + 1
            End If
            Stock_Ticker = Cells(i, 1).Value
            Open_Price = Cells(i, 3).Value
            Set rng = Range("O" & Summary_Table_Row)
            rng.NumberFormat = "0.00%"
            End If
            
            If i = End_Row Then
            Close_Price = Cells(i, 6).Value
            Percent_Change = (Close_Price - Open_Price) / Open_Price
            Range("O" & Summary_Table_Row).Value = Percent_Change
            End If
    
            If Cells(i, 14).Value <= 0 Then
            Cells(i, 14).Interior.Color = vbRed
            Else
            Cells(i, 14).Interior.Color = vbGreen
            End If
        
    Next i

    Close_Price = Cells(End_Row, 6).Value
    Percent_Change = (Close_Price - Open_Price) / (Open_Price)
    Cells(End_Row, 6).Value = Change_Total
    
    
End Sub

Sub Greatest_Increase_Percent_Value()
    Dim maxCount As Double
    Dim rng As Range
    Dim stockValueIndex As Integer
    Dim stockName As Integer
    
    Set rng = Range("U2")
    rng.NumberFormat = "0.00%"
    maxCount = WorksheetFunction.Max(Range("O2:O3001"))
    Range("T1").Value = "Ticker"
    Range("U1").Value = "Value"
    Range("S2").Value = "Greatest % Increase"
    Range("S3").Value = "Greatest % Decrease"
    Range("S4").Value = "Greatest Total Volume"
    Range("U2").Value = maxCount
    Range("M:U").Columns.AutoFit
    stockValueIndex = WorksheetFunction.Match(maxCount, Range("O2:O3001"), 0)
    Range("T2").Value = Range("M" & stockValueIndex + 1).Value
    
    
                
End Sub

Sub Greatest_Decrease_Percent_Value()
    Dim minCount As Double
    Dim rng As Range
    Dim stockValueIndex As Integer
    Dim stockName As Integer
    
    Set rng = Range("U3")
    rng.NumberFormat = "0.00%"
    minCount = WorksheetFunction.Min(Range("O2:O3001"))
    Range("U3").Value = minCount
    stockValueIndex = WorksheetFunction.Match(minCount, Range("O2:O3001"), 0)
    Range("T3").Value = Range("M" & stockValueIndex + 1).Value
    
    
                
End Sub

Sub Greatest_Total_Volume()
    Dim maxCount As Double
    Dim rng As Range
    Dim stockValueIndex As Integer
    Dim stockName As Integer
    
   
    maxCount = WorksheetFunction.Max(Range("G2:G3001"))
    Range("U4").Value = maxCount
    stockValueIndex = WorksheetFunction.Match(maxCount, Range("G2:G3001"), 0)
    Range("T4").Value = Range("M" & stockValueIndex + 1).Value
    Range("M:U").Columns.AutoFit
    
                
End Sub

