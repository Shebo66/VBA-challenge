Attribute VB_Name = "Module1"
Sub Stocks():

'Declare all Variables

    Dim ticker As String
    Dim year_open As Double
    Dim year_high As Double
    Dim year_low As Double
    Dim year_close As Double
    Dim year_vol As Double
    Dim yearly_change As Double
    Dim percentage_change As Double
    Dim total_stockVolume As Double
    
    Range("A2").Value = ticker
    Range("D2").Value = year_high
    Range("E2").Value = year_low
    Range("F2").Value = year_close
    Range("G2").Value = year_vol
    Range("I2").Value = yearly_change
    Range("J2").Value = percentage_change
    Range("K2").Value = total_stockVolume
    
    MsgBox ("ticker")
    MsgBox ("year_open")
    MsgBox ("year_high")
    MsgBox ("year_low")
    MsgBox ("year_close")
    MsgBox ("year_vol")
    MsgBox ("yearly_change")
    MsgBox ("per_change")
    MsgBox ("total_stockVolume")
      
    For i = 2 To 70926
        If Cells(i, 9).Value >= 0 Then
        Cells(i, 9).Interior.ColorIndex = 4
        
        Else
        
        Cells(i, 9).Interior.ColorIndex = 3
        
       End If
      
        Num1 = Cells(i, 3).Value
        Num2 = Cells(i, 6).Value
        Num3 = Cells(i, 9).Value
       Num4 = Cells(i, 10).Value
      
        
      Cells(i, 9) = Cells(i, 6).Value - Cells(i, 3).Value
     
    Next i
        
    End Sub

Sub PercentageChange():

Dim ticker As String
    Dim year_open As Double
   Dim year_high As Double
   Dim year_low As Double
    Dim year_close As Double
    Dim year_vol As Double
    Dim yearly_change As Double
    Dim per_change As Double
    Dim total_stockVolume As Double
  
    For i = 2 To 70926
    
        Num1 = Cells(i, 3).Value
        Num2 = Cells(i, 6).Value
        Num3 = Cells(i, 10).Value
        per_change = Cells(i, 10).Value
        year_change = Cells(i, 6).Value - Cells(i, 3).Value
        Num3 = ((Num2 - Num1) / Num2)
        
      
       Cells(i, 10) = Cells(i, 9).Value / Cells(i, 6).Value
     
   Next i
   
End Sub


   


    








