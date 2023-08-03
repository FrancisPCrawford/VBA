VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Sheetcode()
    Dim xSH As Worksheet
    Application.ScreenUpdating = False
    For Each xSH In Worksheets
       xSH.Select
        Call stockExpert
   Next
    Application.ScreenUpdating = True

End Sub

Sub stockExpert()
Dim column As Integer
Dim LR As Long
Dim Volume As Double
Dim ticker As Integer
Dim Oprice As Double
Dim Cprice As Double
Dim Change As Double
Dim PC As Double

column = 1
ticker = 2
Volume = 0
LR = Cells(Rows.Count, 2).End(xlUp).Row

Oprice = Cells(2, 3).Value

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

For i = 2 To LR

    If Cells(i + 1, column).Value = Cells(i, column).Value Then
        Volume = Volume + Cells(i, 7).Value
    
    ElseIf Cells(i + 1, column).Value <> Cells(i, column).Value Then
        
         Volume = Volume + Cells(i, 7).Value
        
        Cells(ticker, 9).Value = Cells(i, column).Value
        Cells(ticker, 12).Value = Volume
        
        Cprice = Cells(i, 6).Value
        
        
        Change = Cprice - Oprice
        PC = (Change / Oprice) * 100
               
        Cells(ticker, 10).Value = Change
        

        Cells(ticker, 11).Value = PC
        Oprice = Cells(i + 1, 3).Value
        
        Volume = 0
        ticker = ticker + 1
        Cprice = 0
   
    End If

Next i


For i = 2 To lrow

    If Cells(i, 10).Value > 0 Then
        Cells(i, 10).Interior.ColorIndex = 4
        Cells(i, 11).Interior.ColorIndex = 4
    ElseIf Cells(i, 10).Value < 0 Then
        Cells(i, 10).Interior.ColorIndex = 3
        Cells(i, 11).Interior.ColorIndex = 3
    End If
    
Next i
                ticker = ticker + 1
                Oprice = 0
            

End Sub


