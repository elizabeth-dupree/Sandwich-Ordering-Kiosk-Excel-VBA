Attribute VB_Name = "Transfer"
Sub PrintOrder(SandwichPrice As Double, BreadPrice As Double, CheesePrice As Double, DressingPrice As Double, ToppingsPrice As Double)
    Dim TotalPrice As Double
    Dim Row As Integer

    'identifying empty row
    For Row = 2 To 1000
        If Sheets("OrderSheet").Cells(Row, 1).Value = "" Then
            Exit For
        End If
    Next Row

    'calculating total price & printing information in respected column
    TotalPrice = SandwichPrice + BreadPrice + ToppingsPrice
    Sheets("OrderSheet").Cells(Row, 1).Value = ID
    Sheets("OrderSheet").Cells(Row, 2).Value = Now()
    Sheets("OrderSheet").Cells(Row, 3).Value = SandwichChoice
    Sheets("OrderSheet").Cells(Row, 4).Value = BreadChoice
    Sheets("OrderSheet").Cells(Row, 5).Value = CheeseChoice
    Sheets("OrderSheet").Cells(Row, 6).Value = DressingChoice
    Sheets("OrderSheet").Cells(Row, 7).Value = AlfalfaSprouts
    Sheets("OrderSheet").Cells(Row, 8).Value = MesclunGreens
    Sheets("OrderSheet").Cells(Row, 9).Value = RoastedTomatoes
    Sheets("OrderSheet").Cells(Row, 10).Value = XtraCheese
    Sheets("OrderSheet").Cells(Row, 11).Value = SandwichPrice
    Sheets("OrderSheet").Cells(Row, 12).Value = TotalPrice

 
End Sub


