Attribute VB_Name = "InventoryReport"
Public Sub GenerateInventoryReport()
  
    Dim i As Integer
    Dim Row As Integer
    
    Dim SandwichChoiceArray(13) As String
    Dim CheeseChoiceArray(5) As String
    Dim BreadChoiceArray(4) As String
    Dim DressingChoiceArray(6) As String
    Dim ToppingRow As Integer

    'sandwich - read the sandwich menu options into array and set count for each sandwich to 0
    For Row = 0 To 13
        SandwichChoiceArray(Row) = Worksheets("Inventory Report").Cells(Row + 2, 1).Value
        Worksheets("Inventory Report").Cells(Row + 2, 2).Value = 0
    Next Row

    'cheese - read the cheese menu options into array and set count for each cheese to 0
    For Row = 0 To 5
        CheeseChoiceArray(Row) = Worksheets("Inventory Report").Cells(Row + 2, 4).Value
        Worksheets("Inventory Report").Cells(Row + 2, 5).Value = 0
    Next Row

    'bread - read the bread menu options into array and set count for each bread to 0
    For Row = 0 To 4
        BreadChoiceArray(Row) = Worksheets("Inventory Report").Cells(Row + 2, 7).Value
        Worksheets("Inventory Report").Cells(Row + 2, 8).Value = 0
    Next Row

    'dressing - read the dressing menu options into array and set count for each dressing to 0
    For Row = 0 To 6
        DressingChoiceArray(Row) = Worksheets("Inventory Report").Cells(Row + 2, 10).Value
        Worksheets("Inventory Report").Cells(Row + 2, 11).Value = 0
    Next Row
    
    'toppings - set count for each topping to 0
    For ToppingRow = 2 To 5
        Worksheets("Inventory Report").Cells(ToppingRow, 14).Value = 0
    Next ToppingRow
    
    
    'Counts
    For i = 2 To 100
        
    'Sandwich
        For Row = 0 To 13
            If Worksheets("OrderSheet").Cells(i, 3).Value = SandwichChoiceArray(Row) Then
                Worksheets("Inventory Report").Cells(Row + 2, 2).Value = Worksheets("Inventory Report").Cells(Row + 2, 2).Value + 1
            End If
        Next Row
        
    'Bread
        For Row = 0 To 4
            If Worksheets("OrderSheet").Cells(i, 4).Value = BreadChoiceArray(Row) Then
                Worksheets("Inventory Report").Cells(Row + 2, 8).Value = Worksheets("Inventory Report").Cells(Row + 2, 8).Value + 1
            End If
        Next Row
        
    'Cheese
        For Row = 0 To 5
            If Worksheets("OrderSheet").Cells(i, 5).Value = CheeseChoiceArray(Row) Then
                Worksheets("Inventory Report").Cells(Row + 2, 5).Value = Worksheets("Inventory Report").Cells(Row + 2, 5).Value + 1
            End If
        Next Row
        
    'Dressing
        For Row = 0 To 6
            If Worksheets("OrderSheet").Cells(i, 6).Value = DressingChoiceArray(Row) Then
                Worksheets("Inventory Report").Cells(Row + 2, 11).Value = Worksheets("Inventory Report").Cells(Row + 2, 11).Value + 1
            End If
        Next Row
        
    'Sprouts
    If Worksheets("OrderSheet").Cells(i, 7).Value = 1 Then
        Worksheets("Inventory Report").Cells(2, 14).Value = Worksheets("Inventory Report").Cells(2, 14).Value + 1
    End If

    'Greens
    If Worksheets("OrderSheet").Cells(i, 8).Value = 1 Then
        Worksheets("Inventory Report").Cells(3, 14).Value = Worksheets("Inventory Report").Cells(3, 14).Value + 1
    End If

    'Tomatoes
    If Worksheets("OrderSheet").Cells(i, 9).Value = 1 Then
     Worksheets("Inventory Report").Cells(4, 14).Value = Worksheets("Inventory Report").Cells(4, 14).Value + 1
    End If

    'Extra Cheese
    If Worksheets("OrderSheet").Cells(i, 10).Value = 1 Then
        Worksheets("Inventory Report").Cells(5, 14).Value = Worksheets("Inventory Report").Cells(5, 14).Value + 1
    End If
        
    Next i
    

End Sub
