Attribute VB_Name = "MostOrdered"
Public Sub ShowMostOrdered()

    Dim Row As Integer
    Dim MaxSandwich As Integer
    Dim SandwichName As String
    Dim MaxCheese As Integer
    Dim CheeseName As String
    Dim MaxBread As Integer
    Dim BreadName As String
    Dim MaxDressing As Integer
    Dim DressingName As String
    Dim MaxTopping As Integer
    Dim ToppingName As String
   
    'find name of Sandwich with greatest count
    MaxSandwich = 0
    For Row = 2 To 15
        If Worksheets("Inventory Report").Cells(Row, 2).Value > MaxSandwich Then
            MaxSandwich = Worksheets("Inventory Report").Cells(Row, 2).Value
            SandwichName = Worksheets("Inventory Report").Cells(Row, 1).Value
        End If
    Next Row

    'find name of Cheese with greatest count
    MaxCheese = 0
    For Row = 2 To 7
        If Worksheets("Inventory Report").Cells(Row, 5).Value > MaxCheese Then
            MaxCheese = Worksheets("Inventory Report").Cells(Row, 5).Value
            CheeseName = Worksheets("Inventory Report").Cells(Row, 4).Value
        End If
    Next Row

    'find name of Bread with greatest count
    MaxBread = 0
    For Row = 2 To 6
        If Worksheets("Inventory Report").Cells(Row, 8).Value > MaxBread Then
            MaxBread = Worksheets("Inventory Report").Cells(Row, 8).Value
            BreadName = Worksheets("Inventory Report").Cells(Row, 7).Value
        End If
    Next Row
    
    'find name of Dressing with greatest count
    MaxDressing = 0
    For Row = 2 To 8
        If Worksheets("Inventory Report").Cells(Row, 11).Value > MaxDressing Then
            MaxDressing = Worksheets("Inventory Report").Cells(Row, 11).Value
            DressingName = Worksheets("Inventory Report").Cells(Row, 10).Value
        End If
    Next Row
    
    'find name of Toppings with greatest count
    MaxTopping = 0
    For Row = 2 To 5
        If Worksheets("Inventory Report").Cells(Row, 14).Value > MaxTopping Then
            MaxTopping = Worksheets("Inventory Report").Cells(Row, 14).Value
            ToppingName = Worksheets("Inventory Report").Cells(Row, 13).Value
        End If
    Next Row
    
    'pass Names to another sub to display the message
    Call DisplayMostOrdered(SandwichName, CheeseName, BreadName, DressingName, ToppingName)
 
End Sub

Sub DisplayMostOrdered(Sandwich As String, Cheese As String, Bread As String, _
                        Dressing As String, Topping As String)
    
    'show names of most ordered items and format with a new line after each
    MsgBox ("Most Ordered Sandwich: " & Sandwich & vbNewLine & _
            "Most Ordered Cheese: " & Cheese & vbNewLine & _
            "Most Ordered Bread: " & Bread & vbNewLine & _
            "Most Ordered Dressing: " & Dressing & vbNewLine & _
            "Most Ordered Topping: " & Topping)

End Sub
