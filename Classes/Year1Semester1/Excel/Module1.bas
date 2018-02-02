Attribute VB_Name = "Module1"
Const dxMove = 45
Const dyMove = 37
Public strActiveItem As String
Public items As Object
Dim stoneBlock As New MCItem


Sub SetupItems()
    items.tools.wood = ["pickaxe", "shovel", "axe", "hoe", "sword"]
End Sub

Sub EnableWASD()
    ' reset first, just to be safe
    DisableWASD
    
    Application.OnKey "a", "MoveLeft"
    Application.OnKey "d", "MoveRight"
    Application.OnKey "w", "MoveUp"
    Application.OnKey "s", "MoveDown"
    
    Range("KeyboardStatus").Value = "Enabled"
End Sub

Sub DisableWASD()
    Application.OnKey "a"
    Application.OnKey "d"
    Application.OnKey "w"
    Application.OnKey "s"
    
    Range("KeyboardStatus").Value = "Disabled"
End Sub

Private Sub MoveRight()
    With ActiveSheet.Shapes.Range(Array("SteveImage"))
        .Left = .Left + dxMove
    End With
End Sub

Private Sub MoveLeft()
    With ActiveSheet.Shapes.Range(Array("SteveImage"))
        .Left = .Left - dxMove
    End With
End Sub

Private Sub MoveUp()
    With ActiveSheet.Shapes.Range(Array("SteveImage"))
        .Top = .Top - dyMove
    End With
End Sub

Private Sub MoveDown()
    With ActiveSheet.Shapes.Range(Array("SteveImage"))
        .Top = .Top + dyMove
    End With
End Sub

Function InRange(Range1 As Range, Range2 As Range)
    Dim RangeIntersection As Range
    Set RangeIntersection = Application.Intersect(Range1, Range2)
    InRange = Not RangeIntersection Is Nothing
End Function

Function CRAFT(Value As String)
    If Value = "Wooden Pickaxe" Then
        CRAFT = "Crafting Complete!"
    Else
        CRAFT = "#ERROR/NYI"
    End If
End Function

Sub PlaceItem(strItemToPlace As String, Target As Range)
    If Target.Value = "Sky" Then
        Target.Value = strActiveItem
    End If
End Sub

Sub ClearItem(Target As Range)
    If Target.Value = "Stone" And Not strActiveItem = "Wooden Pickaxe" Then End
    If Not Target.Value = "Sky" Then
        Target.Value = "Sky"
    End If
End Sub

Sub UpdateActiveItem(strNewItem As String)
    strActiveItem = strNewItem
End Sub

Sub UseActiveItem(Target As Range)
    Select Case strActiveItem
        Case "Dirt", "Wood"
            PlaceItem strActiveItem, Target
        Case "Wooden Shovel", "Wooden Axe", "Wooden Pickaxe"
            ClearItem Target
        Case Else
    End Select
End Sub


