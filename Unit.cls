VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Unit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit


Private p_Health As Integer
Private p_Damage As Integer
Private p_Defense As Integer
Private p_Speed As Integer
Private p_Sight As Integer
Private p_Range As Integer
Private p_Exp As Integer
Private p_Mana As Integer
Private p_Type Ad Integer

Private p_X As Long
Private p_Y As Long

Private p_Inventory(255) As Integer
Private p_InventoryPos(255) As Integer

Private p_Setup(17) As Integer

Private p_EffectLen() As Integer
Private p_Effects() As Integer

Public Enum SetupPoint
    e_Boots = 0
    e_Leggings = 1
    e_Chest = 2
    e_Gloves = 3
    e_Helmet = 4
    e_Ring0 = 5
    e_Ring1 = 6
    e_Ring2 = 7
    e_Ring3 = 8
    e_Ring4 = 9
    e_Ring5 = 10
    e_Ring6 = 11
    e_Ring7 = 12
    e_Ring8 = 13
    e_Ring9 = 14
    e_Necklace = 15
    e_PriWeapon = 16
    e_SecWeapon = 17
End Enum

Public Property Let Health(Value As Integer): p_Health = Value: If p_Health =< 0 Then Set Me = Nothing: End Property
Public Property Get Health() As Integer: Health = p_Health: End Property

Public Property Let Damage(Value As Integer): p_Damage = Value: End Property
Public Property Get Damage() As Integer: Damage = p_Damage: End Property

Public Property Let Defense(Value As Integer): p_Defense = Value: End Property
Public Property Get Defense() As Integer: Defense = p_Defense: End Property

Public Property Let Speed(Value As Integer): p_Speed = Value: End Property
Public Property Get Speed() As Integer: Speed = p_Speed: End Property

Public Property Let Sight(Value As Integer): p_Sight = Value: End Property
Public Property Get Sight() As Integer: Sight = p_Sight: End Property

Public Property Let Range(Value As Integer): p_Range = Value: End Property
Public Property Get Range() As Integer: Range = p_Range: End Property

Public Property Let Exp(Value As Integer): p_Exp = Value: End Property
Public Property Get Exp() As Integer: Exp = p_Exp: End Property

Public Property Let Mana(Value As Integer): p_Mana = Value: End Property
Public Property Get Mana() As Integer: Mana = p_Mana: End Property


Public Function Move(X As Integer, Y As Integer) As Integer
    Move = GetTile(CurrentMap, p_X + X, p_Y + Y)
    If Walkable(Move) Then
        p_X = p_X + X
        p_Y = p_Y + Y
    ElseIf IsEnemy(p_Type, p_X + X, p_Y + y) Then
        
    Else
    End If
End Function

Public Function MovePath(Path() As Integer, Index As Integer) As Integer
    MovePath = Move(Path(Index, 0), Path(Index, 1)
End Function

Public Property Let Type(Value As Integer): p_Type = Value: End Property
Public Property Get Type() As Integer: Type = p_Type: End Property

Public Property Let HealthPlus(Value As Integer): p_Health = p_Health + Value: If p_Health =< 0 Then Set Me = Nothing: End Property
Public Property Let HealthMinus(Value As Integer): p_Health = p_Health - Value: End Property

Public Property Let ManaPlus(Value As Integer): p_Mana = p_Mana + Value: End Property
Public Property Let ManaMinus(Value As Integer): p_Mana = p_Mana - Value: End Property

Public Property Let ExpPlus(Value As Integer): p_Exp = p_Exp + Value: End Property

Public Function AddItem(Item As Integer) As Integer
    Dim i As Long
    On Error GoTo Error
    For i = 0 To Ubound(p_InventoryPos)
        If p_InventoryPos(i) = Item Then
            p_Inventory(i) = p_Inventory(i) + 1
            AddItem = i
            Exit Function
        End If
    Next i
    For i = 0 To Ubound(p_InventoryPos)
        If p_InventoryPos(i) = 0 Then
             p_Inventory(i) = Item
             p_InventoryPos(i) = 1
             AddItem = i
             Exit Function
        End If
    Next i
    Error:
    AddItem = -1
End Function

Public Function AddEffect(Effect As Integer, Lenght As Integer)
    Dim Size As Integer
    On Error Resume Next
    Size = -1 
    Size = Ubound(p_Effects)
    For i = 0 To Size
        If p_Effects(i) = Effect Then
            p_EffectLen(i) = p_EffectLen(i) + Lenght
            AddEffect = i
            Exit Function
        End If
    Next i
    Size  = Size + 1
    ReDim Preserve p_Effects(Size)
    ReDim Preserve p_EffectLen(Size)
    p_Effects(Size) = Effect
    p_EffectLen(Size) = Lenght
    AddEffect = Size
End Function

Public Function Equip(Item As Integer, Slot As Integer)
    If AllowedSlot(Item, Slot) = True Then
        p_Setup(Slot) = Item
        Equip = Slot
    Else
        Equip = -1
    End If
End Function

Private Sub Class_Initialize()

End Sub

Private Sub Class_Terminate()

End Sub