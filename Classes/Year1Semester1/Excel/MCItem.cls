VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MCItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Module: Item
Private dName As String
Private dID As Integer
Private dMetaData As Integer
Private dDurability As Integer
Private dDamage As Integer
Private dIcon As Icon
Private dType As String

' Properties
Property Get Name() As String
    Name = dName
End Property

Property Let Name(daName As String)
    dName = daName
End Property

Property Get ID() As Integer
    ID = dID
End Property

Property Let ID(daID As Integer)
    dID = daID
End Property

Property Get MetaData() As Integer
    MetaData = dMetaData
End Property

Property Let MetaData(daMetaData As Integer)
    dMetaData = daMetaData
End Property

Property Get Durability() As Integer
    Durability = dDurability
End Property

Property Let Durability(daDurability As Integer)
    dDurability = daDurability
End Property

Property Get Damage() As Integer
    Damage = dDamage
End Property

Property Let Damage(daDamage As Integer)
    dDamage = daDamage
End Property

Property Get Icon() As Icon
    Icon = dIcon
End Property

Property Let Icon(daIcon As Icon)
    dIcon = daIcon
End Property

Property Get Type() As String
    Type = dType
End Property

Property Let Type(daType As String)
    dType = daType
End Property

Private Sub Class_Initialize()
    dName = "Stone"
End Sub
