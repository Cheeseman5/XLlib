VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CellView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' CellView - handles basic I/O interaction with the smallest UI element (a cell).
'
Private pRg As Range

Public Property Get Pos() As Range
    Set Pos = pRg
End Property
Public Property Let Pos(Value As Range)
    Set pRg = Value
End Property

Private Function IsInit() As Boolean
    IsInit = (Not pRg Is Nothing)
End Function

' Value (Content)
Public Property Get Content() As String
    If IsInit Then
        Content = Pos.Value
    End If
End Property
Public Property Let Content(Value As String)
    If IsInit Then
        Pos.Value = Value
    End If
End Property

' Formula
Public Property Get Formula() As String
    If IsInit Then
        Formula = Pos.Formula
    End If
End Property
Public Property Let Formula(Value As String)
    If IsInit Then
        Pos.Formula = Value
    End If
End Property

' Value
Public Property Get Value() As String
    If IsInit Then
        Value = Pos.Value
    End If
End Property
Public Property Let Value(val As String)
    If IsInit Then
        Pos.Value = val
    End If
End Property

' Text
Public Property Get Text() As String
    If IsInit Then
        Text = Pos.Text
    End If
End Property

' Address
Public Property Get Address() As String
    If IsInit Then
        Address = Pos.Address
    End If
End Property

' NumberFormat
Public Property Get NumberFormat() As String
    If IsInit Then
        NumberFormat = Pos.NumberFormat
    End If
End Property
Public Property Let NumberFormat(Value As String)
    If IsInit Then
        Pos.NumberFormat = Value
    End If
End Property

