VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EmailOutlook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' EmailOutlook - Simple interface to send emails from Outlook.

Private pOutlook As Object
Private pDraft As Object

Private pFrom As String
Private pTo As String
Private pCC As String
Private pBCC As String
Private pHeader As String   ' internal use
Private pBody As String
Private pFooter As String   ' internal use
Private pSignature As String

Private pFontName As String

Enum EmailBodyFontType
    Plain = 1
    RichText = 2
    HTML = 3
End Enum

Private pBodyType As EmailBodyFontType


Private pInit As Boolean

Public Function Init()
    If Not IsInit Then
        
        Set pOutlook = CreateObject("Outlook.Application")
        Set pDraft = pOutlook.CreateItem(0)
        
        pBodyType = Plain
        
        pInit = True
    End If
End Function

Private Function IsInit() As Boolean
    IsInit = pInit
End Function

' Boiler plate

Public Property Get FromField() As String
    If IsInit Then
        FromField = pFrom
    End If
End Property
Public Property Let FromField(value As String)
    If IsInit Then
        pFrom = value
    End If
End Property

Public Property Get ToField() As String
    If IsInit Then
        ToField = pTo
    End If
End Property
Public Property Let ToField(value As String)
    If IsInit Then
        pTo = value
    End If
End Property

Public Property Get CCField() As String
    If IsInit Then
        CCField = pCC
    End If
End Property
Public Property Let CCField(value As String)
    If IsInit Then
        pCC = value
    End If
End Property

Public Property Get BCCField() As String
    If IsInit Then
        BCCField = pBCC
    End If
End Property
Public Property Let BCCField(value As String)
    If IsInit Then
        pBCC = value
    End If
End Property

Public Property Get Body() As String
    If IsInit Then
        Body = pBody
    End If
End Property
Public Property Let Body(value As String)
    If IsInit Then
        pBody = value
    End If
End Property

Public Property Get Signature() As String
    If IsInit Then
        Signature = pSignature
    End If
End Property
Public Property Let Signature(value As String)
    If IsInit Then
        pSignature = value
    End If
End Property

Public Property Get FontName() As String
    If IsInit Then
        FontName = pFontName
    End If
End Property
Public Property Let Signature(value As String)
    If IsInit Then
        pFontName = value
    End If
End Property

' End Boiler plate

Public Function Send()
    If IsInit Then
        
    End If
End Function
























