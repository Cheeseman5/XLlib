VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EmailOutlook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private Const pDEBUG = True

Private pOutlook As Object
Private pDraft As Object


Private Const NEW_LINE = "<br>"
Private Const CHAR_TAB = "&emsp;" '"&nbsp &nbsp &nbsp &nbsp "

Private pFrom As String
Private pTo As String
Private pCC As String
Private pBCC As String
Private pSubject As String
Private pHeader As String   ' internal use
Private pBody As String
Private pFooter As String   ' internal use
Private pSignature As String

Private pFontName As String
Private pInitialSig As String


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
        
        pBodyType = HTML
        
        pInit = True
    End If
End Function

Private Function IsInit() As Boolean
    IsInit = pInit
End Function

Private Function HasOutlook() As Boolean
    HasOutlook = Not pOutlook Is Nothing
End Function

Private Function HasDraft() As Boolean
    HasDraft = Not pDraft Is Nothing
End Function

Private Function IsDebugging() As Boolean
    IsDebugging = pDEBUG
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

Public Property Get Subject() As String
    If IsInit Then
        Subject = pSubject
    End If
End Property
Public Property Let Subject(value As String)
    If IsInit Then
        pSubject = value
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
Public Property Let FontName(value As String)
    If IsInit Then
        pFontName = value
    End If
End Property

' End Boiler plate

Public Function TabChar(Optional count = 1) As String
    TabChar = "<span style='mso-tab-count:" & count & "'>&nbsp;</span>"
End Function

Private Function ConstructEmail()
    If IsInit And HasDraft Then
        
        ' Notes on changing the "From" field
        '   pOutlook.Session.Accounts(index)
        '   pOutlook.SendUsingAccount = pOutlook.Session.Accounts(1)
        '   pOutlook.SentOnBehalfOfName = """SenderName"" <Reply@Address.com>"
        
        With pDraft
            .To = ToField
            .CC = CCField
            .BCC = BCCField
            .Subject = pSubject
            
            ' header
            pHeader = "<font face=" & FontName & ">"
            ' footer
            pFooter = "</font>"
            
            Select Case pBodyType
                Case EmailBodyFontType.HTML
                    .HTMLBody = pHeader & Body & pFooter & Signature
                Case Else
                    '.Body = pHeader & Body & pFooter & Signature
                    .Body = Body & Signature
            End Select
            
        End With
        
    End If
End Function

Public Function Send()
    If IsInit Then
        ConstructEmail
        
        ' HACK: for testing only!!
        Display
        'Destroy
    End If
End Function

Public Function Display()
    If IsInit And (Not pDraft Is Nothing) Then
        pDraft.Display
        
        ConstructEmail
        
        Select Case pBodyType
            Case EmailBodyFontType.HTML
                pInitialSig = pDraft.HTMLBody
            Case Else
                pInitialSig = pDraft.Body
        End Select
    End If
End Function

Public Function GetSignature(signatureName As String) As String
    Dim fso As Object
    Dim ts As Object
    Dim fileName As String
    Dim str As String
    
    fileName = Environ("appdata") & "\Microsoft\Signatures\" & signatureName & ".htm"
    
    If Dir(fileName) <> "" Then
    
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set ts = fso.GetFile(fileName).OpenAsTextStream(1, -2)
        str = ts.ReadAll
        
        If Not ts Is Nothing Then ts.Close
    Else
        MsgBox "Signature """ & signatureName & """ not found!  Please check the name." & vbCrLf & _
            "Using empty signature..."
        str = ""
    End If
    
    GetSignature = str
End Function

' TODO: Make public once a good fix is found.  Currently needs to
'       run pDraft.Display in order to copy signature
Private Function GetMsgSig() As String
    If IsInit And HasDraft Then
    
        GetMsgSig = pInitialSig
        
    End If
End Function



Public Function Destroy()
    Set pOutlook = Nothing
    Set pDraft = Nothing
End Function






















