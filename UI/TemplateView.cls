VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TemplateView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' TemplateView - general template class to copy to new views.
'
Option Explicit

' Replace 'Variant' with actual parent class
Private pParentView As Variant

Private pSett As Settings

' position relative to parent
Private pRg As Range

Private pViewWidth As Integer
Private pViewHeight As Integer

Private pInit As Boolean

' replace 'Variant' with actual parent class
Public Function Init(parentView As Variant)
    If Not pInit Then
        Set pParentView = parentView
        Set pSett = Singletons.GetSettings
        
        pViewWidth = 1
        pViewHeight = 1
    End If
    pInit = True
End Function

'----[ Utils ]---------------------------------------------------
Private Function IsInit() As Boolean
    IsInit = (pInit And (Not pParentView Is Nothing))
End Function

'----[ Position ]---------------------------------------------
' SetPosition usage - used from 'pParentView'; usage snippet:
'
'        ' array of "day" views
'        Private pDayViews(6) As DayView
'
'        ' simple counter
'        Dim d As Integer
'
'        ' Set position of template view
'        Set pRg = rgPosition.Cells(1, 1)
'
'        ' <init/populate pDayViews>
'
'        ' Iterate through all days of the week
'        For d = 0 To 6
'            '                          <current pos> + <view width> + <DayView width>
'            pDayViews(d).SetPosition pRg.Cells(1, 1 + pViewWidth + pDayViews(d).OffsetWidth(d))
'        Next d
'
' rgPosition is the upper-left most cell of this view
Public Function SetPosition(rgPosition As Range)
    If pInit And (Not rgPosition Is Nothing) Then
        Set pRg = rgPosition.Cells(1, 1)
    End If
End Function

' Returns this view's width multiplied by offsetIndex.
Public Function OffsetWidth(offsetIndex As Integer) As Integer
    If IsInit Then
        OffsetWidth = pViewWidth * offsetIndex
    End If
End Function

' Returns this view's height multiplied by offsetIndex.
Public Function OffsetHeight(offsetIndex As Integer) As Integer
    If IsInit Then
        OffsetHeight = pViewHeight * offsetIndex
    End If
End Function

'----[ Properties ]-----------------------------------------------

' Properties...

'----[ Property Utils ]-------------------------------------------

' Property Utils...

'----[ View Formatting ]------------------------------------------

' View Formatting...

'----[ Utils ]----------------------------------------------------

' Utils...

'----[ Methods ]--------------------------------------------------

' Methods...

