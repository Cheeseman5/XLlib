Attribute VB_Name = "Utils"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Utils
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function AnounceError(err As Variant, Optional message As String)
    If IsObject(err) Then
        If Not err Is Nothing Then
            MsgBox "Message: " & message & vbNewLine & err.source & vbNewLine & err.Description
        End If
    Else
        MsgBox "Message: " & message & vbNewLine & "Error code: " & err & vbNewLine & "Error: " & Error$(err)
    End If
    
    'Edit: this call should be up to the user, not run by default
    'Utils.EnableDrawing True
End Function

Function GetSetting(sh As Worksheet, settingName As String) As String
    If Not Utils.Debugging Then
        On Error GoTo ErrEnter
    End If
    
    If (Not sh Is Nothing) And (settingName <> "") Then
        GetSetting = sh.Range("A1").Find(settingName).Offset(0, 1).Text
        
    End If
    
ErrExit:
    Exit Function

ErrEnter:
    GetSetting = ""
    err.Clear
    GoTo ErrExit
End Function

Function ChangeSetting(sh As Worksheet, settingName As String, settingValue As String)
    If Not Utils.Debugging Then
        On Error GoTo ErrEnter
    End If
    
    If (Not sh Is Nothing) And (settingName <> "") Then
        sh.Range("A1").Find(settingName).Offset(0, 1).Value = settingValue
    End If
    
ErrExit:
    Exit Function

ErrEnter:
    err.Clear
    GoTo ErrExit
End Function

' Searches <Ws> for both header names and then their values (in sequential order) until both values are found.
' Returns <srchValue2> Range if both found and 'Nothing' if nothing found.
Function Find(Ws As Worksheet, hdrName1 As String, srchValue1 As String, hdrName2 As String, srchValue2 As String) As Range
    If Not Utils.Debugging Then
        On Error GoTo ErrEnter
    End If
    
    Dim rgResult As Range
    Dim rgHdr1 As Range
    Dim rgHdr2 As Range
    Dim rgSearching As Range
    Dim oldRow As Integer
    
    Set rgHdr1 = Ws.Range("A1").EntireRow.Find(hdrName1, LookAt:=xlWhole)
    Set rgHdr2 = Ws.Range("A1").EntireRow.Find(hdrName2, LookAt:=xlWhole)
    
    ' return Nothing if either header isn't found
    If (rgHdr1 Is Nothing) Or (rgHdr2 Is Nothing) Then
        Set Find = Nothing
        Exit Function
    End If
    
    Set rgSearching = rgHdr1.EntireColumn.Find(srchValue1, LookAt:=xlWhole)
    
    ' return Nothing if 1st value doesn't exist
    If rgSearching Is Nothing Then
        Set Find = Nothing
        Exit Function
    End If
    
    Do
        oldRow = rgSearching.row
        If Not rgSearching Is Nothing Then
            If rgSearching.Offset(0, rgHdr2.Column - rgSearching.Column).Value = srchValue2 Then
                Set Find = rgSearching.Offset(0, rgHdr2.Column - rgSearching.Column)
                Exit Function
            End If
        Else
            Exit Do
        End If
        
        Set rgSearching = rgHdr1.EntireColumn.FindNext(After:=rgSearching)
        
    Loop While (Not rgSearching Is Nothing) And (oldRow < rgSearching.row) ' while rgSearching != nothing
    
ErrExit:
    Set Find = Nothing
    Exit Function

ErrEnter:
    Utils.AnounceError (err)
    GoTo ErrExit
End Function

' Get the row number of the lowest cell (largest row number) with data in it.
Function GetLastUsedRow(sh As Worksheet) As Integer
    If Not Utils.Debugging Then
        On Error GoTo ErrEnter
    End If
    
    If Not sh Is Nothing Then
        GetLastUsedRow = sh.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
    Else
        GetLastUsedRow = 0
    End If
    
ErrExit:
    Exit Function

ErrEnter:
    Utils.AnounceError (err)
    GoTo ErrExit
End Function

' Enables/Disables drawing
Function EnableDrawing(enabled As Boolean)
    Application.ScreenUpdating = enabled
    Application.EnableEvents = enabled
    Application.StatusBar = enabled
    
    If enabled Then
        Application.Calculation = xlCalculationAutomatic
    Else
        Application.Calculation = xlCalculationManual
    End If
    
End Function


Function HitWithWrench()
    Utils.EnableDrawing True
    
    ' Add any other code to reset current project...
End Function

Function Floor(Value As Double) As Double
    If Not Utils.Debugging Then
        On Error GoTo ErrEnter
    End If
    
    Floor = Int(Value)
    
ErrExit:
    Exit Function

ErrEnter:
    Utils.AnounceError (err)
    GoTo ErrExit
End Function

Function StrDate(d As Date) As String
    If Not Utils.Debugging Then
        On Error GoTo ErrEnter
    End If
    
    StrDate = Format(d, "mm/dd/yyyy")
    
ErrExit:
    Exit Function

ErrEnter:
    Utils.AnounceError (err)
    GoTo ErrExit
End Function

Function Debugging() As Boolean
    Debugging = False
End Function

Public Function NullString() As String
    NullString = ""
End Function

Public Function NullDate() As Date
    NullDate = -1
End Function

Public Function NullDouble() As Double
    NullDouble = -1
End Function

