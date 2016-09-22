' Simple way of finding a value within a Range
Public Function FindValInRange(strToFind As String, rgToSearch As Range) As Range
    Dim val As Range
    If Not rgToSearch Is Nothing Then
        Set val = rgToSearch.EntireColumn.Find(strToFind, _
                                            LookAt:=xlPart, _
                                            LookIn:=xlValues, _
                                            SearchOrder:=xlByColumns)
        If Not val Is Nothing Then
            Set FindValInRange = val
            Set val = Nothing
            Exit Function
        End If
    End If
    
    Set val = Nothing
    Set FindValInRange = Nothing
End Function

' SheetExists - Checks if a worksheet with a specified name exists within the given workbook.
Public Function SheetExists(wb As Workbook, shName As String, Optional isCaseSensitive = False) As Boolean
    Dim compMethod
    
    If isCaseSensitive Then
        compMethod = vbBinaryCompare
    Else
        compMethod = vbTextCompare
    End If
    
    For Each sh In wb.Worksheets
        If StrComp(sh.Name, shName, compMethod) = 0 Then
            SheetExists = True
            Exit Function
        End If
    Next
End Function

' Check if a file exists
Public Function FileExists(fName As String) As Boolean
    Dim fso As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    
    fName = ParseFPath(fName)
    
    FileExists = fso.FileExists(fName)
    
    Set fso = Nothing
End Function

' Returns file name or "" if no file found.
Public Function FileName(fName As String) As String
    Dim fso As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    
    fName = ParseFPath(fName)
    
    If fso.FileExists(fName) Then
        FileName = fso.GetFileName(fName)
    Else
        FileName = ""
    End If
    
    Set fso = Nothing
End Function

Public Function ParseFPath(fName As String) As String
    If StrComp(".\", Left(fName, 2)) = 0 Then
        fName = Replace(fName, ".\", ActiveWorkbook.path & "\")
    End If
    
    ParseFPath = fName
End Function

' Opens or gets the instance of a workbook
Public Function GetWorkbook(fName As String) As Workbook
    Dim wbName As String
    Dim fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    wbName = fso.GetBaseName(ParseFPath(fName))
    If IsWorkbookOpen(wbName) Then
        Set GetWorkbook = Workbooks(wbName)
    Else
        Set GetWorkbook = Workbooks.Open(ParseFPath(fName))
    End If
    
    Set fso = Nothing
End Function

'Returns TRUE if the workbook is open
Public Function IsWorkbookOpen(strWorkBookName As String) As Boolean
    Dim wb As Workbook

    On Error Resume Next
    ' Check if workbook is open
    Set wb = Workbooks(strWorkBookName)
    If wb Is Nothing Then
        IsWorkbookOpen = False
    Else
        IsWorkbookOpen = True
        Set wb = Nothing
    End If
End Function

' LastUsedRow - Returns last/bottom row with data in it or 0 (zero) if nothing found.
Public Function LastUsedRow(sh As Worksheet) As Integer
    Dim rg As Range
    Set rg = sh.cells.Find("*", _
                            After:=sh.cells(1), _
                            LookAt:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlPrevious)
    
    If Not rg Is Nothing Then
        LastUsedRow = rg.row
    Else
        LastUsedRow = 0
    End If
End Function

' FindHeader - returns the header's Range object or Nothing if not found
' Param <r>: row to search. Default 1.
Public Function FindHeader(sh As Worksheet, hdrName As String, Optional r As Integer = 1) As Range
    If Not sh Is Nothing Then
        Set FindHeader = sh.cells(r).EntireRow.Find(hdrName, LookIn:=xlValues, SearchOrder:=xlByRows)
    Else
        Set FindHeader = Nothing
    End If
End Function

Public Function ExistsInCollection(list As Collection, val As Variant) As Boolean
    On Error Resume Next
    
    list.item val
    ExistsInCollection = (Err.Number = 0)
    
    Err.Clear
    On Error GoTo 0
End Function

Public Function CollectionGetItemSafe(list As Collection, val As Variant) As Variant
    On Error GoTo ErrEnter

    Dim itm As Variant
    
    itm = list.item(val)
    CollectionGetItemSafe = itm
    Exit Function
    
ErrEnter:
    CollectionGetItemSafe = Null
    Err.Clear
End Function


Public Function IsArrayEmpty(arr As Variant) As Boolean
    On Error Resume Next
    If ArrayLenth(arr) = 0 Then
        IsArrayEmpty = True
    Else
        IsArrayEmpty = False
    End If
    On Error GoTo 0
End Function

Public Function ArrayLenth(arr As Variant) As Integer
    On Error GoTo ErrorEnter
    
    ArrayLenth = UBound(arr) - LBound(arr) + 1
    
ErrorExit:
    On Error GoTo 0
    Exit Function
ErrorEnter:
    ArrayLenth = 0
    Err.Clear
    Resume ErrorExit
End Function
