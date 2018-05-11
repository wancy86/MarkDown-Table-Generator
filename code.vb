Sub GenTableBtn()
    Application.ScreenUpdating = False
    
    '1. data start from row #2
    Dim TabSheet
    Set TabSheet = Worksheets("Sheet1")
    
    Dim UsedRows%, UsedCols%
    UsedRows = Application.CountA(TabSheet.Range("A:A"))
    UsedCols = Application.CountA(TabSheet.Range("2:2"))
    'MsgBox (UsedRows)
    'MsgBox (UsedCols)
    
    'get max length
    Dim MaxLenArr() As Integer
    MaxLenArr = GetColMaxLength(UsedRows, UsedCols)
    
    'gen table
    Dim tabHeader$, tabstr$, tabSplit$
    
    For x = 2 To UsedRows + 1
        For y = 1 To UsedCols
            t = Cells(x, y).Value
            
            'MsgBox (MaxLenArr(y - 1) - Len(t))
            'MsgBox (Len(t))
            tabstr = tabstr & "|" & t & String((MaxLenArr(y - 1) - Len(t)), " ")
            If x = 2 Then
                tabSplit = tabSplit & "|:" & String((MaxLenArr(y - 1) - 1), "-")
            End If
        Next y
        If x = 2 Then
            tabHeader = tabstr & "|" & vbLf
            tabstr = ""
        Else
            tabstr = tabstr & "|" & vbLf
        End If
    Next x
    
    tabstr = tabHeader & tabSplit & "|" & vbLf & tabstr
    Worksheets("Sheet2").Cells(1, 1).Value = tabstr

    'send to clipboard
    Dim MyData2 As New DataObject
    Dim clip$
    clip = "123"
    MyData2.SetText clip
    MyData2.PutInClipboard

End Sub


'max length of column save to array
Function GetColMaxLength(UsedRows As Integer, UsedCols As Integer) As Integer()
    Dim maxlen As Integer
    Dim arr() As Integer
    ReDim arr(0 To UsedCols) 'rang array
        
    For y = 1 To UsedCols
        maxlen = 0
        For x = 2 To UsedRows + 1
           If Len(Worksheets("Sheet1").Cells(x, y).Value) > maxlen Then
            maxlen = Len(Worksheets("Sheet1").Cells(x, y).Value)
           End If
            'MsgBox (Worksheets("Sheet1").Cells(x, y).Value)
            'MsgBox (maxlen)
        Next x
        
        arr(y - 1) = maxlen
    Next y
    
    'set return
    GetColMaxLength = arr
End Function



