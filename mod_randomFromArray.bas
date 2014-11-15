Attribute VB_Name = "mod_randomFromArray"

Sub randomFromArray()
    If Selection.Count = 0 Then Exit Sub
    
    Dim str
    str = InputBox("strings to be used?(,)", , "X,Y,Z")
    
    arr = Split(str, ",")
    Dim width
    width = UBound(arr) + 1
    
    If Selection.Count = 1 Then
        Call randomFromArray_col(arr, width)
    Else
        Call randomFromArray_selected(arr, width)
    End If
End Sub

Sub randomFromArray_col(arr, width)
    Dim cosuu
    cosuu = InputBox("How many cells on current column to be randomed?", , 10)
        
    Dim i, y, x
    x = ActiveCell.Column
    y = ActiveCell.Row
    
    For i = 1 To cosuu
        Cells(y, x).Value = arr(Int(Rnd() * width))
        y = y + 1
    Next i
End Sub

Sub randomFromArray_selected(arr, width)
    Dim cl
    For Each cl In Selection
        cl.Value = arr(Int(Rnd() * width))
    Next cl
End Sub
