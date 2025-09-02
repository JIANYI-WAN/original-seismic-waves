Sub TransposeData()


Dim LastRow As Long
Dim i As Long

'获取最后一行的行号
LastRow = 3254

'循环遍历每一行，将数据从横向转换为纵向
For i = 4 To LastRow
    Range("J" & Rows.Count).End(xlUp).Offset(1, 0).Resize(6, 1).Value = _
        Application.Transpose(Range("A" & i & ":F" & i).Value)
Next i

End Sub
