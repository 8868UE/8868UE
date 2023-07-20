Attribute VB_Name = "Module1"
Sub vbafinal()
Attribute vbafinal.VB_Description = "大到小"
Attribute vbafinal.VB_ProcData.VB_Invoke_Func = "q\n14"
'
' vbafinal 巨集
' 偵測
'
' 快速鍵: Ctrl+q
'
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Add2 Key:=Range("B2:B414") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("工作表1").Sort
        .SetRange Range("A1:B414")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWindow.SmallScroll Down:=-15
    ActiveWorkbook.Save
End Sub
Sub vbafinal1()
Attribute vbafinal1.VB_Description = "小到大"
Attribute vbafinal1.VB_ProcData.VB_Invoke_Func = "w\n14"
'
' vbafinal1 巨集
' 小到大
'
' 快速鍵: Ctrl+w
'
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("工作表1").Sort.SortFields.Add2 Key:=Range("B2:B414") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("工作表1").Sort
        .SetRange Range("A1:B414")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
