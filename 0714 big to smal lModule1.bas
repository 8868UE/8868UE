Attribute VB_Name = "Module1"
Sub vbafinal()
Attribute vbafinal.VB_Description = "�j��p"
Attribute vbafinal.VB_ProcData.VB_Invoke_Func = "q\n14"
'
' vbafinal ����
' ����
'
' �ֳt��: Ctrl+q
'
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Add2 Key:=Range("B2:B414") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("�u�@��1").Sort
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
Attribute vbafinal1.VB_Description = "�p��j"
Attribute vbafinal1.VB_ProcData.VB_Invoke_Func = "w\n14"
'
' vbafinal1 ����
' �p��j
'
' �ֳt��: Ctrl+w
'
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Add2 Key:=Range("B2:B414") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("�u�@��1").Sort
        .SetRange Range("A1:B414")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
