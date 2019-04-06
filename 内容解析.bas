Attribute VB_Name = "内容解析"
Public Function 节点内容初始化解析(内容, 节点号)
    Dim 行() As String, i As Long
    节点重置 节点号
    行 = Split(内容, vbCrLf)
    For i = 0 To UBound(行)
        节点内容初始化解析_子函数 行(i), 节点号
    Next
    绘制需求 = True
End Function
Private Function 节点重置(i)
    Dim 新节点 As 节点
    新节点.名字 = 点(i).名字
    新节点.内容 = 点(i).内容
    新节点.坐标 = 点(i).坐标
    新节点.颜色 = 点(i).颜色
    新节点.大小 = 点(i).大小
    点(i) = 新节点
End Function
Private Function 节点内容初始化解析_子函数(行, 节点号)
    Dim 行头 As String
    On Error GoTo Er
    If InStr(1, 行, " ") > 0 Then
        行头 = UCase(Split(行, " ")(0))
    Else
        行头 = UCase(行)
    End If
    Select Case 行头
        Case "去", "Q", "QU"
            点(节点号).去 = Split(行, " ")(1)
        Case "值", "Z"
            点(节点号).权值 = Val(Split(行, " ")(1))
        Case "算", "S"
            点(节点号).运算 = Split(行, " ")(1)
        Case "上限", "SX"
            点(节点号).阈值.上限 = Val(Split(行, " ")(1))
        Case "下限", "XX"
            点(节点号).阈值.下限 = Val(Split(行, " ")(1))
        Case "常", "C"
            点(节点号).常量 = True
    End Select
Er:
End Function
