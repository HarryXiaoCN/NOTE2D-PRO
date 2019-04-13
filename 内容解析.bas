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
    新节点.索引 = 点(i).索引
    新节点.编辑界面偏移 = 点(i).编辑界面偏移
    点(i) = 新节点
End Function

Private Function 节点内容初始化解析_子函数_随机赋值(行, 节点号)
    Dim 命令() As String, 子命令() As String
    命令 = Split(行, " ")
    Select Case UCase(命令(1))
        Case "R", "随机数", "SJS", "SJ", "S", "RND"
            子命令 = Split(命令(2), ",")
            Randomize Val(子命令(0))
            点(节点号).权值 = Rnd * Val(子命令(1)) + Val(子命令(2))
        Case Else
            点(节点号).权值 = Val(命令(1))
    End Select
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
            点(节点号).去缓存 = 绘制连接_子函数_节点名转索引序(点(节点号).去)
        Case "值", "Z"
            节点内容初始化解析_子函数_随机赋值 行, 节点号
    End Select
Er:
End Function
