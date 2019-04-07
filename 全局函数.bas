Attribute VB_Name = "全局函数"
Public Function 新建点(X, Y, Optional N As String)
    Dim 节点名缓存 As String
    On Error GoTo Er
        If N = "" Then
            节点名缓存 = 节点默认前缀 & UBound(点)
        Else
            节点名缓存 = N
        End If
        If 节点名重复性检测(节点名缓存, UBound(点)) = False Then
            With 点(UBound(点))
                .名字 = 节点名缓存
                .大小 = 100
                .索引 = UBound(点)
                .内容 = 节点默认内容
                .坐标.X = X: .坐标.Y = Y
                .不应.常 = 1
                .颜色 = 节点默认颜色
                .编辑界面偏移.X = 节点编辑界面横偏移长度
                .编辑界面偏移.Y = 节点编辑界面纵偏移长度
            End With
            绘制需求 = True
            ReDim Preserve 点(UBound(点) + 1)
        Else
            节点名缓存 = InputBox("自动生成节点名已存在，请指定新节点名字：", "新建点", 节点名缓存)
            新建点 X, Y, 节点名缓存
        End If
Er:
End Function

Public Function 字符串转整数表(目标表, 表, 分割符)
    Dim 缓存() As String, i As Long, j As Long
    缓存 = Split(表, 分割符)
    For i = 0 To UBound(缓存)
        If 缓存(i) <> "" Then
            j = j + 1
        End If
    Next
    ReDim 目标表(j - 1)
    j = 0
    For i = 0 To UBound(缓存)
        If 缓存(i) <> "" Then
            目标表(j) = Val(缓存(i))
            j = j + 1
        End If
    Next
End Function

Public Function 区域节点检测(X, Y, Optional 圆心距 As Single = 100) As Long
    Dim i As Long, 距离长 As Double
    区域节点检测 = -1
    On Error GoTo Er
        For i = 0 To UBound(点) - 1
            With 点(i)
                距离长 = (X - .坐标.X) ^ 2 + (Y - .坐标.Y) ^ 2
                If 距离长 < (圆心距 + .大小 + 50) ^ 2 Then
                   区域节点检测 = i: Exit Function
                End If
            End With
        Next
        Exit Function
Er:
    Debug.Print "全局函数[区域节点检测] - 错误！", Err.Description
End Function

Public Function 节点名重复性检测(N, id) As Boolean
    Dim i As Long
    For i = 0 To UBound(点) - 1
        If 点(i).名字 = N And id <> i Then
            节点名重复性检测 = True: Exit Function
        End If
    Next
End Function
