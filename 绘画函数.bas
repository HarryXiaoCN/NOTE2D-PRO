Attribute VB_Name = "绘画函数"
Public Function 绘制连接(绘制面)
    Dim i As Long
    For i = 0 To UBound(点) - 1
        With 点(i)
            If .去 <> "" Or .阈值.超限导向 <> "" Or .阈值.低限导向 <> "" Then
                绘制面.ForeColor = .颜色
                绘制连接_子函数 绘制面, .不应.去, i, 1, 不应导连接符号
                绘制连接_子函数 绘制面, .阈值.低限导向, i, 1, 低导连接符号
                绘制连接_子函数 绘制面, .去, i
                绘制连接_子函数 绘制面, .阈值.超限导向, i, 1, 超导连接符号
            End If
        End With
    Next
    绘制需求 = False
End Function

Public Function 二维坐标求中运算(a As 二维坐标, b As 二维坐标, 中分) As 二维坐标
    二维坐标求中运算.X = (b.X - a.X) * 中分 + a.X
    二维坐标求中运算.Y = (b.Y - a.Y) * 中分 + a.Y
End Function

Private Function 绘制连接_子函数(绘制面, 去, 本, Optional 线宽 As Long = 2, Optional 线意 As String)
    Dim 中点 As 二维坐标, 去缓存() As Long, i As Long
    去缓存 = 绘制连接_子函数_节点名转索引序(去)
    For i = 0 To UBound(去缓存)
        If 去缓存(i) < UBound(点) And 去缓存(i) >= 0 Then
            中点 = 二维坐标求中运算(点(本).坐标, 点(去缓存(i)).坐标, 0.67)
'            中点.X = (点(去缓存(i)).坐标.X - 点(本).坐标.X) / 3 * 2 + 点(本).坐标.X
'            中点.Y = (点(去缓存(i)).坐标.Y - 点(本).坐标.Y) / 3 * 2 + 点(本).坐标.Y
            If 线意 <> "" Then
                绘制面.CurrentX = 中点.X
                绘制面.CurrentY = 中点.Y
                Select Case 线意
                    Case 低导连接符号
                        绘制面.Print 线意 & "<" & 点(本).阈值.下限
                    Case 超导连接符号
                        绘制面.Print 线意 & ">" & 点(本).阈值.上限
                    Case 不应导连接符号
                        绘制面.Print 线意
                End Select
            End If
            绘制面.DrawWidth = 线宽
            绘制面.Line (点(本).坐标.X, 点(本).坐标.Y)-(中点.X, 中点.Y), 点(本).颜色
            绘制面.DrawWidth = 1
            绘制面.Line (中点.X, 中点.Y)-(点(去缓存(i)).坐标.X, 点(去缓存(i)).坐标.Y), 点(去缓存(i)).颜色
        End If
    Next
End Function

Public Function 绘制连接_子函数_节点名转索引序(ByVal 去串) As Variant
    Dim 缓存() As String, i As Long, j As Long, 序集() As Long
    去串 = 去串 & ","
    缓存 = Split(去串, ",")
    ReDim 序集(UBound(缓存) - 1)
    For i = 0 To UBound(序集)
        序集(i) = -1
        For j = 0 To UBound(点) - 1
            If 点(j).名字 = 缓存(i) Then
                序集(i) = j: Exit For
            End If
        Next
    Next
    绘制连接_子函数_节点名转索引序 = 序集
End Function

Public Function 绘制源点(绘制面)
    Dim i As Long, 源点集() As String
    源点集 = Split(启动节点, ",")
    For i = 0 To UBound(源点集) - 1
        With 点(Val(源点集(i))).坐标
            绘制面.FillColor = 源点色
            绘制面.Circle (.X, .Y), 30, 源点色
        End With
    Next
    绘制需求 = False
End Function

Public Function 绘制运算点(坐标 As 二维坐标, 大小 As Single, 颜色 As Long, Optional 绘制间隔 As Long = 400)
    体.面.FillColor = 颜色
    体.面.Circle (坐标.X, 坐标.Y), 大小, 颜色
    DoEvents
    Sleep 绘制间隔
End Function

Public Function 绘制流动运算点(起点 As 二维坐标, 终点 As 二维坐标, Optional 绘制间隔 As Long = 400, Optional 颜色 As Long = 14822282, Optional 宽度 As Long = 3)
    Dim 基本绘制间隔 As Double, 坐标 As 二维坐标
    基本绘制间隔 = 绘制间隔 / 10
    With 体
        .FillColor = 颜色
        .面.DrawWidth = 宽度
        For i = 1 To 基本绘制间隔
            坐标 = 二维坐标求中运算(起点, 终点, i / 基本绘制间隔)
            .面.Line (起点.X, 起点.Y)-(坐标.X, 坐标.Y), 颜色
            DoEvents
            Sleep 10
        Next
    End With
End Function

Public Function 绘制节点(绘制面)
    Dim i As Long
    For i = 0 To UBound(点) - 1
        With 点(i)
            绘制面.FillColor = .颜色
            绘制面.ForeColor = .颜色
            绘制面.Circle (.坐标.X, .坐标.Y), .大小, .颜色
            绘制面.CurrentX = .坐标.X + 节点名绘制横偏移长度
            绘制面.CurrentY = .坐标.Y + 节点名绘制纵偏移长度
            绘制面.Print .名字 & "=" & .权值
            绘制面.CurrentX = .坐标.X + 节点运算符横偏移长度
            绘制面.CurrentY = .坐标.Y + 节点运算符纵偏移长度
            绘制面.Print .运算
            绘制面.CurrentX = .坐标.X + 节点序号横偏移长度
            绘制面.CurrentY = .坐标.Y + 节点序号纵偏移长度
            绘制面.Print .索引 & " 遗忘:" & .遗忘.位 & "/" & .遗忘.期 & " 常应:" & .不应.常位 & "/" & .不应.常
            绘制面.CurrentX = .坐标.X + 节点序号横偏移长度
            绘制面.CurrentY = .坐标.Y + 节点序号纵偏移长度 + 250
            绘制面.Print "   不应:" & .不应.位 & "/" & .不应.期
            If .常量 Then
                绘制面.CurrentX = .坐标.X + 节点流值横偏移长度
                绘制面.CurrentY = .坐标.Y + 节点流值纵偏移长度
                绘制面.Print .流值
            End If
        End With
    Next
    绘制需求 = False
End Function
