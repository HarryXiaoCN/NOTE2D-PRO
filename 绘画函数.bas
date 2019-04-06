Attribute VB_Name = "绘画函数"
Public Function 绘制连接(绘制面)
    Dim i As Long
    For i = 0 To UBound(点) - 1
        With 点(i)
            If .去 <> "" Then
                绘制连接_子函数 绘制面, .去, i
            End If
        End With
    Next
    绘制需求 = False
End Function

Private Function 绘制连接_子函数(绘制面, 去, 本)
    Dim 中点 As 二维坐标, 去缓存() As Long, i As Long
    去缓存 = 绘制连接_子函数_节点名转索引序(去)
    For i = 0 To UBound(去缓存)
        If 去缓存(i) < UBound(点) And 去缓存(i) >= 0 Then
            中点.X = (点(去缓存(i)).坐标.X - 点(本).坐标.X) / 3 * 2 + 点(本).坐标.X
            中点.Y = (点(去缓存(i)).坐标.Y - 点(本).坐标.Y) / 3 * 2 + 点(本).坐标.Y
            绘制面.DrawWidth = 2
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
            绘制面.Print i
            If .常量 Then
                绘制面.CurrentX = .坐标.X + 节点流值横偏移长度
                绘制面.CurrentY = .坐标.Y + 节点流值纵偏移长度
                绘制面.Print .流值
            End If
        End With
    Next
    绘制需求 = False
End Function
