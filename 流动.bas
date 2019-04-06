Attribute VB_Name = "流动"
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Function 启动流(表 As String)
    Dim 节点表() As Long, i As Long
    体.绘制时钟.Enabled = False
    字符串转整数表 节点表, 表, ","
    For i = 0 To UBound(节点表)
        流 节点表(i)
    Next
    绘制需求 = True
    体.绘制时钟.Enabled = True
End Function

Public Function 流(起点 As Long)
    Dim 节点表() As Long, i As Long
    节点表 = 绘制连接_子函数_节点名转索引序(点(起点).去)
    For i = 0 To UBound(节点表)
        If 节点表(i) >= 0 Then
            If 流运算(点(起点), 点(节点表(i))) Then
                体.绘制周期
                体.面.FillColor = 点(起点).颜色
                体.面.Circle (点(节点表(i)).坐标.X, 点(节点表(i)).坐标.Y), 点(节点表(i)).大小 / 2, 点(起点).颜色
                DoEvents
                Sleep 400
                流 节点表(i)
            End If
        End If
    Next
End Function

Public Function 流运算(源点 As 节点, 去点 As 节点) As Boolean
    Dim 源值 As Double
    源值 = 节点流汇入算术值获取(源点)
    If (源值 >= 去点.阈值.下限 And 源值 <= 去点.阈值.上限) Or (去点.阈值.下限 = 0 And 去点.阈值.上限 = 0) Then
        节点流汇入运算 节点流汇入算术值获取(源点), 去点
        流运算 = True
    End If
End Function

Public Function 节点流汇入运算(源值, 去点 As 节点)
    If 去点.常量 Then
        去点.流值 = 运算(源值, 去点.权值, 去点.运算)
    Else
        去点.权值 = 运算(源值, 去点.权值, 去点.运算)
    End If
End Function

Public Function 节点流汇入算术值获取(源点 As 节点) As Double
    If 源点.常量 Then
        节点流汇入算术值获取 = 源点.流值
    Else
        节点流汇入算术值获取 = 源点.权值
    End If
End Function

Public Function 运算(a, b, f) As Double
    On Error GoTo Er
        Select Case f
            Case "+"
                运算 = a + b
            Case "-"
                运算 = a - b
            Case "--"
                运算 = b - a
            Case "*"
                运算 = a * b
            Case "/"
                运算 = a / b
            Case "//"
                运算 = b / a
            Case "="
                运算 = a
        End Select
Er:
End Function
