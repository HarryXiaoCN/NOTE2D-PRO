Attribute VB_Name = "流动"
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Function 启动流(表 As String)
    Dim 节点表() As Long, i As Long
    If 表 <> "" Then
        体.绘制时钟.Enabled = False
        字符串转整数表 节点表, 表, ","
        For i = 0 To UBound(节点表)
            流 节点表(i)
            体.绘制周期
        Next
        绘制需求 = True
        体.绘制时钟.Enabled = True
    End If
End Function

Public Function 流(起点 As Long)
    Dim 节点表() As Long
    节点表 = 绘制连接_子函数_节点名转索引序(点(起点).去)
    流_子函数 起点, 节点表
End Function

Public Function 流_子函数(起点 As Long, 节点表() As Long, Optional 非正常流 As Boolean)
    Dim i As Long
    For i = 0 To UBound(节点表)
        If 节点表(i) >= 0 Then
            If 流运算(点(起点), 点(节点表(i)), 非正常流) Then
                流 节点表(i)
            End If
        End If
    Next
End Function

Public Function 流运算_阈值导向子函数(源点 As 节点, 去点 As 节点, 导向串 As String)
    Dim 节点集() As Long
    节点集 = 绘制连接_子函数_节点名转索引序(导向串)
    绘制流动运算点 源点.坐标, 去点.坐标, 100
    流_子函数 去点.索引, 节点集, True
End Function

Public Function 流运算(源点 As 节点, 去点 As 节点, 非正常流 As Boolean) As Boolean
    Dim 源值 As Double, 去值 As Double
    源值 = 节点流汇入算术值获取(源点)
    Debug.Print 源点.名字; 去点.名字
    If 去点.不应.位 <= 0 Then
        If 去点.不应.常位 > 0 Then
            去点.不应.常位 = 去点.不应.常位 - 1
        Else
            去点.不应.常位 = 去点.不应.常
            去点.不应.位 = 去点.不应.期
        End If
        If 去点.阈值.低限导向 <> "" And 源值 < 去点.阈值.下限 And (去点.阈值.下限 <> 0 And 去点.阈值.上限 <> 0) Then
            流运算_阈值导向子函数 源点, 去点, 去点.阈值.低限导向
        End If
        If (去点.自锁.次 = 0 Or 去点.自锁.次 > 去点.激活数) And (非正常流 = True Or (源值 >= 去点.阈值.下限 And 源值 <= 去点.阈值.上限) Or (去点.阈值.下限 = 0 And 去点.阈值.上限 = 0)) Then
            体.绘制周期
            绘制流动运算点 源点.坐标, 去点.坐标, 100
            节点流汇入运算 节点流汇入算术值获取(源点), 去点
            去点.激活数 = 去点.激活数 + 1
            去值 = 节点流汇入算术值获取(去点)
            If (去值 >= 去点.阈值.输出下限 And 去值 <= 去点.阈值.输出上限) Or (去点.阈值.输出下限 = 0 And 去点.阈值.输出上限 = 0) Then
                流运算 = True
            End If
        End If
        If 去点.阈值.超限导向 <> "" And 源值 > 去点.阈值.上限 And (去点.阈值.下限 <> 0 Or 去点.阈值.上限 <> 0) Then
            流运算_阈值导向子函数 源点, 去点, 去点.阈值.超限导向
        End If
    Else
        去点.不应.位 = 去点.不应.位 - 1
        流运算_阈值导向子函数 源点, 去点, 去点.不应.去
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
                运算 = b + a
            Case "-"
                运算 = b - a
            Case "--"
                运算 = a - b
            Case "*"
                运算 = b * a
            Case "/"
                运算 = b / a
            Case "//"
                运算 = a / b
            Case "="
                运算 = a
            Case "=="
                运算 = b
        End Select
Er:
End Function
