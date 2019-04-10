Attribute VB_Name = "命令系统"
Public Function 命令输入(命令框 As TextBox, 反馈框 As RichTextBox)
    Dim 命令() As String
    On Error GoTo Er
        命令 = Split(命令框.Text, " ")
        Select Case 命令(0)
            Case "存"
                If 存为二进制文件(命令(1), 命令框, 反馈框) = "" Then
                    GoTo Su
                End If
            Case "取"
                取得二进制文件 命令(1)
                Debug.Print UBound(点)
                绘制需求 = True
                GoTo Su
        End Select
        命令框.SelStart = 0: 命令框.SelLength = Len(命令框.Text)
        Exit Function
Su:
    文本框更新值 "【执行成功！】——" & 命令框.Text & vbCrLf, 反馈框
    Exit Function
Er:
    文本框更新值 "【执行错误！" & Err.Description & "】——" & 命令框.Text & vbCrLf, 反馈框
End Function

Private Function 文本框更新值(值 As String, 反馈框 As RichTextBox)
    反馈框.SelStart = Len(反馈框.Text)
    反馈框.SelText = 值
End Function

Public Function 取得二进制文件(路径 As String)
    Dim fN1 As Integer, fN2 As Integer, 缓存 As String, 缓存2() As String
    fN1 = FreeFile
    Open 路径 & ".ini" For Binary As #fN1
        缓存 = Input(LOF(1), #fN1)
    Close #fN1
    ReDim 点(Val(缓存))
    fN2 = FreeFile
    Open 路径 For Binary As fN2
        Get fN2, , 点
    Close fN2
End Function

Public Function 存为二进制文件(路径 As String, 命令框 As TextBox, 反馈框 As RichTextBox) As String
    Dim fN As Integer, fN2 As Integer
    On Error GoTo Er
        fN = FreeFile
        Open 路径 For Binary As #fN
            Put #fN, , 点
        Close #fN
        fN2 = FreeFile
        Open 路径 & ".ini" For Output As #fN2
            Print #fN2, UBound(点)
        Close #fN2
        Exit Function
Er:
    Close #fN
    Close #fN2
    存为二进制文件 = "错误！"
    文本框更新值 "【执行错误！" & Err.Description & "】——" & 命令框.Text & vbCrLf, 反馈框
End Function
