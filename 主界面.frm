VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form 体 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "NOTE2D PRO"
   ClientHeight    =   8880
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18015
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8880
   ScaleWidth      =   18015
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox 控制台容器 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5055
      Left            =   11160
      ScaleHeight     =   5025
      ScaleWidth      =   6825
      TabIndex        =   33
      ToolTipText     =   "“~”键隐藏/显示本界面"
      Top             =   0
      Width           =   6855
      Begin VB.TextBox 命令输入框 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   34
         Top             =   120
         Width           =   6615
      End
      Begin RichTextLib.RichTextBox 命令提示框 
         Height          =   3975
         Left            =   120
         TabIndex        =   35
         Top             =   840
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   7011
         _Version        =   393217
         ScrollBars      =   2
         BulletIndent    =   4
         Appearance      =   0
         TextRTF         =   $"主界面.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "楷体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.PictureBox 节点编辑辅助 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5655
      Left            =   0
      ScaleHeight     =   5625
      ScaleWidth      =   3240
      TabIndex        =   17
      Top             =   0
      Width           =   3275
      Begin VB.TextBox 默认节点前缀 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         MaxLength       =   10
         TabIndex        =   31
         Text            =   "d"
         Top             =   360
         Width           =   3015
      End
      Begin VB.PictureBox 默认色 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H000000FF&
         ForeColor       =   &H80000008&
         Height          =   150
         Index           =   0
         Left            =   120
         ScaleHeight     =   120
         ScaleWidth      =   120
         TabIndex        =   30
         Top             =   120
         Width           =   150
      End
      Begin VB.PictureBox 默认色 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H000080FF&
         ForeColor       =   &H80000008&
         Height          =   150
         Index           =   1
         Left            =   360
         ScaleHeight     =   120
         ScaleWidth      =   120
         TabIndex        =   29
         Top             =   120
         Width           =   150
      End
      Begin VB.PictureBox 默认色 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H80000008&
         Height          =   150
         Index           =   2
         Left            =   600
         ScaleHeight     =   120
         ScaleWidth      =   120
         TabIndex        =   28
         Top             =   120
         Width           =   150
      End
      Begin VB.PictureBox 默认色 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H0000FF00&
         ForeColor       =   &H80000008&
         Height          =   150
         Index           =   3
         Left            =   840
         ScaleHeight     =   120
         ScaleWidth      =   120
         TabIndex        =   27
         Top             =   120
         Width           =   150
      End
      Begin VB.PictureBox 默认色 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFF00&
         ForeColor       =   &H80000008&
         Height          =   150
         Index           =   4
         Left            =   1080
         ScaleHeight     =   120
         ScaleWidth      =   120
         TabIndex        =   26
         Top             =   120
         Width           =   150
      End
      Begin VB.PictureBox 默认色 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FF0000&
         ForeColor       =   &H80000008&
         Height          =   150
         Index           =   5
         Left            =   1320
         ScaleHeight     =   120
         ScaleWidth      =   120
         TabIndex        =   25
         Top             =   120
         Width           =   150
      End
      Begin VB.PictureBox 默认色 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FF00FF&
         ForeColor       =   &H80000008&
         Height          =   150
         Index           =   6
         Left            =   1560
         ScaleHeight     =   120
         ScaleWidth      =   120
         TabIndex        =   24
         Top             =   120
         Width           =   150
      End
      Begin VB.PictureBox 默认色 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FF80FF&
         ForeColor       =   &H80000008&
         Height          =   150
         Index           =   7
         Left            =   1800
         ScaleHeight     =   120
         ScaleWidth      =   120
         TabIndex        =   23
         Top             =   120
         Width           =   150
      End
      Begin VB.PictureBox 默认色 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FF8080&
         ForeColor       =   &H80000008&
         Height          =   150
         Index           =   8
         Left            =   2040
         ScaleHeight     =   120
         ScaleWidth      =   120
         TabIndex        =   22
         Top             =   120
         Width           =   150
      End
      Begin VB.PictureBox 默认色 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFC0&
         ForeColor       =   &H80000008&
         Height          =   150
         Index           =   9
         Left            =   2280
         ScaleHeight     =   120
         ScaleWidth      =   120
         TabIndex        =   21
         Top             =   120
         Width           =   150
      End
      Begin VB.PictureBox 默认色 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H0080FF80&
         ForeColor       =   &H80000008&
         Height          =   150
         Index           =   10
         Left            =   2520
         ScaleHeight     =   120
         ScaleWidth      =   120
         TabIndex        =   20
         Top             =   120
         Width           =   150
      End
      Begin VB.PictureBox 默认色 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   150
         Index           =   11
         Left            =   2760
         ScaleHeight     =   120
         ScaleWidth      =   120
         TabIndex        =   19
         Top             =   120
         Width           =   150
      End
      Begin VB.PictureBox 默认色 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   150
         Index           =   12
         Left            =   3000
         ScaleHeight     =   120
         ScaleWidth      =   120
         TabIndex        =   18
         Top             =   120
         Width           =   150
      End
      Begin RichTextLib.RichTextBox 默认节点内容 
         Height          =   4455
         Left            =   120
         TabIndex        =   32
         ToolTipText     =   "按F2弹出内容解析编码帮助"
         Top             =   960
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   7858
         _Version        =   393217
         ScrollBars      =   2
         BulletIndent    =   4
         Appearance      =   0
         TextRTF         =   $"主界面.frx":009D
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "楷体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.PictureBox 面 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   10000
      Left            =   -120
      ScaleHeight     =   9975
      ScaleWidth      =   18465
      TabIndex        =   0
      Top             =   -120
      Width           =   18500
      Begin VB.Timer 编辑界面关闭时钟 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   1200
         Top             =   8520
      End
      Begin VB.Timer 区域检测时钟 
         Interval        =   100
         Left            =   720
         Top             =   8520
      End
      Begin VB.Timer 绘制时钟 
         Interval        =   30
         Left            =   240
         Top             =   8520
      End
      Begin VB.PictureBox 节点编辑界面 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   4215
         Left            =   6240
         ScaleHeight     =   4185
         ScaleWidth      =   3225
         TabIndex        =   3
         Top             =   1440
         Visible         =   0   'False
         Width           =   3255
         Begin VB.PictureBox 颜色 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00404040&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   12
            Left            =   3000
            ScaleHeight     =   120
            ScaleWidth      =   120
            TabIndex        =   16
            Top             =   600
            Width           =   150
         End
         Begin VB.PictureBox 颜色 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H0080C0FF&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   11
            Left            =   2760
            ScaleHeight     =   120
            ScaleWidth      =   120
            TabIndex        =   15
            Top             =   600
            Width           =   150
         End
         Begin VB.PictureBox 颜色 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H0080FF80&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   10
            Left            =   2520
            ScaleHeight     =   120
            ScaleWidth      =   120
            TabIndex        =   14
            Top             =   600
            Width           =   150
         End
         Begin VB.PictureBox 颜色 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   9
            Left            =   2280
            ScaleHeight     =   120
            ScaleWidth      =   120
            TabIndex        =   13
            Top             =   600
            Width           =   150
         End
         Begin VB.PictureBox 颜色 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FF8080&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   8
            Left            =   2040
            ScaleHeight     =   120
            ScaleWidth      =   120
            TabIndex        =   12
            Top             =   600
            Width           =   150
         End
         Begin VB.PictureBox 颜色 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FF80FF&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   7
            Left            =   1800
            ScaleHeight     =   120
            ScaleWidth      =   120
            TabIndex        =   11
            Top             =   600
            Width           =   150
         End
         Begin VB.PictureBox 颜色 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FF00FF&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   6
            Left            =   1560
            ScaleHeight     =   120
            ScaleWidth      =   120
            TabIndex        =   10
            Top             =   600
            Width           =   150
         End
         Begin VB.PictureBox 颜色 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FF0000&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   5
            Left            =   1320
            ScaleHeight     =   120
            ScaleWidth      =   120
            TabIndex        =   9
            Top             =   600
            Width           =   150
         End
         Begin VB.PictureBox 颜色 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFF00&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   4
            Left            =   1080
            ScaleHeight     =   120
            ScaleWidth      =   120
            TabIndex        =   8
            Top             =   600
            Width           =   150
         End
         Begin VB.PictureBox 颜色 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H0000FF00&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   3
            Left            =   840
            ScaleHeight     =   120
            ScaleWidth      =   120
            TabIndex        =   7
            Top             =   600
            Width           =   150
         End
         Begin VB.PictureBox 颜色 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H0000FFFF&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   2
            Left            =   600
            ScaleHeight     =   120
            ScaleWidth      =   120
            TabIndex        =   6
            Top             =   600
            Width           =   150
         End
         Begin VB.PictureBox 颜色 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H000080FF&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   1
            Left            =   360
            ScaleHeight     =   120
            ScaleWidth      =   120
            TabIndex        =   5
            Top             =   600
            Width           =   150
         End
         Begin VB.PictureBox 颜色 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H000000FF&
            ForeColor       =   &H80000008&
            Height          =   150
            Index           =   0
            Left            =   120
            ScaleHeight     =   120
            ScaleWidth      =   120
            TabIndex        =   4
            Top             =   600
            Width           =   150
         End
         Begin RichTextLib.RichTextBox 节点内容 
            Height          =   3135
            Left            =   120
            TabIndex        =   2
            Top             =   840
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   5530
            _Version        =   393217
            ScrollBars      =   2
            BulletIndent    =   4
            Appearance      =   0
            TextRTF         =   $"主界面.frx":0333
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "楷体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.TextBox 节点名 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "微软雅黑"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            MaxLength       =   10
            TabIndex        =   1
            Top             =   120
            Width           =   3015
         End
      End
   End
End
Attribute VB_Name = "体"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private 节点编辑初始位置 As 二维坐标, 节点移动初始位置 As 二维坐标, 面移动初始位置 As 二维坐标, 控制台初始位置 As 二维坐标
Private 节点默认编辑初始位置 As 二维坐标

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27
            节点编辑界面.Visible = False
        Case vbKeyF1
            初始化全部节点
        Case vbKeyF5
            启动流 启动节点
        Case vbKeyF2
            MsgBox 编码规则提示初始化, 32, "节点内容编码帮助"
        Case vbKeyReturn
            If 命令输入框.Text <> "" Then 命令输入 命令输入框, 命令提示框
        Case 192
            If 控制台容器.Visible Then 控制台容器.Visible = False Else 控制台容器.Visible = True
    End Select
End Sub

Private Function 初始化全部节点()
    Dim i As Long
    For i = 0 To UBound(点) - 1
        节点内容初始化解析 点(i).内容, i
    Next
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            KeyAscii = 0
    End Select
End Sub

Private Sub Form_Load()
    ReDim 点(0), 线(0)
    With 面
        .Top = -20000: .Left = -20000
        .Height = 60000: .Width = 60000
    End With
    节点默认前缀 = "d"
    节点默认内容 = 默认节点内容.Text
    颜色(9).BackColor = 初始节点颜色
    默认色(9).BackColor = 初始节点颜色
    默认色_Click 9
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub 编辑界面关闭时钟_Timer()
    节点编辑界面.Visible = False
    编辑界面关闭时钟.Enabled = False
End Sub

Private Sub 绘制时钟_Timer()
    Dim 绘制时钟计时器 As Double
    If 绘制需求 = True Then
        面.Cls
        绘制时钟计时器 = Timer
        绘制连接 面
        绘制节点 面
        绘制源点 面
        绘制时钟计时器 = Format(Timer - 绘制时钟计时器, "0.000") * 1000
'        Debug.Print 绘制时钟计时器
        If 绘制时钟.Interval <= 绘制时钟计时器 Then
            绘制时钟.Interval = 绘制时钟计时器 + 10
        ElseIf 绘制时钟.Interval >= 绘制时钟计时器 + 15 Then
            绘制时钟.Interval = 绘制时钟计时器 + 10
        End If
    End If
End Sub

Public Function 绘制周期()
    面.Cls
    绘制连接 面
    绘制节点 面
    绘制源点 面
End Function

Private Sub 节点编辑辅助_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button > 0 Then
        节点编辑辅助.Top = Y - 节点默认编辑初始位置.Y + 节点编辑辅助.Top
        节点编辑辅助.Left = X - 节点默认编辑初始位置.X + 节点编辑辅助.Left
    End If
End Sub

Private Sub 节点编辑辅助_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    节点默认编辑初始位置.X = X: 节点默认编辑初始位置.Y = Y
End Sub

Private Sub 节点编辑界面_GotFocus()
    编辑界面关闭时钟.Enabled = False
End Sub

Private Sub 节点编辑界面_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    节点编辑初始位置.X = X: 节点编辑初始位置.Y = Y
End Sub

Private Sub 节点编辑界面_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        节点编辑界面.Top = 节点编辑界面.Top + Y - 节点编辑初始位置.Y
        节点编辑界面.Left = 节点编辑界面.Left + X - 节点编辑初始位置.X
        With 点(编辑界面装载点)
            .编辑界面偏移.X = 节点编辑界面.Left - .坐标.X
            .编辑界面偏移.Y = 节点编辑界面.Top - .坐标.Y
        End With
    End If
End Sub

Private Function 节点名过滤器(N)
    If InStr(1, N, """") > 0 Then N = Replace(N, """", "")
    If InStr(1, N, " ") > 0 Then N = Replace(N, " ", "")
End Function

Private Sub 节点名_Change()
    节点名过滤器 节点名.Text
    If 节点名重复性检测(节点名.Text, 编辑界面装载点) = False Then
        点(编辑界面装载点).名字 = 节点名.Text
        绘制需求 = True
        节点名.ToolTipText = ""
    ElseIf 当前选中点 <> -1 Then
        节点名.ToolTipText = "重名警告！此修改不会被同步！"
    End If
End Sub

Private Sub 节点名_GotFocus()
    节点编辑界面_GotFocus
End Sub

Private Sub 节点内容_Change()
    点(编辑界面装载点).内容 = 节点内容.Text
    节点内容初始化解析 节点内容.Text, 编辑界面装载点
End Sub

Private Sub 节点内容_GotFocus()
    节点编辑界面_GotFocus
End Sub

Private Sub 控制台容器_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    控制台初始位置.X = X: 控制台初始位置.Y = Y
End Sub

Private Sub 控制台容器_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button > 0 Then
        控制台容器.Top = Y - 控制台初始位置.Y + 控制台容器.Top
        控制台容器.Left = X - 控制台初始位置.X + 控制台容器.Left
    End If
End Sub

Private Sub 面_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
        Case 1
            If 当前选中点 >= 0 Then
                If Shift = 0 Then
                    区域检测时钟.Enabled = False
                    节点移动初始位置.X = X: 节点移动初始位置.Y = Y
                    节点编辑界面.Visible = False
                ElseIf Shift = 2 Then
                    Dim 缓存 As String, 缓存2 As String
                    缓存 = "," & 启动节点
                    缓存2 = "," & 当前选中点 & ","
                    If InStr(1, 缓存, 缓存2) = 0 Then
                        启动节点 = 启动节点 & 当前选中点 & ","
                        绘制需求 = True
                    End If
                End If
            Else
                面移动初始位置.X = X: 面移动初始位置.Y = Y
                Form_KeyDown 27, 0
            End If
        Case 2
            If Shift = 0 Then
                If 当前选中点 = -1 And X > 500 And Y > 500 Then
                    新建点 X, Y
                    面扩张 X, Y
                End If
            ElseIf Shift = 2 And 当前选中点 >= 0 Then
                If 启动节点 <> "" Then
                    缓存 = "," & 启动节点
                    缓存2 = "," & 当前选中点 & ","
                    启动节点 = Replace(缓存, 缓存2, "")
                    绘制需求 = True
                End If
            End If
    End Select
End Sub

Private Function 面扩张(X, Y)
    If X > 面.Width - 1000 Then
        面.Width = 面.Width + 3000
    End If
    If Y > 面.Height - 1000 Then
        面.Height = 面.Height + 3000
    End If
End Function

Private Sub 面_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    鼠标位置.X = X: 鼠标位置.Y = Y
    If Shift = 0 Then
        Select Case Button
            Case 1
                If 当前选中点 >= 0 Then
                    With 点(当前选中点)
                        Dim 缓存坐标 As 二维坐标
                        缓存坐标.X = .坐标.X + X - 节点移动初始位置.X
                        缓存坐标.Y = .坐标.Y + Y - 节点移动初始位置.Y
                        面扩张 缓存坐标.X, 缓存坐标.Y
                        If 缓存坐标.X > 500 And 缓存坐标.Y > 500 Then
                            .坐标.X = 缓存坐标.X
                            .坐标.Y = 缓存坐标.Y
                            节点移动初始位置 = .坐标
                            绘制需求 = True
                        End If
                    End With
                ElseIf 面移动初始位置.X <> 0 And 面移动初始位置.Y <> 0 Then
                    面.Top = 面.Top + Y - 面移动初始位置.Y
                    面.Left = 面.Left + X - 面移动初始位置.X
                End If
            Case 2
                
        End Select
    End If
End Sub

Private Sub 面_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
        Case 1
            If 当前选中点 >= 0 Then
                节点编辑界面显示 当前选中点
                区域检测时钟.Enabled = True
            End If
        Case 2
            
    End Select
    
End Sub

Private Sub 默认节点内容_Change()
    节点默认内容 = 默认节点内容.Text
End Sub

Private Sub 默认节点前缀_Change()
    节点名过滤器 默认节点前缀.Text
    节点默认前缀 = 默认节点前缀.Text
End Sub

Private Sub 默认色_Click(Index As Integer)
    节点默认颜色 = 默认色(Index).BackColor
End Sub

Private Sub 区域检测时钟_Timer()
    Dim i As Long
    i = 区域节点检测(鼠标位置.X, 鼠标位置.Y)
    If i <> 当前选中点 Then
        当前选中点 = i
        With 节点编辑界面
            If i >= 0 Then
                编辑界面关闭时钟.Enabled = False
                节点编辑界面显示 i
'            Else
'                编辑界面关闭时钟.Enabled = True
            End If
        End With
    End If
End Sub

Private Function 节点编辑界面显示(i)
    With 节点编辑界面
        编辑界面装载点 = i
        .Top = 点(i).坐标.Y + 点(i).编辑界面偏移.Y
        .Left = 点(i).坐标.X + 点(i).编辑界面偏移.X
        节点名.Text = 点(i).名字
        节点内容.Text = 点(i).内容
        .Visible = True
    End With
End Function

Private Sub 颜色_GotFocus(Index As Integer)
    节点编辑界面_GotFocus
End Sub

Private Sub 颜色_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
        Case 1
            点(编辑界面装载点).颜色 = 颜色(Index).BackColor
            绘制需求 = True
        Case 2
            
    End Select
End Sub
