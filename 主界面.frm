VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form �� 
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
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox ����̨���� 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      ToolTipText     =   "��~��������/��ʾ������"
      Top             =   0
      Width           =   6855
      Begin VB.TextBox ��������� 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "΢���ź�"
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
      Begin RichTextLib.RichTextBox ������ʾ�� 
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
         TextRTF         =   $"������.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.PictureBox �ڵ�༭���� 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Begin VB.TextBox Ĭ�Ͻڵ�ǰ׺ 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "΢���ź�"
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
      Begin VB.PictureBox Ĭ��ɫ 
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
      Begin VB.PictureBox Ĭ��ɫ 
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
      Begin VB.PictureBox Ĭ��ɫ 
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
      Begin VB.PictureBox Ĭ��ɫ 
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
      Begin VB.PictureBox Ĭ��ɫ 
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
      Begin VB.PictureBox Ĭ��ɫ 
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
      Begin VB.PictureBox Ĭ��ɫ 
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
      Begin VB.PictureBox Ĭ��ɫ 
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
      Begin VB.PictureBox Ĭ��ɫ 
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
      Begin VB.PictureBox Ĭ��ɫ 
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
      Begin VB.PictureBox Ĭ��ɫ 
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
      Begin VB.PictureBox Ĭ��ɫ 
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
      Begin VB.PictureBox Ĭ��ɫ 
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
      Begin RichTextLib.RichTextBox Ĭ�Ͻڵ����� 
         Height          =   4455
         Left            =   120
         TabIndex        =   32
         ToolTipText     =   "��F2�������ݽ����������"
         Top             =   960
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   7858
         _Version        =   393217
         ScrollBars      =   2
         BulletIndent    =   4
         Appearance      =   0
         TextRTF         =   $"������.frx":009D
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.PictureBox �� 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Begin VB.Timer �༭����ر�ʱ�� 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   1200
         Top             =   8520
      End
      Begin VB.Timer ������ʱ�� 
         Interval        =   100
         Left            =   720
         Top             =   8520
      End
      Begin VB.Timer ����ʱ�� 
         Interval        =   30
         Left            =   240
         Top             =   8520
      End
      Begin VB.PictureBox �ڵ�༭���� 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "΢���ź�"
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
         Begin VB.PictureBox ��ɫ 
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
         Begin VB.PictureBox ��ɫ 
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
         Begin VB.PictureBox ��ɫ 
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
         Begin VB.PictureBox ��ɫ 
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
         Begin VB.PictureBox ��ɫ 
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
         Begin VB.PictureBox ��ɫ 
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
         Begin VB.PictureBox ��ɫ 
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
         Begin VB.PictureBox ��ɫ 
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
         Begin VB.PictureBox ��ɫ 
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
         Begin VB.PictureBox ��ɫ 
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
         Begin VB.PictureBox ��ɫ 
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
         Begin VB.PictureBox ��ɫ 
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
         Begin VB.PictureBox ��ɫ 
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
         Begin RichTextLib.RichTextBox �ڵ����� 
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
            TextRTF         =   $"������.frx":0333
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.TextBox �ڵ��� 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "΢���ź�"
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
Attribute VB_Name = "��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private �ڵ�༭��ʼλ�� As ��ά����, �ڵ��ƶ���ʼλ�� As ��ά����, ���ƶ���ʼλ�� As ��ά����, ����̨��ʼλ�� As ��ά����
Private �ڵ�Ĭ�ϱ༭��ʼλ�� As ��ά����

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27
            �ڵ�༭����.Visible = False
        Case vbKeyF1
            ��ʼ��ȫ���ڵ�
        Case vbKeyF5
            ������ �����ڵ�
        Case vbKeyF2
            MsgBox ���������ʾ��ʼ��, 32, "�ڵ����ݱ������"
        Case vbKeyReturn
            If ���������.Text <> "" Then �������� ���������, ������ʾ��
        Case 192
            If ����̨����.Visible Then ����̨����.Visible = False Else ����̨����.Visible = True
    End Select
End Sub

Private Function ��ʼ��ȫ���ڵ�()
    Dim i As Long
    For i = 0 To UBound(��) - 1
        �ڵ����ݳ�ʼ������ ��(i).����, i
    Next
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn
            KeyAscii = 0
    End Select
End Sub

Private Sub Form_Load()
    ReDim ��(0), ��(0)
    With ��
        .Top = -20000: .Left = -20000
        .Height = 60000: .Width = 60000
    End With
    �ڵ�Ĭ��ǰ׺ = "d"
    �ڵ�Ĭ������ = Ĭ�Ͻڵ�����.Text
    ��ɫ(9).BackColor = ��ʼ�ڵ���ɫ
    Ĭ��ɫ(9).BackColor = ��ʼ�ڵ���ɫ
    Ĭ��ɫ_Click 9
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub �༭����ر�ʱ��_Timer()
    �ڵ�༭����.Visible = False
    �༭����ر�ʱ��.Enabled = False
End Sub

Private Sub ����ʱ��_Timer()
    Dim ����ʱ�Ӽ�ʱ�� As Double
    If �������� = True Then
        ��.Cls
        ����ʱ�Ӽ�ʱ�� = Timer
        �������� ��
        ���ƽڵ� ��
        ����Դ�� ��
        ����ʱ�Ӽ�ʱ�� = Format(Timer - ����ʱ�Ӽ�ʱ��, "0.000") * 1000
'        Debug.Print ����ʱ�Ӽ�ʱ��
        If ����ʱ��.Interval <= ����ʱ�Ӽ�ʱ�� Then
            ����ʱ��.Interval = ����ʱ�Ӽ�ʱ�� + 10
        ElseIf ����ʱ��.Interval >= ����ʱ�Ӽ�ʱ�� + 15 Then
            ����ʱ��.Interval = ����ʱ�Ӽ�ʱ�� + 10
        End If
    End If
End Sub

Public Function ��������()
    ��.Cls
    �������� ��
    ���ƽڵ� ��
    ����Դ�� ��
End Function

Private Sub �ڵ�༭����_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button > 0 Then
        �ڵ�༭����.Top = Y - �ڵ�Ĭ�ϱ༭��ʼλ��.Y + �ڵ�༭����.Top
        �ڵ�༭����.Left = X - �ڵ�Ĭ�ϱ༭��ʼλ��.X + �ڵ�༭����.Left
    End If
End Sub

Private Sub �ڵ�༭����_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    �ڵ�Ĭ�ϱ༭��ʼλ��.X = X: �ڵ�Ĭ�ϱ༭��ʼλ��.Y = Y
End Sub

Private Sub �ڵ�༭����_GotFocus()
    �༭����ر�ʱ��.Enabled = False
End Sub

Private Sub �ڵ�༭����_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    �ڵ�༭��ʼλ��.X = X: �ڵ�༭��ʼλ��.Y = Y
End Sub

Private Sub �ڵ�༭����_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        �ڵ�༭����.Top = �ڵ�༭����.Top + Y - �ڵ�༭��ʼλ��.Y
        �ڵ�༭����.Left = �ڵ�༭����.Left + X - �ڵ�༭��ʼλ��.X
        With ��(�༭����װ�ص�)
            .�༭����ƫ��.X = �ڵ�༭����.Left - .����.X
            .�༭����ƫ��.Y = �ڵ�༭����.Top - .����.Y
        End With
    End If
End Sub

Private Function �ڵ���������(N)
    If InStr(1, N, """") > 0 Then N = Replace(N, """", "")
    If InStr(1, N, " ") > 0 Then N = Replace(N, " ", "")
End Function

Private Sub �ڵ���_Change()
    �ڵ��������� �ڵ���.Text
    If �ڵ����ظ��Լ��(�ڵ���.Text, �༭����װ�ص�) = False Then
        ��(�༭����װ�ص�).���� = �ڵ���.Text
        �������� = True
        �ڵ���.ToolTipText = ""
    ElseIf ��ǰѡ�е� <> -1 Then
        �ڵ���.ToolTipText = "�������棡���޸Ĳ��ᱻͬ����"
    End If
End Sub

Private Sub �ڵ���_GotFocus()
    �ڵ�༭����_GotFocus
End Sub

Private Sub �ڵ�����_Change()
    ��(�༭����װ�ص�).���� = �ڵ�����.Text
    �ڵ����ݳ�ʼ������ �ڵ�����.Text, �༭����װ�ص�
End Sub

Private Sub �ڵ�����_GotFocus()
    �ڵ�༭����_GotFocus
End Sub

Private Sub ����̨����_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ����̨��ʼλ��.X = X: ����̨��ʼλ��.Y = Y
End Sub

Private Sub ����̨����_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button > 0 Then
        ����̨����.Top = Y - ����̨��ʼλ��.Y + ����̨����.Top
        ����̨����.Left = X - ����̨��ʼλ��.X + ����̨����.Left
    End If
End Sub

Private Sub ��_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
        Case 1
            If ��ǰѡ�е� >= 0 Then
                If Shift = 0 Then
                    ������ʱ��.Enabled = False
                    �ڵ��ƶ���ʼλ��.X = X: �ڵ��ƶ���ʼλ��.Y = Y
                    �ڵ�༭����.Visible = False
                ElseIf Shift = 2 Then
                    Dim ���� As String, ����2 As String
                    ���� = "," & �����ڵ�
                    ����2 = "," & ��ǰѡ�е� & ","
                    If InStr(1, ����, ����2) = 0 Then
                        �����ڵ� = �����ڵ� & ��ǰѡ�е� & ","
                        �������� = True
                    End If
                End If
            Else
                ���ƶ���ʼλ��.X = X: ���ƶ���ʼλ��.Y = Y
                Form_KeyDown 27, 0
            End If
        Case 2
            If Shift = 0 Then
                If ��ǰѡ�е� = -1 And X > 500 And Y > 500 Then
                    �½��� X, Y
                    ������ X, Y
                End If
            ElseIf Shift = 2 And ��ǰѡ�е� >= 0 Then
                If �����ڵ� <> "" Then
                    ���� = "," & �����ڵ�
                    ����2 = "," & ��ǰѡ�е� & ","
                    �����ڵ� = Replace(����, ����2, "")
                    �������� = True
                End If
            End If
    End Select
End Sub

Private Function ������(X, Y)
    If X > ��.Width - 1000 Then
        ��.Width = ��.Width + 3000
    End If
    If Y > ��.Height - 1000 Then
        ��.Height = ��.Height + 3000
    End If
End Function

Private Sub ��_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ���λ��.X = X: ���λ��.Y = Y
    If Shift = 0 Then
        Select Case Button
            Case 1
                If ��ǰѡ�е� >= 0 Then
                    With ��(��ǰѡ�е�)
                        Dim �������� As ��ά����
                        ��������.X = .����.X + X - �ڵ��ƶ���ʼλ��.X
                        ��������.Y = .����.Y + Y - �ڵ��ƶ���ʼλ��.Y
                        ������ ��������.X, ��������.Y
                        If ��������.X > 500 And ��������.Y > 500 Then
                            .����.X = ��������.X
                            .����.Y = ��������.Y
                            �ڵ��ƶ���ʼλ�� = .����
                            �������� = True
                        End If
                    End With
                ElseIf ���ƶ���ʼλ��.X <> 0 And ���ƶ���ʼλ��.Y <> 0 Then
                    ��.Top = ��.Top + Y - ���ƶ���ʼλ��.Y
                    ��.Left = ��.Left + X - ���ƶ���ʼλ��.X
                End If
            Case 2
                
        End Select
    End If
End Sub

Private Sub ��_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
        Case 1
            If ��ǰѡ�е� >= 0 Then
                �ڵ�༭������ʾ ��ǰѡ�е�
                ������ʱ��.Enabled = True
            End If
        Case 2
            
    End Select
    
End Sub

Private Sub Ĭ�Ͻڵ�����_Change()
    �ڵ�Ĭ������ = Ĭ�Ͻڵ�����.Text
End Sub

Private Sub Ĭ�Ͻڵ�ǰ׺_Change()
    �ڵ��������� Ĭ�Ͻڵ�ǰ׺.Text
    �ڵ�Ĭ��ǰ׺ = Ĭ�Ͻڵ�ǰ׺.Text
End Sub

Private Sub Ĭ��ɫ_Click(Index As Integer)
    �ڵ�Ĭ����ɫ = Ĭ��ɫ(Index).BackColor
End Sub

Private Sub ������ʱ��_Timer()
    Dim i As Long
    i = ����ڵ���(���λ��.X, ���λ��.Y)
    If i <> ��ǰѡ�е� Then
        ��ǰѡ�е� = i
        With �ڵ�༭����
            If i >= 0 Then
                �༭����ر�ʱ��.Enabled = False
                �ڵ�༭������ʾ i
'            Else
'                �༭����ر�ʱ��.Enabled = True
            End If
        End With
    End If
End Sub

Private Function �ڵ�༭������ʾ(i)
    With �ڵ�༭����
        �༭����װ�ص� = i
        .Top = ��(i).����.Y + ��(i).�༭����ƫ��.Y
        .Left = ��(i).����.X + ��(i).�༭����ƫ��.X
        �ڵ���.Text = ��(i).����
        �ڵ�����.Text = ��(i).����
        .Visible = True
    End With
End Function

Private Sub ��ɫ_GotFocus(Index As Integer)
    �ڵ�༭����_GotFocus
End Sub

Private Sub ��ɫ_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
        Case 1
            ��(�༭����װ�ص�).��ɫ = ��ɫ(Index).BackColor
            �������� = True
        Case 2
            
    End Select
End Sub
