Attribute VB_Name = "����"
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Function ������(�� As String)
    Dim �ڵ��() As Long, i As Long
    ��.����ʱ��.Enabled = False
    �ַ���ת������ �ڵ��, ��, ","
    For i = 0 To UBound(�ڵ��)
        �� �ڵ��(i)
    Next
    �������� = True
    ��.����ʱ��.Enabled = True
End Function

Public Function ��(��� As Long)
    Dim �ڵ��() As Long, i As Long
    �ڵ�� = ��������_�Ӻ���_�ڵ���ת������(��(���).ȥ)
    For i = 0 To UBound(�ڵ��)
        If �ڵ��(i) >= 0 Then
            If ������(��(���), ��(�ڵ��(i))) Then
                ��.��������
                ��.��.FillColor = ��(���).��ɫ
                ��.��.Circle (��(�ڵ��(i)).����.X, ��(�ڵ��(i)).����.Y), ��(�ڵ��(i)).��С / 2, ��(���).��ɫ
                DoEvents
                Sleep 400
                �� �ڵ��(i)
            End If
        End If
    Next
End Function

Public Function ������(Դ�� As �ڵ�, ȥ�� As �ڵ�) As Boolean
    Dim Դֵ As Double
    Դֵ = �ڵ�����������ֵ��ȡ(Դ��)
    If (Դֵ >= ȥ��.��ֵ.���� And Դֵ <= ȥ��.��ֵ.����) Or (ȥ��.��ֵ.���� = 0 And ȥ��.��ֵ.���� = 0) Then
        �ڵ����������� �ڵ�����������ֵ��ȡ(Դ��), ȥ��
        ������ = True
    End If
End Function

Public Function �ڵ�����������(Դֵ, ȥ�� As �ڵ�)
    If ȥ��.���� Then
        ȥ��.��ֵ = ����(Դֵ, ȥ��.Ȩֵ, ȥ��.����)
    Else
        ȥ��.Ȩֵ = ����(Դֵ, ȥ��.Ȩֵ, ȥ��.����)
    End If
End Function

Public Function �ڵ�����������ֵ��ȡ(Դ�� As �ڵ�) As Double
    If Դ��.���� Then
        �ڵ�����������ֵ��ȡ = Դ��.��ֵ
    Else
        �ڵ�����������ֵ��ȡ = Դ��.Ȩֵ
    End If
End Function

Public Function ����(a, b, f) As Double
    On Error GoTo Er
        Select Case f
            Case "+"
                ���� = a + b
            Case "-"
                ���� = a - b
            Case "--"
                ���� = b - a
            Case "*"
                ���� = a * b
            Case "/"
                ���� = a / b
            Case "//"
                ���� = b / a
            Case "="
                ���� = a
        End Select
Er:
End Function
