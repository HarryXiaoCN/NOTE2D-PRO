Attribute VB_Name = "����"
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Function ������(�� As String)
    Dim �ڵ��() As Long, i As Long
    If �� <> "" Then
        ��.����ʱ��.Enabled = False
        �ַ���ת������ �ڵ��, ��, ","
        For i = 0 To UBound(�ڵ��)
            �� �ڵ��(i)
            ��.��������
        Next
        �������� = True
        ��.����ʱ��.Enabled = True
    End If
End Function

Public Function ��(��� As Long)
    Dim �ڵ��() As Long
    �ڵ�� = ��������_�Ӻ���_�ڵ���ת������(��(���).ȥ)
    ��_�Ӻ��� ���, �ڵ��
End Function

Public Function ��_�Ӻ���(��� As Long, �ڵ��() As Long, Optional �������� As Boolean)
    Dim i As Long
    For i = 0 To UBound(�ڵ��)
        If �ڵ��(i) >= 0 Then
            If ������(��(���), ��(�ڵ��(i)), ��������) Then
                �� �ڵ��(i)
            End If
        End If
    Next
End Function

Public Function ������_��ֵ�����Ӻ���(Դ�� As �ڵ�, ȥ�� As �ڵ�, ���� As String)
    Dim �ڵ㼯() As Long
    �ڵ㼯 = ��������_�Ӻ���_�ڵ���ת������(����)
    ������������� Դ��.����, ȥ��.����
    ��_�Ӻ��� ȥ��.����, �ڵ㼯, True
End Function

Public Function ������(Դ�� As �ڵ�, ȥ�� As �ڵ�, �������� As Boolean) As Boolean
    Dim Դֵ As Double, ȥֵ As Double
    Դֵ = �ڵ�����������ֵ��ȡ(Դ��)
    Debug.Print Դ��.����; ȥ��.����
    If ȥ��.��Ӧ.λ <= 0 Then
        ȥ��.��Ӧ.λ = ȥ��.��Ӧ.��
        If ȥ��.��ֵ.���޵��� <> "" And Դֵ < ȥ��.��ֵ.���� And (ȥ��.��ֵ.���� <> 0 And ȥ��.��ֵ.���� <> 0) Then
            ������_��ֵ�����Ӻ��� Դ��, ȥ��, ȥ��.��ֵ.���޵���
        End If
        If �������� = True Or (Դֵ >= ȥ��.��ֵ.���� And Դֵ <= ȥ��.��ֵ.����) Or (ȥ��.��ֵ.���� = 0 And ȥ��.��ֵ.���� = 0) Then
            ��.��������
            ������������� Դ��.����, ȥ��.����
            �ڵ����������� �ڵ�����������ֵ��ȡ(Դ��), ȥ��
            ȥֵ = �ڵ�����������ֵ��ȡ(ȥ��)
            If (ȥֵ >= ȥ��.��ֵ.������� And ȥֵ <= ȥ��.��ֵ.�������) Or (ȥ��.��ֵ.������� = 0 And ȥ��.��ֵ.������� = 0) Then
                ������ = True
            End If
        End If
        If ȥ��.��ֵ.���޵��� <> "" And Դֵ > ȥ��.��ֵ.���� And (ȥ��.��ֵ.���� <> 0 Or ȥ��.��ֵ.���� <> 0) Then
            ������_��ֵ�����Ӻ��� Դ��, ȥ��, ȥ��.��ֵ.���޵���
        End If
    Else
        ȥ��.��Ӧ.λ = ȥ��.��Ӧ.λ - 1
        ������_��ֵ�����Ӻ��� Դ��, ȥ��, ȥ��.��Ӧ.ȥ
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
                ���� = b + a
            Case "-"
                ���� = b - a
            Case "--"
                ���� = a - b
            Case "*"
                ���� = b * a
            Case "/"
                ���� = b / a
            Case "//"
                ���� = a / b
            Case "="
                ���� = a
            Case "=="
                ���� = b
        End Select
Er:
End Function
