Attribute VB_Name = "���ݽ���"
Public Function �ڵ����ݳ�ʼ������(����, �ڵ��)
    Dim ��() As String, i As Long
    �ڵ����� �ڵ��
    �� = Split(����, vbCrLf)
    For i = 0 To UBound(��)
        �ڵ����ݳ�ʼ������_�Ӻ��� ��(i), �ڵ��
    Next
    �������� = True
End Function
Private Function �ڵ�����(i)
    Dim �½ڵ� As �ڵ�
    �½ڵ�.���� = ��(i).����
    �½ڵ�.���� = ��(i).����
    �½ڵ�.���� = ��(i).����
    �½ڵ�.��ɫ = ��(i).��ɫ
    �½ڵ�.��С = ��(i).��С
    ��(i) = �½ڵ�
End Function
Private Function �ڵ����ݳ�ʼ������_�Ӻ���(��, �ڵ��)
    Dim ��ͷ As String
    On Error GoTo Er
    If InStr(1, ��, " ") > 0 Then
        ��ͷ = UCase(Split(��, " ")(0))
    Else
        ��ͷ = UCase(��)
    End If
    Select Case ��ͷ
        Case "ȥ", "Q", "QU"
            ��(�ڵ��).ȥ = Split(��, " ")(1)
        Case "ֵ", "Z"
            ��(�ڵ��).Ȩֵ = Val(Split(��, " ")(1))
        Case "��", "S"
            ��(�ڵ��).���� = Split(��, " ")(1)
        Case "����", "SX"
            ��(�ڵ��).��ֵ.���� = Val(Split(��, " ")(1))
        Case "����", "XX"
            ��(�ڵ��).��ֵ.���� = Val(Split(��, " ")(1))
        Case "��", "C"
            ��(�ڵ��).���� = True
    End Select
Er:
End Function
