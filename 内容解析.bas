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
    �½ڵ�.���� = ��(i).����
    �½ڵ�.�༭����ƫ�� = ��(i).�༭����ƫ��
    ��(i) = �½ڵ�
End Function

Private Function �ڵ����ݳ�ʼ������_�Ӻ���_�����ֵ(��, �ڵ��)
    Dim ����() As String, ������() As String
    ���� = Split(��, " ")
    Select Case UCase(����(1))
        Case "R", "�����", "SJS", "SJ", "S", "RND"
            ������ = Split(����(2), ",")
            Randomize Val(������(0))
            ��(�ڵ��).Ȩֵ = Rnd * Val(������(1)) + Val(������(2))
        Case Else
            ��(�ڵ��).Ȩֵ = Val(����(1))
    End Select
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
            ��(�ڵ��).ȥ���� = ��������_�Ӻ���_�ڵ���ת������(��(�ڵ��).ȥ)
        Case "ֵ", "Z"
            �ڵ����ݳ�ʼ������_�Ӻ���_�����ֵ ��, �ڵ��
    End Select
Er:
End Function
