Attribute VB_Name = "�滭����"
Public Function ��������(������)
    Dim i As Long
    For i = 0 To UBound(��) - 1
        With ��(i)
            If .ȥ <> "" Then
                ��������_�Ӻ��� ������, .ȥ, i
            End If
        End With
    Next
    �������� = False
End Function

Private Function ��������_�Ӻ���(������, ȥ, ��)
    Dim �е� As ��ά����, ȥ����() As Long, i As Long
    ȥ���� = ��������_�Ӻ���_�ڵ���ת������(ȥ)
    For i = 0 To UBound(ȥ����)
        If ȥ����(i) < UBound(��) And ȥ����(i) >= 0 Then
            �е�.X = (��(ȥ����(i)).����.X - ��(��).����.X) / 3 * 2 + ��(��).����.X
            �е�.Y = (��(ȥ����(i)).����.Y - ��(��).����.Y) / 3 * 2 + ��(��).����.Y
            ������.DrawWidth = 2
            ������.Line (��(��).����.X, ��(��).����.Y)-(�е�.X, �е�.Y), ��(��).��ɫ
            ������.DrawWidth = 1
            ������.Line (�е�.X, �е�.Y)-(��(ȥ����(i)).����.X, ��(ȥ����(i)).����.Y), ��(ȥ����(i)).��ɫ
        End If
    Next
End Function

Public Function ��������_�Ӻ���_�ڵ���ת������(ByVal ȥ��) As Variant
    Dim ����() As String, i As Long, j As Long, ��() As Long
    ȥ�� = ȥ�� & ","
    ���� = Split(ȥ��, ",")
    ReDim ��(UBound(����) - 1)
    For i = 0 To UBound(��)
        ��(i) = -1
        For j = 0 To UBound(��) - 1
            If ��(j).���� = ����(i) Then
                ��(i) = j: Exit For
            End If
        Next
    Next
    ��������_�Ӻ���_�ڵ���ת������ = ��
End Function

Public Function ����Դ��(������)
    Dim i As Long, Դ�㼯() As String
    Դ�㼯 = Split(�����ڵ�, ",")
    For i = 0 To UBound(Դ�㼯) - 1
        With ��(Val(Դ�㼯(i))).����
            ������.FillColor = Դ��ɫ
            ������.Circle (.X, .Y), 30, Դ��ɫ
        End With
    Next
    �������� = False
End Function

Public Function ���ƽڵ�(������)
    Dim i As Long
    For i = 0 To UBound(��) - 1
        With ��(i)
            ������.FillColor = .��ɫ
            ������.ForeColor = .��ɫ
            ������.Circle (.����.X, .����.Y), .��С, .��ɫ
            ������.CurrentX = .����.X + �ڵ������ƺ�ƫ�Ƴ���
            ������.CurrentY = .����.Y + �ڵ���������ƫ�Ƴ���
            ������.Print .���� & "=" & .Ȩֵ
            ������.CurrentX = .����.X + �ڵ��������ƫ�Ƴ���
            ������.CurrentY = .����.Y + �ڵ��������ƫ�Ƴ���
            ������.Print .����
            ������.CurrentX = .����.X + �ڵ���ź�ƫ�Ƴ���
            ������.CurrentY = .����.Y + �ڵ������ƫ�Ƴ���
            ������.Print i
            If .���� Then
                ������.CurrentX = .����.X + �ڵ���ֵ��ƫ�Ƴ���
                ������.CurrentY = .����.Y + �ڵ���ֵ��ƫ�Ƴ���
                ������.Print .��ֵ
            End If
        End With
    Next
    �������� = False
End Function
