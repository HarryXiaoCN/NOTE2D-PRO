Attribute VB_Name = "�滭����"
Public Function ��������(������)
    Dim i As Long
    For i = 0 To UBound(��) - 1
        With ��(i)
            If .ȥ <> "" Then
                ������.ForeColor = .��ɫ
                ��������_�Ӻ��� ������, ��(i), i
            End If
        End With
    Next
    �������� = False
End Function

Public Function ��ά������������(a As ��ά����, b As ��ά����, �з�) As ��ά����
    ��ά������������.X = (b.X - a.X) * �з� + a.X
    ��ά������������.Y = (b.Y - a.Y) * �з� + a.Y
End Function

Private Function ��������_�Ӻ���(������, ȥ�� As �ڵ�, ��, Optional �߿� As Long = 2)
    Dim �е� As ��ά����, ȥ����() As Long, i As Long
    With ȥ��
        For i = 0 To UBound(.ȥ����)
            If .ȥ����(i) < UBound(��) And .ȥ����(i) >= 0 Then
                �е� = ��ά������������(��(��).����, ��(.ȥ����(i)).����, 0.67)
                ������.DrawWidth = �߿�
                ������.Line (��(��).����.X, ��(��).����.Y)-(�е�.X, �е�.Y), ��(��).��ɫ
                ������.DrawWidth = 1
                ������.Line (�е�.X, �е�.Y)-(��(.ȥ����(i)).����.X, ��(.ȥ����(i)).����.Y), ��(.ȥ����(i)).��ɫ
            End If
        Next
    End With
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

Public Function ���������(���� As ��ά����, ��С As Single, ��ɫ As Long, Optional ���Ƽ�� As Long = 400)
    ��.��.FillColor = ��ɫ
    ��.��.Circle (����.X, ����.Y), ��С, ��ɫ
    DoEvents
    Sleep ���Ƽ��
End Function

Public Function �������������(��� As ��ά����, �յ� As ��ά����, Optional ���Ƽ�� As Long = 400, Optional ��ɫ As Long = 14822282, Optional ��� As Long = 3)
    Dim �������Ƽ�� As Double, ���� As ��ά����
    �������Ƽ�� = ���Ƽ�� / 10
    With ��
        .FillColor = ��ɫ
        .��.DrawWidth = ���
        For i = 1 To �������Ƽ��
            ���� = ��ά������������(���, �յ�, i / �������Ƽ��)
            .��.Line (���.X, ���.Y)-(����.X, ����.Y), ��ɫ
            DoEvents
            Sleep 10
        Next
    End With
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
        End With
    Next
    �������� = False
End Function
