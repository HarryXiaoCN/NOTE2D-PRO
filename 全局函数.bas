Attribute VB_Name = "ȫ�ֺ���"
Public Function �½���(X, Y, Optional N As String)
    Dim �ڵ������� As String
    On Error GoTo Er
        If N = "" Then
            �ڵ������� = �ڵ�Ĭ��ǰ׺ & UBound(��)
        Else
            �ڵ������� = N
        End If
        If �ڵ����ظ��Լ��(�ڵ�������, UBound(��)) = False Then
            With ��(UBound(��))
                .���� = �ڵ�������
                .��С = 100
                .���� = UBound(��)
                .���� = �ڵ�Ĭ������
                .����.X = X: .����.Y = Y
                .��Ӧ.�� = 1
                .��ɫ = �ڵ�Ĭ����ɫ
                .�༭����ƫ��.X = �ڵ�༭�����ƫ�Ƴ���
                .�༭����ƫ��.Y = �ڵ�༭������ƫ�Ƴ���
            End With
            �������� = True
            ReDim Preserve ��(UBound(��) + 1)
        Else
            �ڵ������� = InputBox("�Զ����ɽڵ����Ѵ��ڣ���ָ���½ڵ����֣�", "�½���", �ڵ�������)
            �½��� X, Y, �ڵ�������
        End If
Er:
End Function

Public Function �ַ���ת������(Ŀ���, ��, �ָ��)
    Dim ����() As String, i As Long, j As Long
    ���� = Split(��, �ָ��)
    For i = 0 To UBound(����)
        If ����(i) <> "" Then
            j = j + 1
        End If
    Next
    ReDim Ŀ���(j - 1)
    j = 0
    For i = 0 To UBound(����)
        If ����(i) <> "" Then
            Ŀ���(j) = Val(����(i))
            j = j + 1
        End If
    Next
End Function

Public Function ����ڵ���(X, Y, Optional Բ�ľ� As Single = 100) As Long
    Dim i As Long, ���볤 As Double
    ����ڵ��� = -1
    On Error GoTo Er
        For i = 0 To UBound(��) - 1
            With ��(i)
                ���볤 = (X - .����.X) ^ 2 + (Y - .����.Y) ^ 2
                If ���볤 < (Բ�ľ� + .��С + 50) ^ 2 Then
                   ����ڵ��� = i: Exit Function
                End If
            End With
        Next
        Exit Function
Er:
    Debug.Print "ȫ�ֺ���[����ڵ���] - ����", Err.Description
End Function

Public Function �ڵ����ظ��Լ��(N, id) As Boolean
    Dim i As Long
    For i = 0 To UBound(��) - 1
        If ��(i).���� = N And id <> i Then
            �ڵ����ظ��Լ�� = True: Exit Function
        End If
    Next
End Function
