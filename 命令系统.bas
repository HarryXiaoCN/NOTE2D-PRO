Attribute VB_Name = "����ϵͳ"
Public Function ��������(����� As TextBox, ������ As RichTextBox)
    Dim ����() As String
    On Error GoTo Er
        ���� = Split(�����.Text, " ")
        Select Case ����(0)
            Case "��"
                If ��Ϊ�������ļ�(����(1), �����, ������) = "" Then
                    GoTo Su
                End If
            Case "ȡ"
                ȡ�ö������ļ� ����(1)
                Debug.Print UBound(��)
                �������� = True
                GoTo Su
        End Select
        �����.SelStart = 0: �����.SelLength = Len(�����.Text)
        Exit Function
Su:
    �ı������ֵ "��ִ�гɹ���������" & �����.Text & vbCrLf, ������
    Exit Function
Er:
    �ı������ֵ "��ִ�д���" & Err.Description & "������" & �����.Text & vbCrLf, ������
End Function

Private Function �ı������ֵ(ֵ As String, ������ As RichTextBox)
    ������.SelStart = Len(������.Text)
    ������.SelText = ֵ
End Function

Public Function ȡ�ö������ļ�(·�� As String)
    Dim fN1 As Integer, fN2 As Integer, ���� As String, ����2() As String
    fN1 = FreeFile
    Open ·�� & ".ini" For Binary As #fN1
        ���� = Input(LOF(1), #fN1)
    Close #fN1
    ReDim ��(Val(����))
    fN2 = FreeFile
    Open ·�� For Binary As fN2
        Get fN2, , ��
    Close fN2
End Function

Public Function ��Ϊ�������ļ�(·�� As String, ����� As TextBox, ������ As RichTextBox) As String
    Dim fN As Integer, fN2 As Integer
    On Error GoTo Er
        fN = FreeFile
        Open ·�� For Binary As #fN
            Put #fN, , ��
        Close #fN
        fN2 = FreeFile
        Open ·�� & ".ini" For Output As #fN2
            Print #fN2, UBound(��)
        Close #fN2
        Exit Function
Er:
    Close #fN
    Close #fN2
    ��Ϊ�������ļ� = "����"
    �ı������ֵ "��ִ�д���" & Err.Description & "������" & �����.Text & vbCrLf, ������
End Function
