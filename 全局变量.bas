Attribute VB_Name = "ȫ�ֱ���"
Public Type ��ֵ
    '������>����ʱ��������Ϣ����ͨ���ýڵ�
    ���� As Double '������ֵ
    ���� As Double
    ���޵��� As String
    ���޵��� As String
    ������� As Double
    ������� As Double
    End Type

Public Type ��Ӧ����
    λ As Long 'ÿ����һ���źŽ��룬���Ϊ0���򴥷������ұ�Ϊ��Ӧ�ڳ��ȣ���Ϊ0���1����������
    �� As Long
    �� As Long
    ��λ As Long
    ȥ As String
    End Type

Public Type ��
    �� As Long '�ܹ������������Ч��
    End Type

Public Type ��ά����
    X As Single
    Y As Single
    End Type
    
Public Type �ڵ�
    ���� As String
    ���� As Long
    ������ As Long
    ���� As String '���ڽڵ��ڵĹ���
    Ȩֵ As Double '�ڵ��Ȩֵ
    ��ֵ As Double '������ʹ��
    ��С As Single
    ���� As Boolean  '����ýڵ��ǳ����򲻻ᱻ�ı�ֵ������Ȼ���Դ�������ֵ
    ��ֵ As ��ֵ
    ��Ӧ As ��Ӧ����
    ���� As ��
    ���� As String ' +-*/ ���ֻ�������
    ��ɫ As Long
    ��Ϣ�� As String '��������Ϣ���������нڵ�����
    ���� As ��ά����
    ȥ As String
    �༭����ƫ�� As ��ά����
    End Type

Public Type ����
    �� As Long
    ȥ As Long
    End Type
Public ��() As �ڵ�, ��() As ����, ��ǰѡ�е� As Long, �༭����װ�ص� As Long
Public ���λ�� As ��ά����, �����ڵ� As String
Public �������� As Boolean, �ڵ�Ĭ����ɫ As Long, �ڵ�Ĭ��ǰ׺ As String, �ڵ�Ĭ������ As String
