Attribute VB_Name = "�������"
Option Explicit
Option Base 1
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Const KEYEVENTF_KEYUP = &H2 '������� ���������� �������
Const VK_LCONTROL = &HA2 ' ����� Ctrl
Const VK_RCONTROL = &HA3 ' ������ Ctrl
Const VK_ESCAPE = &H1B  '������� Escape
Const VK_LWIN = &H5B '����� �������, ����������� ������� ������ ����
Const VK_RWIN = &H5B '����� �������, ����������� ������� ������ ����
Const VK_LMENU = &HA4 ' ����� Alt
Const VK_RMENU = &HA5 ' ������ Alt
Const VK_SHIFT = &H10 ' Shift
 
Sub Command1_Click()
    Call keybd_event(VK_SHIFT, 0, 0, 0) 'H�������
    Call keybd_event(VK_SHIFT, 0, KEYEVENTF_KEYUP, 0) '���������
End Sub

Function ��������_�_1251(name As String) As String
Dim ����������������_���_������������ As String
Dim ����� As String
Dim i As Integer

����������������_���_������������ = name
    For i = 1 To Len(����������������_���_������������)
        ����� = (Mid(name, i, 1))
        Select Case (�����)
            Case "�"
            Mid(����������������_���_������������, i, 1) = "�"
            Case "�"
            Mid(����������������_���_������������, i, 1) = "�"
            Case "�"
            Mid(����������������_���_������������, i, 1) = "�"
            Case "�"
            Mid(����������������_���_������������, i, 1) = "�"
            Case "�"
            Mid(����������������_���_������������, i, 1) = "�"
            Case "�"
            Mid(����������������_���_������������, i, 1) = "�"
            Case "�"
            Mid(����������������_���_������������, i, 1) = "�"
            Case "�"
            Mid(����������������_���_������������, i, 1) = "�"
            Case "�"
            Mid(����������������_���_������������, i, 1) = "�"
            Case "�"
            Mid(����������������_���_������������, i, 1) = "�"
            Case "�"
            Mid(����������������_���_������������, i, 1) = "�"
            Case "�"
            Mid(����������������_���_������������, i, 1) = "�"
            Case "�"
            Mid(����������������_���_������������, i, 1) = "�"
            Case "�"
            Mid(����������������_���_������������, i, 1) = "�"
            Case "�"
            Mid(����������������_���_������������, i, 1) = "�"
            Case "�"
            Mid(����������������_���_������������, i, 1) = "�"
            Case "�"
            Mid(����������������_���_������������, i, 1) = "�"
            Case "�"
            Mid(����������������_���_������������, i, 1) = "�"
            Case "�"
            Mid(����������������_���_������������, i, 1) = "�"
            Case "�"
            Mid(����������������_���_������������, i, 1) = "�"
            Case "�"
            Mid(����������������_���_������������, i, 1) = "�"
            Case "�"
            Mid(����������������_���_������������, i, 1) = "�"
            Case "�"
            Mid(����������������_���_������������, i, 1) = "�"
            Case "�"
            Mid(����������������_���_������������, i, 1) = "�"
            Case "�"
            Mid(����������������_���_������������, i, 1) = "�"
            Case Chr(152)
            Mid(����������������_���_������������, i, 1) = "ؘ��"
            Case "�"
            Mid(����������������_���_������������, i, 1) = "�"
            Case "�"
            Mid(����������������_���_������������, i, 1) = "�"
            Case "�"
            Mid(����������������_���_������������, i, 1) = "�"
            Case "�"
            Mid(����������������_���_������������, i, 1) = "�"
            Case "�"
            Mid(����������������_���_������������, i, 1) = "�"
            Case "�"
            Mid(����������������_���_������������, i, 1) = "�"
            Case "�"
            Mid(����������������_���_������������, i, 1) = "�"
            End Select
         Next
         ��������_�_1251 = ����������������_���_������������
End Function

Function ���������_������_������(������_������� As Byte, ������_��������_������() As String) As Integer ' ��������� ����� � ��������� ��������
    Dim i As Integer ' ������� ���������� ������
    Dim j As Integer ' ������� �����
    Dim s As String
    freeKanal = FreeFile
    Open "C:\TS_NET\log.txt" For Output As #freeKanal
    i = 0
    For j = 1 To ������_�������
        If ������_��������_������(j) = "" Then
            Exit For
        End If
    Print #freeKanal, ������_��������_������(j)
    i = i + 1
  
    Next
    Close #freeKanal
   
   ���������_������_������ = i
   
End Function

Function ���������_������_������(������_������� As Byte, ������_��������_������() As String) As Integer ' ��������� �� ��������� ������ ������
    Dim i As Integer ' ������� ���������� ������
    Dim j As Integer ' ������� �����
    
    freeKanal = FreeFile
    Open "C:\TS_NET\log.txt" For Input As #freeKanal
    j = 1
    i = 0
    Do While Not EOF(freeKanal)
        Line Input #freeKanal, ������_��������_������(j)
        j = j + 1
        i = i + 1
        �������_������ = �������_������ + 1
        If j = ������_������� - 1 Then
            ������_������� = ������_������� + ������_�������
            ReDim Preserve ������_��������_������(������_�������)
        End If
    Loop
    Close #freeKanal
    
    freeKanal = FreeFile
    Open "C:\TS_NET\log.txt" For Output As #freeKanal
    Close #freeKanal
        
    ���������_������_������ = i
End Function
