VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4380
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   4380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "�������"
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   1080
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form2.frx":0000
      Left            =   960
      List            =   "Form2.frx":0016
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ���_������������ As String
Dim ����������������_���_������������ As String
Dim i As Byte
Dim ����� As String


Private Sub ��������_�_866(name As String)
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
            Case "�"
            Mid(����������������_���_������������, i, 1) = "� ���"
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
End Sub
Private Sub Command1_Click() ' ����� ������������
���_������������ = Combo1.Text
Form1.Caption = "������� ������������: " + ���_������������
Form2.Visible = False
��������_�_866 (���_������������)
freeKanal = FreeFile
Open "C:\�������" For Output As #freeKanal
Print #freeKanal, ����������������_���_������������
Close #freeKanal
End Sub

Private Sub Label1_Click()

End Sub
