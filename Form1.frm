VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "������� ������������: "
   ClientHeight    =   2880
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   7275
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   7275
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer_screensaver 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   6360
      Top             =   2040
   End
   Begin VB.Timer Timer3 
      Left            =   120
      Top             =   1920
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   0
      Top             =   5400
   End
   Begin VB.CommandButton Command2 
      Caption         =   "��������"
      Height          =   615
      Left            =   2760
      TabIndex        =   1
      Top             =   2040
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   9600
      Top             =   5400
   End
   Begin VB.Label Label2 
      Alignment       =   2  '���������
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   6015
   End
   Begin VB.Menu ��������� 
      Caption         =   "���������"
      Begin VB.Menu �����_������������ 
         Caption         =   "����� ������������"
      End
      Begin VB.Menu ����_��������_���� 
         Caption         =   "�������� ����"
      End
      Begin VB.Menu ����_��������_������� 
         Caption         =   "�������� �������"
      End
      Begin VB.Menu ����_��������_�������_�������� 
         Caption         =   "�������� ������� ��������"
      End
      Begin VB.Menu ����_�����_��������� 
         Caption         =   "����� ���������"
      End
   End
   Begin VB.Menu ����_������_����� 
      Caption         =   "������ ������"
      Begin VB.Menu ����_��������_����� 
         Caption         =   "�������� ������"
      End
      Begin VB.Menu ����_��������_������ 
         Caption         =   "�������� ������"
      End
   End
   Begin VB.Menu ����_������������ 
      Caption         =   "������������"
      Begin VB.Menu ����_�������_�����_������� 
         Caption         =   "������� ����� �������"
         Begin VB.Menu ����_�������_������� 
            Caption         =   "������� �������"
         End
         Begin VB.Menu ����_�������_�������� 
            Caption         =   "������� ��������"
         End
      End
      Begin VB.Menu ����_��������� 
         Caption         =   "���������"
         Begin VB.Menu ����_���������_���������� 
            Caption         =   "��������� ����������"
         End
         Begin VB.Menu ����_�������_��������� 
            Caption         =   "������� ���������"
         End
         Begin VB.Menu ����_��������_���������� 
            Caption         =   "������� ��������"
         End
      End
      Begin VB.Menu ����_���������� 
         Caption         =   "SCREENSAVER"
         Begin VB.Menu ����_���������_��������������� 
            Caption         =   "�������������� ������"
         End
         Begin VB.Menu ����_����_������ 
            Caption         =   "��������� ������"
         End
         Begin VB.Menu ����_������_������� 
            Caption         =   "������ �������"
         End
      End
      Begin VB.Menu ����_work_with_katalog 
         Caption         =   "������ � ����������"
         Begin VB.Menu ����_�����_����� 
            Caption         =   "����� ����� ��� ��������"
         End
         Begin VB.Menu ����_copy_in_input 
            Caption         =   "����������� � _INPUT"
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long

Dim ������_��������_������() As String ' ������������ ������ ����� ��� ���� �������� ������
Dim ������_������� As Byte
Dim FSO As Object ' ������ FileSystemObject
Dim FILE ' ��������� ���� �� �������� �����
Dim PATH ' ���� � ����� �������� _INPUT
Dim FOLDER ' ������ ��� ������ ���������
Dim ���_����� As String ' ��� ���������� ����� �� �������� �����
Dim ���_�����_������� As String
Dim �����_���_������� As Integer
Dim ������_����_������������ As Boolean
Dim i As Byte ' ������� ��� �����
Dim ��������_msgbox As Integer
Dim ������_��������_������ As String
Dim �������_���� As String
Dim ������ As Integer ' ������� ��� ������� ����������������

Public Sub Form_Load()

Dim ������� As String
If App.PrevInstance Then End ' ��������� ������ ������ � ����������� ����� ����������
������_������� = 15
�������_������ = 1
ReDim ������_��������_������(������_�������)
Set FSO = CreateObject("Scripting.FileSystemObject")

If FSO.FileExists("C:\TS_NET\log.txt") = True Then
    ���������_������_������ ������_�������, ������_��������_������
End If

If FSO.FolderExists("C:\_INPUT") = False Then ' �������� ��� �������� �� ������� �������� ��������
        FSO.CreateFolder ("C:\_INPUT")
End If

Set ��������_����� = FSO.GetFolder("C:\_INPUT")
Command2.Visible = False ' ������ ������ �����
Timer2.Enabled = False
Timer3.Enabled = False
Timer3.Interval = 20000
Timer_screensaver.Enabled = False
Timer_screensaver.Interval = 60000
����_����_������.Enabled = False ' ��� �������� ���� ��������� ������ �� �������
Form3.Slider1.Value = 1700
Form3.Caption = "����� ������� �������"
Form2.Caption = "����� ������������"
���_��������� = Empty
������_��������� = Empty

If FSO.FileExists("C:\�������") = True Then
    Dim ��������_��������� As Integer
    freeKanal = FreeFile
    Open "C:\�������" For Input As #freeKanal
    Input #freeKanal, �������
    Close #freeKanal
    ������� = ��������_�_1251(�������)
    Form1.Caption = "������� ������������: " + �������
    Else
    ��������_��������� = MsgBox("����� ������� �� ����������, ������� ������ ����?", 20, "������ ������ �����")
     If ��������_��������� = 6 Then
            FSO.CreateTextFile "C:\�������"
            If FSO.FileExists("C:\�������") = True Then
                MsgBox "���� ������� ������� ������", 64, "�����"
                Form1.Caption = "������� ������������: "
            End If
        End If
    
End If

Form_alarm_clock.Visible = False
Form_worked_alarm.Visible = False
Form_period_screensaver.Visible = False
Form_period_screensaver.List1.Text = 3
��������_������� = Form_period_screensaver.List1.Text


End Sub

Private Sub Form_Unload(Cancel As Integer) ' ������� ������� ����� ��� �������� ���������
    ���������_������_������ ������_�������, ������_��������_������
End
End Sub

Public Sub Timer1_Timer()

For Each FILE In ��������_�����.Files
    ������_����_������������ = True
    ���_����� = FSO.GetFileName(FILE)
    For i = 1 To ������_�������
        If ���_����� = ������_��������_������(i) Then
            ������_����_������������ = False
            Exit For
        End If
    Next
    If ������_����_������������ = True Then ' ���� ����� � ������ ��� (����� ����)
    
            If Form1.WindowState = 1 Then ' ���� ���������� �������� �� ���������������
                Form1.WindowState = 0
            End If
            
            If ����_copy_in_input.Checked = True Then ' ���� ����� ����� ���������� � _INPUT
                FSO.CopyFile FILE, "C:\_INPUT\", 1
            End If
            
            Label2.Caption = "���� " & ���_����� & " �� �������" & vbCrLf & "������� ������ ��������"
            Command2.Visible = True
            Timer2.Enabled = True
            Exit For
         
     End If
Next

End Sub

Public Sub Timer2_Timer()
    Beep Form3.Slider1.Value, 50
End Sub

Private Sub Timer3_Timer() ' ������ �������� ����������
    Dim proverka_hour As Integer
    Dim proverka_minute As Integer
    Dim proverka_msgbox As Integer
    
    proverka_hour = DatePart("h", Now)
    proverka_minute = DatePart("n", Now)
    
    If proverka_hour = ���_��������� And proverka_minute = ������_��������� Then
        Timer2.Enabled = True
        Form_worked_alarm.Visible = True
            
    End If
    
End Sub

Private Sub �����_������������_Click() ' ���� ����� ������������
Form2.Visible = True
End Sub


Private Sub ����_��������_����_Click()
    Dim ��������_��������� As Integer
    If FSO.FileExists("C:\����") = True Then
        freeKanal = FreeFile
        Open "C:\����" For Output As #freeKanal
        �������_���� = Date
        Print #freeKanal, �������_����
        Close #freeKanal
    Else
        ��������_��������� = MsgBox("����� ���� �� ����������, ������� ������ ����?", 20, "������ ������ �����")
        If ��������_��������� = 6 Then
            FSO.CreateTextFile "C:\����"
            If FSO.FileExists("C:\����") = True Then
                MsgBox "���� ���� ������� ������", 64, "�����"
                freeKanal = FreeFile
                Open "C:\����" For Output As #freeKanal
                �������_���� = Date
                Print #freeKanal, �������_����
                Close #freeKanal
            End If
        End If
    End If

End Sub

Private Sub ����_��������_�������_Click() ' ����� ���� �������� �������

If FSO.FileExists("C:\TS_NET\jur.txt") = True Then
   Open "C:\TS_NET\jur.txt" For Output As #2
   Print #2, Chr(42) & Chr(42) & Chr(42) & Chr(42) & Chr(42) & Chr(42) & Chr(42) & Chr(42) & Chr(42) & Chr(42) ' ������� ������
End If

If FSO.FileExists("C:\TS_NET\jurotpr.txt") = True Then
    Open "C:\TS_NET\jurotpr.txt" For Output As #3
    Print #3, Chr(42) & Chr(42) & Chr(42) & Chr(42) & Chr(42) & Chr(42) & Chr(42) & Chr(42) & Chr(42) & Chr(42)
End If

If FSO.FileExists("C:\TS_NET\ts_file.prn") = True Then
    Open "C:\TS_NET\ts_file.prn" For Output As #4
    Print #4, Chr(42) & Chr(42) & Chr(42) & Chr(42) & Chr(42) & Chr(42) & Chr(42) & Chr(42) & Chr(42) & Chr(42)
End If

If FSO.FileExists("C:\TS_NET\ts_mfl.txt") = True Then
    Open "C:\TS_NET\ts_mfl.txt" For Output As #5
    Print #5, Chr(42) & Chr(42) & Chr(42) & Chr(42) & Chr(42) & Chr(42) & Chr(42) & Chr(42) & Chr(42) & Chr(42)
End If

If FSO.FileExists("C:\TS_NET\jurotpr.tbi") = True Then
    Open "C:\TS_NET\jurotpr.tbi" For Output As #6
End If

Reset

If FSO.FolderExists("C:\TS_NET\INPUT") = True Then ' �������� ��� �������� �� ������� �������� C:\TS_NET\INPUT
        Set FOLDER = FSO.GetFolder("C:\TS_NET\INPUT")
        For Each FILE In FOLDER.Files ' ������� �������� C:\TS_NET\INPUT
            PATH = FSO.GetAbsolutePathName(FILE)
            FSO.DeleteFile PATH, 1
        Next
End If

If FSO.FolderExists("C:\TS_NET\INKV") = True Then ' �������� ��� �������� �� ������� �������� C:\TS_NET\INKV
        Set FOLDER = FSO.GetFolder("C:\TS_NET\INKV")
        For Each FILE In FOLDER.Files ' ������� �������� C:\TS_NET\INKV
            PATH = FSO.GetAbsolutePathName(FILE)
            FSO.DeleteFile PATH, 1
        Next
End If

If FSO.FolderExists("C:\TS_NET\OUTPUT") = True Then ' �������� ��� �������� �� ������� �������� C:\TS_NET\OUTPUT
        Set FOLDER = FSO.GetFolder("C:\TS_NET\OUTPUT")
        For Each FILE In FOLDER.Files ' ������� �������� C:\TS_NET\OUTPUT
            PATH = FSO.GetAbsolutePathName(FILE)
            FSO.DeleteFile PATH, 1
        Next
End If


End Sub

Private Sub ����_��������_�������_��������_Click() ' ����� ���� �������� ������� ��������

    For Each FILE In ��������_�����.Files ' ������� �������� _INPUT ��� ������� ���������� ��������
        PATH = FSO.GetAbsolutePathName(FILE)
        FSO.DeleteFile PATH, 1
    Next
    
    If ����_copy_in_input.Checked = True Then ' ���� ����� ����� ���������� � _INPUT
        Set FOLDER = FSO.GetFolder("C:\_INPUT")
        For Each FILE In FOLDER.Files ' ������  �������� _INPUT �������������
            PATH = FSO.GetAbsolutePathName(FILE)
            FSO.DeleteFile PATH, 1
        Next
    End If
    
    If FSO.FolderExists("C:\_KVITOK") = True Then ' �������� ��� �������� �� ������� ��������
        Set FOLDER = FSO.GetFolder("C:\_KVITOK")
        
        For Each FILE In FOLDER.Files ' ������  �������� _KVITOK
            PATH = FSO.GetAbsolutePathName(FILE)
            FSO.DeleteFile PATH, 1
        Next
        
    End If
    
    If FSO.FolderExists("C:\WORK\OUTPUT") = True Then ' �������� ��� �������� �� ������� �������� WORK\OUTPUT
        Set FOLDER = FSO.GetFolder("C:\WORK\OUTPUT")
        For Each FILE In FOLDER.Files ' �������  � �������� �������� WORK\OUTPUT
            PATH = FSO.GetAbsolutePathName(FILE)
            FSO.DeleteFile PATH, 1
        Next
    FSO.DeleteFolder FOLDER, 1
    End If
    
    If FSO.FolderExists("C:\WORK\KV_IN") = True Then ' �������� ��� �������� �� ������� �������� C:\WORK\KV_IN
        Set FOLDER = FSO.GetFolder("C:\WORK\KV_IN")
        For Each FILE In FOLDER.Files ' �������  � �������� �������� C:\WORK\KV_IN
            PATH = FSO.GetAbsolutePathName(FILE)
            FSO.DeleteFile PATH, 1
        Next
    FSO.DeleteFolder FOLDER, 1
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    If FSO.FolderExists("C:\������") = True Then ' �������� ��� �������� �� ������� �������� C:\������
        Set FOLDER = FSO.GetFolder("C:\������")
        Dim �����_���_�������� As Integer
        freeKanal = FreeFile
        Open "C:\TS_NET\NUM_TLG.ISH" For Input As #freeKanal
        Input #freeKanal, �����_���_��������
        Close #freeKanal
                    
        For Each FILE In FOLDER.Files
            ���_�����_������� = FSO.GetBaseName(FILE)
            Trim (���_�����_�������)
            If (Right(���_�����_�������, 3) = "TLG" Or Right(���_�����_�������, 3) = "tlg") Then
                �����_���_������� = Val(���_�����_�������)
                If (�����_���_������� >= �����_���_��������) Then
                    �����_���_�������� = �����_���_�������
                    freeKanal = FreeFile
                    Open "C:\TS_NET\NUM_TLG.ISH" For Output As #freeKanal
                    Print #freeKanal, Trim(Str(�����_���_�������� + 1))
                    Close #freeKanal
                End If
            End If
        Next
        For Each FILE In FOLDER.Files ' �������  �������� C:\������
            PATH = FSO.GetAbsolutePathName(FILE)
            FSO.DeleteFile PATH, 1
        Next
    End If
    
    
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     
     If FSO.FolderExists("C:\WORK\INPUT") = True Then ' �������� ��� �������� �� ������� �������� C:\WORK\INPUT
        Set FOLDER = FSO.GetFolder("C:\WORK\INPUT")
       
        ��������_msgbox = MsgBox("������� ������� C:\WORK\INPUT ?", 36, "�����")
        If ��������_msgbox = 6 Then
            For Each FILE In FOLDER.Files ' ������� � �������� �������� C:\WORK\INPUT
                PATH = FSO.GetAbsolutePathName(FILE)
                FSO.DeleteFile PATH, 1
            Next
        End If
        FSO.DeleteFolder FOLDER, 1
    End If
 
End Sub

Private Sub ����_�����_���������_Click() '����� ���� ����� ���������
    ����_��������_����_Click
    ����_��������_�������_Click
    ����_��������_�������_��������_Click
    ����_��������_������_Click
    MsgBox "����� ��������� �����������", 64
End Sub

Private Sub ����_��������_������_Click() ' ����� ���� �������� �����

Erase ������_��������_������
�������_������ = 1
������_������� = 15
ReDim ������_��������_������(������_�������)

End Sub

Private Sub ����_��������_�����_Click() ' ����� ���� �������� �����

������_��������_������ = vbNullString
For i = 1 To ������_�������
    If ������_��������_������(i) = "" Then
        Exit For
        End If
    ������_��������_������ = ������_��������_������ & vbCrLf & i & " " & ������_��������_������(i)
Next
MsgBox ������_��������_������

End Sub

Public Sub Command2_Click() ' ������ ������ ����� ������������ �� �����
    ������_��������_������(�������_������) = ���_�����
    �������_������ = �������_������ + 1
    If �������_������ = ������_������� - 1 Then
        ������_������� = ������_������� + ������_�������
        ReDim Preserve ������_��������_������(������_�������)
    End If
        
    Timer2.Enabled = False
    Timer1.Enabled = True
    Label2.Caption = ""
    Command2.Visible = False
End Sub

Private Sub ����_�������_�������_Click() ' ����� ���� ������� �������
Form3.Visible = True
Form3.Label1.ForeColor = vbBlue
Form3.Label1.Caption = "��������� ��������: " & Form3.Slider1.Value
Form3.Slider1.SelStart = Form3.Slider1.Value
End Sub

Public Sub ����_�������_��������_Click() ' ����� ���� ������� �������� �������
MsgBox Form3.Slider1.Value
End Sub

Private Sub ����_���������_����������_Click() ' ����� ���� ��������� ����������
    Form_alarm_clock.Visible = True
End Sub

Private Sub ����_�������_���������_Click() ' ����� ���� ������� ���������
    ���_��������� = Empty
    ������_��������� = Empty
    Form1.Timer3.Enabled = False
End Sub

Private Sub ����_��������_����������_Click()
    Dim stroka As String
    If ���_��������� = Empty And ������_��������� = Empty Then
        stroka = "��������� �� ����������"
    Else
        stroka = CStr(���_���������) + " " + ":" + " " + CStr(������_���������)
    End If
    MsgBox stroka
End Sub

Private Sub ����_���������_���������������_Click()
    MsgBox "ANTI-SCREENSAVER �������", 64
    Timer_screensaver.Enabled = True
    ������ = 0
    ����_���������_���������������.Enabled = False
    ����_����_������.Enabled = True
    
End Sub

Private Sub ����_����_������_Click()
    MsgBox "ANTI-SCREENSAVER ��������", 48
    Timer_screensaver.Enabled = False
    ����_���������_���������������.Enabled = True
    ����_����_������.Enabled = False
 End Sub



Private Sub ����_������_�������_Click()
    Form_period_screensaver.Visible = True
End Sub

Private Sub Timer_screensaver_Timer()
   ������ = ������ + 1
   
   If ((������ = ��������_�������) Or ((������ Mod ��������_�������) = 0)) Then
        Command1_Click ' ���������� �������
        ������ = 0
   End If
End Sub

Private Sub ����_�����_�����_Click()
     Form_select_catalog.Visible = True
End Sub

Private Sub ����_copy_in_input_Click()
    
    If ����_copy_in_input.Checked = False Then
        ����_copy_in_input.Checked = True
    Else
        ����_copy_in_input.Checked = False
    End If
    
End Sub

Private Sub ���������_Click()
    If FSO.FileExists("C:\����") = True Then
        freeKanal = FreeFile
        Open "C:\����" For Input As #freeKanal
        Input #freeKanal, �������_����
        Close #freeKanal
        If (�������_���� = Date) Then
            ����_��������_����.Checked = True
        Else
        ����_��������_����.Checked = False
    End If
End If


End Sub
