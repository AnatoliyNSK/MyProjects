VERSION 5.00
Begin VB.Form Form_select_catalog 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ВЫБОР КАТАЛОГА ДЛЯ СЛЕЖЕНИЯ"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6375
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   14.25
      Charset         =   204
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000C0&
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   6375
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton btn_select_catalog 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1890
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label Label_catalog 
      Alignment       =   2  'Центровка
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   6135
   End
End
Attribute VB_Name = "Form_select_catalog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_select_catalog_Click()
    Dim FSO As Object ' объект FileSystemObject
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set входящая_папка = FSO.GetFolder(Dir1.PATH)
    Label_catalog.Caption = входящая_папка
    
    If входящая_папка <> "C:\_INPUT" Then
         Form1.меню_режим_прометей.Enabled = True
    End If
  
End Sub

Private Sub Form_Load()
    Form_select_catalog.BackColor = RGB(101, 162, 219)
    Label_catalog.BackColor = RGB(101, 162, 219)
    Label_catalog.ForeColor = vbYellow
    Label_catalog.Caption = входящая_папка
    Dir1.PATH = "C:\"
End Sub




