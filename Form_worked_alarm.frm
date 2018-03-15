VERSION 5.00
Begin VB.Form Form_worked_alarm 
   Caption         =   "БУДИЛЬНИК"
   ClientHeight    =   3030
   ClientLeft      =   12465
   ClientTop       =   9570
   ClientWidth     =   4560
   LinkTopic       =   "Form4"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   Begin VB.CommandButton Command1 
      Caption         =   "ОСТАНОВИТЬ"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   1680
      Width           =   3735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Центровка
      Caption         =   "БУДИЛЬНИК СРАБОТАЛ"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3855
   End
End
Attribute VB_Name = "Form_worked_alarm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Form1.Timer2.Enabled = False
    Form1.Timer3.Enabled = False
    час_будильник = Empty
    минута_будильник = Empty
    Form_worked_alarm.Visible = False
End Sub
