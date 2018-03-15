VERSION 5.00
Begin VB.Form Form_alarm_clock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "УСТАНОВКА БУДИЛЬНИКА"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6735
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   6735
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btn_set_time 
      Caption         =   "УСТАНОВИТЬ ВРЕМЯ"
      Height          =   735
      Left            =   1200
      TabIndex        =   4
      Top             =   2160
      Width           =   3615
   End
   Begin VB.ListBox List_minute 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1140
      ItemData        =   "Form_alarm_clock.frx":0000
      Left            =   3480
      List            =   "Form_alarm_clock.frx":0028
      TabIndex        =   1
      Top             =   840
      Width           =   1335
   End
   Begin VB.ListBox List_hour 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1140
      ItemData        =   "Form_alarm_clock.frx":005C
      Left            =   1200
      List            =   "Form_alarm_clock.frx":00A8
      TabIndex        =   0
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label_minute 
      Alignment       =   2  'Центровка
      Caption         =   "МИНУТЫ"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label_hour 
      Alignment       =   2  'Центровка
      Caption         =   "ЧАСЫ"
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form_alarm_clock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub btn_set_time_Click()
    час_будильник = List_hour.Text
    минута_будильник = List_minute.Text
    Form1.Timer3.Enabled = True
    Form_alarm_clock.Visible = False
End Sub
