VERSION 5.00
Begin VB.Form Form_period_screensaver 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4395
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4395
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "œ–»Õﬂ“‹"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      TabIndex        =   2
      Top             =   1920
      Width           =   2055
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      ItemData        =   "Form_period_screensaver.frx":0000
      Left            =   480
      List            =   "Form_period_screensaver.frx":0022
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Ã»Õ”“€"
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
      Left            =   2280
      TabIndex        =   0
      Top             =   720
      Width           =   1695
   End
End
Attribute VB_Name = "Form_period_screensaver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Form_period_screensaver.Caption = "œ≈–»Œƒ Õ¿∆¿“»ﬂ"
 End Sub
Private Sub Command1_Click()
    ÁÌ‡˜ÂÌËÂ_Ì‡Ê‡ÚËˇ = Form_period_screensaver.List1.Text
    Form_period_screensaver.Visible = False
End Sub


