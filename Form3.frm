VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form3"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4590
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "ѕ–»Ќя“№"
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   1800
      Width           =   1695
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   675
      Left            =   600
      TabIndex        =   0
      Top             =   840
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1191
      _Version        =   393216
      Min             =   1
      Max             =   10000
      SelStart        =   1500
      TickStyle       =   3
      Value           =   1500
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   240
      Width           =   3495
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Command1_Click()
Form3.Visible = False
End Sub

Private Sub Slider1_Click()
    Form3.Label1.Caption = "¬ыбранное значение: " & Form3.Slider1.Value
End Sub
