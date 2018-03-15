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
      Caption         =   "ѕ–»Ќя“№"
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
Dim им€_пользовател€ As String
Dim конвертированное_им€_пользовател€ As String
Dim i As Byte
Dim буква As String


Private Sub ковертер_в_866(name As String)
конвертированное_им€_пользовател€ = name
    For i = 1 To Len(конвертированное_им€_пользовател€)
        буква = (Mid(name, i, 1))
        Select Case (буква)
            Case "ј"
            Mid(конвертированное_им€_пользовател€, i, 1) = "А"
            Case "Ѕ"
            Mid(конвертированное_им€_пользовател€, i, 1) = "Б"
            Case "¬"
            Mid(конвертированное_им€_пользовател€, i, 1) = "В"
            Case "√"
            Mid(конвертированное_им€_пользовател€, i, 1) = "Г"
            Case "ƒ"
            Mid(конвертированное_им€_пользовател€, i, 1) = "Д"
            Case "≈"
            Mid(конвертированное_им€_пользовател€, i, 1) = "Е"
            Case "®"
            Mid(конвертированное_им€_пользовател€, i, 1) = "р"
            Case "∆"
            Mid(конвертированное_им€_пользовател€, i, 1) = "Ж"
            Case "«"
            Mid(конвертированное_им€_пользовател€, i, 1) = "З"
            Case "»"
            Mid(конвертированное_им€_пользовател€, i, 1) = "И"
            Case "…"
            Mid(конвертированное_им€_пользовател€, i, 1) = "Й"
            Case " "
            Mid(конвертированное_им€_пользовател€, i, 1) = "К"
            Case "Ћ"
            Mid(конвертированное_им€_пользовател€, i, 1) = "Л"
            Case "ћ"
            Mid(конвертированное_им€_пользовател€, i, 1) = "М"
            Case "Ќ"
            Mid(конвертированное_им€_пользовател€, i, 1) = "Н"
            Case "ќ"
            Mid(конвертированное_им€_пользовател€, i, 1) = "О"
            Case "ѕ"
            Mid(конвертированное_им€_пользовател€, i, 1) = "П"
            Case "–"
            Mid(конвертированное_им€_пользовател€, i, 1) = "Р"
            Case "—"
            Mid(конвертированное_им€_пользовател€, i, 1) = "С"
            Case "“"
            Mid(конвертированное_им€_пользовател€, i, 1) = "Т"
            Case "”"
            Mid(конвертированное_им€_пользовател€, i, 1) = "У"
            Case "‘"
            Mid(конвертированное_им€_пользовател€, i, 1) = "Ф"
            Case "’"
            Mid(конвертированное_им€_пользовател€, i, 1) = "Х"
            Case "÷"
            Mid(конвертированное_им€_пользовател€, i, 1) = "Ц"
            Case "„"
            Mid(конвертированное_им€_пользовател€, i, 1) = "Ч"
            Case "Ў"
            Mid(конвертированное_им€_пользовател€, i, 1) = "Ш ШШШ"
            Case "ў"
            Mid(конвертированное_им€_пользовател€, i, 1) = "Щ"
            Case "Џ"
            Mid(конвертированное_им€_пользовател€, i, 1) = "Ъ"
            Case "џ"
            Mid(конвертированное_им€_пользовател€, i, 1) = "Ы"
            Case "№"
            Mid(конвертированное_им€_пользовател€, i, 1) = "Ь"
            Case "Ё"
            Mid(конвертированное_им€_пользовател€, i, 1) = "Э"
            Case "ё"
            Mid(конвертированное_им€_пользовател€, i, 1) = "Ю"
            Case "я"
            Mid(конвертированное_им€_пользовател€, i, 1) = "Я"
            End Select
         Next
End Sub
Private Sub Command1_Click() ' ¬џЅќ– ѕќЋ№«ќ¬ј“≈Ћя
им€_пользовател€ = Combo1.Text
Form1.Caption = "“≈ ”ў»… ѕќЋ№«ќ¬ј“≈Ћ№: " + им€_пользовател€
Form2.Visible = False
ковертер_в_866 (им€_пользовател€)
freeKanal = FreeFile
Open "C:\‘јћ»Ћ»я" For Output As #freeKanal
Print #freeKanal, конвертированное_им€_пользовател€
Close #freeKanal
End Sub

Private Sub Label1_Click()

End Sub
