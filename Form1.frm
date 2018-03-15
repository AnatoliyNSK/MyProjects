VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ТЕКУЩИЙ ПОЛЬЗОВАТЕЛЬ: "
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
      Caption         =   "ЗАПИСАТЬ"
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
      Alignment       =   2  'Центровка
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
   Begin VB.Menu Настройка 
      Caption         =   "НАСТРОЙКА"
      Begin VB.Menu Выбор_пользователя 
         Caption         =   "ВЫБОР ПОЛЬЗОВАТЕЛЯ"
      End
      Begin VB.Menu меню_обновить_дату 
         Caption         =   "ОБНОВИТЬ ДАТУ"
      End
      Begin VB.Menu меню_очистить_журналы 
         Caption         =   "ОЧИСТИТЬ ЖУРНАЛЫ"
      End
      Begin VB.Menu меню_очистить_рабочие_каталоги 
         Caption         =   "ОЧИСТИТЬ РАБОЧИЕ КАТАЛОГИ"
      End
      Begin VB.Menu меню_смена_дежурства 
         Caption         =   "СМЕНА ДЕЖУРСТВА"
      End
   End
   Begin VB.Menu меню_список_файло 
      Caption         =   "СПИСОК ФАЙЛОВ"
      Begin VB.Menu меню_показать_файлы 
         Caption         =   "ПОКАЗАТЬ СПИСОК"
      End
      Begin VB.Menu меню_очистить_список 
         Caption         =   "ОЧИСТИТЬ СПИСОК"
      End
   End
   Begin VB.Menu меню_конфигурация 
      Caption         =   "КОНФИГУРАЦИЯ"
      Begin VB.Menu меню_частота_звука_спикера 
         Caption         =   "ЧАСТОТА ЗВУКА СПИКЕРА"
         Begin VB.Menu меню_сменить_частоту 
            Caption         =   "СМЕНИТЬ ЧАСТОТУ"
         End
         Begin VB.Menu меню_текущее_значение 
            Caption         =   "ТЕКУЩЕЕ ЗНАЧЕНИЕ"
         End
      End
      Begin VB.Menu меню_будильник 
         Caption         =   "БУДИЛЬНИК"
         Begin VB.Menu меню_установка_будильника 
            Caption         =   "УСТАНОВКА БУДИЛЬНИКА"
         End
         Begin VB.Menu меню_удалить_будильник 
            Caption         =   "УДАЛИТЬ БУДИЛЬНИК"
         End
         Begin VB.Menu меню_значение_будильника 
            Caption         =   "ТЕКУЩЕЕ ЗНАЧЕНИЕ"
         End
      End
      Begin VB.Menu меню_скрисейвер 
         Caption         =   "SCREENSAVER"
         Begin VB.Menu меню_активация_антискринсейвер 
            Caption         =   "АКТИВИЗИРОВАТЬ ЗАЩИТУ"
         End
         Begin VB.Menu меню_откл_защиты 
            Caption         =   "ОТКЛЮЧИТЬ ЗАЩИТУ"
         End
         Begin VB.Menu меню_период_нажатия 
            Caption         =   "ПЕРИОД НАЖАТИЯ"
         End
      End
      Begin VB.Menu меню_work_with_katalog 
         Caption         =   "РАБОТА С КАТАЛОГАМИ"
         Begin VB.Menu меню_ВЫБОР_ПАПКИ 
            Caption         =   "ВЫБОР ПАПКИ ДЛЯ СЛЕЖЕНИЯ"
         End
         Begin VB.Menu меню_copy_in_input 
            Caption         =   "КОПИРОВАНИЕ В _INPUT"
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

Dim массив_входящих_файлов() As String ' динамический массив строк для имен входящих файлов
Dim размер_массива As Byte
Dim FSO As Object ' объект FileSystemObject
Dim FILE ' очередной файл из входящей папки
Dim PATH ' путь к файлу каталога _INPUT
Dim FOLDER ' объект для разных каталогов
Dim имя_файла As String ' имя очередного файла из входящей папки
Dim имя_файла_кТелеги As String
Dim номер_тлг_текущий As Integer
Dim данный_файл_несуществует As Boolean
Dim i As Byte ' счетчик для цикла
Dim значение_msgbox As Integer
Dim список_принятых_файлов As String
Dim текущая_дата As String
Dim таймер As Integer ' счетчик для таймера антискринсейвера

Public Sub Form_Load()

Dim фамилия As String
If App.PrevInstance Then End ' запрещает запуск второй и последующей копии программмы
размер_массива = 15
счетчик_записи = 1
ReDim массив_входящих_файлов(размер_массива)
Set FSO = CreateObject("Scripting.FileSystemObject")

If FSO.FileExists("C:\TS_NET\log.txt") = True Then
    Загрузить_список_файлов размер_массива, массив_входящих_файлов
End If

If FSO.FolderExists("C:\_INPUT") = False Then ' проверка при загрузке на наличие рабочего каталога
        FSO.CreateFolder ("C:\_INPUT")
End If

Set входящая_папка = FSO.GetFolder("C:\_INPUT")
Command2.Visible = False ' кнопка записи файла
Timer2.Enabled = False
Timer3.Enabled = False
Timer3.Interval = 20000
Timer_screensaver.Enabled = False
Timer_screensaver.Interval = 60000
меню_откл_защиты.Enabled = False ' при загрузке меню отключить защиту не активна
Form3.Slider1.Value = 1700
Form3.Caption = "ВЫБОР ЧАСТОТЫ СПИКЕРА"
Form2.Caption = "ВЫБОР ПОЛЬЗОВАТЕЛЯ"
час_будильник = Empty
минута_будильник = Empty

If FSO.FileExists("C:\ФАМИЛИЯ") = True Then
    Dim проверка_сообщения As Integer
    freeKanal = FreeFile
    Open "C:\ФАМИЛИЯ" For Input As #freeKanal
    Input #freeKanal, фамилия
    Close #freeKanal
    фамилия = ковертер_в_1251(фамилия)
    Form1.Caption = "ТЕКУЩИЙ ПОЛЬЗОВАТЕЛЬ: " + фамилия
    Else
    проверка_сообщения = MsgBox("ФАЙЛА ФАМИЛИЯ НЕ СУЩЕСТВУЕТ, СОЗДАТЬ ДАННЫЙ ФАЙЛ?", 20, "ОШИБКА ЧТЕНИЯ ФАЙЛА")
     If проверка_сообщения = 6 Then
            FSO.CreateTextFile "C:\ФАМИЛИЯ"
            If FSO.FileExists("C:\ФАМИЛИЯ") = True Then
                MsgBox "ФАЙЛ ФАМИЛИЯ УСПЕШНО СОЗДАН", 64, "РУБИН"
                Form1.Caption = "ТЕКУЩИЙ ПОЛЬЗОВАТЕЛЬ: "
            End If
        End If
    
End If

Form_alarm_clock.Visible = False
Form_worked_alarm.Visible = False
Form_period_screensaver.Visible = False
Form_period_screensaver.List1.Text = 3
значение_нажатия = Form_period_screensaver.List1.Text


End Sub

Private Sub Form_Unload(Cancel As Integer) ' убирает процесс Рубин при закрытии программы
    Сохранить_список_файлов размер_массива, массив_входящих_файлов
End
End Sub

Public Sub Timer1_Timer()

For Each FILE In входящая_папка.Files
    данный_файл_несуществует = True
    имя_файла = FSO.GetFileName(FILE)
    For i = 1 To размер_массива
        If имя_файла = массив_входящих_файлов(i) Then
            данный_файл_несуществует = False
            Exit For
        End If
    Next
    If данный_файл_несуществует = True Then ' если файла в списке нет (новый файл)
    
            If Form1.WindowState = 1 Then ' если приложение свернуто то разворачивается
                Form1.WindowState = 0
            End If
            
            If меню_copy_in_input.Checked = True Then ' если стоит галка копировать в _INPUT
                FSO.CopyFile FILE, "C:\_INPUT\", 1
            End If
            
            Label2.Caption = "Файл " & имя_файла & " не записан" & vbCrLf & "нажмите кнопку записать"
            Command2.Visible = True
            Timer2.Enabled = True
            Exit For
         
     End If
Next

End Sub

Public Sub Timer2_Timer()
    Beep Form3.Slider1.Value, 50
End Sub

Private Sub Timer3_Timer() ' таймер проверки будильника
    Dim proverka_hour As Integer
    Dim proverka_minute As Integer
    Dim proverka_msgbox As Integer
    
    proverka_hour = DatePart("h", Now)
    proverka_minute = DatePart("n", Now)
    
    If proverka_hour = час_будильник And proverka_minute = минута_будильник Then
        Timer2.Enabled = True
        Form_worked_alarm.Visible = True
            
    End If
    
End Sub

Private Sub Выбор_пользователя_Click() ' МЕНЮ ВЫБОР ПОЛЬЗОВАТЕЛЯ
Form2.Visible = True
End Sub


Private Sub меню_обновить_дату_Click()
    Dim проверка_сообщения As Integer
    If FSO.FileExists("C:\ДАТА") = True Then
        freeKanal = FreeFile
        Open "C:\ДАТА" For Output As #freeKanal
        текущая_дата = Date
        Print #freeKanal, текущая_дата
        Close #freeKanal
    Else
        проверка_сообщения = MsgBox("ФАЙЛА ДАТА НЕ СУЩЕСТВУЕТ, СОЗДАТЬ ДАННЫЙ ФАЙЛ?", 20, "ОШИБКА ЧТЕНИЯ ФАЙЛА")
        If проверка_сообщения = 6 Then
            FSO.CreateTextFile "C:\ДАТА"
            If FSO.FileExists("C:\ДАТА") = True Then
                MsgBox "ФАЙЛ ДАТА УСПЕШНО СОЗДАН", 64, "РУБИН"
                freeKanal = FreeFile
                Open "C:\ДАТА" For Output As #freeKanal
                текущая_дата = Date
                Print #freeKanal, текущая_дата
                Close #freeKanal
            End If
        End If
    End If

End Sub

Private Sub меню_очистить_журналы_Click() ' пункт меню ОЧИСТИТЬ ЖУРНАЛЫ

If FSO.FileExists("C:\TS_NET\jur.txt") = True Then
   Open "C:\TS_NET\jur.txt" For Output As #2
   Print #2, Chr(42) & Chr(42) & Chr(42) & Chr(42) & Chr(42) & Chr(42) & Chr(42) & Chr(42) & Chr(42) & Chr(42) ' СИМВОЛЫ ЗВЕЗДА
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

If FSO.FolderExists("C:\TS_NET\INPUT") = True Then ' проверка при загрузке на наличие каталога C:\TS_NET\INPUT
        Set FOLDER = FSO.GetFolder("C:\TS_NET\INPUT")
        For Each FILE In FOLDER.Files ' ОЧИСТКА КАТАЛОГА C:\TS_NET\INPUT
            PATH = FSO.GetAbsolutePathName(FILE)
            FSO.DeleteFile PATH, 1
        Next
End If

If FSO.FolderExists("C:\TS_NET\INKV") = True Then ' проверка при загрузке на наличие каталога C:\TS_NET\INKV
        Set FOLDER = FSO.GetFolder("C:\TS_NET\INKV")
        For Each FILE In FOLDER.Files ' ОЧИСТКА КАТАЛОГА C:\TS_NET\INKV
            PATH = FSO.GetAbsolutePathName(FILE)
            FSO.DeleteFile PATH, 1
        Next
End If

If FSO.FolderExists("C:\TS_NET\OUTPUT") = True Then ' проверка при загрузке на наличие каталога C:\TS_NET\OUTPUT
        Set FOLDER = FSO.GetFolder("C:\TS_NET\OUTPUT")
        For Each FILE In FOLDER.Files ' ОЧИСТКА КАТАЛОГА C:\TS_NET\OUTPUT
            PATH = FSO.GetAbsolutePathName(FILE)
            FSO.DeleteFile PATH, 1
        Next
End If


End Sub

Private Sub меню_очистить_рабочие_каталоги_Click() ' пункт меню ОЧИСТИТЬ РАБОЧИЕ КАТАЛОГИ

    For Each FILE In входящая_папка.Files ' ОЧИСТКА КАТАЛОГА _INPUT или другого выбранного каталога
        PATH = FSO.GetAbsolutePathName(FILE)
        FSO.DeleteFile PATH, 1
    Next
    
    If меню_copy_in_input.Checked = True Then ' если стоит галка копировать в _INPUT
        Set FOLDER = FSO.GetFolder("C:\_INPUT")
        For Each FILE In FOLDER.Files ' ОИСТКА  КАТАЛОГА _INPUT принудительно
            PATH = FSO.GetAbsolutePathName(FILE)
            FSO.DeleteFile PATH, 1
        Next
    End If
    
    If FSO.FolderExists("C:\_KVITOK") = True Then ' проверка при загрузке на наличие каталога
        Set FOLDER = FSO.GetFolder("C:\_KVITOK")
        
        For Each FILE In FOLDER.Files ' ОИСТКА  КАТАЛОГА _KVITOK
            PATH = FSO.GetAbsolutePathName(FILE)
            FSO.DeleteFile PATH, 1
        Next
        
    End If
    
    If FSO.FolderExists("C:\WORK\OUTPUT") = True Then ' проверка при загрузке на наличие каталога WORK\OUTPUT
        Set FOLDER = FSO.GetFolder("C:\WORK\OUTPUT")
        For Each FILE In FOLDER.Files ' ОЧИСТКА  и УДАЛЕНИЕ КАТАЛОГА WORK\OUTPUT
            PATH = FSO.GetAbsolutePathName(FILE)
            FSO.DeleteFile PATH, 1
        Next
    FSO.DeleteFolder FOLDER, 1
    End If
    
    If FSO.FolderExists("C:\WORK\KV_IN") = True Then ' проверка при загрузке на наличие каталога C:\WORK\KV_IN
        Set FOLDER = FSO.GetFolder("C:\WORK\KV_IN")
        For Each FILE In FOLDER.Files ' ОЧИСТКА  и УДАЛЕНИЕ КАТАЛОГА C:\WORK\KV_IN
            PATH = FSO.GetAbsolutePathName(FILE)
            FSO.DeleteFile PATH, 1
        Next
    FSO.DeleteFolder FOLDER, 1
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    If FSO.FolderExists("C:\ТЕЛЕГИ") = True Then ' проверка при загрузке на наличие каталога C:\ТЕЛЕГИ
        Set FOLDER = FSO.GetFolder("C:\ТЕЛЕГИ")
        Dim номер_тлг_максимум As Integer
        freeKanal = FreeFile
        Open "C:\TS_NET\NUM_TLG.ISH" For Input As #freeKanal
        Input #freeKanal, номер_тлг_максимум
        Close #freeKanal
                    
        For Each FILE In FOLDER.Files
            имя_файла_кТелеги = FSO.GetBaseName(FILE)
            Trim (имя_файла_кТелеги)
            If (Right(имя_файла_кТелеги, 3) = "TLG" Or Right(имя_файла_кТелеги, 3) = "tlg") Then
                номер_тлг_текущий = Val(имя_файла_кТелеги)
                If (номер_тлг_текущий >= номер_тлг_максимум) Then
                    номер_тлг_максимум = номер_тлг_текущий
                    freeKanal = FreeFile
                    Open "C:\TS_NET\NUM_TLG.ISH" For Output As #freeKanal
                    Print #freeKanal, Trim(Str(номер_тлг_максимум + 1))
                    Close #freeKanal
                End If
            End If
        Next
        For Each FILE In FOLDER.Files ' ОЧИСТКА  КАТАЛОГА C:\ТЕЛЕГИ
            PATH = FSO.GetAbsolutePathName(FILE)
            FSO.DeleteFile PATH, 1
        Next
    End If
    
    
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     
     If FSO.FolderExists("C:\WORK\INPUT") = True Then ' проверка при загрузке на наличие каталога C:\WORK\INPUT
        Set FOLDER = FSO.GetFolder("C:\WORK\INPUT")
       
        значение_msgbox = MsgBox("Удалить каталог C:\WORK\INPUT ?", 36, "РУБИН")
        If значение_msgbox = 6 Then
            For Each FILE In FOLDER.Files ' ОЧИСТКА и удаление КАТАЛОГА C:\WORK\INPUT
                PATH = FSO.GetAbsolutePathName(FILE)
                FSO.DeleteFile PATH, 1
            Next
        End If
        FSO.DeleteFolder FOLDER, 1
    End If
 
End Sub

Private Sub меню_смена_дежурства_Click() 'пункт меню СМЕНА ДЕЖУРСТВА
    меню_обновить_дату_Click
    меню_очистить_журналы_Click
    меню_очистить_рабочие_каталоги_Click
    меню_очистить_список_Click
    MsgBox "СМЕНА ДЕЖУРСТВА ПРОИЗВЕДЕНА", 64
End Sub

Private Sub меню_очистить_список_Click() ' пункт меню очистить файлы

Erase массив_входящих_файлов
счетчик_записи = 1
размер_массива = 15
ReDim массив_входящих_файлов(размер_массива)

End Sub

Private Sub меню_показать_файлы_Click() ' пункт меню показать файлы

список_принятых_файлов = vbNullString
For i = 1 To размер_массива
    If массив_входящих_файлов(i) = "" Then
        Exit For
        End If
    список_принятых_файлов = список_принятых_файлов & vbCrLf & i & " " & массив_входящих_файлов(i)
Next
MsgBox список_принятых_файлов

End Sub

Public Sub Command2_Click() ' кнопка ЗАПИСЬ ФАЙЛА появляющаяся на форме
    массив_входящих_файлов(счетчик_записи) = имя_файла
    счетчик_записи = счетчик_записи + 1
    If счетчик_записи = размер_массива - 1 Then
        размер_массива = размер_массива + размер_массива
        ReDim Preserve массив_входящих_файлов(размер_массива)
    End If
        
    Timer2.Enabled = False
    Timer1.Enabled = True
    Label2.Caption = ""
    Command2.Visible = False
End Sub

Private Sub меню_сменить_частоту_Click() ' пункт меню СМЕНИТЬ ЧАСТОТУ
Form3.Visible = True
Form3.Label1.ForeColor = vbBlue
Form3.Label1.Caption = "Выбранное значение: " & Form3.Slider1.Value
Form3.Slider1.SelStart = Form3.Slider1.Value
End Sub

Public Sub меню_текущее_значение_Click() ' пункт меню ТЕКУЩЕЕ ЗНАЧЕНИЕ ЧАСТОТЫ
MsgBox Form3.Slider1.Value
End Sub

Private Sub меню_установка_будильника_Click() ' пункт меню УСТАНОВКА БУДИЛЬНИКА
    Form_alarm_clock.Visible = True
End Sub

Private Sub меню_удалить_будильник_Click() ' пункт меню УДАЛИТЬ БУДИЛЬНИК
    час_будильник = Empty
    минута_будильник = Empty
    Form1.Timer3.Enabled = False
End Sub

Private Sub меню_значение_будильника_Click()
    Dim stroka As String
    If час_будильник = Empty And минута_будильник = Empty Then
        stroka = "Будильник не установлен"
    Else
        stroka = CStr(час_будильник) + " " + ":" + " " + CStr(минута_будильник)
    End If
    MsgBox stroka
End Sub

Private Sub меню_активация_антискринсейвер_Click()
    MsgBox "ANTI-SCREENSAVER запущен", 64
    Timer_screensaver.Enabled = True
    таймер = 0
    меню_активация_антискринсейвер.Enabled = False
    меню_откл_защиты.Enabled = True
    
End Sub

Private Sub меню_откл_защиты_Click()
    MsgBox "ANTI-SCREENSAVER отключен", 48
    Timer_screensaver.Enabled = False
    меню_активация_антискринсейвер.Enabled = True
    меню_откл_защиты.Enabled = False
 End Sub



Private Sub меню_период_нажатия_Click()
    Form_period_screensaver.Visible = True
End Sub

Private Sub Timer_screensaver_Timer()
   таймер = таймер + 1
   
   If ((таймер = значение_нажатия) Or ((таймер Mod значение_нажатия) = 0)) Then
        Command1_Click ' глобальная функция
        таймер = 0
   End If
End Sub

Private Sub меню_ВЫБОР_ПАПКИ_Click()
     Form_select_catalog.Visible = True
End Sub

Private Sub меню_copy_in_input_Click()
    
    If меню_copy_in_input.Checked = False Then
        меню_copy_in_input.Checked = True
    Else
        меню_copy_in_input.Checked = False
    End If
    
End Sub

Private Sub Настройка_Click()
    If FSO.FileExists("C:\ДАТА") = True Then
        freeKanal = FreeFile
        Open "C:\ДАТА" For Input As #freeKanal
        Input #freeKanal, текущая_дата
        Close #freeKanal
        If (текущая_дата = Date) Then
            меню_обновить_дату.Checked = True
        Else
        меню_обновить_дату.Checked = False
    End If
End If


End Sub
