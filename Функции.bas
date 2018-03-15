Attribute VB_Name = "Функции"
Option Explicit
Option Base 1
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Const KEYEVENTF_KEYUP = &H2 'событие отпускания клавиши
Const VK_LCONTROL = &HA2 ' Левый Ctrl
Const VK_RCONTROL = &HA3 ' Правый Ctrl
Const VK_ESCAPE = &H1B  'клавиша Escape
Const VK_LWIN = &H5B 'левая клавиша, эмулирующая нажатие кнопки ПУСК
Const VK_RWIN = &H5B 'левая клавиша, эмулирующая нажатие кнопки ПУСК
Const VK_LMENU = &HA4 ' Левый Alt
Const VK_RMENU = &HA5 ' Правый Alt
Const VK_SHIFT = &H10 ' Shift
 
Sub Command1_Click()
    Call keybd_event(VK_SHIFT, 0, 0, 0) 'Hажимаем
    Call keybd_event(VK_SHIFT, 0, KEYEVENTF_KEYUP, 0) 'Отпускаем
End Sub

Function ковертер_в_1251(name As String) As String
Dim конвертированное_имя_пользователя As String
Dim буква As String
Dim i As Integer

конвертированное_имя_пользователя = name
    For i = 1 To Len(конвертированное_имя_пользователя)
        буква = (Mid(name, i, 1))
        Select Case (буква)
            Case "Ђ"
            Mid(конвертированное_имя_пользователя, i, 1) = "А"
            Case "Ѓ"
            Mid(конвертированное_имя_пользователя, i, 1) = "Б"
            Case "‚"
            Mid(конвертированное_имя_пользователя, i, 1) = "В"
            Case "ѓ"
            Mid(конвертированное_имя_пользователя, i, 1) = "Г"
            Case "„"
            Mid(конвертированное_имя_пользователя, i, 1) = "Д"
            Case "…"
            Mid(конвертированное_имя_пользователя, i, 1) = "Е"
            Case "р"
            Mid(конвертированное_имя_пользователя, i, 1) = "Ё"
            Case "†"
            Mid(конвертированное_имя_пользователя, i, 1) = "Ж"
            Case "‡"
            Mid(конвертированное_имя_пользователя, i, 1) = "З"
            Case "€"
            Mid(конвертированное_имя_пользователя, i, 1) = "И"
            Case "‰"
            Mid(конвертированное_имя_пользователя, i, 1) = "Й"
            Case "Љ"
            Mid(конвертированное_имя_пользователя, i, 1) = "К"
            Case "‹"
            Mid(конвертированное_имя_пользователя, i, 1) = "Л"
            Case "Њ"
            Mid(конвертированное_имя_пользователя, i, 1) = "М"
            Case "Ќ"
            Mid(конвертированное_имя_пользователя, i, 1) = "Н"
            Case "Ћ"
            Mid(конвертированное_имя_пользователя, i, 1) = "О"
            Case "Џ"
            Mid(конвертированное_имя_пользователя, i, 1) = "П"
            Case "ђ"
            Mid(конвертированное_имя_пользователя, i, 1) = "Р"
            Case "‘"
            Mid(конвертированное_имя_пользователя, i, 1) = "С"
            Case "’"
            Mid(конвертированное_имя_пользователя, i, 1) = "Т"
            Case "“"
            Mid(конвертированное_имя_пользователя, i, 1) = "У"
            Case "”"
            Mid(конвертированное_имя_пользователя, i, 1) = "Ф"
            Case "•"
            Mid(конвертированное_имя_пользователя, i, 1) = "Х"
            Case "–"
            Mid(конвертированное_имя_пользователя, i, 1) = "Ц"
            Case "—"
            Mid(конвертированное_имя_пользователя, i, 1) = "Ч"
            Case Chr(152)
            Mid(конвертированное_имя_пользователя, i, 1) = "Ш"
            Case "™"
            Mid(конвертированное_имя_пользователя, i, 1) = "Щ"
            Case "љ"
            Mid(конвертированное_имя_пользователя, i, 1) = "Ъ"
            Case "›"
            Mid(конвертированное_имя_пользователя, i, 1) = "Ы"
            Case "њ"
            Mid(конвертированное_имя_пользователя, i, 1) = "Ь"
            Case "ќ"
            Mid(конвертированное_имя_пользователя, i, 1) = "Э"
            Case "ћ"
            Mid(конвертированное_имя_пользователя, i, 1) = "Ю"
            Case "џ"
            Mid(конвертированное_имя_пользователя, i, 1) = "Я"
            End Select
         Next
         ковертер_в_1251 = конвертированное_имя_пользователя
End Function

Function Сохранить_список_файлов(размер_массива As Byte, массив_входящих_файлов() As String) As Integer ' сохраняет файлы в текстовый документ
    Dim i As Integer ' счетчик записанных файлов
    Dim j As Integer ' счетчик цикла
    Dim s As String
    freeKanal = FreeFile
    Open "C:\TS_NET\log.txt" For Output As #freeKanal
    i = 0
    For j = 1 To размер_массива
        If массив_входящих_файлов(j) = "" Then
            Exit For
        End If
    Print #freeKanal, массив_входящих_файлов(j)
    i = i + 1
  
    Next
    Close #freeKanal
   
   Сохранить_список_файлов = i
   
End Function

Function Загрузить_список_файлов(размер_массива As Byte, массив_входящих_файлов() As String) As Integer ' загружает из документа список файлов
    Dim i As Integer ' счетчик записанных файлов
    Dim j As Integer ' счетчик цикла
    
    freeKanal = FreeFile
    Open "C:\TS_NET\log.txt" For Input As #freeKanal
    j = 1
    i = 0
    Do While Not EOF(freeKanal)
        Line Input #freeKanal, массив_входящих_файлов(j)
        j = j + 1
        i = i + 1
        счетчик_записи = счетчик_записи + 1
        If j = размер_массива - 1 Then
            размер_массива = размер_массива + размер_массива
            ReDim Preserve массив_входящих_файлов(размер_массива)
        End If
    Loop
    Close #freeKanal
    
    freeKanal = FreeFile
    Open "C:\TS_NET\log.txt" For Output As #freeKanal
    Close #freeKanal
        
    Загрузить_список_файлов = i
End Function
