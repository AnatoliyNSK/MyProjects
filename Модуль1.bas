Attribute VB_Name = "Глобальные_переменные"
Option Explicit

Global счетчик_записи As Byte ' показывает индекс записи в массив принятых файлов
Global freeKanal As Integer ' переменная номера свободного файла для доступа к нему
Global час_будильник As Integer
Global минута_будильник As Integer
Global значение_нажатия ' количество минут для счетчика антискринсейвера
Global входящая_папка


