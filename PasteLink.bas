Attribute VB_Name = "PasteLink"
Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal uFormat As Long) As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal Hwnd As Long) As Long
Private Declare Function GetClipboardData Lib "user32" (ByVal uFormat As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal drop_handle As Long, ByVal UINT As Long, ByVal lpStr As String, ByVal ch As Long) As Long

Private Const CF_HDROP As Long = 15

Private Function GetFiles(ByRef fileCount As Long) As String()
'
' Получить имена файлов, скопированнных в буфер обмена
'
    Dim hDrop As Long, i As Long
    Dim aFiles() As String, sFileName As String * 1024

    fileCount = 0

    If Not CBool(IsClipboardFormatAvailable(CF_HDROP)) Then Exit Function
    If Not CBool(OpenClipboard(0&)) Then Exit Function

    hDrop = GetClipboardData(CF_HDROP)
    If Not CBool(hDrop) Then GoTo done

    fileCount = DragQueryFile(hDrop, -1, vbNullString, 0)

    ReDim aFiles(fileCount - 1)
    For i = 0 To fileCount - 1
        DragQueryFile hDrop, i, sFileName, Len(sFileName)
        aFiles(i) = Left$(sFileName, InStr(sFileName, vbNullChar) - 1)
    Next
    GetFiles = aFiles
done:
    CloseClipboard
End Function

Private Function GetFilenameFromPath(ByVal strPath As String) As String
'
' Получение имени файла из имени и пути
'
    If Right$(strPath, 1) <> "\" And Len(strPath) > 0 Then
        GetFilenameFromPath = GetFilenameFromPath(Left$(strPath, Len(strPath) - 1)) + Right$(strPath, 1)
    End If
End Function

Sub PasleOneLinkFromClipdoard()
Attribute PasleOneLinkFromClipdoard.VB_ProcData.VB_Invoke_Func = "e\n14"
'
' Вставить в текущую ячейку гиперссылку на файл/каталог, скопированный в буфер обмена, если он один
'
    Dim A() As String, fileCount As Long, i As Long
    A = GetFiles(fileCount)
    If (fileCount <> 1) Then
        MsgBox "Эй, чувак! Нет файлов или каталогов в буфере обмена или их больше одного."
    End If
    
    ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:=A(0) ', TextToDisplay:=GetFilenameFromPath(a(0))
End Sub

' НИжЕ ПОКА НЕ ДОРАБОТАННО !!!


Function ClipboardText()
'
' Чтение из буфера обмена
'
 Set A = GetObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
  A.GetFromClipboard
  ClipboardText = A.GetText
End Function


Function ReadFileNames(FolderPath, FileName) As String
'
' Поиск файла
'
 On Error Resume Next
 Set fso = CreateObject("scripting.filesystemobject")
 Set curfold = fso.GetFolder(FolderPath)
 If Not curfold Is Nothing Then
   For Each fil In curfold.Files
    If fil.Name Like "*" & FileName Then
     ReadFileNames = fil.Name
     Exit Function
    End If
   Next
   For Each sfol In curfold.SubFolders
    ReadFileNames sfol.Path, FileName
   Next
 End If
End Function


Sub PasteLinkFromClipboard()
Attribute PasteLinkFromClipboard.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Получаем содержимое буфера, ищем этот файл в указанном каталоге, если найден и только
' один - вставляем в виде гиперссылки в текущую ячейку
'
Dim strFilePath
strFolderPath = "\\Prim-fs-serv\rdu\СРЗА\Уставки\"

strFileName = ClipboardText()

MsgBox ReadFileNames(strFolderPath, strFileName)

End Sub

