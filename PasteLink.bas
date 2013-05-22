Attribute VB_Name = "PasteLink"
' PtrSafe - for x64 support
Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal uFormat As Long) As Long
Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal Hwnd As Long) As Long
Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal uFormat As Long) As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal drop_handle As Long, ByVal UINT As Long, ByVal lpStr As String, ByVal ch As Long) As Long

Private Const CF_HDROP As Long = 15

Private Function GetFiles(ByRef fileCount As Long) As String()
'
' Get file names from clipboard
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
' Get file name from full file path (not used now)
'
    If Right$(strPath, 1) <> "\" And Len(strPath) > 0 Then
        GetFilenameFromPath = GetFilenameFromPath(Left$(strPath, Len(strPath) - 1)) + Right$(strPath, 1)
    End If
End Function

Sub PasleOneLinkFromClipdoard()
Attribute PasleOneLinkFromClipdoard.VB_ProcData.VB_Invoke_Func = "e\n14"
'
' Paste hiperlink to current cell
'
    Dim A() As String, fileCount As Long, i As Long
    A = GetFiles(fileCount)
    If (fileCount <> 1) Then
       MsgBox "No such files in clipboard"
    else    
       ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:=A(0) ', TextToDisplay:=GetFilenameFromPath(a(0))
    end if   
End Sub