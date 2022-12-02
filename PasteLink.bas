Attribute VB_Name = "PasteLink"
Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "User32" (ByVal uFormat As Long) As Long
Private Declare PtrSafe Function OpenClipboard Lib "User32" (ByVal hWnd As Long) As Long
Private Declare PtrSafe Function GetClipboardData Lib "User32" (ByVal uFormat As Long) As Long
Private Declare PtrSafe Function CloseClipboard Lib "User32" () As Long
Private Declare PtrSafe Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal drop_handle As Long, ByVal UINT As Long, _
  ByVal lpStr As String, ByVal ch As Long) As Long

Private Declare PtrSafe Function PathCanonicalize Lib "shlwapi.dll" Alias "PathCanonicalizeA" (ByVal pszBuf As String, ByVal pszPath As String) As Long
' http://www.vbforums.com/showthread.php?214494-PathCanonicalize

Private Const CF_HDROP As Long = 15

Public Function GetFiles(ByRef fileCount As Long) As String()
'
' Get file names from clipboard
'

On Error GoTo done
  

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

Function GetFilenameFromPath(ByVal strPath As String) As String
'
' Get file name from full file path (not used now)
'
    If Right$(strPath, 1) <> "\" And Len(strPath) > 0 Then
        GetFilenameFromPath = GetFilenameFromPath(Left$(strPath, Len(strPath) - 1)) + Right$(strPath, 1)
    End If
End Function

Sub EnableCtrlE()
  Application.OnKey "^e", "PasteOneLinkFromClipdoard"
End Sub

Sub DisableCtrlE()
  Application.OnKey "^e"
End Sub

Sub PasteOneLinkFromClipdoard()
Attribute PasteOneLinkFromClipdoard.VB_ProcData.VB_Invoke_Func = "e\n14"
'
' Paste hiperlink to current cell
'
' Keyboard Shortcut: Ctrl+e
'
    
    Dim a() As String, fileCount As Long, i As Long
    a = GetFiles(fileCount)
    If (fileCount <> 1) Then
        MsgBox "No files copyed"
        Exit Sub
    End If
    
    If ActiveCell.text = "" Then
      ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:=a(0), TextToDisplay:=GetFilenameFromPath(a(0))
    Else
      ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:=a(0)
    End If
End Sub


Sub TestSheetHyperlinks()
'
' Test all hyperlinks on sheet
'

Dim hl As Hyperlink
Dim addr As String
Dim absaddr As String

For Each hl In ActiveSheet.Hyperlinks 
  addr = hl.Address
  absaddr = String(1024, 0)
  PathCanonicalize absaddr, Application.ActiveWorkbook.Path & "\" + addr
  If Dir(absaddr) = "" Then       ' file is not exists
    'MsgBox absaddr
    hl.Range.Font.Bold = True
  End If
Next

End Sub
