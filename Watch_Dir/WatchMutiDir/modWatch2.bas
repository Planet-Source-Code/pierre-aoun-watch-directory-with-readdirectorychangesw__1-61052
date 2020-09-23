Attribute VB_Name = "modWatch2"
Option Explicit
'Get the directory chages using ReadDirectoryChangesW
Private Const FILE_NOTIF_GLOB = FILE_NOTIFY_CHANGE_ATTRIBUTES Or _
                                FILE_NOTIFY_CHANGE_FILE_NAME Or _
                                FILE_NOTIFY_CHANGE_DIR_NAME Or _
                                FILE_NOTIFY_CHANGE_ATTRIBUTES Or _
                                FILE_NOTIFY_CHANGE_LAST_WRITE
    Private Const nBufLen = 1024
    Private nReadLen As Long
    Private sAction As String
    Private fiBuffer As FILE_NOTIFY_INFORMATION
    Private cBuffer(0 To nBufLen) As Byte
    Private fName As String
    Private i As Integer
    Private Action As Long
    Private FileNameLength As Long
    Private FileName As String
    Private UseACollect As Boolean
    
Public Sub StartWatch2()
UseACollect = True
  Do
   If (DirHndl(1) = 0) Or (DirHndl(1) = -1) Then Exit Do
   For i = 0 To nBufLen: cBuffer(i) = 0: Next i
   Call ReadDirectoryChangesW(DirHndl(1), cBuffer(0), nBufLen, WSubFolder(1), FILE_NOTIF_GLOB, nReadLen, 0, 0)
   Action = cBuffer(4)
   FileNameLength = CLng(CStr(cBuffer(11)) + CStr(cBuffer(10)) + CStr(cBuffer(9)) + CStr(cBuffer(8)))
   FileName = ""
   For i = 0 To FileNameLength - 1 Step 2
    FileName = FileName + Chr(cBuffer(i + 12))
   Next i
   Select Case Action
            Case FILE_ACTION_ADDED
                sAction = "Added file"
            Case FILE_ACTION_REMOVED
                sAction = "Removed file"
            Case FILE_ACTION_MODIFIED
                sAction = "Modified file"
            Case FILE_ACTION_RENAMED_OLD_NAME
                sAction = "Renamed from"
            Case FILE_ACTION_RENAMED_NEW_NAME
                sAction = "Renamed to"
            Case Else
                sAction = "Unknown"
   End Select
   fName = sAction + "-" + DirPath(1) + FileName
   If sAction <> "Unknown" Then
    If UseACollect Then  'I can use CollectA2
        CollectA2.Add fName
        If CB2ReadyToRead = False Then CA2ReadyToRead = True: UseACollect = False
    Else             'I can use CollectB1
        CollectB2.Add fName
        If CA2ReadyToRead = False Then CB2ReadyToRead = True: UseACollect = True
    End If
   End If
  Loop
End Sub



