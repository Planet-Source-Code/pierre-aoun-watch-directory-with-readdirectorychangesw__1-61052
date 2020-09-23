Attribute VB_Name = "modFunction"
Option Explicit
Public Const MaxMumDir = 2
Public Const INFINITE = &HFFFF      '  Infinite timeout
Public Const CREATE_SUSPENDED = &H4
Global DirHndl(0 To MaxMumDir - 1) As Long
Global WSubFolder(0 To MaxMumDir - 1) As Boolean
Global CollectA1 As Collection
Global CollectA2 As Collection
Global CollectB1 As Collection
Global CollectB2 As Collection
Global CA1ReadyToRead As Boolean
Global CA2ReadyToRead As Boolean
Global CB1ReadyToRead As Boolean
Global CB2ReadyToRead As Boolean
Global DirPath(0 To MaxMumDir - 1) As String


Type FILE_NOTIFY_INFORMATION
   NextEntryOffset As Long
   Action As Long
   FileNameLength As Long
   FileName As String
End Type
Public Const FILE_FLAG_OVERLAPPED = &H40000000
Public Const FILE_LIST_DIRECTORY = &H1
Public Const FILE_SHARE_READ = &H1&
Public Const FILE_SHARE_DELETE = &H4&
Public Const OPEN_EXISTING = 3
Public Const FILE_FLAG_BACKUP_SEMANTICS = &H2000000
Public Const FILE_NOTIFY_CHANGE_FILE_NAME = &H1&
Public Const FILE_NOTIFY_CHANGE_LAST_WRITE = &H10&
Public Const FILE_SHARE_WRITE As Long = &H2
Public Const FILE_NOTIFY_CHANGE_ATTRIBUTES As Long = &H4
Public Const FILE_NOTIFY_CHANGE_DIR_NAME As Long = &H2

Public Const FILE_ACTION_ADDED = &H1&
Public Const FILE_ACTION_REMOVED = &H2&
Public Const FILE_ACTION_MODIFIED = &H3&
Public Const FILE_ACTION_RENAMED_OLD_NAME = &H4&
Public Const FILE_ACTION_RENAMED_NEW_NAME = &H5&


Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function ReadDirectoryChangesW Lib "kernel32" (ByVal hDirectory As Long, lpBuffer As Any, ByVal nBufferLength As Long, ByVal bWatchSubtree As Long, ByVal dwNotifyFilter As Long, lpBytesReturned As Long, ByVal PassZero As Long, ByVal PassZero As Long) As Long
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal PassZero As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal PassZero As Long) As Long
Public Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As Any, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
Public Declare Function TerminateThread Lib "kernel32" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function ResumeThread Lib "kernel32" (ByVal hThread As Long) As Long
Public Declare Function SuspendThread Lib "kernel32" (ByVal hThread As Long) As Long

Public Function GetDirHndl(ByVal PathDir As String) As Long
 On Error Resume Next
 Dim hDir As Long
 If Right(PathDir, 1) <> "\" Then PathDir = PathDir + "\"
 hDir = CreateFile(PathDir, FILE_LIST_DIRECTORY, FILE_SHARE_READ + FILE_SHARE_WRITE + FILE_SHARE_DELETE, _
                   ByVal 0&, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS Or FILE_FLAG_OVERLAPPED, ByVal 0&)
 GetDirHndl = hDir
End Function

Public Sub ClearHndl(Handle As Long)
 CloseHandle Handle
 Handle = 0
End Sub

