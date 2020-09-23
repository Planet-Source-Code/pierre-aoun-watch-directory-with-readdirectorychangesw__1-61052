VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   6225
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9675
   LinkTopic       =   "Form1"
   ScaleHeight     =   6225
   ScaleWidth      =   9675
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrEND 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   9000
      Top             =   120
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Watch Sub Folders"
      Height          =   255
      Left            =   6600
      TabIndex        =   4
      Top             =   120
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Text            =   "c:"
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   390
      Left            =   2040
      TabIndex        =   2
      Top             =   120
      Width           =   1680
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   390
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1680
   End
   Begin VB.ListBox List1 
      Height          =   5520
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   9375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Get the directory chages using ReadDirectoryChangesW
'Author: Pierre AOUN
'email: pierre_aoun@hotmail.com
Dim ThreadHandle   As Long
Dim Fin As Boolean
Private Sub Check1_Click()
    WSubFolder = Check1.Value
End Sub

Private Sub cmdStart_Click()
Dim Dummy As Long
Dim Changes As String
Dim WaitNum As Long
  WSubFolder = Check1.Value
  WatchStart = True
'Get Folder Handle
  FolderPath = Text1.Text
  If Right(FolderPath, 1) <> "\" Then FolderPath = FolderPath + "\"
  DirHndl = GetDirHndl(FolderPath)
  If (DirHndl = 0) Or (DirHndl = -1) Then MsgBox "Cannot create handle": Exit Sub
  cmdStart.Enabled = False
  cmdStop.Enabled = True
  'Create thread to Watch changes
Do
    ThreadHandle = CreateThread(ByVal 0&, ByVal 0&, AddressOf StartWatch, DirHndl, 0, Dummy)
    Do
    WaitNum = WaitForSingleObject(ThreadHandle, 50)
    DoEvents
    Loop Until (WaitNum = 0) Or (WatchStart = False)
    Changes = ""
    If WaitNum = 0 Then Changes = GetChanges
    If Changes <> "" Then List1.AddItem Changes
Loop Until Not WatchStart
 'Terminate the Thread & Clear Handle
If DirHndl <> 0 Then ClearHndl DirHndl
If ThreadHandle <> 0 Then Call TerminateThread(ThreadHandle, ByVal 0&): ThreadHandle = 0
End Sub
Private Sub cmdStop_Click()
WatchStart = False
cmdStop.Enabled = False
cmdStart.Enabled = True
End Sub
Private Sub Form_Resize()
List1.Width = Me.Width
End Sub
Private Sub Form_Unload(Cancel As Integer)
Cancel = -1
Fin = True
cmdStop_Click
tmrEND.Enabled = True
End Sub
Private Sub tmrEND_Timer()
    End
End Sub
