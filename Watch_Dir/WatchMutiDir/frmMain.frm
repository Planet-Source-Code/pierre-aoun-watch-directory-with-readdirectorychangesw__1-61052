VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9675
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   9675
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrWatch 
      Interval        =   100
      Left            =   120
      Top             =   5880
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   390
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   3240
      Width           =   1680
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   390
      Index           =   1
      Left            =   2040
      TabIndex        =   8
      Top             =   3240
      Width           =   1680
   End
   Begin VB.TextBox txtDir 
      Height          =   375
      Index           =   1
      Left            =   3960
      TabIndex        =   7
      Text            =   "d:"
      Top             =   3240
      Width           =   2415
   End
   Begin VB.CheckBox chkSubDir 
      Caption         =   "Watch Sub Folders"
      Height          =   255
      Index           =   1
      Left            =   6600
      TabIndex        =   6
      Top             =   3240
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.ListBox lstWatch 
      Height          =   2400
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   3720
      Width           =   9375
   End
   Begin VB.CheckBox chkSubDir 
      Caption         =   "Watch Sub Folders"
      Height          =   255
      Index           =   0
      Left            =   6600
      TabIndex        =   4
      Top             =   120
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.TextBox txtDir 
      Height          =   375
      Index           =   0
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
      Index           =   0
      Left            =   2040
      TabIndex        =   2
      Top             =   120
      Width           =   1680
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   390
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1680
   End
   Begin VB.ListBox lstWatch 
      Height          =   2400
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   9375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   -240
      X2              =   18735
      Y1              =   3120
      Y2              =   3120
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
Dim ThreadHandle(0 To MaxMumDir - 1)   As Long
Dim Fin As Boolean
Private Sub chkSubDir_Click(Index As Integer)
    WSubFolder(Index) = chkSubDir(Index).Value
End Sub
Private Sub cmdStart_Click(Index As Integer)
 WSubFolder(Index) = chkSubDir(Index).Value
'Get Folder Handle
  DirPath(Index) = txtDir(Index).Text
  If Right(DirPath(Index), 1) <> "\" Then DirPath(Index) = DirPath(Index) + "\"
  DirHndl(Index) = GetDirHndl(DirPath(Index))
  If (DirHndl(Index) = 0) Or (DirHndl(Index) = -1) Then MsgBox "Cannot create handle": Exit Sub
  cmdStart(Index).Enabled = False
  cmdStop(Index).Enabled = True
'Resume Thread to Watch changes
      ResumeThread ThreadHandle(Index)
      
End Sub
Private Sub cmdStop_Click(Index As Integer)
Dim i As Integer
    SuspendThread ThreadHandle(Index)
    If DirHndl(Index) <> 0 Then ClearHndl DirHndl(Index)
    cmdStop(Index).Enabled = False
    cmdStart(Index).Enabled = True
End Sub



Private Sub Form_Load()
Dim Dummy1 As Long
Dim Dummy2 As Long
Set CollectA1 = New Collection
Set CollectA2 = New Collection
Set CollectB1 = New Collection
Set CollectB2 = New Collection
 
ThreadHandle(0) = CreateThread(ByVal 0&, ByVal 0&, AddressOf StartWatch1, ByVal 0&, CREATE_SUSPENDED, Dummy1)
ThreadHandle(1) = CreateThread(ByVal 0&, ByVal 0&, AddressOf StartWatch2, ByVal 0&, CREATE_SUSPENDED, Dummy2)

End Sub

Private Sub Form_Resize()
lstWatch(0).Width = Me.Width - 300
lstWatch(1).Width = Me.Width - 300
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim i As Integer
Cancel = -1
tmrWatch.Enabled = False
TerminateThread ThreadHandle(1), 0
TerminateThread ThreadHandle(0), 0
Set CollectA1 = Nothing
Set CollectA2 = Nothing
Set CollectB1 = Nothing
Set CollectB2 = Nothing
End
End Sub


Private Sub tmrWatch_Timer()
Dim i As Long
'First Group
If CA1ReadyToRead Then
    For i = 1 To CollectA1.Count
        lstWatch(0).AddItem CollectA1.Item(i)
    Next i
    For i = 1 To CollectA1.Count
        CollectA1.Remove 1
    Next i
    CA1ReadyToRead = False
End If
If CB1ReadyToRead Then
    For i = 1 To CollectB1.Count
        lstWatch(0).AddItem CollectB1.Item(i)
    Next i
    For i = 1 To CollectB1.Count
        CollectB1.Remove 1
    Next i
    CB1ReadyToRead = False
End If
'------------------------------------------------
'Second Group
If CA2ReadyToRead Then
    For i = 1 To CollectA2.Count
        lstWatch(1).AddItem CollectA2.Item(i)
    Next i
    For i = 1 To CollectA2.Count
        CollectA2.Remove 1
    Next i
    CA2ReadyToRead = False
End If
If CB2ReadyToRead Then
    For i = 1 To CollectB2.Count
        lstWatch(1).AddItem CollectB2.Item(i)
    Next i
    For i = 1 To CollectB2.Count
        CollectB2.Remove 1
    Next i
    CB2ReadyToRead = False
End If

End Sub
