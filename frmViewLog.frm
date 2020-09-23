VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmViewLog 
   Caption         =   "IP to Serial Port Connection Log"
   ClientHeight    =   5130
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6870
   Icon            =   "frmViewLog.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5130
   ScaleWidth      =   6870
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox RTB 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   8705
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frmViewLog.frx":0442
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Log"
      Begin VB.Menu mnuClear 
         Caption         =   "&Clear Log"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu mnuFileBar5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuSelect 
         Caption         =   "&Select All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUndo 
         Caption         =   "Undo Selection"
         Shortcut        =   ^Z
      End
   End
End
Attribute VB_Name = "frmViewLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub mnuClear_Click()
On Error Resume Next
Kill App.Path & "\" & "Connection.log"
mnuRefresh_Click
End Sub

Private Sub mnuCopy_Click()
Clipboard.Clear
SendKeys "^C", True
RTB.SelStart = (Len(RTB.Text) + 1)
End Sub

Private Sub mnuFileExit_Click()
  'unload the form
  Unload Me
End Sub

Private Sub Form_Load()

Dim LogFile As Integer
Dim LogText As String

frmCommSetUp.mnuLog.Checked = True

LogFile = FreeFile
On Error GoTo FileERR
Open App.Path & "\" & "Connection.log" For Input As #LogFile
    Do Until EOF(LogFile) = True
        Line Input #LogFile, LogText
        RTB.SelText = LogText & vbCrLf
        RTB.SelStart = (Len(RTB.Text) + 1)
    Loop
RTB.Locked = True
Close #LogFile
Exit Sub
FileERR:
End Sub

Private Sub Form_Resize()
    'Resize the form
    ResizeFormFor Me
    
    If frmViewLog.WindowState <> vbMinimized Then
        mnuRefresh_Click
    End If

End Sub

Private Sub Form_Terminate()
frmCommSetUp.mnuLog.Checked = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmCommSetUp.mnuLog.Checked = False
End Sub

Public Sub mnuRefresh_Click()

Dim LogFile As Integer
Dim LogText As String

RTB.Locked = False
RTB.Text = ""
RTB.SelStart = (Len(RTB.Text) + 1)
LogFile = FreeFile
On Error GoTo RefERR
Open App.Path & "\" & "Connection.log" For Input As #LogFile
    Do Until EOF(LogFile) = True
        Line Input #LogFile, LogText
        RTB.SelText = LogText & vbCrLf
        RTB.SelStart = (Len(RTB.Text) + 1)
    Loop
RTB.Locked = True
Close #LogFile
Exit Sub
RefERR:
End Sub

Private Sub mnuSelect_Click()
RTB.SelStart = 0
RTB.SelLength = Len(RTB.Text)

SendKeys "^A", True
End Sub

Private Sub mnuUndo_Click()
RTB.SelStart = (Len(RTB.Text) + 1)
End Sub

Private Sub RTB_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
PopupMenu mnuEdit
End If
End Sub
