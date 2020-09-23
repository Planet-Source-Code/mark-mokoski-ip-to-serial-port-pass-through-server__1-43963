VERSION 5.00
Begin VB.MDIForm frmMonitor 
   AutoShowChildren=   0   'False
   BackColor       =   &H80000001&
   Caption         =   "IP to Serial Port Data Monitor"
   ClientHeight    =   8490
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10080
   Icon            =   "frmMonitor.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save Settings"
      End
      Begin VB.Menu mnuFileSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "Tile &Horizontal"
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "Tile &Vertical"
      End
      Begin VB.Menu mnuWindowBar1 
         Caption         =   "-"
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
         Caption         =   "&Undo Selection"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEditSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFind 
         Caption         =   "&Find"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuFindNext 
         Caption         =   "Find &Next"
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnnuHelpAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const NAME_COLUMN = 0
Const TYPE_COLUMN = 1
Const SIZE_COLUMN = 2
Const DATE_COLUMN = 3
Private WindowStyle As Integer

Private Sub MDIForm_Load()
WindowStyle = 0 'Cascade
End Sub

Public Sub MDIForm_Resize()
'Resize MDI Child Windows too !
Select Case WindowStyle
    
    Case 0
    'Cascade
        mnuWindowCascade_Click
    Case 1
    'Horizontal
        mnuWindowTileHorizontal_Click
    Case 2
    'Vertical
        mnuWindowTileVertical_Click
    
End Select

End Sub

Private Sub MDIForm_Terminate()
    Dim x As Integer
    
    For x = 0 To 3
        Set frmViewPackets(x) = Nothing
    Next x
frmCommSetUp.mnuMonitor.Checked = False
Unload frmFind

End Sub
Private Sub MDIForm_Unload(Cancel As Integer)
    Dim x As Integer
    For x = 0 To 3
        Set frmViewPackets(x) = Nothing
    Next x
frmCommSetUp.mnuMonitor.Checked = False
Unload frmFind

End Sub
Private Sub mnnuHelpAbout_Click()
    frmAbout.Show
    frmAbout.SetFocus
End Sub

Private Sub mnuCopy_Click()
Clipboard.Clear
SendKeys "^C", True
ActiveForm.RTB.SelStart = (Len(ActiveForm.RTB.Text) + 1)
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub
Private Sub mnuFileSave_Click()
    frmCommSetUp.Command2_Click
End Sub

Private Sub mnuFind_Click()

If frmFind.Visible Then
    frmFind.SetFocus
Else
    Load frmFind
    frmFind.Visible = True
End If

End Sub

Private Sub mnuFindNext_Click()
frmFind.btnFindNext_Click
End Sub

Private Sub mnuSelect_Click()
ActiveForm.RTB.SelStart = 0
ActiveForm.RTB.SelLength = Len(ActiveForm.RTB.Text)
SendKeys "^A", True
End Sub

Private Sub mnuUndo_Click()
ActiveForm.RTB.SelStart = (Len(ActiveForm.RTB.Text) + 1)
End Sub

Private Sub mnuWindowCascade_Click()
    Me.Arrange vbCascade
    WindowStyle = 0
End Sub
Private Sub mnuWindowTileHorizontal_Click()
    Me.Arrange vbTileHorizontal
    WindowStyle = 1
End Sub
Private Sub mnuWindowTileVertical_Click()
    Me.Arrange vbTileVertical
    WindowStyle = 2
End Sub
