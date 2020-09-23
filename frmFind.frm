VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton FindCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton btnFindNext 
      Caption         =   "Find Next"
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton btnFind 
      Caption         =   "Find Text"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox TextFind 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Text to Find"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private FindPos As Long
Dim RTBtext As String

Private Sub btnFind_Click()

FindPos = 0

RTBtext = frmMonitor.ActiveForm.RTB.Text
FindPos = InStr(1, UCase(RTBtext), UCase(TextFind.Text))
    'Set highlight of found text
    If FindPos <> 0 Then
        frmMonitor.ActiveForm.RTB.SetFocus
        frmMonitor.ActiveForm.RTB.SelStart = (FindPos - 1)
        frmMonitor.ActiveForm.RTB.SelLength = Len(TextFind.Text)
    End If
    
End Sub

Public Sub btnFindNext_Click()


RTBtext = frmMonitor.ActiveForm.RTB.Text
FindPos = InStr((FindPos + 1), UCase(RTBtext), UCase(TextFind.Text))
    'Set highlight of found text
    If FindPos <> 0 Then
        frmMonitor.ActiveForm.RTB.SetFocus
        frmMonitor.ActiveForm.RTB.SelStart = (FindPos - 1)
        frmMonitor.ActiveForm.RTB.SelLength = Len(TextFind.Text)
    End If

End Sub

Private Sub FindCancel_Click()
Unload Me
End Sub

Private Sub Form_Load()
Static FindText As String

If frmMonitor.ActiveForm.RTB.SelText <> "" Then
    TextFind.Text = frmMonitor.ActiveForm.RTB.SelText
Else
    If FindText <> "" Then
        TextFind.Text = FindText
    Else
        TextFind.Text = ""
    End If
End If

FindText = TextFind.Text

End Sub


