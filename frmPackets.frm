VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmPackets 
   BackColor       =   &H00808000&
   Caption         =   "Form1"
   ClientHeight    =   5265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7290
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPackets.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5265
   ScaleWidth      =   7290
   Visible         =   0   'False
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "HEX "
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Text Verbose"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Text Raw"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Value           =   -1  'True
      Width           =   1335
   End
   Begin RichTextLib.RichTextBox RTB 
      Height          =   3975
      Left            =   40
      TabIndex        =   0
      Top             =   1080
      Width           =   7200
      _ExtentX        =   12700
      _ExtentY        =   7011
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmPackets.frx":0442
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "Monitor Output View"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   780
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "frmPackets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ViewType  As Integer

Public Sub Form_Resize()
    ResizeFormFor Me
End Sub

Private Sub Option1_Click(Index As Integer)

'Set the View Type
'0 = Raw Text
'1 = Verbose Text
'2 = HEX
ViewType = Index
RTB.SetFocus

End Sub

Private Sub RTB_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
PopupMenu frmMonitor.mnuEdit
End If
End Sub
