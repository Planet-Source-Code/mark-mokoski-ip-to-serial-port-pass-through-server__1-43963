VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCommSetUp 
   BackColor       =   &H80000001&
   Caption         =   "IP to Serial Port"
   ClientHeight    =   5580
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6750
   Icon            =   "CommSetUp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   6750
   StartUpPosition =   3  'Windows Default
   WindowState     =   1  'Minimized
   Begin VB.CommandButton Command2 
      Caption         =   "Save Settings"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3720
      Picture         =   "CommSetUp.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   78
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Monitor Packets"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1680
      Picture         =   "CommSetUp.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   77
      Top             =   4680
      Width           =   1455
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   7858
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   -2147483647
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Link 0"
      TabPicture(0)   =   "CommSetUp.frx":114E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Winsock(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "MSComm(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "ComPortFrame(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Link 1"
      TabPicture(1)   =   "CommSetUp.frx":116A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Winsock(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "MSComm(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "ComPortFrame(1)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Link 2"
      TabPicture(2)   =   "CommSetUp.frx":1186
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Winsock(2)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "MSComm(2)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "ComPortFrame(2)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Link 3"
      TabPicture(3)   =   "CommSetUp.frx":11A2
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "ComPortFrame(3)"
      Tab(3).Control(1)=   "MSComm(3)"
      Tab(3).Control(2)=   "Winsock(3)"
      Tab(3).ControlCount=   3
      Begin VB.Frame ComPortFrame 
         Caption         =   "COM/IP Link 3 Control"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   3135
         Index           =   3
         Left            =   -74760
         TabIndex        =   58
         Top             =   600
         Width           =   6015
         Begin VB.CheckBox Check1 
            Caption         =   "Check1"
            Height          =   195
            Index           =   3
            Left            =   3120
            TabIndex        =   83
            Top             =   2760
            Width           =   255
         End
         Begin VB.ComboBox CommBits 
            Height          =   315
            Index           =   3
            Left            =   1320
            TabIndex        =   68
            Text            =   "8"
            Top             =   1200
            Width           =   1455
         End
         Begin VB.ComboBox CommStop 
            Height          =   315
            Index           =   3
            Left            =   1320
            TabIndex        =   67
            Text            =   "1"
            Top             =   1680
            Width           =   1455
         End
         Begin VB.ComboBox CommSpeed 
            Height          =   315
            Index           =   3
            Left            =   1320
            TabIndex        =   66
            Text            =   "9600"
            Top             =   720
            Width           =   1455
         End
         Begin VB.ComboBox CommFlow 
            Height          =   315
            Index           =   3
            Left            =   1320
            TabIndex        =   65
            Text            =   "XON/XOFF"
            Top             =   2640
            Width           =   1455
         End
         Begin VB.ComboBox CommParity 
            Height          =   315
            Index           =   3
            Left            =   1320
            TabIndex        =   64
            Text            =   "None"
            Top             =   2160
            Width           =   1455
         End
         Begin VB.ComboBox CommPort 
            Height          =   315
            Index           =   3
            Left            =   1320
            TabIndex        =   63
            Text            =   "COM 1"
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox LocalPortTCP 
            Height          =   285
            Index           =   3
            Left            =   4320
            TabIndex        =   62
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox LocalPortUDP 
            Height          =   285
            Index           =   3
            Left            =   4320
            TabIndex        =   61
            Top             =   720
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option1"
            Height          =   255
            Index           =   3
            Left            =   3120
            TabIndex        =   60
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Option2"
            Height          =   255
            Index           =   3
            Left            =   3120
            TabIndex        =   59
            Top             =   720
            Width           =   255
         End
         Begin VB.Label Label3 
            Caption         =   "Link Enabled"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   3
            Left            =   3360
            TabIndex        =   86
            Top             =   2760
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Data Bits"
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   76
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label CommSpeedLabel 
            Caption         =   "Speed"
            ForeColor       =   &H00C00000&
            Height          =   375
            Index           =   3
            Left            =   240
            TabIndex        =   75
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label CommFlowLabel 
            Caption         =   "Flow Control"
            ForeColor       =   &H00C00000&
            Height          =   375
            Index           =   3
            Left            =   240
            TabIndex        =   74
            Top             =   2640
            Width           =   1215
         End
         Begin VB.Label CommParityLabel 
            Caption         =   "Parity"
            ForeColor       =   &H00C00000&
            Height          =   375
            Index           =   3
            Left            =   240
            TabIndex        =   73
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label CommStopLabel 
            Caption         =   "Stop Bits"
            ForeColor       =   &H00C00000&
            Height          =   375
            Index           =   3
            Left            =   240
            TabIndex        =   72
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label CommPortLabel 
            Caption         =   "COM Port"
            ForeColor       =   &H00C00000&
            Height          =   375
            Index           =   3
            Left            =   240
            TabIndex        =   71
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label TCPlabel 
            Caption         =   "TCP Port"
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   3
            Left            =   3360
            TabIndex        =   70
            Top             =   240
            Width           =   855
         End
         Begin VB.Label PortLabel 
            Caption         =   "UDP Port"
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   3
            Left            =   3360
            TabIndex        =   69
            Top             =   720
            Width           =   855
         End
      End
      Begin VB.Frame ComPortFrame 
         Caption         =   "COM/IP Link 2 Control"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   3135
         Index           =   2
         Left            =   -74760
         TabIndex        =   39
         Top             =   600
         Width           =   6015
         Begin VB.CheckBox Check1 
            Caption         =   "Check1"
            Height          =   195
            Index           =   2
            Left            =   3120
            TabIndex        =   82
            Top             =   2760
            Width           =   255
         End
         Begin VB.ComboBox CommBits 
            Height          =   315
            Index           =   2
            Left            =   1320
            TabIndex        =   49
            Text            =   "8"
            Top             =   1200
            Width           =   1455
         End
         Begin VB.ComboBox CommStop 
            Height          =   315
            Index           =   2
            Left            =   1320
            TabIndex        =   48
            Text            =   "1"
            Top             =   1680
            Width           =   1455
         End
         Begin VB.ComboBox CommSpeed 
            Height          =   315
            Index           =   2
            Left            =   1320
            TabIndex        =   47
            Text            =   "9600"
            Top             =   720
            Width           =   1455
         End
         Begin VB.ComboBox CommFlow 
            Height          =   315
            Index           =   2
            Left            =   1320
            TabIndex        =   46
            Text            =   "XON/XOFF"
            Top             =   2640
            Width           =   1455
         End
         Begin VB.ComboBox CommParity 
            Height          =   315
            Index           =   2
            Left            =   1320
            TabIndex        =   45
            Text            =   "None"
            Top             =   2160
            Width           =   1455
         End
         Begin VB.ComboBox CommPort 
            Height          =   315
            Index           =   2
            Left            =   1320
            TabIndex        =   44
            Text            =   "COM 1"
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox LocalPortTCP 
            Height          =   285
            Index           =   2
            Left            =   4320
            TabIndex        =   43
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox LocalPortUDP 
            Height          =   285
            Index           =   2
            Left            =   4320
            TabIndex        =   42
            Top             =   720
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option1"
            Height          =   255
            Index           =   2
            Left            =   3120
            TabIndex        =   41
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Option2"
            Height          =   255
            Index           =   2
            Left            =   3120
            TabIndex        =   40
            Top             =   720
            Width           =   255
         End
         Begin VB.Label Label3 
            Caption         =   "Link Enabled"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   2
            Left            =   3360
            TabIndex        =   85
            Top             =   2760
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Data Bits"
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   57
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label CommSpeedLabel 
            Caption         =   "Speed"
            ForeColor       =   &H00C00000&
            Height          =   375
            Index           =   2
            Left            =   240
            TabIndex        =   56
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label CommFlowLabel 
            Caption         =   "Flow Control"
            ForeColor       =   &H00C00000&
            Height          =   375
            Index           =   2
            Left            =   240
            TabIndex        =   55
            Top             =   2640
            Width           =   1215
         End
         Begin VB.Label CommParityLabel 
            Caption         =   "Parity"
            ForeColor       =   &H00C00000&
            Height          =   375
            Index           =   2
            Left            =   240
            TabIndex        =   54
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label CommStopLabel 
            Caption         =   "Stop Bits"
            ForeColor       =   &H00C00000&
            Height          =   375
            Index           =   2
            Left            =   240
            TabIndex        =   53
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label CommPortLabel 
            Caption         =   "COM Port"
            ForeColor       =   &H00C00000&
            Height          =   375
            Index           =   2
            Left            =   240
            TabIndex        =   52
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label TCPlabel 
            Caption         =   "TCP Port"
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   2
            Left            =   3360
            TabIndex        =   51
            Top             =   240
            Width           =   855
         End
         Begin VB.Label PortLabel 
            Caption         =   "UDP Port"
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   2
            Left            =   3360
            TabIndex        =   50
            Top             =   720
            Width           =   855
         End
      End
      Begin VB.Frame ComPortFrame 
         Caption         =   "COM/IP Link 1 Control"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   3135
         Index           =   1
         Left            =   -74760
         TabIndex        =   20
         Top             =   600
         Width           =   6015
         Begin VB.CheckBox Check1 
            Caption         =   "Check1"
            Height          =   195
            Index           =   1
            Left            =   3120
            TabIndex        =   81
            Top             =   2760
            Width           =   255
         End
         Begin VB.ComboBox CommBits 
            Height          =   315
            Index           =   1
            Left            =   1320
            TabIndex        =   30
            Text            =   "8"
            Top             =   1200
            Width           =   1455
         End
         Begin VB.ComboBox CommStop 
            Height          =   315
            Index           =   1
            Left            =   1320
            TabIndex        =   29
            Text            =   "1"
            Top             =   1680
            Width           =   1455
         End
         Begin VB.ComboBox CommSpeed 
            Height          =   315
            Index           =   1
            Left            =   1320
            TabIndex        =   28
            Text            =   "9600"
            Top             =   720
            Width           =   1455
         End
         Begin VB.ComboBox CommFlow 
            Height          =   315
            Index           =   1
            Left            =   1320
            TabIndex        =   27
            Text            =   "XON/XOFF"
            Top             =   2640
            Width           =   1455
         End
         Begin VB.ComboBox CommParity 
            Height          =   315
            Index           =   1
            Left            =   1320
            TabIndex        =   26
            Text            =   "None"
            Top             =   2160
            Width           =   1455
         End
         Begin VB.ComboBox CommPort 
            Height          =   315
            Index           =   1
            Left            =   1320
            TabIndex        =   25
            Text            =   "COM 1"
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox LocalPortTCP 
            Height          =   285
            Index           =   1
            Left            =   4320
            TabIndex        =   24
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox LocalPortUDP 
            Height          =   285
            Index           =   1
            Left            =   4320
            TabIndex        =   23
            Top             =   720
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option1"
            Height          =   255
            Index           =   1
            Left            =   3120
            TabIndex        =   22
            Top             =   240
            Width           =   255
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Option2"
            Height          =   255
            Index           =   1
            Left            =   3120
            TabIndex        =   21
            Top             =   720
            Width           =   255
         End
         Begin VB.Label Label3 
            Caption         =   "Link Enabled"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   1
            Left            =   3360
            TabIndex        =   84
            Top             =   2760
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Data Bits"
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   38
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label CommSpeedLabel 
            Caption         =   "Speed"
            ForeColor       =   &H00C00000&
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   37
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label CommFlowLabel 
            Caption         =   "Flow Control"
            ForeColor       =   &H00C00000&
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   36
            Top             =   2640
            Width           =   1215
         End
         Begin VB.Label CommParityLabel 
            Caption         =   "Parity"
            ForeColor       =   &H00C00000&
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   35
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label CommStopLabel 
            Caption         =   "Stop Bits"
            ForeColor       =   &H00C00000&
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   34
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label CommPortLabel 
            Caption         =   "COM Port"
            ForeColor       =   &H00C00000&
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   33
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label TCPlabel 
            Caption         =   "TCP Port"
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   1
            Left            =   3360
            TabIndex        =   32
            Top             =   240
            Width           =   855
         End
         Begin VB.Label PortLabel 
            Caption         =   "UDP Port"
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   1
            Left            =   3360
            TabIndex        =   31
            Top             =   720
            Width           =   855
         End
      End
      Begin VB.Frame ComPortFrame 
         Caption         =   "COM/IP Link 0 Control"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   3135
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   6015
         Begin VB.CheckBox Check1 
            Caption         =   "Check1"
            Height          =   195
            Index           =   0
            Left            =   3120
            TabIndex        =   79
            Top             =   2760
            Width           =   255
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Option2"
            Height          =   255
            Index           =   0
            Left            =   3120
            TabIndex        =   19
            Top             =   720
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option1"
            Height          =   255
            Index           =   0
            Left            =   3120
            TabIndex        =   18
            Top             =   240
            Width           =   255
         End
         Begin VB.TextBox LocalPortUDP 
            Height          =   285
            Index           =   0
            Left            =   4320
            TabIndex        =   9
            Top             =   720
            Width           =   1455
         End
         Begin VB.TextBox LocalPortTCP 
            Height          =   285
            Index           =   0
            Left            =   4320
            TabIndex        =   8
            Top             =   240
            Width           =   1455
         End
         Begin VB.ComboBox CommPort 
            Height          =   315
            Index           =   0
            Left            =   1320
            TabIndex        =   7
            Text            =   "COM 1"
            Top             =   240
            Width           =   1455
         End
         Begin VB.ComboBox CommParity 
            Height          =   315
            Index           =   0
            Left            =   1320
            TabIndex        =   6
            Text            =   "None"
            Top             =   2160
            Width           =   1455
         End
         Begin VB.ComboBox CommFlow 
            Height          =   315
            Index           =   0
            Left            =   1320
            TabIndex        =   5
            Text            =   "XON/XOFF"
            Top             =   2640
            Width           =   1455
         End
         Begin VB.ComboBox CommSpeed 
            Height          =   315
            Index           =   0
            Left            =   1320
            TabIndex        =   4
            Text            =   "9600"
            Top             =   720
            Width           =   1455
         End
         Begin VB.ComboBox CommStop 
            Height          =   315
            Index           =   0
            Left            =   1320
            TabIndex        =   3
            Text            =   "1"
            Top             =   1680
            Width           =   1455
         End
         Begin VB.ComboBox CommBits 
            Height          =   315
            Index           =   0
            Left            =   1320
            TabIndex        =   2
            Text            =   "8"
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Label Label3 
            Caption         =   "Link Enabled"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   0
            Left            =   3360
            TabIndex        =   80
            Top             =   2760
            Width           =   1095
         End
         Begin VB.Label PortLabel 
            Caption         =   "UDP Port"
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   0
            Left            =   3360
            TabIndex        =   17
            Top             =   720
            Width           =   855
         End
         Begin VB.Label TCPlabel 
            Caption         =   "TCP Port"
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   0
            Left            =   3360
            TabIndex        =   16
            Top             =   240
            Width           =   855
         End
         Begin VB.Label CommPortLabel 
            Caption         =   "COM Port"
            ForeColor       =   &H00C00000&
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   15
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label CommStopLabel 
            Caption         =   "Stop Bits"
            ForeColor       =   &H00C00000&
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   14
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label CommParityLabel 
            Caption         =   "Parity"
            ForeColor       =   &H00C00000&
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   13
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label CommFlowLabel 
            Caption         =   "Flow Control"
            ForeColor       =   &H00C00000&
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   12
            Top             =   2640
            Width           =   1215
         End
         Begin VB.Label CommSpeedLabel 
            Caption         =   "Speed"
            ForeColor       =   &H00C00000&
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   11
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Data Bits"
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   10
            Top             =   1200
            Width           =   975
         End
      End
      Begin MSCommLib.MSComm MSComm 
         Index           =   0
         Left            =   120
         Top             =   3720
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
         RThreshold      =   1
      End
      Begin MSWinsockLib.Winsock Winsock 
         Index           =   0
         Left            =   720
         Top             =   3840
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSCommLib.MSComm MSComm 
         Index           =   1
         Left            =   -74880
         Top             =   3720
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
         RThreshold      =   1
      End
      Begin MSWinsockLib.Winsock Winsock 
         Index           =   1
         Left            =   -74280
         Top             =   3840
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSCommLib.MSComm MSComm 
         Index           =   2
         Left            =   -74880
         Top             =   3720
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
         RThreshold      =   1
      End
      Begin MSWinsockLib.Winsock Winsock 
         Index           =   2
         Left            =   -74280
         Top             =   3840
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSCommLib.MSComm MSComm 
         Index           =   3
         Left            =   -74880
         Top             =   3720
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
         RThreshold      =   1
      End
      Begin MSWinsockLib.Winsock Winsock 
         Index           =   3
         Left            =   -74280
         Top             =   3840
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save Settings"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileMin 
         Caption         =   "Minimize to System Tray"
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Shutdown IP to Serial Server"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuConLog 
         Caption         =   "&Connection Log"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "C&lear Log"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuLink 
         Caption         =   "Link 0"
         Checked         =   -1  'True
         Index           =   0
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuLink 
         Caption         =   "Link 1"
         Index           =   1
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuLink 
         Caption         =   "Link 2"
         Index           =   2
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuLink 
         Caption         =   "Link 3"
         Index           =   3
         Shortcut        =   {F4}
      End
      Begin VB.Menu windowsep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMonitor 
         Caption         =   "&Monitor Connections"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuLog 
         Caption         =   "&Connection Log"
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
   Begin VB.Menu mnuRestore 
      Caption         =   "Restore"
      Visible         =   0   'False
      Begin VB.Menu mnuResMonitor 
         Caption         =   "&Monitor Connections"
      End
      Begin VB.Menu mnuResLog 
         Caption         =   "&Connection Log"
      End
      Begin VB.Menu mnuResSetUp 
         Caption         =   "&Set-Up Connections"
      End
      Begin VB.Menu mnuResAbout 
         Caption         =   "&About IP to Serial Server"
      End
      Begin VB.Menu mnuRseSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
      Begin VB.Menu mnuResSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuResClose 
         Caption         =   "Shut Down IP to Serial Server"
      End
   End
End
Attribute VB_Name = "frmCommSetUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim frmViewPackets(4)   As Form
Dim x           As Integer
Dim COMdata     As String
Dim COMtext     As String
Dim CloseText   As String
Dim OpenText    As String
Dim IPtext      As String
Dim Parity      As String

Const NAME_COLUMN = 0
Const TYPE_COLUMN = 1
Const SIZE_COLUMN = 2
Const DATE_COLUMN = 3


Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hWnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)

Public Sub mnuClear_Click()
On Error Resume Next
If frmViewLog.Visible = True Then Unload frmViewLog
Kill App.Path & "\" & "Connection.log"
End Sub



Private Sub mnuConLog_Click()

If frmViewLog.Visible = True Then
    frmViewLog.SetFocus
        If frmViewLog.WindowState = vbMinimized Then
            frmViewLog.WindowState = vbNormal
        End If
Else
    frmViewLog.Visible = True
End If

End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuFileMin_Click()
frmCommSetUp.WindowState = vbMinimized
End Sub

Private Sub mnuFileSave_Click()
Command2_Click
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show
    frmAbout.SetFocus
End Sub


Private Sub Check1_Click(Index As Integer)
'Enable Tab settings
x = Index

If Check1(x).Value = Checked Then
    PortEnabled(x) = True
Else
    PortEnabled(x) = False
End If

End Sub

Private Sub Command1_Click()
'Main Monitor Form
'Check to see if from already loaded
If frmMonitor.Visible = True Then
    frmMonitor.SetFocus
        If frmMonitor.WindowState = vbMinimized Then
            frmMonitor.WindowState = vbNormal
        End If
Else
    'Create Packet Monitor Forms
    For x = 0 To 3
            Set frmViewPackets(x) = New frmPackets
            Load frmViewPackets(x)
            frmViewPackets(x).Visible = False
    Next x
    
    'Make acitve ones visible and set caption
    For x = 0 To 3
        If PortEnabled(x) = True Then
            frmViewPackets(x).Visible = True
            If Winsock(x).Protocol = sckTCPProtocol Then
                frmViewPackets(x).Caption = MonitorCaption(x)
            Else
                frmViewPackets(x).Caption = Protocol(x) & " Port " & Str(Winsock(x).LocalPort) & " to " & CommPort(x)
            End If
            frmViewPackets(x).RTB.Locked = True
        End If
    Next x
End If

frmMonitor.Visible = True
frmCommSetUp.mnuMonitor.Checked = True

'Triger a resize to set up cascade properly
frmMonitor.MDIForm_Resize
End Sub

Public Sub Command2_Click()

'Save settings to registry

For x = 0 To 3

LocalComPort(x) = CommPort(x).Text
SaveSetting "IPtoCOM", "Port" & Str(x), "ComPort", CommPort(x).Text
LocalComSpeed(x) = CommSpeed(x).Text
SaveSetting "IPtoCOM", "Port" & Str(x), "ComSpeed", CommSpeed(x).Text
LocalCombits(x) = CommBits(x).Text
SaveSetting "IPtoCOM", "Port" & Str(x), "Combits", CommBits(x).Text
LocalComStop(x) = CommStop(x).Text
SaveSetting "IPtoCOM", "Port" & Str(x), "ComStop", CommStop(x).Text
LocalComParity(x) = CommParity(x).Text
SaveSetting "IPtoCOM", "Port" & Str(x), "ComParity", CommParity(x).Text
LocalComFlow(x) = CommFlow(x).Text
SaveSetting "IPtoCOM", "Port" & Str(x), "ComFlow", CommFlow(x).Text
TCPport(x) = LocalPortTCP(x).Text
SaveSetting "IPtoCOM", "Port" & Str(x), "TCPport", LocalPortTCP(x).Text
UDPport(x) = LocalPortUDP(x).Text
SaveSetting "IPtoCOM", "Port" & Str(x), "UDPport", LocalPortUDP(x).Text
SaveSetting "IPtoCOM", "Port" & Str(x), "Protocol", Protocol(x)
SaveSetting "IPtoCOM", "Port" & Str(x), "PortEnabled", PortEnabled(x)

StartComIP (x)


Next x

'Load and show additional frmViewPackets forms if enabled and frmMonitor
'is currently visible
If frmMonitor.Visible = True Then
    For x = 0 To 3
        If PortEnabled(x) = True Then
            frmViewPackets(x).Visible = True
                If Winsock(x).Protocol = sckTCPProtocol Then
                    frmViewPackets(x).Caption = Protocol(x) & " Port " & Str(Winsock(x).LocalPort) & " to " & CommPort(x) & " Waiting for Connection"
                Else
                    frmViewPackets(x).Caption = Protocol(x) & " Port " & Str(Winsock(x).LocalPort) & " to " & CommPort(x)
                End If
            frmViewPackets(x).RTB.Locked = True
       Else
            frmViewPackets(x).Visible = False
       End If
    Next x
End If
frmMonitor.MDIForm_Resize
End Sub

Private Sub Form_Load()

Dim hSysMenu As Long
hSysMenu = GetSystemMenu(hWnd, False)
RemoveMenu hSysMenu, SC_CLOSE, MF_BYCOMMAND

For x = 0 To 3

    Set frmViewPackets(x) = Nothing
    
    'Set List box selection values
    'Comm ports
    CommPort(x).AddItem "COM 1", 0
    CommPort(x).AddItem "COM 2", 1
    CommPort(x).AddItem "COM 3", 2
    CommPort(x).AddItem "COM 4", 3
    CommPort(x).AddItem "COM 5", 4
    CommPort(x).AddItem "COM 6", 5
    CommPort(x).AddItem "COM 7", 6
    CommPort(x).AddItem "COM 8", 7
    
    'Comm port speed
    CommSpeed(x).AddItem "300", 0
    CommSpeed(x).AddItem "1200", 1
    CommSpeed(x).AddItem "2400", 2
    CommSpeed(x).AddItem "4800", 3
    CommSpeed(x).AddItem "9600", 4
    CommSpeed(x).AddItem "19200", 5
    CommSpeed(x).AddItem "38400", 6
    CommSpeed(x).AddItem "57600", 7
    CommSpeed(x).AddItem "115200", 8
    
    'Comm Databits
    CommBits(x).AddItem "8", 0
    CommBits(x).AddItem "7", 1
    
    'CommStop bits
    CommStop(x).AddItem "1", 0
    CommStop(x).AddItem "2", 1
    
    'Comm Parity bits
    CommParity(x).AddItem "None", 0
    CommParity(x).AddItem "Even", 1
    CommParity(x).AddItem "Odd", 2
    CommParity(x).AddItem "Mark", 3
    CommParity(x).AddItem "Space", 4
    
    'Comm Flow control
    CommFlow(x).AddItem "None", 0
    CommFlow(x).AddItem "XON/XOFF", 1
    CommFlow(x).AddItem "RTS/CTS", 2
    
    'Set current values
    CommPort(x).Text = LocalComPort(x)
    CommSpeed(x).Text = LocalComSpeed(x)
    CommBits(x).Text = LocalCombits(x)
    CommStop(x).Text = LocalComStop(x)
    CommParity(x).Text = LocalComParity(x)
    CommFlow(x).Text = LocalComFlow(x)
    LocalPortTCP(x).Text = TCPport(x)
    LocalPortUDP(x).Text = UDPport(x)
    
    'IP/COM tab enabled
    If PortEnabled(x) = True Then
        Check1(x).Value = Checked
    Else
        Check1(x).Value = Unchecked
    End If
    
    'Protocol type
    If Protocol(x) = "TCP" Then
        'TCP Protocol
        Option1(x).Value = True
    Else
        'Else UDP Protocol
        Option2(x).Value = True
    End If
    
    StartComIP (x)
    
Next x
    
    Command1.Picture = LoadResPicture(101, 1)
    Command2.Picture = LoadResPicture(102, 1)
End Sub

Private Sub Form_Resize()

If Me.WindowState = vbMinimized Then
    Call SystrayOn(frmCommSetUp, "IP to Serial Port Server")
    frmCommSetUp.Hide
End If
End Sub

Private Sub Form_Terminate()
'Get rid of icon in systray
Call SystrayOff(frmCommSetUp)
Unload frmMonitor
Unload frmViewLog
Unload frmAbout
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Get rid of icon in systray
Call SystrayOff(frmCommSetUp)
Unload frmMonitor
Unload frmViewLog
Unload frmAbout
End Sub

Private Sub mnuLink_Click(Index As Integer)
mnuLink.Item(Index).Checked = True
frmCommSetUp.SSTab1.Tab = Index
    For x = 0 To 3
        If x <> Index Then
            mnuLink.Item(x).Checked = False
        End If
    Next x
End Sub

Private Sub mnuLog_Click()
mnuConLog_Click
End Sub

Private Sub mnuMonitor_Click()
Command1_Click
End Sub

Private Sub mnuResAbout_Click()
mnuHelpAbout_Click
End Sub

Private Sub mnuResClose_Click()
Unload Me
End Sub

Private Sub mnuResLog_Click()
mnuConLog_Click
End Sub

Private Sub mnuResMonitor_Click()
Command1_Click
End Sub

Private Sub mnuResSetUp_Click()
'    Call SystrayOff(frmCommSetUp)
    Call SetForegroundWindow(Me.hWnd)
    frmCommSetUp.WindowState = vbNormal
    frmCommSetUp.Show
    frmCommSetUp.SetFocus

End Sub

Private Sub MSComm_OnComm(Index As Integer)
'Get input from COM port
x = Index

'Check Winsock states and protocol

'UDP protocol, RCV data from client only, exit sub
If Winsock(x).Protocol = sckUDPProtocol Then Exit Sub
'Test TCP socket for open
If Winsock(x).State = sckConnected Then
    If MSComm(x).InBufferCount > 0 Then
            COMdata = MSComm(x).Input
            Winsock(x).SendData COMdata
            'Write COM data to RTB if visible
        If frmMonitor.Visible = True Then Call ViewPackets(COMdata, "COM", x)
    End If
End If

End Sub

Private Sub Option1_Click(Index As Integer)
'Select Protocol to TCP
x = Index

If Option1(x) = True Then
    LocalPortTCP(x).Enabled = True
    LocalPortTCP(x).BackColor = vbWhite     '&H80000005 white
    LocalPortUDP(x).Enabled = False
    LocalPortUDP(x).BackColor = &H80000004  'grey
    Protocol(x) = "TCP"
End If

End Sub

Private Sub Option2_Click(Index As Integer)
'Select Protocol to UDP
x = Index

If Option2(x) = True Then
    LocalPortUDP(x).Enabled = True
    LocalPortUDP(x).BackColor = vbWhite     '&H80000005 white
    LocalPortTCP(x).Enabled = False
    LocalPortTCP(x).BackColor = &H80000004  'grey
    Protocol(x) = "UDP"
End If

End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
Dim Index As Integer

Index = SSTab1.Tab
mnuLink.Item(Index).Checked = True
    For x = 0 To 3
        If x <> Index Then
            mnuLink.Item(x).Checked = False
        End If
    Next x

End Sub


Private Sub Winsock_Close(Index As Integer)
Dim ConFile As Integer
Dim ConText As String
x = Index

CloseText = "Connection closed by client " & vbCrLf
MSComm(x).RTSEnable = False
MSComm(x).DTREnable = False
MSComm(x).PortOpen = False
    'Write disconect info to log file
    ConFile = FreeFile
    Open App.Path & "\" & "Connection.log" For Append As #ConFile
    ConText = Date & " at " & Format(Time, "hh:mm:ss") & "  - Host " & Winsock(x).RemoteHostIP & " disconnected from port" & Str(Winsock(x).LocalPort) & " to " & CommPort(x)
    Print #ConFile, ConText
    Close #ConFile
    'If log View visible, update the window
    If frmViewLog.Visible Then
        frmViewLog.mnuRefresh_Click
    End If
    
Winsock(x).Close
MonitorCaption(x) = Protocol(x) & " Port " & Str(Winsock(x).LocalPort) & " to " & CommPort(x) & " Waiting for Connection"
BalloonText = Protocol(x) & " Port " & Str(Winsock(x).LocalPort) & " to " & CommPort(x) & " Connection Closed"
Call PopupBalloon(frmCommSetUp, BalloonText, "IP to COM Connection Closed")

If frmMonitor.Visible = True Then
    If frmViewPackets(x).Visible = True Then
        frmViewPackets(x).RTB.SelColor = vbBlack
        frmViewPackets(x).RTB.SelText = CloseText
            If Winsock(x).Protocol = sckTCPProtocol Then
                frmViewPackets(x).Caption = MonitorCaption(x)
            End If
        frmViewPackets(x).RTB.SelStart = Len(frmViewPackets(x).RTB.Text) + 1
    End If
End If
Winsock(x).Listen

End Sub

Private Sub Winsock_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Dim ConFile As Integer
Dim ConText As String
x = Index

' Check if the control's State is closed. If not,
' close the connection before accepting the new
' connection.
If Winsock(x).State <> sckClosed Then Winsock(x).Close
' Accept the request with the requestID parameter.
On Error GoTo CommERR
MSComm(x).PortOpen = True
MSComm(x).DTREnable = True
MSComm(x).RTSEnable = True

'Open port and accept connection
Winsock(x).Accept requestID
OpenText = "Connection opened by client " & vbCrLf
    'Write connection info to log file
    ConFile = FreeFile
    Open App.Path & "\" & "Connection.log" For Append As #ConFile
    ConText = Date & " at " & Format(Time, "hh:mm:ss") & "  - Host " & Winsock(x).RemoteHostIP & " connected on port" & Str(Winsock(x).LocalPort) & " to " & CommPort(x)
    Print #ConFile, ConText
    Close #ConFile
    'If log View visible, update the window
    If frmViewLog.Visible Then
        frmViewLog.mnuRefresh_Click
    End If
MonitorCaption(x) = Protocol(x) & " Port " & Str(Winsock(x).LocalPort) & " to " & CommPort(x) & "  Connected with  " & Winsock(x).RemoteHostIP & " on port " & Str(Winsock(x).RemotePort)
BalloonText = MonitorCaption(x)
Call PopupBalloon(frmCommSetUp, BalloonText, "IP to COM Connection Opened")

If frmMonitor.Visible = True Then
    If frmViewPackets(x).Visible = True Then
        frmViewPackets(x).RTB.SelColor = vbBlack
        frmViewPackets(x).RTB.SelText = OpenText
        frmViewPackets(x).Caption = MonitorCaption(x)
        frmViewPackets(x).RTB.SelStart = Len(frmViewPackets(x).RTB.Text) + 1
    End If
End If
Exit Sub

CommERR:
'Put any error code here
'Most likely error is COM port used by another App
'So, just close the current socket and relisten
Winsock(x).Close
Winsock(x).Listen

End Sub

Private Sub Winsock_DataArrival(Index As Integer, ByVal bytesTotal As Long)

x = Index

Winsock(x).GetData IPtext

'Write incomming data to RTB
If frmMonitor.Visible = True Then
    If frmViewPackets(x).Visible = True Then Call ViewPackets(IPtext, "IP", x)
End If
'Send data to RS232 Port

        'Set up COM Port
        If MSComm(x).PortOpen = False Then
            MSComm(x).PortOpen = True 'Just to make sure !!

        End If
        
        MSComm(x).Output = IPtext
        'Wait until all IPtext is sent to COM port
        Do Until MSComm(x).OutBufferCount = 0
          DoEvents
        Loop
        
        Exit Sub
        'If Com Port error, display message box warning
ComPortErr:
        MsgBox "ERROR Opening Commuications Port" & vbCrLf & _
                "Port may be in use by another application or" & vbCrLf & _
                "is not a installed COM port." & vbCrLf & _
                "Check Properties for selected COM port", vbExclamation, "Communications Port ERROR"

End Sub

Private Sub StartComIP(Index As Integer)

x = Index
'Close exisiting Winsock and COM

Winsock(x).Close

If MSComm(x).PortOpen = True Then MSComm(x).PortOpen = False

'Winsock settings

If Check1(x).Value = Checked Then
    On Error GoTo WinsockError
    Select Case Protocol(x)
        
        Case "TCP"
            Winsock(x).Protocol = sckTCPProtocol
            Winsock(x).LocalPort = Val(TCPport(x))
            Winsock(x).Bind Val(TCPport(x))
            Winsock(x).Listen
        
        Case Else
            Winsock(x).Protocol = sckUDPProtocol
            Winsock(x).LocalPort = Val(UDPport(x))
            Winsock(x).Bind Val(UDPport(x))
            
    End Select

'COM settings

        On Error GoTo ComPortErr
        MSComm(x).CommPort = Val(Mid(LocalComPort(x), 5, 1))
        Parity = Mid(LocalComParity(x), 1, 1)
        MSComm(x).RThreshold = 1
        
        Select Case LocalComFlow(x)
            Case "None"
                MSComm(x).Handshaking = 0
            Case "XON/XOFF"
                MSComm(x).Handshaking = 1
            Case "RTS/CTS"
                MSComm(x).Handshaking = 2
        End Select
        'If error opening port, display message box warning
        MSComm(x).Settings = LocalComSpeed(x) & "," & Parity & "," & LocalCombits(x) & LocalComStop(x)
        
        'If protocol is UDP, open comport
        If Winsock(x).Protocol = sckUDPProtocol Then
            MSComm(x).PortOpen = True
        End If
End If

Exit Sub

        'If Com Port error, display message box warning
ComPortErr:
        MsgBox "ERROR Opening Commuications Port" & vbCrLf & _
                "Port may be in use by another application or" & vbCrLf & _
                "is not a installed COM port." & vbCrLf & _
                "Check Properties for selected COM port", vbExclamation, "Communications Port ERROR"
Exit Sub

        'If Winsock error, display message box warning
WinsockError:
        MsgBox "ERROR Opening TCP/IP Port" & vbCrLf & _
                "Port may be in use by another application or" & vbCrLf & _
                "Check Properties for selected IP port", vbExclamation, "TCP/IP Port ERROR"

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Static lngMsg As Long
    Dim blnflag As Boolean, lngResult As Long

    lngMsg = x / Screen.TwipsPerPixelX
    If blnflag = False Then
        blnflag = True
        
        Debug.Print x
        
        Select Case lngMsg
        Case WM_RBUTTONCLK      'to popup on left-click
            Call SetForegroundWindow(Me.hWnd)
            PopupMenu mnuRestore
                        
        Case WM_LBUTTONDBLCLK   'open on left-dblclick
            'Call SystrayOff(frmCommSetUp)
            Call SetForegroundWindow(Me.hWnd)
            frmCommSetUp.WindowState = vbNormal
            frmCommSetUp.Show
'            FormOnTop frmCommSetUp
            frmCommSetUp.SetFocus
            
        End Select
        blnflag = False
    End If
End Sub

Public Sub ViewPackets(COMdata As String, Source As String, Index As Integer)
'Decode and format output to Rich Text Box in format selected
Dim TEMPtext As String
Dim TEMPchr  As Integer
Dim y As Integer

frmViewPackets(Index).RTB.Locked = False

Select Case frmViewPackets(Index).ViewType

    Case 0  'Raw Text output
        If Source = "IP" Then
            COMtext = "IP: -> COM: " & COMdata & vbCrLf
            frmViewPackets(Index).RTB.SelColor = vbRed
        Else
            COMtext = "COM: -> IP: " & COMdata & vbCrLf
            frmViewPackets(Index).RTB.SelColor = vbBlue
        End If
    
    Case 1  'Verbose Text output
        TEMPtext = ""
        For y = 1 To Len(COMdata)
            TEMPchr = Asc(Mid(COMdata, y, 1))
                Select Case TEMPchr
                    'ASCII Codes
                    Case 0      'Null
                        TEMPtext = TEMPtext & "<NUL> "
                    Case 1      'Start Of Header
                        TEMPtext = TEMPtext & "<SOH> "
                    Case 2      'Start Of Text
                        TEMPtext = TEMPtext & "<STX> "
                    Case 3      'End Of Text
                        TEMPtext = TEMPtext & "<ETX> "
                    Case 4      'End Of Transmission
                        TEMPtext = TEMPtext & "<EOT> "
                    Case 5      'Enquiry
                        TEMPtext = TEMPtext & "<ENQ> "
                    Case 6      'Acknowledge
                        TEMPtext = TEMPtext & "<ACK> "
                    Case 7      'Bell
                        TEMPtext = TEMPtext & "<BEL>"
                    Case 8      'Backspace
                        TEMPtext = TEMPtext & "<BS>"
                    Case 9      'Horizontal Tab
                        TEMPtext = TEMPtext & "<TAB>"
                    Case 10     'Line Feed
                        TEMPtext = TEMPtext & "<LF>"
                    Case 11     'Vertical Tab
                        TEMPtext = TEMPtext & "<VT>"
                    Case 12     'Form Feed
                        TEMPtext = TEMPtext & "<FF>"
                    Case 13     'Carriage Return
                        TEMPtext = TEMPtext & "<CR>"
                    Case 14     'Shift Out
                        TEMPtext = TEMPtext & "<SO>"
                    Case 15     'Shift In
                        TEMPtext = TEMPtext & "<SI>"
                    Case 16     'Data Link Escape
                        TEMPtext = TEMPtext & "<DLE>"
                    Case 17     'Device Control 1 (XON)
                        TEMPtext = TEMPtext & "<XON>"
                    Case 18     'Device Control 2
                        TEMPtext = TEMPtext & "<DC2>"
                    Case 19     'Device Control 3(XOFF)
                        TEMPtext = TEMPtext & "<XOF>"
                    Case 20     'Device Control  4
                        TEMPtext = TEMPtext & "<DC4>"
                    Case 21     'Negative Acknowledge
                        TEMPtext = TEMPtext & "<NAK>"
                    Case 22     'Synchronous Idle
                        TEMPtext = TEMPtext & "<SYN>"
                    Case 23     'End of Transmission Block
                        TEMPtext = TEMPtext & "<ETB>"
                    Case 24     'Cancel
                        TEMPtext = TEMPtext & "<CAN>"
                    Case 25     'End of Medium
                        TEMPtext = TEMPtext & "<EM>"
                    Case 26     'Substitue
                        TEMPtext = TEMPtext & "<SUB>"
                    Case 27     'Escape
                        TEMPtext = TEMPtext & "<ESC>"
                    Case 28     'File Separator
                        TEMPtext = TEMPtext & "<FS>"
                    Case 29     'Group Separator
                        TEMPtext = TEMPtext & "<GS>"
                    Case 30     'Record Separator
                        TEMPtext = TEMPtext & "<RS>"
                    Case 31     'Unit Separator
                        TEMPtext = TEMPtext & "<US>"
                    Case 32     'Space
                        TEMPtext = TEMPtext & "<SP>"
                    Case 127    'Delete
                        TEMPtext = TEMPtext & "<DEL>"
                    Case Is > 127   'Extended ASCII (8 bit) codes
                        TEMPtext = TEMPtext & "<" & Hex(Asc(Mid(COMdata, y, 1))) & "h>"
                    Case Else   'Text, ASCII codes 33-126
                        TEMPtext = TEMPtext & Mid(COMdata, y, 1)
                End Select
        Next y
        
        If Source = "IP" Then
            COMtext = "IP: -> COM: " & TEMPtext & vbCrLf
            frmViewPackets(Index).RTB.SelColor = vbRed
        Else
            COMtext = "COM: -> IP: " & TEMPtext & vbCrLf
            frmViewPackets(Index).RTB.SelColor = vbBlue
        End If
    
    Case 2  'HEX output
        TEMPtext = ""
        For y = 1 To Len(COMdata)
        TEMPchr = Asc(Mid(COMdata, y, 1))
        If TEMPchr < 17 Then
            TEMPtext = TEMPtext & "0" & Hex(Asc(Mid(COMdata, y, 1))) & " "
        Else
            TEMPtext = TEMPtext & Hex(Asc(Mid(COMdata, y, 1))) & " "
        End If
        Next y
        
        If Source = "IP" Then
            COMtext = "IP: -> COM: " & TEMPtext & vbCrLf
            frmViewPackets(Index).RTB.SelColor = vbRed
        Else
            COMtext = "COM: -> IP: " & TEMPtext & vbCrLf
            frmViewPackets(Index).RTB.SelColor = vbBlue
        End If
        
End Select

'Set start of next line to end of current buffer
frmViewPackets(Index).RTB.SelStart = Len(frmViewPackets(Index).RTB.Text) + 1
'Write the text
frmViewPackets(Index).RTB.SelText = COMtext
'Set start of next line to end of current buffer
frmViewPackets(Index).RTB.SelStart = Len(frmViewPackets(Index).RTB.Text) + 1
            
frmViewPackets(Index).RTB.Locked = True

End Sub

