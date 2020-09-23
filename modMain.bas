Attribute VB_Name = "modMain"
'***************************************************************************
'   IP to Serial port Pass-Through Server Project
'   23-FEB-2003
'
'   Copyright © 2003 Mark Mokoski
'   markm@cmtelephone.com
'   www.cmtelephone.com
'
'   This project takes TCP and UDP messages from other applications
'   and routes them to the proper serial port to control Serial port
'   devices (ex. Telephone Sytems, Router Console Port)or other device
'   on a PC Serial Port.
'
'   Control can be one way (App --> COM Port) via UDP packets, or
'   Control can be two way (App <-> COM Port) via a TCP connection.
'
'   Version Info
'
'   0.1.X   23-FEB-2003
'   First development.  UDP packets working, packet monitor form
'
'   0.2.X   02-MAR-2003
'   TCP connections working (as server)
'
'   0.3.X   05-MAR-2003
'   Cleaned up some small issues. (changed some form references)
'   Added resize code to monitor form and MDI children.
'
'   0.4.X   07-MAR-2003
'   Cleaned up user interface, colors, icons etc.
'   Set some error trapping on COM and IP conections.
'   Refuse addional connections (TCP) once a COM port is in use.
'   Close COM port when no TCP connection present (For local use of port)
'
'   1.0.X   08-MAR-2003
'   Test release.  Works OK via telnet to control Kenwood TS870 radio and
'   MFJ TNC-2 control.
'
'   1.1.X   09-MAR-2003
'   Added App minimized to systray on startup
'
'   1.2.X   13-MAR-2003
'   Cleaned up some Resize isuues in frmMonitor dealing with the
'   child forms (frmPackets).  Added Raw Text, Verbose Text and HEX
'   monitor modes on frmPackets. Some Dead code clean up and renamed
'   Tabs from "Ports" to "Links" ( I think that's more in line with
'   the projects operation!). Added "Connected to" in the frmPackets
'   caption showing remote IP and remote port of connection.
'   Expanded ASCII control character symbols for verbose text mode.
'
'   1.3.X   15-MAR-2003
'   Added menu bar to frmCommSetUP.  Added connection log and view form
'   (frmViewLog).  Added About form (frmAbout). Added Pop-up menu when
'   App is in SysTray (Right click to Popup)
'
'   1.4.X   17-MAR-2003
'   Added some clipboard features (Cut, Select All, Undo Selection, Find)
'   on Connection Log and Monitor forms
'
'   1.5.x   28-JAN-2004
'   Fixed a problem with the "X" to close app while in connections setup
'   Fixed problem with restoring settings on link's 1 to 3
'
'****************************************************************************
Option Explicit

Public LocalComPort(4)      As String
Public LocalCombits(4)      As String
Public LocalComSpeed(4)     As String
Public LocalComStop(4)      As String
Public LocalComParity(4)    As String
Public LocalComFlow(4)      As String
Public UDPport(4)           As String
Public TCPport(4)           As String
Public Protocol(4)          As String
Public PortEnabled(4)       As Boolean
Public frmViewPackets(4)    As Form
Public MonitorCaption(4)    As String
Public BalloonText          As String

Public Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long

Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
    Public Const SC_CLOSE = &HF060&
    Public Const MF_BYCOMMAND = &H0&


Dim x                       As Integer
Sub Main()
    ' * Test to see if App is already running
    ' * If App is running, terminate copy
    If App.PrevInstance Then
        MsgBox "IP to Comm Port Control application is already running" & vbCrLf & vbCrLf & _
        "Only one instance (copy) of program this can be running" & vbCrLf & _
        "for proper operation", vbInformation, "Application ERROR"
        End
    Else
        '  MsgBox "This is the first instance of your application"
    End If
    'Load settings saved in Registry
    'LocalPort Comm Port settings
    For x = 0 To 3
        LocalComPort(x) = GetSetting("IPtoCOM", "Port" & Str(x), "ComPort", "COM 1")
        LocalCombits(x) = GetSetting("IPtoCOM", "Port" & Str(x), "ComBits", "8")
        LocalComSpeed(x) = GetSetting("IPtoCOM", "Port" & Str(x), "ComSpeed", "9600")
        LocalComStop(x) = GetSetting("IPtoCOM", "Port" & Str(x), "ComStop", "1")
        LocalComParity(x) = GetSetting("IPtoCOM", "Port" & Str(x), "ComParity", "None")
        LocalComFlow(x) = GetSetting("IPtoCOM", "Port" & Str(x), "ComFlow", "XON/XOFF")
        UDPport(x) = GetSetting("IPtoCOM", "Port" & Str(x), "UDPport", "8003")
        TCPport(x) = GetSetting("IPtoCOM", "Port" & Str(x), "TCPport", "8001")
        Protocol(x) = GetSetting("IPtoCOM", "Port" & Str(x), "Protocol", "TCP")
        PortEnabled(x) = GetSetting("IPtoCOM", "Port" & Str(x), "PortEnabled", False)
    Next x
    
    For x = 0 To 3
    
        MonitorCaption(x) = Protocol(x) & " Port " & Str(frmCommSetUp.Winsock(x).LocalPort) & " to " & frmCommSetUp.CommPort(x) & " Waiting for Connection"

    Next x
    'Make setup form visable
    Load frmCommSetUp
    frmCommSetUp.Visible = True
End Sub
