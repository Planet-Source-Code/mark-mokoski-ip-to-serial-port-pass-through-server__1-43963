Attribute VB_Name = "modSystray"
'Systray Module
Option Explicit

Public blnClick                  As Boolean
Public vbTray                    As NOTIFYICONDATA

Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    
Public Type NOTIFYICONDATA
   cbSize                        As Long
   hWnd                          As Long
   uID                           As Long
   uFlags                        As Long
   uCallbackMessage              As Long
   hIcon                         As Long
   szTip                         As String * 128
   dwState                       As Long
   dwStateMask                   As Long
   szInfo                        As String * 256
   uTimeout                      As Long
   szInfoTitle                   As String * 64
   dwInfoFlags                   As Long
End Type

Public Const SWP_NOMOVE          As Long = &H2
Public Const SWP_NOSIZE          As Long = &H1
Public Const FLAGS               As Long = SWP_NOMOVE Or SWP_NOSIZE
Public Const WM_RBUTTONUP        As Long = &H205
Public Const WM_RBUTTONCLK       As Long = &H204
Public Const WM_LBUTTONCLK       As Long = &H202
Public Const WM_LBUTTONDBLCLK    As Long = &H203
Public Const WM_MOUSEMOVE        As Long = &H200
Public Const NIM_ADD             As Long = &H0
Public Const NIM_DELETE          As Long = &H2
Public Const NIF_ICON            As Long = &H2
Public Const NIF_MESSAGE         As Long = &H1
Public Const NIM_MODIFY          As Long = &H1
Public Const NIF_TIP             As Long = &H4
Public Const NIF_INFO            As Long = &H10
Public Const NIS_HIDDEN          As Long = &H1
Public Const NIS_SHAREDICON      As Long = &H2
Public Const NIIF_NONE           As Long = &H0
Public Const NIIF_WARNING        As Long = &H2
Public Const NIIF_ERROR          As Long = &H3
Public Const NIIF_INFO           As Long = &H1
Public Const NIIF_GUID           As Long = &H4
Public Const HWND_NOTOPMOST      As Long = -2
Public Const HWND_TOPMOST        As Long = -1
Public Sub SystrayOn(frm As Form, IconTooltipText As String)
    'adds icon to systray
    With vbTray
        .cbSize = Len(vbTray)
        .hWnd = frm.hWnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .szTip = Trim(IconTooltipText$) & vbNullChar
        .hIcon = frm.Icon
    End With
    
    Call Shell_NotifyIcon(NIM_ADD, vbTray)
    App.TaskVisible = False
    
End Sub
Public Sub SystrayOff(frm As Form)
    'removes icon from systray
    With vbTray
        .cbSize = Len(vbTray)
        .hWnd = frm.hWnd
        .uID = vbNull
    End With
    
    Call Shell_NotifyIcon(NIM_DELETE, vbTray)
    
End Sub
Public Sub FormOnTop(frm As Form)
    'puts your form ontop of all the other windows!
    Call SetWindowPos(frm.hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub

Public Sub ChangeSystrayToolTip(frm As Form, IconTooltipText As String)
    With vbTray
        .cbSize = Len(vbTray)
        .hWnd = frm.hWnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .szTip = Trim(IconTooltipText$) & vbNullChar
        .hIcon = frm.Icon
    End With
    
    Call Shell_NotifyIcon(NIM_MODIFY, vbTray)
    
End Sub
Public Sub PopupBalloon(frm As Form, Message As String, Title As String)
'Set a Balloon tip on Systray
'This removes any currenT Balloon Tip that is active
Call RemoveBalloon(frm)
   With vbTray
        .cbSize = Len(vbTray)
        .hWnd = frm.hWnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE Or NIM_MODIFY 'Or NIF_TIP 'NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = frm.Icon
        .dwState = 0
        .dwStateMask = 0
        .szInfo = Message & Chr(0)
        .szInfoTitle = Title & Chr(0)
        .dwInfoFlags = NIIF_INFO
    End With
    
    Call Shell_NotifyIcon(NIM_MODIFY, vbTray)
End Sub

Public Sub RemoveBalloon(frm As Form)
'Kill any current Ballonn tip on screen
  
   With vbTray
        .cbSize = Len(vbTray)
        .hWnd = frm.hWnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE Or NIM_MODIFY
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = frm.Icon
        .dwState = 0
        .dwStateMask = 0
        .szInfo = Chr(0)
        .szInfoTitle = Chr(0)
        .dwInfoFlags = NIIF_INFO
    End With
    
    Call Shell_NotifyIcon(NIM_MODIFY, vbTray)

End Sub

