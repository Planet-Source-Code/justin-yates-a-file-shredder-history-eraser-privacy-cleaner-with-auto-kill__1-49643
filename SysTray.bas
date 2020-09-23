Attribute VB_Name = "SysTray"
'System Tray API
Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

'System Tray Constants
Public Const WM_MOUSEMOVE = &H200
Public Const NIF_ICON = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_TIP = &H4
Public Const NIM_ADD = &H0
Public Const NIM_DELETE = &H2
Public Const MAX_TOOLTIP As Integer = 64

Public Const MouseMove = 7680
Public Const LeftMouseDown = 7695
Public Const LeftMouseUp = 7710
Public Const LeftDblClick = 7725
Public Const RightMouseDown = 7740
Public Const RightMouseUp = 7755
Public Const RightDblClick = 7770

'System Tray Type
Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * MAX_TOOLTIP
End Type

Public nfIconData As NOTIFYICONDATA

Public Sub ShowIcon(fForm As Form, Descrpition As String, Icon As Long)
    With nfIconData
        .hWnd = fForm.hWnd
        .uID = fForm.Icon
        .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Icon
        .szTip = Descrpition & Chr$(0)
        .cbSize = Len(nfIconData)
    End With
    
    Call Shell_NotifyIcon(NIM_ADD, nfIconData)
End Sub

Public Sub HideIcon()
    Call Shell_NotifyIcon(NIM_DELETE, nfIconData)
End Sub

