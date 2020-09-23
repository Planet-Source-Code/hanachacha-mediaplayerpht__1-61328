Attribute VB_Name = "mdlSys"
Option Explicit
'API Add  SystemTray Icon
Declare Function Shell_NotifyIcon _
Lib "shell32.dll" Alias _
"Shell_NotifyIconA" (ByVal dwMessage As Long, _
lpData As NOTIFYICONDATA) As Long

Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4

Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205

Public Type NOTIFYICONDATA
    cbSize              As Long
    hWnd                As Long
    uID                 As Long
    uFlags              As Long
    uCallbackMessage    As Long
    hIcon               As Long
    szTip               As String * 64
End Type
Public tData As NOTIFYICONDATA
Public Sub SysTrayIcon(pic As Picture)
    With tData
        .hIcon = pic.Handle
        .uFlags = NIF_ICON
    End With
    Shell_NotifyIcon NIM_MODIFY, tData
End Sub
Public Sub AddIcon(frm As Form, mnu As Menu)
    With tData
        .cbSize = Len(tData)
        .hWnd = frm.hWnd
        .uCallbackMessage = WM_MOUSEMOVE
        .uID = 1&
        .hIcon = frm.Icon.Handle
        .uFlags = NIF_ICON Or NIF_MESSAGE
    End With
        Shell_NotifyIcon NIM_ADD, tData
End Sub
Public Sub RemoveIcon()
    With tData
        .uFlags = 0
    End With
    Shell_NotifyIcon NIM_DELETE, tData
End Sub
Public Sub SysTip(ToolTip As String)
    With tData
        .szTip = ToolTip & vbNullChar
        .uFlags = NIF_TIP
    End With
    Shell_NotifyIcon NIM_MODIFY, tData
End Sub
