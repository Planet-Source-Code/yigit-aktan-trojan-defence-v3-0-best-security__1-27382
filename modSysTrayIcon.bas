Attribute VB_Name = "modSysTrayIcon"
'**************************************
'Windows API/Global Declarations for :Wi
'     ndows System Tray
'thanxs to arno pijnappels
'modifications also by Rich Jones <rich_
'     jones@wmg.com>
' By: Found on the World Wide Web
'**************************************

Option Explicit

Public Const WM_LBUTTONDBLCLICK = &H203
Public Const WM_RBUTTONUP = &H205
Public Const WM_MOUSEMOVE = &H200
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4


Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
    End Type
    Public VBGTray As NOTIFYICONDATA


Declare Function Shell_NotifyIcon Lib "shell32" Alias _
                "Shell_NotifyIconA" (ByVal dwMessage As Long, _
                            pnid As NOTIFYICONDATA) As Boolean


Public Declare Function SetForegroundWindow Lib "user32" _
                (ByVal hwnd As Long) As Long

