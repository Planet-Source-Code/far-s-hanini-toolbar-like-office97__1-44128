Attribute VB_Name = "Module1"
Option Explicit

Public Const WM_USER = &H400
Public Const TB_SETSTYLE = WM_USER + 56
Public Const TB_GETSTYLE = WM_USER + 57
Public Const TBSTYLE_FLAT = &H800

Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" _
(ByVal hwnd As Long, _
ByVal wMsg As Long, _
ByVal wParam As Long, _
ByVal lParam As Long) As Long

Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
(ByVal hWnd1 As Long, _
ByVal hWnd2 As Long, _
ByVal lpsz1 As String, _
ByVal lpsz2 As String) As Long

