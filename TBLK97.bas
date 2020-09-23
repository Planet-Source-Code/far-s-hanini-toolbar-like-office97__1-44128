Attribute VB_Name = "ToolBarLook97"
Option Explicit

Public Const WM_USER = &H400
Public Const TB_SETSTYLE = WM_USER + 56
Public Const TB_GETSTYLE = WM_USER + 57
Public Const TBSTYLE_FLAT = &H800

Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Public Sub MakeToolBarFlat(ToolBar1 As Toolbar)
    Dim style As Long
    Dim hToolbar As Long
    Dim r As Long
    'get the handle of the toolbar
    hToolbar = FindWindowEx(ToolBar1.hwnd, 0&, "ToolbarWindow32", vbNullString)
    'retreive the toolbar styles
    style = SendMessageLong(hToolbar, TB_GETSTYLE, 0&, 0&)
    'Set the new style flag
    If style And TBSTYLE_FLAT Then
    style = style Xor TBSTYLE_FLAT
    Else
    style = style Or TBSTYLE_FLAT
    End If
    'apply the new style to the toolbar
    r = SendMessageLong(hToolbar, TB_SETSTYLE, 0, style)
End Sub
