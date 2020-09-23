Attribute VB_Name = "Variables"
'Public Declare Function SendMessage Lib "User32" Alias _
    '    "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
    '        ByVal wParam As Long, lParam As Long) As Long

Option Explicit
Public Const WM_USER = &H400
Public Const EM_HIDESELECTION = WM_USER + 63

'// Constants
Public Const EM_GETLINECOUNT = &HBA        '// Total Line Count
Public Const EM_GETFIRSTVISIBLELINE = &HCE '// First Visible Line
Public Const WM_VSCROLL = &H115            '// Vertical Scrolling
