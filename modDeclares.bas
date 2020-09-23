Attribute VB_Name = "modDeclares"
Option Explicit

Public Enum ModifyTypes
    AddText = 0
    DeleteText = 1
    ReplaceText = 2
    CutText = 3
    PasteText = 4
End Enum
'
'Public Type CHARRANGE
'    cpMin As Long
'    cpMax As Long
'End Type
'
'Public Type TEXTRANGE
'    chrg As CHARRANGE
'    lpstrText As Long    ' /* allocated by caller, zero terminated by RichEdit */
'End Type
'
'Private Const WM_USER = &H400
'
'Public Const EM_GETTEXTRANGE = (WM_USER + 75)
Private Const CB_SETDROPPEDWIDTH = &H160&
'
'Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Function SetComboDropDownWidth(cboCombo As ComboBox, lWidth As Long) As Boolean
'Return value is new width if call was successfull
    SetComboDropDownWidth = (SendMessage(cboCombo.hwnd, CB_SETDROPPEDWIDTH, lWidth, 0&) = lWidth)
End Function
