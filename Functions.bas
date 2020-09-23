Attribute VB_Name = "Module1"
Option Explicit
Public Const WM_USER = &H400
Public Const EM_SETTEXTMODE = (WM_USER + 89)
Public Const EM_UNDO = &HC7
Public Const EM_REDO = (WM_USER + 84)
Public Const EM_CANPASTE = (WM_USER + 50)
Public Const EM_CANUNDO = &HC6&
Public Const EM_CANREDO = (WM_USER + 85)
Public Const EM_GETUNDONAME = (WM_USER + 86)
Public Const EM_GETREDONAME = (WM_USER + 87)
Public Const WM_COPY = &H301
Public Const WM_CUT = &H300
Public Const WM_PASTE = &H302
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Const GWL_STYLE = (-16)
Const ES_NUMBER = &H2000&
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Function FileGetName(FileName As String) As String
  Dim Pos As Integer
    Pos = InStrRev(FileName, "\")
    FileGetName = Right(FileName, Len(FileName) - Pos)
End Function
Public Function FileCheck(Path$) As Boolean
'USAGE: If FileCheck("C:\windows\kewl.exe") then msgbox "it was found"
    FileCheck = True 'Assume Success
    On Error Resume Next
    Dim Disregard As Long
      Disregard = FileLen(Path)
      If Err <> 0 Then FileCheck = False
End Function
Public Function returnRelPath(ByVal strHome As String, ByVal strNewLoc As String) As String
  Dim ip1 As Long, i As Integer
  Dim lCount As Long
  Dim blDone As Boolean
    
'remove a file name if there is one
'    strHome = Left$(strHome, InStrRev(strHome, "\"))
    
'if the left part of the string of the new location
'mathces strHome, then we have an easy one

    If Mid$(strNewLoc, 4, Len(strHome)) = strHome Then
        returnRelPath = Replace(Mid$(strNewLoc, Len(strHome)), "\", "/")
      Else
        'else we have to loop to find the relative path
'        MsgBox strNewLoc
'        MsgBox strHome
        ip1 = Len(strHome)
        Do
            ip1 = InStrRev(strHome, "\", ip1)
            If ip1 <> 0 Then
                If Left$(strNewLoc, ip1) = Left$(strHome, ip1) Then
                    blDone = True
                  Else
                    lCount = lCount + 1
                End If
                ip1 = ip1 - 1
            End If
        Loop Until (ip1 = 0) Or (blDone = True)
      
        'build up a string of "../"
        For i = 1 To lCount
            returnRelPath = returnRelPath & "../"
        Next i
      
        'and add the path on the end of it
        returnRelPath = returnRelPath & Replace(Mid$(strNewLoc, ip1 + 2), "\", "/")
    End If
  
End Function

Public Sub SetNumber(NumberText As TextBox, Flag As Boolean)
  Dim curstyle As Long
  Dim newstyle As Long

    curstyle = GetWindowLong(NumberText.hwnd, GWL_STYLE)
    If Flag Then
        curstyle = curstyle Or ES_NUMBER
      Else
        curstyle = curstyle And (Not ES_NUMBER)
    End If

    newstyle = SetWindowLong(NumberText.hwnd, GWL_STYLE, curstyle)
    NumberText.Refresh
End Sub
Public Function OpenIt(ToOpen As String)
    ShellExecute &O0, "Open", ToOpen, &O0, &O0, 1
End Function
