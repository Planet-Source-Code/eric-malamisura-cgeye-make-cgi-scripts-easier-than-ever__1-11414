Attribute VB_Name = "SyntaxColorize"
Option Explicit
Public Keywords As String

Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageByNum Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

'// Win API Const
Private Const EM_GETFIRSTVISIBLELINE = &HCE
Private Const EM_LINEINDEX = &HBB
Private Const EM_GETRECT = &HB2
Private Const WM_GETFONT = &H31

'// Variables

'//Variables for FirstVisible/LastVisibles
Dim FirstVisibleLine As Long
Dim LastVisibleLine As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type TEXTMETRIC
    tmHeight As Long
    tmAscent As Long
    tmDescent As Long
    tmInternalLeading As Long
    tmExternalLeading As Long
    tmAveCharWidth As Long
    tmMaxCharWidth As Long
    tmWeight As Long
    tmOverhang As Long
    tmDigitizedAspectX As Long
    tmDigitizedAspectY As Long
    tmFirstChar As Byte
    tmLastChar As Byte
    tmDefaultChar As Byte
    tmBreakChar As Byte
    tmItalic As Byte
    tmUnderlined As Byte
    tmStruckOut As Byte
    tmPitchAndFamily As Byte
    tmCharSet As Byte
End Type
Public Sub Colorize(RTFBox As RichTextBox, CommentColor As Long, ScriptColor As Long, StringColor As Long, Optional All As Boolean = False)
  Dim lTextSelPos As Long, lTextSelLen As Long

    LockWindowUpdate RTFBox.hwnd

    On Error GoTo ErrHandler
  Dim i As Long
  Dim sBuffer As String, lBufferLen As Long
  Dim lSelPos As Long, lSelLen As Long
  Dim sTempBuffer As String
  Dim sSearchChar As String, lSearchCharLen As Long
  Dim SelectColor As Long
  Dim ByteArray() As Byte
  

    With RTFBox
sBuffer = .Text & " "
lBufferLen = Len(sBuffer)
.SelColor = vbBlack
If All = True Then
.SelStart = 1
Else
lTextSelPos = RTFBox.SelStart
End If

  ByteArray() = StrConv(sBuffer, vbFromUnicode)  'used this for speedier parsing...
  
  'this module isnt all that great...I'm not a total mathmatician so it takes me a couple
  'mins to figure out how to do some of this crap...but it works decent enough!!!
  
        For i = FirstVisibleChar(RTFBox, All) To LastVisibleChar(RTFBox, lBufferLen, All)
            Select Case ByteArray(i - 1)

              Case 39, 60, 35
              
                 If Mid(sBuffer, i, 4) = "<!--" Then '// HTML Comment
                    sSearchChar = "-->"
                    lSearchCharLen = 3
                    SelectColor = CommentColor
                  ElseIf Mid(sBuffer, i, 1) = "#" Then    '//Perl Comment
                    sSearchChar = vbCrLf
                    lSearchCharLen = 0
                  
  Dim Buffer As Variant
                    
                    If InStrRev(sBuffer, vbCrLf, i) > 0 Then
                        Buffer = Right(sBuffer, Len(sBuffer) - InStrRev(sBuffer, vbCrLf, i))
                        Buffer = Left(Buffer, InStr(Buffer, vbCrLf))
                        If Not Buffer = "" Then Buffer = Left(Buffer, InStr(Buffer, "#") - 1)
                    End If

                    If Buffer = "" Or InStr(LCase(Buffer), "color:") = 0 And InStr(LCase(Buffer), "color=") = 0 _
                        And InStr(LCase(Buffer), "alink=") = 0 And InStr(LCase(Buffer), "vlink=") = 0 _
                        And InStr(LCase(Buffer), "link=") = 0 And InStr(LCase(Right(Buffer, 2)), """") = 0 Then
                        If i = 1 Then
                            SelectColor = &H800080
                          Else
                            SelectColor = CommentColor
                        End If
                      Else
                        SelectColor = vbBlack
                    End If
                        
                  Else                                    '// None
                    GoTo ExitComment
                End If
                '// Kill TempBuffer
                sTempBuffer = ""
          
                '// Colorize the comment string
                .SelStart = i - 1
                lSelLen = InStr(i, sBuffer, sSearchChar) + lSearchCharLen
                If lSelLen <> lSearchCharLen Then '// FileEnd ?
                    lSelLen = lSelLen - i
                  Else
                    lSelLen = lBufferLen - i
                End If
                .SelLength = lSelLen
          
                .SelColor = SelectColor
                i = .SelStart + .SelLength
          
ExitComment:

                '--------------------Color Script
              Case 97 To 122, 65 To 90, 92
                '// a to  z ,  A to Z , \
          
                If sTempBuffer = "" Then lSelPos = i
                sTempBuffer = sTempBuffer & Mid(sBuffer, i, 1)
        
              Case Else
                If Trim(sTempBuffer) <> "" Then
                    If InStr(1, Keywords, "|" & sTempBuffer & "|", 1) <> 0 Then
                        .SelStart = lSelPos - 1
                        .SelLength = Len(sTempBuffer)
                        .SelColor = ScriptColor
                    End If
                End If
      
                sTempBuffer = ""
            End Select
        Next
    End With

ErrHandler:

    RTFBox.SelStart = lTextSelPos
    RTFBox.SelColor = vbBlack 'this ensures the typed color is black..
    LockWindowUpdate 0
End Sub

Private Function FirstVisibleChar(RTFBox As RichTextBox, Optional All As Boolean = False) As Long

    FirstVisibleLine = SendMessage(RTFBox.hwnd, EM_GETFIRSTVISIBLELINE, 0, 0&)
    FirstVisibleChar = SendMessageByNum(RTFBox.hwnd, EM_LINEINDEX, FirstVisibleLine, 0&)
    If FirstVisibleChar = 0 Then FirstVisibleChar = 1
End Function

Private Function LastVisibleChar(RTFBox As RichTextBox, LenFile As Long, All As Boolean) As Long
  Dim rc As RECT
  Dim tm As TEXTMETRIC
  Dim hdc As Long
  Dim lFont As Long
  Dim OldFont As Long
  Dim di As Long
  Dim lc As Long
  Dim VisibleLines As Long

    lc = SendMessage(RTFBox.hwnd, EM_GETRECT, 0, rc)
    lFont = SendMessage(RTFBox.hwnd, WM_GETFONT, 0, 0)
    hdc = GetDC(RTFBox.hwnd)
    If lFont <> 0 Then OldFont = SelectObject(hdc, lFont)
    di = GetTextMetrics(hdc, tm)
    If lFont <> 0 Then lFont = SelectObject(hdc, OldFont)

    If All = True Then
        VisibleLines = Len(RTFBox.Text)
      Else
        VisibleLines = (rc.Bottom - rc.Top) / tm.tmHeight + 50
    End If
  
    di = ReleaseDC(RTFBox.hwnd, hdc)
  
    LastVisibleLine = SendMessage(RTFBox.hwnd, EM_GETFIRSTVISIBLELINE, 0, 0&)
    LastVisibleLine = LastVisibleLine + VisibleLines
  
    LastVisibleChar = SendMessageByNum(RTFBox.hwnd, EM_LINEINDEX, LastVisibleLine, 0&)
    If LastVisibleChar = -1 Or LastVisibleChar = 0 Then LastVisibleChar = LenFile
End Function
