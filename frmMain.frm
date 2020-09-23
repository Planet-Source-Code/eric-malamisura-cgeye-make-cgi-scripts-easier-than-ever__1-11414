VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "New Document"
   ClientHeight    =   7605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8655
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7605
   ScaleWidth      =   8655
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   6960
   End
   Begin VB.PictureBox picLines 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   6925
      Left            =   0
      ScaleHeight     =   6930
      ScaleWidth      =   300
      TabIndex        =   1
      Top             =   20
      Width           =   300
   End
   Begin RichTextLib.RichTextBox txtMain 
      Height          =   6975
      Left            =   360
      TabIndex        =   0
      Top             =   0
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   12303
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      ScrollBars      =   3
      RightMargin     =   50000
      TextRTF         =   $"frmMain.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public txtChanged As TriState

Public Enum TriState
    tsTrue = -1
    tsFalse = 0
    tsnone = -2
End Enum

Dim TextHeigth As Long, fTop As Integer
Dim LineCountChange As Integer
Dim FirstLine As Long
Dim FirstLineNow As Long

Public WindowNumber As Integer

Public bRedoing As Boolean
Public UndoStack As New Collection
Public RedoStack As New Collection
Public lUndoCount As Long

Private Type CHARRANGE
    cpMin As Long
    cpMax As Long
End Type

Private Type TEXTRANGE
    chrg As CHARRANGE
    lpstrText As Long    ' /* allocated by caller, zero terminated by RichEdit */
End Type

Private Const WM_USER = &H400

Private Const EM_GETTEXTRANGE = (WM_USER + 75)
Private Const CB_SETDROPPEDWIDTH = &H160&

Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Property Get CanPaste() As Boolean
    CanPaste = SendMessageLong(txtMain.hwnd, EM_CANPASTE, 0, 0)
End Property
Public Sub UpdateStatus()
    mdiMain.cmdRedo.Enabled = Not (RedoStack.Count = 0)
    mdiMain.cboRedo.Enabled = Not (RedoStack.Count = 0)
    mdiMain.cmdUndo.Enabled = Not (UndoStack.Count = 0)
    mdiMain.cboUndo.Enabled = Not (UndoStack.Count = 0)
End Sub
Public Function InsertTag(Tag1 As String, Tag2 As String, Optional PrintString As Boolean = False)
    On Error Resume Next
    Dim cUndo As New clsUndo

      cUndo.lStart = txtMain.SelStart
    Dim Buf1 As Integer, Buf2 As Integer
      Buf1 = txtMain.SelStart
      cUndo.lStart = txtMain.SelStart
      If txtMain.SelLength > 0 Then
          cUndo.sDelText = txtMain.SelText
          Buf2 = txtMain.SelLength
          txtMain.SelLength = 0
          txtMain.SelStart = Buf1
          
          If PrintString = True Then
              txtMain.SelText = "print " & """" & Tag1
              txtMain.SelStart = Buf1 + Buf2 + Len(Tag1) + 7
              txtMain.SelText = Tag2 & """"
              txtMain.SelStart = Buf1
              txtMain.SelLength = Buf2 + Len(Tag1) + Len(Tag2) + 8
            Else
              txtMain.SelText = Tag1
              txtMain.SelStart = Buf1 + Buf2 + Len(Tag1)
              txtMain.SelText = Tag2
              txtMain.SelStart = Buf1
              txtMain.SelLength = Buf2 + Len(Tag1) + Len(Tag2)
          End If
          cUndo.sAddText = txtMain.SelText
        Else
        
          If PrintString = True Then
              txtMain.SelText = "print " & """" & Tag1 & Tag2 & """"
              txtMain.SelStart = Buf1 + Len(Tag1) + 7
              cUndo.sAddText = "print " & """" & Tag1 & Tag2 & """"
            Else
              txtMain.SelText = Tag1 & Tag2
              txtMain.SelStart = Buf1 + Len(Tag1)
              cUndo.sAddText = Tag1 & Tag2
          End If
        
      End If
      AddToUndoStack cUndo
      txtMain.SetFocus
End Function

Public Sub InsertFont(Face As String, Size As Integer, Style As String, Color As String)
  Dim sBuf1 As String
  Dim sBuf2 As String

    If Size = 0 And Face = "" Then
        sBuf1 = "<font>"

      ElseIf Size > 0 And Face = "" And Color = "" Then
        sBuf1 = "<font size=\" & """" & Size & "\" & """" & ">"

      ElseIf Size > 0 And Face = "" And Not Color = "" Then
        sBuf1 = "<font size=\" & """" & Size & "\" & """" & " color=\" & """" & Color & "\" & """" & ">"
      ElseIf Not Face = "" And Size = 0 And Color = "" Then
        sBuf1 = "<font face=\" & """" & Face & "\" & """" & ">"
      ElseIf Not Face = "" And Size = 0 And Not Color = "" Then
        sBuf1 = "<font face=\" & """" & Face & "\" & """" & " color=\" & """" & Color & "\" & """" & ">"
      ElseIf Not Face = "" And Size > 0 And Color = "" Then
        sBuf1 = "<font face=\" & """" & Face & "\" & """" & " size=\" & """" & Size & "\" & """" & ">"
      ElseIf Not Face = "" And Size > 0 And Not Color = "" Then
        sBuf1 = "<font face=\" & """" & Face & "\" & """" & " size=\" & """" & Size & "\" & """" & " color=\" & """" & Color & "\" & """" & ">"
    End If

    sBuf2 = "</font>"

    If Style = "Italic" Then
        sBuf1 = sBuf1 & "<I>"
        sBuf2 = "</I></font>"
      ElseIf Style = "Bold" Then
        sBuf1 = sBuf1 & "<B>"
        sBuf2 = "</B></font>"
      ElseIf Style = "Bold Italic" Then
        sBuf1 = sBuf1 & "<B><I>"
        sBuf2 = "</I></B></font>"
    End If

    InsertTag sBuf1, sBuf2
End Sub
Public Sub InsertURL(URL As String, Optional Frame As Boolean, Optional FrameType As String)
    If Frame = True Then
        InsertTag "<a href=\" & """" & URL & "\" & """" & " " & "target=\" & """" & FrameType & "\" & """" & ">", "</a>", True
      Else
        InsertTag "<a href=\" & """" & URL & "\" & """" & ">", "</a>", True
    End If

End Sub
Public Sub InsertImage(URL As String, Border As Integer, Optional Sizes As Boolean, Optional Width As Integer, Optional Height As Integer)
    If Sizes = False Then
        InsertTag "<img " & "border=\" & """" & Border & "\" & """" & " src=\" & """" & URL & "\" & """" & ">", "", True
      Else
        InsertTag "<img " & "border=\" & """" & Border & "\" & """" & " src=\" & """" & URL & "\" & """" & " width=\" & """" & Width & "\" & """" & " " & "height=\" & """" & Height & "\" & """" & ">", "", True
    End If
End Sub

Private Sub Form_GotFocus()
    UpdateStatus
End Sub

Private Sub Form_Load()
    mdiMain.WindowCount = mdiMain.WindowCount + 1
    WindowNumber = mdiMain.WindowCount
    If txtMain.FileName = "" Then Caption = "Document" & WindowNumber & ".cgi"
    
    lUndoCount = mdiMain.varUndoLimit
    SetComboDropDownWidth mdiMain.cboUndo, 110
    SetComboDropDownWidth mdiMain.cboRedo, 110

   
    'this will be editable
    txtChanged = tsFalse
    TextHeigth = txtMain.Font.Size  '// We need this to find out about the size of font
    UpdateStatus

End Sub
Public Function CloseDocument() As Boolean
    Select Case txtChanged
      Case tsTrue
  Dim YesNo
        YesNo = MsgBox("The following document: " & vbCrLf & Right(Me.Caption, Len(Me.Caption) - 1) & vbCrLf & vbCrLf & "Has had changes made to it." & vbCrLf & vbCrLf & "Would you like to save the changes?", vbYesNoCancel + vbQuestion, "Save Changes?")
        If YesNo = vbYes Then
            mdiMain.ShowSave
            CloseDocument = False
          ElseIf YesNo = vbNo Then
            CloseDocument = False
          ElseIf YesNo = vbCancel Then
            CloseDocument = True
        End If
      Case Else
        CloseDocument = False
    End Select
End Function

Private Sub Form_Paint()

    If mdiMain.varShowLines = True Then
        DrawNumbers
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = CloseDocument
End Sub

Public Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub 'if you minimize it why do you want to resize?
    
picLines.Move 0, 20, 610, Me.ScaleHeight - 60 'size the piclines
    
     If mdiMain.varShowLines = True Then 'check to see if lines are showing
        picLines.Visible = True
        txtMain.Move picLines.Width, 0, Me.ScaleWidth - picLines.Width, Me.ScaleHeight
      Else
        txtMain.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
        picLines.Visible = False
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

'disable undo/redo because form in focus is no longer valid..

UpdateStatus

mdiMain.cmdUndo.Enabled = False
mdiMain.cmdRedo.Enabled = False
mdiMain.cboUndo.Enabled = False
mdiMain.cboRedo.Enabled = False

mdiMain.WindowCount = mdiMain.WindowCount - 1 'subtract the count for the windows
End Sub

Private Sub Timer1_Timer()
'// Get first visible line in rtfText
    DoEvents
    FirstLine = SendMessage(txtMain.hwnd, EM_GETFIRSTVISIBLELINE, 0&, 0&)
    FirstLine = FirstLine   '// Change start from 0 to 1 if necessary
    DoEvents
    If Not FirstLineNow = FirstLine Then DrawNumbers '// I can't hook to a scrollbar so I used a sucker-timer
    DoEvents
End Sub

Private Sub txtMain_Change()
  Dim LineCount As Long
    If txtChanged = tsFalse Then
        Me.Caption = "*" & Me.Caption
    End If
    txtChanged = tsTrue

    '// Get number of lines in Rtftext
    LineCount = SendMessage(txtMain.hwnd, EM_GETLINECOUNT, 0&, 0&)
    LineCount = LineCount - 1  '// Change start from 0 to 1

    If LineCount = LineCountChange Then
        GoTo skip:    '// Line count is still the same
      Else
        DrawNumbers '// new Line count is required
    End If
    
skip:

    If bRedoing Then
        bRedoing = False
        ClearStack RedoStack
    End If

    UpdateStatus
    UpdateCopyPaste

End Sub

Private Sub txtMain_Click()
'MsgBox Asc(txtMain.Text)
End Sub

Private Sub txtMain_GotFocus()
    On Error Resume Next
    Dim Control As Control
      For Each Control In Controls
          Control.TabStop = False
      Next Control
      UpdateStatus
      UpdateCopyPaste
End Sub

Sub DrawNumbers()
  Dim LineCount As Long '// How many lines in total
  Dim i As Long      '// Just an integer
  Dim TempBuf As String
  Static WidthCount As Integer
'// Get number of lines in Rtftext
    LineCount = SendMessage(txtMain.hwnd, EM_GETLINECOUNT, 0&, 0&)
    LineCount = LineCount - 1  '// Change start from 0 to 1

    '// Same lines ?
    LineCountChange = LineCount

    '// Get first visible line in rtfText
    FirstLine = SendMessage(txtMain.hwnd, EM_GETFIRSTVISIBLELINE, 0&, 0&)
    FirstLine = FirstLine   '// Change start from 0 to 1 if necessary

    picLines.Cls '// Clear the PicLines
    picLines.CurrentY = 40  '// Move the .top text by 40 twips

    '// Print the number of each line on a picture
    For i = 0 To LineCount - FirstLine
        picLines.CurrentY = picLines.CurrentY + 7.49 '// Where on Y
        picLines.CurrentX = 20 '-2                   '// Where on X
        picLines.Print i + FirstLine + 1             '// print the number
    Next
    picLines.Refresh
    'LineCountChange = LineCount '// Remember the last line count
    FirstLineNow = FirstLine     '// Is the first visible line still the same ?
End Sub
Public Sub GotoLine(LineNum As Long, Highlight As Boolean)
    On Error GoTo done:
  Dim Temp As Integer
  Dim Num As Integer
  Dim Pos  As Integer
  Dim LastPos As Integer
  Dim Cut As Integer
    If LineNum = 0 Then Exit Sub
    Pos = 1
    Num = 1
    Temp = 0
    Do
        LastPos = Temp
        Temp = InStr(Pos, txtMain.Text, vbLf)
        If Temp = 0 Then GoTo redo:
        If Temp >= 1 Then
            Num = Num + 1
            Pos = Temp + 2
        End If
    
    Loop Until Num >= LineNum

    Cut = 1

redo:
    If Temp = 0 Then
        LastPos = 0
        Temp = Len(txtMain.Text)
        Cut = 0
    End If

    If LineNum = 1 Then
        Temp = 0
        LastPos = InStr(1, txtMain.Text, vbLf)
        If LastPos = 0 Then
            LastPos = Len(txtMain.Text)
        End If

        Cut = 0
    End If

    txtMain.SelStart = Temp
    If Highlight = True Then txtMain.SelLength = LastPos - Cut
    txtMain.SetFocus
done:
End Sub

Public Function GetUndoText(ModifyType As ModifyTypes) As String
    Select Case ModifyType
      Case DeleteText
        GetUndoText = "Delete Text"
      Case AddText
        GetUndoText = "Add Text"
      Case ReplaceText
        GetUndoText = "Replace Text"
      Case PasteText
        GetUndoText = "Paste Text"
      Case CutText
        GetUndoText = "Cut Text"
    End Select
End Function
Public Sub AddToUndoStack(cUndo As clsUndo)
    If UndoStack.Count = lUndoCount Then
        UndoStack.Remove (1)
    End If
    UndoStack.Add cUndo
    UpdateStatus
End Sub
Private Sub ClearStack(Stack As Collection)
    On Error Resume Next
    Dim i As Long
      For i = 1 To Stack.Count Step 1
          Stack.Remove (i)
      Next
End Sub
Public Sub ClearUndoRedo()
  Dim i As Long

    For i = 1 To UndoStack.Count
        UndoStack.Remove (1)
    Next i

    i = 1

    For i = 1 To RedoStack.Count
        RedoStack.Remove (1)
    Next i

    UpdateStatus
End Sub
Private Sub txtMain_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim cUndo As New clsUndo
  Dim lAmount As Long
  Dim lOldPos As Long

    If IsMoveKey(KeyCode) Then Exit Sub

    If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        If KeyCode = vbKeyBack Then
            lAmount = -1
          Else
            lAmount = 0
        End If
        With cUndo
            If txtMain.SelLength = 0 Then
                '// there is no text being deleted
                If lAmount = -1 And txtMain.SelStart = 0 Then
                    '// we aren't going anywhere!
                    GoTo exitundo
                End If
                '// set the start pos
                .lStart = IIf(txtMain.SelStart = 0, 0, txtMain.SelStart + lAmount)
                '// see what we are going to delete
                .sDelText = TextInRange(txtMain.SelStart + lAmount, 1)
                
                '// if there is part of vbCrLf selected
                '// set the length to 2 instead
                If InStr(1, Chr(10) & Chr(13), .sDelText) Then
                    .sDelText = vbCrLf
                    If .sDelText = Chr(10) Then
                        '// deleting end of CrLf
                        .lStart = .lStart - 1
                    End If
                End If
              Else
                '// save the text that is being deleted
                .lStart = txtMain.SelStart
                .sDelText = txtMain.SelText
            End If
            .ModifyType = DeleteText
            AddToUndoStack cUndo
exitundo:
        End With
    End If
    
    If Shift = vbCtrlMask And KeyCode <> vbKeyControl Then
        With cUndo
            Select Case KeyCode
              Case vbKeyV
                '// add the pasted text to the Undo stack
                .lStart = txtMain.SelStart
                .sAddText = Clipboard.GetText(vbCFText)
                .sDelText = txtMain.SelText
                MsgBox .sDelText
                .ModifyType = PasteText
                AddToUndoStack cUndo
                txtMain.SelText = .sAddText
                KeyCode = 0
              Case vbKeyX
                '// cut
                .lStart = txtMain.SelStart
                .sDelText = txtMain.SelText
                
                .ModifyType = CutText
                AddToUndoStack cUndo
              Case vbKeyZ
                mdiMain.cmdUndo_Click
                KeyCode = 0
              Case vbKeyY
                mdiMain.cmdRedo_Click
                KeyCode = 0
            End Select
        End With
    End If
End Sub
    
Public Property Get TextInRange(ByVal lStart As Long, ByVal lLen As Long)
  Dim tR As TEXTRANGE
  Dim lR As Long
  Dim sText As String
  Dim b() As Byte
  Dim lEnd As Long

    lEnd = lStart + lLen
    
    tR.chrg.cpMin = lStart
    tR.chrg.cpMax = lEnd
    
    sText = String$(lEnd - lStart + 1, 0)
    b = StrConv(sText, vbFromUnicode)
    ' VB won't do the terminating null for you!
    ReDim Preserve b(0 To UBound(b) + 1) As Byte
    b(UBound(b)) = 0
    tR.lpstrText = VarPtr(b(0))
    
    lR = SendMessage(txtMain.hwnd, EM_GETTEXTRANGE, 0, tR)
    If (lR > 0) Then
        sText = StrConv(b, vbUnicode)
        TextInRange = Left$(sText, lR)
    End If
End Property

Private Function IsMoveKey(KeyCode As Integer) As Boolean
    Select Case KeyCode
      Case vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown, vbKeyPageUp, vbKeyPageDown, vbKeyHome, vbKeyEnd, vbKeyShift
        IsMoveKey = True
    End Select
End Function

Private Sub txtMain_KeyPress(KeyAscii As Integer)
  Dim cUndo As New clsUndo
    
    If KeyAscii = vbKeyBack Then
        '// ignore. Note that the Delete key does not trigger this event
      ElseIf KeyAscii >= 32 Or KeyAscii = 13 Then
        '// ignore keycodes under 32
        With cUndo
            .lStart = txtMain.SelStart
            If KeyAscii = 13 Then
                .sAddText = vbCrLf
              Else
                .sAddText = Chr(KeyAscii)
            End If
            .sDelText = txtMain.SelText
            .ModifyType = IIf(.sDelText = "", AddText, ReplaceText)
        End With
        AddToUndoStack cUndo
    End If
    Set cUndo = Nothing
    
    If KeyAscii = 13 Then
        Colorize txtMain, &H8000&, &HFF0000, &H80&, False
    End If

End Sub
Public Sub CommentBlock()
  Dim Buffer As Variant
  Dim cUndo As New clsUndo

    If txtMain.SelLength = 0 Then
        cUndo.lStart = txtMain.SelStart
        txtMain.SelText = "#"
        cUndo.sAddText = "#"
        AddToUndoStack cUndo
      Else
        cUndo.lStart = txtMain.SelStart
        cUndo.sDelText = txtMain.SelText
        Buffer = txtMain.SelText
        Buffer = "#" & Buffer
        Buffer = Replace(Left(Buffer, Len(Buffer)), vbLf, vbLf & "#")
        txtMain.SelText = Buffer
        cUndo.sAddText = Buffer
        AddToUndoStack cUndo
    End If

End Sub

Public Sub UncommentBlock()
  Dim Buffer As Variant
  Dim FirstLineBuffer As Variant
  Dim cUndo As New clsUndo
  Dim Proceed As Boolean
  Dim i As Integer
    cUndo.lStart = txtMain.SelStart
    cUndo.sDelText = txtMain.SelText
    Buffer = txtMain.SelText
    If InStr(Buffer, vbLf) Then
        FirstLineBuffer = Left(Buffer, InStr(Buffer, vbLf))
      Else
        FirstLineBuffer = Buffer
    End If
    If InStr(FirstLineBuffer, "#") Then
        For i = 1 To Len(FirstLineBuffer)
            If Mid(FirstLineBuffer, i, 1) = "" Or Mid(FirstLineBuffer, i, 1) = " " Or Mid(FirstLineBuffer, i, 1) = "#" Then
                Proceed = True
                Exit For
              Else
                Proceed = False
            End If
        Next i
        If Proceed = True Then
            Buffer = Right(Buffer, Len(Buffer) - InStr(FirstLineBuffer, "#"))
        End If
    End If
skiptrim:
    If InStr(Buffer, vbLf & "#") Then
        Buffer = Replace(Left(Buffer, Len(Buffer)), vbLf & "#", vbLf)
    End If
    txtMain.SelText = Buffer
    cUndo.sAddText = Buffer
    AddToUndoStack cUndo
End Sub
Public Sub UpdateCopyPaste()
    If CanPaste = True Then
        mdiMain.cmdButtonBar(5).Enabled = True
      Else
        mdiMain.cmdButtonBar(5).Enabled = False
    End If

    If Len(txtMain.SelText) > 0 Then
        mdiMain.cmdButtonBar(3).Enabled = True
        mdiMain.cmdButtonBar(4).Enabled = True
      Else
        mdiMain.cmdButtonBar(3).Enabled = False
        mdiMain.cmdButtonBar(4).Enabled = False
    End If
End Sub

Private Sub txtMain_LostFocus()
    UpdateCopyPaste
End Sub

Private Sub txtMain_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    UpdateCopyPaste
End Sub

Private Sub txtMain_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    UpdateCopyPaste
End Sub
