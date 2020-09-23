VERSION 5.00
Begin VB.MDIForm mdiMain 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   Caption         =   "CgEye By Elucid Software (Preview Release 3)"
   ClientHeight    =   7155
   ClientLeft      =   180
   ClientTop       =   735
   ClientWidth     =   10260
   Icon            =   "mdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin VB.PictureBox picStatusBar 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   0
      ScaleHeight     =   285
      ScaleWidth      =   10260
      TabIndex        =   30
      Top             =   6870
      Width           =   10260
      Begin VB.Label txtStatus 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Thank you for using CgEye By Elucid Software"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   31
         Top             =   35
         Width           =   10215
      End
   End
   Begin VB.PictureBox picToolbar 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   10260
      TabIndex        =   0
      Top             =   0
      Width           =   10260
      Begin VB.CommandButton cmdButtonBar 
         Enabled         =   0   'False
         Height          =   300
         Index           =   19
         Left            =   3990
         Picture         =   "mdiMain.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Page Colors"
         Top             =   55
         Width           =   330
      End
      Begin VB.CommandButton cmdButtonBar 
         Height          =   300
         Index           =   18
         Left            =   8490
         Picture         =   "mdiMain.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Insert Underline Tag"
         Top             =   55
         Width           =   330
      End
      Begin VB.CommandButton cmdButtonBar 
         Height          =   300
         Index           =   17
         Left            =   8110
         Picture         =   "mdiMain.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Insert Italic Tag"
         Top             =   55
         Width           =   330
      End
      Begin VB.CommandButton cmdButtonBar 
         Height          =   300
         Index           =   16
         Left            =   7740
         Picture         =   "mdiMain.frx":08A8
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Insert Bold Tag"
         Top             =   55
         Width           =   330
      End
      Begin VB.CommandButton cmdButtonBar 
         Height          =   300
         Index           =   14
         Left            =   8960
         Picture         =   "mdiMain.frx":09F2
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Find"
         Top             =   55
         Width           =   330
      End
      Begin VB.CommandButton cmdButtonBar 
         Height          =   300
         Index           =   13
         Left            =   6870
         Picture         =   "mdiMain.frx":0B3C
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Undo Comment"
         Top             =   55
         Width           =   330
      End
      Begin VB.CommandButton cmdButtonBar 
         Height          =   300
         Index           =   12
         Left            =   6005
         Picture         =   "mdiMain.frx":0E46
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Insert Right Tags"
         Top             =   55
         Width           =   330
      End
      Begin VB.CommandButton cmdButtonBar 
         Height          =   300
         Index           =   11
         Left            =   5630
         Picture         =   "mdiMain.frx":0F90
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Insert Center Tags"
         Top             =   55
         Width           =   330
      End
      Begin VB.CommandButton cmdButtonBar 
         Height          =   300
         Index           =   10
         Left            =   5255
         Picture         =   "mdiMain.frx":10DA
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Insert Left Tags"
         Top             =   55
         Width           =   330
      End
      Begin VB.CommandButton cmdButtonBar 
         Height          =   300
         Index           =   9
         Left            =   6480
         Picture         =   "mdiMain.frx":1224
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Comment Block"
         Top             =   55
         Width           =   330
      End
      Begin VB.CommandButton cmdButtonBar 
         Height          =   300
         Index           =   8
         Left            =   4740
         Picture         =   "mdiMain.frx":152E
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Insert Image Tag"
         Top             =   55
         Width           =   330
      End
      Begin VB.CommandButton cmdButtonBar 
         Height          =   300
         Index           =   7
         Left            =   4370
         Picture         =   "mdiMain.frx":1678
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Insert Link Tag"
         Top             =   55
         Width           =   330
      End
      Begin VB.CommandButton cmdButtonBar 
         Height          =   300
         Index           =   6
         Left            =   7360
         Picture         =   "mdiMain.frx":17C2
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Insert Font"
         Top             =   55
         Width           =   330
      End
      Begin VB.CommandButton cmdButtonBar 
         Height          =   300
         Index           =   5
         Left            =   2050
         Picture         =   "mdiMain.frx":190C
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Paste Text"
         Top             =   55
         Width           =   330
      End
      Begin VB.CommandButton cmdButtonBar 
         Height          =   300
         Index           =   4
         Left            =   1680
         Picture         =   "mdiMain.frx":1A56
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Copy Text"
         Top             =   55
         Width           =   330
      End
      Begin VB.CommandButton cmdButtonBar 
         Height          =   300
         Index           =   3
         Left            =   1300
         Picture         =   "mdiMain.frx":1BA0
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Cut Text"
         Top             =   55
         Width           =   330
      End
      Begin VB.CommandButton cmdButtonBar 
         Height          =   300
         Index           =   2
         Left            =   775
         Picture         =   "mdiMain.frx":1CEA
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Save project to file..."
         Top             =   55
         Width           =   330
      End
      Begin VB.CommandButton cmdButtonBar 
         Height          =   300
         Index           =   1
         Left            =   400
         Picture         =   "mdiMain.frx":1E34
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Open New Project"
         Top             =   55
         Width           =   330
      End
      Begin VB.CommandButton cmdButtonBar 
         Height          =   300
         Index           =   0
         Left            =   25
         Picture         =   "mdiMain.frx":1F7E
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "New Project"
         Top             =   55
         Width           =   330
      End
      Begin VB.Frame Frame1 
         Height          =   25
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   10335
      End
      Begin VB.PictureBox picUndoContainer 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   2550
         ScaleHeight     =   315
         ScaleWidth      =   1305
         TabIndex        =   1
         Top             =   55
         Width           =   1305
         Begin VB.CommandButton cmdRedo 
            Height          =   255
            Left            =   700
            MaskColor       =   &H000000C0&
            Picture         =   "mdiMain.frx":20C8
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Redo Last Undo"
            Top             =   30
            Width           =   330
         End
         Begin VB.ComboBox cboRedo 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   670
            Style           =   2  'Dropdown List
            TabIndex        =   2
            ToolTipText     =   "Redo"
            Top             =   0
            Width           =   630
         End
         Begin VB.CommandButton cmdUndo 
            Height          =   255
            Left            =   30
            Picture         =   "mdiMain.frx":2212
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Undo Last Event"
            Top             =   30
            Width           =   330
         End
         Begin VB.ComboBox cboUndo 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   0
            Style           =   2  'Dropdown List
            TabIndex        =   3
            ToolTipText     =   "Undo"
            Top             =   0
            Width           =   630
         End
      End
      Begin VB.Frame Frame3 
         Height          =   385
         Index           =   0
         Left            =   1200
         TabIndex        =   12
         Top             =   -25
         Width           =   40
      End
      Begin VB.Frame Frame3 
         Height          =   385
         Index           =   1
         Left            =   2460
         TabIndex        =   16
         Top             =   -25
         Width           =   40
      End
      Begin VB.Frame Frame3 
         Height          =   385
         Index           =   2
         Left            =   3900
         TabIndex        =   17
         Top             =   -25
         Width           =   40
      End
      Begin VB.Frame Frame3 
         Height          =   385
         Index           =   3
         Left            =   5160
         TabIndex        =   22
         Top             =   -25
         Width           =   40
      End
      Begin VB.Frame Frame3 
         Height          =   385
         Index           =   4
         Left            =   6400
         TabIndex        =   26
         Top             =   -25
         Width           =   40
      End
      Begin VB.Frame Frame3 
         Height          =   385
         Index           =   5
         Left            =   7280
         TabIndex        =   27
         Top             =   -25
         Width           =   40
      End
      Begin VB.Frame Frame3 
         Height          =   385
         Index           =   6
         Left            =   8880
         TabIndex        =   35
         Top             =   -25
         Width           =   40
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   10260
      TabIndex        =   8
      Top             =   375
      Width           =   10260
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   10260
      TabIndex        =   9
      Top             =   375
      Width           =   10260
   End
   Begin VB.Menu mnu_File 
      Caption         =   "&File"
      Begin VB.Menu mnu_New 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnu_Open 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnu_Close 
         Caption         =   "&Close"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu_line1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Save 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu mnu_saveas 
         Caption         =   "&Save As..."
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnu_SaveAll 
         Caption         =   "Save A&ll"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnu_revert 
         Caption         =   "&Revert"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu3_line2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_printsetup 
         Caption         =   "Print &Setup"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu_print 
         Caption         =   "&Print"
         Enabled         =   0   'False
         Shortcut        =   ^P
      End
      Begin VB.Menu mnu_Line4 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_documents 
         Caption         =   "<blank>"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu mnu_line3 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnu_edit 
      Caption         =   "&Edit"
      Begin VB.Menu mnu_undo 
         Caption         =   "&Undo"
         Enabled         =   0   'False
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnu_redo 
         Caption         =   "&Redo"
         Enabled         =   0   'False
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnu2_line1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_find 
         Caption         =   "&Find"
         Enabled         =   0   'False
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnu_findnext 
         Caption         =   "Find &Next"
         Enabled         =   0   'False
         Shortcut        =   ^{F5}
      End
      Begin VB.Menu mnu_replace 
         Caption         =   "&Replace"
         Enabled         =   0   'False
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnu_replacenext 
         Caption         =   "Replace N&ext"
         Enabled         =   0   'False
         Shortcut        =   ^{F8}
      End
      Begin VB.Menu mnu2_line4 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_cut 
         Caption         =   "Cu&t"
         Enabled         =   0   'False
         Shortcut        =   ^X
      End
      Begin VB.Menu mnu_copy 
         Caption         =   "&Copy"
         Enabled         =   0   'False
         Shortcut        =   ^C
      End
      Begin VB.Menu mnu_Paste 
         Caption         =   "&Paste"
         Enabled         =   0   'False
         Shortcut        =   ^V
      End
      Begin VB.Menu mnu2_line2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_gotoline 
         Caption         =   "&Goto Line"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnu_script 
      Caption         =   "&Script"
      Begin VB.Menu mnu_CommentBlock 
         Caption         =   "&Comment Block"
      End
      Begin VB.Menu mnu_uncommentblock 
         Caption         =   "&Uncomment Block"
      End
      Begin VB.Menu mnu_line2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_printline 
         Caption         =   "Print Line"
      End
      Begin VB.Menu mnu_Text 
         Caption         =   "&Text"
         Begin VB.Menu mnu_Bold 
            Caption         =   "&Bold"
         End
         Begin VB.Menu mnu_Italic 
            Caption         =   "&Italic"
         End
         Begin VB.Menu mnu_Underline 
            Caption         =   "&Underline"
         End
      End
      Begin VB.Menu mnu_insert 
         Caption         =   "&Insert Tags"
         Begin VB.Menu mnu_LeftTag 
            Caption         =   "&Left"
         End
         Begin VB.Menu mnu_CenterTag 
            Caption         =   "&Center"
         End
         Begin VB.Menu mnu_RightTag 
            Caption         =   "&Right"
         End
         Begin VB.Menu line1 
            Caption         =   "-"
         End
         Begin VB.Menu mnu_inserturl 
            Caption         =   "&URL"
         End
         Begin VB.Menu mnu_insertfont 
            Caption         =   "&Font"
         End
         Begin VB.Menu mnu_insertpicture 
            Caption         =   "&Picture"
         End
      End
   End
   Begin VB.Menu mnu_View 
      Caption         =   "&View"
      Begin VB.Menu mnu_toolbar 
         Caption         =   "&Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnu_statusbar 
         Caption         =   "&Status Bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnu_linenumbers 
         Caption         =   "&Line Numbers"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnu5_line1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_settings 
         Caption         =   "&Settings"
      End
   End
   Begin VB.Menu mnu_Window 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnu_tilehorizontally 
         Caption         =   "Tile &Horizontally"
      End
      Begin VB.Menu mnu_tilevertically 
         Caption         =   "Tile &Vertically"
      End
      Begin VB.Menu mnu_cascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnu_arrangeicons 
         Caption         =   "Arrange &Icons"
      End
   End
   Begin VB.Menu about_mnu 
      Caption         =   "&About"
      Begin VB.Menu ElucidSoftwareWebpage 
         Caption         =   "Elucid Software Webpage"
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const cFilters As String = "All CGI(*.pl *.cgi)|*.pl;*.cgi|Cgi(*.cgi)|*.cgi|Perl(*.pl)|*.pl|All Files(*.*)|*.*"
' vHKEY_LOCAL_MACHINE, "\Software\" & App.Title & "\Settings", "Proxy IP", ""
Public WindowCount As Integer
Private Document(255) As frmMain

'These are for the stupid registry shit that pisses me off
Public varUndoLimit As Integer
Public varClearUndo As Boolean
Public varDocuments As Integer
Public varUseRelative As Boolean
Public varDefaultFolder As String
Public varShowLines As Boolean
Private Sub SaveWindowSettings()
  Dim CReg As New CRegister
    Set CReg = New CRegister
    CReg.REGSaveSetting vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\", "Window State", Me.WindowState
    CReg.REGSaveSetting vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\", "Window LeftPos", Me.Left
    CReg.REGSaveSetting vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\", "Window TopPos", Me.Top
    Set CReg = Nothing
End Sub

Private Sub LoadWindowSettings()
  Dim CReg As New CRegister
    Set CReg = New CRegister

    If CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\", "Window State", vbMaximized) = 2 Then
        Me.WindowState = vbMaximized
        Exit Sub
      Else
        Me.WindowState = CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\", "Window State", vbMaximized)
    End If

    If CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\", "Window LeftPos", Me.Left) < 0 Then
        Me.Left = 0
      Else
        Me.Left = CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\", "Window LeftPos", Me.Left)
    End If

    If CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\", "Window TopPos", Me.Top) < 0 Then
        Me.Top = 0
      Else
        Me.Top = CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\", "Window TopPos", Me.Top)
    End If

    Set CReg = Nothing
End Sub

Private Sub cboRedo_DropDown()
  Dim i As Long
    cboRedo.Clear
    For i = ActiveForm.RedoStack.Count To 1 Step -1
        cboRedo.AddItem "Redo " & ActiveForm.GetUndoText(ActiveForm.RedoStack(i).ModifyType)
    Next
End Sub

Private Sub cboUndo_Click()
  Dim i As Long
    For i = ActiveForm.UndoStack.Count To (ActiveForm.UndoStack.Count - cboUndo.ListIndex) Step -1
        Call cmdUndo_Click
    Next
End Sub

Private Sub cboUndo_DropDown()
  Dim i As Long
    cboUndo.Clear
    For i = ActiveForm.UndoStack.Count To 1 Step -1
        cboUndo.AddItem "Undo " & ActiveForm.GetUndoText(ActiveForm.UndoStack(i).ModifyType)
    Next
End Sub

Private Sub cmdButtonBar_Click(Index As Integer)

    If Index > 1 Then
        If ActiveForm Is Nothing Then Exit Sub
    End If

    Select Case Index
      Case 0 'New Button
        mnu_New_Click
      Case 1 'Open Button
        mnu_Open_Click
      Case 2 'Save Button
        If ActiveForm.txtChanged = -1 Then
            If Len(ActiveForm.txtMain.FileName) > 0 Then
                mnu_save_Click
              Else
                mnu_saveas_Click
            End If
        End If
      Case 3 'Cut Button
        mnu_cut_Click
      Case 4 'Copy Button
        mnu_copy_Click
      Case 5 'Paste Button
        mnu_Paste_Click
      Case 6 'Font
        frmFont.Show , Me
      Case 7 'URL
        frmInsertURL.Show , Me

      Case 8 'Image
        frmInsertImage.Show , Me
      Case 9 'Comment Code
        ActiveForm.CommentBlock
      Case 10 'Left Align
        ActiveForm.InsertTag "<p align=""left"" >", "</p>"
      Case 11 'Center Align
        ActiveForm.InsertTag "<p align=""center"" >", "</p>"
      Case 12 'Right Align
        ActiveForm.InsertTag "<p align=""right"" >", "</p>"
      Case 13 'Uncomment Code
        ActiveForm.UncommentBlock
      Case 14 'Find
        frmFind.Show , Me
        '      Case 15 'Find Next
      Case 16 'Bold
        ActiveForm.InsertTag "<B>", "</B>"
      Case 17 'Italic
        ActiveForm.InsertTag "<I>", "</I>"
      Case 18 'Underline
        ActiveForm.InsertTag "<U>", "</U>"
    End Select

End Sub

Private Sub cmdButtonBar_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    txtStatus.Caption = cmdButtonBar(Index).ToolTipText
End Sub

Private Sub ElucidSoftwareWebpage_Click()
OpenIt "http://elucidsoftware.hypermart.net"
End Sub

Private Sub mnu_arrangeicons_Click()
    Me.Arrange vbArrangeIcons
End Sub

Private Sub mnu_Bold_Click()
    Call cmdButtonBar_Click(16)
End Sub

Private Sub mnu_cascade_Click()
    Me.Arrange vbCascade
End Sub

Private Sub mnu_CenterTag_Click()
    cmdButtonBar_Click (11)
End Sub

Private Sub mnu_CommentBlock_Click()
    cmdButtonBar_Click (9)
End Sub

Private Sub mnu_copy_Click()
    SendMessageLong ActiveForm.txtMain.hwnd, WM_COPY, 0, 0
End Sub

Private Sub mnu_cut_Click()
  Dim cUndo As clsUndo
    Set cUndo = New clsUndo

    cUndo.lStart = ActiveForm.txtMain.SelStart
    cUndo.sDelText = ActiveForm.txtMain.SelText
    SendMessageLong ActiveForm.txtMain.hwnd, WM_CUT, 0, 0
    ActiveForm.AddToUndoStack cUndo
End Sub

Private Sub mnu_documents_Click(Index As Integer)
    NewDocument
    ActiveForm.txtChanged = -2
    ActiveForm.txtMain.LoadFile mnu_documents(Index).Tag, rtfText
    ActiveForm.txtChanged = 0
End Sub

Private Sub mnu_gotoline_Click()
    frmGotoLine.Show 0, Me
End Sub
Private Sub mnu_insertfont_Click()
    Call cmdButtonBar_Click(6)
End Sub

Private Sub mnu_insertpicture_Click()
    Call cmdButtonBar_Click(8)
End Sub

Private Sub mnu_inserturl_Click()
    Call cmdButtonBar_Click(7)
End Sub

Private Sub mnu_Italic_Click()
    Call cmdButtonBar_Click(17)
End Sub

Private Sub mnu_LeftTag_Click()
    cmdButtonBar_Click (10)
End Sub

Private Sub mnu_linenumbers_Click()

    If ActiveForm.picLines.Visible = False Then
        ActiveForm.picLines.Visible = True
        mnu_linenumbers.Checked = True
        varShowLines = True
      Else
        ActiveForm.picLines.Visible = False
        mnu_linenumbers.Checked = False
        varShowLines = False
    End If
    ActiveForm.Form_Resize

  Dim CReg As New CRegister
    Set CReg = New CRegister
    CReg.REGSaveSetting vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "ShowLinenumbers", mnu_linenumbers.Checked
    Set CReg = Nothing

End Sub

Private Sub mnu_Paste_Click()
  Dim cUndo As New clsUndo
    Set cUndo = New clsUndo

    cUndo.lStart = ActiveForm.txtMain.SelStart
    SendMessageLong ActiveForm.txtMain.hwnd, WM_PASTE, 0, 0
    cUndo.sAddText = Clipboard.GetText(vbCFText)
    ActiveForm.AddToUndoStack cUndo
End Sub

Private Sub mnu_revert_Click()
  Dim YesNo As String
    YesNo = MsgBox("Are you sure you wish to revert back to last saved version of this file?" & vbCrLf & vbCrLf & "This will cause all unsaved changes to be lost!", vbYesNo, "Are you sure?")
    
    If YesNo = vbYes Then
        ActiveForm.ClearStack
        ActiveForm.txtMain.LoadFile ActiveForm.txtMain.FileName, rtfText
    End If

End Sub

Private Sub mnu_RightTag_Click()
    cmdButtonBar_Click (12)
End Sub

Private Sub mnu_settings_Click()
    frmSettings.Show 1, Me
End Sub

Private Sub mnu_statusbar_Click()

    If picStatusBar.Visible = True Then
        picStatusBar.Visible = False
        mnu_statusbar.Checked = False
      Else
        picStatusBar.Visible = True
        mnu_statusbar.Checked = True
    End If
    ActiveForm.Form_Resize
    
  Dim CReg As New CRegister
    Set CReg = New CRegister
    CReg.REGSaveSetting vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "ShowStatusbar", mnu_statusbar.Checked
    Set CReg = Nothing
End Sub

Private Sub mnu_tilehorizontally_Click()
    Me.Arrange vbHorizontal
End Sub

Private Sub mnu_tilevertically_Click()
    Me.Arrange vbVertical
End Sub

Private Sub mnu_toolbar_click()
  Dim CReg As New CRegister
    Set CReg = New CRegister

    If picToolbar.Visible = True Then

        picToolbar.Visible = False
        mnu_toolbar.Checked = False
      Else
        picToolbar.Visible = True
        mnu_toolbar.Checked = True
    End If
    ActiveForm.Form_Resize

    CReg.REGSaveSetting vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "ShowToolbar", mnu_toolbar.Checked
    Set CReg = Nothing
End Sub
Private Sub cboRedo_Click()
  Dim i As Long
    For i = ActiveForm.RedoStack.Count To (ActiveForm.RedoStack.Count - cboRedo.ListIndex) Step -1
        Call cmdRedo_Click
    Next
End Sub

Public Sub cmdRedo_Click()
  Dim cRedo As clsUndo
    ActiveForm.bRedoing = False
    If ActiveForm.RedoStack.Count = 0 Then Exit Sub
    '// get the current Redo item
    Set cRedo = ActiveForm.RedoStack(ActiveForm.RedoStack.Count)
    '// add it to the undo stack, and remove it from the Redo stack
    ActiveForm.UndoStack.Add cRedo
    ActiveForm.RedoStack.Remove (ActiveForm.RedoStack.Count)
    '// freeze updates
    LockWindowUpdate ActiveForm.txtMain.hwnd
    '// Redo the text edit
    ActiveForm.txtMain.SelStart = cRedo.lStart
    '// delete any text that was deleted
    ActiveForm.txtMain.SelLength = Len(cRedo.sDelText)
    '// replace the text that was added
    ActiveForm.txtMain.SelText = cRedo.sAddText
    LockWindowUpdate 0
    ActiveForm.bRedoing = True
    ActiveForm.txtMain.SetFocus
End Sub

Public Sub cmdUndo_Click()
  Dim cUndo As clsUndo
    
    If ActiveForm.UndoStack.Count = 0 Then Exit Sub
    '// get the current Undo item
    Set cUndo = ActiveForm.UndoStack(ActiveForm.UndoStack.Count)
    '// add it to the redo stack, and remove it from the Undo stack
    ActiveForm.RedoStack.Add cUndo
    ActiveForm.UndoStack.Remove (ActiveForm.UndoStack.Count)
    '// freeze updates
    LockWindowUpdate ActiveForm.txtMain.hwnd
    '// Undo the text edit
    ActiveForm.txtMain.SelStart = cUndo.lStart
    '// delete any text that was added
    ActiveForm.txtMain.SelLength = Len(cUndo.sAddText)
    '// replace the text that was deleted
    ActiveForm.txtMain.SelText = cUndo.sDelText

    LockWindowUpdate 0
    ' ActiveForm.txtmain.SetFocus

End Sub

Private Sub MDIForm_Resize()
    picToolbar.Move 0, 0, Me.Width, 375

End Sub

Private Sub mnu_Edit_Click()
    If ActiveForm Is Nothing Then Exit Sub

    If ActiveForm.RedoStack.Count > 1 Then mnu_redo.Enabled = True
    If ActiveForm.UndoStack.Count > 1 Then mnu_undo.Enabled = True

    If ActiveForm.CanPaste Then
        mnu_Paste.Enabled = True
      Else
        mnu_Paste.Enabled = False
    End If

    If Len(ActiveForm.txtMain.SelText) > 0 Then
        mnu_cut.Enabled = True
        mnu_copy.Enabled = True
      Else
        mnu_cut.Enabled = False
        mnu_copy.Enabled = False
    End If

    If Len(ActiveForm.txtMain.Text) > 0 Then
        mnu_gotoline.Enabled = True
        mnu_find.Enabled = True
        mnu_findnext.Enabled = True
        mnu_replace.Enabled = True
        mnu_replacenext.Enabled = True
    End If

End Sub
Private Sub mnu_File_Click()
    If Not ActiveForm Is Nothing Then
        '       mnu_rename.Enabled = True
        mnu_Close.Enabled = True
        mnu_printsetup.Enabled = True
        mnu_print.Enabled = True
      Else
        Exit Sub
    End If

    If ActiveForm.txtChanged = -1 Then
        If Len(ActiveForm.txtMain.FileName) > 0 Then
            mnu_Save.Enabled = True
            mnu_revert.Enabled = True
            '            mnu_rename.Enabled = True
          Else
            mnu_Save.Enabled = False
            
            mnu_revert.Enabled = False
            '            mnu_rename.Enabled = False
        End If
        
        mnu_saveas.Enabled = True
      Else
        mnu_Save.Enabled = False
        mnu_saveas.Enabled = False
    End If
    
    If WindowCount > 1 Then
        mnu_SaveAll.Enabled = True
      Else
        mnu_SaveAll.Enabled = False
    End If

  Dim CReg As New CRegister
    Set CReg = New CRegister

End Sub

Private Sub MDIForm_Load()
    LoadSettings
    LoadWindowSettings
    picToolbar.Move 0, 0, Me.Width, 375
    Frame1.Left = -100
    Frame1.Width = Screen.Width
    txtStatus.Width = picStatusBar.Width
    NewDocument
    'MsgBox "This is a preview release and is only for testing purposes.  We need you to submit any bugs you can find and features you would like to see.  As of now we are aware that some of the features are inoperable.  This is becuase we felt it would be better to get your opinion on how to implement them.", vbInformation, "Preview Release"
    GetRecentList
    GetKeywords
End Sub
Public Sub GotoLine(LineNumber As Long, Highlight As Boolean)
    ActiveForm.GotoLine LineNumber, Highlight
End Sub
Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveWindowSettings
End Sub

Private Sub mnu_Close_Click()
    Unload ActiveForm

End Sub
Private Sub mnu_New_Click()
    NewDocument
End Sub
Private Sub NewDocument()
  Dim Index As Byte
    Index = UBound(Document)
    Set Document(Index) = New frmMain
    Document(Index).txtChanged = tsFalse
    Document(Index).Show
End Sub
Private Sub mnu_Open_Click()
  Dim CmdDlg As New cCommonDialog
    Set CmdDlg = New cCommonDialog
    CmdDlg.Filter = cFilters
    CmdDlg.DialogTitle = "Open Script"
    CmdDlg.hwnd = Me.hwnd
    CmdDlg.flags = OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST
    CmdDlg.FileTitle = mdiMain.varDefaultFolder
    CmdDlg.ShowOpen
    If CmdDlg.FileName = "" Then Exit Sub
    If ActiveForm Is Nothing Then
        NewDocument
      ElseIf ActiveForm.txtChanged = True Or Len(ActiveForm.txtMain.Text) > 0 Then
        NewDocument
    End If
    ActiveForm.txtChanged = tsnone
    ActiveForm.Caption = CmdDlg.FileTitle
    ActiveForm.txtMain.LoadFile CmdDlg.FileName, rtfText
    Colorize ActiveForm.txtMain, &H8000&, &HFF0000, &H80&, True
    ActiveForm.Caption = FileGetName(CmdDlg.FileName)
    ActiveForm.txtChanged = tsFalse
    AddRecentList CmdDlg.FileName
    Set CmdDlg = Nothing

End Sub

Private Sub mnu_print_Click()
    On Error Resume Next
    Dim c As New cCommonDialog
      With c
          .DialogTitle = "Choose Printer"
          .hwnd = Me.hwnd
          .PrinterDefault = True
          .Object = Printer
          .ShowPrinter
          ActiveForm.txtMain.SelPrint Printer.hdc
      End With
End Sub

Private Sub mnu_printsetup_Click()
    On Error Resume Next
    Dim c As New cCommonDialog
      With c
          .DialogTitle = "Choose Printer"
          .hwnd = Me.hwnd
          .PrinterDefault = True
          .Object = Printer
          .flags = PD_PRINTSETUP
          .ShowPrinter
      End With
End Sub

Private Sub mnu_redo_Click()
    cmdRedo_Click
End Sub

Private Sub mnu_save_Click()

  Dim hFile As Long
    hFile = FreeFile

    Open ActiveForm.txtMain.FileName For Output As hFile
    Print #hFile, ActiveForm.txtMain.Text
    Close

    ActiveForm.Caption = FileGetName(ActiveForm.txtMain.FileName)
    ActiveForm.txtChanged = tsFalse
    If Me.varClearUndo = True Then ActiveForm.ClearUndoRedo
End Sub

Private Sub mnu_saveas_Click()
    ShowSave
End Sub
Public Sub ShowSave()
  Dim CmdDlg As New cCommonDialog
    Set CmdDlg = New cCommonDialog
    CmdDlg.Filter = cFilters

    If ActiveForm.txtMain.FileName = "" Then

        If Left(ActiveForm.Caption, 1) = "*" Then
            CmdDlg.FileName = Right(ActiveForm.Caption, Len(ActiveForm.Caption) - 1)
          Else
            CmdDlg.FileName = ActiveForm.Caption
        End If
      Else
        CmdDlg.FileName = ActiveForm.txtMain.FileName
    End If
    CmdDlg.FileTitle = mdiMain.varDefaultFolder
    CmdDlg.DialogTitle = "Save Script"
    CmdDlg.flags = OFN_OVERWRITEPROMPT
    CmdDlg.hwnd = Me.hwnd
    CmdDlg.ShowSave
    If CmdDlg.FileName = "" Then Exit Sub
'    ActiveForm.txtMain.Text = Replace(ActiveForm.txtMain.Text, vbCrLf, vbCr)
    Dim hFile As Long
    hFile = FreeFile

    Open CmdDlg.FileName For Output As hFile
    Print #hFile, ActiveForm.txtMain.Text
    Close

    ActiveForm.Caption = FileGetName(CmdDlg.FileName)
    AddRecentList CmdDlg.FileName
    ActiveForm.txtChanged = tsFalse
    If Me.varClearUndo = True Then ActiveForm.ClearUndoRedo
    Set CmdDlg = Nothing
End Sub

Private Sub mnu_uncommentblock_Click()
    cmdButtonBar_Click (13)
End Sub

Private Sub mnu_Underline_Click()
    Call cmdButtonBar_Click(18)
End Sub

Private Sub mnu_undo_Click()
    cmdUndo_Click
End Sub

Private Sub mnu_Windows_Click(Index As Integer)

End Sub

Private Sub mnu_View_Click()

    If picToolbar.Visible = True Then
        mnu_toolbar.Checked = True
      Else
        mnu_toolbar.Checked = False
    End If

    If picStatusBar.Visible = True Then
        mnu_statusbar.Checked = True
      Else
        mnu_statusbar.Checked = False
    End If

    If ActiveForm Is Nothing Then Exit Sub
    If ActiveForm.picLines.Visible = True Then
        mnu_linenumbers.Checked = True
      Else
        mnu_linenumbers.Checked = False
    End If
End Sub
Public Sub ClearRecentList()
  Dim i  As Integer

    For i = 1 To mnu_documents.UBound
        Unload mnu_documents(i)
    Next i

End Sub
Public Sub AddRecentList(FileToAdd As String)
  Dim a As Integer
  Dim sBuf As Variant

  Dim b As Integer
    a = FreeFile

    If FileCheck(App.Path & "\recent.lst") = True Then

        If FileLen(App.Path & "\recent.lst") > 0 Then
            Open App.Path & "\recent.lst" For Input As #a
            sBuf = Input(LOF(a), #a)
            Close #a
        End If

    End If

    b = FreeFile

    Open App.Path & "\recent.lst" For Output As #b
    sBuf = FileToAdd & vbCrLf & sBuf
    Print #b, sBuf
    Close b

    ClearRecentList
    GetRecentList
End Sub
Public Sub GetRecentList()
  Dim a As Integer
  Dim sBuf As String
  Dim NewIndex As Integer
  Dim Count As Integer

    a = FreeFile
    If FileCheck(App.Path & "\recent.lst") = False Then Exit Sub
    Open App.Path & "\recent.lst" For Input As #a

  Dim i As Integer
    For i = 1 To Me.varDocuments

        If EOF(a) Or i > Me.varDocuments Then GoTo closeit:
        If i > 0 Then mnu_documents(0).Visible = False
        Line Input #a, sBuf
        If FileCheck(sBuf) = False Then GoTo skipit:
        NewIndex = mnu_documents.UBound + 1
        Load mnu_documents(NewIndex)
        mnu_documents(NewIndex).Tag = sBuf 'keep the entire path in the tag in case its trimmed
        If Len(sBuf) > 35 Then
            sBuf = "..." & Right(sBuf, 32)  'make sure this thing isnt to long
        End If
        Count = Count + 1
        mnu_documents(NewIndex).Caption = "&" & Count & " " & sBuf
        mnu_documents(NewIndex).Enabled = True
        mnu_documents(NewIndex).Visible = True
skipit:
    Next i
closeit:
    Close a

End Sub
Public Sub InsertFont(Face As String, Size As Integer, Style As String, lColor As String)
    ActiveForm.InsertFont Face, Size, Style, lColor
End Sub
Public Sub SaveSettings()

End Sub
Public Sub LoadSettings()
  Dim CReg As New CRegister
    Set CReg = New CRegister
    varUndoLimit = CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "UndoLimit", 100)
    varClearUndo = CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "ClearUndoSave", vbUnchecked)
    varDocuments = Int(CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "Documents", 4))
    varUseRelative = CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "UseRelative", vbChecked)
    varDefaultFolder = CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "DefaultFolder", App.Path)
    picToolbar.Visible = CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "ShowToolbar", True)
    picStatusBar.Visible = CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "ShowStatusbar", True)
    varShowLines = CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "ShowLinenumbers", True)

    mnu_toolbar.Checked = CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "ShowToolbar", True)
    mnu_statusbar.Checked = CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "ShowStatusbar", True)
    mnu_linenumbers.Checked = CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "ShowLinenumbers", True)
    Set CReg = Nothing
End Sub
Public Sub LineNumbers(IsVisible As Boolean)
    ActiveForm.picLines.Visible = IsVisible
    ActiveForm.Form_Resize
End Sub
Private Sub picToolbar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    txtStatus.Caption = "Thank you for using CgEye By Elucid Software"
End Sub
Private Sub GetKeywords()
  Dim Num%, Buf$
    Num = FreeFile
    Keywords = "|"
    If FileCheck(App.Path & "\keywords.dat") Then
        Open App.Path & "\keywords.dat" For Input As #Num
        While Not EOF(Num)
            Line Input #Num, Buf$
            Keywords = Keywords & Buf$ & "|"
        Wend
        Close #Num
    End If

End Sub
