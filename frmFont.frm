VERSION 5.00
Begin VB.Form frmFont 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Font"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6270
   Icon            =   "frmFont.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   6270
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command4 
      Caption         =   "&Reset"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   20
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   19
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   18
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtHex 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Select Color"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Height          =   25
      Left            =   480
      TabIndex        =   13
      Top             =   1920
      Width           =   2055
   End
   Begin VB.PictureBox picColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1440
      ScaleHeight     =   345
      ScaleWidth      =   1065
      TabIndex        =   12
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox txtPreview 
      Alignment       =   2  'Center
      Height          =   1215
      Left            =   120
      TabIndex        =   9
      Text            =   "Preview Text"
      Top             =   3120
      Width           =   6015
   End
   Begin VB.ListBox lstSize 
      Height          =   1035
      ItemData        =   "frmFont.frx":014A
      Left            =   4800
      List            =   "frmFont.frx":0163
      TabIndex        =   8
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox txtSize 
      Height          =   285
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "Normal"
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox txtStyle 
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "Regular"
      Top             =   360
      Width           =   1575
   End
   Begin VB.ListBox lstStyle 
      Height          =   1035
      ItemData        =   "frmFont.frx":01B3
      Left            =   2880
      List            =   "frmFont.frx":01C3
      TabIndex        =   4
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox txtFont 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "[Default Font]"
      Top             =   360
      Width           =   2415
   End
   Begin VB.ListBox lstFont 
      Height          =   1035
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   720
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Height          =   25
      Left            =   720
      TabIndex        =   10
      Top             =   3000
      Width           =   5415
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "You may type in preview box to see what certain character may look like!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   4320
      Width           =   6015
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Hex:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   2520
      Width           =   345
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Color"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "Preview"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Size:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Font Style:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2880
      TabIndex        =   3
      Top             =   120
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   " Font:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   435
   End
End
Attribute VB_Name = "frmFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Regular As Boolean
Dim Italic As Boolean
Dim Bold As Boolean
Dim BoldItalic As Boolean
Dim sFontName As String
Dim iFontsize As Integer
Dim FontStyle As String
Dim lColor As String

Private Sub Command1_Click()
  Dim CmdDlg As New cCommonDialog
    Set CmdDlg = New cCommonDialog
  Dim i As Byte
  Dim iLength As Integer
    CmdDlg.hwnd = Me.hwnd
    CmdDlg.ShowColor

    lColor = Hex$(Str(CmdDlg.Color))
    iLength = Len(lColor)
    If iLength = 5 Then lColor = "0" & lColor
    If iLength = 4 Then lColor = "00" & lColor
    If iLength = 3 Then lColor = "000" & lColor
    If iLength = 2 Then lColor = "0000" & lColor
    If iLength = 1 Then lColor = "00000" & lColor
  Dim One$, Two$, Three$

    One = Left(lColor, 2)
    Two = Mid(lColor, 3, 2)
    Three = Right(lColor, 2)

    lColor = Three & Two & One
    txtHex.Text = "#" & lColor

    picColor.BackColor = CmdDlg.Color
    txtPreview.ForeColor = CmdDlg.Color

End Sub

Private Sub Command2_Click()

    If Italic = True Then
        FontStyle = "Italic"
      ElseIf Bold = True Then
        FontStyle = "Bold"
      ElseIf BoldItalic = True Then
        FontStyle = "Bold Italic"
    End If

    mdiMain.InsertFont sFontName, iFontsize, FontStyle, txtHex.Text
    Unload Me

End Sub

Private Sub Command3_Click()
    Unload Me
    mdiMain.SetFocus
End Sub

Private Sub Command4_Click()
    Regular = False
    Italic = False
    Bold = False
    BoldItalic = False
    sFontName = ""
    iFontsize = 0
    FontStyle = ""
    lColor = ""
    lstStyle.ListIndex = -1
    lstFont.ListIndex = -1
    lstSize.ListIndex = -1
    txtPreview.Font.Bold = False
    txtPreview.Font.Italic = False
    txtPreview.FontName = "MS Sans Serif"
    txtPreview.Font.Size = 8
    txtPreview.ForeColor = vbBlack
    picColor.BackColor = vbBlack
    txtHex = ""
End Sub

Private Sub Form_Load()
  Dim i As Integer
    For i = 1 To Screen.FontCount
        lstFont.AddItem Screen.Fonts(i - 1)
    Next i

    Regular = False
    Italic = False
    Bold = False
    BoldItalic = False
    sFontName = ""
    iFontsize = 0
    FontStyle = ""
    lColor = ""
End Sub
Private Function FixLen(ByVal sIn As String, ByVal sMask As String) As String

    If Len(sIn) < Len(sMask) Then
        FixLen = Left$(sMask, Len(sMask) - Len(sIn)) & sIn
      Else
        FixLen = Right$(sIn, Len(sMask))
    End If
    
End Function

Private Sub lstFont_Click()
    txtFont.Text = lstFont.List(lstFont.ListIndex)
    sFontName = lstFont.List(lstFont.ListIndex)
    If lstFont.ListIndex >= 0 Then txtPreview.Font.Name = lstFont.List(lstFont.ListIndex)
End Sub

Private Sub lstSize_Click()
    iFontsize = lstSize.ListIndex + 1

    Select Case lstSize.ListIndex
      Case 0
        txtPreview.Font.Size = 8
      Case 1
        txtPreview.Font.Size = 10
      Case 2
        txtPreview.Font.Size = 12
      Case 3
        txtPreview.Font.Size = 14
      Case 4
        txtPreview.Font.Size = 18
      Case 5
        txtPreview.Font.Size = 24
      Case 6
        txtPreview.Font.Size = 36
    End Select

End Sub

Private Sub lstStyle_Click()
    Bold = False
    Italic = False
    BoldItalic = False
    Regular = False

    txtStyle.Text = lstStyle.List(lstStyle.ListIndex)
    Select Case lstStyle.List(lstStyle.ListIndex)
      Case "Regular"
        Regular = True

      Case "Italic"
        Italic = True
      Case "Bold"
        Bold = True
      Case "Bold Italic"
        BoldItalic = True
    End Select

    If Regular = False Then
        If BoldItalic = False Then
            txtPreview.Font.Bold = Bold
            txtPreview.Font.Italic = Italic
          Else
            txtPreview.Font.Bold = True
            txtPreview.Font.Italic = True
        End If
    
        If Italic = True Then
            txtPreview.Font.Bold = False
            txtPreview.Font.Italic = True
        End If

        If Bold = True Then
            txtPreview.Font.Italic = False
            txtPreview.Font.Bold = True
        End If

      Else

        If Regular = True Then
            txtPreview.Font.Bold = False
            txtPreview.Font.Italic = False
        End If

    End If

End Sub
