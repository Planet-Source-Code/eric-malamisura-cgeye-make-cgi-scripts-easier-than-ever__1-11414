VERSION 5.00
Begin VB.Form frmInsertImage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Insert Image"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8835
   Icon            =   "frmInsertImage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   8835
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
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
      Left            =   6120
      TabIndex        =   12
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CheckBox chkRelative 
      Alignment       =   1  'Right Justify
      Caption         =   "Get Relative Path"
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
      Left            =   5400
      TabIndex        =   11
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "&Select File"
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
      Left            =   7320
      TabIndex        =   10
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox txtWidth 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   960
      TabIndex        =   9
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox txtHeight 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   960
      TabIndex        =   8
      Top             =   1080
      Width           =   855
   End
   Begin VB.CheckBox chkSize 
      Caption         =   "Specify Image Size:"
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
      TabIndex        =   5
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox txtBorder 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3840
      TabIndex        =   4
      Text            =   "0"
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   8655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Default         =   -1  'True
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
      Left            =   7440
      TabIndex        =   0
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Width:"
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
      Left            =   360
      TabIndex        =   7
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Height:"
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
      Left            =   360
      TabIndex        =   6
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Border Size:"
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
      Left            =   2880
      TabIndex        =   3
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Image File:"
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
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmInsertImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkRelative_Click()
  Dim CReg As New CRegister
    Set CReg = New CRegister
    CReg.REGSaveSetting vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "UseRelative", chkRelative.Value
    mdiMain.varUseRelative = chkRelative.Value
    Set CReg = Nothing
End Sub

Private Sub chkSize_Click()
    If chkSize.Value = Checked Then
        txtHeight.Enabled = True
        txtWidth.Enabled = True
        txtHeight.BackColor = &H80000005
        txtWidth.BackColor = &H80000005
      Else
        txtWidth.BackColor = &H8000000F
        txtHeight.BackColor = &H8000000F
        txtHeight.Enabled = False
        txtWidth.Enabled = False
    End If
End Sub

Private Sub Command1_Click()
    If chkSize.Value = Checked Then
        mdiMain.ActiveForm.InsertImage Text1, txtBorder, True, txtWidth, txtHeight
      Else
        mdiMain.ActiveForm.InsertImage Text1, txtBorder, False
    End If
    Unload Me
    mdiMain.SetFocus
End Sub

Private Sub Command2_Click()
    Unload Me
    mdiMain.SetFocus
End Sub

Private Sub Command3_Click()
  Dim CmdDlg As New cCommonDialog
    Set CmdDlg = New cCommonDialog

    CmdDlg.Filter = "All Images(*.jpg *.gif)|*.jpg;*.gif|Jpeg(*.jpg)|*.jpg|Gif(*.gif)|*.gif|All Files(*.*)|*.*"
    CmdDlg.DialogTitle = "Select Image"
    CmdDlg.FileTitle = mdiMain.varDefaultFolder
    CmdDlg.hwnd = mdiMain.hwnd
    CmdDlg.ShowOpen
    If CmdDlg.FileName = "" Then Exit Sub

  Dim RelativePath As String
  Dim Temp As String
  Dim i As Integer

    Text1.Text = returnRelPath(mdiMain.varDefaultFolder, CmdDlg.FileName)
End Sub

Private Sub Form_Load()

    If mdiMain.varUseRelative = True Then
        chkRelative.Value = Checked
    End If

    SetNumber txtHeight, True
    SetNumber txtWidth, True
    SetNumber txtBorder, True
End Sub
