VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9660
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   9660
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame ctSyntax 
      Caption         =   "Syntax Highlighting (This code is in extreme testing phase.)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   2760
      TabIndex        =   42
      Top             =   120
      Width           =   6735
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Left            =   240
         TabIndex        =   60
         Top             =   360
         Width           =   1935
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   4440
         ScaleHeight     =   345
         ScaleWidth      =   1185
         TabIndex        =   58
         Top             =   2160
         Width           =   1215
      End
      Begin VB.ListBox lstKeywords 
         Height          =   1035
         Left            =   3840
         TabIndex        =   45
         Top             =   720
         Width           =   1695
      End
      Begin VB.Frame Frame4 
         Height          =   25
         Left            =   600
         TabIndex        =   52
         Top             =   1920
         Width           =   6015
      End
      Begin VB.PictureBox picShebang 
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1320
         ScaleHeight     =   345
         ScaleWidth      =   1185
         TabIndex        =   51
         Top             =   2160
         Width           =   1215
      End
      Begin VB.PictureBox picComments 
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1320
         ScaleHeight     =   345
         ScaleWidth      =   1185
         TabIndex        =   50
         Top             =   2760
         Width           =   1215
      End
      Begin VB.PictureBox picKeywords 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   4440
         ScaleHeight     =   345
         ScaleWidth      =   1185
         TabIndex        =   49
         Top             =   2760
         Width           =   1215
      End
      Begin VB.CheckBox chkShebang 
         Caption         =   "Shebang Line"
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
         Left            =   2160
         TabIndex        =   48
         Top             =   840
         Width           =   1335
      End
      Begin VB.CheckBox chkComments 
         Caption         =   "Comments"
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
         Left            =   720
         TabIndex        =   47
         Top             =   840
         Width           =   1095
      End
      Begin VB.CheckBox chkKeywords 
         Caption         =   "Keywords"
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
         Left            =   3840
         TabIndex        =   46
         Top             =   360
         Width           =   2055
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Add"
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
         Left            =   5640
         TabIndex        =   44
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton Command6 
         Caption         =   "&Remove"
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
         Left            =   5640
         TabIndex        =   43
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label18 
         Caption         =   "Syntax highlighting takes about 2 - 3 seconds on opening of script to process for every 600 lines."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   59
         Top             =   1320
         Width           =   3495
      End
      Begin VB.Label Label19 
         Caption         =   "String Text:"
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
         Left            =   3000
         TabIndex        =   57
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Shebang Line:"
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
         Left            =   240
         TabIndex        =   56
         Top             =   2160
         Width           =   1020
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Colors"
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
         TabIndex        =   55
         Top             =   1800
         Width           =   450
      End
      Begin VB.Label Label16 
         Caption         =   "Comments:"
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
         Left            =   240
         TabIndex        =   54
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label Label17 
         Caption         =   "Script Keywords:"
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
         Left            =   3000
         TabIndex        =   53
         Top             =   2760
         Width           =   1335
      End
   End
   Begin VB.Frame ctGeneral 
      Caption         =   "General"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   2760
      TabIndex        =   2
      Top             =   120
      Width           =   6735
      Begin VB.TextBox txtUndo 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1080
         TabIndex        =   38
         Text            =   "100"
         Top             =   1680
         Width           =   495
      End
      Begin VB.Frame Frame3 
         Height          =   25
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   6495
      End
      Begin VB.TextBox txtDocuments 
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
         Left            =   1440
         TabIndex        =   10
         Text            =   "4"
         Top             =   2880
         Width           =   375
      End
      Begin VB.CheckBox chkClearUndo 
         Caption         =   "Clear undo buffer on save"
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
         Left            =   600
         TabIndex        =   9
         Top             =   2040
         Width           =   2655
      End
      Begin VB.CheckBox chkRelative 
         Caption         =   "Always use relative paths"
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
         Left            =   1320
         TabIndex        =   8
         Top             =   840
         Width           =   2655
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Select"
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
         Left            =   5160
         TabIndex        =   7
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtDefaultFolder 
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
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   360
         Width           =   5535
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Recent Documents:"
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
         TabIndex        =   41
         Top             =   2520
         Width           =   1410
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Undo/Redo Settings:"
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
         TabIndex        =   40
         Top             =   1320
         Width           =   1500
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "time(s) before clearing old undo's"
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
         Left            =   1680
         TabIndex        =   39
         Top             =   1680
         Width           =   2385
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Undo"
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
         Left            =   600
         TabIndex        =   37
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "recent documents in file menu"
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
         Left            =   1920
         TabIndex        =   13
         Top             =   2880
         Width           =   2775
      End
      Begin VB.Label Label2 
         Caption         =   "Allow up to"
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
         Left            =   600
         TabIndex        =   12
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Root Folder:"
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
         TabIndex        =   5
         Top             =   360
         Width           =   900
      End
   End
   Begin VB.Frame ctHotkeys 
      Caption         =   "Hotkeys"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   2760
      TabIndex        =   27
      Top             =   120
      Width           =   6735
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmSettings.frx":06EA
         Left            =   120
         List            =   "frmSettings.frx":06F4
         TabIndex        =   32
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox Text4 
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
         TabIndex        =   31
         Top             =   1320
         Width           =   2415
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Assign &Key"
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
         Left            =   1440
         TabIndex        =   30
         Top             =   1680
         Width           =   1095
      End
      Begin VB.ListBox List2 
         Columns         =   2
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         ItemData        =   "frmSettings.frx":0714
         Left            =   2880
         List            =   "frmSettings.frx":0716
         TabIndex        =   29
         Top             =   600
         Width           =   3735
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Top             =   2520
         Width           =   6495
      End
      Begin VB.Label Label9 
         Caption         =   "Key Mask"
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
         TabIndex        =   36
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label10 
         Caption         =   "Trigger Key:"
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
         TabIndex        =   35
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label28 
         Caption         =   "Assigned Keys"
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
         TabIndex        =   34
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Text To Assign Key To"
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
         TabIndex        =   33
         Top             =   2280
         Width           =   1605
      End
   End
   Begin VB.Frame ctAppearance 
      Caption         =   "Appearance"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   2760
      TabIndex        =   14
      Top             =   120
      Width           =   6735
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   360
         ScaleHeight     =   255
         ScaleWidth      =   5535
         TabIndex        =   24
         Top             =   1320
         Width           =   5535
         Begin VB.OptionButton chkStatusbar2 
            Caption         =   "Never Show"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2160
            TabIndex        =   26
            Top             =   0
            Width           =   1935
         End
         Begin VB.OptionButton chkStatusbar1 
            Caption         =   "Always Show"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   25
            Top             =   0
            Value           =   -1  'True
            Width           =   1815
         End
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   480
         ScaleHeight     =   255
         ScaleWidth      =   5775
         TabIndex        =   20
         Top             =   2040
         Width           =   5775
         Begin VB.OptionButton chkLinenumbers2 
            Caption         =   "Never Show"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2040
            TabIndex        =   22
            Top             =   0
            Width           =   1695
         End
         Begin VB.OptionButton chkLinenumbers1 
            Caption         =   "Always Show"
            BeginProperty Font 
               Name            =   "Verdana"
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
            Top             =   0
            Value           =   -1  'True
            Width           =   1935
         End
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   360
         ScaleHeight     =   255
         ScaleWidth      =   4935
         TabIndex        =   16
         Top             =   600
         Width           =   4935
         Begin VB.OptionButton chkToolbar2 
            Caption         =   "Never Show"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2160
            TabIndex        =   18
            Top             =   0
            Width           =   1695
         End
         Begin VB.OptionButton chkToolbar1 
            Caption         =   "Always Show"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   17
            Top             =   0
            Value           =   -1  'True
            Width           =   1455
         End
      End
      Begin VB.Label Label8 
         Caption         =   "Status Bar:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   23
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Line Numbers:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   19
         Top             =   1800
         Width           =   1245
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Toolbar:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   15
         Top             =   360
         Width           =   720
      End
   End
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
      Left            =   6600
      TabIndex        =   4
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Apply"
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
      Left            =   8040
      TabIndex        =   3
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Section"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2985
         ItemData        =   "frmSettings.frx":0718
         Left            =   120
         List            =   "frmSettings.frx":0725
         TabIndex        =   1
         Top             =   240
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub SaveSettings()
  Dim CReg As New CRegister
    Set CReg = New CRegister
    CReg.REGSaveSetting vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "UndoLimit", txtUndo
    CReg.REGSaveSetting vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "ClearUndoSave", chkClearUndo.Value
    CReg.REGSaveSetting vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "Documents", Int(txtDocuments.Text)
    CReg.REGSaveSetting vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "UseRelative", chkRelative.Value
    CReg.REGSaveSetting vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "DefaultFolder", txtDefaultFolder.Text
    CReg.REGSaveSetting vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "ShowToolbar", chkToolbar1.Value
    CReg.REGSaveSetting vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "ShowStatusbar", chkStatusbar1.Value
    CReg.REGSaveSetting vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "ShowLinenumbers", chkLinenumbers1.Value
    Set CReg = Nothing
End Sub

Public Sub LoadSettings()
  Dim CReg As New CRegister
    Set CReg = New CRegister
    txtUndo = CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "UndoLimit", 100)
    chkClearUndo.Value = CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "ClearUndoSave", vbUnchecked)
    txtDocuments.Text = Int(CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "Documents", 4))
    chkRelative.Value = CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "UseRelative", vbChecked)
    txtDefaultFolder.Text = CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "DefaultFolder", App.Path)

    If CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "ShowToolbar", True) = True Then
        chkToolbar1.Value = True
      Else
        chkToolbar2.Value = True
    End If

    If CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "ShowStatusbar", True) = True Then
        chkStatusbar1.Value = True
      Else
        chkStatusbar2.Value = True
    End If

    If CReg.REGGetSetting(vHKEY_LOCAL_MACHINE, "\Software\ElucidSoftware\CgEye\Settings", "ShowLinenumbers", True) = True Then
        chkLinenumbers1.Value = True
      Else
        chkLinenumbers2.Value = True
    End If

    Set CReg = Nothing
End Sub

Private Sub Check3_Click()

End Sub

Private Sub Command1_Click()
    SaveSettings

    If Not mdiMain.varDocuments = Int(txtDocuments.Text) Then
        mdiMain.varDocuments = Int(txtDocuments.Text)
        mdiMain.ClearRecentList
        mdiMain.GetRecentList
    End If
    mdiMain.varDefaultFolder = txtDefaultFolder.Text
    Unload Me
    mdiMain.SetFocus
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command5_Click()
    MsgBox chkStatusbar2.Value
End Sub

Private Sub Command3_Click()
    frmFolder.Show 1, Me
End Sub

Private Sub Form_Load()
    LoadSettings
    SetNumber txtUndo, True
    SetNumber txtDocuments, True
    ctGeneral.ZOrder 0
    GetKeywords

End Sub

Private Sub List1_Click()
    Select Case List1.ListIndex
      Case 0  ' General
        ctGeneral.ZOrder 0
        'ctGeneral.Visible = True
      Case 1 'Appearance
        ctAppearance.ZOrder 0
        'ctAppearance.Visible = True
      Case 2 'Syntax
        ctSyntax.ZOrder 0
    End Select
End Sub
Private Sub GetKeywords()
  Dim Num%, Buf$
    Num = FreeFile
    Keywords = "|"
    If FileCheck(App.Path & "\keywords.dat") Then
        Open App.Path & "\keywords.dat" For Input As #Num
        While Not EOF(Num)
            Line Input #Num, Buf$
            lstKeywords.AddItem Buf
        Wend
        Close #Num
    End If

End Sub
