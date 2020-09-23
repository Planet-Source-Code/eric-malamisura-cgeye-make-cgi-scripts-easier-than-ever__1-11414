VERSION 5.00
Begin VB.Form frmFolder 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select Folder"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
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
      Left            =   2160
      TabIndex        =   3
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
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
      Left            =   3600
      TabIndex        =   2
      Top             =   4200
      Width           =   1335
   End
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3465
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   4695
   End
   Begin VB.DriveListBox Drive1 
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
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frmFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    If Not Len(Dir1.Path) = 3 Then
        frmSettings.txtDefaultFolder = Dir1.Path & "\"
      Else
        frmSettings.txtDefaultFolder = Dir1.Path
    End If
    
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

