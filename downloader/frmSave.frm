VERSION 5.00
Begin VB.Form frmSave 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Save File To...."
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   240
   ClientWidth     =   3870
   Icon            =   "frmSave.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   3870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   120
      TabIndex        =   9
      Top             =   2880
      Width           =   3615
   End
   Begin VB.CommandButton cmdCancel 
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
      Left            =   2400
      TabIndex        =   4
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton cmdOk 
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
      Left            =   120
      TabIndex        =   3
      Top             =   4560
      Width           =   1215
   End
   Begin VB.TextBox FilePath 
      Height          =   285
      Left            =   120
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   4200
      Width           =   3615
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3615
   End
   Begin VB.DirListBox Dir1 
      Height          =   1665
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   3615
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Current Files:"
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
      TabIndex        =   8
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Location Path To File:"
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
      TabIndex        =   7
      Top             =   3960
      Width           =   1545
   End
   Begin VB.Label Label2 
      Caption         =   "Select Folder:"
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
      TabIndex        =   6
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Select Drive:"
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
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    frmMain.Animation1.AutoPlay = False
    Unload frmSave
    
End Sub

Private Sub cmdOk_Click()
    FilePathName = FilePath.Text
    
    If FileCheck(FilePathName) Then
        frmExist.Show , frmSave
        
    Else
        frmMain.Winsock.Connect strSvrURL, 80
        Unload frmSave
        
    End If
    
End Sub

Private Sub Dir1_Change()
    FilePathName = Dir1.Path & "\" & Filename
    
    If InStr(FilePathName, "\\") Then 'this prevents from a double / if you goto the root of the drive
        FilePathName = Dir1.Path & Filename
        
    End If
    
    FilePath.Text = FilePathName
    File1.Path = Dir1.Path
    
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
    
End Sub

Private Sub File1_Click()
    FilePath = File1.Path & "\" & File1.Filename
    
End Sub

Private Sub Form_Load()
    'frmTool.commDialog.HelpFile = App.Path & "\Cub Scout.chm"
    
    FilePathName = Me.Dir1.Path & "\" & Filename
    
    If InStr(FilePathName, "\\") Then 'this prevents from a double / if you goto the root of the drive
        FilePathName = Me.Dir1.Path & Filename
    End If
    
    FilePath.Text = FilePathName
    
End Sub
