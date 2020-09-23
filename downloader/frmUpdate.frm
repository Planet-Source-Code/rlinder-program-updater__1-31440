VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Begin VB.Form frmUpdate 
   Caption         =   "Online Soft Web.Com Online Update"
   ClientHeight    =   3435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6570
   Icon            =   "frmUpdate.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3435
   ScaleWidth      =   6570
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   6120
      Top             =   3000
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit update"
      Height          =   465
      Left            =   4440
      TabIndex        =   5
      Top             =   2880
      Width           =   1455
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   6000
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Get New update"
      Enabled         =   0   'False
      Height          =   465
      Left            =   2640
      TabIndex        =   0
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   1800
      Left            =   2280
      Picture         =   "frmUpdate.frx":0442
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4335
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   3435
      Left            =   0
      Picture         =   "frmUpdate.frx":73B64
      Top             =   0
      Width           =   2310
   End
   Begin VB.Label lblNewVersion 
      AutoSize        =   -1  'True
      Caption         =   "lblNewVersion"
      Height          =   195
      Left            =   4440
      TabIndex        =   4
      Top             =   2400
      Width           =   1005
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Newest Version Available:"
      Height          =   195
      Left            =   2370
      TabIndex        =   3
      Top             =   2400
      Width           =   1845
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      Caption         =   "lblVersion"
      Height          =   195
      Left            =   4440
      TabIndex        =   2
      Top             =   1920
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "You Have Version:"
      Height          =   195
      Left            =   2880
      TabIndex        =   1
      Top             =   1920
      Width           =   1335
   End
End
Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Function HyperJump(ByVal Url As String) As Long
    HyperJump = ShellExecute(0&, vbNullString, Url, vbNullString, vbNullString, vbNormalFocus)
End Function

Private Sub cmdExit_Click()
    Inet1.Cancel
    Unload frmUpdate
    
End Sub

Private Sub cmdUpdate_Click()
    Dim strVersion As String, News As String
    On Error GoTo ErrorMessage
    frmUpdate.MousePointer = 11
    'The Version File Has Been Set To A Free Download From My Site
    strVersion = Inet1.OpenURL("http://www.onlinesoftweb.com/application.ver")
    'You can try this function on Your local disk, but You must change adresses:
    'for example: "file://c:\path\application.ver"
    If strVersion = "" Then GoTo Skip 'if file not found or file is empty then exit
    If strVersion <= App.Major & "." & App.Minor Then
        MsgBox "No newer version was released.", vbInformation
        GoTo Skip
    End If
    'now display MessageBox with news in newer strVersion(s) of application and two buttons Yes(update), No(end)
    News = Inet1.OpenURL("http://www.onlinesoftweb.com/news.txt")
    If MsgBox(Mid(News, 1, InStr(1, News, App.Major & "." & App.Minor) - 9), vbYesNo, "You can update from Version " & App.Major & "." & App.Minor & " to Version " & strVersion) = vbYes Then
        frmMain.Show
        Unload frmUpdate
    End If
Skip:
    frmUpdate.MousePointer = 0
    Inet1.Cancel
    Unload frmUpdate
    Exit Sub
ErrorMessage:
    frmUpdate.MousePointer = 0
    MsgBox "An Internet Lag Error has occured. Update failed." & Chr(10) & "Please Try Again Or You Can download new Version of this application manually at http://www.onlinesoftweb.com/cubnet.exe", vbCritical
End Sub

Private Sub Form_Load()
    MsgBox "The Version File Has Been Set To A Free Download From My Site." & vbNewLine & "If You Find This Program Usefull" & vbNewLine _
        & "Please Vote", vbCritical, "Online Soft Web.Com"
    frmUpdate.Caption = "Online Soft Web.Com Updater " & App.Major & "." & App.Minor
    lblVersion.Caption = App.Major & "." & App.Minor
    
    lblNewVersion.Caption = "Checking For New Version" & vbNewLine & "Please Wait!"
    
End Sub

Private Sub Timer1_Timer()
On Error GoTo MyError
    Dim strVersion As String
    
    strVersion = Inet1.OpenURL("http://www.onlinesoftweb.com/application.ver")
    lblNewVersion.Caption = strVersion
    cmdUpdate.Enabled = True
    
    Exit Sub
    
MyError:
    
End Sub
