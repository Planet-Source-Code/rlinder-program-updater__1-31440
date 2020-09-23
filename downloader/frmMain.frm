VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Caption         =   "Onlinesoftweb.com Update Downloader"
   ClientHeight    =   3720
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6855
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   6855
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Please Vote"
      Height          =   495
      Left            =   5880
      TabIndex        =   20
      Top             =   1320
      Width           =   735
   End
   Begin VB.TextBox FilePath 
      Height          =   285
      Left            =   2160
      ScrollBars      =   2  'Vertical
      TabIndex        =   18
      Text            =   "C:\Program Files\Cub Scout.Net Explorer"
      Top             =   3840
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "&Run"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1800
      TabIndex        =   16
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Timer tmrUpdateProgress 
      Interval        =   1
      Left            =   1200
      Top             =   3720
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   720
      TabIndex        =   4
      Text            =   "http://www.onlinesoftweb.com/cubnet.exe"
      Top             =   3360
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Timer tmrTimeLeft 
      Interval        =   1000
      Left            =   720
      Top             =   3720
   End
   Begin VB.CommandButton cmdDownload 
      Caption         =   "&Download"
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
      Left            =   600
      TabIndex        =   3
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "&File Download Progress"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   600
      TabIndex        =   1
      Top             =   1200
      Width           =   5055
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00C00000&
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         ScaleHeight     =   225
         ScaleWidth      =   4785
         TabIndex        =   2
         Top             =   240
         Width           =   4815
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Visible         =   0   'False
         Width           =   4815
      End
      Begin VB.Label lblSize 
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
         Left            =   960
         TabIndex        =   14
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblRecieve 
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
         TabIndex        =   13
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblSpeed 
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
         Left            =   4320
         TabIndex        =   12
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblElapsed 
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
         Left            =   3720
         TabIndex        =   11
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblRemaining 
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
         TabIndex        =   10
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Elapsed Time:"
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
         Left            =   2640
         TabIndex        =   9
         Top             =   960
         Width           =   990
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Time Remaining:"
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
         Top             =   960
         Width           =   1170
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Speed:"
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
         Left            =   3720
         TabIndex        =   7
         Top             =   600
         Width           =   510
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Bytes Recieved:"
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
         TabIndex        =   6
         Top             =   600
         Width           =   1290
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "File Size:"
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
         Top             =   600
         Width           =   630
      End
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "&Stop"
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
      Left            =   4560
      TabIndex        =   0
      Top             =   2760
      Width           =   1095
   End
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   1680
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   80
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "&Pause"
      Height          =   375
      Left            =   3480
      TabIndex        =   15
      Top             =   2760
      Width           =   1095
   End
   Begin MSComCtl2.Animation Animation1 
      Height          =   1095
      Left            =   1200
      TabIndex        =   19
      Top             =   0
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   1931
      _Version        =   393216
      FullWidth       =   265
      FullHeight      =   73
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
      Begin VB.Menu mnuOnline 
         Caption         =   "&Online Soft Web"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Data As String
Dim Percent%
Dim BeginTransfer As Single
Dim BytesAlreadySent As Single
Dim BytesRemaining As Single
Dim Header As Variant
Dim Status As String
Dim TransferRate As Single

Function ConvertTime(TheTime As Single)
    Dim NewTime As String
    Dim Sec As Single
    Dim Min As Single
    Dim H As Single

    If TheTime > 60 Then
        Sec = TheTime
        Min = Sec / 60
        Min = Int(Min)
        Sec = Sec - Min * 60
        H = Int(Min / 60)
        Min = Min - H * 60
        NewTime = H & ":" & Min & ":" & Sec
        If H < 0 Then H = 0
        If Min < 0 Then Min = 0
        If Sec < 0 Then Sec = 0
        NewTime = Format(NewTime, "HH:MM:SS")
        ConvertTime = NewTime
    End If


    If TheTime < 60 Then
        NewTime = "00:00:" & TheTime
        NewTime = Format(NewTime, "HH:MM:SS")
        ConvertTime = NewTime
    End If
End Function
Public Function StartUpdate(strURL As String)
    BytesAlreadySent = 1
    If strURL = "" Then Exit Function
    Url = strURL
    Dim Pos%, Length%, NextPos%, LENGTH2%, POS2%, POS3%
        Pos = InStr(strURL, "://") 'Record position of ://
        LENGTH2 = Len("://") 'Record the length of it
        Length = Len(strURL) 'Length of the entire url
            If InStr(strURL, "://") Then  ' check if they entered the http:// or ftp://
            strURL = Right(strURL, Length - LENGTH2 - Pos + 1) ' remove http:// or ftp://
            End If
                If InStr(strURL, "/") Then 'looks for the first / mark going from left to right
                POS2 = InStr(strURL, "/") 'gets the position of the / mark
    '-----------------GET THE FILENAME-------------
                Dim StrFile$: StrFile = strURL 'load the variables into each other
                Do Until InStr(StrFile, "/") = 0 'Do the loop until all is left is the filename
                LENGTH2 = Len(StrFile) 'get the length of the filename every time its passed over by the loop
                POS3 = InStr(StrFile, "/") 'find the / mark
                StrFile = Right(strURL, LENGTH2 - POS3) 'slash it down removing everything before the / mark including the / mark...
                Loop
                Filename = StrFile
    '----------------END GET FILE NAME--------------
                strSvrURL = Left(strURL, POS2 - 1) 'removes everything after the / mark leaving just the server name as the end result
    End If
    '-----------END TRIM THE URL FOR THE SERVER NAME-----------

End Function
Public Sub Reset()
    CloseSocket
    Data = ""
    Percent = 0
    BeginTransfer = 0
    BytesAlreadySent = 0
    BytesRemaining = 0
    Status = ""
    Header = ""
    RESUMEFILE = False
    UpdateProgress Picture1, 0
    cmdDownload.Enabled = True
    
End Sub
Public Sub CloseSocket()
    Do Until Winsock.State = 0
        Winsock.Close
        Winsock.LocalPort = 0
        Close #1
    Loop
    
End Sub

Private Sub cmdDownload_Click()
    StartUpdate Text1.Text
    frmSave.Show
    lblStatus.Visible = False
    Animation1.AutoPlay = True
    
End Sub

Private Sub cmdPause_Click()
    If BytesRemaining > BytesAlreadySent Then
        If Winsock.State > 0 Then
            Data = ""
            BeginTransfer = 0
            Status = ""
            Header = ""
            CloseSocket
            Picture1.Visible = False
            lblStatus.Visible = True
            lblStatus.Caption = "Download Paused"
            cmdPause.Caption = "Restart"
            cmdStop.Enabled = False
            Animation1.AutoPlay = False
            
        Else
            Picture1.Visible = True
            lblStatus.Visible = False
            FileLength = FileLen(FilePathName)
            RESUMEFILE = True
            frmMain.Winsock.Connect strSvrURL, 80
            cmdPause.Caption = "Pause"
            cmdStop.Enabled = True
            Animation1.AutoPlay = True
            
        End If
    
    End If
    
End Sub

Private Sub cmdRun_Click()
    Const conBtns As Integer = vbYesNoCancel + vbExclamation _
                            + vbDefaultButton3 + vbApplicationModal
    Const conMsg As String = "Do you want Install Cub Scout Explorer Updates"
    Dim intUserResponse As Integer
                   'document was changed since last save
        intUserResponse = MsgBox(conMsg, conBtns, "Cub Scout.Net")
        Select Case intUserResponse
            Case vbYes                      'user wants to Open Program Updates
                OpenIt frmMain, FilePathName
                End
            Case vbNo
                'Do nothing user does not want to Open Program Updates
            Case vbCancel
                'Do nothing return to Program-don't unload form
        End Select
        
End Sub

Private Sub cmdStop_Click()
    Const conBtns As Integer = vbYesNoCancel + vbExclamation _
                            + vbDefaultButton3 + vbApplicationModal
    Const conMsg As String = "Do you want Stop The Download"
    Dim intUserResponse As Integer
    
        intUserResponse = MsgBox(conMsg, conBtns, "Online Soft Web.Com Updater")
        Select Case intUserResponse
            Case vbYes        'user wants to Stop Download
                If Winsock.State > 0 Then
                    CloseSocket
                    MsgBox "Download Aborted!", vbExclamation, "Download Aborted"
                    Animation1.AutoPlay = False
                    Reset
        
                End If
                
            Case vbNo                       'user does not want Stop The Download
                Exit Sub
            Case vbCancel
                Exit Sub                 'user does not want Stop The Download
        End Select

    
End Sub

Private Sub Command1_Click()
    OpenIt Me, "http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=31440&lngWId=1"
End Sub

Private Sub exit_Click()
    Unload frmMain
    
End Sub

Private Sub Form_Load()
    RESUMEFILE = False
    strFormLoaded = "Main"
    Animation1.Open (App.Path & "\" & "Filemove.avi")
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    CloseSocket
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CloseSocket
    
End Sub

Private Sub mnuOnline_Click()
    OpenIt Me, "http://www.onlinesoftweb.com"
    
End Sub

Private Sub tmrTimeLeft_Timer()
'On Error Resume Next
    If BytesRemaining > 0 And BytesAlreadySent > 0 Then
        If BytesRemaining <= BytesAlreadySent Then
            lblSpeed = 0
            CloseSocket
            lblElapsed = Format(Hr & ":" & Min & ":" & Sec, "HH:MM:SS")
            cmdDownload.Enabled = False
            cmdRun.Enabled = True
            Picture1.Visible = False
            lblStatus.Visible = True
            lblStatus.Caption = "Download Completed"
            Reset
        Else
            Sec = Sec + 1
            If Sec >= 60 Then
            Sec = 0
            Min = Min + 1
            ElseIf Min >= 60 Then
            Min = 0
            Hr = Hr + 1
            End If
            cmdDownload.Enabled = False
            cmdRun.Enabled = False
            lblElapsed = Format(Hr & ":" & Min & ":" & Sec, "HH:MM:SS")
            'The reason I divide the difference of bytesalreadysent and bytesremaining is becuase they are in bytes right now.. I want it to be in KB so it can be Kbps and not bps
            lblRemaining = ConvertTime(Int(((BytesRemaining - BytesAlreadySent) / 1024) / TransferRate))
            lblSpeed = TransferRate
        End If
    
    End If
End Sub

Private Sub tmrUpdateProgress_Timer()
On Error Resume Next
    If BytesAlreadySent > 0 And BytesRemaining > 0 Then
        lblRecieve = File_ByteConversion(BytesAlreadySent)
        lblSize = File_ByteConversion(BytesRemaining)
        Percent = Format((BytesAlreadySent / BytesRemaining) * 100, "00") 'calculates the percentage completed
        UpdateProgress Picture1, Percent 'updates progress bar with new percentage rate
    End If
    
End Sub

Private Sub Winsock_Connect()
On Error Resume Next
    Dim strCommand As String
    strCommand = "GET " + Url + " HTTP/1.0" + vbCrLf 'tells server to GET the file if you just want the header info and not the data change "GET " to "HEAD "
    strCommand = strCommand + "Accept: *.*, */*" + vbCrLf
    If RESUMEFILE = True Then
        strCommand = strCommand + "Range: bytes=" & FileLength & "-" & vbCrLf
    End If
    
    strCommand = strCommand + "User-Agent: Online Soft Web.Com" & vbCrLf
    strCommand = strCommand + "Referer: " & strSvrURL & vbCrLf
    strCommand = strCommand + vbCrLf
    Winsock.SendData strCommand 'sends a header to the server instructing it what to do!
    BeginTransfer = Timer 'start timer for transfer rate
    
End Sub

Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)
    Winsock.GetData Data, vbString
    If InStr(Data, "Content-Type:") Then 'find out if this chunk has the header..you can change that to anything that the header contains
            
            If RESUMEFILE = True Then 'check to see if its gonna resume ok or not..This is actually the worst way to check this.
                If InStr(Data, "HTTP/1.1 206 Partial Content") = 0 Then
                MsgBox "Server did not accept resuming.", vbCritical, "No Resuming Support"
                Exit Sub
                Reset
                CloseSocket
                End If
            End If
            
        Dim Pos%, Length%, HEAD$
        Pos = InStr(Data, vbCrLf & vbCrLf) ' find out where the header and the data is split apart
        Length = Len(Data) 'get the length of the data chunk
        HEAD = Left(Data, Pos - 1) 'Get the header from the chunk of data and ignore the data content
        Data = Right(Data, Length - Pos - 3) 'Get the data from the first chunk that contains the header also
        Header = Header & HEAD 'Append the header to header text box
    
        If RESUMEFILE = True Then
            BytesAlreadySent = FileLength + 1
            BytesRemaining = GETDATAHEAD(Header, "Content-Length:")
            BytesRemaining = BytesRemaining + FileLength
        Else
            BytesRemaining = GETDATAHEAD(Header, "Content-Length:")
        End If
        
    End If

'-----------BEGIN WRITE CHUNK TO FILE CODE--------
        Open FilePathName For Binary Access Write As #1 'opens file for output
        Put #1, BytesAlreadySent, Data 'writes data to the end of file
        BytesAlreadySent = Seek(1)
        Close #1 'close file for now until next data chunk is available
'--------------------------------------------------

'Lets explain this a bit..The variable BeginTransfer is given the starting value of the
'timer which in case you dont know is the amount of seconds til midnight but that has
'nothing to do with this. Anyways so its given the amount for the start time and then
'when this event below is fired for the first time the timer will be given the value again
'since your system clock was ticking along while the operation between the two of these
'events happened the number will be different.  The two values are subtracted and divided
'by the amount recieved and then by 1000 and put into a readable format
If RESUMEFILE = False Then
'This is pretty straightforward if you ever taken math before you can tell what im doing!
TransferRate = Format(Int(BytesAlreadySent / (Timer - BeginTransfer)) / 1000, "####.00")
Else
'If you dont subtract the difference you will get a really large and odd download speed hehe.
TransferRate = Format(Int((BytesAlreadySent - FileLength) / (Timer - BeginTransfer)) / 1000, "####.00")
End If
End Sub

