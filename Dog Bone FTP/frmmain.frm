VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmmain 
   Caption         =   "Dog Bone FTP"
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   8160
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraStatus 
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   120
      TabIndex        =   13
      Top             =   5400
      Width           =   7965
      Begin VB.Label lblStatus 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   7695
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   9128
      _Version        =   393216
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Connect"
      TabPicture(0)   =   "frmmain.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblPassword"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblUserName"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txbURL"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txbPassword"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txbUserName"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdConnect"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdDisconnect"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Files"
      TabPicture(1)   =   "frmmain.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdDownLoad"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdUpload"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "fraLocalFiles"
      Tab(1).Control(3)=   "fraServerFiles"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "About"
      TabPicture(2)   =   "frmmain.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label2"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label3"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label4"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      Begin VB.CommandButton cmdDownLoad 
         Caption         =   "Download "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72120
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "Download file from server to local"
         Top             =   4680
         Width           =   1125
      End
      Begin VB.CommandButton cmdUpload 
         Caption         =   "Upload"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -70860
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "Upload selected local file to selected server dir"
         Top             =   4680
         Width           =   1125
      End
      Begin VB.Frame fraLocalFiles 
         Caption         =   "Local files"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4245
         Left            =   -70920
         TabIndex        =   18
         Top             =   360
         Width           =   3735
         Begin VB.DirListBox Dir1 
            BackColor       =   &H00E0E0E0&
            Height          =   1665
            Left            =   120
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   390
            Width           =   3495
         End
         Begin VB.FileListBox File1 
            BackColor       =   &H00E0E0E0&
            Height          =   1845
            Left            =   120
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   2040
            Width           =   3495
         End
         Begin VB.Label lblLocalFilesHelp 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            Caption         =   " ?"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   225
            Left            =   3480
            TabIndex        =   21
            ToolTipText     =   "Help"
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Frame fraServerFiles 
         Caption         =   "Server files"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4245
         Left            =   -74880
         TabIndex        =   15
         Top             =   360
         Width           =   3915
         Begin VB.ListBox lisServerFiles 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   3765
            Left            =   180
            Sorted          =   -1  'True
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   300
            Width           =   3585
         End
         Begin VB.Label lbServerFilesHelp 
            Appearance      =   0  'Flat
            BackColor       =   &H00808080&
            Caption         =   " ?"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   225
            Left            =   3600
            TabIndex        =   17
            ToolTipText     =   "Help"
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.CommandButton cmdDisconnect 
         Caption         =   "Disconnect"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6480
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Disconnect logon "
         Top             =   2760
         Width           =   1125
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "Connect"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6480
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Connect to server"
         Top             =   2280
         Width           =   1125
      End
      Begin VB.Frame Frame1 
         Caption         =   "Options"
         Height          =   1335
         Left            =   4320
         TabIndex        =   7
         Top             =   840
         Width           =   3375
         Begin VB.CommandButton cmdNil 
            Caption         =   "Connect With UserName"
            Height          =   315
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   10
            TabStop         =   0   'False
            ToolTipText     =   "No Username and Password"
            Top             =   960
            Width           =   3135
         End
         Begin VB.CommandButton cmdPrivate 
            Caption         =   "Use a Registered UserName"
            Height          =   315
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   9
            TabStop         =   0   'False
            ToolTipText     =   "Use registered UserName"
            Top             =   600
            Width           =   3135
         End
         Begin VB.CommandButton cmdPublic 
            Appearance      =   0  'Flat
            Caption         =   "Connect Anomymously"
            Height          =   315
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   8
            TabStop         =   0   'False
            ToolTipText     =   "Use Anonymous as UserName"
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.TextBox txbUserName 
         Height          =   315
         Left            =   1440
         TabIndex        =   4
         Top             =   840
         Width           =   2715
      End
      Begin VB.TextBox txbPassword 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1440
         PasswordChar    =   "|"
         TabIndex        =   3
         Top             =   1200
         Width           =   2715
      End
      Begin VB.TextBox txbURL 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   480
         Width           =   6345
      End
      Begin VB.Label Label4 
         Caption         =   $"frmmain.frx":0054
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00D26D2B&
         Height          =   1455
         Left            =   -74760
         TabIndex        =   26
         Top             =   1560
         Width           =   4695
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "dogbonevb@hotmail.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -69960
         TabIndex        =   25
         Top             =   1560
         Width           =   2670
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Dog Bone FTP"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   -74760
         TabIndex        =   24
         Top             =   480
         Width           =   6540
      End
      Begin VB.Label lblUserName 
         AutoSize        =   -1  'True
         Caption         =   "UserName:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   1185
      End
      Begin VB.Label lblPassword 
         AutoSize        =   -1  'True
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Domain:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   870
      End
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef lpdwFlags _
      As Long, ByVal dwReserved As Long) As Long


Const defaultURL = "ftp://www."
Const defaultUserName = ""
Const defaultPassword = ""
Const defaultEMailAddress = ""

Dim ConnectedFlag As Boolean
Dim ServerDirFlag As Boolean
Dim DownloadFlag As Boolean
Dim UploadFlag As Boolean
Dim FileSizeFlag As Boolean

Dim homeLen As Integer
Dim LocFilespec As String
Dim SerFilespec As String
Dim gFileSize As String

Private Sub Form_Load()
    GetStartingDefaults
    ConnectedFlag = False
    ClearFlags
    UpdButtons
End Sub

Private Sub GetStartingDefaults()
    txbURL.Text = defaultURL
    txbUserName.Text = ""
    txbPassword.Text = ""
End Sub



Private Sub cmdConnect_click()
     On Error Resume Next
     Dim tmp As String
     Dim i As Integer
     
     Inet1.Cancel
     Inet1.Execute , "CLOSE"
    
     Err.Clear
     On Error GoTo errHandler
    
     ClearFlags
    
     If Len(txbURL) < 6 Then
          MsgBox "No URL yet", , "Dog Bone FTP"
          Exit Sub
     End If
    
     If UCase(Left(txbURL, 6)) <> "FTP://" Then
          MsgBox "No FTP protocol entered in URL", , "Dog Bone FTP"
          Exit Sub
     End If
    
     lblStatus.Caption = "To connect ...."
    
       ' (Note we use txtURL.Text here; you can just use txtURL if you wish)
     Inet1.AccessType = icUseDefault
     Inet1.URL = LTrim(Trim(txbURL.Text))
     Inet1.UserName = LTrim(Trim(txbUserName.Text))
     Inet1.Password = LTrim(Trim(txbPassword.Text))
     Inet1.RequestTimeout = 40
            
       ' Will force to bring up Dialup Dialog if not already having a line
     ServerDirFlag = True
     Inet1.Execute , "DIR"
     Do While Inet1.StillExecuting
          DoEvents
          ' Connection not established yet, hence cannot
          ' try to fall back on ConnectedFlag to exit
     Loop
     txbURL.Text = Inet1.URL
     
          ' Home portion
     For i = 7 To Len(txbURL.Text)
          tmp = Mid(txbURL.Text, i, 1)
          If tmp = "/" Then
               Exit For
          End If
     Next i
     homeLen = i - 1
     
     If IsNetConnected() Then
          ConnectedFlag = True
          UpdButtons
     Else
          GoTo errHandler
     End If
     Exit Sub
    
errHandler:
    If icExecuting Then
           ' We place this here in case command for "CLOSE" failed.
           ' With Inet, one can never tell.
         If ConnectedFlag = False Then
              Exit Sub
         End If
        
         If MsgBox("Executing job. Cancel it?", vbYesNo + vbQuestion, "Dog Bone FTP") = vbYes Then
              Inet1.Cancel
              If Inet1.StillExecuting Then
                  lblStatus.Caption = "System failed to cancel job"
              End If
         Else
              Resume
         End If
     End If
     ErrMsgProc "cmdConnect_Click"
End Sub






Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Inet1.Execute , "CLOSE"
    Unload Me
End Sub



Private Sub cmdPublic_Click()
    txbUserName.Text = "anonymous"
    txbPassword.PasswordChar = ""
    txbPassword.Text = defaultEMailAddress
    txbURL.SetFocus
End Sub



Private Sub cmdPrivate_Click()
    txbUserName.Text = defaultUserName
    txbPassword.PasswordChar = "*"
    txbPassword.Text = defaultPassword
    txbURL.SetFocus
End Sub



Private Sub cmdNil_Click()
    txbUserName.Text = ""
    txbPassword.Text = ""
    txbURL.SetFocus
End Sub




Private Sub cmdDisconnect_Click()
    On Error Resume Next
    Inet1.Cancel
    Inet1.Execute , "CLOSE"
    lblStatus.Caption = "Unconnected"
       ' Put back starting default
    GetStartingDefaults
    ConnectedFlag = False
    lisServerFiles.Clear
    ClearFlags
    UpdButtons
End Sub




Private Sub ClearFlags()
    ServerDirFlag = False
    DownloadFlag = False
    UploadFlag = False
    FileSizeFlag = False
End Sub




Private Sub UpdButtons()
    cmdConnect.Enabled = False
    cmdDownLoad.Enabled = False
    cmdUpload.Enabled = False
    cmdDisconnect.Enabled = False
    If ConnectedFlag Then
           ' Once connected, no interference to txbURL.text
         txbURL.Locked = True
         cmdDownLoad.Enabled = True
         cmdUpload.Enabled = True
         cmdDisconnect.Enabled = True
    Else
         txbURL.Locked = False
         cmdConnect.Enabled = True
    End If
End Sub




Private Sub cmdDownLoad_Click()
     On Error GoTo errHandler
     
     If ConnectedFlag = False Then
          MsgBox "No connection yet", , "Dog Bone FTP"
          Exit Sub
     ElseIf lisServerFiles.ListCount = 0 Then
          MsgBox "No server file listed yet", , "Dog Bone FTP"
          Exit Sub
     ElseIf Right(lisServerFiles.Text, 1) = "/" Then
          MsgBox "Selected item is a directory only." & vbCrLf & vbCrLf & _
             "To list files under that dir, double click on it.", , "Dog Bone FTP"
          Exit Sub
     End If
    
     lblStatus.Caption = "Retreiving file..."
     SerFilespec = Right(txbURL.Text, Len(txbURL.Text) - homeLen) & _
               "/" & lisServerFiles.Text
     SerFilespec = Right(SerFilespec, Len(SerFilespec) - 1)
     
        ' Use same file name and store it in current dir of local. Parse
        ' above SerFilespec and take only the file name as LocFileSpec.
     LocFilespec = SerFilespec
     Do While InStr(LocFilespec, "/") <> 0
         LocFilespec = Right(LocFilespec, Len(LocFilespec) - _
              InStr(LocFilespec, "/"))
     Loop
     
     If IsFileThere(LocFilespec) Then
          If MsgBox(LocFilespec & " already exist. Overwrite?", _
               vbYesNo + vbQuestion, "Dog Bone FTP") = vbNo Then
               Exit Sub
          End If
     End If
     
     lblStatus.Caption = "Requesting for file size..."
     
     gFileSize = ""
     FileSizeFlag = True
     Inet1.Execute , "SIZE " & SerFilespec
     Do While Inet1.StillExecuting
          DoEvents
          If ConnectedFlag = False Then
               Exit Sub
          End If
     Loop
         
     If gFileSize = "" Then
          MsgBox "Selected file has 0 byte content.", , "Dog Bone FTP"
          Exit Sub
     Else
          If MsgBox("File size is " & gFileSize & " bytes." & vbCrLf & vbCrLf & _
                  "Proceed to download?", vbYesNo + vbQuestion, "Dog Bone FTP") = vbNo Then
              Exit Sub
          End If
     End If
     
     DownloadFlag = True
     Inet1.Execute , "Get " & SerFilespec & " " & LocFilespec
     Do While Inet1.StillExecuting
          DoEvents
          If ConnectedFlag = False Then
               Exit Sub
          End If
     Loop

     lblStatus.Caption = "Connected"
     File1.Refresh
     Exit Sub
     
errHandler:
    If icExecuting Then
        If ConnectedFlag = False Then
            Exit Sub
        End If
        
        If MsgBox("Executing job. Cancel it?", vbYesNo + vbQuestion, "Dog Bone FTP") = vbYes Then
            Inet1.Cancel
            If Inet1.StillExecuting Then
                lblStatus.Caption = "System failed to cancel job"
            End If
        Else
            Resume
        End If
    End If
    ErrMsgProc "cmdDownLoad_Click"
End Sub




' Assuming you have the appropriate privileges on the server
Private Sub cmdUpLoad_Click()
     On Error GoTo errHandler
     Dim tmpPath As String
     Dim tmpFile As String
     Dim bExist As Boolean
     Dim lFileSize As Long
     Dim i
     
     If ConnectedFlag = False Then
          MsgBox "No connection yet", , "Dog Bone FTP"
          Exit Sub
     ElseIf File1.ListCount = 0 Then
          MsgBox "No local file in current dir yet", , "Dog Bone FTP"
          Exit Sub
     ElseIf Not (Right(lisServerFiles.Text, 1) = "/") Then
          MsgBox "Selected server file item is not a directory", , "Dog Bone FTP"
          Exit Sub
     ElseIf lisServerFiles.Text = "../" Then
          MsgBox "No directory name selected yet", , "Dog Bone FTP"
          Exit Sub
     End If
    
     LocFilespec = tmpPath & File1.List(File1.ListIndex)
     If LocFilespec = "" Then
          MsgBox "No local file selected yet", , "Dog Bone FTP"
          Exit Sub
     End If
     
     lFileSize = FileLen(LocFilespec)
     If MsgBox("File size is " & CStr(lFileSize) & " bytes." & vbCrLf & vbCrLf & _
                  "Proceed to upload?", vbYesNo + vbQuestion, , "Dog Bone FTP") = vbNo Then
         Exit Sub
     End If
    
     lblStatus.Caption = "Uploading file..."
     
     If Right(Dir1.Path, 1) <> "\" Then
          tmpPath = Dir1.Path & "\"
     Else
          tmpPath = Dir1.Path                   ' e.g. root "C:\"
     End If
     
     SerFilespec = Right(txbURL.Text, Len(txbURL.Text) - homeLen) & _
               "/" & lisServerFiles.Text
          ' Remove the front "/" from above
     SerFilespec = Right(SerFilespec, Len(SerFilespec) - 1)
     
     SerFilespec = SerFilespec & File1.List(File1.ListIndex)
     
          ' In order to test whether same file on server already exists
     lblStatus.Caption = "Verifying existence of file of same name..."
     tmpPath = SerFilespec
     ServerDirFlag = True
     Inet1.Execute , "DIR " & tmpPath & "/*.*"
     Do While Inet1.StillExecuting
          DoEvents
          If ConnectedFlag = False Then
               Exit Sub
          End If
     Loop
         
     bExist = False
     If lisServerFiles.ListCount > 0 Then
          For i = 0 To lisServerFiles.ListCount - 1
               tmpFile = lisServerFiles.List(i)
               If tmpFile = File1.List(File1.ListIndex) Then
                    bExist = True
                    Exit For
               End If
          Next i
     End If
         
          ' Go back
     ServerDirFlag = True
     Inet1.Execute , "DIR ../*"
     Do While Inet1.StillExecuting
          DoEvents
          If ConnectedFlag = False Then
               Exit Sub
          End If
     Loop
          
          
     If bExist Then
          If MsgBox("File already exist in selected server dir.  Supersede?", _
                  vbYesNo + vbQuestion, "Dog Bone FTP") = vbNo Then
               Exit Sub
          End If
     End If
     
     Exit Sub
         
     UploadFlag = True
     Inet1.Execute , "PUT " & LocFilespec & " " & SerFilespec
     
     Do While Inet1.StillExecuting
          DoEvents
          If ConnectedFlag = False Then
               Exit Sub
          End If
     Loop

     lblStatus.Caption = "Connected"
     Exit Sub
    
errHandler:
     If icExecuting Then
         If ConnectedFlag = False Then
              Exit Sub
         End If
        
         If MsgBox("Executing job. Cancel it?", vbYesNo + vbQuestion, "Dog Bone FTP") = vbYes Then
              Inet1.Cancel
              If Inet1.StillExecuting Then
                   lblStatus.Caption = "System failed to cancel job"
              End If
         Else
              Resume
         End If
     End If
     MsgBox "cmdUpload_Click", , "Dog Bone FTP"
End Sub



Private Sub lbServerFilesHelp_Click()
     MsgBox "Help:" & vbCrLf & vbCrLf & _
          "To change dir, double click a directory item on list." & vbCrLf & _
          "   (To go up one level, click the '../' item)" & vbCrLf & vbCrLf & _
          "To select a file for download, highlight it then" & vbCrLf & _
          "   click Download button (will report file size)." & vbCrLf & vbCrLf & _
          "To upload a local file, highlight a server dir first," & vbCrLf & _
          "   highlight a local file, then click Upload button." & vbCrLf & vbCrLf, , "Dog Bone FTP"
End Sub



Private Sub lblLocalFilesHelp_Click()
     MsgBox "Help:" & vbCrLf & vbCrLf & _
          "To see file size of a local file, double click the" & vbCrLf & _
            "   local file item." & vbCrLf & vbCrLf & _
          "For other Help, refer Server Files." & vbCrLf & vbCrLf, , "Dog Bone FTP"
End Sub



' For local files, we have FileSystem control to go up and down of dir hierachy
' and list individual files under a dir, but for server files listing, we have to
' provide a similar facility.
Private Sub lisServerFiles_dblClick()
     On Error GoTo errHandler
     
     If Not (Right(lisServerFiles.Text, 1) = "/") Then
          Exit Sub
     End If
     
     Dim tmpDir As String, tmp As String
     Dim i
     If Trim(lisServerFiles.Text) = "../" Then
          For i = Len(txbURL.Text) To 7 Step -1
               tmp = Mid(txbURL.Text, i, 1)
               If tmp = "/" Then
                    Exit For
               End If
          Next i
          If i = 7 Then
               MsgBox "No upper level of dir", , "Dog Bone FTP"
               Exit Sub
          End If
          txbURL.Text = Left(txbURL.Text, i - 1)
             ' Relative dir
          tmpDir = "../*"
     Else
          txbURL.Text = txbURL.Text & "/" & _
                   Left(lisServerFiles.Text, Len(lisServerFiles.Text) - 1)
          tmpDir = Right(txbURL.Text, Len(txbURL.Text) - homeLen) & "/*"
     End If
     ServerDirFlag = True
     Inet1.Execute , "DIR " & tmpDir
     Do While Inet1.StillExecuting
          DoEvents
          If ConnectedFlag = False Then
               Exit Sub
          End If
     Loop
     Exit Sub
    
errHandler:
    Select Case Err.Number
        Case icExecuting
             Resume
        Case Else
             ErrMsgProc "lisServerFiles_dblClick"
     End Select
End Sub



Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub



Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub



Private Sub File1_dblClick()
    If File1.ListCount = 0 Then
         Exit Sub
    End If
    Dim lFileSize As Long
    lFileSize = FileLen(File1.List(File1.ListIndex))
    MsgBox CStr(lFileSize) & " bytes", , "Dog Bone FTP"
End Sub




Private Sub Inet1_StateChanged(ByVal State As Integer)
    On Error Resume Next
    Select Case State
        Case icError                                      ' 11
            lblStatus = Inet1.ResponseCode & ": " & Inet1.ResponseInfo
            Inet1.Execute , "CLOSE"
            lblStatus.Caption = "Unconnected"
            lisServerFiles.Clear
            ConnectedFlag = False
            ServerDirFlag = False
            DownloadFlag = False
            UpdButtons
            
        Case icResponseCompleted                          ' 12
            Dim bDone As Boolean
            Dim tmpData As Variant       ' GetChunk returns Variant type
            
            If ServerDirFlag = True Then
                 Dim dirData As String
                 Dim strEntry As String
                 Dim i As Integer, k As Integer
            
                 tmpData = Inet1.GetChunk(4096, icString)
                 dirData = dirData & tmpData
            
                 If dirData <> "" Then
                     lisServerFiles.Clear
                       ' Use relative address to allow one dir level up
                     lisServerFiles.AddItem ("../")
                     For i = 1 To Len(dirData) - 1
                          k = InStr(i, dirData, vbCrLf)        ' We don't want CRLF
                          strEntry = Mid(dirData, i, k - i)
                          If Right(strEntry, 1) = "/" Then
                               strEntry = Left(strEntry, Len(strEntry) - 1) & "/"
                          End If
                          If Trim(strEntry) <> "" Then
                               lisServerFiles.AddItem strEntry
                          End If
                          i = k + 1
                          DoEvents
                     Next i
                     lisServerFiles.ListIndex = 0
                 End If
                 
                 ServerDirFlag = False
                 lblStatus.Caption = "Dir completed"
                 
            ElseIf DownloadFlag Then
                 Dim varData As Variant
                 
                 bDone = False

                 Open LocFilespec For Binary Access Write As #1
    
                   ' Get first chunk
                 tmpData = Inet1.GetChunk(10240, icByteArray)
                 DoEvents
                 If Len(tmpData) = 0 Then
                      bDone = True
                 End If
                 Do While Not bDone
                      varData = tmpData
                      Put #1, , varData
                      tmpData = Inet1.GetChunk(10240, icByteArray)
                      DoEvents
                      If ConnectedFlag = False Then
                           Exit Sub
                      End If
                      If Len(tmpData) = 0 Then
                            bDone = True
                      End If
                 Loop
                 Close #1
                 DownloadFlag = False
                 DoEvents
                 lblStatus.Caption = "Download completed"
                 DownloadFlag = False
                 MsgBox "Download completed:" & vbCrLf & vbCrLf & _
                     "File in current dir, named  " & LocFilespec, , "Dog Bone FTP"
                 
            ElseIf UploadFlag Then
                 lblStatus.Caption = "Connected"
                 UploadFlag = False
                 MsgBox "Download completed: File in " & LocFilespec, , "Dog Bone FTP"
                 
            ElseIf FileSizeFlag Then
                 Dim sizeData As String
            
                 tmpData = Inet1.GetChunk(1024, icString)
                 DoEvents
                 If Len(tmpData) > 0 Then
                      sizeData = sizeData & tmpData
                 End If
                 
                 gFileSize = sizeData
                 FileSizeFlag = False
                 
            Else
                 lblStatus.Caption = "Connected"
            End If
            
            
        Case icNone                                       ' 0
            lblStatus.Caption = "No state to report"
        Case icResolvingHost                              ' 1
            lblStatus.Caption = "Resolving host..."
        Case icHostResolved                               ' 2
            lblStatus.Caption = "Host resolved - found its IP address"
        Case icConnecting                                 ' 3
            lblStatus.Caption = "Connecting..."
        Case icConnected                                  ' 4
            lblStatus.Caption = "Connected"
        Case icRequesting                                 ' 5
            lblStatus.Caption = "Sending requesst..."
        Case icRequestSent                                ' 6
            lblStatus.Caption = "Request sent"
        Case icReceivingResponse                          ' 7
            lblStatus = "Receiving data..."
        Case icResponseReceived                           ' 8
            lblStatus = "Response received"
        Case icDisconnecting                              ' 9
            lblStatus.Caption = "Disconnecting..."
        Case icDisconnected                               '10
            lblStatus = "Disconnected"
    End Select
End Sub



Function IsNetConnected() As Boolean
    IsNetConnected = InternetGetConnectedState(0, 0)
End Function
                  


Sub ErrMsgProc(mMsg As String)
    MsgBox mMsg & vbCrLf & Err.Number & Space(5) & Err.Description, , "Dog Bone FTP"
End Sub



Function IsFileThere(inFileSpec As String) As Boolean
    On Error Resume Next
    Dim i
    i = FreeFile
    Open inFileSpec For Input As i
    If Err Then
        IsFileThere = False
    Else
        Close i
        IsFileThere = True
    End If
End Function



