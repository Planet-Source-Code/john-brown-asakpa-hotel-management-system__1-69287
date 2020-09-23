VERSION 5.00
Begin VB.Form frmBackup 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Database Backup"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frameCurrBackUp 
      Caption         =   "Choose Path for BackUp"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2895
      Left            =   240
      TabIndex        =   7
      Top             =   2760
      Width           =   6855
      Begin VB.DriveListBox Drive1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   480
         TabIndex        =   12
         Top             =   480
         Width           =   3015
      End
      Begin VB.DirListBox Dir1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1665
         Left            =   480
         TabIndex        =   11
         Top             =   960
         Width           =   3015
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&BackUp"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5160
         TabIndex        =   10
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5160
         Picture         =   "frmBackup.frx":0000
         TabIndex        =   9
         Top             =   1440
         Width           =   1095
      End
      Begin VB.FileListBox File1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   3600
         TabIndex        =   8
         Top             =   2040
         Visible         =   0   'False
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Last Backup Detail"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   6855
      Begin VB.Label lblLastPath 
         AutoSize        =   -1  'True
         Caption         =   "Last BackUp Path"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   1440
         TabIndex        =   6
         Top             =   360
         Width           =   1650
      End
      Begin VB.Label lblPath 
         AutoSize        =   -1  'True
         Caption         =   "Path"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Width           =   435
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   360
         TabIndex        =   4
         Top             =   720
         Width           =   450
      End
      Begin VB.Label Time 
         AutoSize        =   -1  'True
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   360
         TabIndex        =   3
         Top             =   1080
         Width           =   435
      End
      Begin VB.Label lblLastDate 
         AutoSize        =   -1  'True
         Caption         =   "Last BackUp Path"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   1440
         TabIndex        =   2
         Top             =   720
         Width           =   1650
      End
      Begin VB.Label lblLastTime 
         AutoSize        =   -1  'True
         Caption         =   "Last BackUp Path"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   1440
         TabIndex        =   1
         Top             =   1080
         Width           =   1650
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "BACK UP DATABSE"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   1800
      TabIndex        =   13
      Top             =   240
      Width           =   2490
   End
End
Attribute VB_Name = "frmBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Note : Microsoft Scripting Runtime Library is Referenced
'       For Making Object of File System Object


Dim Fsys As New FileSystemObject
Dim bckupFile As file

'Reading Previously Backup Details
Private Sub Form_Load()
    Connect
    
    Dim lastPath As String
    Dim lastDate As String
    Dim lastTime As String
    
    'Read Registry for previous settings stored
    lastPath = GetSetting(App.Title, "Settings", "BackupPath")
    lastDate = GetSetting(App.Title, "Settings", "BackupDate")
    lastTime = GetSetting(App.Title, "Settings", "BackupTime")
    
    If lastPath = "" Then
        lblLastPath.Caption = "No Backup made previously"
        lblLastDate.Caption = " "
        lblLastTime.Caption = " "
    Else
        lblLastPath.Caption = lastPath
        lblLastDate.Caption = lastDate & "  (mm-dd-yy)"
        lblLastTime.Caption = lastTime
    End If
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

'Backup Cmd Btn
Private Sub cmdsave_click()


On Error Resume Next
    Dim dbname As String
    
    dbname = "Backup_HMS-" & Format$(Date, "mm-dd-yyyy") & ".mdb"
    cmdSave.Enabled = False
   

    Label1.BackColor = vbGreen
    Label1.ForeColor = vbYellow
    Dim destination As String
    Dim Source As String
    Dim currDate, currTime As String
    currDate = Format$(Now, "mm - dd - yy")
    currTime = Format$(Now, "hh:mm:ss AM/PM")
    
    destination = File1.Path & "\" & dbname
    'destination = File1.Path & "\" & "HMS2.mdb"
    
    Source = App.Path & "\database\HMS.mdb"
    
    'MsgBox "Source : " & source
    'MsgBox "Destination : " & destination
    Set bckupFile = Fsys.GetFile(finalpath)
    bckupFile.Attributes = Compressed
    Fsys.CopyFile Source, destination, True
    'Saving Current Backup Details
    SaveSetting App.Title, "Settings", "BackupPath", destination
    SaveSetting App.Title, "Settings", "BackupDate", currDate
    SaveSetting App.Title, "Settings", "BackupTime", currTime
    
    
    cmdSave.Enabled = True
    
   
        MsgBox "BackUp Process Over", vbInformation, "Backup"
        
        With RS_Userlog
       .AddNew
       .Fields(0) = UserName
       .Fields(1) = "Database Backup"
       .Fields(2) = Date
       .Fields(3) = Time
       .Fields(4) = "Successful"
       .Update
    End With

   
End Sub

Private Sub Drive1_Change()
    On Error Resume Next
    Dir1.Path = Drive1.Drive
End Sub

Private Sub Dir1_Change()
    On Error Resume Next
    File1.Path = Dir1.Path
    
    
End Sub

