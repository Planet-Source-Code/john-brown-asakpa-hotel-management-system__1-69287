VERSION 5.00
Begin VB.Form frmWelcome 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3075
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   3075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2310
      Left            =   0
      Picture         =   "frmWelcome.frx":0000
      ScaleHeight     =   2250
      ScaleWidth      =   3000
      TabIndex        =   0
      ToolTipText     =   "Press any key to continue"
      Top             =   0
      Width           =   3060
      Begin VB.Label lbl_user 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "User name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   0
         TabIndex        =   3
         ToolTipText     =   "Press any key to continue"
         Top             =   720
         Width           =   3015
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   3000
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label lbl_time 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Login at :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   0
         TabIndex        =   2
         ToolTipText     =   "Press any key to continue"
         Top             =   1200
         Width           =   3015
      End
      Begin VB.Label lbl_day 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Today :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   0
         TabIndex        =   1
         ToolTipText     =   "Press any key to continue"
         Top             =   1680
         Width           =   3015
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Press any key to continue"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   2520
      Width           =   2535
   End
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Const HWND_TOPMOST = -1
'Const HWND_NOTOPMOST = -2
'Const SWP_NOSIZE = &H1
'Const SWP_NOMOVE = &H2
'Const SWP_NOACTIVATE = &H10
'Const SWP_SHOWWINDOW = &H40

Private Declare Sub SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim i As Integer
Private Sub popup()
On Error Resume Next
    Picture1.Visible = True
    i = Me.Height
    Me.Height = 0
    While Me.Height < i
        Me.Height = Me.Height + 2
        Me.Top = Me.Top - 2
        DoEvents
    Wend
End Sub
Private Sub popdown()
On Error Resume Next
    i = Me.Height
    While Me.Height > 500
        Me.Height = Me.Height - 2
        Me.Top = Me.Top + 2
        DoEvents
    Wend
End Sub
Private Sub Form_Activate()
On Error Resume Next
    mdi_start.Enabled = False
    lbl_user.Caption = UserName
    lbl_time.Caption = "Login at:" & Format$(Now, "hh:mm:ss AM/PM")
    lbl_day.Caption = "Today:" & Format$(Date, "dd-MMM-yy")
    Call popup
    'SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    'Sleep welcometime 'Wait for 1 Seconds
    'Call popdown
main_menu.Enabled = True
'Unload Me
End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
Unload Me
End Sub

Private Sub Form_Load()
frmWelcome.Caption = UserName

On Error Resume Next
    Me.Left = Screen.Width - (Me.Width + 50)
    Me.Top = Screen.Height - 450 '450 assumed height for taskbar
    Picture1.Visible = False
End Sub



