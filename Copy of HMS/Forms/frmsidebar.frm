VERSION 5.00
Begin VB.Form frmsidebar 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Side Bar"
   ClientHeight    =   11595
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3120
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   11595
   ScaleWidth      =   3120
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      Caption         =   "   User Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1095
      Left            =   120
      TabIndex        =   5
      Top             =   8880
      Width           =   2775
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Label18"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   765
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Label17"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   765
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   0
         Picture         =   "frmsidebar.frx":0000
         Top             =   0
         Width           =   240
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00404040&
      Caption         =   "    Today"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   10080
      Width           =   2775
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Date:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00404040&
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   420
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   120
         Picture         =   "frmsidebar.frx":058A
         Top             =   0
         Width           =   240
      End
   End
   Begin VB.Image Image19 
      Height          =   480
      Left            =   120
      Picture         =   "frmsidebar.frx":0914
      Top             =   3120
      Width           =   480
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "Make Payment"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   720
      TabIndex        =   25
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Image Image18 
      Height          =   480
      Left            =   120
      Picture         =   "frmsidebar.frx":15DE
      Top             =   3600
      Width           =   480
   End
   Begin VB.Image Image17 
      Height          =   480
      Left            =   120
      Picture         =   "frmsidebar.frx":22A8
      Top             =   4080
      Width           =   480
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "Process Payroll"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   720
      TabIndex        =   24
      Top             =   3720
      Width           =   1545
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "Make Reservation"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   720
      TabIndex        =   23
      Top             =   4200
      Width           =   1770
   End
   Begin VB.Image Image16 
      Height          =   480
      Left            =   120
      Picture         =   "frmsidebar.frx":2F72
      Top             =   4560
      Width           =   480
   End
   Begin VB.Image Image15 
      Height          =   480
      Left            =   120
      Picture         =   "frmsidebar.frx":3C3C
      Top             =   5040
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   120
      Picture         =   "frmsidebar.frx":4906
      Top             =   5520
      Width           =   480
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "Bar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   720
      TabIndex        =   22
      Top             =   4680
      Width           =   345
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "Daily Evaluation Report"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   720
      TabIndex        =   21
      Top             =   5160
      Width           =   2325
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "Laundry Services"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   720
      TabIndex        =   20
      Top             =   5640
      Width           =   1725
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   720
      TabIndex        =   17
      Top             =   8520
      Width           =   375
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "Log Off"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   720
      TabIndex        =   16
      Top             =   8040
      Width           =   735
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "Change Password"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   720
      TabIndex        =   15
      Top             =   7560
      Width           =   1770
   End
   Begin VB.Image Image14 
      Height          =   480
      Left            =   120
      Picture         =   "frmsidebar.frx":55D0
      Top             =   8400
      Width           =   480
   End
   Begin VB.Image Image13 
      Height          =   480
      Left            =   120
      Picture         =   "frmsidebar.frx":629A
      Top             =   7920
      Width           =   480
   End
   Begin VB.Image Image12 
      Height          =   480
      Left            =   120
      Picture         =   "frmsidebar.frx":6F64
      Top             =   7440
      Width           =   480
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "Notepad"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   720
      TabIndex        =   14
      Top             =   7080
      Width           =   825
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "Calculator"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   720
      TabIndex        =   13
      Top             =   6600
      Width           =   1005
   End
   Begin VB.Image Image11 
      Height          =   480
      Left            =   120
      Picture         =   "frmsidebar.frx":7C2E
      Top             =   6960
      Width           =   480
   End
   Begin VB.Image Image10 
      Height          =   480
      Left            =   120
      Picture         =   "frmsidebar.frx":88F8
      Top             =   6480
      Width           =   480
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "Control Panel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   720
      TabIndex        =   12
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "Search Employee"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   720
      TabIndex        =   11
      Top             =   2760
      Width           =   1710
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "Status Of Hotel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   720
      TabIndex        =   10
      Top             =   2280
      Width           =   1515
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "Search Guest"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   720
      TabIndex        =   9
      Top             =   1800
      Width           =   1320
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "Check Out"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   720
      TabIndex        =   8
      Top             =   1320
      Width           =   1020
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "Check In"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   720
      TabIndex        =   7
      Top             =   840
      Width           =   870
   End
   Begin VB.Image Image9 
      Height          =   480
      Left            =   120
      Picture         =   "frmsidebar.frx":95C2
      Top             =   6000
      Width           =   480
   End
   Begin VB.Image Image8 
      Height          =   480
      Left            =   120
      Picture         =   "frmsidebar.frx":A28C
      Top             =   2640
      Width           =   480
   End
   Begin VB.Image Image7 
      Height          =   480
      Left            =   120
      Picture         =   "frmsidebar.frx":AF56
      Top             =   2160
      Width           =   480
   End
   Begin VB.Image Image6 
      Height          =   480
      Left            =   120
      Picture         =   "frmsidebar.frx":BC20
      Top             =   1680
      Width           =   480
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   120
      Picture         =   "frmsidebar.frx":C8EA
      Top             =   1200
      Width           =   480
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   120
      Picture         =   "frmsidebar.frx":D5B4
      Top             =   720
      Width           =   480
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "PICK A TASK"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   480
      TabIndex        =   6
      Top             =   240
      Width           =   1620
   End
End
Attribute VB_Name = "frmsidebar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "Align Left"
            'ToDo: Add 'Align Left' button code.
            MsgBox "Add 'Align Left' button code."
        Case "Align Right"
            'ToDo: Add 'Align Right' button code.
            MsgBox "Add 'Align Right' button code."
        Case "Bold"
            'ToDo: Add 'Bold' button code.
            MsgBox "Add 'Bold' button code."
        Case "Button"
            'ToDo: Add 'Button' button code.
            MsgBox "Add 'Button' button code."
        Case "Center"
            'ToDo: Add 'Center' button code.
            MsgBox "Add 'Center' button code."
        Case "Copy"
            'ToDo: Add 'Copy' button code.
            MsgBox "Add 'Copy' button code."
        Case "Cut"
            'ToDo: Add 'Cut' button code.
            MsgBox "Add 'Cut' button code."
        Case "Italic"
            'ToDo: Add 'Italic' button code.
            MsgBox "Add 'Italic' button code."
        Case "Justify"
            'ToDo: Add 'Justify' button code.
            MsgBox "Add 'Justify' button code."
    End Select
End Sub

Private Sub Command2_Click()
On Error GoTo Err
Shell "calc.exe", vbNormalFocus
Exit Sub
Err:
    MsgBox "You don't have a Calculator installed in your computer.", vbExclamation, "CSRS version 1"

End Sub

Private Sub Command3_Click()
On Error GoTo Err
Shell "notepad.exe", vbNormalFocus
Exit Sub
Err:
    MsgBox "You don't have a NotePad  installed in your computer.", vbExclamation, "CSRS version 1"

End Sub

Private Sub Command4_Click()
frmChangePassword.Show
End Sub


Private Sub Form_Load()

Me.Left = 0
Me.Top = 0
Label2.Caption = Format(Date, "long date")
Label1.Caption = Format(Now, "long time")
Label18.Caption = "You are logged on as: "
Label17.Caption = UserName
End Sub

Private Sub Form_Unload(Cancel As Integer)
Exit Sub
End Sub

Private Sub Image10_Click()
Label12_Click
End Sub

Private Sub Image11_Click()
Label13_Click
End Sub

Private Sub Image12_Click()
Label14_Click
End Sub

Private Sub Image13_Click()
Label15_Click
End Sub

Private Sub Image14_Click()
Label16_Click
End Sub

Private Sub Image15_Click()
frmReport.Show 1
End Sub

Private Sub Image16_Click()
frmBar.Show 1
End Sub

Private Sub Image4_Click()
Label6_Click
End Sub

Private Sub Image5_Click()
Label7_Click
End Sub

Private Sub Image6_Click()
Label8_Click
End Sub

Private Sub Image7_Click()
Label9_Click
End Sub

Private Sub Image8_Click()
Label10_Click
End Sub

Private Sub Image9_Click()
Label11_Click
End Sub

Private Sub Label10_Click()
frmSearchEmp.Show 1
End Sub

Private Sub Label11_Click()
Shell "c:\windows\system32\control.exe", vbNormalFocus
End Sub

Private Sub Label12_Click()
Shell "calc.exe", vbNormalFocus
End Sub

Private Sub Label13_Click()
Shell "notepad", vbNormalFocus
End Sub

Private Sub Label14_Click()
frmChangePassword.Show 1
End Sub

Private Sub Label15_Click()
If MsgBox("Are you sure you want to Log Off User?", vbYesNo) = vbYes Then
MDIMain.Hide
frmLogin.Show
End If
End Sub

Private Sub Label16_Click()
If MsgBox("Are you sure you want to quit?", vbInformation + vbYesNo) = vbYes Then
MsgBox "Thank you for using this application  " & login, vbOKOnly, "  " & Date & " " & Time
 Else: Exit Sub
End If
With RS_Userlog
       .AddNew
       .Fields(0) = UserName
       .Fields(1) = "Log Out"
       .Fields(2) = Date
       .Fields(3) = Time
       .Fields(4) = "Successful"
       .Update
    End With
 End

End Sub

Private Sub Label20_Click()
frmReport.Show 1
End Sub

Private Sub Label21_Click()
Image16_Click
End Sub

Private Sub Label6_Click()
frmcheckIn.Show 1
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.MousePointer = hourglass
End Sub

Private Sub Label7_Click()
frmCheckOut.Show 1
End Sub

Private Sub Label8_Click()
frmSearchGuest.Show 1
End Sub

Private Sub Label9_Click()
frmStatus.Show 1
End Sub
