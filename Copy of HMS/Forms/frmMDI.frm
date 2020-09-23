VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMDI 
   BackColor       =   &H80000004&
   Caption         =   "Hotel Management System"
   ClientHeight    =   8565
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10665
   Icon            =   "frmMDI.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H80000018&
      Height          =   1815
      Left            =   0
      ScaleHeight     =   1755
      ScaleWidth      =   10605
      TabIndex        =   1
      Top             =   0
      Width           =   10665
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   2400
         TabIndex        =   7
         Top             =   1080
         Width           =   12135
         Begin VB.Label lblTicker 
            AutoSize        =   -1  'True
            BackColor       =   &H80000018&
            Caption         =   "Label6"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   720
         End
      End
      Begin VB.Timer tmrTicker 
         Interval        =   100
         Left            =   8880
         Top             =   1920
      End
      Begin VB.Image Image2 
         Height          =   1695
         Left            =   16560
         Picture         =   "frmMDI.frx":0442
         Stretch         =   -1  'True
         Top             =   0
         Width           =   2535
      End
      Begin VB.Image Image1 
         Height          =   1815
         Left            =   0
         Picture         =   "frmMDI.frx":967E
         Stretch         =   -1  'True
         Top             =   0
         Width           =   2295
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2400
         TabIndex        =   6
         Top             =   720
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2400
         TabIndex        =   5
         Top             =   480
         Width           =   675
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Old English Text MT"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7560
         TabIndex        =   4
         Top             =   840
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Old English Text MT"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7320
         TabIndex        =   3
         Top             =   600
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Old English Text MT"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   480
         Left            =   6600
         TabIndex        =   2
         Top             =   120
         Width           =   1260
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   4800
      Top             =   3000
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   8190
      Width           =   10665
      _ExtentX        =   18812
      _ExtentY        =   661
      SimpleText      =   "]"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4234
            MinWidth        =   4234
            Object.ToolTipText     =   "Active User"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   4235
            MinWidth        =   4235
            TextSave        =   "04/09/2007"
            Object.ToolTipText     =   "Date"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   4234
            MinWidth        =   4234
            Picture         =   "frmMDI.frx":124C2
            TextSave        =   "5:07 PM"
            Object.ToolTipText     =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
            Text            =   "**This Project is Developed By :- Brown Asakpa Oghenekevbe John**"
            TextSave        =   "**This Project is Developed By :- Brown Asakpa Oghenekevbe John**"
            Object.ToolTipText     =   "Developer's Name"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            TextSave        =   "CAPS"
         EndProperty
      EndProperty
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuGuests 
      Caption         =   "&Guests"
      Begin VB.Menu res 
         Caption         =   "Reservations"
      End
      Begin VB.Menu mnuGuCheckIn 
         Caption         =   "CheckIn"
      End
      Begin VB.Menu mnuGuEdit 
         Caption         =   "Edit Entry"
      End
      Begin VB.Menu payment 
         Caption         =   "Make Payment"
      End
      Begin VB.Menu mnuGuCheckOut 
         Caption         =   "CheckOut"
      End
   End
   Begin VB.Menu mnuEmp 
      Caption         =   "&Employee Detail"
      Begin VB.Menu mnuEmpAdd 
         Caption         =   "Add New Employee"
      End
      Begin VB.Menu mnuEmpEdit 
         Caption         =   "Edit Employee"
      End
      Begin VB.Menu mnuEmpDelete 
         Caption         =   "Delete Employee"
      End
      Begin VB.Menu ser 
         Caption         =   "-"
      End
      Begin VB.Menu pay 
         Caption         =   "Staff Payroll"
      End
   End
   Begin VB.Menu mnuSrch 
      Caption         =   "&Search"
      Begin VB.Menu mnuSrchGuest 
         Caption         =   "Guest"
      End
      Begin VB.Menu mnuSrchEmp 
         Caption         =   "Employee"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "View"
      Begin VB.Menu mnuViewStatus 
         Caption         =   "Status Of Hotel"
      End
      Begin VB.Menu mnuViewCharges 
         Caption         =   "Charges"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "Reports"
      Begin VB.Menu dai 
         Caption         =   "Daily Evaluation"
      End
      Begin VB.Menu mnuReportsGuests 
         Caption         =   "Guests"
      End
      Begin VB.Menu mnuReportsEmployees 
         Caption         =   "Employees"
         Begin VB.Menu id 
            Caption         =   "By Employer ID"
         End
         Begin VB.Menu allemp 
            Caption         =   "All Employee"
         End
      End
   End
   Begin VB.Menu admin 
      Caption         =   "Administrator"
      Begin VB.Menu adduser 
         Caption         =   "Add User"
      End
      Begin VB.Menu viewus 
         Caption         =   "View Users"
      End
      Begin VB.Menu e 
         Caption         =   "-"
      End
      Begin VB.Menu comp 
         Caption         =   "Company Info"
      End
      Begin VB.Menu mnuUtiChangeCharges 
         Caption         =   "Change Charges"
      End
      Begin VB.Menu sepr 
         Caption         =   "-"
      End
      Begin VB.Menu backup 
         Caption         =   "Backup Database"
      End
   End
   Begin VB.Menu mnuUtilities 
      Caption         =   "Utilities"
      Begin VB.Menu mnuUtiChangePass 
         Caption         =   "Change Password"
      End
      Begin VB.Menu se 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUtiCal 
         Caption         =   "Calculator"
      End
      Begin VB.Menu mnuUtiNotepad 
         Caption         =   "Notepad"
      End
   End
   Begin VB.Menu set 
      Caption         =   "Settings"
      Begin VB.Menu sb 
         Caption         =   "Side Bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuLogOff 
         Caption         =   "&Log Off"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu appskin 
         Caption         =   "Application Skin"
         Begin VB.Menu def 
            Caption         =   "Default"
         End
         Begin VB.Menu xpblue 
            Caption         =   "XP Blue"
         End
         Begin VB.Menu winclassic 
            Caption         =   "Win Classic"
         End
         Begin VB.Menu macgrey 
            Caption         =   "Mac Grey"
         End
         Begin VB.Menu liviolet 
            Caption         =   "Light Violet"
         End
         Begin VB.Menu lightbr 
            Caption         =   "Light Brown"
         End
         Begin VB.Menu coolgreen 
            Caption         =   "Cool Green"
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu mnuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "frmMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c As Integer
Dim prevHour As Byte

Private Sub adduser_Click()
frmAddUser.Show 1
End Sub

Private Sub allemp_Click()
DataReport1.Show
End Sub

Private Sub backup_Click()
frmBackup.Show 1
End Sub

Private Sub comp_Click()
frmCompany.Show 1
End Sub

Private Sub coolgreen_Click()
Call select_color_type(3)
sys_color = "3"

End Sub

Private Sub dai_Click()
frmReport.Show 1
End Sub

Private Sub def_Click()
Call select_color_type(0)
sys_color = "0"

End Sub

Private Sub id_Click()
rptEmpReport.Show 1
End Sub

Private Sub lightbr_Click()
Call select_color_type(5)
sys_color = "5"

End Sub

Private Sub liviolet_Click()
Call select_color_type(4)
sys_color = "4"

End Sub

Private Sub macgrey_Click()
Call select_color_type(1)
sys_color = "1"

End Sub

Private Sub MDIForm_Load()
StatusBar.Panels(2).Text = Format(Date, "long date")
Connect
With RS_Company
Label1.Caption = .Fields(0)
Label2.Caption = .Fields(1)
Label3.Caption = .Fields(2)
Label4.Caption = "You are Logged in as  : - " & UserName
Label5.Caption = Now
frmsidebar.Show
frmTip.Show
End With
loadTicker
End Sub


Private Sub mnuEmpAdd_Click()
frmEmployeeAdd.Show 1
End Sub

Private Sub mnuEmpDelete_Click()
frmDeleteEmployee.Show 1
End Sub

Private Sub mnuEmpEdit_Click()
frmEditemployee.Show 1
End Sub

Private Sub mnuExit_Click()
If MsgBox("Are You Sure ?", vbYesNo + vbInformation, "Warning") = vbYes Then
    End
    Unload frmSYSTRAYICON
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

End Sub

Private Sub mnuGuCheckIn_Click()
frmcheckIn.Show 1
End Sub

Private Sub mnuGuCheckOut_Click()
frmCheckOut.Show 1
End Sub




Private Sub mnuGuEdit_Click()
frmEditGuest.Show 1
End Sub

Private Sub mnuHelpAbout_Click()
Forms (12)
End Sub

Private Sub mnuLogOff_Click()
If MsgBox("Are you sure you want to Log Off?", vbYesNo) = vbYes Then
Unload Me
Load fmLogin
fmLogin.Show
Else
Exit Sub
End If
End Sub

Private Sub mnuSrchEmp_Click()
frmSearchEmp.Show 1
End Sub

Private Sub mnuSrchGuest_Click()
frmSearchGuest.Show 1
End Sub

Private Sub mnuUtiCal_Click()
On Error GoTo errHandle
    Dim a As Double
    a = Shell("calc.exe", vbNormalFocus)
    Exit Sub
errHandle:
    MsgBox "Unable to run Calculator Utility on your computer", vbInformation, "Error in opening!!!"
    Resume Next
End Sub



Private Sub mnuUtiChangeCharges_Click()
frmChangeRate.Show 1
End Sub

Private Sub mnuUtiChangePass_Click()
frmChangePassword.Show 1
End Sub

Private Sub mnuUtiNotepad_Click()
On Error GoTo errHandle
    Dim a As Double
    a = Shell("notepad.exe", vbNormalFocus)
    Exit Sub
errHandle:
    MsgBox "Unable to run Notepad Utility on your computer", vbInformation, "Error in opening!!!"
    Resume Next
End Sub

Private Sub mnuViewCharges_Click()
frmCharges.Show 1
End Sub

Private Sub mnuViewStatus_Click()
frmStatus.Show 1
End Sub

Private Sub pay_Click()
frmPayroll.Show 1
End Sub

Private Sub payment_Click()
frmGuestPayment.Show 1
End Sub

Private Sub res_Click()
frmreservations.Show 1
End Sub

Private Sub sb_Click()
If frmsidebar.Visible = True Then
frmsidebar.Visible = False
ElseIf frmsidebar.Visible = False Then
frmsidebar.Visible = True
End If
End Sub

Private Sub Timer1_Timer()
StatusBar.Panels(4).Text = Right(StatusBar.Panels(4).Text, Len(StatusBar.Panels(4).Text) - 1) & Left(StatusBar.Panels(4).Text, 1)
End Sub

Private Sub viewus_Click()
frmViewUsers.Show 1
End Sub

Private Sub winclassic_Click()
Call select_color_type(6)
sys_color = "6"

End Sub

Private Sub xpblue_Click()
Call select_color_type(2)
sys_color = "2"
End Sub
Private Sub loadTicker()
Dim tickSQL As String
tickSQL = " SELECT msgTitle,msgText FROM Ticker "

'Dim rs_ticker As Recordset
'RS_ticker.Open tickSQL, cnn, adOpenDynamic, adLockPessimistic
'First load an array of labels
Dim i As Integer
'Assign recordset to labels
i = 0
While Not RS_ticker.EOF
    On Error Resume Next
    lblTicker(i).Container = frmTick
    lblTicker(i).Visible = False
    lblTicker(i).Caption = RS_ticker("msgTitle") & vbCrLf & RS_ticker("msgText")
    i = i + 1
    Load lblTicker
    RS_ticker.MoveNext
Wend
'RS_ticker.Close
'Set RS_ticker = Nothing
tmrTicker.Enabled = True
c = 0
End Sub
Private Sub tmrTicker_Timer()
Dim currMin As Byte
currMin = Minute(Now())
If currMin > prevmin Then 'Refreshes every hour
    prevmin = currMin
    destroyTicker
    loadTicker
End If
moveTicker 15
End Sub
Private Sub moveTicker(ByVal amt As Integer)
lblTicker(c).ZOrder vbBringToFront
lblTicker(c).Visible = True
lblTicker(c).Top = lblTicker(c).Top - amt

If lblTicker(c).Top < 0 - lblTicker(c).Height Then
    'hide the current lbl
    lblTicker(c).Visible = False
    c = c + 1
    If c > lblTicker.UBound Then
        c = 0
    End If
    tickerStart
End If
End Sub
Private Sub destroyTicker()
Dim i As Integer
For i = lblTicker.LBound To lblTicker.UBound
    'Suppose to free memory
    If i <> 0 Then
        Unload lblTicker(i)
    End If
Next i
End Sub
Private Sub tickerStart()
'Starts the news ticker
'Set the first ticker visible and position on top of frame
lblTicker(c).Visible = True
'lblTicker(c).Top = frmTick.Height
lblTicker(c).Left = 120
tmrTicker.Enabled = True
End Sub

Private Sub tickerStop()
'Stops the news ticker
tmrTicker.Enabled = False
destroyTicker
End Sub

