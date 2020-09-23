VERSION 5.00
Begin VB.Form frmEmployeeAdd 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "New Employee"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8475
   ControlBox      =   0   'False
   Icon            =   "frmEmployeeAdd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   8475
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   6720
      TabIndex        =   26
      Top             =   3000
      Width           =   1575
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   11
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00FF0000&
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.TextBox txtpin 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      MaxLength       =   6
      TabIndex        =   4
      Top             =   3240
      Width           =   3015
   End
   Begin VB.TextBox txtphone 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      TabIndex        =   5
      Top             =   3600
      Width           =   3015
   End
   Begin VB.TextBox txtstate 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      TabIndex        =   3
      Top             =   2880
      Width           =   3015
   End
   Begin VB.TextBox txtcity 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      TabIndex        =   2
      Top             =   2520
      Width           =   3015
   End
   Begin VB.TextBox txtaddress 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      TabIndex        =   1
      Top             =   2160
      Width           =   3615
   End
   Begin VB.TextBox txtemployeename 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      TabIndex        =   0
      Top             =   1800
      Width           =   3015
   End
   Begin VB.TextBox txtExp 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   5280
      Width           =   3015
   End
   Begin VB.ComboBox cboDepartment 
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
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   4920
      Width           =   3015
   End
   Begin VB.TextBox txtdesignation 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2640
      TabIndex        =   7
      Top             =   4560
      Width           =   3015
   End
   Begin VB.TextBox txteduqualification 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   2640
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   3960
      Width           =   3015
   End
   Begin VB.Label lblpin 
      AutoSize        =   -1  'True
      Caption         =   "Pin Code"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   480
      TabIndex        =   25
      Top             =   3240
      Width           =   870
   End
   Begin VB.Label lblemployeeNo 
      AutoSize        =   -1  'True
      Caption         =   "Employee No."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   480
      TabIndex        =   24
      Top             =   1200
      Width           =   1350
   End
   Begin VB.Label lblPhone 
      AutoSize        =   -1  'True
      Caption         =   "Phone/Mobile"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   480
      TabIndex        =   23
      Top             =   3600
      Width           =   1380
   End
   Begin VB.Label lblstate 
      AutoSize        =   -1  'True
      Caption         =   "State"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   480
      TabIndex        =   22
      Top             =   2880
      Width           =   525
   End
   Begin VB.Label lblcity 
      AutoSize        =   -1  'True
      Caption         =   "City"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   480
      TabIndex        =   21
      Top             =   2520
      Width           =   390
   End
   Begin VB.Label lbladdress 
      AutoSize        =   -1  'True
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   480
      TabIndex        =   20
      Top             =   2160
      Width           =   795
   End
   Begin VB.Label lblemployeename 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   480
      TabIndex        =   19
      Top             =   1800
      Width           =   570
   End
   Begin VB.Label lblID 
      AutoSize        =   -1  'True
      Caption         =   "ID"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   2640
      TabIndex        =   18
      Top             =   1200
      Width           =   240
   End
   Begin VB.Label lblExp 
      AutoSize        =   -1  'True
      Caption         =   "Experience Summary"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   480
      TabIndex        =   17
      Top             =   5280
      Width           =   2100
   End
   Begin VB.Label lblDepartment 
      AutoSize        =   -1  'True
      Caption         =   "Department"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   480
      TabIndex        =   16
      Top             =   4920
      Width           =   1170
   End
   Begin VB.Label lblqualification 
      AutoSize        =   -1  'True
      Caption         =   "Edu. Qualification"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   480
      TabIndex        =   15
      Top             =   3960
      Width           =   1740
   End
   Begin VB.Label lbldesignation 
      AutoSize        =   -1  'True
      Caption         =   "Appointed As"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   480
      TabIndex        =   14
      Top             =   4560
      Width           =   1290
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   8400
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   4800
      TabIndex        =   13
      Top             =   1200
      Width           =   510
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "ADD EMPLOYEE RECORD"
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
      TabIndex        =   12
      Top             =   240
      Width           =   3300
   End
End
Attribute VB_Name = "frmEmployeeAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Variable Declaration
Dim str As String
Dim strMonth, result As String
Dim num, i As Integer

Public Sub Blank()
    txtemployeename.Text = ""
    txtaddress.Text = ""
    txtcity.Text = ""
    txtState.Text = ""
    txtpin.Text = ""
    txtphone.Text = ""
    txteduqualification.Text = ""
    txtdesignation.Text = ""
    txtExp.Text = ""
    cboDepartment.ListIndex = -1
End Sub



Private Sub cmdAdd_Click()

If txtemployeename.Text = "" Or txtaddress.Text = "" Or _
   txtcity.Text = "" Or txtState.Text = "" Or _
   txtpin.Text = "" Or txtphone.Text = "" Or _
   txteduqualification.Text = "" Or txtExp.Text = "" Then
   MsgBox "Fill Complete Information", vbInformation, "Information"
   txtemployeename.SetFocus
   Exit Sub
ElseIf cboDepartment.ListIndex = -1 Then
   MsgBox "Select the department", vbInformation, "Information"
   cboDepartment.SetFocus
   Exit Sub
Else
   With rs
     .AddNew
     .Fields(0) = lblID.Caption
     .Fields(1) = txtemployeename.Text
     .Fields(2) = txtaddress.Text
     .Fields(3) = txtcity.Text
     .Fields(4) = txtState.Text
     .Fields(5) = txtpin.Text
     .Fields(6) = txtphone.Text
     .Fields(7) = lblDate.Caption
     .Fields(8) = txteduqualification.Text
     .Fields(9) = txtdesignation.Text
     .Fields(10) = cboDepartment.List(cboDepartment.ListIndex)
     .Fields(11) = txtExp.Text
     .Update
     
     MsgBox "Record Entered Successfully"
     Blank
     If .RecordCount = 0 Then
        lblID.Caption = str & strMonth & "-" & num
     Else
        .MoveLast
        i = .RecordCount
        num = num + i
        lblID.Caption = str & strMonth & "-" & num
     End If
   End With
End If
   
   

End Sub

Private Sub cmdCancel_Click()
Unload Me
Me.Visible = False
End Sub

Private Sub Form_Load()
Me.Top = 3000
Me.Left = 3000

num = 100
str = "EmpID"
strMonth = Month(Date)
lblDate.Caption = Date

cboDepartment.AddItem "Front Office"
cboDepartment.AddItem "House Keeping"
cboDepartment.AddItem "Food & Beverage"
cboDepartment.AddItem "Security"
cboDepartment.AddItem "Maintenance"
cboDepartment.AddItem "Purchase & Stores"
cboDepartment.AddItem "Sales & Marketing"

Connect
         

Blank

With rs
   If .RecordCount = 0 Then
    lblID.Caption = str & strMonth & "-" & num
   Else
    .MoveLast
    i = .RecordCount
    num = num + i
  lblID.Caption = str & strMonth & "-" & num
   End If
End With
End Sub

