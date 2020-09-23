VERSION 5.00
Begin VB.Form frmPayroll 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Employee Payroll"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
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
      Height          =   4335
      Left            =   120
      TabIndex        =   9
      Top             =   3000
      Width           =   6855
      Begin VB.CommandButton Command5 
         Caption         =   "Reset"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5400
         TabIndex        =   36
         ToolTipText     =   "Delete Existing Payroll Record from the database"
         Top             =   2640
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5400
         TabIndex        =   35
         ToolTipText     =   "Update an Existing Payroll Record"
         Top             =   2160
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5400
         TabIndex        =   31
         ToolTipText     =   "Save New Payroll Record"
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
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
         Height          =   375
         Left            =   5400
         TabIndex        =   30
         Top             =   3120
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Process"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5400
         TabIndex        =   29
         ToolTipText     =   "Calculate Payroll"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox Text10 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   3480
         Width           =   2175
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   26
         Top             =   3120
         Width           =   2175
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   2760
         Width           =   2175
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   2040
         Width           =   2175
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   20
         Top             =   1680
         Width           =   2175
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   19
         Top             =   1320
         Width           =   2175
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   18
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   17
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         TabIndex        =   16
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label18 
         Caption         =   "Net Pay"
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
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Pension"
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
         Left            =   120
         TabIndex        =   24
         Top             =   3120
         Width           =   750
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Tax"
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
         Left            =   120
         TabIndex        =   23
         Top             =   2760
         Width           =   330
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "DEDUCTIONS"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1680
         TabIndex        =   22
         Top             =   2400
         Width           =   1395
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Gross Pay"
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
         Left            =   120
         TabIndex        =   15
         Top             =   2040
         Width           =   945
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Medical allowance"
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
         Left            =   120
         TabIndex        =   14
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Transport Allowance"
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
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   1920
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Basic Salary"
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
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   1125
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Hourly Rate"
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
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Hours worked"
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
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1320
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Employee Details"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   6855
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2160
         TabIndex        =   32
         Top             =   1800
         Width           =   2415
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2160
         TabIndex        =   2
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label17 
         Caption         =   "Period"
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
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Label7"
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
         Left            =   2160
         TabIndex        =   8
         Top             =   1440
         Width           =   630
      End
      Begin VB.Label Label6 
         Caption         =   "Designation"
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
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Label5"
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
         Left            =   2160
         TabIndex        =   6
         Top             =   1080
         Width           =   630
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Label4"
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
         Left            =   2160
         TabIndex        =   5
         Top             =   720
         Width           =   630
      End
      Begin VB.Label Label3 
         Caption         =   "Department"
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
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Staff Name"
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
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Staff ID"
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
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "PROCESS STAFF PAYROLL"
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
      Left            =   1320
      TabIndex        =   34
      Top             =   240
      Width           =   3480
   End
End
Attribute VB_Name = "frmPayroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim HW, HR, BS, TA, MA, GP As String
Dim TAX, PEN, SA, NP As String

Private Sub Combo1_Click()
With Rs_Details
.MoveFirst
While Not .EOF
If Combo1.List(Combo1.ListIndex) = .Fields(0) Then
Label4.Caption = .Fields(1)
Label5.Caption = .Fields(10)
Label7.Caption = .Fields(9)
End If
.MoveNext
Wend
End With
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text10.Text = ""
End Sub

Private Sub Combo2_Click()
With RS_Payroll
.MoveFirst
While Not .EOF
If Combo2.List(Combo2.ListIndex) = .Fields(4) Then
Text1.Text = .Fields(5)
Text2.Text = .Fields(6)
Text3.Text = .Fields(7)
Text4.Text = .Fields(8)
Text5.Text = .Fields(9)
Text6.Text = .Fields(10)
Text7.Text = .Fields(11)
Text8.Text = .Fields(12)
Text10.Text = .Fields(13)
MsgBox "Record Already Exist for the month"
End If


.MoveNext
Wend
End With
End Sub

Private Sub Command1_Click()
BS = Val(Text3.Text)
TA = Val(Text4.Text)
MA = Val(Text5.Text)
TAX = Val(Text7.Text)
PEN = Val(Text8.Text)
NP = Val(Text10.Text)
GP = BS + TA + MA
Text6.Text = GP
TAX = BS / 100 * 5
Text7.Text = TAX
NP = GP - TAX - Val(Text8.Text)
Text10.Text = NP
End Sub

Private Sub Command2_Click()
Blank
Me.Hide
End Sub

Private Sub Command3_Click()
If Text1.Text = "" Or Text2.Text = "" Or Text4.Text = "" Or Text5.Text = "" Then
MsgBox "Pls Input the Required Fields", vbCritical
Exit Sub
End If

With RS_Payroll
.MoveFirst
While Not .EOF
If Combo1.Text = .Fields(0) Then
    If Combo2.Text = .Fields(4) Then
    MsgBox "U have already Processed Payment for this Month", vbCritical, "Process Error"
    Exit Sub
    End If
End If
.MoveNext
Wend
.AddNew
.Fields(0) = Combo1.Text
.Fields(1) = Label4.Caption
.Fields(2) = Label5.Caption
.Fields(3) = Label7.Caption
.Fields(4) = Combo2.Text
.Fields(5) = Text1.Text
.Fields(6) = Text2.Text
.Fields(7) = Text3.Text
.Fields(8) = Text4.Text
.Fields(9) = Text5.Text
.Fields(10) = Text6.Text
.Fields(11) = Text7.Text
.Fields(12) = Text8.Text
.Fields(13) = Text10.Text
.Fields(14) = Date
.Fields(15) = UserName
.Update
MsgBox "Payroll Record Successfully Updated", vbInformation, "Update Record"
Blank
End With
End Sub

Private Sub Command5_Click()
ResetPay.Show 1
End Sub

Private Sub Form_Load()
Me.Top = 3000
Me.Left = 3000
Connect
With Rs_Details
.MoveFirst
While Not .EOF
Combo1.AddItem .Fields(0)
.MoveNext
Wend
End With
Label4.Caption = ""
Label5.Caption = ""
Label7.Caption = ""

Combo2.AddItem "January"
Combo2.AddItem "February"
Combo2.AddItem "March"
Combo2.AddItem "April"
Combo2.AddItem "May"
Combo2.AddItem "June"
Combo2.AddItem "July"
Combo2.AddItem "August"
Combo2.AddItem "September"
Combo2.AddItem "October"
Combo2.AddItem "November"
Combo2.AddItem "December"
End Sub

Private Sub Text3_GotFocus()
If Text1.Text = "" Or Text2.Text = "" Then
MsgBox "Pls Enter values for the required fields for Basic Salary to be calculate", vbOKOnly
Exit Sub
Else
HW = Val(Text1.Text)
HR = Val(Text2.Text)
BS = Val(HW) * Val(HR)
Text3.Text = BS
End If
End Sub
Sub Blank()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text10.Text = ""
Combo1.ListIndex = -1
Combo2.ListIndex = -1
Label4.Caption = ""
Label5.Caption = ""
Label7.Caption = ""

End Sub
