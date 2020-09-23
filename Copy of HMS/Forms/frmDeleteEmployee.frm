VERSION 5.00
Begin VB.Form frmDeleteEmployee 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Delete Employee Record"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8160
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtDOJ 
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
      TabIndex        =   27
      Top             =   4680
      Width           =   3135
   End
   Begin VB.TextBox txtDepartment 
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
      TabIndex        =   25
      Top             =   5400
      Width           =   3135
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
      Height          =   645
      Left            =   2640
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   3960
      Width           =   3135
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
      TabIndex        =   11
      Top             =   5040
      Width           =   3135
   End
   Begin VB.ComboBox cboID 
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
      TabIndex        =   0
      Top             =   1080
      Width           =   2415
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
      TabIndex        =   10
      Top             =   5760
      Width           =   3135
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
      TabIndex        =   9
      Top             =   1800
      Width           =   3135
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
      TabIndex        =   8
      Top             =   2160
      Width           =   3135
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
      TabIndex        =   7
      Top             =   2520
      Width           =   3135
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
      TabIndex        =   6
      Top             =   2880
      Width           =   3135
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
      Width           =   3135
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
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   6240
      TabIndex        =   3
      Top             =   2520
      Width           =   1695
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00FF0000&
         Caption         =   "Delete"
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
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
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
         TabIndex        =   2
         Top             =   960
         Width           =   1215
      End
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      Caption         =   "Date"
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
      Left            =   5520
      TabIndex        =   28
      Top             =   960
      Width           =   465
   End
   Begin VB.Label lblDOJ 
      AutoSize        =   -1  'True
      Caption         =   "Date Of Joining"
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
      TabIndex        =   26
      Top             =   4680
      Width           =   1530
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "DELETE EMPLOYEE RECORD"
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
      TabIndex        =   24
      Top             =   240
      Width           =   3750
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   8160
      Y1              =   840
      Y2              =   840
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
      TabIndex        =   23
      Top             =   5040
      Width           =   1290
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
      TabIndex        =   22
      Top             =   3960
      Width           =   1740
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
      TabIndex        =   21
      Top             =   5400
      Width           =   1170
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
      TabIndex        =   20
      Top             =   5760
      Width           =   2100
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
      TabIndex        =   18
      Top             =   2160
      Width           =   795
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
      TabIndex        =   17
      Top             =   2520
      Width           =   390
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
      TabIndex        =   16
      Top             =   2880
      Width           =   525
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
      TabIndex        =   15
      Top             =   3600
      Width           =   1380
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
      TabIndex        =   14
      Top             =   1080
      Width           =   1350
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
      TabIndex        =   13
      Top             =   3240
      Width           =   870
   End
End
Attribute VB_Name = "frmDeleteEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Variable Declaration
Dim i, J As Integer

Public Sub LCK()
 txtemployeename.Locked = True
 txtaddress.Locked = True
 txtcity.Locked = True
 txtState.Locked = True
 txtpin.Locked = True
 txtphone.Locked = True
 txteduqualification.Locked = True
 txtdesignation.Locked = True
 txtDepartment.Locked = True
 txtExp.Locked = True
 txtDOJ.Locked = True
End Sub

Private Sub cboID_Click()
  With RS_Delete
    .MoveLast
    .MoveFirst
    While .EOF = False
     If cboID.List(cboID.ListIndex) = .Fields(0) Then
      txtemployeename.Text = .Fields(1)
      txtaddress.Text = .Fields(2)
      txtcity.Text = .Fields(3)
      txtState.Text = .Fields(4)
      txtpin.Text = .Fields(5)
      txtphone.Text = .Fields(6)
      txtDOJ.Text = .Fields(7)
      txteduqualification.Text = .Fields(8)
      txtdesignation.Text = .Fields(9)
      txtDepartment.Text = .Fields(10)
      txtExp.Text = .Fields(11)
      .MoveNext
     Else
      .MoveNext
     End If
    Wend
  End With
  cmdDelete.Enabled = True
End Sub

Private Sub cmdCancel_Click()
Unload Me
Me.Visible = False
End Sub

Private Sub cmdDelete_Click()

 With RS_Enter
    .AddNew
     .Fields(0) = cboID.List(cboID.ListIndex)
     .Fields(1) = txtemployeename.Text
     .Fields(2) = txtaddress.Text
     .Fields(3) = txtcity.Text
     .Fields(4) = txtState.Text
     .Fields(5) = txtpin.Text
     .Fields(6) = txtphone.Text
     .Fields(7) = txteduqualification.Text
     .Fields(8) = txtdesignation.Text
     .Fields(9) = txtDepartment.Text
     .Fields(10) = txtExp.Text
     .Fields(11) = txtDOJ.Text
     .Fields(12) = lblDate.Caption
     .Update
 End With
 
 txtemployeename.Text = ""
 txtaddress.Text = ""
 txtcity.Text = ""
 txtState.Text = ""
 txtpin.Text = ""
 txtphone.Text = ""
 txteduqualification.Text = ""
 txtdesignation.Text = ""
 txtDepartment.Text = ""
 txtExp.Text = ""
 txtDOJ.Text = ""
 
 With RS_Delete
  .MoveLast
  .MoveFirst
  While Not .EOF
    If cboID.List(cboID.ListIndex) = .Fields(0) Then
     .Delete
     cboID.RemoveItem cboID.ListIndex
     MsgBox "Record Is Deleted"
     cmdDelete.Enabled = False
     cboID.SetFocus
     Exit Sub
    Else
     .MoveNext
     End If
   Wend
   End With
 End Sub


Private Sub Form_Load()

Me.Top = 3000
Me.Left = 3000

lblDate.Caption = Date

Call Connect


cmdDelete.Enabled = False
LCK
With RS_Delete
 If .RecordCount = 0 Then
  MsgBox "There are no records"
 ' Unload Me
  Exit Sub
 Else
  .MoveLast
  i = .RecordCount
   .MoveFirst
 For J = 1 To i
  cboID.AddItem .Fields(0)
  .MoveNext
  Next
 End If
 End With
End Sub


