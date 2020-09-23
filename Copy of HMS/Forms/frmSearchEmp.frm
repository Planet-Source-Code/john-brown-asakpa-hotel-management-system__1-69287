VERSION 5.00
Begin VB.Form frmSearchEmp 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Searching Employee"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10005
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   10005
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "View All"
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
      Left            =   6360
      TabIndex        =   30
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox txtName 
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
      Left            =   2880
      TabIndex        =   4
      Top             =   960
      Width           =   2415
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
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
      TabIndex        =   3
      Top             =   960
      Width           =   855
   End
   Begin VB.ListBox lstEmpName 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3840
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Frame fraDetails 
      Caption         =   "Other Details"
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
      Height          =   5175
      Left            =   2520
      TabIndex        =   1
      Top             =   1920
      Width           =   6735
      Begin VB.TextBox txtDept 
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   3120
         Width           =   2535
      End
      Begin VB.TextBox txtDOJ 
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   3480
         Width           =   2535
      End
      Begin VB.TextBox txteduqualification 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   2400
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   27
         Top             =   3840
         Width           =   2535
      End
      Begin VB.TextBox txtdesignation 
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   2760
         Width           =   2535
      End
      Begin VB.TextBox txtExp 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2400
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         Top             =   4440
         Width           =   2535
      End
      Begin VB.TextBox txtEmpNo 
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox txtaddress 
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   960
         Width           =   3615
      End
      Begin VB.TextBox txtcity 
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox txtstate 
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1680
         Width           =   2535
      End
      Begin VB.TextBox txtphone 
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   2400
         Width           =   2535
      End
      Begin VB.TextBox txtpin 
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
         Left            =   2400
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   19
         Top             =   2040
         Width           =   2535
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         Caption         =   "Date Of Joining"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   3480
         Width           =   1470
      End
      Begin VB.Label lbldesignation 
         AutoSize        =   -1  'True
         Caption         =   "Appointed As"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   2760
         Width           =   1290
      End
      Begin VB.Label lblqualification 
         AutoSize        =   -1  'True
         Caption         =   "Edu. Qualification"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   3840
         Width           =   1710
      End
      Begin VB.Label lblDepartment 
         AutoSize        =   -1  'True
         Caption         =   "Department"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   3120
         Width           =   1155
      End
      Begin VB.Label lblExp 
         AutoSize        =   -1  'True
         Caption         =   "Experience Summary"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   4560
         Width           =   2085
      End
      Begin VB.Label lbladdress 
         AutoSize        =   -1  'True
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   960
         Width           =   795
      End
      Begin VB.Label lblcity 
         AutoSize        =   -1  'True
         Caption         =   "City"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label lblstate 
         AutoSize        =   -1  'True
         Caption         =   "State"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   1680
         Width           =   510
      End
      Begin VB.Label lblPhone 
         AutoSize        =   -1  'True
         Caption         =   "Phone/Mobile"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   2400
         Width           =   1350
      End
      Begin VB.Label lblemployeeNo 
         AutoSize        =   -1  'True
         Caption         =   "Employee No."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lblpin 
         AutoSize        =   -1  'True
         Caption         =   "Pin Code"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   2040
         Width           =   840
      End
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
      Height          =   375
      Left            =   7440
      TabIndex        =   0
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Search Employee Record"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   1680
      TabIndex        =   7
      Top             =   120
      Width           =   3990
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   9600
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "Enter Employee's Name"
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
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   2595
   End
   Begin VB.Label lblcount 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   2880
      TabIndex        =   5
      Top             =   1440
      Width           =   60
   End
End
Attribute VB_Name = "frmSearchEmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim StrSql As String

Public Sub FindX()

    ' connection and recordset variables
    Dim Cnxn As New ADODB.Connection
    Dim rstName As New ADODB.Recordset
    Dim strCnxn As String
    Dim strSQLEmpName As String
    
     ' record variables
    Dim mark As Variant
    Dim count As Integer
    
     ' open connection
     lblcount.Caption = ""
    Set Cnxn = New ADODB.Connection
    strCnxn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" _
              & App.Path & "\database\HMS.mdb;Persist Security Info=False"
              
    Cnxn.Open strCnxn
       
    ' open recordset with default parameters which are
    ' sufficient to search forward through a Recordset
    Set rstName = New ADODB.Recordset
    strSQLEmpName = "SELECT Name FROM PresentEmp_Table"
    rstName.Open strSQLEmpName, Cnxn, adOpenStatic, adLockReadOnly, adCmdText

    count = 0
    rstName.Find "Name LIKE '" & txtName.Text & "%'"
    Do While Not rstName.EOF
        'continue if last find succeeded
       lstEmpName.AddItem rstName!Name
        'count the last title found
       count = count + 1
        ' note current position
       mark = rstName.Bookmark
       rstName.Find "Name LIKE '" & txtName.Text & "%'", 1, adSearchForward, mark
        ' above code skips current record to avoid finding the same row repeatedly;
        ' last arg (bookmark) is redundant because Find searches from current position
    Loop
    If count = 0 Then
     MsgBox "No Match Found", vbOKOnly + vbInformation, "Information"
     txtName.SetFocus
    Else
     lblcount.Caption = "Total Matches found " & count
    End If
     ' clean up
    rstName.Close
    Cnxn.Close
    Set rstName = Nothing
    Set Cnxn = Nothing

End Sub


Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdGo_Click()

lstEmpName.Clear
fraDetails.Visible = False

If txtName.Text = "" Then
 MsgBox "Enter the name", vbOKOnly + vbCritical, "Error"
 txtName.SetFocus
 Exit Sub
End If

FindX


End Sub

Private Sub Command1_Click()
FindAll
End Sub

Private Sub Form_Load()
fraDetails.Visible = False
Me.Top = 3000
Me.Left = 3000

Call Connect


lstEmpName.Clear

End Sub


Private Sub lstEmpName_Click()
With Rs_Details
 .MoveFirst
 While Not .EOF
  If lstEmpName.List(lstEmpName.ListIndex) = .Fields(1) Then
   txtEmpNo.Text = .Fields(0)
   txtaddress.Text = .Fields(2)
   txtcity.Text = .Fields(3)
   txtState.Text = .Fields(4)
   txtpin.Text = .Fields(5)
   txtphone.Text = .Fields(6)
   txtDOJ.Text = .Fields(7)
   txteduqualification.Text = .Fields(8)
   txtdesignation.Text = .Fields(9)
   txtDept.Text = .Fields(10)
   txtExp.Text = .Fields(11)
  End If
  .MoveNext
 Wend
 fraDetails.Visible = True
 End With
End Sub
Public Sub FindAll()

    ' connection and recordset variables
    Dim Cnxn As New ADODB.Connection
    Dim rstName As New ADODB.Recordset
    Dim strCnxn As String
    Dim strSQLEmpName As String
    
     ' record variables
    Dim mark As Variant
    Dim count As Integer
    
     ' open connection
     lblcount.Caption = ""
    Set Cnxn = New ADODB.Connection
    strCnxn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" _
              & App.Path & "\database\HMS.mdb;Persist Security Info=False"
              
    Cnxn.Open strCnxn
       
    ' open recordset with default parameters which are
    ' sufficient to search forward through a Recordset
    Set rstName = New ADODB.Recordset
    strSQLEmpName = "SELECT * FROM PresentEmp_Table"
    rstName.Open strSQLEmpName, Cnxn, adOpenStatic, adLockReadOnly, adCmdText

    count = 0
'    rstName.Find "Name "
    With rstName
    .MoveFirst
    While Not .EOF
    lstEmpName.AddItem .Fields(1)
    .MoveNext
    Wend
     count = .RecordCount
    End With
    
    If count = 0 Then
     MsgBox "No Match Found", vbOKOnly + vbInformation, "Information"
     txtName.SetFocus
    Else
     lblcount.Caption = "Total Matches found " & count
    End If
     ' clean up
    rstName.Close
    Cnxn.Close
    Set rstName = Nothing
    Set Cnxn = Nothing



End Sub
