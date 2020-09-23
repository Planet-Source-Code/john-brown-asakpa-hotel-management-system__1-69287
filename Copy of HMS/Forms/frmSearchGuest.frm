VERSION 5.00
Begin VB.Form frmSearchGuest 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Searching The Guest"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10275
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   10275
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "View all"
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
      Left            =   6600
      TabIndex        =   34
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   6
      Top             =   840
      Width           =   1095
   End
   Begin VB.Frame fraDetails 
      Caption         =   "Other Details"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   4935
      Left            =   2520
      TabIndex        =   5
      Top             =   1560
      Width           =   7095
      Begin VB.TextBox Text13 
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
         Left            =   2520
         TabIndex        =   33
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox Text12 
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
         Left            =   2520
         TabIndex        =   32
         Top             =   4440
         Width           =   2415
      End
      Begin VB.TextBox Text11 
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
         Left            =   2520
         TabIndex        =   31
         Top             =   4080
         Width           =   2415
      End
      Begin VB.TextBox Text10 
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
         Left            =   2520
         TabIndex        =   30
         Top             =   3720
         Width           =   2415
      End
      Begin VB.TextBox Text9 
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
         Left            =   2520
         TabIndex        =   29
         Top             =   3360
         Width           =   2415
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
         Left            =   2520
         TabIndex        =   28
         Top             =   3000
         Width           =   2415
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
         Left            =   5040
         TabIndex        =   27
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox Text6 
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
         Left            =   2520
         TabIndex        =   26
         Top             =   2640
         Width           =   2415
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
         Left            =   2520
         TabIndex        =   25
         Top             =   2280
         Width           =   2415
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
         Left            =   2520
         TabIndex        =   24
         Top             =   1920
         Width           =   2415
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
         Left            =   2520
         TabIndex        =   23
         Top             =   1560
         Width           =   2415
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
         Left            =   2520
         TabIndex        =   22
         Top             =   1200
         Width           =   3975
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
         Left            =   2520
         TabIndex        =   21
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label lblCompany 
         AutoSize        =   -1  'True
         Caption         =   "Company"
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
         Left            =   360
         TabIndex        =   20
         Top             =   2640
         Width           =   915
      End
      Begin VB.Label lblDesignation 
         AutoSize        =   -1  'True
         Caption         =   "Designation"
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
         Left            =   360
         TabIndex        =   19
         Top             =   3000
         Width           =   1170
      End
      Begin VB.Label lblPhone 
         AutoSize        =   -1  'True
         Caption         =   "Telephone/Mobile"
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
         Left            =   360
         TabIndex        =   18
         Top             =   3360
         Width           =   1785
      End
      Begin VB.Label lbltime 
         AutoSize        =   -1  'True
         Caption         =   "Time"
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
         Left            =   4440
         TabIndex        =   17
         Top             =   360
         Width           =   480
      End
      Begin VB.Label lblRoomType 
         AutoSize        =   -1  'True
         Caption         =   "Room Type"
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
         Left            =   360
         TabIndex        =   16
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label lblRoomNo 
         AutoSize        =   -1  'True
         Caption         =   "Room No."
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
         Left            =   360
         TabIndex        =   15
         Top             =   4080
         Width           =   945
      End
      Begin VB.Label lblAdvance 
         AutoSize        =   -1  'True
         Caption         =   "Advance(If Any)"
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
         Left            =   360
         TabIndex        =   14
         Top             =   4440
         Width           =   1650
      End
      Begin VB.Label lbldate 
         AutoSize        =   -1  'True
         Caption         =   "Date Of CheckIn"
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
         Left            =   360
         TabIndex        =   13
         Top             =   360
         Width           =   1635
      End
      Begin VB.Label lblCity 
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
         Left            =   360
         TabIndex        =   12
         Top             =   1560
         Width           =   390
      End
      Begin VB.Label lbladdress1 
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
         Left            =   360
         TabIndex        =   11
         Top             =   1200
         Width           =   795
      End
      Begin VB.Label lblState 
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
         Left            =   360
         TabIndex        =   10
         Top             =   1920
         Width           =   525
      End
      Begin VB.Label lblCountry 
         AutoSize        =   -1  'True
         Caption         =   "Country"
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
         Left            =   360
         TabIndex        =   9
         Top             =   2280
         Width           =   780
      End
      Begin VB.Label lblGuestID 
         AutoSize        =   -1  'True
         Caption         =   "Guest ID"
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
         Left            =   360
         TabIndex        =   8
         Top             =   840
         Width           =   870
      End
   End
   Begin VB.ListBox lstGuestName 
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
      TabIndex        =   4
      Top             =   1680
      Width           =   2175
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   3
      Top             =   840
      Width           =   855
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
      Left            =   3000
      TabIndex        =   2
      Top             =   840
      Width           =   2415
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
      Left            =   3120
      TabIndex        =   7
      Top             =   1200
      Width           =   60
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "Enter Guest's Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1950
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   9840
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Searching The Guests"
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
      Left            =   2400
      TabIndex        =   0
      Top             =   120
      Width           =   3525
   End
End
Attribute VB_Name = "frmSearchGuest"
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
    Dim strSQLTitles As String
    
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
    strSQLTitles = "SELECT Name FROM CheckIn_Table"
    rstName.Open strSQLTitles, Cnxn, adOpenStatic, adLockReadOnly, adCmdText

    count = 0
    rstName.Find "Name LIKE '" & txtName.Text & "%'"
    Do While Not rstName.EOF
        'continue if last find succeeded
       lstGuestName.AddItem rstName!Name
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

lstGuestName.Clear
fraDetails.Visible = False

If txtName.Text = "" Then
 MsgBox "Enter the name", vbOKOnly + vbCritical, "Error"
 txtName.SetFocus
 Exit Sub
End If

FindX


End Sub

Private Sub Command1_Click()


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
    strSQLEmpName = "SELECT * FROM Checkin_Table"
    rstName.Open strSQLEmpName, Cnxn, adOpenStatic, adLockReadOnly, adCmdText

    count = 0
'    rstName.Find "Name "
    With rstName
    .MoveFirst
    While Not .EOF
    lstGuestName.AddItem .Fields(2)
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


Private Sub Form_Load()
fraDetails.Visible = False
Me.Top = 3000
Me.Left = 3000

Connect


lstGuestName.Clear

End Sub

Private Sub lstGuestName_Click()

With Rs_Detail
 .MoveFirst
 While Not .EOF
  If lstGuestName.List(lstGuestName.ListIndex) = .Fields(2) Then
   Text13.Text = .Fields(0)
   Text1.Text = .Fields(1)
   Text2.Text = .Fields(3)
   Text3.Text = .Fields(4)
   Text4.Text = .Fields(5)
   Text5.Text = .Fields(6)
   Text6.Text = .Fields(8)
   Text7.Text = .Fields(7)
   Text8.Text = .Fields(9)
   Text9.Text = .Fields(10)
   Text10.Text = .Fields(11)
   Text11.Text = .Fields(12)
   Text12.Text = .Fields(13)
   Text1.Locked = True
   Text2.Locked = True
   Text3.Locked = True
   Text4.Locked = True
   Text5.Locked = True
   Text6.Locked = True
   Text7.Locked = True
   Text8.Locked = True
   Text9.Locked = True
   Text10.Locked = True
   Text11.Locked = True
   Text12.Locked = True
   Text13.Locked = True
  End If
  .MoveNext
 Wend
   
  fraDetails.Visible = True
 
End With

End Sub
