VERSION 5.00
Begin VB.Form frmreservations 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Reservations"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   8985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      Left            =   2640
      TabIndex        =   31
      Top             =   4440
      Width           =   2655
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
      Height          =   315
      Left            =   2640
      TabIndex        =   30
      Top             =   6240
      Width           =   2655
   End
   Begin VB.TextBox txtadvance 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2640
      TabIndex        =   13
      Top             =   5760
      Width           =   2655
   End
   Begin VB.ComboBox cboRoomNo 
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
      Left            =   2640
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   5280
      Width           =   2655
   End
   Begin VB.ComboBox cboRoomType 
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
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   4800
      Width           =   2655
   End
   Begin VB.TextBox txttelephone 
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
      Left            =   2640
      TabIndex        =   10
      Top             =   4080
      Width           =   2655
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
      Left            =   2640
      TabIndex        =   9
      Top             =   2280
      Width           =   3495
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
      Left            =   2640
      TabIndex        =   8
      Top             =   3720
      Width           =   2655
   End
   Begin VB.TextBox txtcompany 
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
      Left            =   2640
      TabIndex        =   7
      Top             =   3360
      Width           =   2655
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
      Height          =   330
      Left            =   2640
      TabIndex        =   6
      Top             =   1920
      Width           =   3495
   End
   Begin VB.TextBox txtname 
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
      Left            =   2640
      TabIndex        =   5
      Top             =   1560
      Width           =   2655
   End
   Begin VB.TextBox txtState 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2640
      TabIndex        =   4
      Top             =   2580
      Width           =   2655
   End
   Begin VB.TextBox txtCountry 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2640
      TabIndex        =   3
      Top             =   3000
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   5880
      TabIndex        =   0
      Top             =   2760
      Width           =   1575
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00FF0000&
         Caption         =   "Make Reservation"
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
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1335
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
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   1335
      End
   End
   Begin VB.Label lbltime 
      AutoSize        =   -1  'True
      Caption         =   "Time"
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
      Left            =   6360
      TabIndex        =   32
      Top             =   840
      Width           =   525
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Date of Entry"
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
      TabIndex        =   29
      Top             =   6240
      Width           =   1305
   End
   Begin VB.Label lblAdvance 
      AutoSize        =   -1  'True
      Caption         =   "50% Advance"
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
      TabIndex        =   28
      Top             =   5760
      Width           =   1395
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
      TabIndex        =   27
      Top             =   5280
      Width           =   945
   End
   Begin VB.Label lblRoomType 
      AutoSize        =   -1  'True
      Caption         =   "Select Room Type"
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
      TabIndex        =   26
      Top             =   4920
      Width           =   1770
   End
   Begin VB.Label lbldate 
      AutoSize        =   -1  'True
      Caption         =   "Check In Date"
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
      TabIndex        =   25
      Top             =   4440
      Width           =   1395
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      X1              =   0
      X2              =   8640
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label lbltitle 
      AutoSize        =   -1  'True
      Caption         =   "Reservation"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   2640
      TabIndex        =   24
      Top             =   120
      Width           =   1935
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
      TabIndex        =   23
      Top             =   4080
      Width           =   1785
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
      TabIndex        =   22
      Top             =   2280
      Width           =   390
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
      TabIndex        =   21
      Top             =   3720
      Width           =   1170
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
      Top             =   3360
      Width           =   915
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
      TabIndex        =   19
      Top             =   1920
      Width           =   795
   End
   Begin VB.Label lblname 
      AutoSize        =   -1  'True
      Caption         =   "Name of Guest"
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
      Top             =   1560
      Width           =   1455
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
      TabIndex        =   17
      Top             =   2640
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
      TabIndex        =   16
      Top             =   3000
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
      TabIndex        =   15
      Top             =   1200
      Width           =   870
   End
   Begin VB.Label lblGID 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ReID"
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
      Height          =   300
      Left            =   2640
      TabIndex        =   14
      Top             =   1080
      Width           =   570
   End
End
Attribute VB_Name = "frmreservations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim str As String

Private Sub cboRoomNo_Click()
If cboRoomType.ListIndex = -1 Then
 MsgBox "Select The Room Type"
 cboRoomType.SetFocus
 Exit Sub
End If
End Sub

Private Sub cboRoomType_Click()
If cboRoomType.List(cboRoomType.ListIndex) = "Single Room" Then
 cboRoomNo.Clear
 With RS_SingleRoom
 .MoveFirst
  i = 0
  For i = 1 To .RecordCount
   cboRoomNo.AddItem .Fields(0)
   .MoveNext
  Next
 End With
ElseIf cboRoomType.List(cboRoomType.ListIndex) = "Double Room" Then
 cboRoomNo.Clear
 With RS_DoubleRoom
 .MoveFirst
  i = 0
  For i = 1 To .RecordCount
   cboRoomNo.AddItem .Fields(0)
   .MoveNext
  Next
 End With
ElseIf cboRoomType.List(cboRoomType.ListIndex) = "Suite Room" Then
 cboRoomNo.Clear
 With RS_SuiteRoom
 .MoveFirst
  i = 0
  For i = 1 To .RecordCount
   cboRoomNo.AddItem .Fields(0)
   .MoveNext
  Next
 End With
ElseIf cboRoomType.List(cboRoomType.ListIndex) = "Deluxe Suite" Then
 cboRoomNo.Clear
 With RS_DeluxeSuite
 .MoveFirst
  i = 0
  For i = 1 To .RecordCount
   cboRoomNo.AddItem .Fields(0)
   .MoveNext
  Next
 End With
End If
End Sub

Private Sub cmdAdd_Click()
If txtname.Text = "" Or txtaddress.Text = "" Or _
   txtcity.Text = "" Or txtState.Text = "" Or _
   txtCountry.Text = "" Or txtcompany.Text = "" Or _
   txtdesignation.Text = "" Or txttelephone.Text = "" Or _
   txtadvance.Text = "" Then
   MsgBox "Fill The Complete Information"
   txtname.SetFocus
   Exit Sub
 ElseIf cboRoomType.ListIndex = -1 Then
    MsgBox "Select Room Type"
    cboRoomType.SetFocus
    Exit Sub
 ElseIf cboRoomNo.ListIndex = -1 Then
    MsgBox "Select Room No."
    cboRoomNo.SetFocus
    Exit Sub
 End If
 
With RS_Payment
.AddNew
.Fields(0) = lblGID.Caption
.Fields(1) = txtname.Text
.Fields(2) = lblDate.Caption
.Fields(3) = "Accomodation"
.Fields(4) = cboRoomType.List(cboRoomType.ListIndex)
.Fields(5) = cboRoomNo.Text
.Fields(6) = txtadvance.Text
.Fields(7) = 0
.Fields(8) = 0
.Fields(9) = 0
.Fields(10) = Date
.Fields(11) = Date
.Fields(12) = UserName
.Update
End With

With RS_Paymentlog
.AddNew
.Fields(0) = lblGID.Caption
.Fields(1) = txtname.Text
.Fields(2) = lblDate.Caption
.Fields(3) = "Accomodation"
.Fields(4) = cboRoomType.List(cboRoomType.ListIndex)
.Fields(5) = cboRoomNo.Text
.Fields(6) = txtadvance.Text
.Fields(7) = 0
.Fields(8) = 0
.Fields(9) = 0
.Fields(10) = Date
.Fields(11) = UserName
.Update
End With

With RS_Guest
 .AddNew
 .Fields(0) = lblDate.Caption
 .Fields(1) = lblGID.Caption
 .Fields(2) = txtname.Text
 .Fields(3) = txtaddress.Text
 .Fields(4) = txtcity.Text
 .Fields(5) = txtState.Text
 .Fields(6) = txtCountry.Text
 .Fields(7) = lbltime.Caption
 .Fields(8) = txtcompany.Text
 .Fields(9) = txtdesignation.Text
 .Fields(10) = txttelephone.Text
 .Fields(11) = cboRoomType.List(cboRoomType.ListIndex)
 .Fields(12) = cboRoomNo.List(cboRoomNo.ListIndex)
 .Fields(13) = txtadvance.Text
 
 If cboRoomType.List(cboRoomType.ListIndex) = "Single Room" Then
  RS_SingleRoom.MoveFirst
Smart_SingleRoom:
   If cboRoomNo.List(cboRoomNo.ListIndex) = RS_SingleRoom.Fields(0) Then
    RS_SingleRoom.Delete
    cboRoomNo.RemoveItem cboRoomNo.ListIndex
   Else
    RS_SingleRoom.MoveNext
    GoTo Smart_SingleRoom
   End If
   
 ElseIf cboRoomType.List(cboRoomType.ListIndex) = "Double Room" Then
  RS_DoubleRoom.MoveFirst
Smart_DoubleRoom:
   If cboRoomNo.List(cboRoomNo.ListIndex) = RS_DoubleRoom.Fields(0) Then
    RS_DoubleRoom.Delete
    cboRoomNo.RemoveItem cboRoomNo.ListIndex
   Else
    RS_DoubleRoom.MoveNext
    GoTo Smart_DoubleRoom
   End If
   
 ElseIf cboRoomType.List(cboRoomType.ListIndex) = "Suite Room" Then
  RS_SuiteRoom.MoveFirst
Smart_SuiteRoom:
   If cboRoomNo.List(cboRoomNo.ListIndex) = RS_SuiteRoom.Fields(0) Then
    RS_SuiteRoom.Delete
    cboRoomNo.RemoveItem cboRoomNo.ListIndex
   Else
    RS_SuiteRoom.MoveNext
    GoTo Smart_SuiteRoom
   End If
 
 ElseIf cboRoomType.List(cboRoomType.ListIndex) = "Deluxe Suite" Then
  RS_DeluxeSuite.MoveFirst
Smart_DeluxeSuite:
   If cboRoomNo.List(cboRoomNo.ListIndex) = RS_DeluxeSuite.Fields(0) Then
    RS_DeluxeSuite.Delete
    cboRoomNo.RemoveItem cboRoomNo.ListIndex
   Else
    RS_DeluxeSuite.MoveNext
    GoTo Smart_DeluxeSuite
   End If
 
  End If
       
  .Update
 

  txtname.SetFocus
  num = 100 + .RecordCount + 1
   lblGID.Caption = "GUID" + CStr(strMonth) _
                   + "-" + CStr(num) + "-" + CStr(strYear)
End With
   
  MsgBox "Record Entered Successfully"
    Blank
    
    With RS_Userlog
       .AddNew
       .Fields(0) = UserName
       .Fields(1) = "Add Guest Record(Reservation)"
       .Fields(2) = Date
       .Fields(3) = Time
        .Fields(4) = "Successful"
       .Update
    End With

    
End Sub


Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Top = 3000
Me.Left = 3000


num = 100
strMonth = Month(Date)
strYear = Year(Date)
str = "ReID"
lblDate.Caption = Date
lbltime.Caption = Time
cboRoomType.AddItem "Single Room"
cboRoomType.AddItem "Double Room"
cboRoomType.AddItem "Suite Room"
cboRoomType.AddItem "Deluxe Suite"

Call Connect

With RS_Guest
 If .RecordCount = 0 Then
  lblGID.Caption = lblGID.Caption + CStr(strMonth) _
                   + "-" + CStr(num) + "-" + CStr(strYear)
 Else
   num = num + .RecordCount + 1
   lblGID.Caption = str + CStr(strMonth) _
                   + "-" + CStr(num) + "-" + CStr(strYear)
   
 End If
End With
End Sub


Public Sub Blank()
    txtname.Text = ""
    txtaddress.Text = ""
    txtcity.Text = ""
    txtState.Text = ""
    txtCountry.Text = ""
    txtcompany.Text = ""
    txtdesignation.Text = ""
    txttelephone.Text = ""
    cboRoomType.ListIndex = -1
    cboRoomNo.ListIndex = -1
    txtadvance.Text = ""
End Sub

