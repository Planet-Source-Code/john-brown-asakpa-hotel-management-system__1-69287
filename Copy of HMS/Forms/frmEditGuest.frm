VERSION 5.00
Begin VB.Form frmEditGuest 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Edit Guest's Information"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7875
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   7875
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtRoomNo 
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
      Locked          =   -1  'True
      TabIndex        =   34
      Top             =   5760
      Width           =   2655
   End
   Begin VB.TextBox txtRoomType 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   33
      Top             =   5295
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
      TabIndex        =   30
      Top             =   6480
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
      TabIndex        =   29
      Top             =   6120
      Width           =   2655
   End
   Begin VB.TextBox txtTOChkIn 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   5010
      Width           =   2655
   End
   Begin VB.TextBox txtDOChkin 
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
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   4635
      Width           =   2655
   End
   Begin VB.ComboBox cboGuID 
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
      TabIndex        =   0
      Top             =   900
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
      Height          =   330
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   2880
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
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   2520
      Width           =   2655
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
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1440
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
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1800
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
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   3240
      Width           =   2655
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
      Height          =   330
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   3555
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
      Height          =   360
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   2160
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
      Height          =   330
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   3915
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
      Height          =   330
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   4275
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   5760
      TabIndex        =   15
      Top             =   2760
      Width           =   1695
      Begin VB.CommandButton cmdUpdate 
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
         Height          =   495
         Left            =   240
         TabIndex        =   13
         Top             =   840
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
         TabIndex        =   14
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton cmdEdit 
         BackColor       =   &H00FF0000&
         Caption         =   "Edit"
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
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Room Type"
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
      Left            =   120
      TabIndex        =   36
      Top             =   5520
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Select New RoomNo"
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
      Left            =   120
      TabIndex        =   35
      Top             =   6600
      Width           =   1920
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Room No."
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
      Left            =   120
      TabIndex        =   32
      Top             =   5880
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select Room Type"
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
      Left            =   120
      TabIndex        =   31
      Top             =   6240
      Width           =   1740
   End
   Begin VB.Label lblCity 
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
      Left            =   120
      TabIndex        =   28
      Top             =   2280
      Width           =   375
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
      Left            =   120
      TabIndex        =   27
      Top             =   960
      Width           =   870
   End
   Begin VB.Label lblCountry 
      AutoSize        =   -1  'True
      Caption         =   "Country"
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
      Left            =   120
      TabIndex        =   26
      Top             =   3000
      Width           =   765
   End
   Begin VB.Label lblState 
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
      Left            =   120
      TabIndex        =   25
      Top             =   2640
      Width           =   510
   End
   Begin VB.Label lblname 
      AutoSize        =   -1  'True
      Caption         =   "Name of Guest"
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
      Left            =   120
      TabIndex        =   24
      Top             =   1560
      Width           =   1425
   End
   Begin VB.Label lbladdress1 
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
      Left            =   120
      TabIndex        =   23
      Top             =   1920
      Width           =   795
   End
   Begin VB.Label lblCompany 
      AutoSize        =   -1  'True
      Caption         =   "Company"
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
      Left            =   120
      TabIndex        =   22
      Top             =   3360
      Width           =   900
   End
   Begin VB.Label lblDesignation 
      AutoSize        =   -1  'True
      Caption         =   "Designation"
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
      Left            =   120
      TabIndex        =   21
      Top             =   3720
      Width           =   1155
   End
   Begin VB.Label lblPhone 
      AutoSize        =   -1  'True
      Caption         =   "Telephone/Mobile"
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
      Left            =   120
      TabIndex        =   20
      Top             =   4080
      Width           =   1770
   End
   Begin VB.Label lblAdvance 
      AutoSize        =   -1  'True
      Caption         =   "Advance"
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
      Left            =   120
      TabIndex        =   19
      Top             =   4440
      Width           =   840
   End
   Begin VB.Label lblDOChk 
      AutoSize        =   -1  'True
      Caption         =   "Date Of CheckIn"
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
      Left            =   120
      TabIndex        =   18
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Label lblTOChkO 
      AutoSize        =   -1  'True
      Caption         =   "Time Of CheckIn"
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
      Left            =   120
      TabIndex        =   17
      Top             =   5160
      Width           =   1605
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
      Caption         =   "EDIT GUEST's INFORMATION"
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
      TabIndex        =   16
      Top             =   240
      Width           =   3915
   End
End
Attribute VB_Name = "frmEditGuest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Variable Declaration
Dim StrSql As String
Dim i As Integer
Public Sub RoomAdd2()
If txtRoomType.Text = "Single Room" Then
   RS_SingleRoom.AddNew
   RS_SingleRoom.Fields(0) = txtRoomNo.Text
   RS_SingleRoom.Update
  End If
  If txtRoomType.Text = "Double Room" Then
  RS_DoubleRoom.AddNew
  RS_DoubleRoom.Fields(0) = txtRoomNo.Text
  RS_DoubleRoom.Update
 End If
 If txtRoomType.Text = "Suite Room" Then
  RS_SuiteRoom.AddNew
  RS_SuiteRoom.Fields(0) = txtRoomNo.Text
  RS_SuiteRoom.Update
 End If
 If txtRoomType.Text = "Deluxe Suite" Then
  RS_DeluxeSuite.AddNew
  RS_DeluxeSuite.Fields(0) = txtRoomNo.Text
  RS_DeluxeSuite.Update
 End If

End Sub
Public Sub RoomAdd1()

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
       
  '.Update
 
End Sub

Public Sub LCK()
    txtname.Locked = False
    txtaddress.Locked = False
    txtcity.Locked = False
    txtState.Locked = False
    txtCountry.Locked = False
    txtcompany.Locked = False
    txtdesignation.Locked = False
    txttelephone.Locked = False
    'txtRoomType.Locked = False
    'txtRoomNo.Locked = False
    txtadvance.Locked = False
    txtDOChkin.Locked = False
    txtTOChkIn.Locked = False
    
End Sub
Public Sub Blank()
    cboGuID.ListIndex = -1
    txtname.Text = ""
    txtaddress.Text = ""
    txtcity.Text = ""
    txtState.Text = ""
    txtCountry.Text = ""
    txtcompany.Text = ""
    txtdesignation.Text = ""
    txttelephone.Text = ""
    txtRoomType.Text = ""
    txtRoomNo.Text = ""
    txtadvance.Text = ""
    txtDOChkin.Text = ""
    txtTOChkIn.Text = ""
    
End Sub

Private Sub cboGuID_Click()
If cboGuID.ListIndex = -1 Then
 cmdEdit.Enabled = False
 Exit Sub
End If
 cmdEdit.Enabled = True
 With RS_Edit
  .MoveFirst
  While Not .EOF
   If cboGuID.List(cboGuID.ListIndex) = .Fields(1) Then
    txtDOChkin.Text = .Fields(0)
    txtname.Text = .Fields(2)
    txtaddress.Text = .Fields(3)
    txtcity.Text = .Fields(4)
    txtState.Text = .Fields(5)
    txtCountry.Text = .Fields(6)
    txtTOChkIn.Text = .Fields(7)
    txtcompany.Text = .Fields(8)
    txtdesignation.Text = .Fields(9)
    txttelephone.Text = .Fields(10)
    txtRoomType.Text = .Fields(11)
    txtRoomNo.Text = .Fields(12)
    txtadvance.Text = .Fields(13)
   End If
   .MoveNext
  Wend
 End With

End Sub

Private Sub cboRoomType_Click()
If cboRoomType.ListIndex = 0 Then
 cboRoomNo.Clear
 Label3.Visible = False
 cboRoomNo.Visible = False
 Exit Sub
End If
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
Label3.Visible = True
cboRoomNo.Visible = True
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdEdit_Click()
 cmdUpdate.Enabled = True
 cmdEdit.Enabled = False
 LCK
 Label1.Visible = True
 cboRoomType.Visible = True
 cboRoomType.ListIndex = -1
 End Sub

Private Sub cmdUpdate_Click()

RS_Edit.MoveFirst

If cboRoomType.ListIndex >= 1 And cboRoomNo.ListIndex <> -1 Then
 StrSql = "UPDATE CheckIn_Table SET Name = '" & txtname.Text & "'," _
                & "Address = '" & txtaddress.Text & "'," _
                & "City = '" & txtcity.Text & "'," _
                & "State = '" & txtState.Text & "'," _
                & "Country = '" & txtCountry.Text & "'," _
                & "TOChk = '" & txtTOChkIn.Text & "'," _
                & "Company = '" & txtcompany.Text & "'," _
                & "Designation = '" & txtdesignation.Text & "'," _
                & "Phone = '" & txttelephone.Text & "'," _
                & "Advance = '" & txtadvance.Text & "', " _
                & "DOChk = '" & txtDOChkin.Text & "', " _
                & "RoomType = '" & cboRoomType.List(cboRoomType.ListIndex) & "', " _
                & "RoomNo = '" & cboRoomNo.List(cboRoomNo.ListIndex) & "' " _
                & "WHERE GuID = '" & cboGuID.List(cboGuID.ListIndex) & "';"
        cnn.Execute StrSql
  RS_Edit.Update
  RoomAdd1
  RoomAdd2
    MsgBox "Record is Updated"

ElseIf cboRoomType.ListIndex <= 0 Then
    StrSql = "UPDATE CheckIn_Table SET Name = '" & txtname.Text & "'," _
                & "Address = '" & txtaddress.Text & "'," _
                & "City = '" & txtcity.Text & "'," _
                & "State = '" & txtState.Text & "'," _
                & "Country = '" & txtCountry.Text & "'," _
                & "TOChk = '" & txtTOChkIn.Text & "'," _
                & "Company = '" & txtcompany.Text & "'," _
                & "Designation = '" & txtdesignation.Text & "'," _
                & "Phone = '" & txttelephone.Text & "'," _
                & "Advance = '" & txtadvance.Text & "', " _
                & "DOChk = '" & txtDOChkin.Text & "', " _
                & "RoomType = '" & txtRoomType.Text & "', " _
                & "RoomNo = '" & txtRoomNo.Text & "' " _
                & "WHERE GuID = '" & cboGuID.List(cboGuID.ListIndex) & "';"
        cnn.Execute StrSql
       RS_Edit.Update
       MsgBox "Record is Updated"
                
ElseIf cboRoomType.ListIndex >= 1 And cboRoomNo.ListIndex = -1 Then
    MsgBox "Select Room Type & Room No", vbOKOnly + vbInformation, "Information"
    cboRoomType.SetFocus
    Exit Sub
End If
  
  cmdUpdate.Enabled = False
    
  Blank
  Label1.Visible = False
  cboRoomType.Visible = False
  Label3.Visible = False
  cboRoomNo.Visible = False
  cboGuID.ListIndex = -1
  cboGuID.SetFocus
RS_Edit.Requery
End Sub

Private Sub Form_Load()
Me.Top = 3000
Me.Left = 3000

Connect
 With RS_Edit
  While Not .EOF
   cboGuID.AddItem .Fields(1)
   .MoveNext
  Wend
 End With
 cmdEdit.Enabled = False
 cmdUpdate.Enabled = False
 Label1.Visible = False
 Label3.Visible = False
 cboRoomType.Visible = False
 cboRoomNo.Visible = False
 
 cboRoomType.AddItem "None"
 cboRoomType.AddItem "Single Room"
 cboRoomType.AddItem "Double Room"
 cboRoomType.AddItem "Suite Room"
 cboRoomType.AddItem "Deluxe Suite"
End Sub


