VERSION 5.00
Begin VB.Form frmCheckOut 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "CheckOut Information"
   ClientHeight    =   7740
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7860
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   7860
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtTotalBill 
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
      TabIndex        =   13
      Top             =   6240
      Width           =   2655
   End
   Begin VB.TextBox txtBalance 
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
      TabIndex        =   12
      Top             =   5880
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
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   6960
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
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   6600
      Width           =   2655
   End
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
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   5160
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
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   4800
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
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   5520
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
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   4440
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
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   2520
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
      Height          =   330
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   4035
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
      Height          =   375
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   3630
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
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   2160
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
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1800
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
      Height          =   300
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   2880
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
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   3240
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
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1320
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   5880
      TabIndex        =   18
      Top             =   3240
      Width           =   1575
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00FF0000&
         Caption         =   "CheckOut"
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
         TabIndex        =   16
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
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   1215
      End
   End
   Begin VB.Label lblTotalBill 
      AutoSize        =   -1  'True
      Caption         =   "Total Bill"
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
      TabIndex        =   37
      Top             =   6240
      Width           =   870
   End
   Begin VB.Label lblBalance 
      AutoSize        =   -1  'True
      Caption         =   "Balance"
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
      TabIndex        =   36
      Top             =   5880
      Width           =   780
   End
   Begin VB.Label lblTOChkO 
      AutoSize        =   -1  'True
      Caption         =   "Time Of CheckIn"
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
      TabIndex        =   35
      Top             =   6960
      Width           =   1650
   End
   Begin VB.Label lblDOChk 
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
      Left            =   120
      TabIndex        =   34
      Top             =   6600
      Width           =   1635
   End
   Begin VB.Label lblAdvance 
      AutoSize        =   -1  'True
      Caption         =   "Previous Payment"
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
      TabIndex        =   33
      Top             =   5520
      Width           =   1800
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
      Left            =   120
      TabIndex        =   32
      Top             =   5160
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
      Left            =   120
      TabIndex        =   31
      Top             =   4800
      Width           =   1770
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
      Left            =   5160
      TabIndex        =   30
      Top             =   960
      Width           =   525
   End
   Begin VB.Label lbldate 
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
      Left            =   360
      TabIndex        =   29
      Top             =   960
      Width           =   510
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      X1              =   0
      X2              =   7800
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label lbltitle 
      AutoSize        =   -1  'True
      Caption         =   "Check Out Information"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   1320
      TabIndex        =   28
      Top             =   240
      Width           =   3720
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
      Left            =   120
      TabIndex        =   27
      Top             =   4440
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
      Left            =   120
      TabIndex        =   26
      Top             =   2640
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
      Left            =   120
      TabIndex        =   25
      Top             =   4080
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
      Left            =   120
      TabIndex        =   24
      Top             =   3720
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
      Left            =   120
      TabIndex        =   23
      Top             =   2280
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
      Left            =   120
      TabIndex        =   22
      Top             =   1920
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
      Left            =   120
      TabIndex        =   21
      Top             =   3000
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
      Left            =   120
      TabIndex        =   20
      Top             =   3360
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
      Left            =   120
      TabIndex        =   19
      Top             =   1560
      Width           =   870
   End
End
Attribute VB_Name = "frmCheckOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Variable Declaration
Dim DOA As String
Dim result, i As Integer

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
    txtBalance.Text = ""
    txtTotalBill.Text = ""
End Sub

Private Sub cboGuID_Click()
 cmdDelete.Enabled = True
 With RS_GuestIn
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
   
   
    DOA = .Fields(0)
    result = Date - CDate(DOA)
    If result = 0 Then
     result = 1
    End If
    With RS_Payment
    .MoveFirst
    While Not .EOF
     If cboGuID.List(cboGuID.ListIndex) = RS_Payment.Fields(0) Then
    txtBalance.Text = Val(RS_Payment.Fields(9)) - Val(RS_Payment.Fields(8))
    txtTotalBill.Text = RS_Payment.Fields(9)
     txtadvance.Text = .Fields(8)
    End If
    .MoveNext
    Wend
    End With
   ' If txtRoomType.Text = "Single Room" Then
  '   txtTotalBill.Text = result * Rs_Rate.Fields(0)
 '    txtBalance.Text = txtTotalBill.Text - txtadvance.Text
 '   End If
 '   If txtRoomType.Text = "Double Room" Then
 '    txtTotalBill.Text = result * Rs_Rate.Fields(1)
 '    txtBalance.Text = txtTotalBill.Text - txtadvance.Text
 '   End If
 '   If txtRoomType.Text = "Suite Room" Then
 '    txtTotalBill.Text = result * Rs_Rate.Fields(2)
 '    txtBalance.Text = txtTotalBill.Text - txtadvance.Text
 '   End If
 '   If txtRoomType.Text = "Deluxe Suite" Then
 '    txtTotalBill.Text = result * Rs_Rate.Fields(3)
 '    txtBalance.Text = txtTotalBill.Text - txtadvance.Text
 '  End If
   Exit Sub
   Else
   .MoveNext
   End If
  Wend
 End With

End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
With RS_GuestOut
 .AddNew
 .Fields(0) = lbldate.Caption
 .Fields(1) = lbltime.Caption
 .Fields(2) = cboGuID.List(cboGuID.ListIndex)
 .Fields(3) = txtname.Text
 .Fields(4) = txtaddress.Text
 .Fields(5) = txtcity.Text
 .Fields(6) = txtState.Text
 .Fields(7) = txtCountry
 .Fields(8) = txtcompany.Text
 .Fields(9) = txtdesignation.Text
 .Fields(10) = txttelephone.Text
 .Fields(11) = txtRoomType.Text
 .Fields(12) = txtRoomNo.Text
 .Fields(13) = txtadvance.Text
 .Fields(14) = txtDOChkin.Text
 .Fields(15) = txtTOChkIn.Text
 .Fields(16) = txtBalance.Text
 .Fields(17) = txtTotalBill.Text
 .Update
 End With
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
 
 With RS_GuestIn
  .MoveFirst
smart:
 If cboGuID.List(cboGuID.ListIndex) = _
  .Fields(1) Then
  .Delete
  cboGuID.RemoveItem cboGuID.ListIndex
   
 Else
  .MoveNext
  GoTo smart
  End If
 MsgBox "Record Is Successfully Checked Out"
 cmdDelete.Enabled = False
 cboGuID.SetFocus
End With

With RS_Userlog
       .AddNew
       .Fields(0) = UserName
       .Fields(1) = "Check Out"
       .Fields(2) = Date
       .Fields(3) = Time
        .Fields(4) = "Successful"
       .Update
    End With
Blank
End Sub

Private Sub Form_Load()
Me.Top = 3000
Me.Left = 3000

lbldate.Caption = Date
lbltime.Caption = Time
cmdDelete.Enabled = False

Call Connect
 With RS_GuestIn
  While Not .EOF
   cboGuID.AddItem .Fields(1)
   .MoveNext
  Wend
 End With
End Sub
