VERSION 5.00
Begin VB.Form frmGuestPayment 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Guest Payment"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   5220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      TabIndex        =   17
      Top             =   1800
      Width           =   2415
   End
   Begin VB.TextBox Text6 
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
      Left            =   2160
      TabIndex        =   16
      Top             =   1440
      Width           =   2415
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
      Left            =   3600
      TabIndex        =   14
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton Command1 
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
      Left            =   2280
      TabIndex        =   13
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox Text5 
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
      Left            =   2160
      TabIndex        =   12
      Top             =   2880
      Width           =   2415
   End
   Begin VB.TextBox Text4 
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
      Left            =   2160
      TabIndex        =   11
      Top             =   2160
      Width           =   2415
   End
   Begin VB.TextBox Text3 
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
      Left            =   2160
      TabIndex        =   10
      Top             =   3240
      Width           =   2415
   End
   Begin VB.TextBox Text2 
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
      Left            =   2160
      TabIndex        =   9
      Top             =   2520
      Width           =   2415
   End
   Begin VB.TextBox Text1 
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
      Left            =   2160
      TabIndex        =   8
      Top             =   1080
      Width           =   2415
   End
   Begin VB.ComboBox Combo1 
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
      Left            =   2160
      TabIndex        =   7
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label Label9 
      Caption         =   "Payment Type"
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
      TabIndex        =   18
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Date of Checkin"
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
      Top             =   1440
      Width           =   1485
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Balance Payment"
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
      TabIndex        =   6
      Top             =   2880
      Width           =   1620
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Total Bill"
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
      TabIndex        =   5
      Top             =   2160
      Width           =   780
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Payment"
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
      TabIndex        =   4
      Top             =   3240
      Width           =   825
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Advance Payment"
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
      TabIndex        =   3
      Top             =   2520
      Width           =   1680
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Name"
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
      TabIndex        =   2
      Top             =   1080
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Guest ID"
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
      TabIndex        =   1
      Top             =   720
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "GUEST ACCOMODATION PAYMENT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   4260
   End
End
Attribute VB_Name = "frmGuestPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DOA As String
Dim result, i As Integer
Private Sub Combo1_Click()
With RS_Guest
.MoveFirst
While Not .EOF
If Combo1.List(Combo1.ListIndex) = .Fields(1) Then
Text1.Text = .Fields(2)
Text2.Text = .Fields(13)
Text6.Text = .Fields(0)
DOA = .Fields(0)
    result = Date - CDate(DOA)
    If result = 0 Then
     result = 1
     End If
     If .Fields(11).Value = "Single Room" Then
     Text4.Text = result * Rs_Rate.Fields(0)
     Text5.Text = Text4.Text - Text2.Text
    End If
    If .Fields(11).Value = "Double Room" Then
     Text4.Text = result * Rs_Rate.Fields(1)
     Text5.Text = Text4.Text - Text2.Text
    End If
    If .Fields(11).Value = "Suite Room" Then
      Text4.Text = result * Rs_Rate.Fields(2)
     Text5.Text = Text4.Text - Text2.Text
    End If
    If .Fields(11).Value = "Deluxe Suite" Then
      Text4.Text = result * Rs_Rate.Fields(3)
     Text5.Text = Text4.Text - Text2.Text
   End If
End If
.MoveNext
Wend
End With
End Sub

Private Sub Command1_Click()
With RS_Paymentlog
.MoveFirst
While Not .EOF
If Combo1.Text = .Fields(0) Then
.AddNew
.Fields(0) = Combo1.Text
.Fields(1) = Text1.Text
.Fields(2) = Text6.Text
.Fields(3) = Combo2.Text
.Fields(6) = Text2.Text
.Fields(7) = Text3.Text
.Fields(8) = Val(Text3.Text) + Val(Text2.Text)
.Fields(9) = Text4.Text
.Fields(10) = Date
.Fields(11) = UserName
.Update
End If
.MoveNext
Wend
End With


With RS_Payment
.MoveFirst
While Not .EOF
If Combo1.Text = .Fields(0) Then
.Update
.Fields(0) = Combo1.Text
.Fields(1) = Text1.Text
.Fields(2) = Text6.Text
.Fields(6) = Text2.Text
.Fields(7) = Text3.Text
.Fields(8) = Val(Text3.Text) + Val(Text2.Text)
.Fields(9) = Text4.Text
.Fields(11) = Date
.Fields(12) = UserName
End If
.MoveNext
Wend
End With

MsgBox "Payment Successfully Updated", vbOKOnly, "Payment"
'Command2_Click
End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Combo1.ListIndex = -1
Me.Hide
End Sub

Private Sub Form_Load()
Me.Top = 3000
Me.Left = 3000
Connect
With RS_Guest
.MoveFirst
While Not .EOF
Combo1.AddItem .Fields(1)
.MoveNext
Wend
End With
Combo2.AddItem "Accomodation"
Combo2.AddItem "Bar"
Combo2.AddItem "Laundry"
Combo2.AddItem "Restaurant"
End Sub
