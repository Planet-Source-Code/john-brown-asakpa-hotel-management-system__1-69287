VERSION 5.00
Begin VB.Form frmChangeRate 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Change Rate"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8205
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   8205
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   7695
      Begin VB.OptionButton opnDeluxeSuite 
         Caption         =   "Deluxe Room"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   11
         Top             =   1680
         Width           =   1575
      End
      Begin VB.OptionButton opnSuiteRoom 
         Caption         =   "Suite Room"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   10
         Top             =   1320
         Width           =   1455
      End
      Begin VB.OptionButton opnDoubleRoom 
         Caption         =   "Double Room"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton cmdcancel 
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
         Left            =   6000
         TabIndex        =   8
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton cmdChange 
         Caption         =   "Change"
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
         Left            =   6000
         TabIndex        =   7
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtChangeRate 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton opnSingleRoom 
         Caption         =   "Single Room"
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
         Left            =   480
         TabIndex        =   2
         Top             =   480
         Width           =   1575
      End
      Begin VB.Line Line6 
         X1              =   5640
         X2              =   5640
         Y1              =   120
         Y2              =   4080
      End
      Begin VB.Label lblPresentRate 
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   4200
         TabIndex        =   5
         Top             =   720
         Width           =   135
      End
      Begin VB.Label lblChangeRate 
         AutoSize        =   -1  'True
         Caption         =   "Change Rate"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   2520
         TabIndex        =   4
         Top             =   1320
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Present Rate"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   2520
         TabIndex        =   3
         Top             =   720
         Width           =   1275
      End
      Begin VB.Line Line5 
         X1              =   2280
         X2              =   2280
         Y1              =   120
         Y2              =   4080
      End
   End
   Begin VB.Label lbltitle 
      AutoSize        =   -1  'True
      Caption         =   "Change Hotel Rates"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   2040
      TabIndex        =   0
      Top             =   360
      Width           =   2610
   End
End
Attribute VB_Name = "frmChangeRate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim StrSql As String


Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdChange_Click()
 With Rs_Rate
  .MoveFirst
   If opnSingleRoom.Value = True Then
   StrSql = "UPDATE Rate_Table SET SingleRoom = '" _
            & txtChangeRate.Text & "'"
   cnn.Execute StrSql
   .Update
   lblPresentRate.Caption = txtChangeRate.Text
  End If
  
  If opnDoubleRoom.Value = True Then
   StrSql = "UPDATE Rate_Table SET DoubleRoom = '" _
            & txtChangeRate.Text & "'"
   cnn.Execute StrSql
   .Update
   lblPresentRate.Caption = txtChangeRate.Text
  End If
  
  If opnSuiteRoom.Value = True Then
   StrSql = "UPDATE Rate_Table SET SuiteRoom = '" _
            & txtChangeRate.Text & "'"
   cnn.Execute StrSql
   .Update
   lblPresentRate.Caption = txtChangeRate.Text
  End If
  
  If opnDeluxeSuite.Value = True Then
   StrSql = "UPDATE Rate_Table SET DeluxeSuite = '" _
            & txtChangeRate.Text & "'"
   cnn.Execute StrSql
   .Update
   lblPresentRate.Caption = txtChangeRate.Text
  End If
 End With
 MsgBox "Rate is changed Successfully"
 
 With RS_Userlog
       .AddNew
       .Fields(0) = UserName
       .Fields(1) = "Change Hotel Rate"
       .Fields(2) = Date
       .Fields(3) = Time
       .Fields(4) = "Successful"
       .Update
    End With

End Sub

Private Sub Form_Load()
Me.Top = 3000
Me.Left = 3000

Call Connect

lblPresentRate.Visible = False
cmdChange.Enabled = False
End Sub

Private Sub opnDeluxeSuite_Click()
If opnDeluxeSuite.Value = True Then
 Rs_Rate.Requery
 txtChangeRate.Text = ""
 lblPresentRate.Visible = True
 lblPresentRate.Caption = Rs_Rate.Fields(3)
 txtChangeRate.Locked = False
End If
 
End Sub

Private Sub opnDoubleRoom_Click()
If opnDoubleRoom.Value = True Then
 Rs_Rate.Requery
 txtChangeRate.Text = ""
 lblPresentRate.Visible = True
 lblPresentRate.Caption = Rs_Rate.Fields(1)
 txtChangeRate.Locked = False
End If
End Sub

Private Sub opnSingleRoom_Click()
 If opnSingleRoom.Value = True Then
  Rs_Rate.Requery
  txtChangeRate.Text = ""
  lblPresentRate.Visible = True
  lblPresentRate.Caption = Rs_Rate.Fields(0)
  txtChangeRate.Locked = False
 End If
End Sub

Private Sub opnSuiteRoom_Click()
If opnSuiteRoom.Value = True Then
 Rs_Rate.Requery
 txtChangeRate.Text = ""
 lblPresentRate.Visible = True
 lblPresentRate.Caption = Rs_Rate.Fields(2)
 txtChangeRate.Locked = False
End If
End Sub

Private Sub txtChangeRate_Change()
cmdChange.Enabled = True
End Sub

Private Sub txtChangeRate_KeyPress(KeyAscii As Integer)
 If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or _
     KeyAscii = vbKeyBack Then
  Exit Sub
 Else
  KeyAscii = 0
 End If
End Sub
