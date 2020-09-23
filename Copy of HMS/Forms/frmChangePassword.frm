VERSION 5.00
Begin VB.Form frmChangePassword 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Change Password"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5490
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   5490
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00C0C0C0&
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
      Left            =   2880
      TabIndex        =   8
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdchange 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Change Password"
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
      Left            =   1440
      TabIndex        =   7
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txtoldpassword 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2160
      TabIndex        =   0
      Top             =   720
      Width           =   3015
   End
   Begin VB.TextBox txtnewpassword 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1080
      Width           =   3015
   End
   Begin VB.TextBox txtconfirmpassword 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1500
      Width           =   3015
   End
   Begin VB.Label lbloldpassword 
      AutoSize        =   -1  'True
      Caption         =   "Old Password"
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
      Left            =   240
      TabIndex        =   6
      Top             =   720
      Width           =   1380
   End
   Begin VB.Label lblnewpassword 
      AutoSize        =   -1  'True
      Caption         =   "New Password"
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
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   1485
   End
   Begin VB.Label lblconfirmpassword 
      AutoSize        =   -1  'True
      Caption         =   "Confirm Password"
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
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
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
      Left            =   1080
      TabIndex        =   3
      Top             =   240
      Width           =   870
   End
End
Attribute VB_Name = "frmChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Dim str, UserName, StrSql As String


Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdChange_Click()

 If txtoldpassword.Text = "" Then
    MsgBox "Enter Complete Information", vbOKOnly + vbInformation, "Information"
    txtoldpassword.SetFocus
    Exit Sub
 End If
 
 If txtnewpassword.Text = "" Then
    MsgBox "Enter Complete Information", vbOKOnly + vbInformation, "Information"
    txtnewpassword.SetFocus
    Exit Sub
 End If
 
 If txtconfirmpassword.Text = "" Then
     MsgBox "Enter Complete Information", vbOKOnly + vbInformation, "Information"
     txtconfirmpassword.SetFocus
    Exit Sub
 End If
 
 With RS_Password
 .MoveLast
 .MoveFirst
 While Not .EOF
  If .Fields(0) = UserName Then
   If .Fields(1) <> txtoldpassword.Text Then
    MsgBox "Old Password Doesn't Match", vbOKOnly + vbCritical, "Error"
    txtoldpassword.SetFocus
    Exit Sub
   ElseIf txtnewpassword.Text <> txtconfirmpassword.Text Then
    MsgBox "New Password Doesn't Match", vbOKOnly + vbCritical, "Error"
    txtnewpassword.SetFocus
    Exit Sub
   Else
    StrSql = "UPDATE Password_Table SET Pass = '" & txtconfirmpassword.Text & "' " _
             & "WHERE UserName = '" & UserName & "';"
    cnn.Execute StrSql
    .Update
    MsgBox "Done"
    Exit Sub
   End If
  End If
  .MoveNext
 Wend
 MsgBox "Password Successfully Changed", vbInformation
 txtoldpassword.Text = ""
 txtnewpassword.Text = ""
 txtconfirmpassword.Text = ""
End With
   
          With RS_Userlog
       .AddNew
       .Fields(0) = UserName
       .Fields(1) = "User Changed Password"
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
str = frmMDI.StatusBar.Panels(1).Text

UserName = Mid(str, 14, Len(str))

 
 lbltitle.Caption = "User Name ----> " & UserName
 



End Sub
