VERSION 5.00
Begin VB.Form frmAddUser 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add New User"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5745
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      Left            =   1920
      TabIndex        =   12
      Top             =   2040
      Width           =   2655
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
      Left            =   3000
      TabIndex        =   10
      Top             =   2760
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
      Left            =   1440
      TabIndex        =   9
      Top             =   2760
      Width           =   975
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
      Left            =   1920
      TabIndex        =   8
      Top             =   1680
      Width           =   2655
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
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   1320
      Width           =   2655
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
      Left            =   1920
      TabIndex        =   6
      Top             =   960
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
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Top             =   600
      Width           =   2655
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Access Type"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   11
      Top             =   2160
      Width           =   1140
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Retype Password"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   1650
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   930
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "User ID"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Full Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ADD NEW USER"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   2520
   End
End
Attribute VB_Name = "frmAddUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Then
MsgBox "Please enter the required records", vbCritical
Exit Sub
End If
If Text3.Text <> Text4.Text Then
MsgBox "Password Does not Match"
ElseIf Text3.Text = Text4.Text Then
With RS_Password
.AddNew
.Fields(0) = Text1.Text
.Fields(1) = Text2.Text
.Fields(2) = Crypt(Text3.Text)
.Fields(3) = Combo1.Text
.Update
MsgBox "User Successfully Added"
End With
End If
Blank

With RS_Userlog
       .AddNew
       .Fields(0) = UserName
       .Fields(1) = "Added New User"
       .Fields(2) = Date
       .Fields(3) = Time
        .Fields(4) = "Successful"
       .Update
    End With


End Sub

Private Sub Command2_Click()
Blank
Me.Hide
End Sub

Private Sub Form_Load()
Me.Top = 3000
Me.Left = 3000
Call Connect
With RS_Password
.MoveFirst
While Not .EOF
Combo1.AddItem .Fields(3)
.MoveNext
Wend
End With
End Sub
Sub Blank()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Combo1.ListIndex = -1
End Sub
