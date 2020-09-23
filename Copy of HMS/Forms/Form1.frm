VERSION 5.00
Begin VB.Form frmViewUsers 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Add User"
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
      Left            =   4440
      TabIndex        =   11
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
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
      Left            =   5760
      TabIndex        =   10
      Top             =   3240
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   1920
      TabIndex        =   1
      Top             =   840
      Width           =   4815
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
         Height          =   405
         Left            =   1680
         TabIndex        =   7
         Top             =   1440
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
         Height          =   405
         Left            =   1680
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
         Height          =   405
         Left            =   1680
         TabIndex        =   5
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Access Rights"
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
         Top             =   1440
         Width           =   1275
      End
      Begin VB.Label Label2 
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
         TabIndex        =   3
         Top             =   960
         Width           =   705
      End
      Begin VB.Label Label1 
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
         TabIndex        =   2
         Top             =   480
         Width           =   915
      End
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2790
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "ALL USER"
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
      Left            =   2400
      TabIndex        =   9
      Top             =   240
      Width           =   1545
   End
   Begin VB.Label lblcount 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   840
      TabIndex        =   8
      Top             =   3840
      Width           =   45
   End
End
Attribute VB_Name = "frmViewUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Me.Hide
End Sub

Private Sub Command2_Click()
Me.Hide
frmAddUser.Show 1
End Sub

Private Sub Form_Load()
Me.Top = 3000
Me.Left = 3000
Connect
With RS_Password
   .MoveFirst
    While Not .EOF
    List1.AddItem .Fields(0)
    .MoveNext
    Wend
        
       ' stsMain.Panels(1) = LVmain.ListItems.count & Space(1) & "Records in the Database"
        'lblcount = List1.count & "  " & "Records Currently in the Database"
    End With
End Sub

Private Sub List1_Click()
With RS_Password
.MoveFirst
While Not .EOF
If List1.List(List1.ListIndex) = .Fields(0) Then
Text1.Text = .Fields(0)
Text2.Text = .Fields(1)
Text3.Text = .Fields(3)
End If
.MoveNext
Wend

End With
End Sub
