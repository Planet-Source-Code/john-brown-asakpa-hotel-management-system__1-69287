VERSION 5.00
Begin VB.Form frmCompany 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Company Information"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
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
      Left            =   3360
      TabIndex        =   12
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
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
      Height          =   375
      Left            =   1920
      TabIndex        =   11
      Top             =   2640
      Width           =   1095
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
      Left            =   1800
      TabIndex        =   10
      Top             =   2160
      Width           =   3495
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
      Left            =   1800
      TabIndex        =   9
      Top             =   1800
      Width           =   3495
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
      Left            =   1800
      TabIndex        =   8
      Top             =   1440
      Width           =   3495
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
      Left            =   1800
      TabIndex        =   7
      Top             =   1080
      Width           =   3495
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
      Left            =   1800
      TabIndex        =   6
      Top             =   720
      Width           =   3495
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "HOTEL INFORMATION"
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
      Left            =   960
      TabIndex        =   5
      Top             =   120
      Width           =   2955
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Website"
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
      Top             =   2280
      Width           =   780
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Email Address"
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
      Top             =   1800
      Width           =   1320
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Phone Number"
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
      Top             =   1440
      Width           =   1395
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Address"
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
      Top             =   1080
      Width           =   765
   End
   Begin VB.Label Label1 
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
      TabIndex        =   0
      Top             =   720
      Width           =   540
   End
End
Attribute VB_Name = "frmCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
With RS_Company
.Update
 Text1.Text = .Fields(0)
 Text2.Text = .Fields(1)
 Text3.Text = .Fields(2)
 Text4.Text = .Fields(3)
Text5.Text = .Fields(4)
 .Fields(6) = Date
MsgBox "Record Successfully updated", vbOKOnly + vbInformation, "Update Records"
'End If
End With
End Sub

Private Sub Command2_Click()
Form_Activate
Me.Hide
End Sub

Private Sub Form_Activate()
Call Connect
End Sub

Private Sub Form_Load()
Me.Top = 3000
Me.Left = 3000

Connect
With RS_Company
Text1.Text = .Fields(0)
Text2.Text = .Fields(1)
Text3.Text = .Fields(2)
Text4.Text = .Fields(3)
Text5.Text = .Fields(4)
End With
End Sub
