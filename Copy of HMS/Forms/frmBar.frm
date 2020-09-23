VERSION 5.00
Begin VB.Form frmBar 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Bar"
   ClientHeight    =   7185
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   7050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option2 
      Caption         =   "Guest"
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
      TabIndex        =   27
      Top             =   1320
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Customer"
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
      TabIndex        =   26
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   24
      Top             =   6600
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "Hotel Guest"
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
      Height          =   2895
      Left            =   240
      TabIndex        =   11
      Top             =   4200
      Width           =   5655
      Begin VB.CommandButton Command6 
         Caption         =   "Payment"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4320
         TabIndex        =   33
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "New Order"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4320
         TabIndex        =   29
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Take Order"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4320
         TabIndex        =   28
         Top             =   1440
         Width           =   1215
      End
      Begin VB.ComboBox Combo3 
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
         Left            =   1680
         TabIndex        =   22
         Top             =   240
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
         Left            =   1560
         TabIndex        =   19
         Top             =   2280
         Width           =   2415
      End
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
         Left            =   1680
         TabIndex        =   13
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Label19"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4440
         TabIndex        =   32
         Top             =   1080
         Width           =   810
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Label15"
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
         Left            =   1680
         TabIndex        =   23
         Top             =   720
         Width           =   750
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Guest Name"
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
         Left            =   240
         TabIndex        =   21
         Top             =   720
         Width           =   1155
      End
      Begin VB.Label Label13 
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
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   825
      End
      Begin VB.Label Label12 
         Caption         =   "Quantity"
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
         Left            =   240
         TabIndex        =   18
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Label11"
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
         Left            =   1680
         TabIndex        =   17
         Top             =   1920
         Width           =   750
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Rate"
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
         Left            =   240
         TabIndex        =   16
         Top             =   1920
         Width           =   435
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Label9"
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
         Left            =   1680
         TabIndex        =   15
         Top             =   1560
         Width           =   630
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Drink Type"
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
         Left            =   240
         TabIndex        =   14
         Top             =   1560
         Width           =   990
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Drinks"
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
         Left            =   240
         TabIndex        =   12
         Top             =   1080
         Width           =   585
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Customer"
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
      Height          =   2175
      Left            =   240
      TabIndex        =   0
      Top             =   1920
      Width           =   5655
      Begin VB.CommandButton Command2 
         Caption         =   "New Order"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4320
         TabIndex        =   10
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Take Order"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4320
         TabIndex        =   9
         Top             =   1200
         Width           =   1215
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
         Left            =   1680
         TabIndex        =   6
         Top             =   1560
         Width           =   2415
      End
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
         Left            =   1680
         TabIndex        =   2
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Label18"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4320
         TabIndex        =   31
         Top             =   840
         Width           =   810
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Label6"
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
         Left            =   1680
         TabIndex        =   8
         Top             =   840
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Drink Type"
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
         TabIndex        =   7
         Top             =   840
         Width           =   990
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Label4"
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
         Left            =   1680
         TabIndex        =   5
         Top             =   1200
         Width           =   630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Quantity"
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
         Top             =   1560
         Width           =   810
      End
      Begin VB.Label Label2 
         Caption         =   "Rate"
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
         TabIndex        =   3
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Drinks"
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
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Image Image1 
      Height          =   1935
      Left            =   4560
      Picture         =   "frmBar.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2535
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "Process Bar Order"
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
      Left            =   1440
      TabIndex        =   30
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Select Order Options"
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
      Left            =   240
      TabIndex        =   25
      Top             =   600
      Width           =   2250
   End
   Begin VB.Menu fi 
      Caption         =   "File"
      Begin VB.Menu set 
         Caption         =   "Settings"
      End
      Begin VB.Menu exi 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
With RS_DrinkPrice
.MoveFirst
While Not .EOF
If Combo1.List(Combo1.ListIndex) = .Fields(0) Then
Label4.Caption = .Fields(2)
Label6.Caption = .Fields(1)
End If
.MoveNext
Wend
End With
End Sub

Private Sub Combo2_Click()
With RS_DrinkPrice
.MoveFirst
While Not .EOF
If Combo1.List(Combo1.ListIndex) = .Fields(0) Then
Label9.Caption = .Fields(2)
Label11.Caption = .Fields(1)
End If
.MoveNext
Wend
End With
End Sub

Private Sub Combo3_Click()
With RS_Guest
.MoveFirst
While Not .EOF
If Combo3.Text = .Fields(1) Then
Label15.Caption = .Fields(2)
End If
.MoveNext
Wend
End With

End Sub

Private Sub Command1_Click()
Label18.Caption = Label4.Caption * Val(Text1.Text)
End Sub

Private Sub Command2_Click()
Combo1.ListIndex = -1
Text1.Text = ""
Label4.Caption = ""
Label6.Caption = ""
Label18.Caption = ""
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
Label19.Caption = Label11.Caption * Val(Text2.Text)

End Sub

Private Sub Command5_Click()
Combo2.ListIndex = -1
Combo3.ListIndex = -1
Text2.Text = ""
Label9.Caption = ""
Label11.Caption = ""
Label15.Caption = ""
Label19.Caption = ""
End Sub

Private Sub exi_Click()
Me.Hide
End Sub

Private Sub Form_Load()
Top = 3000
Left = 3000
Connect

With RS_Drink
.MoveFirst
While Not .EOF
Combo1.AddItem .Fields(0)
Combo2.AddItem .Fields(0)

.MoveNext
Wend
End With
Label4.Caption = ""
Label6.Caption = ""
Label9.Caption = ""
Label11.Caption = ""
Label15.Caption = ""
Label19.Caption = ""
Label18.Caption = ""
With RS_Guest
.MoveFirst
While Not .EOF
Combo3.AddItem .Fields(1)
.MoveNext
Wend
End With
Frame1.Enabled = False
Frame2.Enabled = False
End Sub

Private Sub Option1_Click()
Frame1.Enabled = True
Frame2.Enabled = False
End Sub

Private Sub Option2_Click()
Frame2.Enabled = True
Frame1.Enabled = False
End Sub

Private Sub set_Click()
BarSettings.Show 1
End Sub
