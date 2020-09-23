VERSION 5.00
Begin VB.Form BarSettings 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Bar Settings"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Price Settings"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   7335
      Begin VB.OptionButton Option4 
         Caption         =   "Wine"
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
         Left            =   120
         TabIndex        =   19
         Top             =   1800
         Width           =   1095
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Juice"
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
         Left            =   120
         TabIndex        =   18
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Beer"
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
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Soft Drinks"
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
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Cancel"
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
         TabIndex        =   15
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Add New"
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
         TabIndex        =   14
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Save Price"
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
         TabIndex        =   13
         Top             =   480
         Width           =   1095
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
         Left            =   4680
         TabIndex        =   8
         Top             =   1680
         Width           =   855
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
         Left            =   4680
         TabIndex        =   7
         Top             =   1200
         Width           =   855
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
         Left            =   4680
         TabIndex        =   6
         Top             =   720
         Width           =   855
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
         Left            =   4680
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox Combo4 
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
         Left            =   1440
         TabIndex        =   4
         Top             =   1680
         Width           =   2175
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
         Left            =   1440
         TabIndex        =   3
         Top             =   1200
         Width           =   2175
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
         Left            =   1440
         TabIndex        =   2
         Top             =   720
         Width           =   2175
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
         Left            =   1440
         TabIndex        =   1
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
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
         Left            =   3720
         TabIndex        =   12
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
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
         Left            =   3720
         TabIndex        =   11
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label6 
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
         Height          =   255
         Left            =   3720
         TabIndex        =   10
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Label5"
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
         Left            =   3720
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "BarSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()
With RS_DrinkPrice
.MoveFirst
While Not .EOF
If Combo1.Text = .Fields(0) Then
Label5.Caption = .Fields(2)
End If
.MoveNext
Wend
End With
End Sub
Private Sub Combo2_Click()
With RS_DrinkPrice
.MoveFirst
While Not .EOF
If Combo2.Text = .Fields(0) Then
Label6.Caption = .Fields(2)
End If
.MoveNext
Wend
End With
End Sub
Private Sub Combo3_Click()
With RS_DrinkPrice
.MoveFirst
While Not .EOF
If Combo3.Text = .Fields(0) Then
Label7.Caption = .Fields(2)
End If
.MoveNext
Wend
End With
End Sub
Private Sub Combo4_Click()
With RS_DrinkPrice
.MoveFirst
While Not .EOF
If Combo4.Text = .Fields(0) Then
Label8.Caption = .Fields(2)
End If
.MoveNext
Wend
End With
End Sub

Private Sub Command1_Click()
If Option1.Enabled = True And Text1.Text <> "" Then
With RS_DrinkPrice
.MoveFirst
While Not .EOF
If Combo1.Text = .Fields(0) Then
MsgBox "Record Already Exits", vbOKOnly
Exit Sub
End If

.MoveNext
Wend
.AddNew
.Fields(0) = Combo1.Text
.Fields(1) = Option1.Caption
.Fields(2) = Text1.Text
.Update
MsgBox "Record Updated"

End With

ElseIf Option2.Enabled = True And Text2.Text <> "" Then
With RS_DrinkPrice
.MoveFirst
While Not .EOF
If Combo2.Text = .Fields(0) Then
MsgBox "Record Already Exits", vbOKOnly
Exit Sub
End If
.MoveNext
Wend
.AddNew
.Fields(0) = Combo2.Text
.Fields(1) = Option2.Caption
.Fields(2) = Text2.Text
.Update
MsgBox "Record Updated"

End With

ElseIf Option3.Enabled = True And Text3.Text <> "" Then
With RS_DrinkPrice
.MoveFirst
While Not .EOF
If Combo3.Text = .Fields(0) Then
MsgBox "Record Already Exits", vbOKOnly
Exit Sub
End If
MsgBox "Record Updated"
.MoveNext
Wend
.AddNew
.Fields(0) = Combo3.Text
.Fields(1) = Option3.Caption
.Fields(2) = Text3.Text
.Update
End With

ElseIf Option4.Enabled = True And Text4.Text <> "" Then
With RS_DrinkPrice
.MoveFirst
While Not .EOF
If Combo4.Text = .Fields(0) Then
MsgBox "Record Already Exits", vbOKOnly
Exit Sub
End If
.MoveNext
Wend
.AddNew
.Fields(0) = Combo4.Text
.Fields(1) = Option4.Caption
.Fields(2) = Text4.Text
.Update
MsgBox "Record Updated"

End With

Else: MsgBox "Record is Incomplete", vbCritical, "Error"
Exit Sub
End If
Blank
End Sub



Private Sub Command3_Click()
Blank
Me.Hide
End Sub

Private Sub Form_Load()
Top = 3000
Left = 3000
Label5.Caption = ""
Label6.Caption = ""
Label7.Caption = ""
Label8.Caption = ""


Connect
With RS_Drink
.MoveFirst
While Not .EOF
If .Fields(1) = "Juice" Then
Combo3.AddItem .Fields(0)
ElseIf .Fields(1) = "Beer" Then
Combo2.AddItem .Fields(0)
ElseIf .Fields(1) = "Soft Drinks" Then
Combo1.AddItem .Fields(0)
End If
.MoveNext
Wend
End With
End Sub

Sub Blank()
Combo1.ListIndex = -1
Combo2.ListIndex = -1
Combo3.ListIndex = -1
Combo4.ListIndex = -1
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Label5.Caption = ""
Label6.Caption = ""
Label7.Caption = ""
Label8.Caption = ""

End Sub
