VERSION 5.00
Begin VB.Form ResetPay 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Reset Payroll"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      Left            =   1320
      TabIndex        =   2
      Top             =   1200
      Width           =   855
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
      Left            =   480
      TabIndex        =   1
      Top             =   600
      Width           =   2895
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select Month to be Deleted"
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
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "ResetPay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Combo1.ListIndex = -1 Then
MsgBox "Pls Enter a Month you want to Deleted", vbCritical, "Error"
Exit Sub
End If
With RS_Payrolllog
.AddNew
.Fields(0) = RS_Payroll.Fields(0)
.Fields(1) = RS_Payroll.Fields(1)
.Fields(2) = RS_Payroll.Fields(2)
.Fields(3) = RS_Payroll.Fields(3)
.Fields(4) = RS_Payroll.Fields(4)
.Fields(5) = RS_Payroll.Fields(5)
.Fields(6) = RS_Payroll.Fields(6)
.Fields(7) = RS_Payroll.Fields(7)
.Fields(8) = RS_Payroll.Fields(8)
.Fields(9) = RS_Payroll.Fields(9)
.Fields(10) = RS_Payroll.Fields(10)
.Fields(11) = RS_Payroll.Fields(11)
.Fields(12) = RS_Payroll.Fields(12)
.Fields(13) = RS_Payroll.Fields(13)
.Fields(14) = Date
.Fields(15) = UserName
.Update
End With
With RS_Payroll
.MoveFirst
While Not .EOF
If Combo1.Text = .Fields(4) Then
.Delete adAffectCurrent
End If
.MoveNext
Wend
End With
MsgBox "Records Successfully Deleted", vbInformation
Me.Hide
With RS_Userlog
       .AddNew
       .Fields(0) = UserName
       .Fields(1) = "Reset Payroll Records"
       .Fields(2) = Date
       .Fields(3) = Time
       .Fields(4) = "Successful"
       .Update
    End With

End Sub

Private Sub Form_Load()
Connect
Combo1.AddItem "January"
Combo1.AddItem "February"
Combo1.AddItem "March"
Combo1.AddItem "April"
Combo1.AddItem "May"
Combo1.AddItem "June"
Combo1.AddItem "July"
Combo1.AddItem "August"
Combo1.AddItem "September"
Combo1.AddItem "October"
Combo1.AddItem "November"
Combo1.AddItem "December"
End Sub
