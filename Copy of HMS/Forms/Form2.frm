VERSION 5.00
Begin VB.Form rptEmpReport 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SELECT EMPLOYEE ID"
   ClientHeight    =   1860
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5205
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      Left            =   1800
      TabIndex        =   2
      Top             =   360
      Width           =   2535
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
      Left            =   1920
      TabIndex        =   0
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Employee Name"
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
      Top             =   360
      Width           =   1515
   End
End
Attribute VB_Name = "rptEmpReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim db As ADODB.Connection
Dim status As Boolean
Dim str As String


Private Sub Command1_Click()
'If (Combo1.Text <> "All" And Combo1.Text <> "Book ID") Then
 'MsgBox "Please select proper Book specifications.", vbCritical, "Invalid Data"
'Exit Sub
'End If

'If (Combo1.Text = "All") Then
'str = "Select * from PresentEmp_Table"
If (Combo1.Text <> "") Then
        'If (Text1.Text <> "") Then
        '    If IsNumeric(Text1.Text) Then
            str = "Select * from PresentEmp_Table where Empid=" & Combo1.Text
            'Else
            'MsgBox ("Please enter Book ID Numeric value."), vbExclamation, "Invalid value"
            'Exit Sub
            'End If
        Else
        MsgBox ("Please enter Book ID."), vbExclamation, "Invalid value"
        Exit Sub
        'End If
End If
again:
If (status = False) Then
rs.Open str, db, adOpenDynamic, adLockPessimistic
status = True
Else
rs.Close
status = False
GoTo again
End If
Set DataReport1.DataSource = rs
DataReport1.Show vbModal
End Sub

Private Sub Form_Load()
Set db = New ADODB.Connection
Set rs = New ADODB.Recordset
db.CursorLocation = adUseClient
db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " _
         & App.Path & ".\Database\HMS.mdb;Persist Security Info=False"
         status = False

Connect
With Rs_Details
.MoveFirst
While Not .EOF
Combo1.AddItem .Fields(0)
.MoveNext
Wend
End With

End Sub
