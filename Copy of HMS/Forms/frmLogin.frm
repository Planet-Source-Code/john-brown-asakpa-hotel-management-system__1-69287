VERSION 5.00
Begin VB.Form fmLogin 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Log On Please"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6480
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraControl 
      BackColor       =   &H00FF8080&
      Height          =   1695
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   4455
      Begin VB.TextBox txtpassword 
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
         IMEMode         =   3  'DISABLE
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   960
         Width           =   2295
      End
      Begin VB.ComboBox cboUserName 
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
         Left            =   1800
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label lblpassword 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   240
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   1020
      End
      Begin VB.Label lblusername 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Username:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   240
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1050
      End
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "OK"
      DragIcon        =   "frmLogin.frx":0000
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
End
Attribute VB_Name = "fmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Variable Declaration

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()

If cboUserName.ListIndex = -1 Then
 MsgBox "Select User Name", vbOKOnly + vbInformation, "Error"
 Exit Sub
End If
UserName = cboUserName.List(cboUserName.ListIndex)
With RS_Password
  '.MoveLast
  .MoveFirst
  While Not .EOF
    If UserName = .Fields(1).Value And txtpassword.Text = Decrypt(.Fields(2).Value) Then
      Rights = .Fields(3).Value
      MsgBox "Access Granted", vbOKOnly + vbInformation, "Correct"
    ElseIf UserName = .Fields(1).Value And txtpassword.Text <> Decrypt(.Fields(2).Value) Then
      MsgBox "Incorrect Password", vbOKOnly + vbInformation, "Correct"
      Exit Sub
       frmMDI.Show
       End If
     
    '  Exit Sub
   
    
     .MoveNext
    
  Wend

End With
 
 If Rights = "Administrator" Then
        frmMDI.StatusBar.Panels(1) = "User Name :- " & UserName
        
      ElseIf Rights = "Front Office" Then
        frmMDI.StatusBar.Panels(1) = "User Name :- " & UserName
        frmMDI.mnuGuests.Enabled = True
        frmMDI.mnuEmp.Enabled = False
        frmMDI.mnuSrchEmp.Enabled = False
        frmMDI.mnuReportsEmployees.Enabled = False
        frmMDI.mnuUtiChangeCharges.Enabled = False
        frmMDI.admin.Enabled = False
        
      ElseIf Rights = "Personnel" Then
        frmMDI.StatusBar.Panels(1) = "User Name :- " & UserName
        frmMDI.mnuEmp.Enabled = True
        frmMDI.mnuGuests.Enabled = False
        frmMDI.mnuSrchGuest.Enabled = False
        frmMDI.mnuReportsGuests.Enabled = False
        frmMDI.mnuUtiChangeCharges.Enabled = False
        frmMDI.admin.Enabled = False
        
       End If

With RS_Userlog
       .AddNew
       .Fields(0) = UserName
       .Fields(1) = "Log In"
       .Fields(2) = Date
       .Fields(3) = Time
        .Fields(4) = "Successful"
       .Update
    End With
      Load frmSYSTRAYICON
     Me.Hide
frmWelcome.Show 1
End Sub

Private Sub Form_Load()

Call Connect
With RS_Password
 .MoveLast
 .MoveFirst
  While Not .EOF
   cboUserName.AddItem .Fields(1).Value
   .MoveNext
  Wend
End With
End Sub


