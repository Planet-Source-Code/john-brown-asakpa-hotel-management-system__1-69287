VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "&H8000000F&"
   ClientHeight    =   3675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6945
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   6600
      Top             =   2880
   End
   Begin MSComctlLib.ProgressBar PB 
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3240
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   390
      Left            =   720
      TabIndex        =   6
      Top             =   2040
      Width           =   1425
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
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
      TabIndex        =   5
      Top             =   3000
      Width           =   60
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "System"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   585
      Left            =   4320
      TabIndex        =   3
      Top             =   1320
      Width           =   2010
   End
   Begin VB.Label lblProductName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Management"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   585
      Left            =   1320
      TabIndex        =   2
      Top             =   840
      Width           =   3660
   End
   Begin VB.Label lblCompanyProduct 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hotel"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   585
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   1650
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Developed By:- Brown Asakpa O. J."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   240
      Left            =   2880
      TabIndex        =   0
      Top             =   120
      Width           =   3825
   End
   Begin VB.Image Image1 
      Height          =   2895
      Left            =   0
      Picture         =   "frmSplash.frx":0442
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6975
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Dim warning1 As String
Dim warning2 As String
Dim warning3 As String
Dim Expday As Date
Dim Exp As String
'// This Code is for Licence Verification
Open App.Path & "\reg.dll" For Input As #1
Line Input #1, a
Line Input #1, B
Line Input #1, C
Line Input #1, D

warning1 = a
warning2 = B
warning3 = C
Exp = D
Close #1
Expday = Date

'-------above reads the Reg.DLL and loads the labels with its info
If warning1 = Date Then
MsgBox "Please Note that you have till " & Exp & " to renew your Software Licence", vbInformation
End If

If warning2 = Date Then
MsgBox "Please Note that you have till " & Exp & " to renew your Software Licence", vbInformation
End If

If warning3 = Date Then
MsgBox "Please Note that you have till " & Exp & " to renew your Software Licence", vbInformation
End If

If Exp = Expday Then
MsgBox "Sorry Licence Period Expired,You can no longer use this Software. Contact Administrator", vbInformation
End
End If

Connect

With RS_Company
Label3.Caption = .Fields(0)
End With
PB.min = 1
PB.max = 200
PB.Value = 1
End Sub

Private Sub Form_Load()
fmLogin.Enabled = False
End Sub

Private Sub Timer1_Timer()
PB.Value = PB.Value + 3
If PB.Value = 40 Then
Label2.Caption = "Loading Personal Settings"
ElseIf PB.Value = 70 Then
Label2.Caption = "Applying Your Personal Settings"
ElseIf PB.Value = 100 Then
Label2.Caption = "Loading Database"
ElseIf PB.Value = 160 Then
Label2.Caption = "Please wait ...."
ElseIf PB.Value = 180 Then
Label2.Caption = "..........."
End If

If PB.Value > 198 Then
    Timer1.Enabled = False
    fmLogin.Enabled = True
    fmLogin.Show
    Unload Me
End If
End Sub
