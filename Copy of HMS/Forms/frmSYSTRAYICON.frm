VERSION 5.00
Begin VB.Form frmSYSTRAYICON 
   Caption         =   "Form1"
   ClientHeight    =   2055
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   2520
   Icon            =   "frmSYSTRAYICON.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2055
   ScaleWidth      =   2520
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picHook 
      Height          =   975
      Left            =   240
      ScaleHeight     =   915
      ScaleWidth      =   1395
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu checkin 
         Caption         =   "CheckIn"
      End
      Begin VB.Menu checkout 
         Caption         =   "Checkout"
      End
      Begin VB.Menu se 
         Caption         =   "-"
      End
      Begin VB.Menu view 
         Caption         =   "View Status of Hotel"
      End
      Begin VB.Menu logof 
         Caption         =   "LogOff"
      End
      Begin VB.Menu ser 
         Caption         =   "-"
      End
      Begin VB.Menu ab 
         Caption         =   "About"
      End
      Begin VB.Menu min 
         Caption         =   "Minimize Window"
      End
      Begin VB.Menu max 
         Caption         =   "Maximize Window"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmSYSTRAYICON"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ab_Click()
frmAbout.Show
End Sub

Private Sub checkin_Click()
frmcheckIn.Show
End Sub

Private Sub checkout_Click()
frmCheckOut.Show
End Sub

Private Sub exit_Click()
End
End Sub

Private Sub Form_Load()
    On Error Resume Next
    t.cbSize = Len(t)
    'Set the window's handle (this will be used to hook the specified window)
    t.hWnd = picHook.hWnd
    'Application-defined identifier of the taskbar icon
    t.uId = 1&
    'Set the flags
    t.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    'Set the callback message
    t.ucallbackMessage = WM_MOUSEMOVE
    'Set the picture (must be an icon!)
    t.hIcon = Me.Icon
    'Set the tooltiptext
    t.szTip = "Hotel Management System.  " & company & Chr$(0)
    'Create the icon
    Shell_NotifyIcon NIM_ADD, t
    Me.Hide
    App.TaskVisible = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
    t.cbSize = Len(t)
    t.hWnd = picHook.hWnd
    t.uId = 1&
    Shell_NotifyIcon NIM_DELETE, t
End Sub


Private Sub logof_Click()
frmMDI.Hide
fmLogin.Show
End Sub

Private Sub max_Click()
frmMDI.WindowState = 2

End Sub

Private Sub min_Click()
frmMDI.WindowState = 1
End Sub

Private Sub pichook_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    Static rec As Boolean, msg As Long
    msg = X / Screen.TwipsPerPixelX
    If rec = False Then
        rec = True
        Select Case msg
            Case WM_LBUTTONDBLCLK:
                
            Case WM_LBUTTONDOWN:
            Case WM_LBUTTONUP:
            Case WM_RBUTTONDBLCLK:
            Case WM_RBUTTONDOWN:
            Case WM_RBUTTONUP:
                Me.PopupMenu file
        End Select
        rec = False
    End If
End Sub

'***************************************
' System Tray Icon  [end]
'***************************************

Private Sub view_Click()
frmStatus.Show
End Sub
