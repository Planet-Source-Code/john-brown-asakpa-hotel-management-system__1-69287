VERSION 5.00
Begin VB.Form frmCharges 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Charges Of Rooms"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7170
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   7170
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraList 
      Caption         =   "List Of Rates"
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
      Height          =   2655
      Left            =   480
      TabIndex        =   1
      Top             =   1560
      Width           =   5895
      Begin VB.Label lblDoubleRoom 
         AutoSize        =   -1  'True
         Caption         =   "2. Double Room"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   480
         TabIndex        =   9
         Top             =   960
         Width           =   1950
      End
      Begin VB.Label lblSuiteRoom 
         AutoSize        =   -1  'True
         Caption         =   "3. Suite Room"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   480
         TabIndex        =   8
         Top             =   1440
         Width           =   1740
      End
      Begin VB.Label lblDeluxeSuite 
         AutoSize        =   -1  'True
         Caption         =   "4. Deluxe Suite"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   480
         TabIndex        =   7
         Top             =   1920
         Width           =   1860
      End
      Begin VB.Label lblsingleroom 
         AutoSize        =   -1  'True
         Caption         =   "1. Single Room"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   480
         TabIndex        =   6
         Top             =   480
         Width           =   1845
      End
      Begin VB.Label lblRtSingleRoom 
         AutoSize        =   -1  'True
         Caption         =   "Label"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   3600
         TabIndex        =   5
         Top             =   540
         Width           =   660
      End
      Begin VB.Label lblrtDoubleRoom 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   3600
         TabIndex        =   4
         Top             =   960
         Width           =   825
      End
      Begin VB.Label lblRtSuiteRoom 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   3600
         TabIndex        =   3
         Top             =   1440
         Width           =   825
      End
      Begin VB.Label lblRtDeluxeSuite 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   3600
         TabIndex        =   2
         Top             =   1920
         Width           =   825
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   0
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "List of Rates of Rooms"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   1200
      TabIndex        =   11
      Top             =   240
      Width           =   3630
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      X1              =   0
      X2              =   7080
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label lblinformation2 
      AutoSize        =   -1  'True
      Caption         =   "All Charges are for Min. 24 Hours Stay"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   600
      TabIndex        =   10
      Top             =   1200
      Width           =   4695
   End
End
Attribute VB_Name = "frmCharges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Top = 3000
Me.Left = 3000


Call Connect

With Rs_Rate
'.MoveFirst
lblRtSingleRoom.Caption = .Fields(0)
lblrtDoubleRoom.Caption = .Fields(1)
lblRtSuiteRoom.Caption = .Fields(2)
lblRtDeluxeSuite.Caption = .Fields(3)
End With

End Sub

