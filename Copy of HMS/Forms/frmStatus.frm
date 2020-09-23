VERSION 5.00
Begin VB.Form frmStatus 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Status Of Hotel"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11310
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   11310
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdcloase 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   0
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label lbltotalrooms 
      AutoSize        =   -1  'True
      Caption         =   "Total Rooms"
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
      Left            =   240
      TabIndex        =   31
      Top             =   1680
      Width           =   1650
   End
   Begin VB.Label lblstatustitle 
      AutoSize        =   -1  'True
      Caption         =   "Status Of The Hotel"
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
      Left            =   3480
      TabIndex        =   30
      Top             =   360
      Width           =   3195
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      X1              =   0
      X2              =   11160
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label lbltotalroomvalue 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
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
      Index           =   0
      Left            =   2520
      TabIndex        =   29
      Top             =   1680
      Width           =   870
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      X1              =   3480
      X2              =   3480
      Y1              =   1080
      Y2              =   3960
   End
   Begin VB.Label llblSingleRoom 
      AutoSize        =   -1  'True
      Caption         =   "Single Rooms"
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
      Index           =   0
      Left            =   240
      TabIndex        =   28
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label lblSingleRoomValue 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
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
      Index           =   0
      Left            =   2520
      TabIndex        =   27
      Top             =   2280
      Width           =   720
   End
   Begin VB.Label lblDoubleRoom 
      AutoSize        =   -1  'True
      Caption         =   "Double Rooms"
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
      Index           =   0
      Left            =   240
      TabIndex        =   26
      Top             =   2640
      Width           =   1545
   End
   Begin VB.Label lblDoubleRoomValue 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
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
      Index           =   0
      Left            =   2520
      TabIndex        =   25
      Top             =   2640
      Width           =   720
   End
   Begin VB.Label lblSuiteRoom 
      AutoSize        =   -1  'True
      Caption         =   "Suite Rooms"
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
      Index           =   0
      Left            =   240
      TabIndex        =   24
      Top             =   3000
      Width           =   1350
   End
   Begin VB.Label lblSuiteRoomValue 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
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
      Index           =   0
      Left            =   2520
      TabIndex        =   23
      Top             =   3000
      Width           =   720
   End
   Begin VB.Label lblDeluxeRoom 
      AutoSize        =   -1  'True
      Caption         =   "Deluxe Suite"
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
      Index           =   0
      Left            =   240
      TabIndex        =   22
      Top             =   3360
      Width           =   1365
   End
   Begin VB.Label lblDeluxeSuiteValue 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
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
      Index           =   0
      Left            =   2520
      TabIndex        =   21
      Top             =   3360
      Width           =   720
   End
   Begin VB.Label lblAvailableRooms 
      AutoSize        =   -1  'True
      Caption         =   "Available Rooms"
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
      Index           =   1
      Left            =   3720
      TabIndex        =   20
      Top             =   1680
      Width           =   2145
   End
   Begin VB.Label lblDeluxeRoom 
      AutoSize        =   -1  'True
      Caption         =   "Deluxe Suite"
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
      Index           =   1
      Left            =   3720
      TabIndex        =   19
      Top             =   3360
      Width           =   1365
   End
   Begin VB.Label lblSuiteRoom 
      AutoSize        =   -1  'True
      Caption         =   "Suite Rooms"
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
      Index           =   1
      Left            =   3720
      TabIndex        =   18
      Top             =   3000
      Width           =   1350
   End
   Begin VB.Label lblDoubleRoom 
      AutoSize        =   -1  'True
      Caption         =   "Double Rooms"
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
      Index           =   1
      Left            =   3720
      TabIndex        =   17
      Top             =   2640
      Width           =   1545
   End
   Begin VB.Label llblSingleRoom 
      AutoSize        =   -1  'True
      Caption         =   "Single Rooms"
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
      Index           =   1
      Left            =   3720
      TabIndex        =   16
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label lblDeluxeSuiteValue 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
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
      Index           =   1
      Left            =   6240
      TabIndex        =   15
      Top             =   3360
      Width           =   720
   End
   Begin VB.Label lblSuiteRoomValue 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
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
      Index           =   1
      Left            =   6240
      TabIndex        =   14
      Top             =   3000
      Width           =   720
   End
   Begin VB.Label lblDoubleRoomValue 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
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
      Index           =   1
      Left            =   6240
      TabIndex        =   13
      Top             =   2640
      Width           =   720
   End
   Begin VB.Label lblSingleRoomValue 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
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
      Index           =   1
      Left            =   6240
      TabIndex        =   12
      Top             =   2280
      Width           =   720
   End
   Begin VB.Label lbltotalroomvalue1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
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
      Left            =   6240
      TabIndex        =   11
      Top             =   1680
      Width           =   870
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      X1              =   0
      X2              =   11280
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000FF&
      X1              =   7320
      X2              =   7320
      Y1              =   1080
      Y2              =   3960
   End
   Begin VB.Label lblSingleRoomValue 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
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
      Index           =   2
      Left            =   10200
      TabIndex        =   10
      Top             =   2280
      Width           =   720
   End
   Begin VB.Label lblDoubleRoomValue 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
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
      Index           =   2
      Left            =   10200
      TabIndex        =   9
      Top             =   2640
      Width           =   720
   End
   Begin VB.Label lblSuiteRoomValue 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
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
      Index           =   2
      Left            =   10200
      TabIndex        =   8
      Top             =   3000
      Width           =   720
   End
   Begin VB.Label lblDeluxeSuiteValue 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
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
      Index           =   2
      Left            =   10200
      TabIndex        =   7
      Top             =   3360
      Width           =   720
   End
   Begin VB.Label llblSingleRoom 
      AutoSize        =   -1  'True
      Caption         =   "Single Rooms"
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
      Index           =   2
      Left            =   7440
      TabIndex        =   6
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label lblDoubleRoom 
      AutoSize        =   -1  'True
      Caption         =   "Double Rooms"
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
      Index           =   2
      Left            =   7440
      TabIndex        =   5
      Top             =   2640
      Width           =   1545
   End
   Begin VB.Label lblSuiteRoom 
      AutoSize        =   -1  'True
      Caption         =   "Suite Rooms"
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
      Index           =   2
      Left            =   7440
      TabIndex        =   4
      Top             =   3000
      Width           =   1350
   End
   Begin VB.Label lblDeluxeRoom 
      AutoSize        =   -1  'True
      Caption         =   "Deluxe Suite"
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
      Index           =   2
      Left            =   7440
      TabIndex        =   3
      Top             =   3360
      Width           =   1365
   End
   Begin VB.Label lblOccupiedRooms 
      AutoSize        =   -1  'True
      Caption         =   "Occupied Rooms"
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
      Left            =   7440
      TabIndex        =   2
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label lbloccupiedRoomvalue1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
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
      Left            =   10200
      TabIndex        =   1
      Top             =   1680
      Width           =   870
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SR, DR, SUR, DSUR, i, J, SR1, DR1, SUR1, DSUR1 As Integer

Private Sub cmdcloase_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Top = 3000
Me.Left = 3000

Call Connect

lbltotalroomvalue(0).Caption = "100"
lblSingleRoomValue(0).Caption = "40"
lblDoubleRoomValue(0).Caption = "40"
lblSuiteRoomValue(0).Caption = "10"
lblDeluxeSuiteValue(0).Caption = "10"

With RS_SingleRoom
 .MoveFirst
 SR = .RecordCount
 lblSingleRoomValue(1).Caption = SR
End With

With RS_DoubleRoom
 .MoveFirst
 DR = .RecordCount
 lblDoubleRoomValue(1).Caption = DR
End With

With RS_SuiteRoom
 .MoveFirst
 SUR = .RecordCount
 lblSuiteRoomValue(1).Caption = SUR
End With
 
With RS_DeluxeSuite
 .MoveFirst
 DSUR = .RecordCount
 lblDeluxeSuiteValue(1).Caption = DSUR
End With
 i = SR + DR + SUR + DSUR
 lbltotalroomvalue1.Caption = i
 J = lbltotalroomvalue(0).Caption - i
  lbloccupiedRoomvalue1.Caption = J
 
 SR1 = lblSingleRoomValue(0).Caption - SR
 lblSingleRoomValue(2).Caption = SR1
 
 DR1 = lblDoubleRoomValue(0).Caption - DR
 lblDoubleRoomValue(2).Caption = DR1
 
 SUR1 = lblSuiteRoomValue(0).Caption - SUR
 lblSuiteRoomValue(2).Caption = SUR1
 
 DSUR1 = lblDeluxeSuiteValue(0).Caption - DSUR
 lblDeluxeSuiteValue(2).Caption = DSUR1
End Sub

