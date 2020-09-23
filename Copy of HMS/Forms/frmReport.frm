VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReport 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Daily Evaluation Report"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11235
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   11235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Print"
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
      Left            =   9840
      TabIndex        =   15
      Top             =   6600
      Width           =   975
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
      Left            =   9360
      TabIndex        =   11
      Top             =   6120
      Width           =   1455
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
      Left            =   9360
      TabIndex        =   10
      Top             =   5760
      Width           =   1455
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
      Left            =   9360
      TabIndex        =   9
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Report Query"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   10455
      Begin VB.CommandButton Command3 
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
         Height          =   375
         Left            =   8880
         TabIndex        =   7
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Refresh"
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
         Left            =   7800
         TabIndex        =   6
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Search"
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
         Left            =   6720
         TabIndex        =   5
         Top             =   480
         Width           =   975
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
         Left            =   4200
         TabIndex        =   4
         Top             =   480
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
         TabIndex        =   3
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Search Criteria"
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
         TabIndex        =   2
         Top             =   480
         Width           =   1395
      End
   End
   Begin MSComctlLib.ListView lvReport 
      Height          =   2775
      Left            =   240
      TabIndex        =   0
      Top             =   2520
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   4895
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Total VAT"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7320
      TabIndex        =   14
      Top             =   6120
      Width           =   825
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Total Service Charge"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7320
      TabIndex        =   13
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Total Amount Received"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7320
      TabIndex        =   12
      Top             =   5400
      Width           =   1980
   End
   Begin VB.Label lblcount2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   720
      TabIndex        =   8
      Top             =   2160
      Width           =   555
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'fillgrid
If Combo1.ListIndex = -1 Then
MsgBox "Select a Search Criteria", vbOKOnly + vbCritical, "Error"
End If
'Exit Sub
If Text1.Text = "" Then
 MsgBox "Enter Text to Search", vbOKOnly + vbCritical, "Error"
 Text1.SetFocus
 Exit Sub
End If
If Combo1.ListIndex = 0 Then
Search_Name
ElseIf Combo1.ListIndex = 1 Then
Search_Date
ElseIf Combo1.ListIndex = 2 Then
'Search_payment
End If
End Sub

Private Sub Command2_Click()
Text1.Text = ""
Combo1.ListIndex = -1
filllist
End Sub

Private Sub Command3_Click()
Text1.Text = ""
Combo1.ListIndex = -1
Me.Hide
End Sub


Private Sub Command4_Click()
If lvReport.ListItems.count >= 1 Then
Printer.CurrentY = 200
Printer.CurrentX = 3000
Printer.Print UCase(Company)
Printer.CurrentX = 3200
Printer.Print Add
Printer.Print ""
Printer.CurrentX = 200
Printer.Print "User Name:  " & UCase(UserName) & "                                                                                                                                                                   Date: " & Date
Printer.Print ""
Printer.Print ""
Printer.CurrentX = 100
Printer.Print "                                                                                        DAILY EVALUATION REPORT  "
Printer.Print "----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
Printer.CurrentY = 1700
Printer.CurrentX = 200
Printer.Print "S/No"
Printer.CurrentY = 1700
Printer.CurrentX = 1000
Printer.Print "Payment Type"
Printer.CurrentY = 1700
Printer.CurrentX = 2500
Printer.Print "Name"
Printer.CurrentY = 1700
Printer.CurrentX = 4500
Printer.Print "Room No."
Printer.CurrentY = 1700
Printer.CurrentX = 5500
Printer.Print "Payment Date"
Printer.CurrentY = 1700
Printer.CurrentX = 7000
Printer.Print "Amount"
Printer.CurrentY = 1700
Printer.CurrentX = 8500
Printer.Print "Service Charge"
Printer.CurrentY = 1700
Printer.CurrentX = 10000
Printer.Print "V.A.T"
'Printer.CurrentX = 100
Printer.Print "----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
'For i = lvReport. To lvReport.ListItems.Add
'Do While lvReport.ListItems

Printer.CurrentY = 2100
Printer.CurrentX = 1000
Printer.Print lvReport.ListItems.Item(i).SubItems(1)
Printer.CurrentY = 2100
Printer.CurrentX = 2500
Printer.Print lvReport.ListItems.Item(i).SubItems(3)
Printer.CurrentY = 2100
Printer.CurrentX = 4500
Printer.Print lvReport.ListItems.Item(i).SubItems(4)
Printer.CurrentY = 2100
Printer.CurrentX = 5500
Printer.Print lvReport.ListItems.Item(i).SubItems(5)
Printer.CurrentY = 2100
Printer.CurrentX = 7000
Printer.Print lvReport.ListItems.Item(i).SubItems(6)
Printer.CurrentY = 2100
Printer.CurrentX = 8500
Printer.Print lvReport.ListItems.Item(i).SubItems(7)
Printer.CurrentY = 2100
Printer.CurrentX = 10000
Printer.Print lvReport.ListItems.Item(i).SubItems(8)
'Wend
End If
End Sub

Private Sub Form_Load()
Top = 3000
Left = 3000
Connect
filllist
Combo1.AddItem "Name"
Combo1.AddItem "Date"
Combo1.AddItem "Payment Type"
lblcount2.Caption = ""
'fillgrid
End Sub
Sub filllist()
With lvReport
  
    .view = lvwReport
    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "S/No"
    .ColumnHeaders.Add , , "Payment Type"
    .ColumnHeaders.Add , , "Guest ID"
     .ColumnHeaders(3).Width = 2500
    .ColumnHeaders.Add , , "Name"
     .ColumnHeaders(4).Width = 3000
    .ColumnHeaders.Add , , "Room No"
    .ColumnHeaders.Add , , "Date of Payment"
    .ColumnHeaders(6).Width = 2500
    .ColumnHeaders.Add , , "Amount"
    .ColumnHeaders.Add , , "Service Charge"
    .ColumnHeaders(8).Width = 2500
    .ColumnHeaders.Add , , "VAT"
  

End With
Call Connect
If RS_Paymentlog.State = adStateOpen Then RS_Paymentlog.Close
RS_Paymentlog.Open "Select * from payment_log order by guestid", cnn, adOpenStatic, adLockOptimistic
lvReport.ListItems.Clear
With RS_Paymentlog
    While Not .EOF
    Set itm = lvReport.ListItems.Add(, , .RecordCount)
    itm.SubItems(1) = .Fields(3)
    itm.SubItems(2) = .Fields(0)
    itm.SubItems(3) = .Fields(1)
    itm.SubItems(4) = .Fields(5)
    itm.SubItems(5) = .Fields(10)
    itm.SubItems(6) = .Fields(6)
    itm.SubItems(7) = Val(.Fields(6)) / 100 * 10
    itm.SubItems(8) = Val(.Fields(6)) / 100 * 5
    .MoveNext
    Wend
    lblcount2.Caption = "Total Records = " & lvReport.ListItems.count
End With

Dim var As Integer
Dim var2 As Integer
Dim var3 As Integer
    Dim i As Integer
    'Declare variables
    '
    ListCount = lvReport.ListItems.count
    'count how many rows are in ListView
    '

    For i = 1 To ListCount
        'Go from 1st row to last
        var = var + lvReport.ListItems(i).SubItems(6)
        var2 = var2 + lvReport.ListItems(i).SubItems(7)
        var3 = var3 + lvReport.ListItems(i).SubItems(8)
        'each loop, add value from listview to v
        '     ar
        'replace the 1 in "(1)" with the column
        '     number
        '*column # starts at 0
        '(i) is the row nuber
    Next i

    'loop
    '
    Text2.Text = var
    Text3.Text = var2
    Text4.Text = var3
End Sub
Sub Search_Name()
 lvReport.ListItems.Clear

     ' record variables
    Dim mark As Variant
    Dim count As Integer
    
   Call Connect
              
    
    count = 0
    With RS_Payment
   'listview1= .RecordCount + 1
r = 1

    .Find "name LIKE '" & Text1.Text & "%'"
    '.MoveFirst
    Do While Not .EOF
        'continue if last find succeeded
    Set itm = lvReport.ListItems.Add(, , .Fields(3))
    itm.SubItems(1) = .Fields(0)
    itm.SubItems(2) = .Fields(1)
    itm.SubItems(3) = .Fields(5)
    itm.SubItems(4) = .Fields(11)
    itm.SubItems(5) = .Fields(6)
    itm.SubItems(6) = Val(.Fields(6)) / 100 * 10
    itm.SubItems(7) = Val(.Fields(6)) / 100 * 5

   
        'count the last title found
       count = count + 1
        ' note current position
       mark = .Bookmark
      .Find "name LIKE '" & Text1.Text & "%'", 1, adSearchForward, mark
        
        ' above code skips current record to avoid finding the same row repeatedly;
        ' last arg (bookmark) is redundant because Find searches from current position
    'r = r + 1
'.MoveNext
Loop
  
    '
    If count = 0 Then

     MsgBox "No Match Found", vbOKOnly + vbInformation, "Information"
     filllist
    Text1.SetFocus
    Else
     lblcount2.Caption = "Total Matches found " & count
    End If
     ' clean up
    RS_Payment.Close
    End With
'
End Sub

Sub Search_Date()
 lvReport.ListItems.Clear

     ' record variables
    Dim mark As Variant
    Dim count As Integer
    
   Call Connect
              
    
    count = 0
    With RS_Payment
   'listview1= .RecordCount + 1
r = 1

    .Find "Date_Modified LIKE '" & Text1.Text & "%'"
    '.MoveFirst
    Do While Not .EOF
        'continue if last find succeeded
    Set itm = lvReport.ListItems.Add(, , .Fields(3))
    itm.SubItems(1) = .Fields(0)
    itm.SubItems(2) = .Fields(1)
    itm.SubItems(3) = .Fields(5)
    itm.SubItems(4) = .Fields(11)
    itm.SubItems(5) = .Fields(6)
    itm.SubItems(6) = Val(.Fields(6)) / 100 * 10
    itm.SubItems(7) = Val(.Fields(6)) / 100 * 5

   
        'count the last title found
       count = count + 1
        ' note current position
       mark = .Bookmark
      .Find "Date_Modified LIKE '" & Text1.Text & "%'", 1, adSearchForward, mark
        
        ' above code skips current record to avoid finding the same row repeatedly;
        ' last arg (bookmark) is redundant because Find searches from current position
    'r = r + 1
'.MoveNext
Loop
  
    '
    If count = 0 Then

     MsgBox "No Match Found", vbOKOnly + vbInformation, "Information"
     filllist
    Text1.SetFocus
    Else
     lblcount2.Caption = "Total Matches found " & count
    End If
     ' clean up
    RS_Payment.Close
    End With
'
End Sub


Sub fillgrid()
Connect
'flex1.Clear
flex1.FormatString = "Enroll No|<Candidate Name              |Father Name           |Course Enrolled       |DOB               "
'If RS_Payment.State = adStateOpen Then RS_Payment.Close
If RS_Payment.RecordCount > 0 Then
flex1.Rows = RS_Payment.RecordCount + 1
r = 1
With RS_Payment
.MoveFirst
Do While Not .EOF
flex1.TextMatrix(r, 0) = RS_Payment!GuestID
If IsNull(!GuestID) = False Then flex1.TextMatrix(r, 1) = !GuestID
If IsNull(!Name) = False Then flex1.TextMatrix(r, 2) = !Name
If IsNull(!Payment_Type) = False Then flex1.TextMatrix(r, 3) = !Payment_Type
If IsNull(!payment) = False Then flex1.TextMatrix(r, 4) = !Advance
r = r + 1
.MoveNext
Loop
'End If
'End If
'End If
'End If
End With

'End If
End If
End Sub

