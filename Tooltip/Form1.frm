VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleMode       =   0  'User
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   3360
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   1440
      Width           =   480
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   1320
      Top             =   2880
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      ToolTipText     =   "It's not used any more !"
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0C000&
      Caption         =   "Please check!"
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "E-mail  :  nobody@post.hinet.hr"
      ForeColor       =   &H0000FFFF&
      Height          =   975
      Left            =   4920
      TabIndex        =   6
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Vote for this !"
      ForeColor       =   &H0000FFFF&
      Height          =   975
      Left            =   840
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "E x i t"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Index           =   1
      Left            =   6840
      MouseIcon       =   "Form1.frx":08CA
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   30
      Width           =   450
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tooltip Multiline 1.0  By Nobody  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   300
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7470
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim e As Integer
Dim e1 As Integer
Dim z As Integer

Private Sub Check1_Click()
   Unload Form2
   If Check1.Value = vbChecked Then
   e1 = 16
   Else
   e1 = 7
   End If
   e = 0
End Sub
Private Sub Check1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
e = 1
End Sub

Private Sub Check1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If z = 0 Then z = 1: Exit Sub
   If e = 1 Then Exit Sub
   X1 = X: Y1 = Y
   Top1 = Check1.Top: Left1 = Check1.Left
   Call Module1.Tooltip
   Form2.Label2.BackColor = vbRed
   Form2.Label2.ForeColor = vbYellow
   Form2.Label2.FontName = "OCR A Extended"
   Form2.Label2.FontSize = e1
   Form2.Label2.FontItalic = False
   Form2.Label2.Caption = Form2.Label2.Caption & "Today is : " _
   & Date & vbCrLf & "Time : " & Time
   Call Module1.RealLabel
End Sub

Private Sub Command1_Click()
   End
End Sub

Private Sub Form_Load()
   e1 = 7
   e = 0
   z = 1
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Unload Form2
End Sub
Private Sub Label1_Click(Index As Integer)
   End
End Sub
Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If z = 0 Then z = 1: Exit Sub
   X1 = X: Y1 = Y
   Top1 = Label2.Top: Left1 = Label2.Left
   Call Module1.Tooltip
   Form2.Label2.FontItalic = True
   Form2.Label2.FontName = "Bookman Old Style"
   Form2.Label2.FontSize = 12
   Form2.Label2.BackColor = vbBlue
   Form2.Label2.ForeColor = vbWhite
   Form2.Label2.Caption = Form2.Label2.Caption & "Simple the" _
   & vbCrLf & "best !"
   Call Module1.RealLabel
End Sub
Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If z = 0 Then z = 1: Exit Sub
   X1 = X: Y1 = Y
   Top1 = Label3.Top: Left1 = Label3.Left
   Call Module1.Tooltip
   Form2.Label2.FontItalic = False
   Form2.Label2.FontName = "MS Serif"
   Form2.Label2.FontSize = 7
   Form2.Label2.BackColor = vbBlue
   Form2.Label2.ForeColor = vbWhite
   Form2.Label2.Caption = Form2.Label2.Caption & "N o b o d y" _
   & vbCrLf & "       w a s " & vbCrLf & "     h e r e !"
   Call Module1.RealLabel
End Sub
Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If z = 0 Then z = 1: Exit Sub
   X1 = X: Y1 = Y
   Top1 = Picture1.Top: Left1 = Picture1.Left
   Call Module1.Tooltip
   Form2.Label2.BackColor = vbYellow
   Form2.Label2.ForeColor = vbBlack
   Form2.Label2.FontName = "Computerfont"
   Form2.Label2.FontSize = 10
   Form2.Label2.FontItalic = False
   Form2.Label2.Caption = Form2.Label2.Caption & "Only for you " _
   & vbCrLf & "another color."
   Call Module1.RealLabel
End Sub
Private Sub Timer1_Timer()
    If z = 1 Then
    z = 0
    Unload Form2
    End If
End Sub
 
 
    
   
   
 

 

