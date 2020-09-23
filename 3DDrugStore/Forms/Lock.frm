VERSION 5.00
Begin VB.Form frmLock 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5340
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
   ScaleHeight     =   3030
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4665
      Top             =   2430
   End
   Begin VB.TextBox txtUnlock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      IMEMode         =   3  'DISABLE
      Left            =   683
      PasswordChar    =   "•"
      TabIndex        =   1
      Top             =   1440
      Width           =   3975
   End
   Begin VB.Image imgkey 
      Height          =   315
      Left            =   2760
      Picture         =   "Lock.frx":0000
      Stretch         =   -1  'True
      Top             =   2370
      Width           =   330
   End
   Begin VB.Image imgclose 
      Height          =   480
      Left            =   2220
      Picture         =   "Lock.frx":0442
      Top             =   2280
      Width           =   480
   End
   Begin VB.Image imgopen 
      Height          =   480
      Left            =   2220
      Picture         =   "Lock.frx":0884
      Top             =   2280
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type the correct password to UNLOCK the system"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   855
      TabIndex        =   0
      Top             =   210
      Width           =   4290
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   210
      Picture         =   "Lock.frx":0CC6
      Top             =   60
      Width           =   480
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   0
      Picture         =   "Lock.frx":1990
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5760
   End
End
Attribute VB_Name = "frmLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub forminit()
    imgkey.Left = 3120
    Timer1_Timer
End Sub

Private Sub Form_Activate()
    Call forminit
    Me.Top = 2000
    Me.Left = 5000
End Sub

Private Sub Timer1_Timer()
    If imgclose.Visible = True Then
        imgclose.Visible = False
        imgopen.Visible = True
        imgkey.Left = 2760
    ElseIf imgopen.Visible = True Then
        imgopen.Visible = False
        imgclose.Visible = True
        imgkey.Left = 3120
    End If
End Sub

Private Sub txtUnlock_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtUnlock = frmMain.Text1 Then
            Unload Me
        Else
            MsgBox "           I N V A L I D   P A S S W O R D            " & vbCrLf _
                & "Please type the correct password to unlock the system", vbCritical + vbOKOnly, "3D Drug Store POS"
                Call Highlight(txtUnlock)
        End If
    End If
End Sub
