VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   0  'None
   Caption         =   "User Password"
   ClientHeight    =   3270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5625
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   -60
      TabIndex        =   12
      Top             =   2505
      Width           =   5730
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   30
      Top             =   885
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change &Time(F6)"
      Height          =   525
      Left            =   495
      TabIndex        =   0
      Top             =   1815
      Width           =   1035
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   405
      Left            =   2805
      TabIndex        =   3
      Top             =   2700
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   405
      Left            =   4185
      TabIndex        =   4
      Top             =   2700
      Width           =   1275
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1650
      PasswordChar    =   "•"
      TabIndex        =   2
      Top             =   1335
      Width           =   2940
   End
   Begin VB.ComboBox cboUser 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1650
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   795
      Width           =   2940
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "hh:mm:ss"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2910
      TabIndex        =   11
      Top             =   2145
      Width           =   930
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "mm-dd-yyyy"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2910
      TabIndex        =   10
      Top             =   1815
      Width           =   1260
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "System Time:"
      Height          =   195
      Left            =   1590
      TabIndex        =   9
      Top             =   2145
      Width           =   1200
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "System Date:"
      Height          =   195
      Left            =   1620
      TabIndex        =   8
      Top             =   1815
      Width           =   1185
   End
   Begin VB.Image Image3 
      Height          =   240
      Left            =   4665
      Picture         =   "Login.frx":0000
      Top             =   825
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   4665
      Picture         =   "Login.frx":038A
      Top             =   1365
      Width           =   240
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   360
      Picture         =   "Login.frx":0714
      Top             =   45
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Password:"
      Height          =   195
      Left            =   645
      TabIndex        =   7
      Top             =   1395
      Width           =   885
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Username:"
      Height          =   195
      Left            =   585
      TabIndex        =   6
      Top             =   855
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To begin Select a Username"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1005
      TabIndex        =   5
      Top             =   195
      Width           =   2415
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   -30
      Picture         =   "Login.frx":13DE
      Stretch         =   -1  'True
      Top             =   -15
      Width           =   5670
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Call forminit
End Sub

Sub forminit()
    'Call CenterForm(frmLogin)
    Me.Top = 1700
    Me.Left = 5000
    Call loadcbouser
End Sub

Private Sub cboUser_Click()
    txtPassword.SetFocus
End Sub

Private Sub cmdCancel_Click()
    End
End Sub

Private Sub Command1_Click()
    Shell "rundll32.exe shell32.dll,Control_RunDLL timedate.cpl", vbNormalFocus
End Sub

Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
    If vbKeyF6 Then
        Command1_Click
    End If
End Sub

Sub loadcbouser()
    Dim strSQL As String
    
    strSQL = "SELECT * FROM tblUserInfo"
    
    rs.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    With rs
        Do While Not rs.EOF
            cboUser.AddItem !sUserName
            cboUser.ItemData(cboUser.NewIndex) = CLng(!iUserID)
            .MoveNext
        Loop
    End With
    Set rs = Nothing
End Sub

Private Sub cmdOK_Click()
    If cboUser = "" Then cboUser.SetFocus: Exit Sub
    If txtPassword = "" Then txtPassword.SetFocus: Exit Sub
    
    Dim rsUsers As Recordset
    Static attempt As Integer
    
    Set rsUsers = New ADODB.Recordset
    'flgFirstUse = 2
    With rsUsers
        If .State = adStateOpen Then .Close
        .Open "SELECT * FROM tbluserinfo WHERE sUserName='" & cboUser & "' AND sUserName='" & txtPassword & "'", cn, adOpenDynamic, adLockOptimistic
        If Not .EOF Then
            Unload Me
            If !sUserTaskLevel = "Administrator" Then
                frmMain.Text1 = !sUserPassword
                frmMain.Text2 = !sUserTaskLevel
                frmMain.Frame1.Visible = True
                With frmMain
                    .Toolbar1.Visible = True
                    .mnuFMaintenance.Visible = True
                    .mnuFSecurity.Visible = True
                    .mnuTransaction.Visible = True
                    .mnuTDelivery.Visible = True
                    .mnuTProduct.Visible = True
                    .mnuTReturn.Visible = True
                    .mnuView.Visible = True
                    .mnuVSummaryInfo.Visible = True
                    .mnuVProducts.Visible = True
                    .mnuReportS.Visible = True
                    .Frame1.Visible = True
                    .Toolbar1.Buttons(2).Enabled = True
                    .Toolbar1.Buttons(3).Enabled = True
                    .Toolbar1.Buttons(5).Enabled = True
                    .Toolbar1.Buttons(7).Enabled = True
                    .Toolbar1.Buttons(9).Enabled = True
                    .Toolbar1.Buttons(10).Enabled = True
                End With
            ElseIf !sUserTaskLevel = "Cashier" Then
                frmMain.Text1 = !sUserPassword
                frmMain.Text2 = !sUserTaskLevel
                With frmMain
                    .Toolbar1.Visible = True
                    .mnuFMaintenance.Visible = False
                    .mnuFSecurity.Visible = False
                    .mnuTDelivery.Visible = False
                    .mnuTProduct.Visible = False
                    .mnuTReturn.Visible = False
                    .mnuView.Visible = False
                    .mnuReportS.Visible = True
                    .Frame1.Visible = False
                    .Toolbar1.Buttons(2).Enabled = False
                    .Toolbar1.Buttons(3).Enabled = False
                    .Toolbar1.Buttons(5).Enabled = False
                    .Toolbar1.Buttons(7).Enabled = True
                    .Toolbar1.Buttons(9).Enabled = False
                    .Toolbar1.Buttons(10).Enabled = True
                End With
                    
            Else
                frmMain.Text1 = !sUserPassword
                frmMain.Text2 = !sUserTaskLevel
                With frmMain
                    .Toolbar1.Visible = True
                    .mnuFMaintenance.Visible = False
                    .mnuFSecurity.Visible = False
                    .mnuTransaction.Visible = False
                    .mnuView.Visible = True
                    .mnuVSummaryInfo.Visible = False
                    .mnuVProducts.Visible = True
                    .mnuReportS.Visible = False
                    .Frame1.Visible = False
                    .Toolbar1.Buttons(2).Enabled = False
                    .Toolbar1.Buttons(3).Enabled = False
                    .Toolbar1.Buttons(5).Enabled = False
                    .Toolbar1.Buttons(7).Enabled = False
                    .Toolbar1.Buttons(9).Enabled = True
                    .Toolbar1.Buttons(10).Enabled = False
                End With
            End If
        Else
            attempt = attempt + 1
            MsgBox "A C C E S S   D E N I E D " & vbCrLf & _
            "Please type the correct password", vbCritical, "This is your " & attempt & " attemp"
            Call Highlight(txtPassword)
            If attempt = 3 Then
                MsgBox "This will terminate the applicatin", vbCritical, "You already used all attempt"
                End
            End If
        End If
    End With
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        cboUser.ListIndex = -1
        txtPassword = ""
        cmdCancel.SetFocus
    End If
End Sub

Private Sub Timer1_Timer()
    Label6.Caption = Format(Date, "mmmm dd, yyyy")
    Label7.Caption = Format(Time, "hh:mm:ss")
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdOK_Click
    End If
End Sub
