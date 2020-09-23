VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUserInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Profile"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6825
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "UserInfo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   6825
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lsvUsers 
      Height          =   2295
      Left            =   4605
      TabIndex        =   17
      Top             =   960
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   4048
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "id"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Username"
         Object.Width           =   4057
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Full Name"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.CheckBox chkViewPassword 
      Caption         =   "&View Password"
      Height          =   330
      Left            =   4605
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3195
      Left            =   90
      TabIndex        =   7
      Top             =   540
      Width           =   4440
      Begin VB.TextBox txtdummy 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   3855
         TabIndex        =   19
         Top             =   615
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtUsername 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1785
         TabIndex        =   3
         Top             =   1470
         Width           =   2490
      End
      Begin VB.TextBox txtConfirmPassword 
         Appearance      =   0  'Flat
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1785
         PasswordChar    =   "•"
         TabIndex        =   6
         Top             =   2745
         Width           =   2490
      End
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1785
         PasswordChar    =   "•"
         TabIndex        =   5
         Top             =   2325
         Width           =   2490
      End
      Begin VB.ComboBox cboUserLevel 
         Height          =   315
         ItemData        =   "UserInfo.frx":038A
         Left            =   1785
         List            =   "UserInfo.frx":0397
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1893
         Width           =   2490
      End
      Begin VB.TextBox txtUserLastName 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1785
         TabIndex        =   2
         Top             =   1044
         Width           =   2490
      End
      Begin VB.TextBox txtUserMi 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1785
         TabIndex        =   1
         Top             =   627
         Width           =   570
      End
      Begin VB.TextBox txtUserFirstName 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1785
         TabIndex        =   0
         Top             =   210
         Width           =   2490
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Username:"
         Height          =   195
         Left            =   780
         TabIndex        =   15
         Top             =   1521
         Width           =   945
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Confirm Password:"
         Height          =   195
         Left            =   90
         TabIndex        =   13
         Top             =   2798
         Width           =   1635
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Password:"
         Height          =   195
         Left            =   840
         TabIndex        =   12
         Top             =   2378
         Width           =   885
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "User Task Level:"
         Height          =   195
         Left            =   285
         TabIndex        =   11
         Top             =   1953
         Width           =   1440
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Last Name:"
         Height          =   195
         Left            =   750
         TabIndex        =   10
         Top             =   1097
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "M.I.:"
         Height          =   195
         Left            =   1320
         TabIndex        =   9
         Top             =   680
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "First Name:"
         Height          =   195
         Left            =   735
         TabIndex        =   8
         Top             =   263
         Width           =   990
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   330
      Left            =   5490
      TabIndex        =   14
      Top             =   150
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "add"
            Object.ToolTipText     =   "Add New User"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "update"
            Object.ToolTipText     =   "Update Existing User"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "delete"
            Object.ToolTipText     =   "Delete Existing User"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5910
      Top             =   900
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserInfo.frx":03BA
            Key             =   "add"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserInfo.frx":0754
            Key             =   "edit"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserInfo.frx":0AEE
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserInfo.frx":0E88
            Key             =   "search"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "USERS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   4605
      TabIndex        =   18
      Top             =   645
      Width           =   2130
   End
End
Attribute VB_Name = "frmUserInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
     Call forminit
End Sub

Sub forminit()
    Call CenterForm(frmUserInfo)
    Toolbar1.Buttons(4).Enabled = False
    Toolbar1.Buttons(5).Enabled = False
    Call loadlsvusers
End Sub

Private Sub chkViewPassword_Click()
    If chkViewPassword.Value = 1 Then
        txtPassword.PasswordChar = ""
        txtConfirmPassword.PasswordChar = ""
    Else
        txtPassword.PasswordChar = "•"
        txtConfirmPassword.PasswordChar = "•"
    End If
End Sub

Function complete()
    If txtUserFirstName = "" Or txtUserMi = "" Or txtUserLastName = "" Or _
        txtUsername = "" Or cboUserLevel = "" Or txtPassword = "" Or _
        txtConfirmPassword = "" Then
        complete = False
    Else
        complete = True
    End If
End Function

Sub textclear()
    txtUserFirstName = ""
    txtUserMi = ""
    txtUserLastName = ""
    txtUsername = ""
    cboUserLevel.ListIndex = -1
    txtPassword = ""
    txtConfirmPassword = ""
    chkViewPassword.Value = 0
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Call textclear
        Toolbar1.Buttons(2).Enabled = True
        Toolbar1.Buttons(4).Enabled = False
        Toolbar1.Buttons(5).Enabled = False
        txtUserFirstName.SetFocus
    End If
End Sub

Private Sub lsvUsers_Click()
    Dim X As Integer
    Dim strSQL As String
    Dim dummy
    Dim row
    Toolbar1.Buttons(2).Enabled = False
    Toolbar1.Buttons(4).Enabled = True
    Toolbar1.Buttons(5).Enabled = True
    
    row = lsvUsers.SelectedItem.Index
    dummy = lsvUsers.ListItems.Item(row).Text
    
    strSQL = "SELECT * FROM tblUserInfo "
    strSQL = strSQL & "WHERE iUserId=" & dummy
    
    rs.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    
    With rs
        txtdummy = dummy
        txtUserFirstName = !sUserFirstname
        txtUserMi = !susermi
        txtUserLastName = !sUserLastname
        txtUsername = !sUserName
        cboUserLevel = !sUserTaskLevel
        txtPassword = !sUserPassword
        txtConfirmPassword = !sCUserPassword
    End With
        txtUserFirstName.SetFocus
    Set rs = Nothing
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "add"
            Dim strSQLAdd As String
            Dim rsAdd As New Recordset
            If complete = True Then
                If confirmpass = True Then
                    strSQLAdd = "SELECT * FROM  tblUserInfo"
                    rsAdd.Open strSQLAdd, cn, adOpenDynamic, adLockOptimistic
                    With rsAdd
                        .AddNew
                        !sUserName = txtUsername
                        !sUserPassword = txtPassword
                        !sCUserPassword = txtConfirmPassword
                        !sUserLastname = txtUserLastName
                        !susermi = txtUserMi
                        !sUserFirstname = txtUserFirstName
                        !sUserTaskLevel = cboUserLevel
                        .Update
                    End With
                        MsgBox "New record added to the database", vbOKOnly, "User Maintenance"
                        Call textclear
                        Call loadlsvusers
                        txtUserFirstName.SetFocus
                        Set rsAdd = Nothing
                Else
                    MsgBox "Re - type you password", vbOKOnly, "Password mismatch"
                    txtConfirmPassword.SetFocus
                    Exit Sub
                End If
            Else
                If txtUserFirstName = "" Then
                    MsgBox "Please enter user first name", vbOKOnly, "User Maintenance"
                    txtUserFirstName.SetFocus
                    Exit Sub
                ElseIf txtUserMi = "" Then
                    MsgBox "Please enter user middle initial", vbOKOnly, "User Maintenance"
                    txtUserMi.SetFocus
                    Exit Sub
                ElseIf txtUserLastName = "" Then
                    MsgBox "Please enter user last name", vbOKOnly, "User Maintenance"
                    txtUserLastName.SetFocus
                    Exit Sub
                ElseIf txtUsername = "" Then
                    MsgBox "Please enter username", vbOKOnly, "User Maintenance"
                    txtUsername.SetFocus
                    Exit Sub
                ElseIf cboUserLevel = "" Then
                    MsgBox "Please specify user task level", vbOKOnly, "User Maintenance"
                    cboUserLevel.SetFocus
                    Exit Sub
                ElseIf txtPassword = "" Then
                    MsgBox "Please enter user password", vbOKOnly, "User Maintenance"
                    txtPassword.SetFocus
                    Exit Sub
                ElseIf txtConfirmPassword = "" Then
                    MsgBox "Please confirm you password", vbOKOnly, "User Maintenace"
                    txtConfirmPassword.SetFocus
                    Exit Sub
               End If
            End If
        Case "update"
            Dim strSQLEdit As String
            Dim rsEdit As New Recordset
            If complete = True Then
                If confirmpass = True Then
                    strSQLEdit = "SELECT * FROM  tblUserInfo"
                    strSQLEdit = strSQLEdit & " WHERE iUserID=" & txtdummy
                    rsEdit.Open strSQLEdit, cn, adOpenDynamic, adLockOptimistic
                    With rsEdit
                        !sUserName = txtUsername
                        !sUserPassword = txtPassword
                        !sCUserPassword = txtConfirmPassword
                        !sUserLastname = txtUserLastName
                        !susermi = txtUserMi
                        !sUserFirstname = txtUserFirstName
                        !sUserTaskLevel = cboUserLevel
                        .Update
                    End With
                        MsgBox "The changes you made was successfully updated", vbOKOnly, "User Maintenance"
                        Call textclear
                        Call loadlsvusers
                        txtUserFirstName.SetFocus
                        Set rsEdit = Nothing
                Else
                    MsgBox "Re - type you password", vbOKOnly, "Password mismatch"
                    txtConfirmPassword.SetFocus
                    Exit Sub
                End If
            Else
                If txtUserFirstName = "" Then
                    MsgBox "Please enter user first name", vbOKOnly, "User Maintenance"
                    txtUserFirstName.SetFocus
                    Exit Sub
                ElseIf txtUserMi = "" Then
                    MsgBox "Please enter user middle initial", vbOKOnly, "User Maintenance"
                    txtUserMi.SetFocus
                    Exit Sub
                ElseIf txtUserLastName = "" Then
                    MsgBox "Please enter user last name", vbOKOnly, "User Maintenance"
                    txtUserLastName.SetFocus
                    Exit Sub
                ElseIf txtUsername = "" Then
                    MsgBox "Please enter username", vbOKOnly, "User Maintenance"
                    txtUsername.SetFocus
                    Exit Sub
                ElseIf cboUserLevel = "" Then
                    MsgBox "Please specify user task level", vbOKOnly, "User Maintenance"
                    cboUserLevel.SetFocus
                    Exit Sub
                ElseIf txtPassword = "" Then
                    MsgBox "Please enter user password", vbOKOnly, "User Maintenance"
                    txtPassword.SetFocus
                    Exit Sub
                ElseIf txtConfirmPassword = "" Then
                    MsgBox "Please confirm you password", vbOKOnly, "User Maintenace"
                    txtConfirmPassword.SetFocus
                    Exit Sub
               End If
            End If
            
        Case "delete"
            Dim strSQLDelete As String
            Dim rsDelete As New ADODB.Recordset
            Dim answerDelete As String
            
            strSQLDelete = "SELECT * FROM tblUserInfo "
            strSQLDelete = strSQLDelete & "WHERE iUserID=" & txtdummy
            
            answerDelete = MsgBox("Are you sure you want DELETE this record?", vbQuestion + vbYesNo, "Category Maintenance")
            
            If answerDelete = vbYes Then
                rsDelete.Open strSQLDelete, cn, adOpenDynamic, adLockOptimistic
                With rsDelete
                    .Delete
                End With
                MsgBox "The record was successfully deleted", , "User Maintenance"
            End If
            Set rsDelete = Nothing
            Call textclear
            txtUserFirstName.SetFocus
            Call loadlsvusers
    End Select
End Sub

Function confirmpass()
    If txtPassword <> txtConfirmPassword Then
        confirmpass = False
    Else
        confirmpass = True
    End If
End Function

Sub loadlsvusers()
    On Error GoTo err_handler:
    Dim X
    Dim strSQL As String
    
    strSQL = "SELECT * FROM tblUserinfo"
    
    rs1.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    lsvUsers.ListItems.Clear
    If rs1.EOF = True Then
        lsvUsers.ListItems.Clear
    Else
        With rs1
            Do While Not rs1.EOF
                Set X = lsvUsers.ListItems.Add(, , !iUserID)
                    X.SubItems(1) = !sUserName
                    X.SubItems(2) = Trim(!sUserFirstname) & " " & Trim(!susermi) & " " & Trim(!sUserLastname)
                    .MoveNext
            Loop
        End With
        Set rs1 = Nothing
    End If
    Exit Sub
err_handler:
    Set rs1 = Nothing

End Sub
