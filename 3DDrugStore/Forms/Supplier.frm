VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSupplier 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Supplier Maintenance"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6765
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Supplier.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   6765
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFind 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   780
      TabIndex        =   11
      Top             =   3615
      Visible         =   0   'False
      Width           =   2535
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
      Height          =   2700
      Left            =   135
      TabIndex        =   5
      Top             =   720
      Width           =   6495
      Begin VB.TextBox txtSupplierNotes 
         Appearance      =   0  'Flat
         Height          =   555
         Left            =   1785
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   1965
         Width           =   4500
      End
      Begin VB.TextBox txtSupplierContactNumber 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1785
         TabIndex        =   3
         Top             =   1599
         Width           =   2100
      End
      Begin VB.TextBox txtSupplierContactPerson 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1785
         TabIndex        =   2
         Top             =   1236
         Width           =   3030
      End
      Begin VB.TextBox txtSupplierName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         TabIndex        =   0
         Top             =   240
         Width           =   4500
      End
      Begin VB.TextBox txtSupplierAddress 
         Appearance      =   0  'Flat
         Height          =   555
         Left            =   1785
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   603
         Width           =   3375
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Notes:"
         Height          =   195
         Left            =   1080
         TabIndex        =   14
         Top             =   1965
         Width           =   555
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Contact Number:"
         Height          =   195
         Left            =   165
         TabIndex        =   10
         Top             =   1644
         Width           =   1470
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Contact Person:"
         Height          =   195
         Left            =   255
         TabIndex        =   9
         Top             =   1281
         Width           =   1380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   195
         Left            =   1065
         TabIndex        =   7
         Top             =   285
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Address:"
         Height          =   195
         Left            =   870
         TabIndex        =   6
         Top             =   603
         Width           =   765
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   330
      Left            =   4890
      TabIndex        =   8
      Top             =   150
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "add"
            Object.ToolTipText     =   "Add New Supplier"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "update"
            Object.ToolTipText     =   "Update Existing Supplier"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "delete"
            Object.ToolTipText     =   "Delete Existing Supplier"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "search"
            Object.ToolTipText     =   "Find Supplier"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   450
      Top             =   585
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
            Picture         =   "Supplier.frx":038A
            Key             =   "add"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Supplier.frx":0724
            Key             =   "edit"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Supplier.frx":0ABE
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Supplier.frx":0E58
            Key             =   "search"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lsvSupplier 
      Height          =   1320
      Left            =   135
      TabIndex        =   12
      Top             =   3990
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   2328
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "id"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Phone Number"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Contact Person"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Notes"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.TextBox txtDummy 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3510
      TabIndex        =   15
      Top             =   3615
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Name:"
      Height          =   195
      Left            =   135
      TabIndex        =   13
      Top             =   3660
      Visible         =   0   'False
      Width           =   570
   End
End
Attribute VB_Name = "frmSupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    Call forminit
End Sub

Sub forminit()
    Call CenterForm(frmSupplier)
    Call textclear
    Call loadlsvSupplier
    Toolbar1.Buttons(4).Enabled = False
    Toolbar1.Buttons(5).Enabled = False
End Sub

Function complete()
    If txtSupplierName = "" Or txtSupplierAddress = "" Or txtSupplierContactPerson = "" _
        Or txtSupplierContactNumber = "" Then
        complete = False
    Else
        complete = True
    End If
End Function

Sub textclear()
    txtSupplierName = ""
    txtSupplierAddress = ""
    txtSupplierContactPerson = ""
    txtSupplierContactNumber = ""
    txtSupplierNotes = ""
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Call textclear
        txtSupplierName.SetFocus
        Toolbar1.Buttons(2).Enabled = True
        Toolbar1.Buttons(4).Enabled = False
        Toolbar1.Buttons(5).Enabled = False
    End If
End Sub

Private Sub lsvSupplier_Click()
    Dim X As Integer
    Dim strSQL As String
    Dim dummy
    Dim row
    Toolbar1.Buttons(2).Enabled = False
    Toolbar1.Buttons(4).Enabled = True
    Toolbar1.Buttons(5).Enabled = True
    
    row = lsvSupplier.SelectedItem.Index
    dummy = lsvSupplier.ListItems.Item(row).Text
    
    strSQL = "SELECT * FROM tblSupplier "
    strSQL = strSQL & "WHERE iSupplierId=" & dummy
    
    rs.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    
    With rs
        txtdummy = dummy
        txtSupplierName = !sSupplierName
        txtSupplierAddress = !sSupplierAddress
        txtSupplierContactPerson = !sSupplierContact
        txtSupplierContactNumber = !sSupplierNumber
        txtSupplierNotes = !sSupplierNotes
    End With
        txtSupplierName.SetFocus
    Set rs = Nothing
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "add"
            Dim strSQLAdd As String
            Dim rsAdd As New Recordset
            If complete = True Then
                strSQLAdd = "SELECT * FROM  tblSupplier"
                rsAdd.Open strSQLAdd, cn, adOpenDynamic, adLockOptimistic
                With rsAdd
                    .AddNew
                    !sSupplierName = txtSupplierName
                    !sSupplierAddress = txtSupplierAddress
                    !sSupplierContact = txtSupplierContactPerson
                    !sSupplierNumber = txtSupplierContactNumber
                    !sSupplierNotes = txtSupplierNotes
                    .Update
                End With
                    MsgBox "New record added to the database", vbOKOnly, "Supplier Maintenance"
                    Call textclear
                    Call loadlsvSupplier
                    txtSupplierName.SetFocus
                    Set rsAdd = Nothing
            Else
                If txtSupplierName = "" Then
                    MsgBox "Please enter supplier name", vbOKOnly, "Supplier Maintenance"
                    txtSupplierName.SetFocus
                    Exit Sub
                ElseIf txtSupplierAddress = "" Then
                    MsgBox "Please enter supplier address", vbOKOnly, "Supplier Maintenance"
                    txtSupplierAddress.SetFocus
                    Exit Sub
                ElseIf txtSupplierContactPerson = "" Then
                    MsgBox "Please enter supplier contact person", vbOKOnly, "Supplier Maintenance"
                    txtSupplierContactPerson.SetFocus
                    Exit Sub
                ElseIf txtSupplierContactNumber = "" Then
                    MsgBox "Please enter supplier contact number", vbOKOnly, "Supplier Maintenace"
                    txtSupplierContactNumber.SetFocus
                End If
            End If
        Case "update"
            Dim strSQLEdit As String
            Dim rsEdit As New Recordset
            If complete = True Then
                strSQLEdit = "SELECT * FROM  tblSupplier"
                strSQLEdit = strSQLEdit & " WHERE iSupplierID=" & txtdummy
                rsEdit.Open strSQLEdit, cn, adOpenDynamic, adLockOptimistic
                With rsEdit
                    !sSupplierName = txtSupplierName
                    !sSupplierAddress = txtSupplierAddress
                    !sSupplierContact = txtSupplierContactPerson
                    !sSupplierNumber = txtSupplierContactNumber
                    !sSupplierNotes = txtSupplierNotes
                    .Update
                End With
                    MsgBox "The changes you made was successfully updated", vbOKOnly, "Supplier Maintenance"
                    Call textclear
                    Call loadlsvSupplier
                    txtSupplierName.SetFocus
                    Set rsEdit = Nothing
            Else
                If txtSupplierName = "" Then
                    MsgBox "Please enter supplier name", vbOKOnly, "Supplier Maintenance"
                    txtSupplierName.SetFocus
                    Exit Sub
                ElseIf txtSupplierEditress = "" Then
                    MsgBox "Please enter supplier Editress", vbOKOnly, "Supplier Maintenance"
                    txtSupplierEditress.SetFocus
                    Exit Sub
                ElseIf txtSupplierContactPerson = "" Then
                    MsgBox "Please enter supplier contact person", vbOKOnly, "Supplier Maintenance"
                    txtSupplierContactPerson.SetFocus
                    Exit Sub
                ElseIf txtSupplierContactNumber = "" Then
                    MsgBox "Please enter supplier contact number", vbOKOnly, "Supplier Maintenace"
                    txtSupplierContactNumber.SetFocus
                End If
            End If
        Case "delete"
            Dim strSQLDelete As String
            Dim rsDelete As New ADODB.Recordset
            Dim answerDelete As String
            
            strSQLDelete = "SELECT * FROM tblSupplier "
            strSQLDelete = strSQLDelete & "WHERE iSupplierID=" & txtdummy
            
            answerDelete = MsgBox("Are you sure you want DELETE this record?", vbQuestion + vbYesNo, "Supplier Maintenance")
            
            If answerDelete = vbYes Then
                rsDelete.Open strSQLDelete, cn, adOpenDynamic, adLockOptimistic
                With rsDelete
                    .Delete
                End With
                MsgBox "The record was successfully deleted", , "Supplier Maintenance"
            End If
            Set rsDelete = Nothing
            Call textclear
            txtFind = ""
            txtFind.Visible = True
            Label3.Visible = True
            txtFind.SetFocus
            Call loadlsvSupplier
        Case "search"
            Label5.Visible = True
            txtFind.Visible = True
            txtFind.SetFocus
    End Select
    
End Sub

Sub loadlsvSupplier()
    Dim X
    Dim strSQL As String
    
    strSQL = "SELECT * FROM tblSupplier"
    
    rs.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    lsvSupplier.ListItems.Clear
    If rs.EOF = True Then
        lsvSupplier.ListItems.Clear
    Else
        With rs
            Do While Not rs.EOF
                Set X = lsvSupplier.ListItems.Add(, , !iSupplierID)
                    X.SubItems(1) = !sSupplierName
                    X.SubItems(2) = !sSupplierNumber
                    X.SubItems(3) = !sSupplierContact
                    X.SubItems(4) = !sSupplierNotes
                    .MoveNext
            Loop
        End With
    End If
    Set rs = Nothing
End Sub

Private Sub txtFind_Change()
    Dim X
    Dim strSQL As String
    strSQL = "SELECT * FROM tblSupplier "
    strSQL = strSQL & "WHERE sSupplierName LIKE'" & txtFind.Text & "%'"
    
    rs.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    lsvSupplier.ListItems.Clear
    With rs
        Do While Not rs.EOF
        Set X = lsvSupplier.ListItems.Add(, , !iSupplierID)
            X.SubItems(1) = !sSupplierName
            X.SubItems(2) = !sSupplierNumber
            X.SubItems(3) = !sSupplierContact
            X.SubItems(4) = !sSupplierNotes
        .MoveNext
        Loop
    End With
    Set rs = Nothing
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyCode = vbKeyDown Then
        If lsvSupplier.ListItems.Count <> 0 Then
            lsvSupplier.SetFocus
        End If
    End If
End Sub
