VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCategory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Category Maintenance"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4995
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Category.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   4995
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFind 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   825
      TabIndex        =   7
      Top             =   2460
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   330
      Left            =   2985
      TabIndex        =   6
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
            Object.ToolTipText     =   "Add New Category"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "update"
            Object.ToolTipText     =   "Update Existing Category"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "delete"
            Object.ToolTipText     =   "Delete Existing Category"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "search"
            Object.ToolTipText     =   "Find Category"
            ImageIndex      =   4
         EndProperty
      EndProperty
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
      Height          =   1545
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   4725
      Begin VB.TextBox txtDescription 
         Appearance      =   0  'Flat
         Height          =   795
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   600
         Width           =   3030
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   0
         Top             =   240
         Width           =   3030
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Description:"
         Height          =   195
         Left            =   225
         TabIndex        =   4
         Top             =   600
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   195
         Left            =   690
         TabIndex        =   3
         Top             =   285
         Width           =   570
      End
   End
   Begin MSComctlLib.ListView lsvCategory 
      Height          =   1140
      Left            =   135
      TabIndex        =   5
      Top             =   2835
      Width           =   4725
      _ExtentX        =   8334
      _ExtentY        =   2011
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
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Notes"
         Object.Width           =   5786
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   105
      Top             =   255
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
            Picture         =   "Category.frx":038A
            Key             =   "add"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Category.frx":0724
            Key             =   "edit"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Category.frx":0ABE
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Category.frx":0E58
            Key             =   "search"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtdummy 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3000
      TabIndex        =   9
      Top             =   2460
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Name:"
      Height          =   195
      Left            =   135
      TabIndex        =   8
      Top             =   2505
      Visible         =   0   'False
      Width           =   570
   End
End
Attribute VB_Name = "frmCategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    'call from the procedure sub forminit
    Call forminit
End Sub

'this procedure will be call later in the program during
'the Form_Activate
Sub forminit()
    Call CenterForm(frmCategory)
    Call loadlsvCategory
    Toolbar1.Buttons(4).Enabled = False
    Toolbar1.Buttons(5).Enabled = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'this code will be perform if you press escape key 27 is the corresponding key for escape key
    If KeyAscii = 27 Then
        Call textclear
        txtName.SetFocus
        txtFind = ""
        Toolbar1.Buttons(2).Enabled = True
    Toolbar1.Buttons(4).Enabled = False ' disabled the 4th toolbar
    Toolbar1.Buttons(5).Enabled = False ' disabled the 5th toolbar
    End If
End Sub

Private Sub lsvCategory_Click()
    'this code will perform when you click the list view control(lsvCategory)
    Dim X As Integer
    Dim strSQL As String
    Dim dummy
    Dim row
    Toolbar1.Buttons(2).Enabled = False
    Toolbar1.Buttons(4).Enabled = True
    Toolbar1.Buttons(5).Enabled = True
    
    row = lsvCategory.SelectedItem.Index
    dummy = lsvCategory.ListItems.Item(row).Text
    
    strSQL = "SELECT * FROM tblCategory "
    strSQL = strSQL & "WHERE icategoryId=" & dummy
    
    rs.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    
    With rs
        txtdummy = dummy
        txtName = !sCategoryName
        txtDescription = !sNotes
    End With
        txtName.SetFocus
    Set rs = Nothing
End Sub

Private Sub lsvCategory_KeyPress(KeyAscii As Integer)
    'when you press enter it call the procedure lsvCategory_Click
    lsvCategory_Click
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    'Codes for the toolbar
    Select Case Button.Key
        'codes for the the first button, add
        Case "add"
            Dim strSQLAdd As String
            Dim rsAdd As New Recordset
            If complete = True Then
                strSQLAdd = "SELECT * FROM  tblCategory"
                rsAdd.Open strSQLAdd, cn, adOpenDynamic, adLockOptimistic
                With rsAdd
                    .AddNew
                    !sCategoryName = txtName
                    !sNotes = txtDescription
                    .Update
                End With
                    MsgBox "New record added to the database", vbOKOnly, "Category Maintenance"
                    Call textclear
                    Call loadlsvCategory
                    txtName.SetFocus
                    Set rsAdd = Nothing
            Else
                If txtName = "" Then
                    MsgBox "Please enter category name", vbOKOnly, "Category Maintenance"
                    txtName.SetFocus
                    Exit Sub
                ElseIf txtDescription = "" Then
                    MsgBox "Please enter category notes", vbOKOnly, "Category Maintenance"
                    txtDescription.SetFocus
                    Exit Sub
                End If
            End If
        'codes for the second button, the update
        Case "update"
            Dim strSQLEdit As String
            Dim rsEdit As New Recordset
            If complete = True Then
                strSQLEdit = "SELECT * FROM  tblCategory"
                strSQLEdit = strSQLEdit & " WHERE iCategoryID=" & txtdummy
                
                rsEdit.Open strSQLEdit, cn, adOpenDynamic, adLockOptimistic
                With rsEdit
                    !sCategoryName = txtName
                    !sNotes = txtDescription
                    .Update
                End With
                    MsgBox "The changes you made was successfully updated", vbOKOnly, "Category Maintenance"
                    Call textclear
                    Call loadlsvCategory
                    txtName.SetFocus
                    Set rsEdit = Nothing
            Else
                If txtName = "" Then
                    MsgBox "Please enter category name", vbOKOnly, "Category Maintenance"
                    txtName.SetFocus
                    Exit Sub
                ElseIf txtDescription = "" Then
                    MsgBox "Please enter category notes", vbOKOnly, "Category Maintenance"
                    txtDescription.SetFocus
                    Exit Sub
                End If
            End If
        
        'codes for the third button, the delete
        Case "delete"
            Dim strSQLDelete As String
            Dim rsDelete As New ADODB.Recordset
            Dim answerDelete As String
            
            strSQLDelete = "SELECT * FROM tblCategory "
            strSQLDelete = strSQLDelete & "WHERE iCategoryID=" & txtdummy
            
            answerDelete = MsgBox("Are you sure you want DELETE this record?", vbQuestion + vbYesNo, "Category Maintenance")
            
            If answerDelete = vbYes Then
                rsDelete.Open strSQLDelete, cn, adOpenDynamic, adLockOptimistic
                With rsDelete
                    .Delete
                End With
                MsgBox "The record was successfully deleted", , "Category Maintenance"
            End If
            Set rsDelete = Nothing
            Call textclear
            txtFind = ""
            txtFind.Visible = True
            Label3.Visible = True
            txtFind.SetFocus
            Call loadlsvCategory
        
        'codes for the last button in the toolbar, the search
        Case "search"
            Label3.Visible = True
            txtFind.Visible = True
            txtFind.SetFocus
        
    End Select
End Sub

'this function will check if the textbox is empty or not
Function complete()
    If txtName = "" Or txtDescription = "" Then
        complete = False
    Else
        complete = True
    End If
End Function

'function to clear the textbox
Sub textclear()
    txtName = ""
    txtDescription = ""
End Sub

'procedure that display records in the list view
Sub loadlsvCategory()
    Dim X As Integer
    
    rs.Open "SELECT * FROM tblCategory", cn, adOpenDynamic, adLockOptimistic
    
    lsvCategory.ListItems.Clear
    
    While Not rs.EOF
      Set lst = lsvCategory.ListItems.Add(, , rs(0))
    For X = 1 To 2
       lst.SubItems(X) = rs(X)
    Next X
    rs.MoveNext
Wend
    Set rs = Nothing
End Sub

'procedure that filters records from the list view
Private Sub txtFind_Change()
    Dim X
    Dim strSQL As String
    strSQL = "SELECT * FROM tblCategory "
    strSQL = strSQL & "WHERE sCategoryName LIKE'" & txtFind.Text & "%'"
    
    rs.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    lsvCategory.ListItems.Clear
    With rs
        Do While Not rs.EOF
            Set X = lsvCategory.ListItems.Add(, , !icategoryID)
                X.SubItems(1) = !sCategoryName
                X.SubItems(2) = !sNotes
        .MoveNext
        Loop
    End With
    Set rs = Nothing
End Sub

'codes will be perform when you press the arrow down in txtfind
Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lsvCategory.ListItems.Count <> 0 Then
            lsvCategory.SetFocus
        End If
    End If
End Sub
