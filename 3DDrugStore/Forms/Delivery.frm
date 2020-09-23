VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmDelivery 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Product Delivery"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6705
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Delivery.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   6705
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFind 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   780
      TabIndex        =   21
      Top             =   4905
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
      Height          =   4140
      Left            =   120
      TabIndex        =   11
      Top             =   645
      Width           =   6480
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   300
         Left            =   3600
         TabIndex        =   2
         Top             =   615
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   529
         _Version        =   393216
         Format          =   59047937
         CurrentDate     =   38753
      End
      Begin VB.TextBox txtDateDelivery 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Top             =   615
         Width           =   2010
      End
      Begin VB.ComboBox cboMeasurement 
         Height          =   315
         Left            =   2325
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2085
         Width           =   1785
      End
      Begin VB.Frame Frame3 
         Height          =   90
         Left            =   135
         TabIndex        =   20
         Top             =   3225
         Width           =   6180
      End
      Begin VB.TextBox txtTotalCost 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   2850
         Width           =   1560
      End
      Begin VB.TextBox txtSupplierCost 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   7
         Top             =   2490
         Width           =   1560
      End
      Begin VB.Frame Frame2 
         Height          =   90
         Left            =   135
         TabIndex        =   17
         Top             =   1860
         Width           =   6180
      End
      Begin VB.TextBox txtQuantity 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   5
         Top             =   2115
         Width           =   555
      End
      Begin VB.ComboBox cboSupplier 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1395
         Width           =   4500
      End
      Begin VB.ComboBox cboCategory 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   990
         Width           =   1860
      End
      Begin VB.TextBox txtProductName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   0
         Top             =   240
         Width           =   4500
      End
      Begin VB.TextBox txtNotes 
         Appearance      =   0  'Flat
         Height          =   555
         Left            =   1560
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   3450
         Width           =   4500
      End
      Begin VB.TextBox txtCategoryDummy 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3045
         TabIndex        =   24
         Top             =   1005
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtSupplierDummy 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5655
         TabIndex        =   25
         Top             =   1425
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtMeasurementDummy 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3720
         TabIndex        =   26
         Top             =   2115
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtDeliveryDummy 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3195
         TabIndex        =   28
         Top             =   615
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Date Recieved:"
         Height          =   195
         Left            =   165
         TabIndex        =   27
         Top             =   645
         Width           =   1320
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Total Cost:"
         Height          =   195
         Left            =   480
         TabIndex        =   19
         Top             =   2895
         Width           =   945
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Supplier Cost:"
         Height          =   195
         Left            =   195
         TabIndex        =   18
         Top             =   2535
         Width           =   1230
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Quantity:"
         Height          =   195
         Left            =   630
         TabIndex        =   16
         Top             =   2160
         Width           =   795
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Supplier:"
         Height          =   195
         Left            =   645
         TabIndex        =   15
         Top             =   1455
         Width           =   780
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Category:"
         Height          =   195
         Left            =   555
         TabIndex        =   14
         Top             =   1035
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Product Name:"
         Height          =   195
         Left            =   150
         TabIndex        =   13
         Top             =   285
         Width           =   1275
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Notes:"
         Height          =   195
         Left            =   870
         TabIndex        =   12
         Top             =   3450
         Width           =   555
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   330
      Left            =   4875
      TabIndex        =   10
      Top             =   180
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
            Object.ToolTipText     =   "Add New Product"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "update"
            Object.ToolTipText     =   "Update Existing Product"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "delete"
            Object.ToolTipText     =   "Delete Existing Product"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "search"
            Object.ToolTipText     =   "Find Product"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5925
      Top             =   675
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
            Picture         =   "Delivery.frx":038A
            Key             =   "add"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Delivery.frx":0724
            Key             =   "edit"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Delivery.frx":0ABE
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Delivery.frx":0E58
            Key             =   "search"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lsvDelivery 
      Height          =   1320
      Left            =   135
      TabIndex        =   22
      Top             =   5280
      Width           =   6480
      _ExtentX        =   11430
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
      NumItems        =   8
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
         Text            =   "Category"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Quantity"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Cost"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Total Cost"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Supplier"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Delivery Date"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Name:"
      Height          =   195
      Left            =   135
      TabIndex        =   23
      Top             =   4950
      Visible         =   0   'False
      Width           =   570
   End
End
Attribute VB_Name = "frmDelivery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public dummyTotalCost
Dim FindItem As Boolean

Private Sub Form_Activate()
    Call forminit
End Sub

Sub forminit()
    DTPicker1.Value = Date
    Call loadcboCategory
    Call loadcboSupplier
    Call loadcboMeasurement
    Call loadlsvDelivery
    Toolbar1.Buttons(4).Enabled = False
    Toolbar1.Buttons(5).Enabled = False
End Sub

Private Sub cboCategory_Click()
    On Error Resume Next
    Dim sindex
    Dim strSQL As String
    sindex = cboCategory.ItemData(cboCategory.ListIndex)
    
    strSQL = "SELECT * FROM tblCategory"
    strSQL = strSQL & " WHERE iCategoryID=" & sindex
    
    rs.Open strSQL, cn, adOpenDynamic, adLockOptimistic
     
    With rs
        txtCategoryDummy = !icategoryID
    End With
    Set rs = Nothing
End Sub

Private Sub cboMeasurement_Click()
    On Error Resume Next
    Dim sindex
    Dim strSQL As String
    sindex = cboMeasurement.ItemData(cboMeasurement.ListIndex)
    
    strSQL = "SELECT * FROM tblMeasurement"
    strSQL = strSQL & " WHERE iMeasurementID=" & sindex
    
    rs.Open strSQL, cn, adOpenDynamic, adLockOptimistic
     
    With rs
        txtMeasurementDummy = !iMeasurementID
    End With
    Set rs = Nothing
End Sub

Private Sub cboSupplier_Click()
    On Error Resume Next
    Dim sindex
    Dim strSQL As String
    sindex = cboSupplier.ItemData(cboSupplier.ListIndex)
    
    strSQL = "SELECT * FROM tblSupplier"
    strSQL = strSQL & " WHERE iSupplierID=" & sindex
    
    rs.Open strSQL, cn, adOpenDynamic, adLockOptimistic
     
    With rs
        txtSupplierDummy = !iSupplierID
    End With
    Set rs = Nothing
End Sub

Private Sub DTPicker1_Change()
    txtDateDelivery.Text = Format(DTPicker1.Value, "MMMM DD, YYYY")
End Sub

Sub loadcboCategory()
    Dim strSQL As String
    Dim rscategory As New ADODB.Recordset
    
    strSQL = "SELECT * FROM tblCategory"
    
    rscategory.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    With rscategory
        Do While Not rscategory.EOF
            cboCategory.AddItem !sCategoryName
            cboCategory.ItemData(cboCategory.NewIndex) = CLng(!icategoryID)
            .MoveNext
        Loop
    End With
    Set rscategory = Nothing
End Sub

Sub loadcboSupplier()
    Dim strSQL As String
    
    strSQL = "SELECT * FROM tblSupplier"
    
    rs.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    With rs
        Do While Not rs.EOF
            cboSupplier.AddItem !sSupplierName
            cboSupplier.ItemData(cboSupplier.NewIndex) = CLng(!iSupplierID)
            .MoveNext
        Loop
    End With
    Set rs = Nothing
End Sub

Sub loadcboMeasurement()
    Dim strSQL As String
    
    strSQL = "SELECT * FROM tblMeasurement"
    
    rs.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    With rs
        Do While Not rs.EOF
            cboMeasurement.AddItem !sMeasurementName
            cboMeasurement.ItemData(cboMeasurement.NewIndex) = CLng(!iMeasurementID)
            .MoveNext
        Loop
    End With
    Set rs = Nothing
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Call textclear
        txtProductName.SetFocus
        Toolbar1.Buttons(2).Enabled = True
        Toolbar1.Buttons(4).Enabled = False
        Toolbar1.Buttons(5).Enabled = False
    End If
End Sub

Private Sub lsvDelivery_Click()
    Dim X As Integer
    Dim strSQL As String
    Dim dummy
    Dim row
    Toolbar1.Buttons(2).Enabled = False
    Toolbar1.Buttons(4).Enabled = True
    Toolbar1.Buttons(5).Enabled = True
    
    row = lsvDelivery.SelectedItem.Index
    dummy = lsvDelivery.ListItems.Item(row).Text
    
    strSQL = "SELECT * FROM qryDelivery "
    strSQL = strSQL & "WHERE iDeliveryId=" & dummy
    
    rs.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    
    With rs
        txtDeliveryDummy = dummy
        txtDateDelivery = !dtDatedelivered
        txtProductName = !sproductname
        cboCategory.ListIndex = ListFindItem(cboCategory, CLng(!icategoryID))
        cboSupplier.ListIndex = ListFindItem(cboSupplier, CLng(!iSupplierID))
        txtQuantity = !iQuantity
        cboMeasurement.ListIndex = ListFindItem(cboMeasurement, CLng(!iMeasurementID))
        txtSupplierCost = !cCost
        txtTotalCost = !cTotalCost
        txtNotes = !sNotes
    End With
        txtProductName.SetFocus
    Set rs = Nothing
End Sub

Private Sub lsvDelivery_KeyPress(KeyAscii As Integer)
    lsvDelivery_Click
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "add"
            Dim strSQLAdd As String
            Dim rsAdd As New Recordset
            Call ListCheck
            If FindItem = False Then
                If complete = True Then
                    strSQLAdd = "SELECT * FROM  tblDelivery"
                    rsAdd.Open strSQLAdd, cn, adOpenDynamic, adLockOptimistic
                    With rsAdd
                        .AddNew
                        !dtDatedelivered = txtDateDelivery
                        !sproductname = txtProductName
                        !iSupplierID = cboSupplier.ItemData(cboSupplier.ListIndex)
                        !icategoryID = cboCategory.ItemData(cboCategory.ListIndex)
                        !iQuantity = txtQuantity
                        !iMeasurementID = cboMeasurement.ItemData(cboMeasurement.ListIndex)
                        !cCost = txtSupplierCost
                        !cTotalCost = dummyTotalCost
                        !sNotes = txtNotes
                        .Update
                    End With
                        MsgBox "New record added to the database", vbOKOnly, "Product Delivery"
                        Call textclear
                        Call loadlsvDelivery
                        txtProductName.SetFocus
                        Set rsAdd = Nothing
                Else
                    If txtProductName = "" Then
                        MsgBox "Please enter product name", vbOKOnly, "Product Delivery"
                        txtProductName.SetFocus
                        Exit Sub
                    ElseIf txtDateDelivery = "" Then
                        MsgBox "Please enter delivery date", vbOKOnly, "Product Delivery"
                        txtDateDelivery.SetFocus
                        Exit Sub
                    ElseIf cboSupplier = "" Then
                        MsgBox "Please select product supplier", vbOKOnly, "Product Delivery"
                        cboSupplier.SetFocus
                        Exit Sub
                    ElseIf cboCategory = "" Then
                        MsgBox "Please select product category", vbOKOnly, "Product Delivery"
                        cboCategory.SetFocus
                        Exit Sub
                    ElseIf txtQuantity = "" Then
                        MsgBox "Please enter product quantity", vbOKOnly, "Product Delivery"
                        txtQuantity.SetFocus
                        Exit Sub
                    ElseIf cboMeasurement = "" Then
                        MsgBox "Please select measurement", vbOKOnly, "Product Delivery"
                    ElseIf txtSupplierCost = "" Then
                        MsgBox "Please enter supplier cost", vbOKOnly, "Product Delivery"
                        txtSupplierCost.SetFocus
                        Exit Sub
                    End If
                End If
            End If
            
        Case "update"
            Dim strSQLEdit As String
            Dim rsEdit As New Recordset
            
            If complete = True Then
                strSQLEdit = "SELECT * FROM  tblDelivery"
                strSQLEdit = strSQLEdit & " WHERE iDeliveryID=" & txtDeliveryDummy
                
                rsEdit.Open strSQLEdit, cn, adOpenDynamic, adLockOptimistic
                
                With rsEdit
                    !dtDatedelivered = txtDateDelivery
                    !sproductname = txtProductName
                    !iSupplierID = cboSupplier.ItemData(cboSupplier.ListIndex)
                    !icategoryID = cboCategory.ItemData(cboCategory.ListIndex)
                    !iQuantity = txtQuantity
                    !iMeasurementID = cboMeasurement.ItemData(cboMeasurement.ListIndex)
                    !cCost = txtSupplierCost
                    !cTotalCost = dummyTotalCost
                    !sNotes = txtNotes
                    .Update
                End With
                    MsgBox "The changes you made was successfully updated", vbOKOnly, "Product Delivery"
                    Call textclear
                    Call loadlsvDelivery
                    txtProductName.SetFocus
                    Set rsEdit = Nothing
            Else
                If txtProductName = "" Then
                    MsgBox "Please enter product name", vbOKOnly, "Product Delivery"
                    txtProductName.SetFocus
                    Exit Sub
                ElseIf txtDateDelivery = "" Then
                    MsgBox "Please enter delivery date", vbOKOnly, "Product Delivery"
                    txtDateDelivery.SetFocus
                    Exit Sub
                ElseIf cboSupplier = "" Then
                    MsgBox "Please select product supplier", vbOKOnly, "Product Delivery"
                    cboSupplier.SetFocus
                    Exit Sub
                ElseIf cboCategory = "" Then
                    MsgBox "Please select product category", vbOKOnly, "Product Delivery"
                    cboCategory.SetFocus
                    Exit Sub
                ElseIf txtQuantity = "" Then
                    MsgBox "Please enter product quantity", vbOKOnly, "Product Delivery"
                    txtQuantity.SetFocus
                    Exit Sub
                ElseIf cboMeasurement = "" Then
                    MsgBox "Please select measurement", vbOKOnly, "Product Delivery"
                ElseIf txtSupplierCost = "" Then
                    MsgBox "Please enter supplier cost", vbOKOnly, "Product Delivery"
                    txtSupplierCost.SetFocus
                    Exit Sub
                End If
            End If
        
        Case "delete"
            Dim strSQLDelete As String
            Dim rsDelete As New ADODB.Recordset
            Dim answerDelete As String
            
            strSQLDelete = "SELECT * FROM tblDelivery "
            strSQLDelete = strSQLDelete & "WHERE iDeliveryID=" & txtDeliveryDummy
            
            answerDelete = MsgBox("Are you sure you want DELETE this record?", vbQuestion + vbYesNo, "Product Delivery")
            
            If answerDelete = vbYes Then
                rsDelete.Open strSQLDelete, cn, adOpenDynamic, adLockOptimistic
                With rsDelete
                    .Delete
                End With
                MsgBox "The record was successfully deleted", , "Product Delivery"
            End If
            Set rsDelete = Nothing
            Call textclear
            txtFind = ""
            txtFind.Visible = True
            Label3.Visible = True
            txtFind.SetFocus
            Call loadlsvDelivery
        
        Case "search"
            Label8.Visible = True
            txtFind.Visible = True
            txtFind.SetFocus
    End Select
End Sub

Sub ListCheck()
    Dim litmfound As ListItem
    Set litmfound = lsvDelivery.FindItem(txtProductName, 1, , 0)

    If litmfound Is Nothing Then
        
        FindItem = False
    Else
        MsgBox "This PRODUCT is already in the list" & vbCrLf _
               & "Input another PRODUCT", vbCritical + vbOKOnly, "Duplicate Item"
        litmfound.EnsureVisible
        litmfound.Selected = True
        FindItem = True
        txtProductName = ""
        txtProductName.SetFocus
    End If
End Sub

Function complete()
    If txtDateDelivery = "" Or txtProductName = "" Or cboCategory = "" Or cboSupplier = "" Or _
        txtQuantity = "" Or cboMeasurement = "" Or txtSupplierCost = "" Or txtTotalCost = "" Then
        complete = False
    Else
        complete = True
    End If
End Function

Sub textclear()
    txtDateDelivery = ""
    txtProductName = ""
    cboCategory.ListIndex = -1
    cboSupplier.ListIndex = -1
    txtQuantity = ""
    cboMeasurement.ListIndex = -1
    txtSupplierCost = ""
    txtTotalCost = ""
    txtNotes = ""
    DTPicker1.Value = Date
End Sub

Private Sub txtFind_Change()
    Dim X
    Dim strSQL As String
    strSQL = "SELECT * FROM qryDelivery "
    strSQL = strSQL & "WHERE sProductName LIKE'" & txtFind.Text & "%'"
    
    rs.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    lsvDelivery.ListItems.Clear
    With rs
        Do While Not rs.EOF
            Set X = lsvDelivery.ListItems.Add(, , !iDeliveryID)
                X.SubItems(1) = !sproductname
                X.SubItems(2) = !sCategoryName
                X.SubItems(3) = !iQuantity & " " & !sMeasurementName
                X.SubItems(4) = !cCost
                X.SubItems(5) = !cTotalCost
                X.SubItems(6) = !sSupplierName
                X.SubItems(7) = Format(!dtDatedelivered, "MMMM DD, YYYY")
        .MoveNext
        Loop
    End With
    Set rs = Nothing
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lsvDelivery.ListItems.Count <> 0 Then
            lsvDelivery.SetFocus
        End If
    End If
End Sub

Private Sub txtSupplierCost_Change()
    txtTotalCost.Text = Val(txtQuantity.Text) * Val(txtSupplierCost.Text)
    dummyTotalCost = Val(txtTotalCost.Text)
    txtTotalCost = Format(txtTotalCost, "P ###,###,###.00")
End Sub

Sub loadlsvDelivery()
    Dim X
    Dim strSQL As String
    
    strSQL = "SELECT * FROM qryDelivery"
    
    rs.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    lsvDelivery.ListItems.Clear
    If rs.EOF = True Then
        lsvDelivery.ListItems.Clear
    Else
        With rs
            Do While Not rs.EOF
                Set X = lsvDelivery.ListItems.Add(, , !iDeliveryID)
                    X.SubItems(1) = !sproductname
                    X.SubItems(2) = !sCategoryName
                    X.SubItems(3) = !iQuantity & " " & !sMeasurementName
                    X.SubItems(4) = !cCost
                    X.SubItems(5) = !cTotalCost
                    X.SubItems(6) = !sSupplierName
                    X.SubItems(7) = Format(!dtDatedelivered, "MMMM DD, YYYY")
                    .MoveNext
            Loop
        End With
    End If
    Set rs = Nothing
End Sub
