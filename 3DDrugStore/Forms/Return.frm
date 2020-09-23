VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmReturn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Return Info"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7650
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Return.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   7650
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtProductQtyDummy 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6840
      TabIndex        =   39
      Top             =   480
      Visible         =   0   'False
      Width           =   345
   End
   Begin MSComctlLib.ListView lsvProducts 
      Height          =   945
      Left            =   1470
      TabIndex        =   32
      Top             =   420
      Visible         =   0   'False
      Width           =   4785
      _ExtentX        =   8440
      _ExtentY        =   1667
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "id"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   8290
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Return Product"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5955
      Picture         =   "Return.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   4155
      Width           =   1620
   End
   Begin VB.TextBox txtQtyReturn 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5640
      TabIndex        =   3
      Top             =   525
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1470
      TabIndex        =   0
      Top             =   150
      Width           =   4785
   End
   Begin VB.TextBox txtSupplierDummy 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5925
      TabIndex        =   35
      Top             =   150
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.TextBox txtMeasurementDummy 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5505
      TabIndex        =   34
      Top             =   150
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.TextBox txtCategoryDummy 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6855
      TabIndex        =   33
      Top             =   150
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.TextBox txtDeliveryDummy 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4230
      TabIndex        =   31
      Top             =   150
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.TextBox txtProductDummy 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4650
      TabIndex        =   30
      Top             =   150
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Frame Frame1 
      Height          =   3225
      Left            =   90
      TabIndex        =   5
      Top             =   870
      Width           =   7485
      Begin VB.Frame Frame3 
         Height          =   90
         Left            =   135
         TabIndex        =   15
         Top             =   2340
         Width           =   7200
      End
      Begin VB.TextBox txtTotalCost 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1935
         Width           =   1560
      End
      Begin VB.TextBox txtSupplierCost 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1575
         Width           =   1560
      End
      Begin VB.Frame Frame2 
         Height          =   90
         Left            =   150
         TabIndex        =   12
         Top             =   975
         Width           =   7200
      End
      Begin VB.TextBox txtOnHand 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1200
         Width           =   555
      End
      Begin VB.TextBox txtNotes 
         Appearance      =   0  'Flat
         Height          =   555
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   2550
         Width           =   4500
      End
      Begin VB.TextBox txtSupplier 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   9
         Top             =   585
         Width           =   4230
      End
      Begin VB.TextBox txtCategory 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Top             =   210
         Width           =   1860
      End
      Begin VB.TextBox txtWarranty 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4830
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1935
         Width           =   2205
      End
      Begin VB.TextBox txtReOrderLevel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4830
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1575
         Width           =   555
      End
      Begin VB.TextBox txtSellingPrice 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4830
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1200
         Width           =   1560
      End
      Begin VB.TextBox txtdatedelivered 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1455
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   1935
         Visible         =   0   'False
         Width           =   1560
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Total Cost:"
         Height          =   195
         Left            =   405
         TabIndex        =   24
         Top             =   1980
         Width           =   945
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Supplier Cost:"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   1620
         Width           =   1230
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "On - Hand:"
         Height          =   195
         Left            =   405
         TabIndex        =   22
         Top             =   1245
         Width           =   945
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Supplier:"
         Height          =   195
         Left            =   570
         TabIndex        =   21
         Top             =   630
         Width           =   780
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Notes:"
         Height          =   195
         Left            =   795
         TabIndex        =   20
         Top             =   2550
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Category:"
         Height          =   195
         Left            =   480
         TabIndex        =   19
         Top             =   255
         Width           =   870
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Warranty:"
         Height          =   195
         Left            =   3870
         TabIndex        =   18
         Top             =   1980
         Width           =   870
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Re - Order Level:"
         Height          =   195
         Left            =   3240
         TabIndex        =   17
         Top             =   1620
         Width           =   1500
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Selling Price:"
         Height          =   195
         Left            =   3615
         TabIndex        =   16
         Top             =   1245
         Width           =   1125
      End
   End
   Begin VB.TextBox txtReturnDate 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1470
      TabIndex        =   1
      Top             =   525
      Width           =   2010
   End
   Begin VB.TextBox txtFind 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   720
      TabIndex        =   4
      Top             =   4710
      Width           =   2535
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   300
      Left            =   3540
      TabIndex        =   25
      Top             =   540
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   529
      _Version        =   393216
      Format          =   20578305
      CurrentDate     =   38753
   End
   Begin MSComctlLib.ListView lsvReturn 
      Height          =   1320
      Left            =   60
      TabIndex        =   26
      Top             =   5040
      Width           =   7485
      _ExtentX        =   13203
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
      NumItems        =   12
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
         Object.Width           =   1764
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
         Text            =   "Price"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Delivered"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Entered"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Warranty"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Expiration"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Returned"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Notes"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Supplier"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Qty to Return:"
      Height          =   195
      Left            =   4305
      TabIndex        =   37
      Top             =   570
      Width           =   1230
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Return Date:"
      Height          =   195
      Left            =   240
      TabIndex        =   29
      Top             =   540
      Width           =   1110
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Name:"
      Height          =   195
      Left            =   75
      TabIndex        =   27
      Top             =   4755
      Width           =   570
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Product:"
      Height          =   195
      Left            =   630
      TabIndex        =   28
      Top             =   180
      Width           =   720
   End
End
Attribute VB_Name = "frmReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    Call forminit
End Sub

Sub forminit()
    Call CenterForm(frmReturn)
    Call loadlsvProducts
    'Call loadlsvReturn
End Sub

Private Sub Command1_Click()
    Dim strSQLAdd As String
    Dim rsAdd As New Recordset
    Dim strSQL1 As String
    Dim strSQL2 As String
    
    If txtReturnDate <> "" Then
        
        strSQLAdd = "SELECT * FROM  tblreturn"
        rsAdd.Open strSQLAdd, cn, adOpenDynamic, adLockOptimistic
        
        With rsAdd
            .AddNew
            !iproductID = txtProductDummy
            !dtReturnedDate = txtReturnDate
            !iQtyReturn = txtQtyReturn
            !sNotes = txtNotes
            .Update
        End With
            MsgBox "Product(s) has been returned", vbOKOnly, "Product Return"
            
            'Call loadlsvReturn
            txtReturnDate.SetFocus
            Set rsAdd = Nothing
    Else
        MsgBox "Please enter return date", vbOKOnly, "Return Product"
        txtReturnDate.SetFocus
    End If
    
    strSQL1 = "SELECT * FROM tblDelivery"
    strSQL1 = strSQL1 & " WHERE iDeliveryID=" & txtDeliveryDummy
    
    rs2.Open strSQL1, cn, adOpenDynamic, adLockOptimistic
            
    With rs2
        qtyreturndummy = !iQuantity - Val(txtProductQtyDummy.Text)
        !iQuantity = qtyreturndummy
        .Update
    End With
    Set rs2 = Nothing
    Call textclear
    'Call loadlsvReturn
End Sub

Private Sub DTPicker1_Change()
        txtReturnDate.Text = Format(DTPicker1.Value, "MMMM DD, YYYY")
End Sub

Sub loadlsvProducts()
    Dim X
    Dim strSQL As String
    
    strSQL = "SELECT * FROM qryProducts"
    
    rs.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    lsvProducts.ListItems.Clear
    If rs.EOF = True Then
        lsvProducts.ListItems.Clear
    Else
        With rs
            Do While Not rs.EOF
                Set X = lsvProducts.ListItems.Add(, , !iproductID)
                    X.SubItems(1) = !sproductname
                    .MoveNext
            Loop
        End With
        Set rs = Nothing
    End If
End Sub
'Sub loadlsvReturn()
'    Dim x
'    Dim strSQL As String
'
'    strSQL = "SELECT * FROM qryproducts"
'
'    rs.Open strSQL, cn, adOpenDynamic, adLockOptimistic
'    lsvReturn.ListItems.Clear
'    If rs.EOF = True Then
'        lsvReturn.ListItems.Clear
'    Else
'        With rs
'            Do While Not rs.EOF
'                Set x = lsvReturn.ListItems.Add(, , !iproductID)
'                    x.SubItems(1) = !sProductName
'                    x.SubItems(2) = !sCategoryName
'                    x.SubItems(3) = !iQuantity
'                    x.SubItems(4) = !cSellingPrice
'                    x.SubItems(5) = !dtDateDelivered
'                    x.SubItems(6) = !dtDateEntered
'                    x.SubItems(7) = !dtWarranty
'                    x.SubItems(8) = !dtExpiration
'                    'x.SubItems(9) = !dtReturnedDate
'                    x.SubItems(10) = !snotes
'                    x.SubItems(11) = !sSupplierName
'                    .MoveNext
'            Loop
'        End With
'        Set rs = Nothing
'    End If
'End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Call textclear
        txtReturnDate.SetFocus
        Command1.Enabled = False
    End If
End Sub

Private Sub lsvProducts_Click()
    Dim X As Integer
    Dim strSQL As String
    Dim dummy
    Dim row
    
    lsvProducts.Visible = False
    txtReturnDate.SetFocus
    Command1.Enabled = True
    row = lsvProducts.SelectedItem.Index
    dummy = lsvProducts.ListItems.Item(row).Text
    
    strSQL = "SELECT * FROM qryProducts "
    strSQL = strSQL & "WHERE iProductId=" & dummy
    
    rs3.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    
    With rs3
        txtProductDummy = dummy
        txtDeliveryDummy = !iDeliveryID
        txtCategoryDummy = !icategoryID
        txtMeasurementDummy = !iMeasurementID
        txtSupplierDummy = !iSupplierID
        Text2 = !sproductname
        txtCategory = !sCategoryName
        txtSupplier = !sSupplierName
        txtOnHand = !iQuantity
        txtSupplierCost = !cCost
        txtTotalCost = !cTotalCost
        txtdatepurchased = Format(!dtDatedelivered, "MMMM DD, YYYY")
        txtSellingPrice = Format(!cSellingPrice, "P ###,###,###.00")
        txtReOrderLevel = !iReOrderLevel
        txtdatedelivered = !dtDatedelivered
        txtWarranty = Format(!dtWarranty)
    End With
        
        Set rs3 = Nothing
End Sub

Private Sub lsvProducts_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        lsvProducts_Click
    End If
End Sub

Private Sub Text2_Change()
    Dim X
    Dim strSQL As String
    strSQL = "SELECT * FROM qryProducts "
    strSQL = strSQL & "WHERE sProductName LIKE'" & Text2.Text & "%'"
    
    rs2.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    lsvProducts.ListItems.Clear
    With rs2
        Do While Not rs2.EOF
            Set X = lsvProducts.ListItems.Add(, , !iproductID)
                X.SubItems(1) = !sproductname
        .MoveNext
        Loop
    End With
    Set rs2 = Nothing
End Sub

Private Sub Text2_GotFocus()
    lsvProducts.Visible = True
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lsvProducts.ListItems.Count <> 0 Then
            lsvProducts.SetFocus
        End If
    End If
End Sub

Sub textclear()
    txtReturnDate = ""
    Text2 = ""
    txtCategory = ""
    txtSupplier = ""
    txtOnHand = ""
    txtSupplierCost = ""
    txtTotalCost = ""
    txtSellingPrice = ""
    txtReOrderLevel = ""
    txtWarranty = ""
    txtNotes = ""
    txtFind = ""
    txtQtyReturn = ""
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        lsvProducts_Click
    End If
End Sub

Private Sub txtQtyReturn_Change()
    txtProductQtyDummy.Text = txtQtyReturn.Text
End Sub
