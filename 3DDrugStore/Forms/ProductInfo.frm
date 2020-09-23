VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmProductInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Product Info"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6645
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ProductInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   6645
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5730
      Left            =   -15
      TabIndex        =   0
      Top             =   0
      Width           =   6705
      _ExtentX        =   11827
      _ExtentY        =   10107
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Product List"
      TabPicture(0)   =   "ProductInfo.frx":038A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lsvProduct"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtFindProductList"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Re - Order Info"
      TabPicture(1)   =   "ProductInfo.frx":03A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text1"
      Tab(1).Control(1)=   "txtFindProductReOrderLevel"
      Tab(1).Control(2)=   "lsvReOrder"
      Tab(1).Control(3)=   "Label2"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Expiration Info"
      TabPicture(2)   =   "ProductInfo.frx":03C2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtFindProductExpiration"
      Tab(2).Control(1)=   "lsvExpiration"
      Tab(2).Control(2)=   "Label4"
      Tab(2).Control(3)=   "Label3"
      Tab(2).ControlCount=   4
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   -69480
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox txtFindProductExpiration 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -73545
         TabIndex        =   9
         Top             =   525
         Width           =   2775
      End
      Begin VB.TextBox txtFindProductReOrderLevel 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -73545
         TabIndex        =   6
         Top             =   525
         Width           =   2775
      End
      Begin VB.TextBox txtFindProductList 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1455
         TabIndex        =   3
         Top             =   525
         Width           =   2775
      End
      Begin MSComctlLib.ListView lsvProduct 
         Height          =   4365
         Left            =   90
         TabIndex        =   1
         Top             =   900
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   7699
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
            Text            =   "Product Name"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Category"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Quantity"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Supplier Cost"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Selling Price"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Date Delivered"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Supplier"
            Object.Width           =   5292
         EndProperty
      End
      Begin MSComctlLib.ListView lsvReOrder 
         Height          =   4365
         Left            =   -74910
         TabIndex        =   4
         Top             =   900
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   7699
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
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
            Text            =   "Product Name"
            Object.Width           =   4358
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Qty On Hand"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Re - Order Level"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Status"
            Object.Width           =   4304
         EndProperty
      End
      Begin MSComctlLib.ListView lsvExpiration 
         Height          =   4365
         Left            =   -74910
         TabIndex        =   7
         Top             =   900
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   7699
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "id"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Product Name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Date Delivered"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Warranty Validity"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Expiration Date"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Status"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   -74880
         TabIndex        =   10
         Top             =   5400
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Product Name:"
         Height          =   195
         Left            =   -74910
         TabIndex        =   8
         Top             =   570
         Width           =   1275
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Product Name:"
         Height          =   195
         Left            =   -74910
         TabIndex        =   5
         Top             =   570
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Product Name:"
         Height          =   195
         Left            =   90
         TabIndex        =   2
         Top             =   570
         Width           =   1275
      End
   End
End
Attribute VB_Name = "frmProductInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    Call forminit
    Dim strSQLMonitor As String
    Dim rsMonitor As New ADODB.Recordset
    
    strSQLMonitor = "SELECT * FROM tblProduct"
    
    rsMonitor.Open strSQLMonitor, cn, adOpenDynamic, adLockOptimistic
    
    With rsMonitor
        .MoveFirst
        Do While Not rsMonitor.EOF
            If !dtexpiration <= Date Then
                !sStatus = "EXPIRED"
            Else
                !sStatus = "OK"
            End If
        .Update
        .MoveNext
        Loop
    End With
End Sub

Sub forminit()
    Call loadlsvproduct
    Call loadlsvreorder
    Call loadlsvExpiration
    SSTab1.Tab = 0
    txtFindProductList.SetFocus
End Sub

Sub loadlsvproduct()
    Dim X
    Dim strSQL As String
    Dim rsproduct As New ADODB.Recordset
    
    strSQL = "SELECT * FROM qryproducts"
    
    rsproduct.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    lsvProduct.ListItems.Clear
    If rsproduct.EOF = True Then
        lsvProduct.ListItems.Clear
    Else
        With rsproduct
            Do While Not rsproduct.EOF
                Set X = lsvProduct.ListItems.Add(, , !iproductID)
                X.SubItems(1) = !sproductname
                X.SubItems(2) = !sCategoryName
                X.SubItems(3) = !iQuantity & " " & !sMeasurementName
                X.SubItems(4) = Format(!cCost, "P ###,###,###.00")
                X.SubItems(5) = Format(!cSellingPrice, "P ###,###,###.00")
                X.SubItems(6) = Format(!dtDatedelivered, "MMMM DD, YYYY")
                X.SubItems(7) = !sSupplierName
                .MoveNext
            Loop
        End With
        Set rsproduct = Nothing
    End If
End Sub

Sub loadlsvreorder()
    Dim X
    Dim strSQL As String
    Dim rsReOrder As New ADODB.Recordset
    Dim row
    
    strSQL = "SELECT * FROM qryproducts"
    
    rsReOrder.Open strSQL, cn, adOpenDynamic, adLockOptimistic

    lsvReOrder.ListItems.Clear
    With rsReOrder
'        If !iQuantity = !iReOrderLevel Then
'            lsvReOrder.ForeColor = &HFF&
'            lsvDelivery.ListItems.Item (row)
'        Else
'            lsvReOrder.ForeColor = &H80000012
'        End If
        Do While Not rsReOrder.EOF
                Set X = lsvReOrder.ListItems.Add(, , !iproductID)
                X.SubItems(1) = !sproductname
                X.SubItems(2) = !iQuantity & " " & !sMeasurementName
                X.SubItems(3) = !iReOrderLevel
                X.SubItems(4) = !sStatus1
            
                .MoveNext
        Loop
    End With
        Set rsReOrder = Nothing
End Sub

Sub loadlsvExpiration()
    Dim X
    Dim strSQL As String
    Dim rsExpiration As New ADODB.Recordset
    
    strSQL = "SELECT * FROM qryproducts"
    
    rsExpiration.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    lsvExpiration.ListItems.Clear
    If rsExpiration.EOF = True Then
        lsvExpiration.ListItems.Clear
    Else
        With rsExpiration
            Do While Not rsExpiration.EOF
                Set X = lsvExpiration.ListItems.Add(, , !iproductID)
                X.SubItems(1) = !sproductname
                X.SubItems(2) = Format(!dtDatedelivered, "MMMM DD, YYYY")
                X.SubItems(3) = Format(!dtdateentered, "MMMM DD, YYYY")
                X.SubItems(4) = !dtexpiration
                X.SubItems(5) = !sStatus
                .MoveNext
            Loop
        End With
        Set rsreexpiration = Nothing
    End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 0 Then
        txtFindProductList.SetFocus
    ElseIf SSTab1.Tab = 1 Then
        txtFindProductReOrderLevel.SetFocus
        Dim strSQLReOrder As String
        Dim rsReOrder As New ADODB.Recordset
        Dim rsReOrder1 As New ADODB.Recordset
        Dim rsReOrder2 As New ADODB.Recordset
        
        strSQLReOrder = "SELECT * FROM qryProducts"
        rsReOrder.Open strSQLReOrder, cn, adOpenDynamic, adLockOptimistic
        
        With rsReOrder
            Do While Not rsReOrder.EOF
                If !iQuantity = !iReOrderLevel Then
                    MsgBox "Check Re Order Level " & !sproductname, vbInformation + vbOKOnly, "Re Order Check up"
                    Text1 = !iproductID
                    rsReOrder1.Open "SELECT * FROM tblProduct WHERE iProductID=" & Text1, cn, adOpenDynamic, adLockOptimistic
                        With rsReOrder1
                            !sStatus1 = "Need to be order"
                            .Update
                        End With
                            Set rsReOrder1 = Nothing
                End If
                .MoveNext
            Loop
        End With
        Set rsReOrder = Nothing
        Call loadlsvreorder
    ElseIf SSTab1.Tab = 2 Then
        txtFindProductExpiration.SetFocus
        Dim strSQLExpiration As String
        Dim rsExpiration As New ADODB.Recordset
        Call loadlsvExpiration
        strSQLExpiration = "SELECT * FROM qryProducts"
        rsExpiration.Open strSQLExpiration, cn, adOpenDynamic, adLockOptimistic
        With rsExpiration
        .MoveFirst
            Do While Not rsExpiration.EOF
                If !dtexpiration <= Date Then
                    dummyexpire = dummyexpires + 1
                    MsgBox "Check the expiration of " & !sproductname, vbOKOnly + vbInformation, "Expiration Check up"
                End If
                .MoveNext
            Loop
        End With

        Set rsExpiration = Nothing
        txtFindProductExpiration.SetFocus
    End If
End Sub

Private Sub txtFindProductExpiration_Change()
    Dim X
    Dim strSQL As String
    strSQL = "SELECT * FROM qryProducts "
    strSQL = strSQL & "WHERE sProductName LIKE'" & txtFindProductExpiration.Text & "%'"
    
    rs1.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    lsvExpiration.ListItems.Clear
    With rs1
        Do While Not rs1.EOF
          Set X = lsvExpiration.ListItems.Add(, , !iproductID)
                X.SubItems(1) = !sproductname
                X.SubItems(2) = Format(!dtDatedelivered, "MMMM DD, YYYY")
                X.SubItems(3) = Format(!dtdateentered, "MMMM DD, YYYY")
                X.SubItems(4) = Format(!dtexpiration, "MMMM DD, YYYY")
        .MoveNext
        Loop
    End With
    Set rs1 = Nothing
End Sub

Private Sub txtFindProductExpiration_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyDown Then
        If lsvExpiration.ListItems.Count <> 0 Then
            lsvExpiration.SetFocus
        End If
    End If
End Sub

Private Sub txtFindProductList_Change()
    Dim X
    Dim strSQL As String
    strSQL = "SELECT * FROM qryProducts "
    strSQL = strSQL & "WHERE sProductName LIKE'" & txtFindProductList.Text & "%'"
    
    rs2.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    lsvProduct.ListItems.Clear
    With rs2
        Do While Not rs2.EOF
           Set X = lsvProduct.ListItems.Add(, , !iproductID)
            X.SubItems(1) = !sproductname
            X.SubItems(2) = !sCategoryName
            X.SubItems(3) = !iQuantity & " " & !sMeasurementName
            X.SubItems(4) = Format(!cCost, "P ###,###,###.00")
            X.SubItems(5) = Format(!cSellingPrice, "P ###,###,###.00")
            X.SubItems(6) = Format(!dtDatedelivered, "MMMM DD, YYYY")
            X.SubItems(7) = !sSupplierName
        .MoveNext
        Loop
    End With
    Set rs2 = Nothing
End Sub

Private Sub txtFindProductList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lsvProduct.ListItems.Count <> 0 Then
            lsvProduct.SetFocus
        End If
    End If
End Sub

Private Sub txtFindProductReOrderLevel_Change()
    Dim X
    Dim strSQL As String
    strSQL = "SELECT * FROM qryProducts "
    strSQL = strSQL & "WHERE sProductName LIKE'" & txtFindProductReOrderLevel.Text & "%'"
    
    rs1.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    lsvReOrder.ListItems.Clear
    With rs1
        Do While Not rs1.EOF
          Set X = lsvReOrder.ListItems.Add(, , !iproductID)
            X.SubItems(1) = !sproductname
            X.SubItems(2) = !iQuantity & " " & !sMeasurementName
            X.SubItems(3) = !iReOrderLevel
        .MoveNext
        Loop
    End With
    Set rs1 = Nothing
End Sub

Private Sub txtFindProductReOrderLevel_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lsvReOrder.ListItems.Count <> 0 Then
            lsvReOrder.SetFocus
        End If
    End If
End Sub
