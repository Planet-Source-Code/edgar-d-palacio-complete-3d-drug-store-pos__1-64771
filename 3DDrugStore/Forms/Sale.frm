VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSale 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sale"
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11775
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Sale.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   11775
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7080
      TabIndex        =   30
      Top             =   6465
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.TextBox txtFind 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   255
      TabIndex        =   29
      Top             =   270
      Width           =   2520
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7080
      TabIndex        =   28
      Top             =   6990
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.TextBox txtDeliveryDummy 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7215
      TabIndex        =   25
      Top             =   5370
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&New Transaction"
      Height          =   525
      Left            =   9675
      TabIndex        =   24
      Top             =   7830
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2490
      Left            =   135
      TabIndex        =   2
      Top             =   5925
      Width           =   6855
      Begin VB.TextBox txtAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   4200
         TabIndex        =   23
         Text            =   "0"
         Top             =   1050
         Width           =   2520
      End
      Begin VB.Label lblChange 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   4200
         TabIndex        =   7
         Top             =   2055
         Width           =   2520
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   345
         Left            =   4200
         TabIndex        =   6
         Top             =   495
         Width           =   2520
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CHANGE:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   2055
         Width           =   1590
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AMOUNT RECEIVED:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   345
         Left            =   120
         TabIndex        =   4
         Top             =   1065
         Width           =   3255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   345
         Left            =   120
         TabIndex        =   3
         Top             =   495
         Width           =   1155
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5145
      Left            =   7185
      TabIndex        =   1
      Top             =   150
      Width           =   4470
      Begin MSComctlLib.ListView lsvList 
         Height          =   4860
         Left            =   105
         TabIndex        =   17
         Top             =   180
         Width           =   4245
         _ExtentX        =   7488
         _ExtentY        =   8573
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "id"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "NAME"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "QTY"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "PRICE"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "DATE"
            Object.Width           =   2540
         EndProperty
      End
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
      Height          =   5190
      Left            =   135
      TabIndex        =   0
      Top             =   60
      Width           =   6915
      Begin VB.Frame fmeQty 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   2655
         Left            =   765
         TabIndex        =   9
         Top             =   1500
         Visible         =   0   'False
         Width           =   5340
         Begin VB.TextBox txtDate 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3600
            TabIndex        =   26
            Top             =   960
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox txtPrice 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1920
            TabIndex        =   22
            Top             =   930
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox txtPriceDummy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2625
            TabIndex        =   21
            Top             =   930
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox txtProductDummy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   195
            TabIndex        =   20
            Top             =   45
            Width           =   345
         End
         Begin VB.CommandButton cmdBack 
            Height          =   420
            Left            =   3930
            Picture         =   "Sale.frx":038A
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   1935
            Width           =   510
         End
         Begin VB.CommandButton cmdAdd 
            Height          =   420
            Left            =   4605
            Picture         =   "Sale.frx":0714
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   1935
            Width           =   510
         End
         Begin VB.Frame Frame5 
            Height          =   30
            Left            =   210
            TabIndex        =   14
            Top             =   1590
            Width           =   4890
         End
         Begin VB.TextBox txtQty 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1215
            TabIndex        =   13
            Top             =   930
            Width           =   615
         End
         Begin VB.TextBox txtProduct 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1215
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   420
            Width           =   3900
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Product:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   210
            TabIndex        =   11
            Top             =   465
            Width           =   915
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qty:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Left            =   615
            TabIndex        =   10
            Top             =   975
            Width           =   465
         End
      End
      Begin MSComctlLib.ListView lsvProducts 
         Height          =   4455
         Left            =   105
         TabIndex        =   8
         Top             =   615
         Width           =   6705
         _ExtentX        =   11827
         _ExtentY        =   7858
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "idproducts"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "NAME"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "PRICE"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "CATEGORY"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "QTY"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "EXPIRATION"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "SUPPLIER"
            Object.Width           =   4410
         EndProperty
      End
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   345
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   2520
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   11325
      TabIndex        =   19
      Top             =   5415
      Width           =   135
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Count:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   10545
      TabIndex        =   18
      Top             =   5400
      Width           =   720
   End
End
Attribute VB_Name = "frmSale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim counter As Integer
Dim total
Dim totaldummy
Dim qtyreturndummy As Integer
Public FindItem As Boolean

Private Sub Form_Activate()
    Call forminit
End Sub

Sub forminit()
    txtDate = Now
    Call CenterForm(frmSale)
    Call loadlsvproduct
    txtFind.SetFocus
    Call loadlsvList
End Sub

Private Sub cmdSave_Click()
    Dim rsSale As Recordset
    Dim rsSale1 As Recordset
    
    Dim strSQL As String
    Dim strSQL1 As String
    
    Dim i As Integer
    Dim X As Integer
    
    If Val(txtAmount) < Val(Text2) Then
        MsgBox "Invalid amount", vbCritical + vbOKOnly, "3D Drug Store - Sale"
        txtAmount.SetFocus
    Else
        MsgBox "Ready for new transaction", vbInformation, "3D Drug Store POS - Sale"
        
        Set rsSale1 = New Recordset
        strSQL1 = "SELECT * FROM tblSale"
        rsSale1.Open strSQL1, cn, adOpenDynamic, adLockOptimistic
        
        With rsSale1
            For i = 1 To lsvList.ListItems.Count
                .AddNew
                !iSaleId = lsvList.ListItems(i)
                !sproductname = lsvList.ListItems(i).SubItems(1)
                !iQty = lsvList.ListItems(i).SubItems(2)
                !cPrice = lsvList.ListItems(i).SubItems(3)
                !dtTransactionDate = lsvList.ListItems(i).SubItems(4)
                .Update
                Set rsSale = Nothing
            Next
        End With
        Load frmSaleReceipt
        frmSaleReceipt.Show 1
        
        Set rsSale = New Recordset
        strSQL = "SELECT * FROM tblSaleTemp"
        rsSale.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    
        With rsSale
            .MoveFirst
            Do While Not .EOF
                .Delete
                .MoveNext
            Loop
            Call loadlsvList
            Text2 = ""
            Text1 = ""
            lblChange = "0.00"
            lblTotal = "0.00"
            txtAmount = ""
            lblChange = ""
        End With
    End If
End Sub

Private Sub cmdAdd_Click()
    Dim strSQL As String
    Dim strSQL1 As String
    fmeQty.Visible = True
    txtQty.SetFocus
    'row = lsvProducts.SelectedItem.Index
    'dummy = lsvProducts.ListItems.Item(row).Text
    
    cmdSave.Visible = True
    fmeQty.Visible = False
    strSQL = "SELECT * FROM tblSaleTemp "
    Call ListCheck
    If FindItem = False Then
        If txtQty <> "" Then
            rs1.Open strSQL, cn, adOpenDynamic, adLockOptimistic
            lsvList.ListItems.Clear
            With rs1
                .AddNew
                !sproductname = txtProduct
                !iQty = txtQty
                !cPrice = txtPriceDummy
                !dtTransactionDate = Format(txtDate, "MMMM DD, YYYY")
                .Update
            End With
            Call loadlsvList
            txtFind.SetFocus
            'Call loadlsvproduct
            'lsvProducts.Refresh
            Set rs1 = Nothing
        Else
            MsgBox "Please enter quantity", vbInformation, "Quantity missing"
            txtQty = ""
            fmeQty.Visible = True
            txtQty.SetFocus
        End If
  
        strSQL1 = "SELECT * FROM tblDelivery"
        strSQL1 = strSQL1 & " WHERE iDeliveryID=" & txtDeliveryDummy
        
        rs2.Open strSQL1, cn, adOpenDynamic, adLockOptimistic
        
        With rs2
            qtyreturndummy = !iQuantity - Val(txtQty.Text)
            !iQuantity = qtyreturndummy
            .Update
        End With
            Call loadlsvproduct
            Set rs2 = Nothing
    End If
End Sub

Sub loadlsvList()
    On Error GoTo err_handler:
    Dim X
    Dim strSQL As String
    
    strSQL = "SELECT * FROM tblSaleTemp"
    
    rs2.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    lsvList.ListItems.Clear
    If rs2.EOF = True Then
        lsvList.ListItems.Clear
    Else
    counter = 0
    total = 0
        With rs2
            Do While Not rs2.EOF
                Set X = lsvList.ListItems.Add(, , !iSaleId)
                X.SubItems(1) = !sproductname
                X.SubItems(2) = !iQty
                X.SubItems(3) = !cPrice
                X.SubItems(4) = !dtTransactionDate
                counter = counter + 1
                total = total + !cPrice
                .MoveNext
            Loop
        End With
        Label10.Caption = counter
        totaldummy = total
        Text2 = total
        lblTotal.Caption = Format(total, "P ###,###,###.00")
        Set rs2 = Nothing
    End If
err_handler:
        Set rs2 = Nothing
End Sub

Private Sub cmdBack_Click()
    fmeQty.Visible = False
    txtFind.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        fmeQty.Visible = False
        txtFind.SetFocus
    End If
End Sub

Private Sub ListView1_DblClick()
    MsgBox "Do you want to remove this item from the list?", vbQuestion, "Delete from List"
End Sub

Private Sub lsvList_DblClick()
    Dim strSQLUpdate As String
    Dim answerUpdate As String
    
    If MsgBox("Are you sure you want to REMOVE  " & lsvList.SelectedItem.SubItems(1), vbYesNo, "Delete Item") = vbYes Then
        cn.Execute "DELETE FROM tblSaleTemp WHERE iSaleID=" & lsvList.SelectedItem.Text
        Call loadlsvList
    
    
        strSQLUpdate = "SELECT * FROM tblDelivery"
        strSQLUpdate = strSQLUpdate & " WHERE iDeliveryID=" & txtDeliveryDummy
        
        rsUpdate.Open strSQLUpdate, cn, adOpenDynamic, adLockOptimistic
        With rsUpdate
            qtyreturndummy = !iQuantity + Val(txtQty.Text)
            !iQuantity = qtyreturndummy
            .Update
        End With
        Call loadlsvproduct
        Set rsUpdate = Nothing
    End If
End Sub

Private Sub lsvProducts_Click()
    
    Dim X As Integer
    Dim strSQL As String
    Dim dummy
    Dim row
    
    fmeQty.Visible = True
    txtProduct = ""
    txtQty = ""
    txtQty.SetFocus
    row = lsvProducts.SelectedItem.Index
    dummy = lsvProducts.ListItems.Item(row).Text
    
    strSQL = "SELECT * FROM qryproducts "
    strSQL = strSQL & "WHERE ideliveryId=" & dummy
    
    rs1.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    
    With rs1
        txtProductDummy = !iproductID
        txtDeliveryDummy = !iDeliveryID
        txtFind = !sproductname
        txtProduct = !sproductname
        txtPrice = !cSellingPrice
    End With
        Set rs1 = Nothing
End Sub

Sub loadlsvproduct()
    On Error GoTo err_handler:
    Dim X
    Dim strSQL As String
    
    strSQL = "SELECT * FROM qryproducts"
    
    rs2.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    lsvProducts.ListItems.Clear
    If rs2.EOF = True Then
        lsvProducts.ListItems.Clear
    Else
        With rs2
            Do While Not rs2.EOF
                Set X = lsvProducts.ListItems.Add(, , !iproductID)
                X.SubItems(1) = !sproductname
                X.SubItems(2) = Format(!cSellingPrice, "P ###,###,###.00")
                X.SubItems(3) = !sCategoryName
                X.SubItems(4) = !iQuantity & " " & !sMeasurementName
                X.SubItems(5) = Format(!dtexpiration)
                X.SubItems(6) = !sSupplierName
                .MoveNext
            Loop
        End With
        Set rs2 = Nothing
err_handler:
        Set rs2 = Nothing
    End If
End Sub

Private Sub lsvProducts_KeyPress(KeyAscii As Integer)
    lsvProducts_Click
End Sub

Private Sub txtAmount_Change()
    lblChange = Val(txtAmount.Text) - Val(totaldummy)
    Text1 = lblChange
    lblChange = Format(lblChange, "P ###,###,###.00")
End Sub

Private Sub txtAmount_GotFocus()
    txtAmount = ""
End Sub

Private Sub txtFind_Change()
    Dim X
    Dim strSQL As String
    strSQL = "SELECT * FROM qryProducts"
    strSQL = strSQL & " WHERE sProductName LIKE'" & txtFind.Text & "%'"
    
    rs3.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    lsvProducts.ListItems.Clear
    With rs3
        Do While Not rs3.EOF
            Set X = lsvProducts.ListItems.Add(, , !iproductID)
                X.SubItems(1) = !sproductname
                X.SubItems(2) = Format(!cSellingPrice, "P ###,###,###.00")
                X.SubItems(3) = !sCategoryName
                X.SubItems(4) = !iQuantity & " " & !sMeasurementName
                X.SubItems(5) = Format(!dtexpiration)
                X.SubItems(6) = !sSupplierName
        .MoveNext
        Loop
    End With
    Set rs3 = Nothing
End Sub

Private Sub txtFind_GotFocus()
    Call Highlight(txtFind)
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lsvProducts.ListItems.Count <> 0 Then
            lsvProducts.SetFocus
        End If
    End If
End Sub

Private Sub txtQty_Change()
    txtPriceDummy = Val(txtQty.Text) * Val(txtPrice.Text)
End Sub
Sub ListCheck()
    Dim litmfound As ListItem
    Set litmfound = lsvList.FindItem(txtFind, 1, , 0)

    If litmfound Is Nothing Then
        
        FindItem = False
    Else
        MsgBox "This PRODUCT is already in the list" & vbCrLf _
               & "Input another PRODUCT", vbCritical + vbOKOnly, "Duplicate Item"
        litmfound.EnsureVisible
        litmfound.Selected = True
        FindItem = True
        fmeQty.Visible = False
        txtFind.SetFocus
    End If
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdAdd_Click
    End If
End Sub
