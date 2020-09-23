VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProductDisplay 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "List of Products"
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
   Icon            =   "ProductDisplay.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   11775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3150
      TabIndex        =   3
      Top             =   300
      Width           =   1980
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   570
      Left            =   10440
      TabIndex        =   2
      Top             =   60
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   1005
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   120
      TabIndex        =   1
      Top             =   300
      Width           =   3000
   End
   Begin MSComctlLib.ListView lsvProduct 
      Height          =   7695
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   11520
      _ExtentX        =   20320
      _ExtentY        =   13573
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
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "id"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "PRODUCT NAME"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "CATEGORY"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "QUANTITY"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "PRICE"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "EXPIRATION DATE"
         Object.Width           =   4762
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9150
      Top             =   525
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ProductDisplay.frx":038A
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmProductDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    Call forminit
End Sub

Sub forminit()
    Call CenterForm(frmProductDisplay)
    Call loadlsvproduct
    Text1.SetFocus
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
                X.SubItems(4) = Format(!cSellingPrice, "P###,###,###.00")
                X.SubItems(5) = Format(!dtDatedelivered, "MMMM DD, YYYY")
                .MoveNext
            Loop
        End With
        Set rsproduct = Nothing
    End If
End Sub

Private Sub Text1_Change()
    Dim X
    Dim strSQL As String
    strSQL = "SELECT * FROM qryProducts "
    strSQL = strSQL & "WHERE sProductName LIKE'" & Text1.Text & "%'"
    
    rs2.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    lsvProduct.ListItems.Clear
    With rs2
        Do While Not rs2.EOF
           Set X = lsvProduct.ListItems.Add(, , !iproductID)
           X.SubItems(1) = !sproductname
            X.SubItems(2) = !sCategoryName
            X.SubItems(3) = !iQuantity & " " & !sMeasurementName
            X.SubItems(4) = Format(!cSellingPrice, "P###,###,###.00")
            X.SubItems(5) = Format(!dtDatedelivered, "MMMM DD, YYYY")
        .MoveNext
        Loop
    End With
    Set rs2 = Nothing
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lsvProduct.ListItems.Count <> 0 Then
            lsvProduct.SetFocus
        End If
    End If
End Sub

Private Sub Text2_Change()
    Dim X
    Dim strSQL As String
    strSQL = "SELECT * FROM qryProducts "
    strSQL = strSQL & "WHERE sCategoryName LIKE'" & Text2.Text & "%'"
    
    rs2.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    lsvProduct.ListItems.Clear
    With rs2
        Do While Not rs2.EOF
           Set X = lsvProduct.ListItems.Add(, , !iproductID)
           X.SubItems(1) = !sproductname
            X.SubItems(2) = !sCategoryName
            X.SubItems(3) = !iQuantity & " " & !sMeasurementName
            X.SubItems(4) = Format(!cSellingPrice, "P###,###,###.00")
            X.SubItems(5) = Format(!dtDatedelivered, "MMMM DD, YYYY")
        .MoveNext
        Loop
    End With
    Set rs2 = Nothing
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error GoTo Err
    Shell "calc.exe", vbNormalFocus
    Exit Sub
Err:
    MsgBox "You don't have a Calculator installed in your computer.", vbExclamation, "CSRS version 1"
End Sub
