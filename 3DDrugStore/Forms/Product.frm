VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmProduct 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Product Entry"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8040
   Icon            =   "Product.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   8040
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFind 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   795
      TabIndex        =   28
      Top             =   5400
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Height          =   4485
      Left            =   150
      TabIndex        =   11
      Top             =   750
      Width           =   7725
      Begin VB.TextBox txtExpiration 
         Appearance      =   0  'Flat
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
         Left            =   5055
         TabIndex        =   8
         Top             =   3120
         Width           =   2205
      End
      Begin MSComctlLib.ListView lsvProducts 
         Height          =   945
         Left            =   1410
         TabIndex        =   33
         Top             =   555
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
      Begin VB.TextBox txtDateEnter 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1410
         TabIndex        =   1
         Top             =   655
         Width           =   2010
      End
      Begin VB.TextBox txtProductDummy 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6630
         TabIndex        =   35
         Top             =   300
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.TextBox txtDeliveryDummy 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6240
         TabIndex        =   34
         Top             =   300
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.TextBox txtdatepurchased 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
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
         Left            =   1665
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   3120
         Width           =   2100
      End
      Begin VB.TextBox txtSellingPrice 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   5070
         TabIndex        =   5
         Top             =   2010
         Width           =   1560
      End
      Begin VB.TextBox txtReOrderLevel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   5055
         TabIndex        =   6
         Top             =   2380
         Width           =   555
      End
      Begin VB.TextBox txtWarranty 
         Appearance      =   0  'Flat
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
         Left            =   5055
         TabIndex        =   7
         Top             =   2750
         Width           =   2205
      End
      Begin VB.TextBox txtCategory 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1410
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1025
         Width           =   2400
      End
      Begin VB.TextBox txtSupplier 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1410
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1395
         Width           =   4230
      End
      Begin VB.TextBox txtNotes 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   1410
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   3765
         Width           =   4500
      End
      Begin VB.TextBox txtOnHand 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
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
         Left            =   1665
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   2010
         Width           =   1560
      End
      Begin VB.Frame Frame2 
         Height          =   90
         Left            =   135
         TabIndex        =   15
         Top             =   1785
         Width           =   7200
      End
      Begin VB.TextBox txtSupplierCost 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
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
         Left            =   1665
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   2380
         Width           =   1560
      End
      Begin VB.TextBox txtTotalCost 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
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
         Left            =   1665
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   2750
         Width           =   1560
      End
      Begin VB.Frame Frame3 
         Height          =   90
         Left            =   135
         TabIndex        =   12
         Top             =   3555
         Width           =   7200
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   300
         Left            =   7290
         TabIndex        =   27
         Top             =   2745
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   529
         _Version        =   393216
         Format          =   59179009
         CurrentDate     =   38753
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1410
         TabIndex        =   0
         Top             =   285
         Width           =   4785
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   300
         Left            =   3450
         TabIndex        =   2
         Top             =   660
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   529
         _Version        =   393216
         Format          =   59179009
         CurrentDate     =   38753
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   300
         Left            =   7290
         TabIndex        =   37
         Top             =   3120
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   529
         _Version        =   393216
         Format          =   59179009
         CurrentDate     =   38753
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Expiration:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4035
         TabIndex        =   38
         Top             =   3165
         Width           =   930
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Date Enter:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   315
         TabIndex        =   36
         Top             =   700
         Width           =   990
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Date Purchased:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   32
         Top             =   3165
         Width           =   1425
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Selling Price:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3840
         TabIndex        =   26
         Top             =   2055
         Width           =   1125
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Re - Order Level:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3465
         TabIndex        =   25
         Top             =   2430
         Width           =   1500
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Warranty:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4095
         TabIndex        =   24
         Top             =   2790
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Category:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   450
         TabIndex        =   23
         Top             =   1070
         Width           =   870
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Notes:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   765
         TabIndex        =   22
         Top             =   3765
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Product:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   600
         TabIndex        =   21
         Top             =   330
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Supplier:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   540
         TabIndex        =   20
         Top             =   1440
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "On - Hand:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   630
         TabIndex        =   19
         Top             =   2055
         Width           =   945
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Supplier Cost:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   345
         TabIndex        =   18
         Top             =   2425
         Width           =   1230
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Total Cost:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   630
         TabIndex        =   17
         Top             =   2790
         Width           =   945
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   330
      Left            =   5925
      TabIndex        =   10
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
      Left            =   75
      Top             =   0
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
            Picture         =   "Product.frx":038A
            Key             =   "add"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Product.frx":0724
            Key             =   "edit"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Product.frx":0ABE
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Product.frx":0E58
            Key             =   "search"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lsvProduct 
      Height          =   1395
      Left            =   150
      TabIndex        =   29
      Top             =   5775
      Width           =   7725
      _ExtentX        =   13626
      _ExtentY        =   2461
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "id"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   3528
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
         Object.Width           =   3528
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
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Selling Price"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Re - Order Level"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Warranty"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Date Entered"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Supplier"
         Object.Width           =   4410
      EndProperty
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   150
      TabIndex        =   30
      Top             =   5445
      Visible         =   0   'False
      Width           =   570
   End
End
Attribute VB_Name = "frmProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public FindItem As Boolean

Private Sub Form_Activate()
    Call forminit
End Sub

Sub forminit()
    Call CenterForm(frmProduct)
    Toolbar1.Buttons(4).Enabled = False
    Toolbar1.Buttons(5).Enabled = False
    Call loadlsvProducts
    Call loadlsvproduct
    Text2.Enabled = True
End Sub

Private Sub DTPicker1_Change()
    txtDateEnter.Text = Format(DTPicker1.Value, "MMMM DD, YYYY")
End Sub

Private Sub DTPicker3_Change()
    txtExpiration.Text = Format(DTPicker3.Value, "MMMM DD, YYYY")
End Sub

Private Sub DTPicker2_Change()
    txtWarranty.Text = Format(DTPicker2.Value, "MMMM DD, YYYY")
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Call textclear
        Toolbar1.Buttons(2).Enabled = True
        Toolbar1.Buttons(4).Enabled = False
        Toolbar1.Buttons(5).Enabled = False
        lsvProducts.Visible = False
        txtDateEnter.SetFocus
    End If
End Sub

Private Sub lsvProduct_Click()
    Dim X As Integer
    Dim strSQL As String
    Dim dummy
    Dim row
    
    Text2.Enabled = False
    Toolbar1.Buttons(2).Enabled = False
    Toolbar1.Buttons(4).Enabled = True
    Toolbar1.Buttons(5).Enabled = True
    row = lsvProduct.SelectedItem.Index
    dummy = lsvProduct.ListItems.Item(row).Text
    
    strSQL = "SELECT * FROM qryProducts "
    strSQL = strSQL & "WHERE iProductId=" & dummy
    
    rs.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    
    With rs
        txtProductDummy = dummy
        txtDeliveryDummy = !iDeliveryID
        txtDateEnter = !dtdateentered
        Text2 = !sproductname
        txtCategory = !sCategoryName
        txtSupplier = !sSupplierName
        txtOnHand = !iQuantity & " " & !sMeasurementName
        txtSupplierCost = !cCost
        txtTotalCost = !cTotalCost
        txtdatepurchased = !dtDatedelivered
        txtSellingPrice = !cSellingPrice
        txtReOrderLevel = !iReOrderLevel
        txtWarranty = !dtWarranty
        txtNotes = !sNotes
    End With
        txtDateEnter.SetFocus
        Set rs = Nothing
End Sub

Private Sub lsvProduct_KeyPress(KeyAscii As Integer)
    lsvProduct_Click
End Sub

Private Sub lsvProducts_Click()
    Dim X As Integer
    Dim strSQL As String
    Dim dummy
    Dim row
    
    lsvProducts.Visible = False
    row = lsvProducts.SelectedItem.Index
    dummy = lsvProducts.ListItems.Item(row).Text
    
    strSQL = "SELECT * FROM qrydelivery "
    strSQL = strSQL & "WHERE ideliveryId=" & dummy
    
    rs1.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    
    With rs1
        txtDeliveryDummy = dummy
        Text2 = !sproductname
        txtCategory = !sCategoryName
        txtSupplier = !sSupplierName
        txtOnHand = !iQuantity & " " & !sMeasurementName
        txtSupplierCost = Format(!cCost, "P ###,###,###.00")
        txtTotalCost = Format(!cTotalCost, "P ###,###,###.00")
        txtdatepurchased = !dtDatedelivered
        dummydate = !dtDatedelivered + 10
        txtWarranty.Text = Format(dummydate, "MMMM DD, YYYY")
    End With
        txtDateEnter.SetFocus
        Set rs1 = Nothing
End Sub

Private Sub lsvProducts_KeyPress(KeyAscii As Integer)
    lsvProducts_Click
End Sub

Private Sub Text2_Change()
    Dim X
    Dim strSQL As String
    strSQL = "SELECT * FROM qryDelivery "
    strSQL = strSQL & "WHERE sProductName LIKE'" & Text2.Text & "%'"
    
    rs3.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    lsvProducts.ListItems.Clear
    With rs3
        Do While Not rs3.EOF
            Set X = lsvProducts.ListItems.Add(, , !iDeliveryID)
                X.SubItems(1) = !sproductname
        .MoveNext
        Loop
    End With
    Set rs3 = Nothing
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

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    
        Case "add"
            Dim strSQLAdd As String
            Dim rsAdd As New Recordset
            Call ListCheck
            If FindItem = False Then
                If complete = True Then
                    strSQLAdd = "SELECT * FROM  tblproduct"
                    rsAdd.Open strSQLAdd, cn, adOpenDynamic, adLockOptimistic
                    
                        With rsAdd
                            .AddNew
                            !iDeliveryID = txtDeliveryDummy
                            !dtdateentered = txtDateEnter
                            !cSellingPrice = txtSellingPrice
                            !iReOrderLevel = txtReOrderLevel
                            !dtWarranty = txtWarranty
                            !dtexpiration = txtExpiration
                            !sNotes = txtNotes
                            !sStatus = "Ok"
                            !sStatus1 = "Ok"
                            .Update
                        End With
                        MsgBox "New record added to the database", vbOKOnly, "Product Entry"
                        Call textclear
                        Call loadlsvproduct
                        lsvProducts.Visible = False
                        Text2.SetFocus
                        Set rsAdd = Nothing
                Else
                    
                    If txtDeliveryDummy = "" Then
                        MsgBox "Please choose a product", vbOKOnly, "Product Entry"
                        Text2.SetFocus
                        Exit Sub
                    ElseIf txtDateEnter = "" Then
                        MsgBox "Please specify date enter", vbOKOnly, "Product Entry"
                        txtDateEnter.SetFocus
                        Exit Sub
                    ElseIf txtSellingPrice = "" Then
                        MsgBox "Please enter selling price", vbOKOnly, "Product Entry"
                        txtSellingPrice.SetFocus
                        Exit Sub
                    ElseIf txtReOrderLevel = "" Then
                        MsgBox "Please entr re - order level", vbOKOnly, "Product Entry"
                        txtReOrderLevel.SetFocus
                        Exit Sub
                    ElseIf txtWarranty = "" Then
                        MsgBox "Please specify warranty", vbOKOnly, "Product Entry"
                        txtWarranty.SetFocus
                        Exit Sub
                    ElseIf txtExpiration = "" Then
                        MsgBox "Please specify expiration", vbOKOnly, "Product Entry"
                        txtExpiration.SetFocus
                    End If
                End If
            End If
            Call loadlsvproduct
        Case "update"
            Dim strSQLEdit As String
            Dim rsEdit As New Recordset
            If complete = True Then
                strSQLEdit = "SELECT * FROM  tblproduct"
                 strSQLEdit = strSQLEdit & " WHERE iProductID =" & txtProductDummy
                 
                rsEdit.Open strSQLEdit, cn, adOpenDynamic, adLockOptimistic
                With rsEdit
                    !iDeliveryID = txtDeliveryDummy
                    !dtdateentered = txtDateEnter
                    !cSellingPrice = txtSellingPrice
                    !iReOrderLevel = txtReOrderLevel
                    !dtWarranty = txtWarranty
                    !dtexpiration = txtExpiration
                    !sNotes = txtNotes
                    .Update
                End With
                    MsgBox "The changes you made was successfully updated", vbOKOnly, "Product Entry"
                    Call textclear
                    Call loadlsvproduct
                    Text2.SetFocus
                    Set rsEdit = Nothing
            Else
                
                If txtDeliveryDummy = "" Then
                    MsgBox "Please choose a product", vbOKOnly, "Product Entry"
                    Text2.SetFocus
                    Exit Sub
                ElseIf txtDateEnter = "" Then
                    MsgBox "Please specify date enter", vbOKOnly, "Product Entry"
                    txtDateEnter.SetFocus
                    Exit Sub
                ElseIf txtSellingPrice = "" Then
                    MsgBox "Please enter selling price", vbOKOnly, "Product Entry"
                    txtSellingPrice.SetFocus
                    Exit Sub
                ElseIf txtReOrderLevel = "" Then
                    MsgBox "Please entr re - order level", vbOKOnly, "Product Entry"
                    txtReOrderLevel.SetFocus
                    Exit Sub
                ElseIf txtWarranty = "" Then
                    MsgBox "Please specify warranty", vbOKOnly, "Product Entry"
                    txtWarranty.SetFocus
                    Exit Sub
                ElseIf txtExpiration = "" Then
                    MsgBox "Please specify expiration", vbOKOnly, "Product Entry"
                    txtExpiration.SetFocus
                End If
            End If
        Case "delete"
            Dim strSQLDelete As String
            Dim rsDelete As New ADODB.Recordset
            Dim answerDelete As String
            
            strSQLDelete = "SELECT * FROM tblProduct "
            strSQLDelete = strSQLDelete & "WHERE iProductID=" & txtProductDummy
            
            answerDelete = MsgBox("Are you sure you want DELETE this record?", vbQuestion + vbYesNo, "Category Maintenance")
            
            If answerDelete = vbYes Then
                rsDelete.Open strSQLDelete, cn, adOpenDynamic, adLockOptimistic
                With rsDelete
                    .Delete
                End With
                MsgBox "The record was successfully deleted", , "Product Entry"
            End If
            Set rsDelete = Nothing
            Call textclear
            txtFind = ""
            txtFind.Visible = True
            Label3.Visible = True
            txtFind.SetFocus
            Call loadlsvproduct
        Case "search"
            Label11.Visible = True
            txtFind.Visible = True
            txtFind.SetFocus
    End Select
End Sub

Sub ListCheck()
    Dim litmfound As ListItem
    Set litmfound = lsvProduct.FindItem(Text2, 1, , 0)

    If litmfound Is Nothing Then
        
        FindItem = False
    Else
        MsgBox "This PRODUCT is already in the list" & vbCrLf _
               & "Input another PRODUCT", vbCritical + vbOKOnly, "Duplicate Item"
        litmfound.EnsureVisible
        litmfound.Selected = True
        FindItem = True
        Text2 = ""
        Text2.SetFocus
    End If
End Sub

Sub loadlsvProducts()
    Dim X
    Dim strSQL As String
    
    strSQL = "SELECT * FROM qryDelivery"
    
    rs.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    lsvProducts.ListItems.Clear
    If rs.EOF = True Then
        lsvProducts.ListItems.Clear
    Else
        With rs
            Do While Not rs.EOF
                Set X = lsvProducts.ListItems.Add(, , !iDeliveryID)
                    X.SubItems(1) = !sproductname
                    .MoveNext
            Loop
        End With
        Set rs = Nothing
    End If
End Sub

Function complete()
    If txtSellingPrice = "" Or txtReOrderLevel = "" Or txtWarranty = "" Or txtDateEnter = "" Or txtExpiration = "" Then
        complete = False
    Else
        complete = True
    End If
End Function

Sub textclear()
    txtDateEnter = ""
    Text2 = ""
    txtCategory = ""
    txtSupplier = ""
    txtOnHand = ""
    txtSupplierCost = ""
    txtTotalCost = ""
    txtdatepurchased = ""
    txtSellingPrice = ""
    txtReOrderLevel = ""
    txtWarranty = ""
    txtExpiration = ""
    txtNotes = ""
End Sub

Sub loadlsvproduct()
    On Error GoTo err_handler:
    Dim X
    Dim strSQL As String
    
    strSQL = "SELECT * FROM qryproducts"
    
    rs2.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    lsvProduct.ListItems.Clear
    If rs2.EOF = True Then
        lsvProduct.ListItems.Clear
    Else
        With rs2
            Do While Not rs2.EOF
                Set X = lsvProduct.ListItems.Add(, , !iproductID)
                X.SubItems(1) = !sproductname
                X.SubItems(2) = !sCategoryName
                X.SubItems(3) = !iQuantity & " " & !sMeasurementName
                X.SubItems(4) = Format(!cCost, "P ###,###,###.00")
                X.SubItems(5) = Format(!cTotalCost, "P ###,###,###.00")
                X.SubItems(6) = Format(!cSellingPrice, "P ###,###,###.00")
                X.SubItems(7) = !iReOrderLevel
                X.SubItems(8) = !dtWarranty
                X.SubItems(9) = !dtdateentered
                X.SubItems(10) = !sSupplierName
                .MoveNext
            Loop
        End With
        Set rs2 = Nothing
err_handler:
        Set rs2 = Nothing
    End If
End Sub

Private Sub txtDateEnter_Change()
    Call ListCheck
End Sub

Private Sub txtFind_Change()
    Dim X
    Dim strSQL As String
    strSQL = "SELECT * FROM qryProducts"
    strSQL = strSQL & " WHERE sProductName LIKE'" & txtFind.Text & "%'"
    
    rs.Open strSQL, cn, adOpenDynamic, adLockOptimistic
    lsvProduct.ListItems.Clear
    With rs
        Do While Not rs.EOF
            Set X = lsvProduct.ListItems.Add(, , !iproductID)
            X.SubItems(1) = !sproductname
            X.SubItems(2) = !sCategoryName
            X.SubItems(3) = !iQuantity & " " & !sMeasurementName
            X.SubItems(4) = Format(!cCost, "P ###,###,###.00")
            X.SubItems(5) = Format(!cTotalCost, "P ###,###,###.00")
            X.SubItems(6) = Format(!cSellingPrice, "P ###,###,###.00")
            X.SubItems(7) = !iReOrderLevel
            X.SubItems(8) = !dtWarranty
            X.SubItems(9) = !dtdateentered
            X.SubItems(10) = !sSupplierName
        .MoveNext
        Loop
    End With
    Set rs = Nothing
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lsvProduct.ListItems.Count <> 0 Then
            lsvProduct.SetFocus
        End If
    End If
End Sub
