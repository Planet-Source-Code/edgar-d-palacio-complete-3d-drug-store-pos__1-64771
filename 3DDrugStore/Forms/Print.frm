VERSION 5.00
Begin VB.Form frmPrint 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Selection"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Print.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Report Option"
      Height          =   2910
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.OptionButton optReturn 
         Caption         =   "Return"
         Height          =   375
         Left            =   480
         TabIndex        =   10
         Top             =   2280
         Width           =   1935
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Preview"
         Height          =   375
         Left            =   2880
         TabIndex        =   9
         Top             =   2280
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Preview"
         Height          =   375
         Left            =   2880
         TabIndex        =   8
         Top             =   1800
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Preview"
         Height          =   375
         Left            =   2880
         TabIndex        =   7
         Top             =   1320
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Preview"
         Height          =   375
         Left            =   2880
         TabIndex        =   6
         Top             =   840
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Preview"
         Height          =   375
         Left            =   2880
         TabIndex        =   5
         Top             =   360
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.OptionButton optSales 
         Caption         =   "Sales"
         Height          =   375
         Left            =   480
         TabIndex        =   4
         Top             =   1800
         Width           =   1935
      End
      Begin VB.OptionButton optWarranty 
         Caption         =   "Product Warranty"
         Height          =   375
         Left            =   480
         TabIndex        =   3
         Top             =   1320
         Width           =   1935
      End
      Begin VB.OptionButton optExpiration 
         Caption         =   "Product Expiration"
         Height          =   375
         Left            =   480
         TabIndex        =   2
         Top             =   840
         Width           =   1935
      End
      Begin VB.OptionButton optProducts 
         Caption         =   "List of Products"
         Height          =   375
         Left            =   480
         TabIndex        =   1
         Top             =   360
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Load frmReportProductList
    frmReportProductList.Show 1
End Sub

Private Sub Command2_Click()
    Load frmReportExpiration
    frmReportExpiration.Show 1
End Sub

Private Sub Command3_Click()
    Load frmReportWarranty
    frmReportWarranty.Show 1
End Sub

Private Sub Command4_Click()
    Load frmReportSale
    frmReportSale.Show 1
End Sub

Private Sub Command5_Click()
    Load frmReportReturn
    frmReportReturn.Show 1
End Sub

Private Sub Form_Activate()
    If frmMain.Text2 = "Administrator" Then
        optProducts.Visible = True
        optExpiration.Visible = True
        optWarranty.Visible = True
        optSales.Visible = True
        optReturn.Visible = True
    ElseIf frmMain.Text2 = "Cashier" Then
        optProducts.Visible = False
        optExpiration.Visible = False
        optWarranty.Visible = False
        optSales.Visible = True
        optReturn.Visible = False
    End If
End Sub

Private Sub Form_Load()
    optProducts.Value = False
    optExpiration.Value = False
    optWarranty.Value = False
    optSales.Value = False
    optReturn.Value = False
End Sub

Private Sub optExpiration_Click()
    If optExpiration.Value = True Then
        Command2.Visible = True
        Command1.Visible = False
        Command3.Visible = False
        Command4.Visible = False
        Command5.Visible = False
    End If
End Sub

Private Sub optProducts_Click()
    If optProducts.Value = True Then
        Command1.Visible = True
        Command2.Visible = False
        Command3.Visible = False
        Command4.Visible = False
        Command5.Visible = False
    End If
End Sub

Private Sub optReturn_Click()
    If optReturn.Value = True Then
        Command5.Visible = True
        Command1.Visible = False
        Command2.Visible = False
        Command3.Visible = False
        Command4.Visible = False
    End If
End Sub

Private Sub optSales_Click()
    If optSales.Value = True Then
        Command4.Visible = True
        Command1.Visible = False
        Command2.Visible = False
        Command3.Visible = False
        Command5.Visible = False
    End If
End Sub

Private Sub optWarranty_Click()
    If optWarranty.Value = True Then
        Command3.Visible = True
        Command1.Visible = False
        Command2.Visible = False
        Command4.Visible = False
        Command5.Visible = False
    End If
End Sub
