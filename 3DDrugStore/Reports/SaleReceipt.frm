VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmSaleReceipt 
   Caption         =   "Receipt"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10065
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "SaleReceipt.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   10065
   StartUpPosition =   2  'CenterScreen
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   4410
      Left            =   75
      TabIndex        =   0
      Top             =   45
      Width           =   10020
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   0   'False
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   0   'False
      EnableSearchControl=   0   'False
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmSaleReceipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim Report As New SaleReceipt

Private Sub Form_Load()
    Call LoadRecieptData

    Screen.MousePointer = vbHourglass
    CRViewer1.ReportSource = SaleReceipt
    CRViewer1.ViewReport
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth

End Sub
Private Sub LoadRecieptData()
Dim i As Integer
     With SaleReceipt
         .txtDate.SetText frmSale.txtDate
         .txtTotal.SetText Format(frmSale.Text2, "P ###,###,###.00")
         For i = 1 To frmSale.lsvList.ListItems.Count
             '.txtProductName.SetText .txtProductName.Text & vbCrLf & frmSale.lsvList.ListItems.Item(i).SubItems(1) & " - " & frmSale.lsvList.ListItems(i).SubItems(2)
             .txtProductName.SetText .txtProductName.Text & vbCrLf & frmSale.lsvList.ListItems(i).SubItems(2) & " - " & frmSale.lsvList.ListItems.Item(i).SubItems(1) & " - " & Format(frmSale.lsvList.ListItems.Item(i).SubItems(3), "P###,###,###.00")
             '.txtVideo.SetText .txtVideo.Text & vbCrLf & frmRent.lsvList.ListItems.Item(i).SubItems(1) & vbTab & frmRent.lsvList.ListItems.Item(i).SubItems(2)
         Next i
         .txtAmtReceived.SetText Format(frmSale.txtAmount, "P ###,###,###.00")
         .txtChange.SetText Format(frmSale.Text1, "P ###,###,###.00")
    End With
End Sub
