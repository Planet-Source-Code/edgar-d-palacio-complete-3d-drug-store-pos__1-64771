Attribute VB_Name = "Module1"
Option Explicit
Public cn As New ADODB.Connection
Public rs As New ADODB.Recordset
Public rs1 As New ADODB.Recordset
Public rs2 As New ADODB.Recordset
Public rs3 As New ADODB.Recordset
Public rsUpdate As New ADODB.Recordset
Public lst As ListItem
Public dummyqtyreturn As Integer
Public dummydate

'Database connection
Public Sub DBConnect()
    On Error GoTo err_handler:
    
    cn.Open "Provider=MSDASQL.1;Persist Security Info=False;Data Source=POS"
    Exit Sub
err_handler:
    MsgBox Err.Description, vbCritical + vbOKOnly, "Database Connection Error"
    Exit Sub
End Sub

Public Sub DBClose()
    On Error GoTo err_handler:
    cn.Close
    Set cn = Nothing
err_handler:
    Exit Sub
End Sub

'Procedure for text highligh
Public Sub Highlight(ByRef sText As TextBox)
    With sText
        .SelStart = 0
        .SelLength = Len(sText.Text)
    End With
End Sub

'Procedure to center the form
Public Sub CenterForm(frm As Form)
    Dim TopCorner As Integer
    Dim LeftCorner As Integer
    
    If frm.WindowState <> 0 Then Exit Sub
    
    TopCorner = (Screen.Height - frm.Height) \ 2
    LeftCorner = (Screen.Width - frm.Width) \ 2
    frm.Move LeftCorner, TopCorner
End Sub

Function ListFindItem(lstCtrl As Control, lngSearch As Long) As Integer
   'just returns the position, does not set it
   'used to see if item is in list
   Dim intLen As Integer
   Dim intLoop As Integer
   Dim intPos As Integer

   intLen = lstCtrl.ListCount - 1
   intPos = -1
   For intLoop = 0 To intLen
      If lstCtrl.ItemData(intLoop) = lngSearch Then
         intPos = intLoop
         Exit For
      End If
   Next intLoop
   ListFindItem = intPos
End Function
