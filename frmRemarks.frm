VERSION 5.00
Begin VB.Form frmRemarks 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   2160
   ScaleWidth      =   6840
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRemark 
      Alignment       =   1  'Right Justify
      Height          =   1935
      Left            =   120
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   6615
   End
End
Attribute VB_Name = "frmRemarks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public txt As String

'Private con As connection
'
'
'
'Public Sub Initialize(con_ As connection, strRequestDataId As String)
'On Error GoTo ERR_Initialize
'    Dim rs As Recordset
'    Dim sql As String
'
'    Set con = con_
'
'    sql = " select r.DESCRIPTION "
'    sql = sql & " from lims_sys.u_extra_request_data rd, "
'    sql = sql & "      lims_sys.u_extra_request_data_user rdu, "
'    sql = sql & "      lims_sys.u_extra_request r"
'    sql = sql & "  where rd.U_EXTRA_REQUEST_DATA_ID=rdu.U_EXTRA_REQUEST_DATA_ID"
'    sql = sql & "  and   r.U_EXTRA_REQUEST_ID=rdu.U_EXTRA_REQUEST_ID"
'    sql = sql & "  and   rd.U_EXTRA_REQUEST_DATA_ID=" & strRequestDataId
'
'    Set rs = con.Execute(sql)
'
'    If rs.EOF = True Then Exit Sub
'
'    txtRemark.Text = nte(rs("DESCRIPTION"))
'
'    Exit Sub
'ERR_Initialize:
'MsgBox "ERR_Initialize" & vbCrLf & Err.description
'End Sub

Public Sub Initialize(txt As String, locked As Boolean, strCaption As String)
32190 On Error GoTo ERR_Initialize

32200     txtRemark.Text = txt
32210     txtRemark.locked = locked
32220     frmRemarks.Caption = strCaption
      '    txtRemark.enabled = enabled

32230     Exit Sub
ERR_Initialize:
32240 MsgBox "ERR_Initialize" & vbCrLf & Err.Description
End Sub



Private Function nte(e As Variant) As Variant
32250     nte = IIf(IsNull(e), "", e)
End Function

'Private Sub Form_Click()
'    Unload Me
'End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
32260     If KeyAscii = vbKeyEscape Then
32270         txt = txtRemark.Text
32280         Me.Hide
32290     End If
End Sub

Private Sub Form_Click()
32300     txt = txtRemark.Text
32310     Me.Hide
End Sub

Private Sub txtRemark_KeyPress(KeyAscii As Integer)
32320     If KeyAscii = vbKeyEscape Then
32330         txt = txtRemark.Text
32340         Me.Hide
32350     End If
End Sub

'Private Sub txtRemark_Click()
'    txt = txtRemark.Text
'    Me.Hide
'End Sub

'Private Sub txtRemark_Click()
'    Unload Me
'    'Me.Hide
'End Sub

Private Sub Form_Unload(Cancel As Integer)
32360     txt = txtRemark.Text
End Sub

