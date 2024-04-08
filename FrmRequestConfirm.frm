VERSION 5.00
Begin VB.Form FrmRequestConfirm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Request Confirm"
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3570
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FrmRequestConfirm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   3570
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox TxtRequetsConfirm 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton CmdRequetsConfirm 
      Caption         =   "Confirm"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label LblRequetsConfirm 
      AutoSize        =   -1  'True
      Caption         =   "Request:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   2
      Top             =   200
      Width           =   945
   End
End
Attribute VB_Name = "FrmRequestConfirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const RED = &HFF&

Public ConfirmSucceeded As Boolean
Public Con As ADODB.connection
Public SdgName As String

Private Sub CmdClose_Click()
28670     ConfirmSucceeded = False
28680     Unload Me
End Sub

Private Sub CmdRequetsConfirm_Click()
28690     ConfirmSucceeded = CheckConfirm()
28700     If ConfirmSucceeded Then
28710         Unload Me
28720     End If
End Sub

Private Sub TxtRequetsConfirm_KeyDown(KeyCode As Integer, Shift As Integer)
28730     If Not KeyCode = vbKeyReturn Then Exit Sub

28740     ConfirmSucceeded = CheckConfirm()
28750     If ConfirmSucceeded Then
28760         Unload Me
28770     End If
End Sub

Private Sub Form_Activate()
28780     TxtRequetsConfirm.Text = ""
28790     Call TxtRequetsConfirm.SetFocus
28800     ConfirmSucceeded = False
End Sub

Private Sub Form_GotFocus()
28810     Call TxtRequetsConfirm.SetFocus
End Sub

Private Function CheckConfirm() As Boolean
28820     On Error GoTo ErrEnd
          Dim strSQL As String
          Dim ReqRs As ADODB.Recordset
          Dim RequestName As String
          Dim strMsg As String

28830     CheckConfirm = False

28840     If Trim(TxtRequetsConfirm.Text) = "" Then
28850         TxtRequetsConfirm.BackColor = RED
28860         MsgBox _
    " ! אנא הקש מספר דרישה התואמת את מספר הדרישה המיועדת לאישור ", , _
    "Nautilus - Request Confirm"
28870         TxtRequetsConfirm.BackColor = vbWhite
28880         TxtRequetsConfirm.Text = ""
28890         Call TxtRequetsConfirm.SetFocus
28900         Exit Function
28910     End If

28920     RequestName = Trim(TxtRequetsConfirm.Text)
28930     If InStr(RequestName, ".") > 0 Then
28940         RequestName = Left(RequestName, InStr(RequestName, ".") - 1)
28950     End If

28960     strSQL = "select " & "d.name request_name " & "from " & _
    "lims_sys.sdg d " & "where d.name = '" & UCase(RequestName) & "'"
28970     Set ReqRs = Con.Execute(strSQL)
28980     If ReqRs.EOF Then
28990         TxtRequetsConfirm.BackColor = RED
29000         MsgBox " ! הדרישה שהוקלדה, איננה קיימת במערכת ", , _
    "Nautilus - Request Confirm"
29010         TxtRequetsConfirm.BackColor = vbWhite
29020         Call TxtRequetsConfirm.SetFocus
29030         Exit Function
29040     End If

29050     If Trim(SdgName) <> Trim(nte(ReqRs("request_name"))) Then
29060         TxtRequetsConfirm.BackColor = RED
29070         strMsg = RequestName & " הדרישה שהוקלדה " & vbCrLf & SdgName & _
    " איננה תואמת את מספר הדרישה המיועדת לאישור "
29080         MsgBox strMsg, , "Nautilus - Request Confirm"
29090         TxtRequetsConfirm.BackColor = vbWhite
29100         Call TxtRequetsConfirm.SetFocus
29110         Exit Function
29120     End If

29130     CheckConfirm = True
29140     ReqRs.Close
29150     Exit Function

ErrEnd:
29160     MsgBox "CheckConfirm... " & vbCrLf & Err.Description
End Function

Private Function nte(e As Variant) As String
29170     nte = IIf(IsNull(e), "", e)
End Function
