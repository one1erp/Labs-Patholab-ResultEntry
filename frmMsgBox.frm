VERSION 5.00
Begin VB.Form frmMsgBox 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label lblMsg 
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label lblSize 
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "frmMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MaxSize = 8000
Private Sub CmdOk_MouseUp(Button As Integer, Shift As Integer, X As Single, Y _
    As Single)
29400     Unload Me
End Sub

Public Sub ShowMsg(Message As String, Optional Title As String = "ResultEntry")
          
          Dim Size As Integer
29410     lblMsg.Caption = Message
29420     Me.Caption = Title
29430     lblSize.Caption = lblMsg.Caption
29440     lblSize.AutoSize = True
29450     If lblSize.Width > MaxSize Then
29460         Size = MaxSize
29470     Else
29480         Size = lblSize.Width
29490     End If
29500     lblMsg.Width = Size
29510     lblMsg.AutoSize = True
29520     lblMsg.WordWrap = True

29530     Me.Width = lblMsg.Width + 200
29540     lblMsg.Left = 100
29550     lblMsg.Top = 100
29560     Me.Width = lblMsg.Left + lblMsg.Width + 300
29570     Me.Height = lblMsg.Height * 8
29580     cmdOK.Top = lblMsg.Top + lblMsg.Height + 200
29590     cmdOK.Left = Me.Width / 2 - cmdOK.Width / 2
29600     Me.Show vbModal
End Sub

