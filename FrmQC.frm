VERSION 5.00
Begin VB.Form FrmQC 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "QC Confirm"
   ClientHeight    =   1275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3540
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FrmQC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   3540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameQC 
      Height          =   615
      Left            =   150
      TabIndex        =   3
      Top             =   0
      Width           =   3255
      Begin VB.CheckBox CheckQC 
         Caption         =   "QC Confirm"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   720
         TabIndex        =   0
         Top             =   175
         Width           =   1815
      End
   End
   Begin VB.CommandButton CmdQCConfirm 
      Caption         =   "Confirm"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   750
      Width           =   1215
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   750
      Width           =   1215
   End
End
Attribute VB_Name = "FrmQC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const RED = &HFF&

Public ConfirmSucceeded As Boolean

Private Sub CmdClose_Click()
29180     ConfirmSucceeded = False
29190     Unload Me
End Sub

Private Sub CmdQCConfirm_Click()
29200     ConfirmSucceeded = CheckConfirm()
29210     If ConfirmSucceeded Then
29220         Unload Me
29230     End If
End Sub

Private Sub Form_Activate()
29240     CheckQC.value = 0
29250     Call CheckQC.SetFocus
29260     ConfirmSucceeded = False
End Sub

Private Sub Form_GotFocus()
29270     Call CheckQC.SetFocus
End Sub

Private Function CheckConfirm() As Boolean
29280     On Error GoTo ErrEnd

29290     CheckConfirm = False

29300     If CheckQC.value = 0 Then
29310         CheckQC.BackColor = RED
29320         MsgBox " ! אנא בחר את תיבת הסימון לאישור QC ", , _
    "Nautilus - QC Confirm"
29330         CheckQC.BackColor = vbWhite
29340         Call CheckQC.SetFocus
29350         Exit Function
29360     End If
29370     CheckConfirm = True
29380     Exit Function

ErrEnd:
29390     MsgBox "CheckConfirm... " & vbCrLf & Err.Description
End Function

