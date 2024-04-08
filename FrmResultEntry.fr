VERSION 5.00
Object = "{114CD60F-7B12-42B6-A320-BFE95C4D00F4}#1.0#0"; "ResultEntry.ocx"
Begin VB.Form FrmResultEntry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nautilus - Result Entry"
   ClientHeight    =   9420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15240
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9420
   ScaleWidth      =   15240
   Begin ResultEntry.ResultEntryCtrl ResultEntryCtrl 
      Height          =   9135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   15615
      _ExtentX        =   27543
      _ExtentY        =   16113
   End
End
Attribute VB_Name = "FrmResultEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    Me.Width = ResultEntryCtrl.Width
    Me.Height = ResultEntryCtrl.Height
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 0
End Sub

Private Sub ResultEntryCtrl_CloseClicked()
    Me.Hide
End Sub



