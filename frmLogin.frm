VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1545
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   912.837
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   495
      TabIndex        =   4
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2100
      TabIndex        =   5
      Top             =   1020
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "&User Name:"
      Height          =   195
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   212
      Width           =   840
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "&Password:"
      Height          =   195
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   592
      Width           =   735
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean
Public Pass As String

Private Sub cmdCancel_Click()
          'set the global var to false
          'to denote a failed login
27530     LoginSucceeded = False
27540     txtPassword = ""
27550     Me.Hide
End Sub

Private Sub cmdOK_Click()
          'check for correct password
27560     If txtPassword = Pass Then
              'place code to here to pass the
              'success to the calling sub
              'setting a global var is the easiest
27570         LoginSucceeded = True
27580         txtPassword = ""
27590         Me.Hide
27600     Else
27610         MsgBox "Invalid Password, try again!", , "Login"
27620         txtPassword.SetFocus
27630         SendKeys "{Home}+{End}"
27640     End If
End Sub

Private Sub Form_Activate()
27650     Call txtPassword.SetFocus
27660     LoginSucceeded = False
End Sub

Private Sub Form_GotFocus()
27670     txtPassword.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
27680     Cancel = 1
27690     cmdCancel_Click
End Sub
