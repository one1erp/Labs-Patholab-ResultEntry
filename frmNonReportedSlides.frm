VERSION 5.00
Begin VB.Form frmNonReportedSlides 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "סליידים לא צבועים"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5565
   ScaleWidth      =   4365
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmStep2 
      Height          =   1455
      Left            =   -480
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   3000
      Width           =   3615
      Begin VB.CheckBox chkLetter 
         Alignment       =   1  'Right Justify
         Caption         =   "שליחת מכתב תשובה לפורטל"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   720
         Width           =   3135
      End
      Begin VB.CheckBox chkRevise 
         Alignment       =   1  'Right Justify
         Caption         =   "פתיחת רוויזיה"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame frmStep1 
      Caption         =   "האם להמשיך בחתימה?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   360
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   2640
      Width           =   3615
      Begin VB.OptionButton opt 
         Alignment       =   1  'Right Justify
         Caption         =   "לא"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   480
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   840
         Width           =   2895
      End
      Begin VB.OptionButton opt 
         Alignment       =   1  'Right Justify
         Caption         =   "כן"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   360
         Width           =   3015
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "אישור"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   4320
      Width           =   3615
   End
   Begin VB.ListBox lstSlides 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1680
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   3615
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "קיימים סליידים לא צבועים:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   3615
   End
End
Attribute VB_Name = "frmNonReportedSlides"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private shouldAuthorise As Boolean
Private shouldRvise As Boolean
Private shouldSendLetter As Boolean

Private nStep As Integer


Public Function GetAuthorise() As Boolean
32370     GetAuthorise = shouldAuthorise
End Function

Public Function GetRevise() As Boolean
32380     GetRevise = shouldRvise
End Function

Public Function GetSendLetter() As Boolean
32390     GetSendLetter = shouldSendLetter
End Function

Public Sub Initialize(d As Dictionary)
32400 On Error GoTo ERR_Initialize

          Dim i As Integer
          
          
32410     shouldAuthorise = False
32420     shouldRvise = False
32430     nStep = 1
      '    lblQuestion.Visible = True
      '    lblQuestion.Caption = "האם להמשיך בחתימה?"
          
32440     Call lstSlides.Clear
          
32450     For i = 0 To d.Count - 1
32460         Call lstSlides.AddItem(d.Keys(i))
32470     Next i
          
32480     opt(0).value = False
32490     opt(1).value = False
32500     cmdOK.Enabled = False
32510     chkRevise.value = 0
32520     chkLetter.value = 0
          
32530     frmStep1.Visible = True
32540     frmStep2.Visible = False
32550     frmStep2.Top = frmStep1.Top
32560     frmStep2.Left = frmStep1.Left

32570     Exit Sub
ERR_Initialize:
32580 MsgBox "ERR_Initialize" & vbCrLf & Err.Description
End Sub


Private Sub chkRevise_Click()
32590 On Error GoTo ERR_chkRevise_Click

32600     chkLetter.Visible = chkRevise.value <> 0
         
32610     Exit Sub
ERR_chkRevise_Click:
32620 MsgBox "ERR_chkRevise_Click" & vbCrLf & Err.Description
End Sub

Private Sub cmdOK_Click()
32630 On Error GoTo ERR_cmdOK_Click


32640     Select Case nStep
              Case 1
32650             If opt(0).value = False Then
32660                 Me.Hide
32670             Else
32680                 shouldAuthorise = True
32690             End If
                  
32700             frmStep1.Visible = False
32710             frmStep2.Visible = True
32720             chkRevise.value = 1
32730             chkLetter.value = 1
                  
32740         Case 2
32750             shouldRvise = chkRevise.value <> 0
32760             shouldSendLetter = chkLetter.value <> 0
      '            shouldRvise = opt(0).Value
32770             Me.Hide
              
32780     End Select
          
32790     nStep = nStep + 1

32800     Exit Sub
ERR_cmdOK_Click:
32810 MsgBox "ERR_cmdOK_Click" & vbCrLf & Err.Description
End Sub



Private Sub opt_Click(index As Integer)
32820 On Error GoTo ERR_opt_Click

32830     cmdOK.Enabled = True

32840     Exit Sub
ERR_opt_Click:
32850 MsgBox "ERR_opt_Click" & vbCrLf & Err.Description
End Sub
