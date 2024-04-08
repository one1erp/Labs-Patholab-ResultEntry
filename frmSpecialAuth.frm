VERSION 5.00
Begin VB.Form frmSpecialAuth 
   Caption         =   "Special Authorization"
   ClientHeight    =   3465
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5325
   LinkTopic       =   "Form1"
   ScaleHeight     =   3465
   ScaleWidth      =   5325
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraWizard 
      Height          =   2655
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   5055
      Begin VB.ComboBox cmbAuthOrder 
         Height          =   315
         ItemData        =   "frmSpecialAuth.frx":0000
         Left            =   840
         List            =   "frmSpecialAuth.frx":0007
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1920
         Width           =   975
      End
      Begin VB.ComboBox cmbAuthDoctor 
         Height          =   315
         Left            =   840
         TabIndex        =   10
         Top             =   840
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "Authorizing pathologist"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   840
         TabIndex        =   9
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.Frame fraWizard 
      Height          =   2655
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   5055
      Begin VB.ComboBox cmbExternal 
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Top             =   2040
         Width           =   3615
      End
      Begin VB.ComboBox cmbReferringDoc 
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   3615
      End
      Begin VB.Label Label3 
         Caption         =   "External Consultant"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1560
         Width           =   2895
      End
      Begin VB.Label Label2 
         Caption         =   "Referring pathologist"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame fraWizard 
      Height          =   2655
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   5055
      Begin VB.OptionButton optType 
         Caption         =   "Examined by another pathologist"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   4455
      End
      Begin VB.OptionButton optType 
         Caption         =   "External consultant"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   8
         Top             =   1320
         Width           =   3015
      End
      Begin VB.OptionButton optType 
         Caption         =   "Authorize on behalf of"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   3495
      End
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton btnBack 
      Caption         =   "<< Back"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton btnNext 
      Caption         =   "Next >>"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   2880
      Width           =   1215
   End
End
Attribute VB_Name = "frmSpecialAuth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Con As ADODB.connection
Public sdgId As Double

Private Doctors As New Dictionary
Private Externals As New Dictionary
Private HasRecord As Boolean

Private Sub btnBack_Click()
27700     btnNext.Caption = "Next >>"
27710     btnBack.Enabled = False
27720     fraWizard(1).Visible = False
27730     fraWizard(2).Visible = False
27740     fraWizard(0).Visible = True
End Sub

Private Sub btnClose_Click()
27750     Unload Me
End Sub

Private Sub btnNext_Click()
          Dim Operator As String
          Dim sender As String
          Dim utype As String
          Dim rst As ADODB.Recordset
          Dim AuthOrder As String
          
27760     AuthOrder = "null"
          
27770     If btnNext.Caption = "Next >>" Then
27780         btnNext.Caption = "Finish"
27790         btnBack.Enabled = True
27800         fraWizard(0).Visible = False
27810         If optType(2).value = True Then
27820             fraWizard(2).Visible = True
27830         Else
27840             If optType(1).value Then
27850                 cmbAuthOrder.Visible = True
27860             Else
27870                 cmbAuthOrder.Visible = False
27880             End If
27890             fraWizard(1).Visible = True
27900         End If
27910     Else
27920         If optType(0).value Then
27930             utype = "I"
27940             Operator = Doctors(cmbAuthDoctor.Text)
27950             sender = "null"
27960         ElseIf optType(1).value Then
27970             utype = "A"
27980             Operator = Doctors(cmbAuthDoctor.Text)
27990             sender = "null"
28000             AuthOrder = cmbAuthOrder.Text
28010         Else
28020             utype = "C"
28030             Operator = Doctors(cmbReferringDoc.Text)
28040             sender = Externals(cmbExternal.Text)
28050         End If
28060         If HasRecord Then
28070             Call Con.Execute("update lims_sys.u_special_inspection_user " _
    & "set u_type = '" & utype & "', " & "u_operator = " & Operator & ", " & _
    "u_sender_pathologist = " & sender & ", " & "u_order = " & AuthOrder & " " & _
    "where u_sdg_id = " & sdgId)
28080         Else
28090             Set rst = _
    Con.Execute("select lims.sq_u_special_inspection.nextval from dual")
28100             Con.BeginTrans
28110             Call Con.Execute("insert into lims_sys.u_special_inspection " _
    & "(u_special_inspection_id, name, version, version_status) " & "values (" & _
    rst(0) & ",'" & sdgId & "','1','A')")
28120             Call _
    Con.Execute("insert into lims_sys.u_special_inspection_user " & _
    "(u_special_inspection_id, u_type, u_operator, u_sender_pathologist, u_date, u_sdg_id, u_order)" _
    & "values (" & rst(0) & ",'" & utype & "'," & Operator & "," & sender & _
    ",sysdate, " & sdgId & "," & AuthOrder & ")")
28130             Con.CommitTrans
28140             rst.Close
28150         End If
28160         Unload Me
28170     End If
End Sub

Private Sub Form_Activate()
          Dim rst As ADODB.Recordset
          Dim sql As String
          
28180     HasRecord = False
          
28190     sql = "select ol.full_name int, si.u_type, oe.full_name ext " & _
    "from lims_sys.u_special_inspection_user si, lims_sys.operator ol, lims_sys.operator oe " _
    & "where si.u_operator = ol.operator_id and " & _
    "si.u_sender_pathologist = oe.operator_id(+) and " & "si.u_sdg_id = " & sdgId
28200     Set rst = Con.Execute(sql)
28210     If rst.EOF Then Exit Sub
28220     cmbAuthDoctor.Text = nte(rst("INT"))
28230     cmbExternal.Text = nte(rst("EXT"))
28240     cmbReferringDoc.Text = nte(rst("INT"))
28250     Select Case nte(rst("U_TYPE"))
          Case "A"
28260         optType(1).value = True
28270     Case "I"
28280         optType(0).value = True
28290     Case "C"
28300         optType(2).value = True
28310     End Select
28320     HasRecord = True
28330     rst.Close
End Sub

Private Sub Form_Load()
          Dim rstDoctors As ADODB.Recordset
          Dim rstExternals As ADODB.Recordset
          Dim sql As String
          
28340     sql = " select o.OPERATOR_ID,"
28350     sql = sql & "  o.FULL_NAME"
28360     sql = sql & " from lims_sys.operator o, "
28370     sql = sql & "      lims_sys.operator_user ou,"
28380     sql = sql & "      lims_sys.lims_role r"
28390     sql = sql & " where ou.OPERATOR_ID=o.OPERATOR_ID"
28400     sql = sql & " and   o.ROLE_ID=r.role_id"
28410     sql = sql & " and upper(r.name)='DOCTOR'"
28420     sql = sql & " and ou.U_ORDER > 0 "
28430     sql = sql & " order by ou.U_ORDER"
          
      '    Set rstDoctors = con.Execute("select o.operator_id, o.full_name " & "from lims_sys.operator o, lims_sys.lims_role r, lims_sys.operator_role opr " & "where o.operator_id = opr.operator_id and " & "opr.role_id = r.role_id and " & "upper(r.name) = 'DOCTOR'")
          
28440     Set rstDoctors = Con.Execute(sql)
28450     While Not rstDoctors.EOF
28460         Call Doctors.Add(rstDoctors("FULL_NAME").value, _
    rstDoctors("OPERATOR_ID").value)
28470         cmbAuthDoctor.AddItem (rstDoctors("FULL_NAME"))
28480         cmbReferringDoc.AddItem (rstDoctors("FULL_NAME"))
28490         rstDoctors.MoveNext
28500     Wend
28510     Set rstExternals = Con.Execute("select o.operator_id, o.full_name " & _
    "from lims_sys.operator o, lims_sys.lims_role r, lims_sys.operator_role opr " & _
    "where o.operator_id = opr.operator_id and " & "opr.role_id = r.role_id and " & _
    "upper(r.name) = 'EXTERNAL ADVISOR'")
28520     While Not rstExternals.EOF
28530         Call Externals.Add(rstExternals("FULL_NAME").value, _
    rstExternals("OPERATOR_ID").value)
28540         cmbExternal.AddItem (rstExternals("FULL_NAME"))
28550         rstExternals.MoveNext
28560     Wend
28570     rstDoctors.Close
28580     rstExternals.Close
28590     cmbAuthOrder.list(0) = "1"
28600     cmbAuthOrder.list(1) = "2"
28610     cmbAuthOrder.Text = "2"
          
              
          'do not use the option of external consultant:
28620     optType(2).Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
28630     Doctors.RemoveAll
28640     Externals.RemoveAll
End Sub

Private Sub optType_Click(index As Integer)
28650     btnNext.Enabled = True
End Sub

Private Function nte(e As Variant) As String
28660     nte = IIf(IsNull(e), "", e)
End Function
