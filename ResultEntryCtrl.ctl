VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{4016B910-CCE8-4B27-95FA-006C7152BC93}#2.16#0"; "MacabiShared.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{30CB9C1A-EE46-4D4C-BBDE-1D306015D2DD}#47.8#0"; "RequestRemark.ocx"
Object = "{53DB53AB-C26B-45DB-AD59-AEB893A8A326}#6.1#0"; "Snomed.ocx"
Object = "{E40B1134-8362-494C-99D9-AB6AD0E21EB5}#6.21#0"; "Organ.ocx"
Object = "{309307AC-C459-42D0-A890-5F79AA02EADE}#2.1#0"; "PhraseTemplate.ocx"
Begin VB.UserControl ResultEntryCtrl 
   BackColor       =   &H80000016&
   ClientHeight    =   9720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15480
   KeyPreview      =   -1  'True
   ScaleHeight     =   9720
   ScaleWidth      =   15480
   Begin Organ.OrganCtrl OrganCtrl 
      Height          =   495
      Left            =   14280
      TabIndex        =   112
      Top             =   4920
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
   End
   Begin VB.CommandButton AllOkBtn 
      Caption         =   "הכל תקין"
      Height          =   375
      Left            =   13320
      TabIndex        =   111
      Top             =   1560
      Visible         =   0   'False
      Width           =   1815
   End
   Begin PhraseTemplate.PhraseTemplateCtrl PResultPhrase 
      Height          =   330
      Index           =   0
      Left            =   6840
      TabIndex        =   109
      Top             =   8520
      Visible         =   0   'False
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   582
   End
   Begin VB.CommandButton btnPrintFax 
      BackColor       =   &H80000016&
      Caption         =   "Send Fax"
      Enabled         =   0   'False
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
      Left            =   13320
      TabIndex        =   107
      Top             =   3480
      Width           =   1815
   End
   Begin MacabiShared.DockListCtrl DockListCtrl 
      Height          =   8415
      Left            =   1440
      TabIndex        =   87
      Top             =   480
      Visible         =   0   'False
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   14843
   End
   Begin VB.TextBox txtPapHeader 
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   104
      Text            =   "PAP Type"
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CheckBox chkRefCancel 
      Alignment       =   1  'Right Justify
      Caption         =   "בטל שליפת הפניות"
      Height          =   375
      Left            =   2640
      TabIndex        =   103
      Top             =   9240
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox chkCon 
      Caption         =   "Check1"
      Enabled         =   0   'False
      Height          =   240
      Left            =   13680
      TabIndex        =   101
      Top             =   5520
      Width           =   225
   End
   Begin VB.CheckBox chkConsult 
      Caption         =   "Check1"
      Height          =   255
      Left            =   14880
      TabIndex        =   99
      Top             =   5520
      Width           =   255
   End
   Begin VB.Frame fraMaxFreeText 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   4100
      TabIndex        =   74
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton CmdResponseLetter 
      Caption         =   "מכתב תשובה"
      Enabled         =   0   'False
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
      Left            =   13320
      TabIndex        =   96
      Top             =   3960
      Width           =   1815
   End
   Begin VB.CommandButton cmdAdditionalActions 
      Caption         =   "בקשות חוזרות"
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
      Left            =   13320
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   91
      Top             =   4440
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox TxtAuthorizedOn 
      BackColor       =   &H80000016&
      Height          =   405
      Left            =   13320
      Locked          =   -1  'True
      TabIndex        =   89
      Top             =   8520
      Width           =   1935
   End
   Begin VB.TextBox SdgCompleted 
      BackColor       =   &H80000016&
      Height          =   405
      Left            =   13320
      Locked          =   -1  'True
      TabIndex        =   78
      Top             =   7800
      Width           =   1935
   End
   Begin VB.CheckBox chkPrintFax 
      Caption         =   "Print Fax On Authorise"
      Height          =   255
      Left            =   13200
      TabIndex        =   76
      Top             =   9480
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CheckBox chkPrintFinalLetter 
      Caption         =   "Print On Authorise"
      Height          =   255
      Left            =   13200
      TabIndex        =   75
      Top             =   9240
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.PictureBox picHistory 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2040
      ScaleHeight     =   255
      ScaleWidth      =   375
      TabIndex        =   72
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox PropsGeneralSdgAuthorized 
      BackColor       =   &H80000016&
      Height          =   405
      Index           =   0
      Left            =   13320
      Locked          =   -1  'True
      TabIndex        =   68
      Top             =   6120
      Width           =   1935
   End
   Begin VB.TextBox PropsGeneralSdgAuthorized 
      BackColor       =   &H80000016&
      Height          =   405
      Index           =   1
      Left            =   13320
      Locked          =   -1  'True
      TabIndex        =   67
      Top             =   6600
      Width           =   1935
   End
   Begin VB.TextBox PropsGeneralSdgAuthorized 
      BackColor       =   &H80000016&
      Height          =   405
      Index           =   2
      Left            =   13320
      Locked          =   -1  'True
      TabIndex        =   66
      Top             =   7080
      Width           =   1935
   End
   Begin VB.CommandButton btnPrint 
      BackColor       =   &H80000016&
      Caption         =   "Print / PDF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   13320
      TabIndex        =   65
      Top             =   3000
      Width           =   1815
   End
   Begin MSComctlLib.ImageList HistoryImageList 
      Left            =   1440
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      UseMaskColor    =   0   'False
      _Version        =   393216
   End
   Begin VB.CommandButton SummaryButton 
      BackColor       =   &H80000016&
      Caption         =   "Summary"
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
      Left            =   13320
      TabIndex        =   2
      Top             =   2040
      Width           =   1815
   End
   Begin VB.CommandButton CloseButton 
      BackColor       =   &H80000016&
      Caption         =   "Close"
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
      Left            =   13320
      TabIndex        =   3
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton SaveButton 
      BackColor       =   &H80000016&
      Caption         =   "Save"
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
      Left            =   13320
      MaskColor       =   &H80000016&
      TabIndex        =   1
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Frame PapsResultsfra 
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8055
      Left            =   3960
      TabIndex        =   4
      Top             =   480
      Width           =   9000
      Begin MacabiShared.FreeTextTemplateCtrl PFreeTextResult 
         Height          =   855
         Index           =   1
         Left            =   0
         TabIndex        =   70
         Top             =   7320
         Visible         =   0   'False
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   1508
      End
      Begin VB.Frame frmDiagnosis 
         Height          =   375
         Left            =   0
         TabIndex        =   105
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
         Begin VB.CommandButton cmdOrangeDiagnosis 
            BackColor       =   &H000080FF&
            Caption         =   "Diagnosis"
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
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   106
            Top             =   0
            Width           =   1245
         End
      End
      Begin VB.Frame PTestTabfra 
         BackColor       =   &H80000016&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6615
         Index           =   1
         Left            =   -120
         TabIndex        =   5
         Top             =   720
         Width           =   7815
         Begin VB.CheckBox PResultCheck 
            BackColor       =   &H80000016&
            Caption         =   "Check1"
            Height          =   255
            Index           =   0
            Left            =   1400
            TabIndex        =   40
            Top             =   1560
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox PResultText 
            Height          =   285
            Index           =   0
            Left            =   1440
            TabIndex        =   39
            Text            =   "Text1"
            Top             =   2760
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label PResultDesc 
            BackColor       =   &H80000016&
            Caption         =   "Label3"
            Height          =   255
            Index           =   0
            Left            =   1440
            TabIndex        =   41
            Top             =   600
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Line PResultLine 
            Index           =   0
            Visible         =   0   'False
            X1              =   2520
            X2              =   3960
            Y1              =   1080
            Y2              =   1080
         End
      End
      Begin VB.Frame PSummaryfra 
         BackColor       =   &H80000016&
         Caption         =   "Result Summary"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7095
         Left            =   0
         TabIndex        =   42
         Top             =   240
         Visible         =   0   'False
         Width           =   8775
         Begin VB.TextBox PSummaryText 
            BackColor       =   &H80000016&
            Height          =   6735
            Left            =   0
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   43
            Top             =   240
            Width           =   7695
         End
      End
      Begin MSComctlLib.TabStrip PTestTab 
         Height          =   7095
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   9000
         _ExtentX        =   15875
         _ExtentY        =   12515
         MultiRow        =   -1  'True
         ImageList       =   "HistoryImageList"
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin RequestRemark.RequestRemarkCtrl RequestRemarkCtrl 
      Height          =   495
      Left            =   13320
      TabIndex        =   79
      Top             =   4920
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
   End
   Begin MSFlexGridLib.MSFlexGrid gridAliquots 
      Height          =   255
      Left            =   120
      TabIndex        =   95
      Top             =   9000
      Visible         =   0   'False
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   450
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      ScrollBars      =   1
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Snomed.SnomedCtrl SnomedCtrl 
      Height          =   510
      Index           =   0
      Left            =   720
      TabIndex        =   90
      Top             =   9360
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   900
   End
   Begin VB.CheckBox chkQC 
      Caption         =   "Check1"
      Enabled         =   0   'False
      Height          =   315
      Left            =   15120
      TabIndex        =   98
      TabStop         =   0   'False
      Top             =   5760
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame PatientPropsfra 
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      Caption         =   "General"
      Height          =   7575
      Index           =   1
      Left            =   0
      TabIndex        =   8
      Top             =   840
      Width           =   3615
      Begin VB.CommandButton cmd_assutaPdf 
         BackColor       =   &H80000016&
         Caption         =   "מסמכים מצורפים"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   108
         Top             =   4200
         Width           =   2055
      End
      Begin VB.TextBox PropsGeneralPatientName 
         BackColor       =   &H80000016&
         Height          =   405
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   480
         Width           =   2055
      End
      Begin VB.ComboBox cmbPatholog 
         Height          =   336
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   7080
         Width           =   2052
      End
      Begin VB.TextBox PropsGeneralPatientGender 
         BackColor       =   &H80000016&
         Height          =   405
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox PropsGeneralPatientBirth 
         BackColor       =   &H80000016&
         Height          =   405
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   1920
         Width           =   2055
      End
      Begin VB.TextBox PropsGeneralReferring 
         BackColor       =   &H80000018&
         Height          =   405
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   3000
         Width           =   2055
      End
      Begin VB.TextBox PropsGeneralSubmitting 
         BackColor       =   &H80000018&
         Height          =   405
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   3480
         Width           =   2055
      End
      Begin VB.TextBox PropsGeneralSdgPriority 
         BackColor       =   &H80000016&
         Height          =   405
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   4680
         Width           =   2055
      End
      Begin VB.TextBox PropsGeneralSdgDelivery 
         BackColor       =   &H80000016&
         Height          =   405
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   5160
         Width           =   2055
      End
      Begin VB.TextBox PropsGeneralSdgSlides 
         BackColor       =   &H80000016&
         Height          =   405
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   5640
         Width           =   2055
      End
      Begin VB.TextBox PropsGeneralSdgCollection 
         BackColor       =   &H80000016&
         Height          =   405
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   6120
         Width           =   2055
      End
      Begin VB.TextBox PropsGeneralSdgWeek 
         BackColor       =   &H80000016&
         Height          =   405
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   6600
         Width           =   2055
      End
      Begin VB.TextBox PropsGeneralPatientID 
         BackColor       =   &H80000016&
         Height          =   405
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000016&
         Caption         =   "Name:"
         Height          =   405
         Index           =   0
         Left            =   120
         TabIndex        =   32
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000016&
         Caption         =   "Patholog:"
         Height          =   405
         Index           =   12
         Left            =   120
         TabIndex        =   86
         Top             =   7080
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000016&
         Caption         =   "Gender"
         Height          =   405
         Index           =   2
         Left            =   120
         TabIndex        =   44
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000016&
         Caption         =   "ID Number:"
         Height          =   405
         Index           =   1
         Left            =   120
         TabIndex        =   31
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000016&
         Caption         =   "Year of Birth:"
         Height          =   405
         Index           =   3
         Left            =   120
         TabIndex        =   30
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000016&
         Caption         =   "Referring:"
         Height          =   405
         Index           =   4
         Left            =   120
         TabIndex        =   29
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000016&
         Caption         =   "Submitting:"
         Height          =   405
         Index           =   5
         Left            =   120
         TabIndex        =   28
         Top             =   3480
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000016&
         Caption         =   "Priority:"
         Height          =   405
         Index           =   6
         Left            =   120
         TabIndex        =   27
         Top             =   4680
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000016&
         Caption         =   "Date :"
         Height          =   405
         Index           =   7
         Left            =   120
         TabIndex        =   26
         Top             =   5160
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000016&
         Caption         =   "No. of Slides"
         Height          =   405
         Index           =   8
         Left            =   120
         TabIndex        =   25
         Top             =   5640
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000016&
         Caption         =   "Refferer :"
         Height          =   405
         Index           =   9
         Left            =   120
         TabIndex        =   24
         ToolTipText     =   "גורם שולח"
         Top             =   6120
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000016&
         Caption         =   "Preg Week:"
         Height          =   405
         Index           =   10
         Left            =   120
         TabIndex        =   23
         Top             =   6600
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000016&
         Caption         =   "Patient"
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
         Index           =   14
         Left            =   0
         TabIndex        =   22
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000016&
         Caption         =   "Physicians"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   15
         Left            =   0
         TabIndex        =   21
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000016&
         Caption         =   "Request"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   17
         Left            =   0
         TabIndex        =   20
         Top             =   4200
         Width           =   1815
      End
   End
   Begin VB.Frame PatientPropsfra 
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   7575
      Index           =   5
      Left            =   120
      TabIndex        =   34
      Top             =   840
      Width           =   3615
      Begin VB.TextBox PropsReferralDiagnose 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         Height          =   2535
         Index           =   2
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   93
         Top             =   4320
         Width           =   3255
      End
      Begin VB.TextBox PropsReferralDiagnose 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         Height          =   2535
         Index           =   1
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   92
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         Caption         =   "מבצע:"
         Height          =   255
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   62
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         Caption         =   "מפנה:"
         Height          =   255
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   94
         Top             =   3840
         Width           =   1215
      End
   End
   Begin VB.Frame PatientPropsfra 
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   7575
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   3615
      Begin MSFlexGridLib.MSFlexGrid HistoryGrid 
         Height          =   7335
         Left            =   0
         TabIndex        =   71
         Top             =   240
         Width           =   3585
         _ExtentX        =   6324
         _ExtentY        =   12938
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.ListView HistoryList 
         Height          =   30
         Left            =   0
         TabIndex        =   73
         Top             =   7545
         Visible         =   0   'False
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   53
         View            =   3
         LabelEdit       =   1
         SortOrder       =   -1  'True
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "HistoryImageList"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Request"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Snomed"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin MSComctlLib.TabStrip PatientProps 
      Height          =   8055
      Left            =   0
      TabIndex        =   35
      Top             =   480
      Width           =   3850
      _ExtentX        =   6800
      _ExtentY        =   14208
      TabWidthStyle   =   2
      TabFixedWidth   =   3341
      TabMinWidth     =   3341
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "History"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame PatientPropsfra 
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   7575
      Index           =   3
      Left            =   120
      TabIndex        =   33
      Top             =   840
      Width           =   3615
      Begin VB.TextBox PropsPhysRefAddress 
         BackColor       =   &H80000016&
         Height          =   1605
         Left            =   1440
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   55
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox PropsPhysSubAddress 
         BackColor       =   &H80000016&
         Height          =   1605
         Left            =   1440
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   57
         Top             =   5040
         Width           =   2055
      End
      Begin VB.TextBox PropsPhysColAddress 
         BackColor       =   &H80000016&
         Height          =   1605
         Left            =   1440
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   56
         Top             =   2880
         Width           =   2055
      End
      Begin VB.TextBox PropsPhysSubName 
         BackColor       =   &H80000016&
         Height          =   405
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   54
         Top             =   4560
         Width           =   2055
      End
      Begin VB.TextBox PropsPhysColName 
         BackColor       =   &H80000016&
         Height          =   405
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   2400
         Width           =   2055
      End
      Begin VB.TextBox PropsPhysRefName 
         BackColor       =   &H80000016&
         Height          =   405
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000016&
         Caption         =   "Address:"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   51
         Top             =   5040
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000016&
         Caption         =   "Referring Physician:"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   50
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000016&
         Caption         =   "Address:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   49
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000016&
         Caption         =   "Collection St.:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   48
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000016&
         Caption         =   "Address:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   47
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000016&
         Caption         =   "Submitting Physician:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   46
         Top             =   4560
         Width           =   1695
      End
   End
   Begin VB.Frame PatientPropsfra 
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   7575
      Index           =   4
      Left            =   120
      TabIndex        =   37
      Top             =   840
      Width           =   3615
      Begin VB.TextBox PropsPatientAddress 
         BackColor       =   &H80000016&
         Height          =   1605
         Left            =   960
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   61
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox PropsPatientName 
         BackColor       =   &H80000016&
         Height          =   405
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   60
         Top             =   120
         Width           =   2535
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000016&
         Caption         =   "Address:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   59
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000016&
         Caption         =   "Name:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   58
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.CommandButton AuthoriseButton 
      BackColor       =   &H80000016&
      Caption         =   "Authorise"
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
      Left            =   13320
      TabIndex        =   36
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox SdgName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   13320
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label lblPayingCustomer 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13320
      TabIndex        =   110
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label11 
      Caption         =   "Con"
      Height          =   255
      Left            =   13320
      TabIndex        =   102
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label Label9 
      Caption         =   "התייעצות"
      Height          =   255
      Left            =   14040
      TabIndex        =   100
      Top             =   5520
      Width           =   855
   End
   Begin VB.Label lblSampleCodeRemark 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3000
      TabIndex        =   97
      Top             =   -45
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label LblAuthorizedOn 
      BackColor       =   &H80000016&
      Caption         =   "Authorized On"
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
      Left            =   13320
      TabIndex        =   88
      Top             =   8280
      Width           =   1815
   End
   Begin VB.Label LblRevisionStatus 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11640
      TabIndex        =   85
      Top             =   -45
      Width           =   1695
   End
   Begin VB.Label lblStatusBar 
      Height          =   255
      Left            =   120
      TabIndex        =   84
      Top             =   8640
      Width           =   8415
   End
   Begin VB.Image ImageRes 
      Height          =   240
      Index           =   0
      Left            =   5640
      Top             =   360
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label LblMaterialValue 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   13140
      TabIndex        =   83
      Top             =   8880
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label LblMaterialTitle 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   ":נשאר חומר"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   13620
      TabIndex        =   82
      Top             =   9000
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label LblTotalLines 
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
      Left            =   12240
      TabIndex        =   81
      Top             =   8640
      Width           =   495
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      Caption         =   "Total Lines:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   10920
      TabIndex        =   80
      Top             =   8640
      Width           =   1230
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000016&
      Caption         =   "Completed by"
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
      Index           =   11
      Left            =   13320
      TabIndex        =   77
      Top             =   7560
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000016&
      Caption         =   "Authorised by"
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
      Index           =   18
      Left            =   13320
      TabIndex        =   69
      Top             =   5880
      Width           =   1815
   End
   Begin VB.Label lblRequestTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   64
      Top             =   -45
      Width           =   6015
   End
   Begin VB.Image imgHistory 
      Height          =   270
      Left            =   2640
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000016&
      Caption         =   "QC:"
      Height          =   255
      Left            =   15000
      TabIndex        =   63
      Top             =   5520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image SdgStatusImage 
      Height          =   270
      Left            =   240
      Stretch         =   -1  'True
      Top             =   120
      Width           =   270
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H80000016&
      Caption         =   "מספר פנימי:"
      Height          =   240
      Left            =   13320
      RightToLeft     =   -1  'True
      TabIndex        =   38
      Top             =   0
      Width           =   1815
   End
   Begin VB.Menu MnSnomed 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu MnShowSnomed 
         Caption         =   "Show Snomed"
      End
   End
End
Attribute VB_Name = "ResultEntryCtrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


' +---------------------------------------------------------------------------+
' | ResultEntryCtrl
' |
' | Modifaction history:
' |
' +------------+--------+-----------------------------------------------------+
' | Date       | By     | Description
' +------------+--------+-----------------------------------------------------+
' | DD-MM-YYYY |        | Created
' | 19-05-2005 | Barak  | Added patholog combo
' | 01-08-2011 | Yonatan| Added Print Fax button (request 769)
' | 02-08-2011 | Yonatan| Added multiple PAP slides validation (request 782)
' +------------+--------+-----------------------------------------------------+

Implements LSExtensionWindowLib.IExtensionWindow
Implements LSExtensionWindowLib.IExtensionWindow2

Private Const RESULTSFRAMEWIDTH = 9300
Private Const MALIGNANT_REQUEST = &HFFC0C0

Private ProcessXML As LSSERVICEPROVIDERLib.NautilusProcessXML
Private NtlsCon As LSSERVICEPROVIDERLib.NautilusDBConnection
Private NtlsSite2 As LSExtensionWindowLib.IExtensionWindowSite2
Private NtlsUser As LSSERVICEPROVIDERLib.NautilusUser
Private BarcodeField As String
Private OriginalFreeTextRes() As String
Private DrOnlyIds As String
Private doctorOnly As Boolean


Private sp As LSSERVICEPROVIDERLib.NautilusServiceProvider

'_________________________________________________
'pat002
Private Const PapLbcAliquotWF = "PAPS LBC Aliquot"
Private Const PAP_TEST_CODE_MEDICAL = "8146"
Private Const PAP_LBC_TEST_CODE_MEDICAL = "81460"
Private Const PAP_LBC_TEST_CODE = "81490"
Private Const PAP_SMEAR_HEADER = "PAP Smear"
Private Const PAP_LBC_HEADER = "PAP LBC"
Private Const GREEN = &HC000&
Private Const PINK = &HFF00FF

'__________________________________________________
'___________________________

'-------------------------- GLOBAL SEMAPHORE ---------------------------------

'a handle to the requests's semaphore:
'it is used to prevent the working on the
'same request from different workstations;
'could be:
'1. empty - no open request
'2. handle value - a request is open
Dim strHandle As String

'-------------------------- LOCAL SEMAPHORE -----------------------------------

'data related to the local semaphore:
'prevent opening more than one instance of the application per workstation

'the id of the relevant mutex:
Private lMutexHandle As Long

'a type needed to be used in calling the mutex function:
Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

'functions to operate the mutex:
Private Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" _
    (lpMutexAttributes As SECURITY_ATTRIBUTES, ByVal bInitialOwner As Boolean, _
    ByVal lpName As String) As Long
'Private Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (lpMutexAttributes As SECURITY_ATTRIBUTES, ByVal bInitialOwner As Long, ByVal lpName As String) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As _
    Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function ReleaseMutex Lib "kernel32" (ByVal hMutex As Long) As _
    Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As _
    Long
Private Declare Function LoadKeyboardLayout Lib "user32" Alias _
    "LoadKeyboardLayoutA" (ByVal pwszKLID As String, ByVal Flags As Long) As Long
'Private Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Integer
'Private Declare Function GlobalDeleteAtom Lib "kernel32" (ByVal nAtom As Integer) As Integer
'Private Declare Function GlobalFindAtom Lib "kernel32" Alias "GlobalFindAtomA" (ByVal lpString As String) As Integer
'Private Const MyAtomName As String = "REorDiag"
Private CurrAtom As Integer
Private Const ERROR_ALREADY_EXISTS = 183&
'----------------------------------------------------------------------------------

Private Con As ADODB.connection
Private Sdg As ADODB.Recordset
Private Referring As ADODB.Recordset
Private Aliquots As ADODB.Recordset
Private Implement As ADODB.Recordset
Private Patient As ADODB.Recordset
Private Results As ADODB.Recordset
Private History As ADODB.Recordset
Private SnomedMCalculation As ADODB.Recordset
Private SnomedTCalculation As ADODB.Recordset
Private SampleCodes As ADODB.Recordset
Private OrderAndCostumer As ADODB.Recordset




Private Operaqtor As ADODB.Recordset

Private Role As ADODB.Recordset
Private InspectionLog As ADODB.Recordset
Private PropsCurFrame As Integer
Private TestCurFrame As Integer
Private PResultIndex As Integer
Private PResultCheckIndex As Integer
Private PResultTextIndex As Integer
Private PResultPhraseIndex As Integer
Private PFreeTextResultIndex As Integer
Private PSnomedIndex As Integer
Private OpenedRequest As Boolean
Private PrintFaxResult As Dictionary
Private PrintFax As Boolean
Private PathologCodes As Scripting.Dictionary
Private PathologCoredNumberToName As Scripting.Dictionary
Private Ref As Referrals.Referral

'isMicroTextSaved
Private IsMicroTextSaved As Boolean

Private VisibleResultTab As Dictionary

Const csHeBrEw As String = "iso-8859-8" ' Hebrew character set

Private Const PTestTabFraHeight = 6615
Private Const PTestTabHeight = 7095

Private Const nInch = 1440
Private Const MAX_WIDTH = 14535
Private Const RED = &HFF&
Private Const BLACK = &H80000008

Private Inspect As Boolean
Private PQCParameter As Integer
Private CQCParameter As Integer
Private HQCParameter As Integer
Private QcRank As Integer

Private MandatoryExists As Boolean
Private FreeTextNormalSize As Boolean
Private FreeTextNormalTop As Long
Private FreeTextNormalLeft As Long
Private FreeTextNormalHeight As Long
Private FreeTextNormalWidth As Long
Private FreeTextContainer As Object

Private AuthoriseButtonFlag As Boolean
Private CurrFreeTextIndx As Integer

Private WorkFolder As String

Private sdg_log As New SdgLog.CreateLog
Private sdg_log_desc As String

Private didShowPDFError As Boolean

Private dicResultIdToName As New Dictionary


Public RunFromWindow As Boolean
Public Event CloseClicked()

Dim WithEvents SnomedParser As Snomed.Parser
Attribute SnomedParser.VB_VarHelpID = -1

Type tagPOINT
    X As Long
    Y As Long
End Type

Private Debugging As Boolean
Private TamplatePFResInsex As Integer


'tells if we should read from the RTF_RESULT_BACKUP table;
' 0 - ask the user
' 1 - read from backup
' 2 - do not read
' 3 - do not read but do not delete backup record
Private nRTFResultBackup As Integer

'Private Const WM_USER = &H400
'Private Const EM_SETSCROLLPOS = WM_USER + 222
'Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Function faIndex(row As Integer, col As Integer) As Long
50        On Error GoTo ErrHnd
60        faIndex = row * HistoryGrid.Cols + col
70        Exit Function
ErrHnd:
80        Call ErrHandler("faIndex")
          
End Function

Private Sub AllOkBtn_Click()
90     If NtlsUser.GetWorkstationName = "ONE1PC1517" Then
100   Call sdg_log.InsertLog(Sdg("SDG_ID"), "DEBUG1", sdg_log_desc)
110    End If
120   On Error GoTo ErrHnd
         'copy all results from the default results to the request aliquot
         Dim sql As String
         
          Dim rst As New ADODB.Recordset
           Dim TamplateResults As New ADODB.Recordset
         'see if there are harshaot
130       Set rst = Con.Execute("select * from LIMS_SYS.LIMS_GROUP g, LIMS_SYS.OPERATOR_GROUP og " & _
          " where og.OPERATOR_ID= " & NtlsUser.GetOperatorId & "  and og.GROUP_ID = g.GROUP_ID and g.NAME='Pap Quick Answer' ")
140        If rst.EOF Then
150           MsgBox ("Error. You need to be in 'Pap Quick Answer' group to use this key.")
160           Exit Sub
170        End If
180        If Not (Sdg("status") = "P" Or Sdg("status") = "V") Then
190           MsgBox ("Error. The SDG needs to be in status 'Received' or 'In Progress' to use this key.")
200             Exit Sub
210        End If
220      If NtlsUser.GetWorkstationName = "ONE1PC1517" Then
230           Call sdg_log.InsertLog(Sdg("SDG_ID"), "DEBUG2", sdg_log_desc)
240      End If

'190     Set TamplateResults = Con.Execute("select result.result_id, result.name result_name, " & _
'          "u_result_desc_user.u_bold,u_result_desc_user.u_height, u_result_desc_user.u_width,  " & _
'          "u_result_desc_user.u_visible,u_result_desc_user.u_read_only,u_result_desc_user.u_label," _
'          & _
'          "u_result_desc_user.u_free_text_template, u_result_desc_user.u_template_name,  " & _
'          "u_result_desc_user.u_type,u_result_desc_user.u_rtl,u_result_desc_user.u_phrase_list," _
'          & _
'          "u_result_desc_user.u_needs_review, u_result_desc_user.u_print_fax,  " & _
'          "u_result_desc_user.u_font_color, test.name test_name, result.status, result.description," _
'          & _
'          "test.priority, u_result_desc_user.u_order,u_result_desc_user.u_renk,  " & _
'          "formatted_result, test_template.amount_used " _
'          & _
'          "from lims_sys.result, lims_sys.result_user, lims_sys.test, lims_sys.aliquot, lims_sys.sample, " & _
'          "lims_sys.sdg, lims_sys.u_result_desc_user, lims_sys.result_template, lims_sys.test_template " _
'          & "where result.test_id = test.test_id " & _
'          "and result.result_id = result_user.result_id " & _
'          " and test.aliquot_id = aliquot.aliquot_id " & _
'        " and aliquot.sample_id = sample.sample_id " & _
'            " and result.result_template_id = result_template.result_template_id " & _
'            " and result_template.name = u_result_desc_user.u_template_name " & _
'            " and test_template.test_template_id = test.test_template_id " & _
'            " and sample.sdg_id =sdg.sdg_id " & _
'             " and sdg.name = 'P000001/17' " & _
'             " and test.priority > 0 ")

250   SdgName.Text = Sdg("NAME")
Dim opId As String
260   opId = NtlsUser.GetOperatorId

'190      Call UnloadRequest
'200        Call LoadRecordsets
'220   sql = sql & "  u_result_desc_user.u_bold, "
'230   sql = sql & "  u_result_desc_user.u_height, "
'240   sql = sql & "  u_result_desc_user.u_width, "
'250   sql = sql & "  u_result_desc_user.u_visible, "
'260   sql = sql & "  u_result_desc_user.u_read_only, "
'270   sql = sql & "  u_result_desc_user.u_label, "
'280   sql = sql & "  u_result_desc_user.u_free_text_template, "
'290   sql = sql & "  u_result_desc_user.u_template_name, "
'300   sql = sql & "  u_result_desc_user.u_type, "
'310   sql = sql & "  u_result_desc_user.u_rtl, "
'320   sql = sql & "  u_result_desc_user.u_phrase_list, "
'330   sql = sql & "  u_result_desc_user.u_needs_review, "
'340   sql = sql & "  u_result_desc_user.u_print_fax, "
'350   sql = sql & "  u_result_desc_user.u_font_color, "


'420   sql = sql & "  NVL(result.formatted_result,tr.formatted_result) as formatted_result, "
'430   sql = sql & "  NVL(result.original_result,tr.original_result) as original_result, "

'560   sql = sql & "  lims_sys.u_result_desc_user, "
'570   sql = sql & "  lims_sys.result_template, "
'580   sql = sql & "  lims_sys.test_template "

'630   sql = sql & "AND result.result_template_id      = result_template.result_template_id "
'640   sql = sql & "AND result_template.name           = u_result_desc_user.u_template_name "
'650   sql = sql & "AND test_template.test_template_id = test.test_template_id "
'300   sql = sql & "  u_result_desc_user.u_order, "
'310   sql = sql & "  u_result_desc_user.u_renk "
'340   sql = sql & "  test_template.amount_used "

270    If NtlsUser.GetWorkstationName = "ONE1PC1517" Then
280   Call sdg_log.InsertLog(Sdg("SDG_ID"), "DEBUG3", sdg_log_desc)
290    End If
'''''royyy change to function
300   sql = "select LIMS.SET_DEFAULT_FOR_PAP(" & Sdg("SDG_ID") & "," & opId & ") as res from dual"
310   Set TamplateResults = Con.Execute(sql)
320   TamplateResults.MoveFirst
330        If TamplateResults.EOF Then MsgBox ("Error putting results")
     
340        If TamplateResults("res") <> "True" Then
350     MsgBox ("Error putting results in SDG")
360   End If
'sql = " UPDATE"
'sql = sql & "   (SELECT /*+ INDEX(LIMS_SYS.sample FK_SAMPLE_SDG)   INDEX(td AK_SDG_NAME)*/ result.result_id,"
'sql = sql & "     result.modified,"
'sql = sql & "     result.operator_id,"
'sql = sql & "     result.name result_name,"
'sql = sql & "     test.name test_name,"
'sql = sql & "     test.status,"
'sql = sql & "     test.description,"
'sql = sql & "     test.priority,"
'sql = sql & "     result.formatted_result old_formatted_result,"
'sql = sql & "     tr.formatted_result new_formatted_result,"
'sql = sql & "     result.original_result old_original_result,"
'sql = sql & "     tr.original_result new_original_result"
'sql = sql & "   FROM lims_sys.result,"
'sql = sql & "     lims_sys.result_user,"
'sql = sql & "     lims_sys.test,"
'sql = sql & "     lims_sys.aliquot,"
'sql = sql & "     lims_sys.sample,"
'sql = sql & "     lims_sys.result tr,"
'sql = sql & "     lims_sys.result_user tru,"
'sql = sql & "     lims_sys.test tt,"
'sql = sql & "     lims_sys.aliquot ta,"
'sql = sql & "     lims_sys.sample ts,"
'sql = sql & "     lims_sys.sdg td"
'sql = sql & "   WHERE"
'sql = sql & "    sample.sdg_id            =" & Sdg("SDG_ID")
'sql = sql & "   AND td.name                  ='P000554/17'"
'
'sql = sql & "    AND result.test_id         = test.test_id"
'sql = sql & "   AND result.result_id         = result_user.result_id"
'sql = sql & "   AND test.aliquot_id          = aliquot.aliquot_id"
'sql = sql & "   AND aliquot.sample_id        = sample.sample_id"
'sql = sql & "   AND test.priority            > 0"
'sql = sql & "   AND result.status           <> 'X'"
'sql = sql & "   AND test.status             <> 'X'"
'sql = sql & "   AND aliquot.status          <> 'X'"
'sql = sql & "   AND ts.sdg_id                =td.sdg_id"
'sql = sql & "   AND ta.sample_id             =ts.sample_id"
'sql = sql & "   AND tt.aliquot_id            =ta.aliquot_id"
'sql = sql & "   AND tr.test_id               =tt.test_id"
'sql = sql & "   AND tru.result_id            =tr.result_id"
'sql = sql & "   AND tr.status               <> 'R'"
'sql = sql & "   AND tr.status               <> 'X'"
'sql = sql & "   AND tt.name                  =test.name"
'sql = sql & "   AND tr.name                  =result.name"
'sql = sql & "   AND result.formatted_result IS NULL"
'sql = sql & "   AND result.original_result  IS NULL"
'sql = sql & "   AND tr.formatted_result     IS NOT NULL"
'sql = sql & "   AND tr.original_result      IS NOT NULL"
'sql = sql & "   AND result.name             <> 'Sign by Screener'"
'sql = sql & "   )"
'sql = sql & " SET old_formatted_result= new_formatted_result,"
'sql = sql & "   old_original_result   = new_original_result,"
'sql = sql & "   MODIFIED              ='T' ,"
'sql = sql & "   OPERATOR_ID           = " & opId
''810   sql = sql & "ORDER BY test.priority, "
''820   sql = sql & "  u_result_desc_user.u_order"
'710   Set TamplateResults = Con.Execute(sql)
'720   'Call sdg_log.InsertLog(Sdg("SDG_ID"), "DEBUG4", sdg_log_desc)
''840   Set Results = TamplateResults
'730    Con.Execute ("update lims_sys.result " & _
'            "  set formatted_result='" & opId & "'" & _
'            " , original_result='" & opId & "'" & _
'            " , MODIFIED='T'" & _
'            " ,OPERATOR_ID=" & opId & _
'            "  where result_id in  (select/*+ INDEX(LIMS_SYS.sample FK_SAMPLE_SDG)  */ r2.result_id  " & _
'            "        From  LIMS_SYS.sample, LIMS_SYS.aliquot, LIMS_SYS.test ,lims_sys.Result r2 " & _
'            "         Where " & _
'            "         sample.sdg_id  =" & Sdg("SDG_ID") & _
'            "         AND aliquot.sample_id   = sample.sample_id " & _
'            "         AND test.aliquot_id = aliquot.aliquot_id " & _
'            "         AND r2.test_id = test.test_id " & _
'            "         and   r2.name = 'Sign by Screener' )  ")
                 
'200       Results.MoveFirst
'210       If Results.EOF Then Exit Sub
'220        Do Until Results.EOF
'230         TamplateResults.MoveFirst
'240       If TamplateResults.EOF Then Exit Sub
'250        Do Until TamplateResults.EOF
'260       If Results("result_name") = TamplateResults("result_name") And Results("test_name") = TamplateResults("test_name") Then
'270           Results("formatted_result") = TamplateResults("formatted_result")
'280           Exit Do
'290        End If
'300       TamplateResults.MoveNext
'310      Loop
'320       Results.MoveNext
'330      Loop
'340      Results.MoveFirst
'850   Results.MoveFirst
'860        If Results.EOF Then Exit Sub
'870   Do Until Results.EOF
'880      If Not IsNull(Results("formatted_result")) And Results("result_name") <> "Sign by Screener" Then
'890         Con.Execute ("update lims_sys.result " & _
'             "  set formatted_result='" & Results("formatted_result") & "'" & _
'            " , original_result='" & Results("original_result") & "'" & _
'             " , MODIFIED='T'" & _
'            " ,OPERATOR_ID=" & NtlsUser.GetOperatorId & _
'            " where result_id=" & Results("result_id"))
'900     ElseIf Results("result_name") = "Sign by Screener" Then
'910         Con.Execute ("update lims_sys.result " & _
'            "  set formatted_result='" & NtlsUser.GetOperatorId & "'" & _
'            " , original_result='" & NtlsUser.GetOperatorId & "'" & _
'            " , MODIFIED='T'" & _
'            " ,OPERATOR_ID=" & NtlsUser.GetOperatorId & _
'            " where result_id=" & Results("result_id"))
'920      End If
'
'930        Results.MoveNext
'940        Loop
      'Call sdg_log.InsertLog(Sdg("SDG_ID"), "DEBUG5", sdg_log_desc)
370    Call sdg_log.InsertLog(Sdg("SDG_ID"), "RE.SAVE", sdg_log_desc)
380    Call UpdateStatusToC
390   SdgName.Text = Sdg("name")
400     If NtlsUser.GetWorkstationName = "ONE1PC1517" Then
410          Call sdg_log.InsertLog(Sdg("SDG_ID"), "DEBUG5", sdg_log_desc)
420      End If
430   Call SdgName_KeyDown(vbKeyReturn, 0)
440     If NtlsUser.GetWorkstationName = "ONE1PC1517" Then
450   Call sdg_log.InsertLog(Sdg("SDG_ID"), "DEBUG6", sdg_log_desc)
460    End If
470    DoEvents
480     If NtlsUser.GetWorkstationName = "ONE1PC1517" Then
490   Call sdg_log.InsertLog(Sdg("SDG_ID"), "DEBUG8", sdg_log_desc)
500    End If
'860        Call LoadProps

'870       Call LoadHistory
'
'880     LoadResults
            '"and u_result_desc_user.u_order > 0 " _
            '& "and result.status <> 'X' " & _
           ' "and test.status <> 'X' " & "and aliquot.status <> 'X' " & _

510      AllOkBtn.Visible = False
         
520             Exit Sub
ErrHnd:
530          Call ErrHandler("AllOkBtn_Click")
End Sub

Private Sub btnPrintFax_Click()
540       On Error GoTo ErrHnd
550       If Sdg.State = adStateClosed Then
560           MsgBox "Load a request for printing"
570           Exit Sub
580       End If
590       Call TriggerSdgEvent("Type Number and Send Fax", nte(Sdg("sdg_id")))
600       Call SdgName.SetFocus
610       Exit Sub
ErrHnd:
620       Call ErrHandler("btnPrint_Click")
End Sub

Private Sub btnSpecialAuth_Click()
630       On Error GoTo ErrHnd
640       If Not OpenedRequest Then Exit Sub
650       Set frmSpecialAuth.Con = Con
660       frmSpecialAuth.sdgId = CDbl(Sdg("SDG_ID"))
670       frmSpecialAuth.Show vbModal
680       sdg_log_desc = ""
690       Call sdg_log.InsertLog(Sdg("SDG_ID"), "RE.SPEC", sdg_log_desc)
700       Exit Sub
ErrHnd:
710       Call ErrHandler("btnSpecialAuth_Click")
End Sub

Private Sub cmbPatholog_Change()
'MsgBox cmbPatholog.Text
End Sub

Private Sub cmbPatholog_Click()
720   On Error GoTo ERR_cmbPatholog_Click
          
730       If Sdg("STATUS") <> "A" And PathologCodes.Exists(cmbPatholog.Text) _
    Then

740           Call UpdatePatholog
              'MsgBox cmbPatholog.Text
750       End If
          
          
760       Exit Sub
ERR_cmbPatholog_Click:
770   MsgBox "ERR_cmbPatholog_Click" & vbCrLf & Err.Description
End Sub


Private Sub UpdatePatholog()
780   On Error GoTo ERR_UpdatePatholog

          Dim strOldPatholog As String
          Dim strNewPatholog As String

790       strOldPatholog = PathologCoredNumberToName(nte(Sdg("U_PATHOLOG")))
800       strNewPatholog = cmbPatholog.Text
          
810       If strOldPatholog = strNewPatholog Then Exit Sub

820       Call Con.Execute(" update lims_sys.sdg_user " & " set U_PATHOLOG = '" _
    & PathologCodes(strNewPatholog) & "' " & " where sdg_id = " & Sdg("SDG_ID"))
                           
830       Call sdg_log.InsertLog(Sdg("sdg_id"), "PATHOLOG.UPD", "New: " & _
    strNewPatholog & ", Old: " & strOldPatholog)

840       Exit Sub
ERR_UpdatePatholog:
850   MsgBox "ERR_UpdatePatholog" & vbCrLf & Err.Description
End Sub


Private Sub cmd_assutaPdf_Click()
       
860    frmShowAssutaPdf.assutaMacase = nte(Sdg("U_MACASE"))
870    frmShowAssutaPdf.assutaPdfPath = nte(Sdg("U_ATFILENM"))
880    frmShowAssutaPdf.Init
890   If (frmShowAssutaPdf.assutaPdfPath <> "" And frmShowAssutaPdf.IsRead) Then
900       frmShowAssutaPdf.Show vbModal
910       Call sdg_log.InsertLog(Sdg("SDG_ID"), "ResultEntryAttached.Open", _
    frmShowAssutaPdf.IsReadDescription)

920   End If


End Sub

Private Sub cmdAdditionalActions_Click()
930   On Error GoTo ERR_cmdAdditionalActions_Click
       
940       Call frmAdditionalActions.Initialize(Con, Sdg, NtlsUser, sdg_log, _
    WorkFolder, ProcessXML)
950       Call frmAdditionalActions.Show(vbModal)

960   chkCon.value = IIf(frmAdditionalActions.ConsultStatus, 1, 0)
          
          
970       Call SignalExtraRequest(nte(Sdg("external_reference")))

980   Exit Sub
ERR_cmdAdditionalActions_Click:
990   MsgBox "cmdAdditionalActions_Click" & vbCrLf & Err.Description
End Sub






Private Sub CmdResponseLetter_Click()
1000  On Error GoTo ERR_CmdResponseLetter_Click
          Dim strErr As String

      '    strErr = frmResponseLetter.initialize(Sdg("sdg_id"))
      '    If strErr = "" Then
1010          Call frmResponseLetter.Show(vbModal)
1020          Call frmResponseLetter.Initialize(Sdg("U_PDF_PATH"))

      '    End If
1030  Exit Sub
ERR_CmdResponseLetter_Click:
1040  MsgBox "ERR_CmdResponseLetter_Click" & vbCrLf & Err.Description
End Sub

Private Sub LoalResponseLetter(strSdgId As String)
1050  On Error GoTo ERR_LoalResponseLetter
          
          Dim strErr As String
1060    If IsNull(Sdg("U_PDF_PATH")) Then
          
          
1070         CmdResponseLetter.Enabled = False
1080    Else
1090      strErr = frmResponseLetter.Initialize(Sdg("U_PDF_PATH"))
          
1100      CmdResponseLetter.Enabled = False
1110      If strErr = "" Then
1120          CmdResponseLetter.Enabled = True
1130      ElseIf strErr = "372" And didShowPDFError = False Then
              'an error of missing the program to show PDFs
      '        MsgBox "Error" & vbCrLf & "Missing PDF component"
          
              'show this error only once per opening of the screen:
1140          didShowPDFError = True
1150      End If
1160    End If
1170      Exit Sub
ERR_LoalResponseLetter:
1180  MsgBox "ERR_LoalResponseLetter" & vbCrLf & Err.Description
End Sub



Private Sub DockListCtrl_CloseList()
1190      DockListCtrl.Visible = False
1200      Call SetSelectedTextList
End Sub


'Get the list of selected texts from the DockList;
'The list maps a text phrase to Snomed-M;
'insert that data in the relevant locations:
Private Sub SetSelectedTextList()
1210  On Error GoTo ERR_SetSelectedTextList

          Dim dTextToSnomed As Dictionary
          Dim i As Integer
          Dim iSnomedMIndex As Integer
          
1220      Set dTextToSnomed = DockListCtrl.GetSelectedList()
1230      iSnomedMIndex = GetSnomedMResultIndex
          
1240      For i = 0 To dTextToSnomed.Count - 1

1250          Call AddSnomedMItem(iSnomedMIndex, CStr(dTextToSnomed.Items(i)))
              'Call AddFreeText(CStr(dTextToSnomed.Keys(i)))

1260      Next i

1270      Exit Sub
ERR_SetSelectedTextList:
1280  MsgBox "ERR_SetSelectedTextList" & vbCrLf & Err.Description
End Sub


Private Sub AddSnomedMItem(iSnomedMIndex As Integer, strSnomedM As String)
1290  On Error GoTo ERR_AddSnomedMItem
          
          Dim strSnomedMList As String

1300      strSnomedMList = PResultText(iSnomedMIndex).Text
          
1310      If strSnomedMList = "" Then
1320          strSnomedMList = strSnomedM
1330      ElseIf InStr(1, strSnomedMList, strSnomedM) = 0 Then
1340          strSnomedMList = strSnomedMList & "," & strSnomedM
1350      End If
          
1360      PResultText(iSnomedMIndex).Text = strSnomedMList

1370      Exit Sub
ERR_AddSnomedMItem:
1380  MsgBox "ERR_AddSnomedMItem" & vbCrLf & Err.Description
End Sub

Private Sub AddFreeText(strText As String)
1390  On Error GoTo ERR_AddFreeText


1400      Exit Sub
ERR_AddFreeText:
1410  MsgBox "ERR_AddFreeText" & vbCrLf & Err.Description
End Sub


Private Sub DockListCtrl_DblClick()
1420  On Error GoTo ERR_DockListCtrl_DblClick
          
1430      Call PFreeTextResult(CurrFreeTextIndx).SetFocus
          
          
      'not in usage until the relevant
      'MacabiShared.ocx version is updated
      'ready to retreive the Snomed-M:
      '    Call GetSnomedM
          
1440      Exit Sub
ERR_DockListCtrl_DblClick:
1450  MsgBox "ERR_DockListCtrl_DblClick" & vbCrLf & Err.Description
End Sub


Private Function GetSnomedMResultIndex() As Integer
1460  On Error GoTo ERR_GetSnomedMResultIndex

          Dim i As Integer
          Dim typ, index

1470      GetSnomedMResultIndex = -1

1480      For i = 1 To PResultIndex
              
1490          typ = Mid(PResultDesc(i).Tag, 1, 1)
1500          index = Val(Mid(PResultDesc(i).Tag, 2))
              
1510          If typ = "T" And UCase(PResultDesc(i).DataField) = _
    UCase("Snomed M") Then
                  
1520              GetSnomedMResultIndex = index
1530              Exit Function
                  
1540          End If
              
1550      Next i

1560      Exit Function
ERR_GetSnomedMResultIndex:
1570  MsgBox "ERR_GetSnomedMResultIndex" & vbCrLf & Err.Description
End Function


'could be activated after a text was selected from the DockList;
'gets the Snomed-M associated with the selected sentance;
'adds this Snomed-M to it's result text box on screen, if not already exists;
Private Sub GetSnomedM()
1580  On Error GoTo ERR_GetSnomedM

          Dim i As Integer
          Dim strSnomedM As String
          Dim strSnomedMList As String
          Dim typ, index

1590      strSnomedM = DockListCtrl.GetSnomedM
          
1600      If strSnomedM = "" Then
1610          Exit Sub
1620      End If

1630      For i = 1 To PResultIndex
1640          typ = Mid(PResultDesc(i).Tag, 1, 1)
1650          index = Val(Mid(PResultDesc(i).Tag, 2))
1660          If typ = "T" And UCase(PResultDesc(i).DataField) = _
    UCase("Snomed M") Then
1670              strSnomedMList = PResultText(index).Text
1680              Exit For
1690          End If
1700      Next i
          
1710      If strSnomedMList = "" Then
1720          strSnomedMList = strSnomedM
1730      ElseIf InStr(1, strSnomedMList, strSnomedM) = 0 Then
1740          strSnomedMList = strSnomedMList & "," & strSnomedM
1750      End If
          
1760      PResultText(index).Text = strSnomedMList

1770      Exit Sub
ERR_GetSnomedM:
1780  MsgBox "ERR_GetSnomedM" & vbCrLf & Err.Description
End Sub




Private Sub DockListCtrl_KeyPress(KeyAscii As Integer)
1790      Select Case KeyAscii
              Case vbKeyReturn
1800              Call PFreeTextResult(CurrFreeTextIndx).SetFocus
1810          Case vbKeyEscape
1820              DockListCtrl.Visible = False
1830              Call PFreeTextResult(CurrFreeTextIndx).SetFocus
1840      End Select
End Sub



Private Sub HistoryGrid_DblClick()
1850      On Error GoTo ErrHnd
          Dim OldBarcodeField As String
          'dbl click on request cell
1860      If HistoryGrid.MouseCol = 0 And HistoryGrid.MouseRow <> 0 Then
1870          If HistoryGrid.Rows = 2 And HistoryGrid.TextArray(faIndex(1, 0)) _
    = "" Then Exit Sub
1880          SdgName.Text = HistoryGrid.Text
1890          OldBarcodeField = BarcodeField
1900          BarcodeField = "NAME"
1910          Call SdgName_KeyDown(vbKeyReturn, 0)
          'Call SdgName_KeyUp(vbKeyReturn, 0)
1920          BarcodeField = OldBarcodeField
1930          Exit Sub
1940      End If
          'dbl click on snomed cell
1950      If HistoryGrid.MouseCol = 2 And HistoryGrid.MouseRow <> 0 And _
    HistoryGrid.TextArray(faIndex(HistoryGrid.MouseRow, HistoryGrid.MouseCol)) <> _
    "" Then
1960          SnomedCtrl(0).Left = HistoryGrid.CellLeft + _
    PatientPropsfra(2).Left
1970          SnomedCtrl(0).Top = HistoryGrid.CellTop + PatientPropsfra(2).Top
1980          SnomedCtrl(0).Width = 7000
1990          SnomedCtrl(0).Height = 2500
2000          SnomedCtrl(0).StatusReadWrite = SnomedCtrl(0).CReadOnly

      '        SnomedCtrl(0).Initialize ("select u_snomed from lims_sys.sdg_user,lims_sys.sdg where " & "sdg.sdg_id = sdg_user.sdg_id and sdg.name = '" & HistoryGrid.TextArray(faIndex(HistoryGrid.row, 0)) & "'")

2010          SnomedCtrl(0).Initialize ("select result.ORIGINAL_RESULT " & _
    "from lims_sys.result, lims_sys.test, lims_sys.aliquot, lims_sys.sample, lims_sys.sdg " _
    & "where result.test_id = test.test_id and " & _
    "test.aliquot_id = aliquot.aliquot_id and " & _
    "aliquot.sample_id = sample.sample_id and " & "result.name = 'Snomed T' and " & _
    "sample.sdg_id = sdg.sdg_id and " & "sdg.name = '" & _
    HistoryGrid.TextArray(faIndex(HistoryGrid.row, 0)) & "'")

2020          SnomedCtrl(0).Visible = True
2030          SnomedCtrl(0).SetFocus
2040          Exit Sub
2050      End If
2060      Exit Sub
ErrHnd:
2070      Call ErrHandler("HistoryGrid_DblClick")
End Sub

Private Sub HistoryGrid_MouseDown(Button As Integer, Shift As Integer, X As _
    Single, Y As Single)
2080      On Error GoTo ErrHnd
2090      If Button = vbRightButton Then
2100          If HistoryGrid.MouseCol = 2 And HistoryGrid.MouseRow <> 0 Then
2110              HistoryGrid.row = HistoryGrid.MouseRow
2120              HistoryGrid.col = HistoryGrid.MouseCol
2130              PopupMenu MnSnomed
2140              MnSnomed.Visible = True
2150          End If
2160      End If
2170      Exit Sub
ErrHnd:
2180      Call ErrHandler("HistoryGrid_MouseDown")
End Sub
Private Sub IExtensionWindow2_Close()
End Sub
Public Function IExtensionWindow_CloseQuery() As Boolean


2190      On Error GoTo ErrHnd
      'happens when the user close the window
2200      If OpenedRequest Then
2210          If Not RunFromWindow Then
2220              If Sdg("STATUS") <> "A" Then
2230                  If FreeTextContentChanged Then
2240                      If ExitQueryResult = False Then Exit Function
2250                  ElseIf MsgBox("Are you sure you want to EXIT?", vbYesNo + _
    vbDefaultButton2) = vbNo Then
2260                      SdgName.Text = ""
2270                      Exit Function
2280                  End If
2290              End If
2300          End If
2310      End If
2320      Call UnloadRequest
2330      If Not Role.State = adStateClosed Then Role.Close
2340      Set Sdg = Nothing
2350      Set Referring = Nothing
2360      Set Implement = Nothing
2370      Set Patient = Nothing
2380      Set Results = Nothing
2390      Set History = Nothing
2400      Set Aliquots = Nothing
2410      Set SnomedMCalculation = Nothing
2420      Set SnomedTCalculation = Nothing
2430      Set SampleCodes = Nothing
2440      SnomedCtrl(0).CloseSnomed
2450      IExtensionWindow_CloseQuery = True

2460      Call ReleaseApplicationMutex

2470      Exit Function
ErrHnd:
2480      Call ErrHandler("IExtensionWindow_CloseQuery")
End Function

Public Function IExtensionWindow_DataChange() As _
    LSExtensionWindowLib.WindowRefreshType
2490      On Error GoTo ErrHnd
2500      IExtensionWindow_DataChange = windowRefreshNow
2510      Exit Function
ErrHnd:
2520      Call ErrHandler("IExtensionWindow_DataChange")
End Function

Public Function IExtensionWindow_GetButtons() As _
    LSExtensionWindowLib.WindowButtonsType
2530      On Error GoTo ErrHnd
2540      IExtensionWindow_GetButtons = windowButtonsNone
2550      Exit Function
ErrHnd:
2560      Call ErrHandler("IExtensionWindow_GetButtons")
End Function

Public Sub IExtensionWindow_Internationalise()

End Sub

Public Sub IExtensionWindow_PreDisplay()
2570      On Error GoTo ErrHnd
          Dim row As Integer
2580      row = 1
2590      PapsResultsfra.Width = RESULTSFRAMEWIDTH
2600      row = 2
2610      PTestTab.Width = RESULTSFRAMEWIDTH
2620      row = 2
      '    InPreDisplay = True
          Dim constr As String
2630      row = 3
          Dim HisGridRightX As Double
          Dim HisGridLeftX As Double
          Dim HisGridTopY As Double
          Dim HisGridBottomY As Double
2640      row = 4
2650      FreeTextNormalSize = True
2660      row = 5
2670      Set Sdg = New ADODB.Recordset
2680      Set Referring = New ADODB.Recordset
2690      Set Implement = New ADODB.Recordset
2700      Set Patient = New ADODB.Recordset
2710      Set Results = New ADODB.Recordset
2720      Set History = New ADODB.Recordset
2730      Set Aliquots = New ADODB.Recordset
2740      Set SnomedMCalculation = New ADODB.Recordset
2750      Set SnomedTCalculation = New ADODB.Recordset
2760      Set SampleCodes = New ADODB.Recordset
2770      row = 6
2780      Set PrintFaxResult = New Dictionary
2790      Set VisibleResultTab = New Dictionary
2800      Set Con = New ADODB.connection
2810      row = 7
2820      constr = "Provider=OraOLEDB.Oracle" & ";Data Source=" & _
            NtlsCon.GetServerDetails & ";User ID=" & NtlsCon.GetUsername & ";Password=" & _
             NtlsCon.GetPassword
            If NtlsCon.GetServerIsProxy Then
            constr = "Provider=OraOLEDB.Oracle;Data Source=" & _
            NtlsCon.GetServerDetails & ";User id=/;Persist Security Info=True;"
          End If
    
2830      row = 8
2840      Con.Open constr
2850      row = 9
2860      Con.CursorLocation = adUseClient
2870      row = 10
       
      '    con.Open NtlsCon.GetADOConnectionString
      '    con.CursorLocation = adUseClient

2880      Con.Execute "SET ROLE LIMS_USER"
2890      row = 11
2900      Call ConnectSameSession(CDbl(NtlsCon.GetSessionId))
2910      row = 12
2920      PropsCurFrame = 1
2930      TestCurFrame = 1
          Dim i
2940      row = 13
           'For i = 1 To 5 'PatientPropsfra are just 2, not 5
2950      For i = 1 To 5
2960      row = 140 + i
2970          PatientPropsfra(i).Left = PatientProps.ClientLeft
2980          PatientPropsfra(i).Top = PatientProps.ClientTop
2990          PatientPropsfra(i).Width = PatientProps.ClientWidth
3000          PatientPropsfra(i).Height = PatientProps.ClientHeight
3010          PatientPropsfra(i).Visible = False
3020      Next
3030      row = 15
3040      PatientPropsfra(1).Visible = True
      '    HistoryList.Left = 0 '
      '    HistoryList.Top = 0 '
3050      HistoryList.Width = PatientPropsfra(1).Width '
      '    HistoryList.Height = PatientPropsfra(1).Height '
3060      HistoryGrid.Left = 0
3070      HistoryGrid.Top = 0
3080      HistoryGrid.Width = PatientPropsfra(1).Width
3090      HistoryGrid.row = 0
3100      HistoryGrid.col = 0
3110      HistoryGrid.Text = "Request"
3120      HistoryGrid.col = 1
3130      HistoryGrid.Text = "Date"
3140      HistoryGrid.col = 2
3150      HistoryGrid.Text = "Snomed"
3160      HistoryGrid.Height = PatientPropsfra(1).Height
          'HistoryGrid.ColWidth(-1) = HistoryGrid.Width / HistoryGrid.Cols - 44
3170      HistoryGrid.ColWidth(0) = (HistoryGrid.Width / HistoryGrid.Cols) + 450
3180      HistoryGrid.ColAlignment(0) = flexAlignLeftCenter
3190      HistoryGrid.ColWidth(1) = (HistoryGrid.Width / HistoryGrid.Cols) - 350
3200      HistoryGrid.ColAlignment(1) = flexAlignLeftCenter
3210      HistoryGrid.ColWidth(2) = (HistoryGrid.Width / HistoryGrid.Cols) - 250
3220      HistoryGrid.ColAlignment(2) = flexAlignLeftCenter
3230      HistoryGrid.RowHeightMin = 345
3240      HisGridRightX = PatientPropsfra(2).Left + HistoryGrid.CellLeft + _
    HistoryGrid.CellWidth
3250      HisGridLeftX = PatientPropsfra(2).Left + HistoryGrid.CellLeft
3260      HisGridTopY = HistoryGrid.CellTop
3270      HisGridBottomY = HistoryGrid.CellTop + HistoryGrid.CellHeight
3280      TestCurFrame = 1
3290      PTestTabfra(1).Left = PTestTab.ClientLeft
3300      PTestTabfra(1).Top = PTestTab.ClientTop
3310      PTestTabfra(1).Width = PTestTab.ClientWidth
3320      PTestTabfra(1).Height = PTestTab.ClientHeight
3330      PTestTabfra(1).Visible = True
          
3340      row = 16
3350      Call HistoryImageList.ListImages.Add(, "A 1", _
    LoadPicture("Resource\sdga.ico"))
3360      Call HistoryImageList.ListImages.Add(, "V 1", _
    LoadPicture("Resource\sdgv.ico"))
3370      Call HistoryImageList.ListImages.Add(, "X 1", _
    LoadPicture("Resource\sdgx.ico"))
3380      Call HistoryImageList.ListImages.Add(, "C 1", _
    LoadPicture("Resource\sdgc.ico"))
3390      Call HistoryImageList.ListImages.Add(, "I 1", _
    LoadPicture("Resource\sdgi.ico"))
3400      Call HistoryImageList.ListImages.Add(, "P 1", _
    LoadPicture("Resource\sdgp.ico"))
3410      Call HistoryImageList.ListImages.Add(, "R 1", _
    LoadPicture("Resource\sdgr.ico"))
3420      Call HistoryImageList.ListImages.Add(, "S 1", _
    LoadPicture("Resource\sdgs.ico"))
3430      Call HistoryImageList.ListImages.Add(, "U 1", _
    LoadPicture("Resource\sdgu.ico"))
3440      Call HistoryImageList.ListImages.Add(, "W 1", _
    LoadPicture("Resource\sdgw.ico"))
          'Set imgHistory.Picture = LoadPicture("Resource\Shift Down.ico")

3450      Set imgHistory.Picture = LoadPicture("Resource\Led On.ico")

3460      Set Role = Con.Execute("select * from lims_sys.lims_role " & _
    "where role_id = " & NtlsUser.GetRoleId)

3470      If UCase(Role("NAME")) <> UCase("doctor") And UCase(Role("NAME")) <> _
    UCase("pap inspector") Then
      '        QCtxt.BackColor = &H8000000F
      '        QCtxt.locked = True
3480      End If
      '    InPreDisplay = False

3490      Call RequestRemarkCtrl.InitializeConnection(Con)
3500      Call RequestRemarkCtrl.GetOperatorId(NtlsUser.GetOperatorId)
3510      RequestRemarkCtrl.Visible = False
3520      OrganCtrl.connection = Con
3530      OrganCtrl.OperatorName = NtlsUser.GetOperatorName
3540      OrganCtrl.SessionId = NtlsCon.GetSessionId
3550      OrganCtrl.Visible = False
3560      row = 21
          Dim f As New StdFont
3570      f.name = SaveButton.FontName
3580      f.Size = SaveButton.FontSize
3590      f.Bold = SaveButton.FontBold
3600      OrganCtrl.Font = f
3610      row = 22


3620      Set sdg_log.Con = Con
3630      sdg_log.Session = CDbl(NtlsCon.GetSessionId)
3640      row = 23
3650      SaveFreeTextContent
3660      row = 24

3670      Exit Sub
ErrHnd:
3680      Call ErrHandler("IExtensionWindow_PreDisplay Row = " & row)
End Sub

Public Sub IExtensionWindow_refresh()
'code for refreshing the window
'    Call RefreshWindow
End Sub

Public Sub IExtensionWindow_RestoreSettings(ByVal hKey As Long)

End Sub

Public Function IExtensionWindow_SaveData() As Boolean

End Function

Public Sub IExtensionWindow_SaveSettings(ByVal hKey As Long)

End Sub

Public Sub IExtensionWindow_SetParameters(ByVal parameters As String)
3690      On Error GoTo ErrHnd

          Dim strMain As String
          
3700      strMain = parameters
3710      BarcodeField = getNextStr(strMain, ",")
3720      PQCParameter = getNextStr(strMain, ",")
3730      CQCParameter = getNextStr(strMain, ",")
3740      HQCParameter = getNextStr(strMain, ",")
3750      frmAdditionalActions.strLettersFolder = getNextStr(strMain, ",")

      '    Dim Index As Integer
      '    Index = InStr(1, parameters, ",")
      '    BarcodeField = Mid(parameters, 1, Index - 1)
      '    PQCParameter = Mid(parameters, Index + 1, InStr(Index + 1, parameters, ",") - Index)
      '    Index = InStr(Index + 1, parameters, ",")
      '    CQCParameter = Mid(parameters, Index + 1, InStr(Index + 1, parameters, ",") - Index)
      '    HQCParameter = Mid(parameters, InStr(Index + 1, parameters, ",") + 1)
              
3760      Exit Sub
ErrHnd:
3770      Call ErrHandler("IExtensionWindow_SetParameters")
End Sub

Public Sub IExtensionWindow_SetServiceProvider(ByVal ServiceProvider As Object)
3780      On Error GoTo ErrHnd
          'Dim sp As LSSERVICEPROVIDERLib.NautilusServiceProvider
3790      Set sp = ServiceProvider
3800      Set ProcessXML = sp.QueryServiceProvider("ProcessXML")
3810      Set NtlsCon = sp.QueryServiceProvider("DBConnection")
3820      Set NtlsUser = sp.QueryServiceProvider("User")
3830      Exit Sub
ErrHnd:
3840      Call ErrHandler("IExtensionWindow_SetServiceProvider")
End Sub

Public Sub IExtensionWindow_SetSite(ByVal Site As Object)
3850      On Error GoTo ErrHnd
3860      Set NtlsSite2 = Site
3870      If RunFromWindow Then Exit Sub
3880      NtlsSite2.SetWindowInternalName ("MacabiResultEntry")
3890      NtlsSite2.SetWindowRegistryName ("MacabiResultEntry")
3900      Exit Sub
ErrHnd:
3910      Call ErrHandler("IExtensionWindow_SetSite")
 End Sub

 Public Sub IExtensionWindow_Setup()
           Dim phrase As ADODB.Recordset
           Dim sql As String

3920      On Error GoTo ErrHnd

3930   cmdOrangeDiagnosis.ZOrder (ZOrderConstants.vbSendToBack)
 
          'check is there is already an open instance of Result Entry
          'running on this workstation:
3940      If IsFirstApplicationInstance() = False Then
3950          MsgBox _
    "There is already an open instance of Result Entry or Diagnosis on this station"
3960          Call ReleaseApplicationMutex
3970          If RunFromWindow Then
3980              RaiseEvent CloseClicked
3990          Else
4000              NtlsSite2.CloseWindow
4010          End If
4020      End If

          
4030      OpenedRequest = False

4040  sql = " select phrase_description  "
4050  sql = sql & "from lims_sys.phrase_entry  "
4060  sql = sql & "where phrase_id = (select phrase_id from lims_sys.phrase_header where  "
4070  sql = sql & "    name = 'System Parameters')  "
4080  sql = sql & "and  "
4090  sql = sql & "phrase_name='Doctor Only' "

4100    Set phrase = Con.Execute(sql)
4110    If Not phrase.EOF Then
            Dim DrOnlyCsv As String
4120        DrOnlyCsv = " " & Replace(phrase.Fields(0), ",", " , ") & " "
4130        DrOnlyIds = " " & Replace(phrase.Fields(0), ",", " , ") & " "
4140    End If
4150  doctorOnly = False
4160  If InStr(1, DrOnlyIds, " " & NtlsUser.GetOperatorId & " ") > 0 Then
4170      doctorOnly = True
4180  End If




          'Init the Patholog combo
      '    Set phrase = con.Execute("select phrase_description, phrase_name from lims_sys.phrase_entry " & "where phrase_id = (select phrase_id from lims_sys.phrase_header where " & "name = 'Patholog') " & "order by order_number")

4190      sql = " select o.OPERATOR_ID,"
4200      sql = sql & "  ou.U_HEBREW_NAME"
4210      sql = sql & " from lims_sys.operator o, "
4220      sql = sql & "      lims_sys.operator_user ou,"
4230      sql = sql & "      lims_sys.lims_role r"
4240      sql = sql & " where ou.OPERATOR_ID=o.OPERATOR_ID"
4250      sql = sql & " and   o.ROLE_ID=r.role_id"
4260      sql = sql & " and upper(r.name)='DOCTOR'"
4270      sql = sql & " and ou.U_ORDER > 0 "
4280      sql = sql & " order by ou.U_ORDER, ou.U_HEBREW_NAME"
          
          
4290      Set phrase = Con.Execute(sql)




4300      cmbPatholog.list(0) = "None"
4310      Set PathologCodes = New Scripting.Dictionary
4320      Set PathologCoredNumberToName = New Scripting.Dictionary
4330      phrase.MoveFirst
4340      Do Until phrase.EOF
4350          cmbPatholog.list(cmbPatholog.ListCount) = phrase("U_HEBREW_NAME")
4360          Call PathologCodes.Add(CStr(phrase("U_HEBREW_NAME").value), _
         CStr(phrase("OPERATOR_ID").value))
4370          Call _
    PathologCoredNumberToName.Add(CStr(phrase("OPERATOR_ID").value), _
    CStr(phrase("U_HEBREW_NAME").value))
      '        cmbPatholog.list(cmbPatholog.ListCount) = phrase("PHRASE_DESCRIPTION")
      '        Call PathologCodes.Add(CStr(phrase("PHRASE_DESCRIPTION").Value), CStr(phrase("PHRASE_NAME").Value))
4380          phrase.MoveNext
4390      Loop

4400      Call zLang.English
4410      SdgName.Alignment = vbLeftJustify
4420      SdgName.RightToLeft = False

4430      WorkFolder = ""
4440      WorkFolder = _
    xmlManager.GetDefaultFolderFromWorkStation(NtlsUser.GetWorkstationId, Con)

4450      If Trim(WorkFolder) <> "" Then
4460          xmlManager.XmlFolder = WorkFolder & "\ResultEntry\"
4470      End If
4480      If Not RunFromWindow Then
4490          Call SdgName.SetFocus
4500           If UCase(NtlsCon.GetUsername) = "LIMS_SYS" Then
4510              chkRefCancel.Visible = False
4520              chkRefCancel.value = vbChecked
'4840              If NtlsUser.GetWorkstationName = "XP" Then
'4850                  SdgName.Text = "B022902/10"
'                     ' Debugging = True
'4860                  Call SdgName_KeyDown(vbKeyReturn, 1)
'
'4870              End If
'
'4880              If NtlsUser.GetWorkstationName = "ONE1PC1517" Then
'4890                ''  SdgName.Text = "P008455/17"
'                      'SdgName.Text = "B001301/16"
'                      'Debugging = True
'                     ' Call SdgName_KeyDown(vbKeyReturn, 1)
'
'4900              End If
                      
4530          End If
4540      End If

4550      didShowPDFError = False
           

4560      Exit Sub
ErrHnd:
4570      Call ErrHandler("IExtensionWindow_Setup")
End Sub

Public Function IExtensionWindow_ViewRefresh() As _
    LSExtensionWindowLib.WindowRefreshType
4580      On Error GoTo ErrHnd
4590      IExtensionWindow_ViewRefresh = windowRefreshNone
4600      Exit Function
ErrHnd:
4610      Call ErrHandler("IExtensionWindow_ViewRefresh")
End Function

Private Sub SaveResults()
4620      On Error GoTo ErrHnd
          Dim Xmldoc As New DOMDocument
          Dim Xmlres As New DOMDocument
          Dim XmlELimsReq As IXMLDOMElement
          Dim XmlEResultReq As IXMLDOMElement
          Dim XmlELoad As IXMLDOMElement
          Dim XmlEResultEntry As IXMLDOMElement
          Dim DateFormatSyntax As String
          Dim DocXMLFileName As String
          Dim ResXMLFileName As String
          Dim FileName As String
          Dim i As Long

4630      If Sdg("STATUS") = "A" Then
4640          MsgBox "This request is already authorized" & vbCrLf & _
    "and therefore results can not be changed!"
4650          Exit Sub
4660      End If

4670      If Sdg("STATUS") = "I" Then Call UnauthoriseResults


4680      Set XmlELimsReq = Xmldoc.createElement("lims-request")
4690      XmlELimsReq.setAttribute "version", "1"
4700      Set XmlEResultReq = Xmldoc.createElement("result-request")
4710      XmlEResultReq.setAttribute "version", "1"
4720      Set XmlELoad = Xmldoc.createElement("load")
4730      XmlELoad.setAttribute "entity", "SDG"
4740      XmlELoad.setAttribute "id", Sdg("SDG_ID")
4750      XmlELoad.setAttribute "mode", "entry"

4760       If NtlsUser.GetWorkstationName = "ONE1PC1517" Then
4770      Call sdg_log.InsertLog(Sdg("SDG_ID"), "Save DEBUG4", sdg_log_desc)
4780   End If
          'calculate snomed only for PAPs (05.06.2006)
4790      If Left(Sdg("name"), 1) = "P" Then
4800   If NtlsUser.GetWorkstationName = "ONE1PC1517" Then
4810      Call sdg_log.InsertLog(Sdg("SDG_ID"), "Save DEBUG5", sdg_log_desc)
4820   End If
4830          CalculateSnomeds
4840      End If
4850   If NtlsUser.GetWorkstationName = "ONE1PC1517" Then
4860      Call sdg_log.InsertLog(Sdg("SDG_ID"), "Save DEBUG6", sdg_log_desc)
4870   End If
4880      For i = 1 To PResultTextIndex
4890          Set XmlEResultEntry = Xmldoc.createElement("result-entry")
4900          XmlEResultEntry.setAttribute "result-id", PResultText(i).Tag
4910          XmlEResultEntry.setAttribute "original-result", _
    PResultText(i).Text
4920          Call XmlELoad.appendChild(XmlEResultEntry)
4930      Next i
        
4940      For i = 1 To PResultPhraseIndex

             ' If Trim(PResultPhrase(i).getValue) <> "" Then
4950              Set XmlEResultEntry = Xmldoc.createElement("result-entry")
4960              XmlEResultEntry.setAttribute "result-id", PResultPhrase(i).Tag
4970              XmlEResultEntry.setAttribute "original-result", _
                        PResultPhrase(i).getValue
4980              Call XmlELoad.appendChild(XmlEResultEntry)
            '  End If
4990      Next i
5000      Inspect = False
5010      For i = 1 To PResultCheckIndex
5020          Set XmlEResultEntry = Xmldoc.createElement("result-entry")
5030          XmlEResultEntry.setAttribute "result-id", PResultCheck(i).Tag
5040          XmlEResultEntry.setAttribute "original-result", _
    IIf(PResultCheck(i).value = 1, "T", "F")
5050          If PResultCheck(i).Caption = "T" And PResultCheck(i).value = 1 _
    Then Inspect = True
5060          If PrintFaxResult.Exists(PResultCheck(i).Tag) And _
    PResultCheck(i).value = 1 Then PrintFax = True
5070          Call XmlELoad.appendChild(XmlEResultEntry)
5080      Next i

5090      If PFreeTextResultIndex > 1 Then
5100          For i = 2 To PFreeTextResultIndex
5110              Set XmlEResultEntry = Xmldoc.createElement("result-entry")
5120              XmlEResultEntry.setAttribute "result-id", PFreeTextResult(i).Tag
5130              XmlEResultEntry.setAttribute "original-result", Mid(PFreeTextResult(i).GetContent, 1, 1000)
5140              Call XmlELoad.appendChild(XmlEResultEntry)
5150              Call UpdateRtfResult(PFreeTextResult(i).Tag, _
    PFreeTextResult(i))
5160          Next i
5170      Else
5180          Set XmlEResultEntry = Xmldoc.createElement("result-entry")
5190          XmlEResultEntry.setAttribute "result-id", PFreeTextResult(1).Tag
5200          XmlEResultEntry.setAttribute "original-result", Mid(PFreeTextResult(1).GetContent, 1, 1000)
5210          Call XmlELoad.appendChild(XmlEResultEntry)
5220          Call UpdateRtfResult(PFreeTextResult(1).Tag, PFreeTextResult(1))
5230      End If

5240      If PSnomedIndex > 1 Then
5250          For i = 1 To PSnomedIndex
5260              Set XmlEResultEntry = Xmldoc.createElement("result-entry")
5270              XmlEResultEntry.setAttribute "result-id", SnomedCtrl(i).Tag
5280              XmlEResultEntry.setAttribute "original-result", _
    SnomedCtrl(i).getSnomeds
5290              Call XmlELoad.appendChild(XmlEResultEntry)
5300          Next i
5310      End If

          ' update Patholog
5320   If NtlsUser.GetWorkstationName = "ONE1PC1517" Then
5330      Call sdg_log.InsertLog(Sdg("SDG_ID"), "Save DEBUG 7", sdg_log_desc)
5340   End If
5350      If (cmbPatholog.Text <> "") And (cmbPatholog.ListIndex <> 0) Then
              
5360          Call UpdatePatholog
              'Call con.Execute("update lims_sys.sdg_user set U_PATHOLOG = '" & PathologCodes(cmbPatholog.Text) & "' " & " where sdg_id = " & Sdg("SDG_ID"))
5370      End If
5380   If NtlsUser.GetWorkstationName = "ONE1PC1517" Then
5390      Call sdg_log.InsertLog(Sdg("SDG_ID"), "Save DEBUG7", sdg_log_desc)
5400   End If
5410      DateFormatSyntax = Format(Now(), "yyyymmddhhmmss")
5420      DocXMLFileName = "c:\ResEntryXML\doc_" & Trim(nte(Sdg("SDG_ID"))) & _
    "-" & DateFormatSyntax & ".xml"
5430      ResXMLFileName = "c:\ResEntryXML\res_" & Trim(nte(Sdg("SDG_ID"))) & _
    "-" & DateFormatSyntax & ".xml"

5440      Call XmlEResultReq.appendChild(XmlELoad)
5450      Call XmlELimsReq.appendChild(XmlEResultReq)
5460      Call AddXmlFireEventNode(XmlELimsReq, Xmldoc)
5470      Call Xmldoc.appendChild(XmlELimsReq)

          If UCase(Role("NAME")) = "DEBUG" Then Xmldoc.Save ("c:\result.xml")
       '  Xmldoc.Save (DocXMLFileName)

'5710      If Trim(WorkFolder) <> "" Then
'5720          FileName = "ResultEntry_" & Trim(nte(Sdg("SDG_ID"))) & "_DOC1"
'5730          Call xmlManager.SaveXmlFile(Xmldoc, FileName)
'5740      End If
5480   If NtlsUser.GetWorkstationName = "ONE1PC1517" Then
5490      Call sdg_log.InsertLog(Sdg("SDG_ID"), "Save DEBUG  b4 xml 8", sdg_log_desc)
5500   End If
5510      Call ProcessXML.ProcessXMLWithResponse(Xmldoc, Xmlres)
5520   If NtlsUser.GetWorkstationName = "ONE1PC1517" Then
5530      Call sdg_log.InsertLog(Sdg("SDG_ID"), "Save DEBUG9 after ", sdg_log_desc)
5540   End If
          If UCase(Role("NAME")) = "DEBUG" Then Xmlres.Save ("c:\resultres.xml")
  '       Xmlres.Save (ResXMLFileName)

'5760      If Trim(WorkFolder) <> "" Then
'5770          FileName = "ResultEntry_" & Trim(nte(Sdg("SDG_ID"))) & "_RES1"
'5780          Call xmlManager.SaveXmlFile(Xmlres, FileName)
'5790      End If

5550      Exit Sub

ErrHnd:
5560      Call ErrHandler("SaveResults")
End Sub

Private Sub AddXmlFireEventNode(ByRef e As IXMLDOMElement, ByRef doc As _
    DOMDocument)
5570      On Error GoTo ErrHnd
          Dim xmlLogin As IXMLDOMElement
          Dim xmlSdg As IXMLDOMElement
          Dim element As IXMLDOMElement
          
5580      Set xmlLogin = doc.createElement("login-request")
5590      Call e.appendChild(xmlLogin)
5600      Set xmlSdg = doc.createElement("SDG")
5610      Call xmlLogin.appendChild(xmlSdg)
5620      Set element = doc.createElement("find-by-id")
5630      element.Text = Sdg("SDG_ID")
5640      Call xmlSdg.appendChild(element)
5650      Set element = doc.createElement("fire-event")
5660      element.Text = "Calculate Results"
5670      Call xmlSdg.appendChild(element)
5680      Exit Sub
ErrHnd:
5690      Call ErrHandler("AddXmlFireEventNode")
End Sub

Private Sub AuthoriseResults(NewStatus As String)
5700      On Error GoTo ErrHnd
          Dim Xmldoc As New DOMDocument
          Dim Xmlres As New DOMDocument
          Dim XmlELimsReq As IXMLDOMElement
          Dim XmlELoginReq As IXMLDOMElement
          Dim XmlESdg As IXMLDOMElement
          Dim XmlEFind As IXMLDOMElement
          Dim XmlEStatus As IXMLDOMElement
          Dim FileName As String
              
          Dim SdgStatus As ADODB.Recordset
5710      Set SdgStatus = _
    Con.Execute("select status from lims_sys.sdg, lims_sys.sdg_user where " & _
    "sdg.sdg_id = sdg_user.sdg_id and sdg.sdg_id = " & Sdg("SDG_ID"))
          
5720      If NewStatus = "A" And (SdgStatus("STATUS") <> "C" And _
    SdgStatus("STATUS") <> "I") Then
5730          MsgBox "Process is incomplete and therefore cannot be authorise"
5740          SdgStatus.Close
5750          Exit Sub
5760      End If
5770      SdgStatus.Close
          

5780      ChangInspectionForDoctorOnly
          
5790      Set XmlELimsReq = Xmldoc.createElement("lims-request")
      '    XmlELimsReq.setAttribute "version", "1"
5800      Set XmlELoginReq = Xmldoc.createElement("login-request")
5810      Set XmlESdg = Xmldoc.createElement("SDG")
5820      Set XmlEFind = Xmldoc.createElement("find-by-id")
5830      XmlEFind.Text = Sdg("SDG_ID")
5840      Set XmlEStatus = Xmldoc.createElement("STATUS")
5850      XmlEStatus.Text = NewStatus
5860      Call XmlESdg.appendChild(XmlEFind)
5870      Call XmlESdg.appendChild(XmlEStatus)
5880      Call XmlELoginReq.appendChild(XmlESdg)
5890      Call XmlELimsReq.appendChild(XmlELoginReq)
5900      Call Xmldoc.appendChild(XmlELimsReq)

      '    Xmldoc.Save ("c:\auth.xml")
          
5910      If Trim(WorkFolder) <> "" Then
5920          FileName = "ResultEntry_" & Trim(nte(Sdg("SDG_ID"))) & "_DOC2"
5930          Call xmlManager.SaveXmlFile(Xmldoc, FileName)
5940      End If

'5215      Con.Execute("

5950      Call ProcessXML.ProcessXMLWithResponse(Xmldoc, Xmlres)

         ' Xmlres.Save ("c:\authres" & NewStatus & ".xml")

5960      If Trim(WorkFolder) <> "" Then
5970          FileName = "ResultEntry_" & Trim(nte(Sdg("SDG_ID"))) & "_RES2"
5980          Call xmlManager.SaveXmlFile(Xmlres, FileName)
5990      End If
           
'5410      Con.Execute "call lims.sdg_snomed_proc.update_sdg_snomeds('" & _
'    Trim(nte(Sdg("SDG_ID"))) & "')"
'
'
'5420      If CheckIsMalignant(Sdg("SDG_ID")) Then
'5430          Con.Execute _
'    "Update lims_sys.Result r set r.original_result = 'T',r.formatted_result='True' where " _
'    & "r.result_id = (select r1.result_id " & _
'    "from lims_sys.sdg s, lims_sys.sample sa, lims_sys.aliquot a, lims_sys.test t, lims_sys.result r1 " _
'    & "Where s.sdg_id = sa.sdg_id " & "and sa.sample_id = a.sample_id " & _
'    "and a.aliquot_id = t.aliquot_id " & "and t.test_id = r1.test_id " & _
'    "and r1.name = 'Malignant' " & "and s.sdg_id = " & Trim(nte(Sdg("SDG_ID"))) & _
'    ")"
'5440      End If
'
6000      Exit Sub
ErrHnd:
6010      Call ErrHandler("AuthoriseResults")
End Sub


Private Sub CompleteAliquot(AliquotID As String, NewStatus As String)
6020      On Error GoTo ErrHnd
          Dim Xmldoc As New DOMDocument
          Dim Xmlres As New DOMDocument
          Dim XmlELimsReq As IXMLDOMElement
          Dim XmlELoginReq As IXMLDOMElement
          Dim XmlESdg As IXMLDOMElement
          Dim XmlEFind As IXMLDOMElement
          Dim XmlEStatus As IXMLDOMElement
          Dim FileName As String
              
          
6030      Set XmlELimsReq = Xmldoc.createElement("lims-request")
      '    XmlELimsReq.setAttribute "version", "1"
6040      Set XmlELoginReq = Xmldoc.createElement("login-request")
6050      Set XmlESdg = Xmldoc.createElement("ALIQUOT")
6060      Set XmlEFind = Xmldoc.createElement("find-by-id")
6070      XmlEFind.Text = AliquotID
6080      Set XmlEStatus = Xmldoc.createElement("STATUS")
6090      XmlEStatus.Text = NewStatus
6100      Call XmlESdg.appendChild(XmlEFind)
6110      Call XmlESdg.appendChild(XmlEStatus)
6120      Call XmlELoginReq.appendChild(XmlESdg)
6130      Call XmlELimsReq.appendChild(XmlELoginReq)
6140      Call Xmldoc.appendChild(XmlELimsReq)

      '    Xmldoc.Save ("c:\auth.xml")
          
6150      If Trim(WorkFolder) <> "" Then
6160          FileName = "ResultEntry_UpdateRevStat" & Replace(Sdg("name"), "/", "_") & "_" & AliquotID & "_DOC2"
6170          Call xmlManager.SaveXmlFile(Xmldoc, FileName)
6180      End If

'5215      Con.Execute("

6190      Call ProcessXML.ProcessXMLWithResponse(Xmldoc, Xmlres)

         ' Xmlres.Save ("c:\authres" & NewStatus & ".xml")

6200      If Trim(WorkFolder) <> "" Then
6210          FileName = "ResultEntry_UpdateRevStat" & Replace(Sdg("name"), "/", "_") & "_" & AliquotID & "_RES2"
6220          Call xmlManager.SaveXmlFile(Xmlres, FileName)
6230      End If
           
'5410      Con.Execute "call lims.sdg_snomed_proc.update_sdg_snomeds('" & _
'    Trim(nte(Sdg("SDG_ID"))) & "')"
'
'
'5420      If CheckIsMalignant(Sdg("SDG_ID")) Then
'5430          Con.Execute _
'    "Update lims_sys.Result r set r.original_result = 'T',r.formatted_result='True' where " _
'    & "r.result_id = (select r1.result_id " & _
'    "from lims_sys.sdg s, lims_sys.sample sa, lims_sys.aliquot a, lims_sys.test t, lims_sys.result r1 " _
'    & "Where s.sdg_id = sa.sdg_id " & "and sa.sample_id = a.sample_id " & _
'    "and a.aliquot_id = t.aliquot_id " & "and t.test_id = r1.test_id " & _
'    "and r1.name = 'Malignant' " & "and s.sdg_id = " & Trim(nte(Sdg("SDG_ID"))) & _
'    ")"
'5440      End If
'
6240      Exit Sub
ErrHnd:
6250      Call ErrHandler("AuthoriseResults")
End Sub

Private Sub OrganCtrl_Click()
6260      If Sdg("status") = "A" Then
6270          MsgBox "Request is authorised. Can't change Organ."
6280      End If
6290      SetOrgansSnomedT
         
End Sub



Private Sub PropsReferralDiagnose_DblClick(index As Integer)
      'hila-old code
      '    If PropsReferralDiagnose(Index).Tag <> "" Then
      '        Ref.ShowReferral (CInt(PropsReferralDiagnose(Index).Tag))
      '    End If
      'hila- end old code
      'new code
6300      If PropsReferralDiagnose(index).Tag <> "" And _
    (Right(nte(Sdg("EXTERNAL_REFERENCE")), 1) <> "B") Then 'hila- cancel the use of _
    ref in case Histology "B"
6310          Ref.ShowReferral (CInt(PropsReferralDiagnose(index).Tag))
6320      End If
      'hila-end new code
End Sub

Private Sub SdgName_GotFocus()
      'change Language to english
6330      LoadKeyboardLayout "00000409", 1
          
         
End Sub


Private Sub SnomedParser_PhraseResult(ResultName As String, Operator As String, _
    value As String, result As String)
          Dim i
          Dim typ, index
          Dim rValue As String
6340      On Error GoTo ErrHnd
6350      result = "FALSE"
6360      For i = 1 To PResultIndex
6370          typ = Mid(PResultDesc(i).Tag, 1, 1)
6380          index = Val(Mid(PResultDesc(i).Tag, 2))
6390          If typ = "B" And UCase(PResultDesc(i).DataField) = ResultName Then
6400              rValue = PResultCheck(index).value
6410              If Operator = SnomedParser.EqOperator Then
6420                  If (rValue = 1 And value = "T") Or (rValue = 0 And value _
    = "F") Then
6430                      result = "TRUE"
6440                  End If
6450              ElseIf Operator = SnomedParser.NotEqOperator Then
6460                  If (rValue = 0 And value = "T") Or (rValue = 1 And value _
    = "F") Then
6470                      result = "TRUE"
6480                      Exit Sub
6490                  End If
6500              End If
6510              Exit Sub
6520          ElseIf typ = "F" And UCase(PResultDesc(i).DataField) = ResultName _
    And Operator = SnomedParser.ConOperator Then
6530              If InStr(1, UCase(PFreeTextResult(index).GetContent), value) _
    Then
6540                  result = "TRUE"
6550                  Exit Sub
6560              End If
6570              Exit Sub
6580          End If
6590      Next i
6600      Exit Sub
ErrHnd:
6610      Call ErrHandler("SnomedParser_PhraseResult")
End Sub

 



Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
          Dim strVer As String

6620      If KeyCode = vbKeyF10 And Shift = 1 Then
6630          strVer = "Name: " & App.EXEName & vbCrLf & vbCrLf & "Path: " & _
    App.Path & vbCrLf & vbCrLf & "Version: " & "[" & App.Major & "." & App.Minor & _
    "." & App.Revision & "]" & vbCrLf & vbCrLf & _
    "Company: One Software Technologies (O.S.T) Ltd."
6640          chkRefCancel.Visible = True
6650          chkRefCancel.value = vbChecked
6660          MsgBox strVer, vbInformation, "Nautilus - Project Properties"
6670          Call SdgName.SetFocus
             
              
6680      End If
End Sub

Private Sub UnauthoriseResults()
6690      On Error GoTo ErrHnd
          Dim i
          
6700      If Sdg("STATUS") <> "A" And Sdg("STATUS") <> "I" Then Exit Sub
          
6710      Con.BeginTrans

6720      Con.Execute ("update lims_sys.sdg set status = 'C' where sdg_id = " & _
    Sdg("SDG_ID"))
6730      Con.Execute ("update lims_sys.sample set status = 'C' where sdg_id = " _
    & Sdg("SDG_ID"))
6740      Con.Execute _
    ("update lims_sys.aliquot set status = 'C' where sample_id in " & _
    "(select sample_id from lims_sys.sample where sdg_id = " & Sdg("SDG_ID") & ")")
6750      Con.Execute _
    ("update lims_sys.test set status = 'C' where aliquot_id in " & _
    "(select aliquot_id from lims_sys.aliquot where sample_id in " & _
    "(select sample_id from lims_sys.sample where sdg_id = " & Sdg("SDG_ID") & "))")
6760      Con.Execute _
    ("update lims_sys.result set status = 'C' where test_id in " & _
    "(select test_id from lims_sys.test where aliquot_id in " & _
    "(select aliquot_id from lims_sys.aliquot where sample_id in " & _
    "(select sample_id from lims_sys.sample where sdg_id = " & Sdg("SDG_ID") & _
    "))) ")
6770      Con.CommitTrans
6780      Exit Sub
ErrHnd:
6790      Call ErrHandler("UnauthoriseResults")
End Sub

Private Sub MnShowSnomed_Click()
6800      SnomedCtrl(0).Left = HistoryGrid.CellLeft + PatientPropsfra(2).Left
6810      SnomedCtrl(0).Top = HistoryGrid.CellTop + PatientPropsfra(2).Top
6820      SnomedCtrl(0).Width = 4335
6830      SnomedCtrl(0).Height = 2500
6840      SnomedCtrl(0).StatusReadWrite = SnomedCtrl(0).CReadOnly
6850      SnomedCtrl(0).Initialize ("select nvl(result.ORIGINAL_RESULT,'') " & _
    "from lims_sys.result, lims_sys.test, lims_sys.aliquot, lims_sys.sample, lims_sys.sdg " _
    & "where result.test_id = test.test_id and " & _
    "test.aliquot_id = aliquot.aliquot_id and " & _
    "aliquot.sample_id = sample.sample_id and " & "sample.sdg_id = sdg.sdg_id and " _
    & "result.name = 'Snomed T' and " & "sdg.name = '" & _
    HistoryGrid.TextArray(faIndex(HistoryGrid.row, 0)) & "'")
6860      SnomedCtrl(0).Visible = True
6870      SnomedCtrl(0).SetFocus
End Sub


'the FreeTextTemplate control notifies us that
'a record exists in the backup RTF table,
'which means that the last work on this result
'ended in a crash;
'the user can decide to read the backup text:
Private Sub PFreeTextResult_BackupRecordExists(index As Integer)
6880  On Error GoTo ERR_TxtFreeText_BackupRecordExists

6890      If nRTFResultBackup = 0 Then
          
          
              '09.12.2007:
              'always read from a backup record
              'if exists:
               '28-07-1010 -
              'never load records
6900          nRTFResultBackup = 2
          
          
      '        Dim res As VbMsgBoxResult
      '
      '        res = MsgBox(" קיימת רשומת שחזור לתוצאת טקסט חופשי. לשחזר? ", '                      vbYesNoCancel + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading)
      '
      '        Select Case res
      '
      '            Case vbYes
      '                nRTFResultBackup = 1
      '
      '            Case vbNo
      '                nRTFResultBackup = 2
      '
      '            Case vbCancel
      '                nRTFResultBackup = 3
      '
      '        End Select

6910      End If
          

6920      If nRTFResultBackup = 1 Then
                  
6930          Call PFreeTextResult(index).ReadFromBackup
              
6940      End If

6950      Exit Sub
ERR_TxtFreeText_BackupRecordExists:
6960  MsgBox "ERR_TxtFreeText_BackupRecordExists" & vbCrLf & Err.Description
End Sub


Private Sub PFreeTextResult_DblClick(index As Integer)
6970      On Error GoTo ErrHnd
6980      If FreeTextNormalSize Then
6990          Set FreeTextContainer = PFreeTextResult(index).Container
7000          FreeTextNormalTop = PFreeTextResult(index).Top
7010          FreeTextNormalLeft = PFreeTextResult(index).Left
7020          FreeTextNormalHeight = PFreeTextResult(index).Height
7030          FreeTextNormalWidth = PFreeTextResult(index).Width
7040          fraMaxFreeText.Width = PFreeTextResult(index).Width + 50 '11400
7050          fraMaxFreeText.Height = 8415
7060          Set PFreeTextResult(index).Container = fraMaxFreeText
7070          PFreeTextResult(index).Top = 0
7080          PFreeTextResult(index).Left = 0
7090          PFreeTextResult(index).Height = fraMaxFreeText.Height - 50
              'PFreeTextResult(Index).Width = fraMaxFreeText.Width - 50
7100          fraMaxFreeText.Visible = True
7110          FreeTextNormalSize = False
7120          fraMaxFreeText.ZOrder 0
7130          DockListCtrl.ZOrder 0
              'PFreeTextResult(Index).ZOrder 0
7140      Else
7150          Set PFreeTextResult(index).Container = FreeTextContainer
7160          fraMaxFreeText.Width = 255
7170          fraMaxFreeText.Height = 255
7180          PFreeTextResult(index).Top = FreeTextNormalTop
7190          PFreeTextResult(index).Left = FreeTextNormalLeft
7200          PFreeTextResult(index).Height = FreeTextNormalHeight
7210          PFreeTextResult(index).Width = FreeTextNormalWidth
7220          fraMaxFreeText.Visible = False
7230          FreeTextNormalSize = True
7240      End If
7250      Exit Sub
ErrHnd:
7260      Call ErrHandler("PFreeTextResult_DblClick")
End Sub

Private Sub PFreeTextResult_GotFocus(index As Integer)
7270      CurrFreeTextIndx = index
End Sub

Private Sub PFreeTextResult_OnChange(index As Integer)
7280      On Error GoTo ErrHnd
          Dim i As Integer
          Dim MaxLines As Integer

7290      MaxLines = 0
7300      For i = 2 To PFreeTextResultIndex
7310          MaxLines = MaxLines + PFreeTextResult(i).Lines
7320      Next i
7330      LblTotalLines.Caption = MaxLines
7340      Exit Sub
ErrHnd:
7350      Call ErrHandler("PFreeTextResult_OnChange")
End Sub

Private Sub PFreeTextResult_ShowList(index As Integer, ListIndex As Integer)
7360      Call PFreeTextResult(index).AssignList2RTF(DockListCtrl, ListIndex)
7370      DockListCtrl.Visible = True
7380      Call DockListCtrl.SetFocus
End Sub

Private Sub PResultCheck_Click(index As Integer)
7390      On Error GoTo ErrHnd
7400      If Left(Sdg("NAME"), 1) <> "P" Then Exit Sub
          Dim SumBoolRes As Integer

7410      SumBoolRes = _
    CInt(VisibleResultTab(CStr(PResultCheck(index).Container.index)))
7420      If PResultCheck(index).value = 1 Then
7430          SumBoolRes = SumBoolRes + 1
7440          Call _
    VisibleResultTab.Remove(CStr(PResultCheck(index).Container.index))
7450          Call _
    VisibleResultTab.Add(CStr(PResultCheck(index).Container.index), SumBoolRes)
7460          If _
    CInt(VisibleResultTab(CStr(PResultCheck(index).Container.index))) > 0 Then
7470              ImageRes(PResultCheck(index).Container.index).Visible = True
7480          Else
7490              ImageRes(PResultCheck(index).Container.index).Visible = False
7500          End If
7510      Else
7520          SumBoolRes = SumBoolRes - 1
7530          Call _
    VisibleResultTab.Remove(CStr(PResultCheck(index).Container.index))
7540          Call _
    VisibleResultTab.Add(CStr(PResultCheck(index).Container.index), SumBoolRes)
7550          If _
    CInt(VisibleResultTab(CStr(PResultCheck(index).Container.index))) > 0 Then
7560              ImageRes(PResultCheck(index).Container.index).Visible = True
7570          Else
7580              ImageRes(PResultCheck(index).Container.index).Visible = False
7590          End If
7600      End If
          
7610      Call UpdateCheckForSpecialResults(index)
          
7620      Exit Sub
ErrHnd:
7630      Call ErrHandler("PResultCheck_Click")
End Sub

'for special results, the tail of the text is held by the following result,
'that has no check box of itself;
'for showing the right Summary, this code keeps the consistency:
Private Sub UpdateCheckForSpecialResults(index As Integer)
7640  On Error GoTo ERR_UpdateCheckForSpecialResults

      '    If GetResultName(PResultCheck(index).Tag) = "interp_11" Then
          
7650      If dicResultIdToName(PResultCheck(index).Tag) = "interp_11" Then
              
7660          If PResultCheck.Count > index + 1 Then
7670              PResultCheck(index + 1).value = PResultCheck(index).value
7680          End If
              
7690      End If

7700      Exit Sub
ERR_UpdateCheckForSpecialResults:
7710  MsgBox "ERR_UpdateCheckForSpecialResults" & vbCrLf & Err.Description
End Sub


Private Function GetResultName(strResultId As String) As String
7720  On Error GoTo ERR_GetResultName

          Dim rs As Recordset
          Dim sql As String
          
7730      sql = " select r.NAME"
7740      sql = sql & " from lims_sys.result r"
7750      sql = sql & " where r.RESULT_ID='" & strResultId & "'"

7760      Set rs = Con.Execute(sql)
          
7770      If Not rs.EOF Then
7780          GetResultName = rs("NAME")
7790      End If
          
7800      Exit Function
ERR_GetResultName:
7810  MsgBox "ERR_GetResultName" & vbCrLf & Err.Description
End Function



Private Sub RequestRemarkCtrl_StatusChanged(NewStatus As String)
7820      On Error GoTo ErrHnd
7830      If AuthoriseButtonFlag = False Then Exit Sub
          ' status = "P" = non completed
7840      If NewStatus = "P" Then
7850          AuthoriseButton.Enabled = False
7860      ElseIf NewStatus <> "P" Then
7870          AuthoriseButton.Enabled = True
7880      End If
         
7890      Exit Sub
ErrHnd:
7900      Call ErrHandler("RequestRemarkCtrl_StatusChanged")
End Sub

Private Sub OrganCtrl_StatusChanged(NewStatus As String)
7910      On Error GoTo ErrHnd
           Dim mbres As VbMsgBoxResult
           
7920       DoEvents
           
7930       If InStr(1, NewStatus, "+") >= 1 Then
7940          SetOrgansSnomedT
              
7950          If Not IsMicroTextSaved Then
7960              mbres = _
    MsgBox("? בוצע אישור במסך איבר. האם ברצונך לטעון מחדש את מסך הכנסת תוצאות" & _
    vbCrLf & "!שינויים שנעשו לא ישמרו", vbYesNo + vbDefaultButton2 + vbCritical)
7970              If mbres = vbYes Then
          '        If mbres = vbYes And SaveButton.Enabled And SaveButton.Visible Then
          '            Call SaveButton_Click
                      'Will not be notified for another Save
7980                  Call SaveFreeTextContent
7990                  SdgName = Sdg("name")
8000                  Call SdgName_KeyDown(vbKeyReturn, 0)
8010              End If
8020          Else 'IsMicroTextSaved
8030              MsgBox (".שים לב, האיבר הוחלף אך הכותרת והאבחנה לא")
8040          End If
8050      End If
               
         ' If AuthoriseButtonFlag = False Then Exit Sub
          ' status = "P" = non completed
      '    If NewStatus = "P" Then
      '        AuthoriseButton.Enabled = False
      '
      '    ElseIf RequestRemarkCtrl.GetRemarkStatus(RequestRemarkCtrl.SdgName) <> "P" Then
      '        AuthoriseButton.Enabled = True
      '    End If
8060      Exit Sub
ErrHnd:
8070      Call ErrHandler("OrganCtrl_StatusChanged")
End Sub


'Private Sub PFreeTextResult_GotFocus(Index As Integer)
'    Dim p As tagPOINT
'    Dim res As Long
'    p.x = 15
'    p.y = 0
'    res = SendMessage(PFreeTextResult(Index).RTBHandle, EM_SETSCROLLPOS, 0, p)
'End Sub

Private Sub SaveButton_Click()
8080      On Error GoTo ErrHnd
8090   If NtlsUser.GetWorkstationName = "ONE1PC1517" Then
8100      Call sdg_log.InsertLog(Sdg("SDG_ID"), "Save DEBUG1", sdg_log_desc)
8110   End If
8120      If Not OpenedRequest Then Exit Sub
           
8130      If Sdg("STATUS") = "I" And UCase(Role("NAME")) <> UCase("doctor") And _
    (UCase(Role("NAME")) <> UCase("cytoscreener") Or GetInspection <> "CC") And _
    UCase(Role("NAME")) <> UCase("pap inspector") Then
8140          MsgBox _
    "This request is in inspection and waits for a physician or pap inspector authorization!"
8150          Exit Sub
8160      End If
8170   If NtlsUser.GetWorkstationName = "ONE1PC1517" Then
8180      Call sdg_log.InsertLog(Sdg("SDG_ID"), "Save DEBUG2", sdg_log_desc)
8190   End If
8200      If Not IsNull(Sdg("u_is_last_update")) And Sdg("u_is_last_update") = "T" And IsNull(Sdg("completed_by")) Then
8210          Con.Execute ("update lims_sys.sdg set  completed_by=" & NtlsUser.GetOperatorId & ", completed_on=sysdate where sdg_id= " & Sdg("sdg_id"))
8220      End If
8230   If NtlsUser.GetWorkstationName = "ONE1PC1517" Then
8240      Call sdg_log.InsertLog(Sdg("SDG_ID"), "Save DEBUG3", sdg_log_desc)
8250   End If
8260      Call SaveResults
8270   If NtlsUser.GetWorkstationName = "ONE1PC1517" Then
8280      Call sdg_log.InsertLog(Sdg("SDG_ID"), "Save DEBUG end 1", sdg_log_desc)
8290   End If
      '    OpenedRequest = False
 
8300   If NtlsUser.GetWorkstationName = "ONE1PC1517" Then
8310      Call sdg_log.InsertLog(Sdg("SDG_ID"), "Save DEBUG end 2", sdg_log_desc)
8320   End If
          ' This is temporary code for the time until the inlab process is fuly functionning
          '
          '
          '
      'MsgBox 1
8330      If Sdg("STATUS") <> "A" And Sdg("STATUS") <> "I" Then
      '        con.Execute ("update lims_sys.aliquot set status = 'V' " & "where aliquot_id in (select aliquot_id " & "from lims_sys.aliquot a, lims_sys.sample s " & "where a.sample_id = s.sample_id " & "and a.status = 'U' and s.sdg_id = " & Sdg("SDG_ID") & ")")
      'MsgBox 2
      '        con.Execute ("update lims_sys.sdg set status = 'C' " & "where sdg_id = " & Sdg("SDG_ID"))
      'MsgBox 3
      '        con.Execute ("update lims_sys.sample set status = 'C' " & "where sample.status in ('V','P') and sdg_id = " & Sdg("SDG_ID"))
      'MsgBox 4
      '        con.Execute ("update lims_sys.aliquot set status = 'C' " & "where aliquot.status in ('V','P') and sample_id in (select sample_id from lims_sys.sample " & "where sdg_id = " & Sdg("SDG_ID") & ")")
      'MsgBox 5
      '        con.Execute ("update lims_sys.test set status = 'C' " & "where test.status in ('V','P') and aliquot_id in(select a.aliquot_id from lims_sys.aliquot a, " & "lims_sys.sample s where a.sample_id=s.sample_id " & "and s.sdg_id = " & Sdg("SDG_ID") & ")")
      'MsgBox 6
8340         If NtlsUser.GetWorkstationName = "ONE1PC1517" Then
8350      Call sdg_log.InsertLog(Sdg("SDG_ID"), "Save DEBUG change status", sdg_log_desc)
8360   End If
8370      Call ChangeAliquotStatus
8380   If NtlsUser.GetWorkstationName = "ONE1PC1517" Then
8390      Call sdg_log.InsertLog(Sdg("SDG_ID"), "Save DEBUG update to c", sdg_log_desc)
8400   End If
8410      Call UpdateStatusToC
      '        con.Execute ("update lims_sys.result set status = 'C' " & "where result.status = 'V' and test_id in (select t.test_id from lims_sys.test t, " & "lims_sys.aliquot a, lims_sys.sample s " & "where t.aliquot_id = a.aliquot_id " & "and a.sample_id = s.sample_id " & "and s.sdg_id = " & Sdg("SDG_ID") & ")")
                  
                  
              'con.
8420
8430      End If
8440             If NtlsUser.GetWorkstationName = "ONE1PC1517" Then
8450      Call sdg_log.InsertLog(Sdg("SDG_ID"), "Save DEBUG refres window", sdg_log_desc)
8460   End If
8470       Call RefreshWindow
      '    SaveButton.Enabled = False
      '    AuthoriseButton.Enabled = False
8480      gridAliquots.Top = lblStatusBar.Top + lblStatusBar.Height
8490      gridAliquots.Visible = False
          
8500      lblStatusBar.Caption = "Results where saved successfully on " & _
    Format(Now, "hh:mm:ss")
          
8510      If cmbPatholog.Text <> "" And cmbPatholog.Text <> "None" Then
8520          sdg_log_desc = Replace(cmbPatholog.Text, "'", "''")
8530      End If
8540      Call sdg_log.InsertLog(Sdg("SDG_ID"), "RE.SAVE", sdg_log_desc)
8550      Call SaveFreeTextContent
8560  AllOkBtn.Visible = False



8570      Exit Sub
ErrHnd:
8580      Call ErrHandler("SaveButton_Click")
End Sub


Private Sub UpdateStatusToC()
8590  On Error GoTo ERR_UpdateStatusToC
          
8600      Con.Execute ("update lims_sys.result set status = 'C' " & _
    "where result.status = 'V' and test_id in (select t.test_id from lims_sys.test t, " _
    & "lims_sys.aliquot a, lims_sys.sample s " & _
    "where t.aliquot_id = a.aliquot_id " & "and a.sample_id = s.sample_id " & _
    "and s.sdg_id = " & Sdg("SDG_ID") & ")")
          
8610      Exit Sub
ERR_UpdateStatusToC:
8620  MsgBox "ERR_UpdateStatusToC" & vbCrLf & _
    "The results were saved successfully, " & _
    "but the request status was not changed to complete"

      'MsgBox "ERR_UpdateStatusToC" & vbCrLf & Err.Description
End Sub
Private Sub ChangeAliquotStatus()
8630       On Error GoTo ErrHnd
           Dim rs As New ADODB.Recordset
8640       If Aliquots.EOF Then Exit Sub
8650       Do Until Aliquots.EOF
8660           Set rs = _
    Con.Execute("select test_id from lims_sys.test where test.aliquot_id = " & _
    Aliquots("ALIQUOT_ID"))
8670           If rs.EOF Then
8680               Con.Execute _
    ("update lims_sys.aliquot set status='C' where aliquot_id = " & _
    Aliquots("ALIQUOT_ID"))
'8830        Call CompleteAliquot(Aliquots("ALIQUOT_ID"), "C")
8690           End If
8700           rs.Close
8710           Aliquots.MoveNext
8720       Loop
8730       Exit Sub
ErrHnd:
8740       Call ErrHandler("ChangeAliquotStatus")
 End Sub


'get all slides that didn't pass a stain report (station 6):
'd (out) - the slide list
'strSdgId (in) -
Private Sub SlideWithoutStainReport(d As Dictionary, strSgdId As String)
8750  On Error GoTo ERR_SlidesWithoutStainReport

          Dim sql As String
          Dim rs As Recordset
          
          
8760      sql = " select a.NAME"
8770      sql = sql & " from lims_sys.sample s,"
8780      sql = sql & "      lims_sys.aliquot a,"
8790      sql = sql & "   lims_sys.aliquot_user au"
8800      sql = sql & " where s.SDG_ID='" & strSgdId & "'"
8810      sql = sql & " and   a.SAMPLE_ID=s.SAMPLE_ID"
8820      sql = sql & " and   au.ALIQUOT_ID=a.ALIQUOT_ID"
8830      sql = sql & " and   au.U_COLOR_TYPE <> 'רזרבה' "
8840      sql = sql & " and   a.STATUS <> 'X' "
8850      sql = sql & " and   exists"
8860      sql = sql & " ("
8870      sql = sql & "   select 1 "
8880      sql = sql & "   from lims_sys.aliquot_formulation af"
8890      sql = sql & "   where af.CHILD_ALIQUOT_ID=a.ALIQUOT_ID"
8900      sql = sql & " )"
8910      sql = sql & _
    " and ( au.U_ALIQUOT_STATION     is null  or  instr(au.U_ALIQUOT_STATION,     '6')=0 ) "
8920      sql = sql & _
    " and ( au.U_OLD_ALIQUOT_STATION is null  or  instr(au.U_OLD_ALIQUOT_STATION, '6')=0 )"
8930      sql = sql & " order by a.ALIQUOT_ID"
          
          
8940      Set rs = Con.Execute(sql)
          
8950      While Not rs.EOF
8960          Call d.Add(nte(rs("NAME")), "")
          
8970          rs.MoveNext
8980      Wend

8990      Exit Sub
ERR_SlidesWithoutStainReport:
9000  MsgBox "ERR_SlideWithoutStainReport" & vbCrLf & Err.Description
End Sub


Private Function GetSdgStatus(strSdgId As String) As String
9010  On Error GoTo ERR_GetSdgStatus

          Dim rs As Recordset
          Dim sql As String
          
9020      sql = " select d.STATUS"
9030      sql = sql & " from lims_sys.sdg d"
9040      sql = sql & " where d.SDG_ID='" & strSdgId & "'"

9050      Set rs = Con.Execute(sql)
          
9060      If Not rs.EOF Then
9070          GetSdgStatus = nte(rs("STATUS"))
9080      End If

9090      Exit Function
ERR_GetSdgStatus:
9100  MsgBox "ERR_GetSdgStatus" & vbCrLf & Err.Description
End Function

Private Sub UpdateReoprted(strSdgId As String, strValue As String)
9110  On Error GoTo ERR_UpdateReported

          Dim sql As String
          Dim rs As Recordset
           
9120      sql = "  update lims_sys.sdg d "
9130      sql = sql & "  set d.reported = '" & strValue & "' "
9140      sql = sql & "   where d.sdg_id = '" & strSdgId & "' "
      '    sql = sql & "   and d.status = 'A' "
          
9150      Call Con.Execute(sql)
           
9160      Exit Sub
ERR_UpdateReported:
9170  MsgBox "ERR_UpdateReported" & vbCrLf & Err.Description
End Sub


'if there are non reported slides guide the user actions:
'continue in authorisation?
'if yes -
'  create revision?
'  send letter to former version or not?
Private Sub AuthoriseWithRevision()
9180  On Error GoTo ERR_AuthoriseWithRevision

          Dim cg As Revision.CopyGenerator
          
          Dim d As New Dictionary
          Dim shouldRevise As Boolean
          Dim shouldSendLetter As Boolean
          Dim strSdgId As String
          Dim strSdgName As String

9190      shouldRevise = False
9200      shouldSendLetter = False

9210      strSdgId = nte(Sdg("sdg_id"))
9220      strSdgName = nte(Sdg("name"))

9230      Call SlideWithoutStainReport(d, strSdgId)

          'there are non reported slides:
9240      If d.Count > 0 Then

9250          Call frmNonReportedSlides.Initialize(d)
9260          frmNonReportedSlides.Top = 9945 - frmNonReportedSlides.Height '4350
9270          frmNonReportedSlides.Left = 13215 - frmNonReportedSlides.Width '8550
9280          Call frmNonReportedSlides.Show(vbModal)
              

9290          If frmNonReportedSlides.GetAuthorise = False Then
9300              Exit Sub
9310          End If

9320          shouldRevise = frmNonReportedSlides.GetRevise
9330          shouldSendLetter = frmNonReportedSlides.GetSendLetter

9340      End If


9350      Call NewAuthorise
          

9360      If shouldRevise = True And GetSdgStatus(strSdgId) = "A" Then
              
              'create a revision to the authorised SDG:
9370          Set cg = New Revision.CopyGenerator
9380          Call cg.Initialize(sp, strSdgId, "30")
9390          Call cg.Execute
              
              'load the new version:
9400          Call InitiateSdg(strSdgName)
              
              'signal not to send letter for the old version:
9410          If shouldSendLetter = False Then
9420              Call UpdateReoprted(strSdgId, "T")
9430              Call sdg_log.InsertLog(CLng(strSdgId), "REPORTED.TRUE", "")
9440          End If
              
9450      End If

9460      Exit Sub
ERR_AuthoriseWithRevision:
9470   Call ErrHandler("ERR_AuthoriseWithRevision")
End Sub


Private Sub AuthoriseButton_Click()
9480  On Error GoTo ERR_AuthoriseButton_Click
Dim sdgId As String
''9670      'Call sdg_log.InsertLog(Sdg("SDG_ID"), "DEBUG", Erl)
9490      sdgId = nte(Sdg("SDG_ID"))

        'anyone can authorize an SDG
        'if the SDG in an state I, only doctor from "doctor only"  and system can authorise
9500  If Sdg("STATUS") = "I" And UCase(NtlsUser.GetRoleName) <> "SYSTEM" Then
9510      If UCase(NtlsUser.GetRoleName) <> "DOCTOR" Or Not doctorOnly Then
       'check if doctor only
9520        MsgBox "Can't Authorize" & vbCrLf & "The request is in inspection. Only a system operator r an authorised doctors can authorize"
9530        Exit Sub
9540      End If
9550  End If

          ' Validate multiple slides for PAP
9560      If Left(Sdg("NAME"), 1) = "P" Then

9570          If Not multipleSlidesValidation(Sdg("SDG_ID")) Then Exit Sub

9580      End If
'patholab 06/16
'8450      If (Right(nte(Sdg("EXTERNAL_REFERENCE")), 1) = "B" Or _
'    Right(nte(Sdg("EXTERNAL_REFERENCE")), 1) = "C") And _
'    Trim(PResultText(GetSnomedMResultIndex()).Text) = "" Then
'8460          MsgBox "Snomed M value us missing. Request cannot be authorized.", _
'    vbCritical, "Result Entry"
'8470          Exit Sub
'8480      End If


9590      If Not IsNull(Sdg("u_is_last_update")) And Sdg("u_is_last_update") = "T" And IsNull(Sdg("completed_by")) Then
9600          Con.Execute ("update lims_sys.sdg set  completed_by=" & NtlsUser.GetOperatorId & ", completed_on=sysdate where sdg_id= " & Sdg("sdg_id"))

9610      End If

9620      Call ChangeAliquotStatus
 
9630      Call UpdateStatusToC
 
 
9640      If Right(nte(Sdg("EXTERNAL_REFERENCE")), 1) = "B" And _
    nte(Sdg("STATUS")) <> "A" Then
 
9650          Call AuthoriseWithRevision
 
9660      Else
 
9670          Call NewAuthorise
 
9680      End If


9690      Call PrintFaxResult.RemoveAll
9700      Call VisibleResultTab.RemoveAll
9710  DoEvents


9720      Exit Sub
ERR_AuthoriseButton_Click:
9730  Call ErrHandler("ERR_AuthoriseButton_Click")

End Sub

Private Sub NewAuthorise()

9740      On Error GoTo ErrHnd
          Dim save_sdg_id As Long
          Dim frc As FrmRequestConfirm
         
9750      If Not OpenedRequest Then Exit Sub
9760         SaveResults

9770      If Mandatory Then
           ' MsgBox "Missing Mandatory Result/s."
9780        Exit Sub
9790      End If
          'exit is mandatory result is missing:
9800      If Left(Sdg("NAME"), 1) = "P" Then
9810
9820                If ProblemWithPapsResults Then Exit Sub
9830       End If


'8610      If Sdg("STATUS") = "I" And UCase(Role("NAME")) <> UCase("doctor") And _
'    (UCase(Role("NAME")) <> UCase("cytoscreener") Or GetInspection <> "CC") And _
'    UCase(Role("NAME")) <> UCase("pap inspector") Or Sdg("STATUS") = "I" And UCase(Role("NAME")) <> UCase("system") Then
'8620          MsgBox _
'    "This request is in inspection and waits for a physician or pap inspector authorization!"
'8630          Exit Sub
'8640      End If

'       ask user to give password
'8640      frmLogin.Pass = NtlsCon.GetPassword
'8650      frmLogin.txtUserName.Text = NtlsCon.GetUsername
'8660      frmLogin.Show vbModal
'8670      If Not frmLogin.LoginSucceeded Then Exit Sub
'8680      frmLogin.Pass = ""
          
'8700      If Right(Sdg("EXTERNAL_REFERENCE"), 1) = "P" Then

'8690          Set frc = New FrmRequestConfirm
'8700          Set frc.Con = Con
'8710          frc.SdgName = Sdg("NAME")
'8720          Call frc.Show(vbModal)
'8730          If frc.ConfirmSucceeded Then
9840              NewPapAuthorise
'8750          End If
'8760          frc.ConfirmSucceeded = False
'8770          Set frc = Nothing

'8800      Else
'8810          NewCyHyAuthorise
'
'8820      End If


          'do not proceed in the authorisation process
          'if did not authorise because of a missing mandatory result
          '(do not unload the request etc):
'8830      If MandatoryExists = True Then
'8840          Exit Sub
'8850      End If
9850      If Right(Sdg("EXTERNAL_REFERENCE"), 1) = "B" Then
9860          Call UpdateTransferalToTissueArchiveHis(Sdg("SDG_ID"))
9870      Else
9880          Call UpdateTransferalToTissueArchiveCytoPap(Sdg("SDG_ID"))
9890      End If

9900      save_sdg_id = nte(Sdg("SDG_ID"))
'8920      If UCase(Role("NAME")) = UCase("doctor") Then
'8930          Call Con.Execute("update lims_sys.sdg_user set u_qc = '" & QcRank _
'    & "' " & "where sdg_id = " & Sdg("SDG_ID"))
'8940      End If
9910      OpenedRequest = False
9920      Call RefreshWindow
          
'8970      If Sdg("STATUS") = "I" Then
'8980          MsgBox "This request should be rechecked"
'8990      End If
          
      '    If Right(Sdg("EXTERNAL_REFERENCE"), 1) = "P" Then
      '        PapAuthoriseMsg
      '    End If
          
      '    If Sdg("STATUTestCurFrameS") = "A" Then Call TriggerSdgEvent("Print Final Letter")
          'If PrintFax Or chkPrintFax.Value = 1 Then
          
          'print only checked for print (not because result is malignant):
'9000      If chkPrintFax.value = 1 Then
'9010          Call TriggerSdgEvent("Print Fax", nte(Sdg("sdg_id")))
'9020      ElseIf chkPrintFinalLetter.value = 1 Then
'9030          If Sdg("STATUS") = "A" Then Call _
'    TriggerSdgEvent("Print Final Letter", nte(Sdg("sdg_id")))
'9040      End If
9930      sdg_log_desc = AuthorizedBy(nte(Sdg("sdg_id")), 1)
          
9940      Call UnloadRequest
9950      Call zLang.English
9960      SdgName.Alignment = vbLeftJustify
9970      SdgName.RightToLeft = False
9980      Call SdgName.SetFocus

9990      AuthoriseButton.Enabled = False
          'SaveButton.Enabled = False
10000     gridAliquots.Top = lblStatusBar.Top + lblStatusBar.Height
10010     gridAliquots.Visible = False
              
10020     CmdResponseLetter.Enabled = False
          
10030     lblStatusBar.Caption = "Results where authorised successfully on " & _
    Format(Now, "hh:mm:ss")
'9900      Call sdg_log.InsertLog(save_sdg_id, "RE.AUTH", sdg_log_desc)
10040     Call SaveFreeTextContent
          
10050     Exit Sub
ErrHnd:
10060     Call ErrHandler("NewAuthorise")
End Sub

Private Sub Authorise()
10070     On Error GoTo ErrHnd
          Dim save_sdg_id As Long
          Dim frc As FrmRequestConfirm
10080 Call NewAuthorise
10090 Exit Sub
'
'8600      If Not OpenedRequest Then Exit Sub
'
'
'8610      If Sdg("STATUS") = "I" And UCase(Role("NAME")) <> UCase("doctor") And _
'    (UCase(Role("NAME")) <> UCase("cytoscreener") Or GetInspection <> "CC") And _
'    UCase(Role("NAME")) <> UCase("pap inspector") Or Sdg("STATUS") = "I" And UCase(Role("NAME")) <> UCase("system") Then
'8620          MsgBox _
'    "This request is in inspection and waits for a physician or pap inspector authorization!"
'8630          Exit Sub
'8640      End If
'8650      frmLogin.Pass = NtlsCon.GetPassword
'8660      frmLogin.txtUserName.Text = NtlsCon.GetUsername
'8670      frmLogin.Show vbModal
'8680      If Not frmLogin.LoginSucceeded Then Exit Sub
'8690      frmLogin.Pass = ""
'
'8700      If Right(Sdg("EXTERNAL_REFERENCE"), 1) = "P" Then
'
'8710          Set frc = New FrmRequestConfirm
'8720          Set frc.Con = Con
'8730          frc.SdgName = Sdg("NAME")
'8740          Call frc.Show(vbModal)
'8750          If frc.ConfirmSucceeded Then
'8760              PapAuthorise
'8770          End If
'8780          frc.ConfirmSucceeded = False
'8790          Set frc = Nothing
'
'8800      Else
'8810          CyHyAuthorise
'8820      End If
'
'
'          'do not proceed in the authorisation process
'          'if did not authorise because of a missing mandatory result
'          '(do not unload the request etc):
'8830      If MandatoryExists = True Then
'8840          Exit Sub
'8850      End If
'
'8860      If Right(Sdg("EXTERNAL_REFERENCE"), 1) = "B" Then
'8870          Call UpdateTransferalToTissueArchiveHis(Sdg("SDG_ID"))
'8880      Else
'8890          Call UpdateTransferalToTissueArchiveCytoPap(Sdg("SDG_ID"))
'8900      End If
'
'
'8910      save_sdg_id = nte(Sdg("SDG_ID"))
'8920      If UCase(Role("NAME")) = UCase("doctor") Then
'8930          Call Con.Execute("update lims_sys.sdg_user set u_qc = '" & QcRank _
'    & "' " & "where sdg_id = " & Sdg("SDG_ID"))
'8940      End If
'8950      OpenedRequest = False
'8960      Call RefreshWindow
'
'8970      If Sdg("STATUS") = "I" Then
'8980          MsgBox "This request should be rechecked"
'8990      End If
'
'      '    If Right(Sdg("EXTERNAL_REFERENCE"), 1) = "P" Then
'      '        PapAuthoriseMsg
'      '    End If
'
'      '    If Sdg("STATUTestCurFrameS") = "A" Then Call TriggerSdgEvent("Print Final Letter")
'          'If PrintFax Or chkPrintFax.Value = 1 Then
'
'          'print only checked for print (not because result is malignant):
'9000      If chkPrintFax.value = 1 Then
'9010          Call TriggerSdgEvent("Print Fax", nte(Sdg("sdg_id")))
'9020      ElseIf chkPrintFinalLetter.value = 1 Then
'9030          If Sdg("STATUS") = "A" Then Call _
'    TriggerSdgEvent("Print Final Letter", nte(Sdg("sdg_id")))
'9040      End If
'9050      sdg_log_desc = AuthorizedBy(nte(Sdg("sdg_id")), 1)
'
'9060      Call UnloadRequest
'9070      Call zLang.English
'9080      SdgName.Alignment = vbLeftJustify
'9090      SdgName.RightToLeft = False
'9100      Call SdgName.SetFocus
'
'9110      AuthoriseButton.Enabled = False
'9120      SaveButton.Enabled = False
'9130      gridAliquots.Top = lblStatusBar.Top + lblStatusBar.Height
'9140      gridAliquots.Visible = False
'
'9150      CmdResponseLetter.Enabled = False
'
'9160      lblStatusBar.Caption = "Results where authorised successfully on " & _
'    Format(Now, "hh:mm:ss")
'9170      Call sdg_log.InsertLog(save_sdg_id, "RE.AUTH", sdg_log_desc)
'9180      Call SaveFreeTextContent
'
10100     Exit Sub
ErrHnd:
10110     Call ErrHandler("Authorise")
End Sub



Private Sub CloseButton_Click()
10120     On Error GoTo ErrHnd
10130     Call zLang.SetOrigLang
           
         '  MsgBox "CloseButton_Click"
10140     Call ReleaseApplicationMutex
10150     If RunFromWindow Then
10160         RaiseEvent CloseClicked
10170     Else
'MsgBox "If Not NtlsSite2 Is Nothing Then NtlsSite2.CloseWindow"
10180        If Not NtlsSite2 Is Nothing Then NtlsSite2.CloseWindow
'MsgBox " Exit Sub CloseButton_Click"
10190     End If

10200     Exit Sub
ErrHnd:
10210     Call ErrHandler("CloseButton_Click")
End Sub

Private Sub ConnectSameSession(ByVal aSessionID)
10220     On Error GoTo ErrHnd
          Dim aProc As New ADODB.Command
          Dim aSession As New ADODB.Parameter
          
10230     aProc.ActiveConnection = Con
10240     aProc.CommandText = "lims.lims_env.connect_same_session"
10250     aProc.CommandType = adCmdStoredProc

10260     aSession.Type = adDouble
10270     aSession.Direction = adParamInput
10280     aSession.value = aSessionID
10290     aProc.parameters.Append aSession

10300     aProc.Execute
10310     Set aSession = Nothing
10320     Set aProc = Nothing
10330     Exit Sub
ErrHnd:
10340     Call ErrHandler("ConnectSameSession")
End Sub

Private Sub UnloadRequest()
10350     On Error GoTo ErrHnd
          Dim i
          'Dim e As Object

10360     For i = 1 To PResultCheckIndex
10370         Unload PResultCheck(i)
10380     Next i
10390     PResultCheckIndex = 0
10400     For i = 1 To PResultIndex
10410         Unload PResultDesc(i)
10420         Unload PResultLine(i)
10430     Next i
10440     PResultIndex = 0
10450     For i = 1 To PResultTextIndex
10460         Unload PResultText(i)
10470     Next i
10480     PResultTextIndex = 0
10490     For i = 1 To PResultPhraseIndex
           ' If i < PResultPhrase.Count Then
10500         Unload PResultPhrase(i)
            ' End If
10510     Next i
10520     PResultPhraseIndex = 0
10530     For i = 2 To PFreeTextResultIndex
           
10540         PFreeTextResult(i).Terminate
10550         If nRTFResultBackup <> 3 Then
10560             Call PFreeTextResult(i).RemoveBackupResult
10570         End If
10580         Unload PFreeTextResult(i)
          
10590     Next i
10600     PFreeTextResultIndex = 1
10610     For i = 1 To PSnomedIndex
10620         SnomedCtrl(i).Terminate
10630         Unload SnomedCtrl(i)
10640     Next i
10650     PSnomedIndex = 0
          
10660     For i = 1 To PTestTabfra.Count
10670         If i = 1 Then PTestTab.Tabs(1).Caption = ""
10680         If i > 1 Then

10690             Unload PTestTabfra(i)
10700             PTestTab.Tabs.Remove (2)
10710         End If
10720     Next i

10730     For i = 1 To ImageRes.Count - 1
10740         Unload ImageRes(i)
10750     Next i

10760     If Not Sdg.State = adStateClosed Then Sdg.Close
10770     If Not Referring.State = adStateClosed Then Referring.Close
10780     If Not Aliquots.State = adStateClosed Then Aliquots.Close
10790     If Not Implement.State = adStateClosed Then Implement.Close
10800     If Not Patient.State = adStateClosed Then Patient.Close
10810     If Not Results.State = adStateClosed Then Results.Close
          
10820     nRTFResultBackup = 0
          
10830     OpenedRequest = False
          
10840     Exit Sub
ErrHnd:
10850     Call ErrHandler("UnloadRequest")
End Sub

Private Sub PatientProps_Click()
10860     On Error GoTo ErrHnd
10870     If PatientProps.SelectedItem.index = PropsCurFrame Then Exit Sub
10880     PatientPropsfra(PatientProps.SelectedItem.index).Visible = True
10890     PatientPropsfra(PropsCurFrame).Visible = False
10900     PropsCurFrame = PatientProps.SelectedItem.index
10910     Exit Sub
ErrHnd:
10920     Call ErrHandler("PatientProps_Click")
End Sub
Private Sub cmdOrangeDiagnosis_Click()
10930     Set PTestTab.SelectedItem = PTestTab.Tabs(1)
10940     PTestTab_Click
End Sub
Private Sub PTestTab_Click()
10950     On Error GoTo ErrHnd
10960     If PTestTab.SelectedItem.index = TestCurFrame Then
10970         SetFirstFocus
10980         Exit Sub
10990     End If
11000     PTestTabfra(PTestTab.SelectedItem.index).Visible = True
11010     PTestTabfra(TestCurFrame).Visible = False
11020     TestCurFrame = PTestTab.SelectedItem.index
11030     SetFirstFocus
          
11040     If UCase(PTestTab.SelectedItem.Caption) = UCase("Histology Macro") _
    Then
       

11050         Call PFreeTextResult(PFreeTextResultIndex).SetFocus
11060     End If
          'Call SdgName.SetFocus
              
11070     Exit Sub
ErrHnd:
11080     Call ErrHandler("PTestTab_Click")
End Sub
Private Function ExitQueryResult() As Boolean
11090     If MsgBox("?בוצעו שינויים שלא נשמרו. האם ברצונך לצאת ללא שמירה  ", _
    vbCritical + vbYesNoCancel + vbDefaultButton2) = vbYes Then
11100           ExitQueryResult = True
11110     Else
11120           ExitQueryResult = False
11130     End If
End Function
Private Sub SdgName_KeyDown(KeyCode As Integer, Shift As Integer)
          
          Dim rst As New ADODB.Recordset
          Dim Patholog As ADODB.Recordset
          Dim row As Integer
          
11140     On Error GoTo ErrHnd
          

          
     
11150     If Not KeyCode = vbKeyReturn Then Exit Sub
11160      btnPrintFax.Enabled = True
11170      AllOkBtn.Visible = False
11180     If FreeTextContentChanged Then
11190         If ExitQueryResult = False Then Exit Sub
11200     End If
11210     lblStatusBar.Caption = ""
11220     chkConsult.value = 0
11230     CmdResponseLetter.Enabled = False
11240     cmdAdditionalActions.Visible = False
11250     SdgName.Text = UCase(SdgName.Text)
11260     SdgName.Text = Replace(SdgName.Text, " ", "")
11270  SdgName.Text = Replace(SdgName.Text, vbCr, "")
11280  SdgName.Text = Replace(SdgName.Text, vbLf, "")
11290   If Mid(SdgName, 8, 1) = "." Then
11300       SdgName = Replace(SdgName, ".", "/", 1, 1)
11310   End If
11320   If Mid(SdgName, 9, 1) = "." Then
11330       SdgName = Replace(SdgName, ".", "/", 1, 1)
11340   End If

      '    If OpenedRequest Then
      '        If Sdg("STATUS") <> "A" Then
      '            If MsgBox("Results weren't saved." & vbCrLf & '                    "Are you sure you want to proceed?", vbYesNo) = vbNo Then
      '                SdgName.Text = ""
      '                Exit Sub
      '            End If
      '        End If
      '    End If
11350     Call PrintFaxResult.RemoveAll
11360     Call VisibleResultTab.RemoveAll
11370     PrintFax = False
11380     Call UnloadRequest
11390     Call LoadRecordsets
              
11400     If Sdg.EOF Then
11410       SdgName.Text = ""
11420       Exit Sub
11430    End If
          
11440     If Sdg("STATUS") = "A" Then
11450         MsgBox ("Sdg is Authorised, No changes can be made!")
11460         btnPrintFax.Enabled = True
11470     End If
          
11480     If Sdg("STATUS") = "S" Or Sdg("STATUS") = "X" Or Sdg("STATUS") = "U" _
    Then
11490         MsgBox "Request is of status " & Sdg("STATUS") & vbCrLf & _
    "and therefore can not be loaded!"
11500         SdgName.Text = ""
11510         Call zLang.English
11520         SdgName.Alignment = vbLeftJustify
11530         SdgName.RightToLeft = False
11540         SdgName.Text = ""
11550         Call SdgName.SetFocus
11560         Exit Sub
11570     End If
11580     Set rst = Con.Execute("select name from lims_sys.sample " & _
    "where sdg_id = " & Sdg("SDG_ID") & " and " & "(status = 'U' or status = 'S')")
11590     If Not rst.EOF Then
11600         MsgBox "Not all samples were received!"
11610         SdgName.Text = ""
11620         Call zLang.English
11630         SdgName.Alignment = vbLeftJustify
11640         SdgName.RightToLeft = False
11650         Call SdgName.SetFocus
11660         rst.Close
11670         SdgName.Text = ""
11680         Exit Sub
11690     End If
11700     rst.Close

11710        cmd_assutaPdf.Enabled = False
             
          'release previously owned semaphore:
11720     If strHandle <> "" Then
11730         Call ReleaseHandle
11740     End If
          
          
11750     lblRequestTitle.BackColor = &HC0FFFF
          
          
          'try and lock this request:
11760     If AllocateHandle("RESULT_ENTRY_" & SdgName.Text) = False Then
11770         strHandle = ""
              
11780         SdgName.Text = ""
11790         Call zLang.English
11800         SdgName.Alignment = vbLeftJustify
11810         SdgName.RightToLeft = False
11820         Call SdgName.SetFocus
11830         Exit Sub
11840     End If

          
11850     If Not InspectionLog Is Nothing Then
11860         If Not InspectionLog.State = adStateClosed Then _
    InspectionLog.Close
11870     End If
11880     Set InspectionLog = Con.Execute("select operator_id from " & _
    "lims_sys.inspection_log log " & "where log.table_name = 'SDG' and " & _
    "log.table_key = " & Sdg("SDG_ID") & " and " & "log.operator_id = " & _
    NtlsUser.GetOperatorId)
              
              'this area of code is very problematic!!!
              
      '    If (Right(Sdg("EXTERNAL_REFERENCE"), 1) = "P" And '       UCase(Role("NAME")) <> UCase("doctor") And '       UCase(Role("NAME")) <> UCase("pap inspector") And '       (Sdg("STATUS") = "I" Or Sdg("STATUS") = "A" Or Sdg("STATUS") = "R")) ' '        Or
11890     If (Right(Sdg("EXTERNAL_REFERENCE"), 1) = "P" And UCase(Role("NAME")) _
    <> UCase("doctor") And UCase(Role("NAME")) <> UCase("pap inspector") And _
    UCase(Role("NAME")) <> UCase("cytoscreener")) Or _
    (Right(Sdg("EXTERNAL_REFERENCE"), 1) <> "P" And UCase(Role("NAME")) = _
    UCase("doctor") And Sdg("STATUS") = "I" And Not InspectionLog.EOF) Or _
    (Right(Sdg("EXTERNAL_REFERENCE"), 1) <> "P" And UCase(Role("NAME")) <> _
    UCase("doctor")) Then
              
      '        SaveButton.Enabled = False
      'roy: anyone can authorize
11900         AuthoriseButton.Enabled = True
11910         AuthoriseButtonFlag = True
11920     Else
11930         SaveButton.Enabled = True
11940         AuthoriseButton.Enabled = True
11950         AuthoriseButtonFlag = True
11960     End If
          
      '    If IsInRevision(Sdg("SDG_ID")) Then
      '        AuthoriseButton.Visible = False
      '    Else
11970         AuthoriseButton.Visible = True
      '    End If
          
11980     If (Left(Sdg("name"), 1)) <> "P" Then
11990         SummaryButton.Enabled = False
12000         txtPapHeader.Visible = False
12010    Else
            'all ok button copies the results from a default aliquot
            'if the user in the all ok group,
            'and if there are no results yet

12020            Set rst = Con.Execute("select * from LIMS_SYS.LIMS_GROUP g, LIMS_SYS.OPERATOR_GROUP og " & _
            " where og.OPERATOR_ID= " & NtlsUser.GetOperatorId & "  and og.GROUP_ID = g.GROUP_ID and g.NAME='Pap Quick Answer' ")
12030           If Not rst.EOF And IsNull(Sdg("u_is_last_update")) And (Sdg("status") = "P" Or Sdg("status") = "V") Then

12040           AllOkBtn.Visible = True
12050           End If

12060
12070         SummaryButton.Enabled = True
              '------------------
              'PAT002
12080         txtPapHeader.Visible = True
12090         Call LoadPapHeader(Sdg("SDG_ID"))
              
12100     End If

12110     If Not RunFromWindow Then
12120           If Not NtlsSite2 Is Nothing Then
12130               NtlsSite2.SetWindowTitle (Sdg("NAME") & " - " & _
                        Patient("U_FIRST_NAME") & " " & Patient("U_LAST_NAME"))
12140           End If
12150     End If
       
12160     Call LoadProps
12170     row = 1
12180     Call LoadHistory
12190     row = 2
12200     Call LoadResults
12210     row = 3

              'ashi - Assuta interface
12220    If (nte(Sdg("U_ATFILENM")) <> "") Then
12230         cmd_assutaPdf.Enabled = True
12240         cmd_assutaPdf.BackColor = &HFFC0C0
12250    Else
12260         cmd_assutaPdf.Enabled = False
12270         cmd_assutaPdf.BackColor = &H80000016
12280    End If
12290    row = 4
          
          

          ' Barak - find the patholog
12300     If Sdg("STATUS") <> "A" And Sdg("STATUS") <> "R" Then
12310         If nte(Sdg("U_PATHOLOG")) <> "" Then
          '        Set Patholog = con.Execute("select phrase_description, phrase_name from lims_sys.phrase_entry " & "where phrase_id = (select phrase_id from lims_sys.phrase_header where " & "name = 'Patholog') and phrase_name = '" & nte(Sdg("U_PATHOLOG")) & "'")
                  'cmbPatholog.Text = nte(Patholog("PHRASE_DESCRIPTION"))
12320             If PathologCoredNumberToName(nte(Sdg("U_PATHOLOG"))) = "" Then
12330                 cmbPatholog.Text = "None"
12340             Else
12350                 cmbPatholog.Text = _
    PathologCoredNumberToName(nte(Sdg("U_PATHOLOG")))
12360             End If
12370         Else
12380             cmbPatholog.Text = "None"
12390         End If
12400     End If
12410     Set SdgStatusImage.Picture = LoadPicture("Resource\sdg" & _
    Sdg("STATUS") & ".ico")
          
12420     PSummaryText.Text = ""
12430     If SummaryButton.Caption = "Entry" Then Call SummaryButton_Click
12440     OpenedRequest = True
      '    Set rst = con.Execute("select 1 from lims_sys.inspection_entry entry, " & '        "lims_sys.inspection_log log " & '        "where log.table_name = 'SDG' and " & '        "log.table_key = " & Sdg("SDG_ID") & " and " & '        "log.order_number = entry.inspection_order and " & '        "entry.inspection_plan_id = " & Sdg("INSPECTION_PLAN_ID") & " and " & '        "log.role_id <> entry.role_id and " & '        "entry.role_id = " & NtlsUser.GetRoleId)
      '    If rst.EOF Then
      '        AuthoriseButton.Enabled = False
      '        SaveButton.Enabled = False
      '    Else
      '        AuthoriseButton.Enabled = True
      '        SaveButton.Enabled = True
      '    End If

12450     RequestRemarkCtrl.Visible = True
12460     Call RequestRemarkCtrl.GetsdgName(Sdg("name"))
12470     RequestRemarkCtrl.Refresh
12480     sdg_log_desc = ""
12490     Call sdg_log.InsertLog(Sdg("SDG_ID"), "RE.SELECT", sdg_log_desc)
          ' shlomi pinto::14-10-2004 -> check if ther is any aliquot with the result name report color that wasent entry/reported
12500     Call CheckReportColorSlides(Sdg("SDG_ID"), "Report Color")
          'do not present left material:
          'Call FillMaterialFromSample
          
          
12510     If Left(Sdg("NAME"), 1) = "P" Then
12520         Call ShowResultTab
12530     End If
          

12540     Call zLang.SetOrigLang
12550     Call InitAliquotGrid
12560     Call LoalResponseLetter(nte(Sdg("sdg_id")))
12570     Call SignalExtraRequest(nte(Sdg("external_reference")))
      '    Call SignalExtraRequest(nte(Sdg("sdg_id")))
          
          
12580     If InStr(1, nte(Sdg("NAME")), "V") = 0 Then
12590         cmdAdditionalActions.Visible = True
12600     End If
          
12610     If InStr(1, nte(Sdg("NAME")), "R") = 0 Then
          
12620       cmdAdditionalActions.Enabled = True
12630     Else
12640       cmdAdditionalActions.Enabled = False
12650     End If

12660 If Left(Sdg("NAME"), 1) = "B" Then
              Dim iSnomedTIndex As Integer
12670         OrganCtrl.Visible = True
12680         OrganCtrl.sdgId = Sdg("sdg_id")
12690         OrganCtrl.Initialize

12700         iSnomedTIndex = GetSnomedTResultIndex
              
              'get the Snomed T result from the Organ control
              'only if the value in the DB is empty:
12710     If iSnomedTIndex <> -1 Then
          'nautilus update
12720         If PResultText(iSnomedTIndex).Text = "" Then
12730             Call SetOrgansSnomedT
12740         End If
12750     End If
12760     Else
12770         OrganCtrl.Visible = False
12780         OrganCtrl.sdgId = ""
12790         OrganCtrl.SampleID = ""
              'OrganCtrl.Initialize
12800     End If
          'before clicking any tab, the focus returns to the initial input:
12810     If Not RunFromWindow Then
12820         Call SdgName.SetFocus
12830     End If
12840     Call SaveFreeTextContent
'11720     cmdOrangeDiagnosis.ZOrder (ZOrderConstants.vbBringToFront)
         
12850    SdgName.Text = ""

12860     Exit Sub
ErrHnd:
12870     Call ErrHandler("SdgName_KeyDown")
End Sub
Private Sub SaveFreeTextContent()
          Dim i As Integer
12880     ReDim OriginalFreeTextRes(PFreeTextResult.LBound To _
    PFreeTextResult.UBound) As String
12890     For i = PFreeTextResult.LBound To PFreeTextResult.UBound
12900         OriginalFreeTextRes(i) = PFreeTextResult(i).GetContent
12910     Next i
End Sub
Private Function FreeTextContentChanged() As Boolean
          Dim i As Integer
          Dim result As Boolean
12920     For i = PFreeTextResult.LBound To PFreeTextResult.UBound
12930         result = (OriginalFreeTextRes(i) <> _
    PFreeTextResult(i).GetContent) Or result
12940     Next i
12950     FreeTextContentChanged = result
End Function


'_____________________________________
'PAT - 002
Private Sub LoadPapHeader(sdgId As String)

12960     On Error GoTo ERR_LoadPapHeader
          Dim rsSample As Recordset
          Dim testCode As String
          Dim sql As String
          'we also neet  su.u_test_code in patholab
12970     sql = " select  su.u_test_code, s.name sample_name  "
12980     sql = sql & " from"
12990     sql = sql & "  lims_sys.sample s , lims_sys.sample_user su"
13000     sql = sql & " where "
13010     sql = sql & "         s.sdg_id =" & sdgId
13020     sql = sql & "     and su.sample_id=s.sample_id"
13030     sql = sql & "     and s.status<>'X' "
13040     Set rsSample = Con.Execute(sql)
13050     If Not rsSample.EOF Then
          'assuming there is only one sample EVER!!
13060         testCode = nte(rsSample("u_test_code"))

13070         If testCode = PAP_LBC_TEST_CODE Or testCode = _
    PAP_LBC_TEST_CODE_MEDICAL Then
13080             txtPapHeader = PAP_LBC_HEADER
13090             txtPapHeader.ForeColor = GREEN

13100         Else
13110             txtPapHeader = PAP_SMEAR_HEADER
13120             txtPapHeader.ForeColor = RED
13130         End If
13140         While Not rsSample.EOF
13150             testCode = nte(rsSample("u_test_code"))
13160             If testCode = PAP_TEST_CODE_MEDICAL Or testCode = _
    PAP_LBC_TEST_CODE_MEDICAL Then
13170                 lblRequestTitle.BackColor = RED
13180             End If
13190             rsSample.MoveNext
13200         Wend


13210     End If
13220     Exit Sub
ERR_LoadPapHeader:
13230 MsgBox "ERR_LoadPapHeader" & vbCrLf & Err.Description
End Sub
'__________

Private Sub ShowResultTab()
13240     On Error GoTo ErrHnd
          Dim i As Integer

13250     For i = 1 To VisibleResultTab.Count
13260         If CInt(VisibleResultTab(CStr(i))) > 0 Then
13270             ImageRes(i).Visible = True
13280         End If
13290     Next i
13300     Exit Sub
ErrHnd:
13310     Call ErrHandler("ShowResultTab")
End Sub

Private Sub FillMaterialFromSample()
13320     On Error GoTo ErrHnd
          Dim SampleRec As ADODB.Recordset

13330     Set SampleRec = Con.Execute("select count(*) " & _
    "from lims_sys.sample, lims_sys.sample_user " & _
    "where sample.sample_id = sample_user.sample_id and " & "sample.sdg_id = '" & _
    Sdg("SDG_ID") & "' and " & "sample_user.u_material = 'T'")

13340     LblMaterialTitle.Visible = False
13350     LblMaterialValue.Visible = False

13360     If Left(SdgName.Text, 1) = "B" Then
              
              
13370         LblMaterialTitle.Visible = True
13380         LblMaterialValue.Visible = True
13390         If SampleRec(0) = 0 Then
13400             LblMaterialValue.Caption = "לא"
13410         Else
13420             LblMaterialValue.Caption = "כן"
13430         End If
13440     End If
13450     Exit Sub
ErrHnd:
13460     Call ErrHandler("FillMaterialFromSample")
End Sub



Private Sub LoadRecordsets()
13470     On Error GoTo ErrHnd

13480     If InStr(SdgName.Text, ".") > 0 Then
13490         SdgName.Text = Left(SdgName.Text, InStr(SdgName.Text, ".") - 1)
13500     End If
'12530     If InStr(SdgName.Text, "_") > 0 Then
'12540         SdgName.Text = Left(SdgName.Text, InStr(SdgName.Text, "_") - 1)
'12550     End If
13510     Set Sdg = _
    Con.Execute("select * from lims_sys.sdg, lims_sys.sdg_user where " & _
    "sdg.sdg_id = sdg_user.sdg_id and " & "( sdg." & BarcodeField & " = '" & _
    SdgName.Text & "' or " & " sdg_user.U_PATHOLAB_NUMBER ='" & SdgName.Text & "')")
13520     If Sdg.EOF Then
13530         MsgBox ("Illegal Request Name or Patholab Number! (" & SdgName.Text & ")")
13540         SdgName.Text = ""
13550         Call zLang.English
13560         SdgName.Alignment = vbLeftJustify
13570         SdgName.RightToLeft = False
13580         Call SdgName.SetFocus
13590         Exit Sub
13600     End If
          ' shlomi pinto::17-10-2004 -> add "'" befor iif and after
13610     Set Referring = Con.Execute("select * from lims_sys.supplier, " & _
    "lims_sys.supplier_user, lims_sys.address " & _
    "where supplier.supplier_id = supplier_user.supplier_id and " & _
    "address_table_name(+) = 'SUPPLIER' and " & "address_line_1(+) = '" & _
    IIf(nte(Sdg("U_CLINIC_CODE")) = "", "0", nte(Sdg("U_CLINIC_CODE"))) & "' and " _
    & "address_item_id(+) = supplier.supplier_id and " & "supplier.supplier_id = " _
    & IIf(nte(Sdg("U_REFERRING_PHYSICIAN")) = "", "0", _
    nte(Sdg("U_REFERRING_PHYSICIAN"))))
          ' shlomi pinto::14-10-2004 -> add "'" befor iif and after
13620     Set Implement = Con.Execute("select * from lims_sys.supplier, " & _
    "lims_sys.supplier_user, lims_sys.address " & _
    "where supplier.supplier_id = supplier_user.supplier_id and " & _
    "address_table_name(+) = 'SUPPLIER' and " & "address_line_1(+) = '" & _
    IIf(nte(Sdg("U_IMPLEMENTING_CLINIC")) = "", "0", _
    nte(Sdg("U_IMPLEMENTING_CLINIC"))) & "' and " & _
    "address_item_id(+) = supplier.supplier_id and " & "supplier.supplier_id = " & _
    IIf(nte(Sdg("U_IMPLEMENTING_PHYSICIAN")) = "", "0", _
    nte(Sdg("U_IMPLEMENTING_PHYSICIAN"))))
13630     Set Patient = Con.Execute("select * from lims_sys.client, " & _
    "lims_sys.client_user, lims_sys.address " & _
    "where client.client_id = client_user.client_id and " & _
    "address_table_name(+) = 'CLIENT' and " & "address_line_1(+) = '0' and " & _
    "address_item_id(+) = client.client_id and " & "client.client_id = " & _
    Sdg("U_PATIENT"))
13640
13650 Set Results = Con.Execute("select result.result_id, result.name result_name, " & _
    "u_result_desc_user.u_bold,u_result_desc_user.u_height, u_result_desc_user.u_width, u_result_desc_user.u_visible,u_result_desc_user.u_read_only,u_result_desc_user.u_label," _
    & _
    "u_result_desc_user.u_free_text_template, u_result_desc_user.u_template_name,u_result_desc_user.u_type,u_result_desc_user.u_rtl,u_result_desc_user.u_phrase_list," _
    & _
    "u_result_desc_user.u_needs_review, u_result_desc_user.u_print_fax, u_result_desc_user.u_font_color, test.name test_name, result.status, result.description," _
    & _
    "test.priority, u_result_desc_user.u_order,u_result_desc_user.u_renk, formatted_result, test_template.amount_used " _
    & _
    "from lims_sys.result, lims_sys.result_user, lims_sys.test, lims_sys.aliquot, lims_sys.sample, lims_sys.u_result_desc_user, lims_sys.result_template, lims_sys.test_template " _
    & "where result.test_id = test.test_id " & _
    "and result.result_id = result_user.result_id " & _
    "and test.aliquot_id = aliquot.aliquot_id " & _
  "and aliquot.sample_id = sample.sample_id " & _
      "and result.result_template_id = result_template.result_template_id " & _
      "and result_template.name = u_result_desc_user.u_template_name " & _
      "and test_template.test_template_id = test.test_template_id " & _
      "and sample.sdg_id = " & Sdg("SDG_ID") & " " & "and test.priority > 0 " & _
      "and u_result_desc_user.u_order > 0 " & "and result.status <> 'X' " & _
      "and test.status <> 'X' " & "and aliquot.status <> 'X' " & _
      "order by test.priority, u_result_desc_user.u_order")
13660     Set Aliquots = _
    Con.Execute("select aliquot.aliquot_id from lims_sys.aliquot, " & _
    "lims_sys.sample,lims_sys.sdg where aliquot.sample_id = sample.sample_id " & _
    "and sample.sdg_id = sdg.sdg_id and sdg.sdg_id = " & Sdg("SDG_ID") & " " & _
    "and aliquot.status in ('V') " & "order by aliquot.aliquot_id desc")
13670     Set SnomedMCalculation = _
    Con.Execute("select description,u_snomed_code from " & _
    "lims_sys.u_snomed_calculation usc, lims_sys.u_snomed_calculation_user uscu " & _
    "where usc.u_snomed_calculation_id = uscu.u_snomed_calculation_id " & _
    "and u_snomed_template = 'Snomed M' and u_sdg_template = '" & Left(Sdg("NAME"), _
    1) & "'")
13680     Set SnomedTCalculation = _
    Con.Execute("select description,u_snomed_code from " & _
    "lims_sys.u_snomed_calculation usc, lims_sys.u_snomed_calculation_user uscu " & _
    "where usc.u_snomed_calculation_id = uscu.u_snomed_calculation_id " & _
    "and u_snomed_template = 'Snomed T' and u_sdg_template = '" & Left(Sdg("NAME"), _
    1) & "'")
13690     Set SampleCodes = Con.Execute("select * from " & _
    "lims_sys.u_sample_code usc, lims_sys.u_sample_code_user uscu " & _
    "where usc.u_sample_code_id = uscu.u_sample_code_id " & "and u_letter = '" & _
    Left(Sdg("NAME"), 1) & "'")
           Dim sql As String
13700     sql = " select o.name order_name, C.NAME Costumer_NAME "
13710     sql = sql & "from lims_sys.U_ORDER o, "
13720     sql = sql & "lims_sys.U_ORDER_USER ou, "
13730     sql = sql & "LIMS_SYS.U_CUSTOMER c, "
13740     sql = sql & "LIMS_SYS.U_CUSTOMER_USER cu "
13750     sql = sql & "where "
13760     sql = sql & "ou.u_sdg_name='" & Sdg("NAME") & "' "
13770     sql = sql & "and O.U_ORDER_ID=OU.U_ORDER_ID "
13780     sql = sql & "and OU.U_CUSTOMER= C.U_CUSTOMER_ID "
13790     sql = sql & "and OU.U_CUSTOMER= CU.U_CUSTOMER_ID "
          
13800     Set OrderAndCostumer = Con.Execute(sql)
          

13810     Exit Sub
          
ErrHnd:
13820     Call ErrHandler("LoadRecordsets")
End Sub

Private Sub LoadProps()
      Dim errorline As Integer

13830     On Error GoTo ErrHnd
          Dim i
          Dim Inspector As ADODB.Recordset
          Dim Collection As ADODB.Recordset
          Dim CompletedBy As ADODB.Recordset
          Dim Revisions As ADODB.Recordset
          Dim ConnectedRefCount As Integer
       

13840     lblRequestTitle.Caption = nte(Sdg("U_PATHOLAB_NUMBER"))
           ''nte(Sdg("NAME")) & " - "
13850     If OrderAndCostumer.EOF Then
13860         lblPayingCustomer.Caption = ""
13870     Else
13880         lblPayingCustomer.Caption = nte(OrderAndCostumer("Costumer_NAME"))
13890     End If
       
13900     Set Revisions = Con.Execute("select sdg.name from lims_sys.sdg " & _
    "where sdg.name = '" & Sdg("NAME") & "V1'")
13910     LblRevisionStatus.Caption = ""
13920     If Not Revisions.EOF Then
13930         LblRevisionStatus.Caption = "רוויזיה"
13940     ElseIf InStr(1, Sdg("NAME"), "V") Then
13950         LblRevisionStatus.Caption = "קיים עדכון"
13960     End If
      '  roy: this area of code never changed title color.
          
13970     If Not SampleCodes.EOF Then
13980         lblSampleCodeRemark.Caption = nte(SampleCodes("u_remark"))
13990     End If
          
14000     If lblSampleCodeRemark <> "" Then
14010         lblSampleCodeRemark.Visible = True
14020         lblRequestTitle.BackColor = RED
14030     Else
14040         lblSampleCodeRemark.Visible = False
      '        lblRequestTitle.BackColor = &HC0FFFF
14050     End If
       
14060     Revisions.Close
14070     If Not Patient.EOF Then
14080         PropsGeneralPatientName.Text = nte(Patient("U_FIRST_NAME")) & " " _
    & nte(Patient("U_LAST_NAME"))
14090         PropsGeneralPatientID.Text = nte(Patient("NAME"))
14100         PropsGeneralPatientGender.Text = nte(Patient("U_GENDER"))
14110         PropsGeneralPatientBirth.Text = ntd(Patient("U_DATE_OF_BIRTH"))
14120         PropsPatientName = nte(Patient("U_FIRST_NAME")) & " " & _
    nte(Patient("U_LAST_NAME"))
14130         lblRequestTitle.Caption = lblRequestTitle & " - " & nte(Patient("NAME")) & " - " & _
    PropsPatientName
14140         PropsPatientAddress = nte(Patient("ADDRESS_LINE_2")) & vbCrLf & _
    nte(Patient("ADDRESS_LINE_3")) & vbCrLf & nte(Patient("ADDRESS_LINE_4")) & _
    vbCrLf & nte(Patient("ADDRESS_LINE_5")) & vbCrLf & nte(Patient("POSTAL_CODE")) _
    & vbCrLf & nte(Patient("PHONE")) & vbCrLf & nte(Patient("FAX")) & vbCrLf
14150     End If
       
14160     PropsGeneralReferring.Text = ""
14170     PropsPhysRefName.Text = ""
14180     PropsPhysRefAddress.Text = ""
14190     If Not Referring.EOF Then
14200         PropsGeneralReferring.Text = nte(Referring("U_FIRST_NAME")) & " " _
    & nte(Referring("U_LAST_NAME")) & " - " & nte(Referring("U_LICENSE_NBR"))
14210         PropsPhysRefName.Text = nte(Referring("U_FIRST_NAME")) & " " & _
    nte(Referring("U_LAST_NAME"))
14220         PropsPhysRefAddress.Text = nte(Referring("ADDRESS_LINE_2")) & _
    vbCrLf & nte(Referring("ADDRESS_LINE_3")) & vbCrLf & _
    nte(Referring("ADDRESS_LINE_4")) & vbCrLf & nte(Referring("ADDRESS_LINE_5")) & _
    vbCrLf & nte(Referring("POSTAL_CODE")) & vbCrLf & nte(Referring("PHONE")) & _
    vbCrLf & nte(Referring("FAX")) & vbCrLf
14230     End If
       
14240     PropsGeneralSubmitting.Text = ""
14250     PropsPhysSubName.Text = ""
14260     PropsPhysSubAddress.Text = ""
14270     If Not Implement.EOF Then
14280         PropsGeneralSubmitting.Text = nte(Implement("U_FIRST_NAME")) & " " _
    & nte(Implement("U_LAST_NAME")) & " - " & nte(Implement("U_LICENSE_NBR"))
14290         PropsPhysSubName.Text = nte(Implement("U_FIRST_NAME")) & " " & _
    nte(Implement("U_LAST_NAME"))
14300         PropsPhysSubAddress.Text = nte(Implement("ADDRESS_LINE_2")) & _
    vbCrLf & nte(Implement("ADDRESS_LINE_3")) & vbCrLf & _
    nte(Implement("ADDRESS_LINE_4")) & vbCrLf & nte(Implement("ADDRESS_LINE_5")) & _
    vbCrLf & nte(Implement("POSTAL_CODE")) & vbCrLf & nte(Implement("PHONE")) & _
    vbCrLf & nte(Implement("FAX"))
14310     End If
14320     PropsGeneralSdgPriority.Text = nte(Sdg("U_PRIORITY"))
14330     PropsGeneralSdgDelivery.Text = nte(Sdg("U_REFERRAL_DATE"))
14340     PropsGeneralSdgSlides.Text = nte(Sdg("U_SLIDE_NBR"))
14350     If IsNull(Sdg("u_implementing_clinic")) Then
14360         PropsGeneralSdgCollection.Text = ""
14370         PropsGeneralSdgCollection.Text = ""
14380     Else
14390         Set Collection = _
    Con.Execute("select * from lims_sys.u_clinic, lims_sys.u_clinic_user, lims_sys.address " _
    & "where u_clinic.u_clinic_id = u_clinic_user.u_clinic_id and " & _
    "address_table_name(+) = 'U_CLINIC' and " & _
    "address_line_1(+) = u_clinic.name and " & _
    "address_item_id(+) = u_clinic.u_clinic_id and " & "u_clinic.u_clinic_id = '" & _
    Sdg("u_implementing_clinic") & "'")
14400         If Not Collection.EOF Then
14410             PropsGeneralSdgCollection.Text = nte(Collection("NAME")) & _
    " - " & nte(Collection("U_CLINIC_NAME"))
14420             PropsPhysColName.Text = nte(Collection("NAME")) & " - " & _
    nte(Collection("U_CLINIC_NAME"))
14430             PropsPhysColAddress.Text = nte(Collection("ADDRESS_LINE_2")) _
    & vbCrLf & nte(Collection("ADDRESS_LINE_3")) & vbCrLf & _
    nte(Collection("ADDRESS_LINE_4")) & vbCrLf & nte(Collection("ADDRESS_LINE_5")) _
    & vbCrLf & nte(Collection("POSTAL_CODE")) & vbCrLf & nte(Collection("PHONE")) & _
    vbCrLf & nte(Collection("FAX"))
14440         End If
14450         Collection.Close
14460     End If
14470     PropsGeneralSdgWeek.Text = nte(Sdg("U_WEEK_NBR"))

          'i = 11
      '    Set Inspector = con.Execute("select operator.name " & '        "from lims_sys.inspection_log, lims_sys.operator " & '        "where table_name = 'SDG' " & '        "and table_key = " & Sdg("SDG_ID") & " " & '        "and inspection_log.operator_id = operator.operator_id " & '        "order by order_number")
14480     PropsGeneralSdgAuthorized(0).ForeColor = BLACK
14490     If Sdg("STATUS") = "I" And UCase(Role("NAME")) = UCase("PAP Inspector") _
    Then
14500         PropsGeneralSdgAuthorized(0).Text = "Checked By QC"
14510         PropsGeneralSdgAuthorized(0).ForeColor = RED
14520         PropsGeneralSdgAuthorized(1).Text = ""
14530         PropsGeneralSdgAuthorized(2).Text = ""
14540     Else

'get the authorizers from 'Sign by ' results and inspection log
Dim sql As String
Dim inspectorIndex As Integer
14550 inspectorIndex = 0
14560 sql = "  select result.original_result as name "
14570 sql = sql & " from lims_sys.result,   lims_sys.sample "
14580 sql = sql & " ,lims_sys.aliquot, lims_sys.test  "
14590 sql = sql & " where test.aliquot_id = aliquot.aliquot_id  "
14600 sql = sql & " and aliquot.sample_id = sample.sample_id  "
14610 sql = sql & " and result.test_id = test.test_id  "
14620 sql = sql & " and sample.sdg_id =" & Sdg("SDG_ID") & " "
14630 sql = sql & " and result.status <> 'X'  "
14640 sql = sql & " and result.name = "






14650  Set Inspector = Con.Execute(sql & "'Sign by 1st.' ")
                  
                    
14660              If Inspector.EOF Then
                       
14670                 PropsGeneralSdgAuthorized(inspectorIndex).Text = ""
14680                   inspectorIndex = 0
14690             Else
                    
14700                 PropsGeneralSdgAuthorized(inspectorIndex).Text = nte(Inspector("NAME"))
14710                 inspectorIndex = 1
          '            Inspector.MoveNext
14720             End If
                   
14730             Inspector.Close


14740  Set Inspector = Con.Execute(sql & "'Sign by 2nd.' ")
                  
                    
14750              If Inspector.EOF Then
                       
14760                 PropsGeneralSdgAuthorized(inspectorIndex).Text = ""
14770                   inspectorIndex = inspectorIndex
14780             Else
                    
14790                 PropsGeneralSdgAuthorized(inspectorIndex).Text = nte(Inspector("NAME"))
14800                 inspectorIndex = inspectorIndex + 1
          '            Inspector.MoveNext
14810             End If
                   
14820             Inspector.Close
' at this point there might be two authoriser or no authorisers out of three spots
14830  For i = 0 To 2 - inspectorIndex



14840             Set Inspector = Con.Execute("select operator.name " & _
    "from lims_sys.operator " & "where operator_id = lims.authorization.signed_by(" _
    & Sdg("SDG_ID") & "," & i + 1 & ")")
                  
                    
14850              If Inspector.EOF Then
                   
14860                 PropsGeneralSdgAuthorized(inspectorIndex + i).Text = ""
14870             Else
                    
14880                 PropsGeneralSdgAuthorized(inspectorIndex + i).Text = nte(Inspector("NAME"))
          '            Inspector.MoveNext
14890             End If
                   
14900             Inspector.Close
14910         Next i
14920     End If
      '    Inspector.Close

      '    Dim Diagnose As ADODB.Recordset
      '    For i = 1 To 4
      '        If Not IsNull(Sdg("U_DIAGNOSIS" & i)) Then
      '            Set Diagnose = con.Execute("select u_diagnos_name from lims_sys.u_diagnos, lims_sys.u_diagnos_user " & '                "where u_diagnos.u_diagnos_id = u_diagnos_user.u_diagnos_id and " & '                "name = '" & Sdg("U_DIAGNOSIS" & i) & "'")
      '            If Not Diagnose.EOF Then
      '                PropsReferralDiagnose(i) = Diagnose("U_DIAGNOS_NAME")
      '            Else
      '                PropsReferralDiagnose(i) = ""
      '            End If
      '            Call Diagnose.Close
      '        Else
      '            PropsReferralDiagnose(i) = ""
      '        End If
      '    Next i
       
14930     If (Right(nte(Sdg("EXTERNAL_REFERENCE")), 1) <> "B") Then 'hila- cancel the use of ref in case Histology "B"
14940         If chkRefCancel.value <> vbChecked Then
14950             If Debugging Then MsgBox 101100
14960             InitReferrals
14970             ConnectedRefCount = Ref.ConnectedRefCount
14980             For i = 0 To ConnectedRefCount - 1
14990                 If i = 2 Then Exit For
15000                 Select Case Ref.RequestType
                          Case "B"
                          
15010                         If Ref.Histology_.IsExecuting(CInt(i)) Then
15020                             PropsReferralDiagnose(1).Text = _
    Ref.GetRefSummary(CInt(i))
15030                             PropsReferralDiagnose(1).Tag = i & ""
15040                         Else
15050                             PropsReferralDiagnose(2).Text = _
    Ref.GetRefSummary(CInt(i))
15060                             PropsReferralDiagnose(2).Tag = i & ""
15070                         End If
15080                     Case "C"
15090                         If Ref.Cytology_.IsExecuting(CInt(i)) Then
15100                             PropsReferralDiagnose(1).Text = _
    Ref.GetRefSummary(CInt(i))
15110                             PropsReferralDiagnose(1).Tag = i & ""
15120                         Else
15130                             PropsReferralDiagnose(2).Text = _
    Ref.GetRefSummary(CInt(i))
15140                             PropsReferralDiagnose(2).Tag = i & ""
15150                         End If
15160                     Case "P"
15170                         PropsReferralDiagnose(1).Text = _
    Ref.GetRefSummary(CInt(i))
15180                         PropsReferralDiagnose(1).Tag = i & ""
15190                 End Select
15200             Next i
15210         End If
15220     End If 'hila
      'If Debugging Then MsgBox 10
      '    QCtxt = nte(Sdg("U_QC"))
'14010     QcRank = ntz(Sdg("U_QC"))
15230 If Debugging Then MsgBox 11

      '    chkQC.Value = IIf(ntz(Sdg("U_ISQC")) = "T", 1, 0)
15240     chkCon.value = IIf(ntz(Sdg("U_ISCONSULT")) = "T", 1, 0)


15250 If Debugging Then MsgBox 12
15260     SdgCompleted.Text = ""
15270     SdgCompleted.ForeColor = BLACK
15280     If nte(Sdg("COMPLETED_BY")) <> "" Then
15290         Set CompletedBy = Con.Execute("select operator.name " & _
    "from lims_sys.operator " & "where operator.operator_id = " & _
    nte(Sdg("COMPLETED_BY")))
15300         If Not CompletedBy.EOF Then
15310             SdgCompleted.Text = CompletedBy("Name")
15320         End If
15330         CompletedBy.Close
15340     End If
15350 If Debugging Then MsgBox 13
15360     If Sdg("STATUS") = "I" And UCase(Role("NAME")) = UCase("PAP Inspector") _
    Then
15370         SdgCompleted.ForeColor = RED
15380         SdgCompleted.Text = "Checked By QC"
15390     End If
15400     TxtAuthorizedOn.Text = Trim(nte(Sdg("AUTHORISED_ON")))
15410     Exit Sub
ErrHnd:
15420     Call ErrHandler("LoadProps ")
End Sub
Private Function TamplateExists(sdgId As String) As Boolean
          Dim sql As String
15430   sql = " select 1 "
15440     sql = sql & "  from "
15450     sql = sql & "  lims_sys.sdg d, "
15460     sql = sql & "  lims_sys.sample s, "
15470     sql = sql & "  lims_sys.sample_user su, "
15480     sql = sql & "  lims_sys.U_LIST l "
15490     sql = sql & "  where "
15500     sql = sql & "  d.SDG_ID= " & sdgId
15510     sql = sql & "  and  su.SAMPLE_ID=s.SAMPLE_ID "
15520     sql = sql & "  and d.sdg_id =s.sdg_id "
15530     sql = sql & "  and  l.name  like '%' ||  su.U_ORGAN_CODE  || '%'"
15540     sql = sql & " and rownum=1"
          ' return True if Tamplate Exists
15550     TamplateExists = Not Con.Execute(sql).EOF
End Function


Private Sub LoadResults()
15560     On Error GoTo ErrHnd
          Dim CurTest As String
          Dim i, j
          Dim w, h, TabCurTop, MaxTop As Integer
          Dim CountBool As Integer
15570     IsMicroTextSaved = True
15580 If Debugging Then MsgBox "load res 1"
15590     Call dicResultIdToName.RemoveAll
15600     i = 1
15610     PResultIndex = 0
15620     PResultCheckIndex = 0
15630     PResultTextIndex = 0
15640     PResultPhraseIndex = 0
15650
15660     PSnomedIndex = 0
15670     TestCurFrame = 1
15680     PTestTab.Height = PTestTabHeight
15690     PTestTabfra(1).Height = PTestTabFraHeight
15700      frmDiagnosis.Visible = False
15710     If Results.EOF Then Exit Sub
15720     Do Until Results.EOF
          
15730         If Not dicResultIdToName.Exists(nte(Results("RESULT_ID"))) Then
15740             Call dicResultIdToName.Add(nte(Results("RESULT_ID")), _
    nte(Results("RESULT_NAME")))
15750         End If
              
              'free text result
15760         If Results("PRIORITY") = 10 Then Exit Do
15770         CurTest = Results("TEST_NAME")
15780         If i > 1 Then
15790             Call PTestTab.Tabs.Add(i, , CurTest)
15800             Load PTestTabfra(i)
15810             PTestTabfra(i).Left = PTestTab.ClientLeft
15820             PTestTabfra(i).Top = PTestTab.ClientTop
15830             PTestTabfra(i).Width = PTestTab.ClientWidth
15840             PTestTabfra(i).Height = PTestTab.ClientHeight
15850             PTestTabfra(i).ZOrder (0)
15860             PTestTabfra(i).Visible = False
15870             PTestTabfra(i).Height = PTestTabFraHeight
15880         End If
15890         PTestTab.Tabs(i).Caption = CurTest
15900         TabCurTop = 100
15910         j = 1
15920         CountBool = 0
15930 If Debugging Then MsgBox 1
              'shlomi pinto::17-10-2004
15940         If Left(Sdg("NAME"), 1) = "P" Then
15950             Load ImageRes(PTestTab.Tabs.Count)
15960             With ImageRes(PTestTab.Tabs.Count)
15970                 .Visible = False
15980                 .Left = nte(Results("AMOUNT_USED"))
15990                 .Picture = LoadPicture("resource\down.ico")
16000             End With
16010         End If
'14900         If UCase(nte(Results("RESULT_NAME"))) = UCase("DIAGNOSIS") And _
'    nte(Results("formatted_result")) = "" Then
'14910                 frmDiagnosis.Visible = True
'14920                 cmdOrangeDiagnosis.ZOrder 0
'14930         End If
16020           If UCase(nte(Results("RESULT_NAME"))) = UCase("HISTOLOGY MICRO") _
    Then
16030                 TamplatePFResInsex = 3
16040                  If TamplateExists(Sdg("sdg_id")) Then _
    PTestTabfra(i).BackColor = &HFF0000
16050                 If nte(Results("formatted_result")) = "" Then
16060                     IsMicroTextSaved = False
                         
16070                 End If
16080             End If
16090         Do Until Results.EOF
16100             If UCase(nte(Results("RESULT_NAME"))) = _
    UCase("HISTOLOGY MICRO") Then
16110                TamplatePFResInsex = 3
16120                 If TamplateExists(Sdg("sdg_id")) Then _
    PTestTabfra(i).BackColor = &HFF0000
16130                 If nte(Results("formatted_result")) = "" Then
16140                     IsMicroTextSaved = False
                         
16150                 End If
16160             End If
16170 If Debugging Then MsgBox 2
16180             If Not dicResultIdToName.Exists(nte(Results("RESULT_ID"))) _
    Then
16190                 Call dicResultIdToName.Add(nte(Results("RESULT_ID")), _
    nte(Results("RESULT_NAME")))
16200             End If
16210 If Debugging Then MsgBox 3
16220             If CurTest <> Results("TEST_NAME") Then Exit Do
16230             PResultIndex = PResultIndex + 1
16240             Load PResultDesc(PResultIndex)
16250             With PResultDesc(PResultIndex)
16260                 Set .Container = PTestTabfra(i)
16270                 .Top = TabCurTop
16280                 .Left = IIf(Results("U_TYPE") = "B", 355, 100)
16290                 .Width = PTestTab.ClientWidth - 300
16300                 .Height = 255
16310                 .FontBold = IIf(Results("U_BOLD") = "T", True, False)
16320                 .Caption = nte(Results("U_LABEL"))
16330                 .Tag = nte(Results("U_TYPE"))
16340                 .DataField = nte(Results("U_TEMPLATE_NAME"))
16350                 .Visible = True
                      'shlomi pinto::17-10--2004
16360                 If nte(Results("U_FONT_COLOR")) <> "" Then
16370                     .ForeColor = nte(Results("U_FONT_COLOR"))
16380                 End If
16390             End With
16400 If Debugging Then MsgBox 4
16410             Load PResultLine(PResultIndex)
16420             With PResultLine(PResultIndex)
16430                 Set .Container = PTestTabfra(i)
16440                 .X1 = 50
16450                 .Y1 = TabCurTop + 290
16460                 .X2 = PTestTab.ClientWidth - 50
16470                 .Y2 = .Y1
16480                 .BorderColor = &H80000011
16490                 If Results("U_TYPE") <> "F" And Results("U_TYPE") <> "S" _
    Then
16500                     .Visible = True
16510                 End If
16520             End With
16530 If Debugging Then MsgBox 5
16540             If Results("U_TYPE") = "B" Then
16550                PResultCheckIndex = PResultCheckIndex + 1
16560                 Load PResultCheck(PResultCheckIndex)
16570                 With PResultCheck(PResultCheckIndex)
16580                     Set .Container = PTestTabfra(i)
16590                     .Tag = Results("RESULT_ID")
16600                     If nte(Results("U_PRINT_FAX")) = "T" Then Call _
    PrintFaxResult.Add(.Tag, "")
16610                     .Top = TabCurTop
16620                     TabCurTop = TabCurTop + 300
16630                     .Left = 100
16640                     .Width = 255
16650                     .Height = 255
'roy : here needs to add  'or ="T"' to support formatted res = T
16660                     If Results("FORMATTED_RESULT") = "True" Then .value = _
    1
16670                     If Results("U_NEEDS_REVIEW") = "T" Then .Caption = "T"
16680                     .Visible = IIf(Results("U_VISIBLE") = "F", False, _
    True)
16690                     PResultDesc(PResultIndex).Tag = _
    PResultDesc(PResultIndex).Tag & PResultCheckIndex
16700 If Debugging Then MsgBox 6
16710                     If .value = 1 Then
16720                         CountBool = CountBool + 1
16730                     End If
16740                     If VisibleResultTab.Exists(CStr(PTestTab.Tabs.Count)) _
    = True Then
16750                         Call _
    VisibleResultTab.Remove(CStr(PTestTab.Tabs.Count))
16760                     End If
16770                     Call VisibleResultTab.Add(CStr(PTestTab.Tabs.Count), _
    CountBool)

16780                 End With
16790             ElseIf Results("U_TYPE") = "F" Then
16800 If Debugging Then MsgBox 7
16810                 PFreeTextResultIndex = PFreeTextResultIndex + 1
16820                 Load PFreeTextResult(PFreeTextResultIndex)
16830                 With PFreeTextResult(PFreeTextResultIndex)
16840                     Set .connection = Con
16850                     .InitContent = getInitStr

                          '-------------------------
                          ' shlomi pinto::23-09-2004
16860                     .FontName = "Arial"
16870                     .RightMargin = nInch * 6
                          '-------------------------
                  
16880                     .Lists = getListsNames
            
16890                     .Initialize
16900                     .Tag = Results("RESULT_ID")
                       
16910                     Set .Container = PTestTabfra(i)
16920                     .Top = TabCurTop
16930                     .Left = 100
16940                     PFreeTextResult(1).Visible = False
16950                     PTestTab.Height = PatientProps.Height
16960                     If nte(Results("U_HEIGHT")) <> "" Then h = _
    nte(Results("U_HEIGHT"))
16970                     .Width = PTestTabfra(i).Width - 100
16980                     .Height = (PTestTab.Height - 480) * h / 100
16990                     TabCurTop = TabCurTop + _
    PFreeTextResult(PFreeTextResultIndex).Height
17000                     .Visible = True
17010                     .locked = IIf(Results("U_READ_ONLY") = "F", False, _
    True)
17020                     .Rtl = IIf(Results("U_RTL") = "F", False, True)
17030                     PResultDesc(PResultIndex).Tag = _
    PResultDesc(PResultIndex).Tag & PFreeTextResultIndex
                          
                          'using backup from rtf_result_backup:
17040                     If nte(Results("STATUS")) <> "A" Then
                              
17050                         .ResultId = Results("RESULT_ID")
                              
17060                     End If
                          
17070                 End With
                      
17080             ElseIf Results("U_TYPE") = "S" Then

17090 If Debugging Then MsgBox 8
17100                 PSnomedIndex = PSnomedIndex + 1
17110                 Load SnomedCtrl(PSnomedIndex)
17120                 With SnomedCtrl(PSnomedIndex)
17130                     Set .connection = Con
17140                     .StatusReadWrite = SnomedCtrl(0).CReadWrite
17150                     .Initialize _
    ("select original_result from lims_sys.result where " & "result.result_id = " & _
    Results("RESULT_ID"))
17160                     .Tag = Results("RESULT_ID")
17170                     Set .Container = PTestTabfra(i)
17180                     .Top = TabCurTop + 300
17190                     .Left = 100
17200                     If nte(Results("U_HEIGHT")) <> "" Then h = _
    nte(Results("U_HEIGHT"))
17210                     .Width = PTestTabfra(i).Width - 150
17220                     h = 35
17230                     .Height = PTestTabfra(i).Height * h / 100
17240                     TabCurTop = SnomedCtrl(PSnomedIndex).Top + _
    SnomedCtrl(PSnomedIndex).Height
17250                     .ShowCloseBtn False
17260                     .Visible = True
17270                     PResultDesc(PResultIndex).Tag = _
    PResultDesc(PResultIndex).Tag & PSnomedIndex
17280                 End With
17290             ElseIf Results("U_TYPE") = "P" Then
17300                 PResultPhraseIndex = PResultPhraseIndex + 1
17310                 Load PResultPhrase(PResultPhraseIndex)
17320                 With PResultPhrase(PResultPhraseIndex)
17330                     Set .Container = PTestTabfra(i)
17340                     Set .connection = Con
17350                     .Tag = Results("RESULT_ID")
17360                     .Top = TabCurTop - 60
17370                     TabCurTop = TabCurTop + 380
                                              
17380                     If nte(Results("U_WIDTH")) <> "" Then w = _
    nte(Results("U_WIDTH"))
                          
17390                     .Width = PTestTabfra(i).Width * w / 100
                          ' .Left = IIf(Results("U_RTL") = "T", 10, PTestTab.ClientWidth - .Width)
17400                      .Left = PTestTab.ClientWidth - .Width
      '                    .Width = 2255
      '                    .Height = 255
17410                     .InitContent = nte(Results("FORMATTED_RESULT"))
17420                     .PhraseName = nte(Results("U_PHRASE_LIST"))
17430                     .Rtl = IIf(Results("U_RTL") = "T", True, False)
17440                     .Initialize
17450                     .Visible = True
17460                     PResultDesc(PResultIndex).Tag = _
    PResultDesc(PResultIndex).Tag & PResultPhraseIndex
17470                 End With
17480             Else ' Results("U_TYPE") = "T" ??

17490 If Debugging Then MsgBox 9
17500                 PResultTextIndex = PResultTextIndex + 1
                    
17510                 Load PResultText(PResultTextIndex)
17520                 With PResultText(PResultTextIndex)
17530                     Set .Container = PTestTabfra(i)
17540                     .Tag = Results("RESULT_ID")
17550                     .Top = TabCurTop
17560                     TabCurTop = TabCurTop + 300
17570                     .Left = PTestTab.ClientWidth - 2300
17580                     .Width = 2255
17590                     .Height = 255
17600                     .Text = nte(Results("FORMATTED_RESULT"))
17610                     .Visible = True
17620                     PResultDesc(PResultIndex).Tag = _
    PResultDesc(PResultIndex).Tag & PResultTextIndex
17630                 End With
          
17640             End If
          
17650 If Debugging Then MsgBox 1000 + j
17660             j = j + 1
17670             Results.MoveNext
17680         Loop

17690         i = i + 1
17700         If MaxTop < TabCurTop Then MaxTop = TabCurTop
17710 If Debugging Then MsgBox 1000 + j + 100 * i
17720     Loop
          
17730     If PFreeTextResultIndex = 1 Then
17740         PTestTab.Height = MaxTop + 600
17750     End If
17760     For i = 1 To PTestTabfra.Count
17770         PTestTabfra(i).Height = PTestTab.Height - 480
17780     Next i
          'free text result
17790 If Debugging Then MsgBox 20
17800     If Not Results.EOF Then
17810 If Debugging Then MsgBox 21
17820         If Results("PRIORITY") = 10 And PFreeTextResultIndex = 1 Then
17830 If Debugging Then MsgBox 22
17840             Set PFreeTextResult(1).connection = Con
17850             PFreeTextResult(1).InitContent = getInitStr

                  '-------------------------
                  ' shlomi pinto::23-09-2004
17860             PFreeTextResult(1).FontName = "Arial"
17870             PFreeTextResult(1).RightMargin = nInch * 6
                  '-------------------------

17880             PFreeTextResult(1).Initialize
17890             PFreeTextResult(1).Tag = Results("RESULT_ID")
17900             PFreeTextResult(1).Visible = True
17910             PFreeTextResult(1).Top = PTestTab.Height + 100
17920             PFreeTextResult(1).Width = PTestTab.Width
17930             If PapsResultsfra.Height - PFreeTextResult(1).Top > 0 Then
17940                 PFreeTextResult(1).Height = PapsResultsfra.Height - _
    PFreeTextResult(1).Top
17950             End If
17960         End If
17970     End If
17980 If Debugging Then MsgBox 22
17990     Call SummaryRefresh
18000 If Debugging Then MsgBox 23
18010     Set PTestTab.SelectedItem = PTestTab.Tabs(1)
18020 If Debugging Then MsgBox 24
18030     PTestTabfra(PTestTab.SelectedItem.index).Visible = True
18040     Call PTestTab_Click
18050 If Debugging Then MsgBox 25
18060     Exit Sub
ErrHnd:
18070     Call ErrHandler("LoadResults")
End Sub
Private Sub LoadHistory()
18080     On Error GoTo ErrHnd
          Dim li As ListItem
          Dim i As Integer
          Dim Snomed As String
          Dim InLoadHistory As Boolean
          Dim FirstMalignant As Boolean

      '    InLoadHistory = True
18090     FirstMalignant = False
18100     HistoryList.ListItems.Clear '
18110     HistoryGrid.Clear
18120     HistoryGrid.Rows = 2
18130     Set SnomedCtrl(0).connection = Con

18140     Set History = _
    Con.Execute("select sdg.name, sdg.created_on, sdg.status, sdg.sdg_id " & _
    "from lims_sys.sdg, lims_sys.sdg_user " & "where sdg.sdg_id = sdg_user.sdg_id " _
    & "and sdg_user.u_patient = " & Sdg("U_PATIENT") & " " & "and sdg.sdg_id <> " & _
    Sdg("SDG_ID") & " and rownum < 100 order by sdg.created_on desc")

18150     Set imgHistory.Picture = LoadPicture("Resource\Led On.ico")
18160     imgHistory.Visible = IIf(History.EOF, False, True)
18170     HistoryGrid.row = 0
18180     HistoryGrid.col = 0
18190     HistoryGrid.Text = "Request"
18200     HistoryGrid.col = 1
18210     HistoryGrid.Text = "Date"
18220     HistoryGrid.col = 2
18230     HistoryGrid.Text = "Snomed"

18240     i = 1
18250     Do Until History.EOF
      '        Set li = HistoryList.ListItems.Add(, , nte(History("NAME")), , History("STATUS") & " 1") '
      '        li.SubItems(1) = Format(nte(History("CREATED_ON")), "dd/mm/yy") '
      '        li.SubItems(2) = nte(History("SNOMED")) '
      '        Snomed = SnomedCtrl(0).getFirstSnomed("select u_snomed from lims_sys.sdg_user where sdg_id = " & nte(History("SDG_ID")))
18260         If History("STATUS") = "C" Or History("STATUS") = "A" Then
18270             Snomed = _
    SnomedCtrl(0).getFirstSnomed("select nvl(result.ORIGINAL_RESULT,'1') " & _
    "from lims_sys.result, lims_sys.test, lims_sys.aliquot, lims_sys.sample " & _
    "where result.test_id = test.test_id and " & _
    "test.aliquot_id = aliquot.aliquot_id and " & _
    "aliquot.sample_id = sample.sample_id and " & "result.name = 'Snomed T' and " & _
    "sample.sdg_id = " & nte(History("SDG_ID")))
18280         Else
18290             Snomed = ""
18300         End If
18310         If i > 1 Then
18320             HistoryGrid.AddItem (nte(History("NAME")) & Chr(9) & _
    Format(nte(History("CREATED_ON")), "dd/mm/yy") & Chr(9) & Snomed)
18330         Else
18340             HistoryGrid.row = 1
18350             HistoryGrid.col = 0
18360             HistoryGrid.Text = nte(History("NAME"))
18370             HistoryGrid.col = 1
18380             HistoryGrid.Text = Format(nte(History("CREATED_ON")), _
    "dd/mm/yy")
18390             HistoryGrid.col = 2
18400             HistoryGrid.Text = Snomed
18410         End If

18420         HistoryGrid.row = i
18430         HistoryGrid.col = 0

18440         HistoryGrid.CellBackColor = vbWhite
18450         If CheckIsMalignant(nte(History("SDG_ID"))) Then
18460             If Not FirstMalignant Then
18470                 FirstMalignant = True
18480                 Set imgHistory.Picture = _
    LoadPicture("Resource\Led Off.ico")
18490             End If
18500             HistoryGrid.CellBackColor = MALIGNANT_REQUEST
18510         End If

18520         HistoryGrid.CellAlignment = flexAlignRightCenter
18530         picHistory.Picture = _
    HistoryImageList.ListImages.Item(History("STATUS") & " 1").Picture
18540         Set HistoryGrid.CellPicture = picHistory.Image
18550         History.MoveNext
              
18560         i = i + 1
18570     Loop
18580     History.Close
18590     SnomedCtrl(0).Terminate
18600     InLoadHistory = False
18610     Exit Sub
ErrHnd:
18620     Call ErrHandler("LoadHistory")
End Sub

Private Function CheckIsMalignant(sdgId As String) As Boolean
18630     On Error GoTo ErrEnd
          Dim strSQL As String
          Dim IsMalignant As String
          Dim MalignantRs As ADODB.Recordset

18640     CheckIsMalignant = False
18650     strSQL = "select lims.is_malignant('" & sdgId & "') " & _
    "from lims_sys.sdg " & "where sdg_id = '" & sdgId & "'"
18660     Set MalignantRs = Con.Execute(strSQL)
18670     If Not MalignantRs.EOF Then
18680         IsMalignant = Trim(nte(MalignantRs(0)))
18690         If IsMalignant = "T" Then
18700             CheckIsMalignant = True
18710         End If
18720     End If
18730     Exit Function

ErrEnd:
18740     MsgBox "CheckIsMalignant... " & vbCrLf & Err.Description
End Function

'Private Function nte(e As Variant) As Variant
'    On Error GoTo ErrHnd
'    nte = IIf(IsNull(e), "", e)
'    Exit Function
'ErrHnd:
'    Call ErrHandler("nte")
'End Function

Private Function nte(e As Variant) As String
'17560     On Error GoTo ErrHnd
          
18750     nte = IIf(IsNull(e), "", e)
          
'17580     Exit Function
'ErrHnd:
'17590     Call ErrHandler("nte")
End Function
Private Function ntd(e As Variant) As String
18760       ntd = IIf(IsNull(e), 0, Format(e, "dd/MM/yyyy"))
 
End Function


Private Function ntz(e As Variant) As Variant
18770     ntz = IIf(IsNull(e), 0, e)
End Function


Private Sub HistoryList_DblClick()
18780     On Error GoTo ErrHnd
          Dim OldBarcodeField As String
          
18790     If HistoryList.ListItems.Count = 0 Then Exit Sub
18800     SdgName.Text = HistoryList.SelectedItem.Text
18810     OldBarcodeField = BarcodeField
18820     BarcodeField = "NAME"
18830     Call SdgName_KeyDown(vbKeyReturn, 0)
          'Call SdgName_KeyUp(vbKeyReturn, 0)
18840     BarcodeField = OldBarcodeField
18850     Exit Sub
ErrHnd:
18860     Call ErrHandler("HistoryList_DblClick")
End Sub


Private Sub SnomedCtrl_CloseClick(index As Integer)
18870     On Error GoTo ErrHnd
18880     If index <> 0 Then Exit Sub
18890     SnomedCtrl(index).Visible = False
18900     Exit Sub
ErrHnd:
18910     Call ErrHandler("SnomedCtrl_CloseClick")
End Sub


Private Sub SnomedCtrl_GotFocus(index As Integer)
18920     On Error GoTo ErrHnd
18930     If index <> 0 Then Exit Sub
18940     SnomedCtrl(index).ZOrder ZOrderConstants.vbBringToFront
18950     Exit Sub
ErrHnd:
18960     Call ErrHandler("SnomedCtrl_GotFocus")
End Sub

Private Sub SnomedCtrl_LostFocus(index As Integer)
18970     On Error GoTo ErrHnd
18980     If index <> 0 Then Exit Sub
18990     SnomedCtrl(index).Terminate
19000     SnomedCtrl(index).Visible = False
19010     Exit Sub
ErrHnd:
19020     Call ErrHandler("SnomedCtrl_LostFocus")
End Sub

Private Sub SummaryButton_Click()
19030     On Error GoTo ErrHnd
19040     If SummaryButton.Caption = "Summary" Then
19050         SummaryButton.Caption = "Entry"
19060         PSummaryfra.Height = PTestTab.Height
19070         PSummaryfra.Top = PTestTab.Top
19080         PSummaryfra.Width = PTestTab.Width
19090         PSummaryText.Height = PSummaryfra.Height - 400
19100         PSummaryText.Width = PSummaryfra.Width - 250
19110         PSummaryfra.Visible = True
19120         PSummaryfra.ZOrder (0)
19130         Call SummaryRefresh
19140     Else
19150         SummaryButton.Caption = "Summary"
19160         PSummaryfra.Visible = False
19170     End If
19180     Exit Sub
ErrHnd:
19190     Call ErrHandler("SummaryButton_Click")
End Sub

Private Sub SummaryRefresh()
19200     On Error GoTo ErrHnd
19210     PSummaryText.Text = GetSummary
          
19220     Exit Sub
ErrHnd:
19230     Call ErrHandler("SummaryRefresh")
End Sub

Private Sub RefreshWindow()
19240     On Error GoTo ErrHnd
          
19250     btnPrintFax.Enabled = True
          
19260     If Sdg.EOF Then Exit Sub
          Dim OldBarcodeField As String
19270     SdgName.Text = Sdg("NAME")
19280     OldBarcodeField = BarcodeField
19290     BarcodeField = "NAME"
19300     Call LoadRecordsets
19310     Call LoadProps
19320     If Sdg.EOF Then Exit Sub
19330     Set SdgStatusImage.Picture = LoadPicture("Resource\sdg" & _
    Sdg("STATUS") & ".ico")
      '    Call SdgName_KeyDown(vbKeyReturn, 0)
      '    Call SdgName_KeyUp(vbKeyReturn, 0)
19340     BarcodeField = OldBarcodeField
19350     SdgName.Text = ""
19360     CmdResponseLetter.Enabled = False
19370     Call zLang.English
19380     SdgName.Alignment = vbLeftJustify
19390     SdgName.RightToLeft = False
19400     Call SdgName.SetFocus
19410 If ((Sdg("STATUS") <> "V" And Sdg("STATUS") <> "P") Or Not IsNull(Sdg("u_is_last_update"))) Then AllOkBtn.Visible = False


19420     Exit Sub
ErrHnd:
19430     Call ErrHandler("RefreshWindow")
End Sub

Private Sub btnPrint_Click()
19440     On Error GoTo ErrHnd
19450     If Sdg.State = adStateClosed Then
19460         MsgBox "Load a request for printing"
19470         Exit Sub
19480     End If
19490     Call TriggerSdgEvent("Print PDF Letter", nte(Sdg("sdg_id")))
19500     Call SdgName.SetFocus
19510     Exit Sub
ErrHnd:
19520     Call ErrHandler("btnPrint_Click")
End Sub

Private Sub TriggerSdgEvent(EventName As String, strSdgId As String)
19530     On Error GoTo ErrHnd
          Dim doc As New DOMDocument
          Dim res As New DOMDocument
          Dim xmlLogin As IXMLDOMElement
          Dim xmlSdg As IXMLDOMElement
          Dim e As IXMLDOMElement
          Dim element As IXMLDOMElement
          Dim FileName As String

19540     Set e = doc.createElement("lims-request")
19550     Call doc.appendChild(e)
19560     Set xmlLogin = doc.createElement("login-request")
19570     Call e.appendChild(xmlLogin)
19580     Set xmlSdg = doc.createElement("SDG")
19590     Call xmlLogin.appendChild(xmlSdg)
19600     Set element = doc.createElement("find-by-id")
19610     element.Text = strSdgId 'Sdg("SDG_ID")
19620     Call xmlSdg.appendChild(element)
19630     Set element = doc.createElement("fire-event")
19640     element.Text = EventName
19650     Call xmlSdg.appendChild(element)

      '    doc.Save ("auth.xml")

      '    If Trim(WorkFolder) <> "" Then
      '        FileName = "C:\ResultEntry_" & Trim(strSdgId) & "_" & EventName & "_DOC3"
      '        Call xmlManager.SaveXmlFile(doc, FileName)
      '    End If
      'doc.Save ("c:\1.xml")
19660     Call ProcessXML.ProcessXMLWithResponse(doc, res)
      ' res.Save ("c:\2.xml")
          'res xml file was not saved befor

      '    If Trim(WorkFolder) <> "" Then
      '        FileName = "C:\ResultEntry_" & Trim(strSdgId) & "_" & EventName & "_RES3"
      '        Call xmlManager.SaveXmlFile(res, FileName)
      '    End If

19670     Exit Sub
ErrHnd:
19680     Call ErrHandler("TriggerSdgEvent")
End Sub
Private Function getInitStr() As String
19690     On Error GoTo ErrHnd
          Dim InitStr As String
          Dim RtfResult As New ADODB.Recordset
19700     If Results("STATUS") = "V" Then
19710         InitStr = nte(Results("U_FREE_TEXT_TEMPLATE"))
19720     Else 'If Results("STATUS") = "C" Then
              'InitStr = nte(Results("FORMATTED_RESULT"))
19730         Call _
    RtfResult.Open("select rtf_text from lims_sys.rtf_result where rtf_result_id = " _
    & Results("RESULT_ID"), Con, adOpenStatic, adLockOptimistic)
19740         If Not RtfResult.EOF Then
19750             InitStr = ReadClob(RtfResult("RTF_TEXT"))
19760         End If
19770         RtfResult.Close
19780     End If
19790     InitStr = Replace(InitStr, "&SdgId&", Sdg("SDG_ID"))
19800     InitStr = Replace(InitStr, "&ResultId&", Results("RESULT_ID"))
19810     InitStr = Replace(InitStr, "&OperatorId&", NtlsUser.GetOperatorId)
19820     InitStr = Replace(InitStr, "&RoleId&", NtlsUser.GetRoleId)
19830     getInitStr = InitStr
19840     Exit Function
ErrHnd:
19850     Call ErrHandler("getInitStr")
End Function

Private Function getListsNames() As String
19860     On Error GoTo ErrHnd
          Dim ListsNames As String
19870     ListsNames = nte(Results("U_PHRASE_LIST"))
19880     ListsNames = Replace(ListsNames, "&SdgId&", Sdg("SDG_ID"))
19890     ListsNames = Replace(ListsNames, "&ResultId&", Results("RESULT_ID"))
19900     ListsNames = Replace(ListsNames, "&OperatorId&", _
    NtlsUser.GetOperatorId)
19910     ListsNames = Replace(ListsNames, "&RoleId&", NtlsUser.GetRoleId)
19920     getListsNames = ListsNames
19930     Exit Function
ErrHnd:
19940     Call ErrHandler("getListsNames")
End Function

Private Sub NewPapAuthorise()
19950     On Error GoTo ErrHnd
          Dim rst As ADODB.Recordset
        Dim sdgId As String
19960   sdgId = Sdg("SDG_ID")
19970     If Not InspectionLog.EOF Then
19980         MsgBox "This request already been signed by " & _
    NtlsUser.GetOperatorName & " and therefore cannot be authorise"
19990         Exit Sub
20000     End If

20010     If (Sdg("STATUS") = _
    "C" Or Sdg("STATUS") = "V" Or Sdg("STATUS") = "P" Or Sdg("STATUS") = "I") Then

'19270         If Not AssignPapInspection Then Exit Sub
20020         AuthoriseResults ("A")
20030         InsertNote

20040         btnPrintFax.Enabled = True
    
20050          Con.Execute ("Update lims_sys.SDG_USER SET U_WEEK_NBR='909 ' WHERE SDG_ID=" & sdgId)
                 
20060         Call TriggerSdgEvent("Print PDF Letter", sdgId)
20070         Exit Sub
20080     End If

'19320     If (UCase(Role("NAME")) = UCase("PAP Inspector")) And (Sdg("STATUS") _
'    = "C" Or Sdg("STATUS") = "V" Or Sdg("STATUS") = "P") Or (Sdg("STATUS") = "I") _
'    Then
'19330         SaveResults
'19340         If Not AssignPapInspection Then Exit Sub
'19350         AuthoriseResults ("A")
'19360         InsertNote
'19370         Exit Sub
'19380     End If

'19390     If UCase(Role("NAME")) = UCase("doctor") Then
'19400         Set rst = _
'    Con.Execute("select instr(old_status,'I') from lims_sys.sdg " & _
'    "where sdg_id = " & Sdg("SDG_ID"))
'19410         If (Sdg("STATUS") = "A" Or Sdg("STATUS") = "R") Then
'19420             InsertIntoInspectionLog
'      '        ElseIf (Sdg("STATUS") = "I") Or CInt(rst(0).Value) > 0 Then
'      '            SaveResults
'      '            AuthoriseResults ("A")
'19430         Else
'19440             SaveResults
'19450             If Not AssignPapInspection Then Exit Sub
'19460             AuthoriseResults ("A")
'19470         End If
'19480         InsertNote
'19490         Exit Sub
'19500     End If

20090     MsgBox "Not the right status for authorization"
20100     Exit Sub
ErrHnd:
20110     Call ErrHandler("NewPapAuthorise")
End Sub
Private Sub PapAuthorise()
20120     On Error GoTo ErrHnd
          Dim rst As ADODB.Recordset
'
'19210     If Not InspectionLog.EOF Then
'19220         MsgBox "This request already been signed by " & _
'    NtlsUser.GetOperatorName & " and therefore cannot be authorise"
'19230         Exit Sub
'19240     End If
'
'19250     If (UCase(Role("NAME")) = UCase("cytoscreener")) And (Sdg("STATUS") = _
'    "C" Or Sdg("STATUS") = "V" Or Sdg("STATUS") = "P") Then
'19260         SaveResults
'19270         If Not AssignPapInspection Then Exit Sub
'19280         AuthoriseResults ("A")
'19290         InsertNote
'19300         Exit Sub
'19310     End If
'
'19320     If (UCase(Role("NAME")) = UCase("PAP Inspector")) And (Sdg("STATUS") _
'    = "C" Or Sdg("STATUS") = "V" Or Sdg("STATUS") = "P") Or (Sdg("STATUS") = "I") _
'    Then
'19330         SaveResults
'19340         If Not AssignPapInspection Then Exit Sub
'19350         AuthoriseResults ("A")
'19360         InsertNote
'19370         Exit Sub
'19380     End If
'
'19390     If UCase(Role("NAME")) = UCase("doctor") Then
'19400         Set rst = _
'    Con.Execute("select instr(old_status,'I') from lims_sys.sdg " & _
'    "where sdg_id = " & Sdg("SDG_ID"))
'19410         If (Sdg("STATUS") = "A" Or Sdg("STATUS") = "R") Then
'19420             InsertIntoInspectionLog
'      '        ElseIf (Sdg("STATUS") = "I") Or CInt(rst(0).Value) > 0 Then
'      '            SaveResults
'      '            AuthoriseResults ("A")
'19430         Else
'19440             SaveResults
'19450             If Not AssignPapInspection Then Exit Sub
'19460             AuthoriseResults ("A")
'19470         End If
'19480         InsertNote
'19490         Exit Sub
'19500     End If

20130     MsgBox "Only Doctor or Cytoscreener or Pap Inspector Can Authorise"
20140     Exit Sub
ErrHnd:
20150     Call ErrHandler("PapAuthorise")
End Sub

Private Sub CyHyAuthorise()
20160     On Error GoTo ErrHnd
      ' TODO: if the sdg is cyto and the role is "super cyto" he can authorize
'19550     If UCase(Role("NAME")) <> UCase("doctor") Then
'19560         MsgBox "Only Doctor Can Authorise"
'19570         Exit Sub
'19580     End If
'
'
'19590     If InspectionLog.EOF And Sdg("STATUS") = "I" Then
'19600         SaveResults
'19610         AuthoriseResults ("A")
'19620         Exit Sub
'19630     End If
'
'19640     If Sdg("STATUS") = "A" Or Sdg("STATUS") = "R" Then
'19650         InsertIntoInspectionLog
'19660         Exit Sub
'19670     End If
'19680     If Sdg("STATUS") = "C" Or Sdg("STATUS") = "V" Or Sdg("STATUS") = "P" _
'    Then
'19690         SaveResults
'      ' see remarks (TODO) in this function
'19700         AssignInspection
'19710         AuthoriseResults ("A")
'19720         Exit Sub
'19730     End If
20170     Exit Sub
ErrHnd:
20180     Call ErrHandler("CyHyAuthorise")
End Sub

Private Sub NewCyHyAuthorise()
20190     On Error GoTo ErrHnd
      ' TODO: if the sdg is cyto and the role is "super cyto" he can authorize
20200     If UCase(Role("NAME")) <> UCase("doctor") Then
20210         MsgBox "Only Doctor Can Authorise"
20220         Exit Sub
20230     End If
          

20240     If InspectionLog.EOF And Sdg("STATUS") = "I" Then
20250         SaveResults
20260         AuthoriseResults ("A")
20270         Exit Sub
20280     End If
          
20290     If Sdg("STATUS") = "A" Or Sdg("STATUS") = "R" Then
20300         InsertIntoInspectionLog
20310         Exit Sub
20320     End If
20330     If Sdg("STATUS") = "C" Or Sdg("STATUS") = "V" Or Sdg("STATUS") = "P" _
    Then
20340         SaveResults
      ' see remarks (TODO) in this function
20350         AssignInspection
20360         AuthoriseResults ("A")
20370         Exit Sub
20380     End If
20390     Exit Sub
ErrHnd:
20400     Call ErrHandler("NewCyHyAuthorise")
End Sub

Private Sub InsertIntoInspectionLog()
20410     On Error GoTo ErrHnd
20420     Call Con.Execute("insert into lims_sys.inspection_log " & _
    "(inspection_log_id, table_name, table_key, inspection_type, " & _
    "role_id, operator_id, order_number, inspection_date) " & _
    "values (lims.sq_inspection_log.nextval, 'SDG', " & Sdg("SDG_ID") & ", " & _
    "'A', " & NtlsUser.GetRoleId & ", " & NtlsUser.GetOperatorId & ", " & _
    "(select count(1) from lims_sys.inspection_log where table_name = 'SDG' and " & _
    "table_key = " & Sdg("SDG_ID") & ") + 1, SYSDATE)")
20430     Exit Sub
ErrHnd:
20440     Call ErrHandler("InsertIntoInspectionLog")
End Sub
Private Sub InsertOperatorIntoInspectionLog(roleIdCommaOperatorId As String)
20450     On Error GoTo ErrHnd
20460     Call Con.Execute("insert into lims_sys.inspection_log " & _
    "(inspection_log_id, table_name, table_key, inspection_type, " & _
    "role_id, operator_id, order_number, inspection_date) " & _
    "values (lims.sq_inspection_log.nextval, 'SDG', " & Sdg("SDG_ID") & ", " & _
    "'A', " & roleIdCommaOperatorId & ", " & _
    "(select count(1) from lims_sys.inspection_log where table_name = 'SDG' and " & _
    "table_key = " & Sdg("SDG_ID") & ") + 1, SYSDATE)")
20470     Exit Sub
ErrHnd:
20480     Call ErrHandler("InsertIntoInspectionLog")
End Sub
Private Sub AssignInspection()
20490     On Error GoTo ErrHnd
          Dim rand As Double
          Dim QCParameter As Integer
          Dim strQC As String
          Dim IsQC As Boolean
          Dim fqc As FrmQC

      '    If MandatoryExists Then
      '        ChangeStatus
      '        Exit Sub
      '    End If

      ' TODO: if the role is "super cyto" and the sdg is cyto change the
      'inspection plan and don't do QC

20500     strQC = Right(Sdg("EXTERNAL_REFERENCE"), 1)
20510     Select Case strQC
          Case "C"
20520         QCParameter = CQCParameter
20530     Case "B"
20540         QCParameter = HQCParameter
20550     End Select
20560     Randomize
20570     IsQC = Rnd < QCParameter / CDbl(100)
          'MsgBox "rand = " & rand & " inspect = " & Inspect & " QCParameter = " & QCParameter & " calc = " & QCParameter / CDbl(100)
20580     If IsQC Or Inspect Then
20590         Call Con.Execute("update lims_sys.sdg set inspection_plan_id = " _
    & "(select inspection_plan_id from lims_sys.inspection_plan " & _
    "where inspection_plan.name = '" & strQC & "QC') " & "where sdg_id = " & _
    Sdg("SDG_ID"))
20600     End If
20610     If Inspect Then
20620         MsgBox "This request should be rechecked by a physician."
20630     End If
20640     If IsQC Then

      '            MsgBox "This request should be rechecked for QC."

20650         Set fqc = New FrmQC
20660         Call fqc.Show(vbModal)
20670         If fqc.ConfirmSucceeded Then
20680             Call Con.Execute("update lims_sys.sdg_user set u_qc = '0' " & _
    "where sdg_id = " & Sdg("SDG_ID"))
20690         End If
20700         fqc.ConfirmSucceeded = False
20710         Set fqc = Nothing

20720     End If
20730     Exit Sub
ErrHnd:
20740     Call ErrHandler("AssignInspection")
End Sub

Private Function IsConsult() As Boolean
20750     On Error GoTo ErrHnd
      '    Dim ConsultResults As ADODB.Recordset
      '
      '    IsConsult = False
      '    Set ConsultResults = con.Execute( '        "select r.original_result " & '        "from lims_sys.result r, " & '             "lims_sys.test t, " & '             "lims_sys.aliquot a, " & '             "lims_sys.sample s " & '        "where t.aliquot_id = a.aliquot_id " & '             "and a.sample_id = s.sample_id " & '             "and r.test_id = t.test_id " & '             "and s.sdg_id = " & Sdg("SDG_ID") & " " & '             "and r.name in ('rem_consult','rem_consult_cito' )")
      '    While Not ConsultResults.EOF
      '        If nte(ConsultResults("ORIGINAL_RESULT")) = "T" Then
      '            IsConsult = True
      '        End If
      '        ConsultResults.MoveNext
      '    Wend
      '    ConsultResults.Close
20760     IsConsult = IIf(chkConsult.value = 0, False, True)
20770     Exit Function

ErrHnd:
20780     Call ErrHandler("IsConsult")
End Function

Private Function ChangInspectionForDoctorOnly() As Boolean
20790 On Error GoTo ErrHnd
Dim Found As Boolean


Dim operqatorId
Dim index As Integer
Dim oldInspection As String, inspectionPlanName As String
    
20800     If doctorOnly Then
    'change inspection
    
20810         oldInspection = nte(Sdg("INSPECTION_PLAN_ID"))
    
20820           inspectionPlanName = "Doctor Only"
20830           If Sdg("STATUS") = "I" Then
20840               inspectionPlanName = "Any + Doctor Only"
20850           End If
                
                Dim inspection As Recordset
    
20860             Set inspection = _
        Con.Execute("select inspection_plan_id from lims_sys.inspection_plan " & _
        "where upper(inspection_plan.name)= upper('" & inspectionPlanName & "') ")
20870         If Not inspection.EOF Then
20880             Call Con.Execute("update lims_sys.sdg set inspection_plan_id = " _
        & inspection("inspection_plan_id") & "where sdg_id = " & Sdg("SDG_ID"))
               
                  Dim sdg_log_description   As String
20890             sdg_log_description = oldInspection & "=>" & _
                     inspection("inspection_plan_id")
20900             Call sdg_log.InsertLog(Sdg("SDG_ID"), "RE.AssignPapInspection", _
                    sdg_log_description)
20910          Else
20920            MsgBox "could not find the ""Doctor Only"" inspection plan"
20930         End If
20940      End If
       
   
20950    Exit Function

ErrHnd:
20960    Call ErrHandler("IsConsult")
End Function


Private Function ProblemWithPapsResults() As Boolean
20970 On Error GoTo ErrHnd

'assume no problem
20980 ProblemWithPapsResults = False

'get the phrase with queries
 Dim phraseWitheQueriesAndErrorMessages As ADODB.Recordset
Dim runQuery As ADODB.Recordset
 
20990 Set phraseWitheQueriesAndErrorMessages = _
    Con.Execute("select phrase_description as query, phrase_info as message, phrase_name from lims_sys.phrase_entry " _
    & "where phrase_id = (select phrase_id from lims_sys.phrase_header where " & _
    "name = 'Check Paps Results') ")
    
21000   While Not phraseWitheQueriesAndErrorMessages.EOF
21010   If Not IsNull(phraseWitheQueriesAndErrorMessages("QUERY")) Then
 
         ' If InStr("SELECT", phraseWitheQueriesAndErrorMessages("QUERY"), vbTextCompare) > 0 Then
21020   If UCase(phraseWitheQueriesAndErrorMessages("QUERY")) <> "T" Then
 
21030       Set runQuery = Con.Execute(Replace(phraseWitheQueriesAndErrorMessages("QUERY"), "#SDG_ID#", Sdg("SDG_ID"), 1, -1, vbTextCompare))
21040       If runQuery.EOF Then
21050           ProblemWithPapsResults = True
21060           MsgBox phraseWitheQueriesAndErrorMessages("message")
21070           phraseWitheQueriesAndErrorMessages.Close
21080           Exit Function
21090       ElseIf Replace$(Trim$(runQuery.Fields(0)), vbTab, "") <> "T" Then
21100           ProblemWithPapsResults = True
21110           MsgBox phraseWitheQueriesAndErrorMessages("message")
21120           phraseWitheQueriesAndErrorMessages.Close
21130           Exit Function
21140       End If
21150     End If
21160  End If
  
21170   phraseWitheQueriesAndErrorMessages.MoveNext

21180   Wend
21190 phraseWitheQueriesAndErrorMessages.Close
21200 Exit Function
ErrHnd:
21210     Call ErrHandler("ProblemWithPapsResults")
End Function
Private Function Mandatory() As Boolean
21220     On Error GoTo ErrHnd
          Dim MandatoryResults As ADODB.Recordset
          Dim sql As String
21230     Mandatory = False
21240     MandatoryExists = False
      '    Set MandatoryResults = con.Execute("select result.original_result," & "u_result_desc_user.u_mandatory_value,u_result_desc_user.u_label " & "from lims_sys.result, lims_sys.u_result_desc_user, lims_sys.result_template, lims_sys.sample " & ",lims_sys.aliquot, lims_sys.test " & "where result.result_template_id = result_template.result_template_id " & "and result_template.name = u_result_desc_user.u_template_name " & "and test.aliquot_id = aliquot.aliquot_id " & "and aliquot.sample_id = sample.sample_id " & "and result.test_id = test.test_id " & "and sample.sdg_id = " & Sdg("SDG_ID") & " " & "and u_result_desc_user.u_mandatory = 'T' " & "and u_result_desc_user.u_mandatory_value = result.original_result ")
          
21250     sql = " select result.original_result,"
21260     sql = sql & _
    " u_result_desc_user.u_mandatory_value,u_result_desc_user.u_label "
21270     sql = sql & _
    " from lims_sys.result, lims_sys.u_result_desc_user, lims_sys.result_template, lims_sys.sample"
21280     sql = sql & " ,lims_sys.aliquot, lims_sys.test "
21290     sql = sql & _
    " where result.result_template_id = result_template.result_template_id "
21300     sql = sql & _
    " and result_template.name = u_result_desc_user.u_template_name "
21310     sql = sql & " and test.aliquot_id = aliquot.aliquot_id "
21320     sql = sql & " and aliquot.sample_id = sample.sample_id "
21330     sql = sql & " and result.test_id = test.test_id "
21340     sql = sql & " and sample.sdg_id = '" & Sdg("SDG_ID") & "' "
21350     sql = sql & " and result.status <> 'X' "
21360     sql = sql & " and u_result_desc_user.u_mandatory = 'T' "
21370     sql = sql & " and "
21380     sql = sql & " ("
21390     sql = sql & "   ("
21400     sql = sql & "     u_result_desc_user.u_mandatory_value is null"
21410     sql = sql & "     and "
21420     sql = sql & "     ("
21430     sql = sql & "       result.original_result is null or"
21440     sql = sql & "       result.original_result = 'F'"
21450     sql = sql & "     )"
21460     sql = sql & "   )"
21470     sql = sql & "   or  "
21480     sql = sql & "   ("
21490     sql = sql & "     u_result_desc_user.u_mandatory_value is  not null"
21500     sql = sql & "     and "
21510     sql = sql & "     ("
21520     sql = sql & _
    "       result.original_result <> u_result_desc_user.u_mandatory_value"
21530     sql = sql & "     )"
21540     sql = sql & "   )"
21550     sql = sql & " )"
          
21560     Set MandatoryResults = Con.Execute(sql)
          
21570     If MandatoryResults.EOF Then
21580         MandatoryResults.Close
21590         Exit Function
21600     End If
          
21610     frmMsgBox.ShowMsg "Mandatory Result Is Missing : " & _
    nte(MandatoryResults("U_LABEL"))
21620     Mandatory = True
21630     MandatoryExists = True
          
      '    If nte(MandatoryResults("U_MANDATORY_VALUE")) = "" Then
      '        If nte(MandatoryResults("ORIGINAL_RESULT")) = "F" Or nte(MandatoryResults("ORIGINAL_RESULT")) = "" Then
      '            frmMsgBox.ShowMsg "Mandatory Result Is Missing : " & nte(MandatoryResults("U_LABEL"))
      '            Mandatory = True
      '        End If
      '    Else
      '        If nte(MandatoryResults("ORIGINAL_RESULT")) <> nte(MandatoryResults("U_MANDATORY_VALUE")) Then
      '            frmMsgBox.ShowMsg "Mandatory Result Is Missing : " & nte(MandatoryResults("U_LABEL"))
      '            Mandatory = True
      '        End If
      '    End If
21640     MandatoryResults.Close
21650     Exit Function
ErrHnd:
21660     Call ErrHandler("Mandatory")
End Function
Private Sub ChangeStatus()
21670     On Error GoTo ErrHnd
          Dim SdgStatus As ADODB.Recordset
21680     Set SdgStatus = _
    Con.Execute("select status from lims_sys.sdg, lims_sys.sdg_user where " & _
    "sdg.sdg_id = sdg_user.sdg_id and sdg.sdg_id = " & Sdg("SDG_ID"))
21690     Set SdgStatusImage.Picture = LoadPicture("Resource\sdg" & _
    SdgStatus("STATUS") & ".ico")
21700     SdgStatus.Close
21710     Exit Sub
ErrHnd:
21720     Call ErrHandler("ChangeStatus")
End Sub

Private Sub UserControl_Terminate()



21730     On Error GoTo ErrHnd
          Dim i As Integer
21740     For i = 0 To SnomedCtrl.Count - 1
21750         SnomedCtrl(i).Terminate
21760     Next i

21770     If strHandle <> "" Then Call ReleaseHandle
          'ashi 06.11.2014 Close Sessions

         ' RunFromWindow = True
          
21780     If Not Con Is Nothing Then
21790         If RunFromWindow And Not Con.State = adStateClosed Then
21800              Con.Close
21810              Set Con = Nothing

                    
21820         End If
21830       Else
21840               If Not Con Is Nothing Then Con.Close
21850           Set Con = Nothing
21860     End If
         
         
              
           
21870     Exit Sub
ErrHnd:
21880     Call ErrHandler("Result_Entry_UserControl_Terminate")
End Sub

Private Sub UserControl_Initialize()
21890     RunFromWindow = False
End Sub

Private Sub UpdateRtfResult(RtfResultId As String, FreeTextCtrl As _
    FreeTextTemplateCtrl)
21900     On Error GoTo ErrHnd
          Dim RtfResult As ADODB.Recordset
          Dim lStream As ADODB.Stream
          
21910     Set RtfResult = _
    Con.Execute("select rtf_result_id from lims_sys.rtf_result " & _
    "where rtf_result_id = " & RtfResultId)
21920     If RtfResult.EOF Then
21930         Call _
    Con.Execute("insert into lims_sys.rtf_result (rtf_result_id) values (" & _
    RtfResultId & ")")
21940     End If
21950     RtfResult.Close
21960     Set RtfResult = New ADODB.Recordset
21970     Call _
    RtfResult.Open("select rtf_text from lims_sys.rtf_result where rtf_result_id = " _
    & RtfResultId, Con, adOpenStatic, adLockOptimistic)
          
          'Call RtfResult("RTF_TEXT").AppendChunk(FreeTextCtrl.GetRTFContent)
       
21980     Set lStream = New ADODB.Stream
21990     lStream.Charset = csHeBrEw
22000     lStream.Type = adTypeText
22010     lStream.Open

22020     lStream.WriteText FreeTextCtrl.GetRTFContent 'TxtFreeText.GetRTFContent
22030     lStream.Position = 0
22040     RtfResult("RTF_TEXT") = lStream.ReadText
22050     lStream.Close
22060     Set lStream = Nothing
          
22070     RtfResult.Update
22080     RtfResult.Close
22090     Set RtfResult = Nothing

22100     Exit Sub
ErrHnd:
22110     Call ErrHandler("UpdateRtfResult")
End Sub

Private Function ReadClob(pFld As ADODB.Field) As String
22120     On Error GoTo ErrHnd
          
          ' Function read a the clob data from the field
          '   using the stream object of the ADODB library
          
          Dim lStream As ADODB.Stream
          Dim lstData As String
          
22130     Set lStream = New ADODB.Stream
22140     lStream.Charset = csHeBrEw
22150     lStream.Type = adTypeText
22160     lStream.Open
          
22170     lStream.WriteText nte(pFld.value)
22180     lStream.Position = 0
22190     lstData = lStream.ReadText

22200     lStream.Close
22210     Set lStream = Nothing
          
22220     ReadClob = lstData
          
22230     Exit Function
ErrHnd:
22240     Call ErrHandler("ReadClob")
End Function


Private Sub CheckReportColorSlides(sdgId As Double, ResultName As String)
22250     On Error GoTo ErrHnd
          Dim ResultRec As ADODB.Recordset
          Dim strSQL As String

22260     strSQL = "select aliquot.name from lims_sys.result, " & _
    "lims_sys.test, lims_sys.aliquot, lims_sys.sample " & "where result.name = '" & _
    ResultName & "' And " & "result.status = 'V' and " & "sample.sdg_id = " & sdgId _
    & " and " & "aliquot.sample_id = sample.sample_id and " & _
    "test.aliquot_id = aliquot.aliquot_id and " & _
    "substr(aliquot.name,1,1)<>'P' and " & "result.test_id = test.test_id"

22270     Set ResultRec = Con.Execute(strSQL)
22280     While Not ResultRec.EOF
22290         MsgBox "The slide: " & ResultRec(0) & ", was not reported !", , _
    "Nautilus - Result Entry"
22300         ResultRec.MoveNext
22310     Wend
22320     ResultRec.Close
22330     Exit Sub
ErrHnd:
22340     Call ErrHandler("CheckReportColorSlides")
End Sub

Private Sub ErrHandler(pstSub As String)
          
          Dim lstErrMsg As String
          
MsgBox "EERR HANDLER1 1"
22350     lstErrMsg = "SUB: " & pstSub & vbCrLf & "Error On Line: " & Erl & _
          vbCrLf & "DESCRIPTION:" & vbCrLf & Err.Description & vbCrLf & _
                        "Please call support."
22360     MsgBox lstErrMsg, vbOKOnly + vbCritical, "ERROR - CALL IT !!!"
 
 MsgBox "EERR HANDLER 2"
22370     If RunFromWindow Then
22380         RaiseEvent CloseClicked
22390     Else
22400          If Not NtlsSite2 Is Nothing Then NtlsSite2.CloseWindow
22410     End If
          MsgBox "EERR HANDLER1 3"
22420     Call ReleaseApplicationMutex
MsgBox "EERR HANDLER 4"

End Sub

Public Sub InitiateSdg(sn As String)
22430     SdgName.Text = sn
22440     Call SdgName_KeyDown(vbKeyReturn, 0)
          'Call SdgName_KeyUp(vbKeyReturn, 0)
End Sub

Private Sub InsertNote()
          Dim DateFormatSyntax As String
          Dim MXRank As ADODB.Recordset
          Dim MNRank As ADODB.Recordset
          Dim MinRank As Integer
          Dim MaxRank As Integer
          Dim qc As Integer
22450     MinRank = 0
22460     MaxRank = 0
22470     qc = 0
22480     QcRank = 0
22490     On Error GoTo ErrHnd
22500     Call Con.Execute("insert into lims_sys.sdg_note " & _
    "(note_entry_id, sdg_id, entry_date, entry_type, " & _
    "session_id, subject, description) " & "values (lims.sq_note_entry.nextval, " & _
    Sdg("SDG_ID") & ", SYSDATE" & ", 'T'" & _
    ", lims.lims_env.session_id, 'Summary', '" & GetSummary & vbCrLf & _
    Trim(PFreeTextResult(1).GetContent) & "')")
          
22510     Call Con.Execute("insert into lims_sys.sdg_note " & _
    "(note_entry_id, sdg_id, entry_date, entry_type, " & _
    "session_id, subject, description) " & "values (lims.sq_note_entry.nextval, " & _
    Sdg("SDG_ID") & ", SYSDATE" & ", 'T'" & ", lims.lims_env.session_id, 'RANK', '" _
    & GetRankSum & "')")
          
22520     Set MXRank = Con.Execute("select max_sn.description as max_rank " & _
    "from lims_sys.sdg_note max_sn " & "Where max_sn.entry_date = " & _
    "(select max(sn1.entry_date) from lims_sys.sdg_note sn1 where " & _
    "max_sn.subject = sn1.subject and max_sn.sdg_id = sn1.sdg_id) " & _
    "and max_sn.subject='RANK' " & "and max_sn.sdg_id=" & Sdg("SDG_ID"))

22530     Set MNRank = Con.Execute("select min_sn.description as min_rank " & _
    "from lims_sys.sdg_note min_sn " & "Where min_sn.entry_date = " & _
    "(select min(sn1.entry_date) from lims_sys.sdg_note sn1 where " & _
    "min_sn.subject = sn1.subject and min_sn.sdg_id = sn1.sdg_id) " & _
    "and min_sn.subject='RANK' " & "and min_sn.sdg_id=" & Sdg("SDG_ID"))
              
22540     If Not MXRank.EOF Then
22550         If nte(MXRank("MAX_RANK")) <> "" Then
22560             MaxRank = CInt(nte(MXRank("MAX_RANK")))
22570         End If
22580     End If
22590     MXRank.Close
22600     If Not MNRank.EOF Then
22610         If nte(MNRank("MIN_RANK")) <> "" Then
22620             MinRank = CInt(nte(MNRank("MIN_RANK")))
22630         End If
22640     End If
22650     MNRank.Close
          
22660     qc = MaxRank - MinRank
      '    QCtxt.Text = Abs(qc)
22670     QcRank = Abs(qc)
      '    chkQC.Value = IIf(ntz(Sdg("U_ISQC")) = "T", 1, 0)
22680     chkCon.value = IIf(ntz(Sdg("U_ISCONSULT")) = "T", 1, 0)


          
22690     Exit Sub

ErrHnd:
22700     Call ErrHandler("InsertNote")
End Sub

Private Function GetSummary() As String
          Dim i
          Dim typ, index
22710     On Error GoTo ErrHnd
22720     GetSummary = ""
22730     For i = 1 To PResultIndex
22740         typ = Mid(PResultDesc(i).Tag, 1, 1)
22750         index = Val(Mid(PResultDesc(i).Tag, 2))
22760         If typ = "B" Then
22770             If PResultCheck(index).value = 1 Then
22780                 GetSummary = GetSummary & PResultDesc(i).Caption & vbCrLf
22790             End If
22800         ElseIf typ = "T" Then
22810             If PResultText(index).Text <> "" Then
22820                 GetSummary = GetSummary & PResultDesc(i).Caption & ": " & _
    PResultText(index).Text & vbCrLf
22830             End If
22840         ElseIf typ = "P" Then
22850             If PResultPhrase(index).getValue <> "" Then
22860                 GetSummary = GetSummary & PResultDesc(i).Caption & ": " & _
    PResultPhrase(index).getValue & vbCrLf
22870             End If
22880         End If
22890     Next i
22900     Exit Function
ErrHnd:
22910     Call ErrHandler("GetSummary")
End Function

Private Sub CalculateSnomed(SnomedCalculation As ADODB.Recordset, ResultName As _
    String)
          Dim i
          Dim typ, index
          Dim rValue As String
          
22920     On Error GoTo ErrHnd
              
22930     Set SnomedParser = SnomedCtrl(0).getParser
          
          'fixed not getting snomeds after "missing data" error
22940     If SnomedCalculation.RecordCount > 0 Then
22950         SnomedCalculation.MoveFirst
22960     End If
          
22970     Do Until SnomedCalculation.EOF
22980         SnomedParser.addPhrase nte(SnomedCalculation("DESCRIPTION")), _
    nte(SnomedCalculation("U_SNOMED_CODE"))
22990         SnomedCalculation.MoveNext
23000     Loop
23010     SnomedParser.CalculateSnomed
          
          
          'enter the new snomed in any case, even if calculation
          'gives us no snomed (05.06.2006 cancel of the IF statement):
          
      '    If SnomedParser.SnomedCodes <> "" Then
              
23020         For i = 1 To PResultIndex
23030             typ = Mid(PResultDesc(i).Tag, 1, 1)
23040             index = Val(Mid(PResultDesc(i).Tag, 2))
23050             If typ = "T" And UCase(PResultDesc(i).DataField) = _
    UCase(ResultName) Then
23060                 PResultText(index).Text = SnomedParser.SnomedCodes
23070                 Exit Sub
23080             End If
23090         Next i

      '    End If
          
23100     Set SnomedParser = Nothing
23110     Exit Sub
ErrHnd:
23120     Call ErrHandler("CalculateSnomed")
          
End Sub


Private Sub CalculateSnomeds()

23130     On Error GoTo ErrHnd
          
23140      CalculateSnomed SnomedMCalculation, "Snomed M"
23150      CalculateSnomed SnomedTCalculation, "Snomed T"
           

23160     Exit Sub
ErrHnd:
23170     Call ErrHandler("CalculateSnomeds")

End Sub

Private Function AssignPapInspection() As Boolean
23180     On Error GoTo ErrHnd
          Dim rand As Double
          Dim QCParameter As Integer
          Dim IsQC As Boolean
          Dim IsPos As Boolean
          Dim fqc As FrmQC
          Dim phrase As ADODB.Recordset
          Dim rstPreviousInspectors As ADODB.Recordset
          Dim phraseEntryName As String
          Dim inspectionPlanName As String
          Dim inspection As ADODB.Recordset
          Dim inspectionPlanId As String
          Dim PreviousInspectors As String
          Dim strRole As String
          Dim strPos As String
          Dim strQC As String
          Dim papqcid As ADODB.Recordset
          Dim oldInspection As Long
          Dim newInspection As Long
          Dim qc As String
          Dim sql As String

23190     qc = "F"

23200     AssignPapInspection = False
          
23210     sql = "   SELECT DISTINCT (REPLACE (SUBSTR (r.NAME, 1, 1), 'P', 'I')"
23220     sql = sql & _
    "                  ) AS role_letter_code,log.INSPECTION_DATE"
23230     sql = sql & _
    "             FROM lims_sys.inspection_log LOG, lims_sys.lims_role r"
23240     sql = sql & "            WHERE LOG.table_name = 'SDG'"
23250     sql = sql & "              AND LOG.table_key = " & Sdg("SDG_ID")
23260     sql = sql & "              AND r.role_id = LOG.role_id"
23270     sql = sql & "           order by log.INSPECTION_DATE asc"
          
          'gets the letter C-cytoscreener I-Pap Inspector D-doctor
          
23280     Set rstPreviousInspectors = Con.Execute(sql)
23290     While Not rstPreviousInspectors.EOF
23300         PreviousInspectors = PreviousInspectors & _
    nte(rstPreviousInspectors(0))
23310         rstPreviousInspectors.MoveNext
23320     Wend
         
23330     Select Case UCase(nte(Role("NAME")))
              Case UCase("doctor")
23340             strRole = "D"
23350         Case UCase("pap inspector")
23360             strRole = "I"
23370         Case UCase("cytoscreener")
23380             strRole = "C"
23390     End Select
          
23400     If CheckIsMalignant(Sdg("SDG_ID")) Or Inspect Then
23410         strPos = "Pos"
23420         IsPos = True
23430     Else
23440         strPos = "Neg"
23450         IsPos = False
23460     End If
           
23470     IsQC = False
          'for call 436 - A revision will not go through QC ( "And nte(Sdg("u_isqc")) = "" " )
23480     If PreviousInspectors = "" And strPos = "Neg" And Not IsConsult And _
    nte(Sdg("u_isqc")) = "" Then
23490         Randomize
23500         IsQC = Rnd < PQCParameter / CDbl(100)
23510     End If
          

          'Pos & Con=T & QC=F & Prev= & Role=C
23520     phraseEntryName = strPos & " & Con=" & btc(IsConsult) & " & QC=" & _
    btc(IsQC) & " & Prev=" & PreviousInspectors & " & Role=" & strRole

23530     Set phrase = _
    Con.Execute("select phrase_description, phrase_name from lims_sys.phrase_entry " _
    & "where phrase_id = (select phrase_id from lims_sys.phrase_header where " & _
    "name = 'Pap Inspection Params') " & _
    "and replace(upper(phrase_name),' ','') = replace(upper('" & phraseEntryName & _
    "'),' ','')")

23540     oldInspection = nte(Sdg("INSPECTION_PLAN_ID"))
23550     If Not phrase.EOF Then
23560         inspectionPlanName = UCase(nte(phrase("PHRASE_DESCRIPTION")))
23570         Set inspection = _
    Con.Execute("select inspection_plan_id from lims_sys.inspection_plan " & _
    "where upper(inspection_plan.name)= '" & inspectionPlanName & "'")
23580         Call Con.Execute("update lims_sys.sdg set inspection_plan_id = " _
    & inspection("inspection_plan_id") & "where sdg_id = " & Sdg("SDG_ID"))
              '----------------------------------------------------------------
              'hila ashi- call 1116- date 20.6.13- insert a new sdg_log recored each time inspection plan update
              'application code = "RE.AssignPapInspection"- new entry in phrase "sdg-log names"
              Dim sdg_log_description   As String
23590         sdg_log_description = "oldInspection " & oldInspection & _
    ", new inspectionPlanName " & inspection("inspection_plan_id") & _
    "  ,inspection plan phraseEntryName "
23600         sdg_log_description = sdg_log_description & phraseEntryName & _
    " , user-role " & NtlsUser.GetOperatorName & "-" & NtlsUser.GetRoleName
23610         Call sdg_log.InsertLog(Sdg("SDG_ID"), "RE.AssignPapInspection", _
    sdg_log_description)
              '----------------------------------------------------------------
23620     Else
23630         MsgBox "Inspection Plan was not updated." & vbCrLf & "Plan:  " & _
    phraseEntryName
23640     End If
          
          Dim RecInspectionLogCountByRole As ADODB.Recordset
          Dim RecInspectionPlanCount As ADODB.Recordset
          Dim InspectionLogCountByRole As Integer
          Dim InspectionPlanCount As Integer

          
23650     Set RecInspectionLogCountByRole = _
    Con.Execute("select count(*) as count from " & "lims_sys.inspection_log log " & _
    "where log.table_name = 'SDG' and " & "log.table_key = " & Sdg("SDG_ID") & _
    " and " & "log.role_id = " & NtlsUser.GetRoleId)
23660     InspectionLogCountByRole = _
    CInt(nte(RecInspectionLogCountByRole("COUNT")))
          
23670     If Not phrase.EOF Then
23680         inspectionPlanId = inspection("inspection_plan_id")
23690     Else
23700         inspectionPlanId = Sdg("INSPECTION_PLAN_ID")
23710     End If
23720     newInspection = inspectionPlanId
          
23730     Set RecInspectionPlanCount = _
    Con.Execute("select count(*) as count from lims_sys.inspection_entry entry " & _
    "where entry.inspection_plan_id = " & inspectionPlanId & " and " & _
    "entry.role_id = " & NtlsUser.GetRoleId)
23740     InspectionPlanCount = CInt(nte(RecInspectionPlanCount("COUNT")))

23750     If InspectionPlanCount <= InspectionLogCountByRole Then
23760         MsgBox "This request cannot be authorise by " & Role("NAME")
23770         Exit Function
23780     End If
          
          
23790     If IsQC Then
23800         Set fqc = New FrmQC
23810         Call fqc.Show(vbModal)
23820         If fqc.ConfirmSucceeded Then
23830             Call Con.Execute("update lims_sys.sdg_user set u_qc = '0' " & _
    "where sdg_id = " & Sdg("SDG_ID"))
23840             qc = "T"
23850         End If
23860         fqc.ConfirmSucceeded = False
23870         Set fqc = Nothing
23880     End If
          
          'update isqc only at the first time
23890     Con.Execute "update lims_sys.sdg_user " & "set u_isqc = '" & _
    btc(IsQC) & "' " & "where u_isqc is null and sdg_id = " & Sdg("SDG_ID")
23900     If IsConsult Then 'update only if true
23910         Con.Execute "update lims_sys.sdg_user " & "set u_isconsult = '" & _
    btc(IsConsult) & "' " & "where sdg_id = " & Sdg("SDG_ID")
23920     End If
          'update each time for the last value
23930     Con.Execute "update lims_sys.sdg_user " & "set u_ispositive = '" & _
    btc(IsPos) & "' " & "where sdg_id = " & Sdg("SDG_ID")
          ''''''''''''''''''''''''''''''''''''''''''
          ' for debug
          ''''''''''''''''''''''''''''''''''''''''''
23940     Set papqcid = Con.Execute("select lims.sq_u_papqc.nextval from dual")
23950     Con.Execute _
    "insert into lims_sys.u_papqc (u_papqc_id, name, version, version_status) " & _
    "values (" & papqcid(0) & ",'" & papqcid(0) & "','1','A')"
23960     Con.Execute _
    "insert into lims_sys.u_papqc_user (u_papqc_id, u_sdg_id, u_operator_id, " & _
    "u_phrase_entry, u_old_inspection_id, u_new_inspection_id, u_qc, u_created_on) " _
    & "values (" & papqcid(0) & "," & Sdg("SDG_ID") & "," & NtlsUser.GetOperatorId _
    & ",'" & phraseEntryName & "'," & oldInspection & "," & newInspection & ",'" & _
    qc & "',sysdate)"
          ''''''''''''''''''''''''''''''''''''''''''''''''''
          
23970     AssignPapInspection = True
23980     Exit Function
ErrHnd:
23990     Call ErrHandler("AssignPapInspection")
End Function

Private Function btc(b As Boolean) As String
24000     btc = IIf(b, "T", "F")
End Function


Private Sub SetFirstFocus()
          Dim i As Integer
          Dim typ, index
24010     On Error GoTo ErrHnd
24020     For i = 1 To PResultIndex
24030         typ = Mid(PResultDesc(i).Tag, 1, 1)
24040         index = Val(Mid(PResultDesc(i).Tag, 2))
24050         If typ = "T" Then
24060            If PResultText(index).Container.Visible = True Then
24070                Call PResultText(index).SetFocus
24080                Exit Sub
24090             End If
24100         End If
24110     Next i
24120     Exit Sub
ErrHnd:
24130     Call ErrHandler("SetFocusFirst")
End Sub

'Private Sub PapAuthoriseMsg()
'    Dim InspectionPlan As ADODB.Recordset
'    If Sdg("STATUS") <> "A" Then
'       Set InspectionPlan = con.Execute("select inspection_plan.name as name from " & '            "lims_sys.inspection_plan,lims_sys.sdg " & '            "where inspection_plan.inspection_plan_id = sdg.inspection_plan_id " & '            "and sdg.sdg_id = " & Sdg("SDG_ID"))
'        If InspectionPlan.EOF Then Exit Sub
'        If UCase(InspectionPlan("NAME")) = UCase("QC and Doctor") Or '            UCase(InspectionPlan("NAME")) = UCase("PQC") Then
'            MsgBox "This request should be rechecked by a Physician"
'        End If
'    End If
'End Sub
Private Function GetRankSum() As Integer
24140     On Error GoTo ErrHnd
          Dim Results As ADODB.Recordset
          
24150     GetRankSum = 0
24160     Set Results = Con.Execute("select sum(rd.u_renk) " & _
    "from lims_sys.u_result_desc_user rd, " & "lims_sys.result_template rt, " & _
    "lims_sys.result r, " & "lims_sys.test t, " & "lims_sys.aliquot a, " & _
    "lims_sys.sample s " & "Where t.aliquot_id = a.aliquot_id " & _
    "and a.sample_id = s.sample_id " & "and r.test_id = t.test_id " & _
    "and s.sdg_id = " & Sdg("SDG_ID") & " " & "and r.original_result = 'T' " & _
    "and r.result_template_id = rt.result_template_id " & _
    "and rt.name = rd.u_template_name")

24170     If Not Results.EOF Then
24180         If nte(Results(0)) <> "" Then
24190             GetRankSum = Results(0)
24200         End If
24210     End If
24220     Results.Close
24230     Exit Function

ErrHnd:
24240     Call ErrHandler("GetRankSum")

End Function

Private Sub InitReferrals()
          Dim phrase As ADODB.Recordset
24250     Set Ref = Nothing
24260     Set Ref = New Referrals.Referral
24270     Set phrase = _
    Con.Execute("select phrase_info, phrase_name from lims_sys.phrase_entry " & _
    "where phrase_id = (select phrase_id from lims_sys.phrase_header where " & _
    "name = 'Refferal Params')")

24280     While Not phrase.EOF
24290         If Trim(CStr(phrase("PHRASE_NAME").value)) = "match" Then
24300             Ref.DayIntervalForMatch = _
    Trim(CStr(phrase("PHRASE_INFO").value))
24310         ElseIf Trim(CStr(phrase("PHRASE_NAME").value)) = "select" Then
24320             Ref.DayIntervalForSelect = _
    Trim(CStr(phrase("PHRASE_INFO").value))
24330         ElseIf Trim(CStr(phrase("PHRASE_NAME").value)) = _
    "Connection String" Then
24340             Ref.ConnectionString = Trim(CStr(phrase("PHRASE_INFO").value))
24350         End If
24360         phrase.MoveNext
24370     Wend

24380     Ref.RequestType = Right(Sdg("EXTERNAL_REFERENCE"), 1)
24390     Ref.requestNumber = Mid(Sdg("EXTERNAL_REFERENCE"), 1, 10)
          
24400     PropsReferralDiagnose(1).Tag = ""
24410     PropsReferralDiagnose(1).Text = ""
24420     PropsReferralDiagnose(2).Tag = ""
24430     PropsReferralDiagnose(2).Text = ""
          
End Sub


'check if there is already an instance of this application
'running at this station, by trying to get a hold of the semaphore:
Private Function IsFirstApplicationInstance() As Boolean
      'On Error GoTo ERR_IsFirstApplicationInstance

24440     IsFirstApplicationInstance = False

          'to handle the mutex:
          Dim X As SECURITY_ATTRIBUTES
          Dim lWaitAnswer As Long
          'Dim Atom As Integer
                 
          'create the mutex and wait for access:
24450     lMutexHandle = CreateMutex(X, True, "RESULT_ENTRY")
24460     If (Err.LastDllError <> ERROR_ALREADY_EXISTS) Then
24470        IsFirstApplicationInstance = True
24480     End If
      '    lWaitAnswer = WaitForSingleObject(lMutexHandle, 10)
      '
      '    'check if got the mutex:
      '    If lWaitAnswer = 0 Then
      '        IsFirstApplicationInstance = True
      '    End If
          
      '    Atom = GlobalFindAtom(MyAtomName)
      '    If Atom = 0 And lWaitAnswer = 0 Then
      '        CurrAtom = GlobalAddAtom(MyAtomName)
      '        IsFirstApplicationInstance = IsFirstApplicationInstance And True
      '    Else
      '        IsFirstApplicationInstance = False
      '    End If
          
24490     Exit Function
ERR_IsFirstApplicationInstance:
24500 MsgBox "ERR_IsFirstApplicationInstance" & vbCrLf & Err.Description
End Function

'release the mutex so a new instance of Result Entry
'could be opened on this workstation:
Private Sub ReleaseApplicationMutex()
24510 On Error GoTo ERR_ReleaseApplicationMutex

24520     ReleaseMutex (lMutexHandle)
24530     CloseHandle (lMutexHandle)
      '    GlobalDeleteAtom CurrAtBom
           
24540     Exit Sub
ERR_ReleaseApplicationMutex:
24550 MsgBox "ERR_ReleaseApplicationMutex" & vbCrLf & Err.Description
End Sub


'report all slides of this request to be moved
'into the tissue archive:
Private Sub UpdateTransferalToTissueArchiveHis(strSdgId As String)
24560 On Error GoTo ERR_UpdateTransferalToTissueArchive
          Dim sql As String

24570     sql = " update lims_sys.aliquot_user au"
24580     sql = sql & " set au.U_ARCHIVE='T'"
24590     sql = sql & " where au.ALIQUOT_ID in"
24600     sql = sql & " ("
24610     sql = sql & "   select a.ALIQUOT_ID"
24620     sql = sql & "   from lims_sys.aliquot a, "
24630     sql = sql & "        lims_sys.sample s"
24640     sql = sql & "   where exists"
24650     sql = sql & "   ("
24660     sql = sql & "      select aliquot_id"
24670     sql = sql & "      from lims_sys.aliquot_formulation"
24680     sql = sql & "      where child_aliquot_id = a.ALIQUOT_ID "
24690     sql = sql & "   )"
24700     sql = sql & "   and a.SAMPLE_ID = s.SAMPLE_ID"
24710     sql = sql & "   and s.SDG_ID = " & strSdgId
24720     sql = sql & " )"
          
24730     Call Con.Execute(sql)
          
24740     Exit Sub
ERR_UpdateTransferalToTissueArchive:
24750 MsgBox "ERR_UpdateTransferalToTissueArchive" & vbCrLf & Err.Description
End Sub

'report all slides of this request to be moved
'into the tissue archive:
Private Sub UpdateTransferalToTissueArchiveCytoPap(strSdgId As String)
24760 On Error GoTo ERR_UpdateTransferalToTissueArchive
          Dim sql As String

24770     sql = " update lims_sys.aliquot_user au"
24780     sql = sql & " set au.U_ARCHIVE='T'"
24790     sql = sql & " where au.ALIQUOT_ID in"
24800     sql = sql & " ("
24810     sql = sql & "   select a.ALIQUOT_ID"
24820     sql = sql & "   from lims_sys.aliquot a, "
24830     sql = sql & "        lims_sys.sample s"
24840     sql = sql & "   where a.SAMPLE_ID = s.SAMPLE_ID"
24850     sql = sql & "   and s.SDG_ID = " & strSdgId
24860     sql = sql & " )"
          
24870     Call Con.Execute(sql)
          
24880     Exit Sub
ERR_UpdateTransferalToTissueArchive:
24890 MsgBox "ERR_UpdateTransferalToTissueArchive" & vbCrLf & Err.Description
End Sub


'we shoe that grid on the buttom-left corner of the screen.
'one row containing all the BLOCKS of the request.
'for each block:
'1. + / - if there's material left
'2. number of tissues
'3. the aliquot name
'the grid is shown OVER the status bar. when the status bar is needed
'the grid goes down and becomes invisible
Private Sub InitAliquotGrid()
24900 On Error GoTo ERR_InitAliquotGrid
      'On Error Resume Next
          Dim X As Integer
          Dim Y As Integer
          Dim iWidth As Integer
          Dim col As Integer
          Dim sql As String
          Dim rs As Recordset
          Dim rsa As Recordset
          Dim s As String
          
      '    sql = " select su.U_MATERIAL, au.U_NUM_OF_TISSUES, a.name"
      '    sql = sql & " from lims_sys.aliquot a,"
      '    sql = sql & "      lims_sys.aliquot_user au, "
      '    sql = sql & "      lims_sys.sample s,"
      '    sql = sql & "   lims_sys.sample_user su"
      '    sql = sql & " where s.SDG_ID = " & Sdg("sdg_id")
      '    sql = sql & " and   s.SAMPLE_ID = su.SAMPLE_ID"
      '    sql = sql & " and   a.SAMPLE_ID = s.SAMPLE_ID"
      '    sql = sql & " and   a.ALIQUOT_ID = au.ALIQUOT_ID"
      '    sql = sql & " and   not exists"
      '    sql = sql & " ("
      '    sql = sql & "   select child_aliquot_id"
      '    sql = sql & "   from lims_sys.aliquot_formulation"
      '    sql = sql & "   where child_aliquot_id = a.ALIQUOT_ID"
      '    sql = sql & " )"
      '    sql = sql & " order by a.aliquot_id"
          
          
24910     s = " select su.u_material, s.sample_id, s.name "
24920     s = s & " from lims_sys.sample s, "
24930     s = s & " lims_sys.sample_user su "
24940     s = s & " where s.sample_id = su.sample_id "
24950     s = s & " and s.sdg_id = " & Sdg("sdg_id")
24960     s = s & " and s.status not in ('X','U','R') "
24970     s = s & " order by su.u_order, s.sample_id "

24980     Set rs = Con.Execute(s)
          
          
          'needed if peviously presented with a scroll
          'and was scrolled right (then we won't see the left cells)
24990     gridAliquots.Clear
25000     gridAliquots.Rows = 0
25010     gridAliquots.Cols = 0
25020     gridAliquots.Width = 0
              
25030     gridAliquots.AllowBigSelection = False
25040     gridAliquots.Enabled = True

25050     gridAliquots.ScrollBars = flexScrollBarNone
25060     gridAliquots.SelectionMode = flexSelectionFree
25070     gridAliquots.AllowUserResizing = flexResizeBoth

25080     gridAliquots.Rows = 1
25090     gridAliquots.Cols = 2
25100     gridAliquots.RowHeight(0) = 240
25110     gridAliquots.Height = 240
25120     gridAliquots.Width = 10695
          
25130     gridAliquots.FixedRows = 0
25140     gridAliquots.FixedCols = 0

25150     gridAliquots.row = 0
          
25160     X = 0
25170     col = 0
25180     iWidth = 0

25190     While Not rs.EOF
25200         col = col + 1
25210         X = X + 1
25220         s = Mid(nte(rs("NAME")), 12)
25230         s = s & IIf(nte(rs("u_material")) = "T", "+", "-")
              's = X & IIf(nte(rs("u_material")) = "T", "+", "-")
25240         gridAliquots.Cols = col
25250         gridAliquots.col = col - 1
25260         gridAliquots.ColWidth(col - 1) = 400
25270         iWidth = iWidth + gridAliquots.ColWidth(col - 1)
25280         gridAliquots.CellAlignment = vbAlignLeft
25290         gridAliquots.Text = s
25300         gridAliquots.CellBackColor = &HC0FFC0

25310         sql = " select au.U_NUM_OF_TISSUES, a.name "
25320         sql = sql & " from lims_sys.aliquot a, "
25330         sql = sql & "      lims_sys.aliquot_user au "
25340         sql = sql & " where a.SAMPLE_ID = " & rs("sample_id")
25350         sql = sql & " and   a.ALIQUOT_ID = au.ALIQUOT_ID"
25360         sql = sql & " and   a.status not in ('X','U','R') "
25370         sql = sql & " and   not exists"
25380         sql = sql & " ("
25390         sql = sql & "   select child_aliquot_id"
25400         sql = sql & "   from lims_sys.aliquot_formulation"
25410         sql = sql & "   where child_aliquot_id = a.ALIQUOT_ID"
25420         sql = sql & " )"
25430         sql = sql & " order by a.aliquot_id"
              
25440         Set rsa = Con.Execute(sql)

25450         Y = 0
25460         While Not rsa.EOF
25470             col = col + 1
25480             Y = Y + 1
25490             s = Mid(nte(rsa("NAME")), 12)
25500             s = s & "(" & nte(rsa("U_NUM_OF_TISSUES")) & ")"
25510             gridAliquots.Cols = col
25520             gridAliquots.col = col - 1
25530             gridAliquots.ColWidth(col - 1) = 1200
25540             iWidth = iWidth + gridAliquots.ColWidth(col - 1)
25550             gridAliquots.CellAlignment = vbLeftJustify
25560             gridAliquots.Text = s
25570             rsa.MoveNext
25580         Wend

25590         rs.MoveNext
25600     Wend
       
      '    While Not rs.EOF
      '        s = IIf(nte(rs("u_material")) = "T", "+", "-")
      '        s = s & ntz(rs("u_num_of_tissues")) & "-"
      '        s = s & Mid(nte(rs("NAME")), 12)
      '
      '        x = x + 1
      '        gridAliquots.Cols = x
      '        gridAliquots.col = x - 1
      '        gridAliquots.ColWidth(x - 1) = 1000
      '        gridAliquots.Text = s
      '
      '        rs.MoveNext
      '    Wend
          
          'scrolling is needed:
          'make place for the scroll bar and show it
25610     If iWidth > gridAliquots.Width Then
25620         gridAliquots.Height = 480
25630         gridAliquots.ScrollBars = flexScrollBarHorizontal
25640     End If
          
          'show the grid OVER the status bar:
25650     gridAliquots.Top = lblStatusBar.Top
25660     gridAliquots.Visible = True
          
25670     Exit Sub
ERR_InitAliquotGrid:

25680 Select Case Err.Number
      Case 6
25690     Resume Next 'overflow - grid issue
25700 Case Else
25710     MsgBox "ERR_InitAliquotGrid" & vbCrLf & Err.Description & vbCrLf & _
    Err.Number
25720 End Select

      'If iWidth > gridAliquots.Width Then
      '    gridAliquots.Height = 480
      '    gridAliquots.ScrollBars = flexScrollBarHorizontal
      'End If
End Sub


'create & lock a semaphore
Private Function AllocateHandle(strSemaphoreName As String) As Boolean
25730 On Error GoTo ERR_AllocateHandle
          
          Dim cmd As New ADODB.Command
          Dim param As ADODB.Parameter
          Dim rs As ADODB.Recordset
          Dim strResult As String
          
          
25740     AllocateHandle = False
          
25750     Set cmd.ActiveConnection = Con
25760     cmd.CommandType = adCmdStoredProc
25770     cmd.CommandText = "LIMS.SEMAPHORE.ALLOCATE_SEMAPHORE"  '"LIMS$ALLOCATE_SEMAPHORE"
          
25780     Set param = New ADODB.Parameter
25790     param.name = "semaphore_name"
25800     param.Type = adVarChar
25810     param.value = strSemaphoreName
25820     param.Size = 50
25830     param.Direction = adParamInput
25840     Call cmd.parameters.Append(param)
          
25850     Set param = New ADODB.Parameter
25860     param.name = "timeout"
25870     param.Type = adInteger
25880     param.value = 2
      '    param.Size = 50
25890     param.Direction = adParamInput
25900     Call cmd.parameters.Append(param)
          
25910     Set param = New ADODB.Parameter
25920     param.name = "retval"
25930     param.Type = adVarChar
25940     param.value = "xxx"
25950     param.Size = 50
25960     param.Direction = adParamReturnValue
25970     Call cmd.parameters.Append(param)

25980     Call cmd.Execute

25990     strResult = CStr(cmd.parameters.Item(2))
26000     Select Case strResult
              Case "Timeout"
26010             Call PrintSemaphoreError(strResult, strSemaphoreName)
26020         Case "Deadlock"
26030             Call PrintSemaphoreError(strResult, strSemaphoreName)
26040         Case "Parameter Error"
26050             Call PrintSemaphoreError(strResult, strSemaphoreName)
26060         Case "Illegal Lock Handle"
26070             Call PrintSemaphoreError(strResult, strSemaphoreName)
26080         Case Else
26090             strHandle = strResult
26100             AllocateHandle = True
26110     End Select

26120     Exit Function
ERR_AllocateHandle:
26130 MsgBox "ERR_AllocateHandle" & vbCrLf & Err.Description
End Function

'release a semaphore held by me:
Private Function ReleaseHandle() As Boolean
26140 On Error GoTo ERR_ReleaseHandle

         
          Dim cmd As New ADODB.Command
          Dim param As ADODB.Parameter
          Dim rs As ADODB.Recordset
          Dim strResult As String
          
26150     Set cmd.ActiveConnection = Con
26160     cmd.CommandType = adCmdStoredProc
26170     cmd.CommandText = "LIMS.SEMAPHORE.RELEASE_SEMAPHORE"
          
26180     Set param = New ADODB.Parameter
26190     param.name = "lockhandle"
26200     param.Type = adVarChar
26210     param.value = strHandle
26220     param.Size = 50
26230     param.Direction = adParamInput
26240     Call cmd.parameters.Append(param)

26250     Set param = New ADODB.Parameter
26260     param.name = "retval"
26270     param.Type = adVarChar
26280     param.value = "xxx"
26290     param.Size = 50
26300     param.Direction = adParamReturnValue
26310     Call cmd.parameters.Append(param)
          
26320     Call cmd.Execute
26330     strResult = CStr(cmd.parameters.Item(1))
          
26340     Select Case strResult
              Case "Timeout"
                  'Call PrintSemaphoreError(strResult)
26350         Case "Deadlock"
                  'Call PrintSemaphoreError(strResult)
26360         Case "Parameter Error"
                  'Call PrintSemaphoreError(strResult)
26370         Case "Illegal Lock Handle"
                  'Call PrintSemaphoreError(strResult)
26380         Case "Do not own lock"
                  
26390         Case Else
26400             ReleaseHandle = True
26410     End Select
          
26420     Exit Function
ERR_ReleaseHandle:
26430 MsgBox "ERR_ReleaseHandle" & vbCrLf & Err.Description
End Function

'called in case we can't use the resource we wanted:
Private Sub PrintSemaphoreError(strMsg As String, strSemaphoreName As String)
          Dim strUser As String
26440     strUser = GetSemaphoreOwner(strSemaphoreName)
          
26450     MsgBox "Cannot view request, " & vbCrLf & _
    "it may be opened at another station" & vbCrLf & strUser & vbCrLf & _
    "DB Message: " & strMsg
End Sub

'get details of the owner of a semaphore;
'fails if there is no lock with such semaphore
Private Function GetSemaphoreOwner(strSemaphoreName) As String
26460 On Error GoTo ERR_GetSemaphoreOwner
          
          Dim cmd As New ADODB.Command
          Dim param As ADODB.Parameter
          Dim rs As ADODB.Recordset
          Dim strResult As String
          
          
26470     Set cmd.ActiveConnection = Con
26480     cmd.CommandType = adCmdStoredProc
26490     cmd.CommandText = "LIMS.SEMAPHORE.GET_OWNER"  '"LIMS$ALLOCATE_SEMAPHORE"
          
26500     Set param = New ADODB.Parameter
26510     param.name = "semaphore_name"
26520     param.Type = adVarChar
26530     param.value = strSemaphoreName
26540     param.Size = 50
26550     param.Direction = adParamInput
26560     Call cmd.parameters.Append(param)
          
26570     Set param = New ADODB.Parameter
26580     param.name = "user_name"
26590     param.Type = adVarChar
26600     param.value = ""
26610     param.Size = 50
26620     param.Direction = adParamReturnValue
26630     Call cmd.parameters.Append(param)
          
26640     Set param = New ADODB.Parameter
26650     param.name = "computer_name"
26660     param.Type = adVarChar
26670     param.value = ""
26680     param.Size = 50
26690     param.Direction = adParamReturnValue
26700     Call cmd.parameters.Append(param)

26710     Call cmd.Execute

26720     strResult = "User Name: " & CStr(cmd.parameters.Item(1)) & _
    ", Station Name: " & CStr(cmd.parameters.Item(2))
               
26730     GetSemaphoreOwner = strResult
               
26740     Exit Function
ERR_GetSemaphoreOwner:
26750 MsgBox "ERR_GetSemaphoreOwner" & vbCrLf & Err.Description
End Function


'paint the extra-requests button red
'if there are extra requests to this SDG:
Private Sub SignalExtraRequest(strExternalRef As String)
26760 On Error GoTo ERR_SignalExtraRequest
          Dim rs As Recordset
          Dim sql As String

26770     sql = " select 1 "
26780     sql = sql & " from lims_sys.u_extra_request_user ru, "
26790     sql = sql & "      lims_sys.sdg d "
26800     sql = sql & " where ru.U_SDG_ID=d.sdg_id "
          
          'sql = sql & " and   d.external_reference='" & strExternalRef & "'"
           'change this to name in patholab
26810     sql = sql & "  and substr(d.name,1,10) ='" & strExternalRef & "'  "
          
26820     sql = sql & " and   ru.U_STATUS <> 'X' "

26830     Set rs = Con.Execute(sql)
          
26840     If Not rs.EOF Then
26850         cmdAdditionalActions.BackColor = vbRed
26860     Else
26870         cmdAdditionalActions.BackColor = vbButtonFace
26880     End If

26890     Exit Sub
ERR_SignalExtraRequest:
26900 MsgBox "ERR_SignalExtraRequest" & vbCrLf & Err.Description
End Sub


'get the nth operator to sign on the request:
Private Function AuthorizedBy(strSdgId As String, nNumberOfAuthorization) As _
    String
26910 On Error GoTo ERR_AuthorizedBy
          Dim rs As Recordset
          Dim sql As String

26920     sql = " select o.FULL_NAME "
26930     sql = sql & " from lims_sys.operator o"
26940     sql = sql & " where o.OPERATOR_ID=lims.AUTHORIZATION.signed_by"
26950     sql = sql & "('" & strSdgId & "'," & nNumberOfAuthorization & ")"
          
26960     Set rs = Con.Execute(sql)
          
26970     If Not rs.EOF Then
26980         AuthorizedBy = nte(rs("FULL_NAME"))
26990     End If

27000     Exit Function
ERR_AuthorizedBy:
27010 MsgBox "ERR_AuthorizedBy" & vbCrLf & Err.Description
End Function


Private Function getNextStr(ByRef s As String, c As String)
          Dim p
          Dim res
27020     p = InStr(1, s, c)
27030     If (p = 0) Then
27040         res = s
27050         s = ""
27060         getNextStr = res
27070     Else
27080         res = Mid$(s, 1, p - 1)
27090         s = Mid$(s, p + Len(c), Len(s))
27100         getNextStr = res
27110     End If
End Function


Private Function GetSnomedTResultIndex() As Integer
27120 On Error GoTo ERR_GetSnomedTResultIndex

          Dim i As Integer
          Dim typ, index

27130     GetSnomedTResultIndex = -1

27140     For i = 1 To PResultIndex
              
27150         typ = Mid(PResultDesc(i).Tag, 1, 1)
27160         index = Val(Mid(PResultDesc(i).Tag, 2))
              
27170         If typ = "T" And UCase(PResultDesc(i).DataField) = _
    UCase("Snomed T") Then
                  
27180             GetSnomedTResultIndex = index
27190             Exit Function
                  
27200         End If
              
27210     Next i

27220     Exit Function
ERR_GetSnomedTResultIndex:
27230 MsgBox "ERR_GetSnomedTResultIndex" & vbCrLf & Err.Description
End Function


'Get the list of snomeds T from Organ Ctrl;
'insert that data in the relevant locations:
Private Sub SetOrgansSnomedT()
27240 On Error GoTo ERR_SetOrgansSnomedT

          Dim i As Integer
          Dim iSnomedTIndex As Integer
          
27250     iSnomedTIndex = GetSnomedTResultIndex
          
27260     If OrganCtrl.SnomedT <> "" Then
27270         Call AddSnomedTItem(iSnomedTIndex, OrganCtrl.SnomedT)
27280     End If

27290     Exit Sub
ERR_SetOrgansSnomedT:
27300 MsgBox "ERR_SetOrgansSnomedT" & vbCrLf & Err.Description
End Sub


Private Sub AddSnomedTItem(iSnomedTIndex As Integer, strSnomedT As String)
27310 On Error GoTo ERR_AddSnomedMItem
          
27320     PResultText(iSnomedTIndex).Text = strSnomedT

27330     Exit Sub
ERR_AddSnomedMItem:
27340 MsgBox "ERR_AddSnomedTItem" & vbCrLf & Err.Description
End Sub

'Private Function IsInRevision(SdgId As String) As Boolean
'    Dim rst As Recordset
'
'    Set rst = Con.Execute("select 1 from lims_sys.u_letter_control_user" & '        " where u_sdg_id = " & SdgId & '        " and u_grp_code = '11'")
'
'    If rst.EOF Then
'        IsInRevision = False
'    Else
'        IsInRevision = True
'    End If
'End Function

Private Function GetInspection()
          Dim rst As ADODB.Recordset
          
27350     Set rst = _
    Con.Execute("select i.name from lims_sys.sdg d, lims_sys.inspection_plan i " & _
    "where d.inspection_plan_id = i.inspection_plan_id " & "and d.sdg_id = " & _
    Sdg("SDG_ID"))
          'nautilus update
27360     GetInspection = nte(rst(0).value)
End Function

'---------------------------------------------------------------------------------------
' Procedure : multipleSlidesValidation
' Author    : Yonatan Amir
' Date      : 02/08/2011
' Purpose   : Display message for multiple PAP slides
'---------------------------------------------------------------------------------------
'
Private Function multipleSlidesValidation(sdgId) As Boolean

          Dim aliquotAmountQuery As String
27370     aliquotAmountQuery = "select count ( aliquot_id ) amount " & _
    "From lims_sys.sample s, lims_sys.aliquot a " & _
    "Where a.sample_id = s.sample_id and s.sdg_id = '" & sdgId & "' " & _
    "group by s.sdg_id"
          
          Dim aliquotAmountRecordset As ADODB.Recordset
27380     Set aliquotAmountRecordset = Con.Execute(aliquotAmountQuery)
          
27390     If CInt(aliquotAmountRecordset("AMOUNT")) > 1 Then
27400         Select Case MsgBox("שים\י לב: קיימים מספר סליידים לדרישה", _
    vbOKCancel Or vbExclamation Or vbMsgBoxRight Or vbMsgBoxRtlReading Or _
    vbDefaultButton1, "סליידים מרובים")
              
                  Case vbOK
27410                 multipleSlidesValidation = True
                      
27420             Case vbCancel
27430                 multipleSlidesValidation = False
                      
27440         End Select
          
27450     Else: multipleSlidesValidation = True
          
27460     End If
          
End Function
'ashi Assuta interface
Private Sub ShowAssutaPdfCtrl1_Opened(desc As String)
27470 Call sdg_log.InsertLog(Sdg("SDG_ID"), "ResultEntryAttached.OPEN", desc)
End Sub
Private Sub ShowAssutaPdfCtrl1_Closed(IsRead As Boolean, desc As String)
27480 If (IsRead) Then
27490  Call sdg_log.InsertLog(Sdg("SDG_ID"), "ResultEntryAttached.CLOSE", desc)
27500 Else
27510  Call sdg_log.InsertLog(Sdg("SDG_ID"), "ResultEntryAttached.CLOSE", desc)
27520 End If
End Sub
'end Assuta interface
