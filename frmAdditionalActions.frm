VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmAdditionalActions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "בקשות חוזרות"
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8130
   ScaleWidth      =   15285
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImgLst 
      Left            =   5160
      Top             =   6840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.TreeView tree 
      Height          =   7935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   13996
      _Version        =   393217
      Indentation     =   706
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ImgLst"
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7935
      Left            =   5400
      TabIndex        =   1
      Top             =   120
      Width           =   10245
      _ExtentX        =   18071
      _ExtentY        =   13996
      _Version        =   393216
      Tabs            =   6
      Tab             =   4
      TabHeight       =   520
      BackColor       =   -2147483638
      TabCaption(0)   =   "צביעות"
      TabPicture(0)   =   "frmAdditionalActions.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "frmNewSlides"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdDrag"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdOKColors"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "frmExistingSlides"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdOKAddBlock"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "הצגת סליידים לרופא"
      TabPicture(1)   =   "frmAdditionalActions.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdOKSlidesFromArchive"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame6"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "ReEmbedding"
      TabPicture(2)   =   "frmAdditionalActions.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "picBlockReembedding"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame2"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cmdOKReEmbedding"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label7"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label8"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label9"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Label2"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "מסירה נוספת"
      TabPicture(3)   =   "frmAdditionalActions.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "picShowSamples"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Frame3"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "cmdOKShowOriginalSample"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Label11"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "שליחה להתיעצות"
      TabPicture(4)   =   "frmAdditionalActions.frx":0070
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "cd"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Frame5"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "cmdExistingLetters"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Frame8"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "Frame7"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "cmdOKAdvisors"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).ControlCount=   6
      TabCaption(5)   =   "היסטורית בקשות חוזרות"
      TabPicture(5)   =   "frmAdditionalActions.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "txtRequest"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "txtCreatedBy"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "txtCreaedOn"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "txtRequestRemarks"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).Control(4)=   "grid"
      Tab(5).Control(4).Enabled=   0   'False
      Tab(5).Control(5)=   "Label17"
      Tab(5).Control(5).Enabled=   0   'False
      Tab(5).Control(6)=   "Label18"
      Tab(5).Control(6).Enabled=   0   'False
      Tab(5).Control(7)=   "Label19"
      Tab(5).Control(7).Enabled=   0   'False
      Tab(5).Control(8)=   "Label20"
      Tab(5).Control(8).Enabled=   0   'False
      Tab(5).ControlCount=   9
      Begin VB.CommandButton cmdOKAddBlock 
         BackColor       =   &H0000C000&
         Caption         =   "יצירת CELL - BLOCK"
         Height          =   615
         Left            =   -66600
         MaskColor       =   &H0000C0C0&
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   94
         ToolTipText     =   "Please select a sample from the tree on the left"
         Top             =   1200
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin VB.Frame frmExistingSlides 
         BackColor       =   &H8000000A&
         Caption         =   "רזרבות (סה""כ 0)"
         Height          =   1830
         Left            =   -74520
         RightToLeft     =   -1  'True
         TabIndex        =   75
         Top             =   3495
         Width           =   7815
         Begin VB.PictureBox picSlides 
            BackColor       =   &H8000000A&
            Height          =   1215
            Left            =   120
            RightToLeft     =   -1  'True
            ScaleHeight     =   1155
            ScaleWidth      =   7515
            TabIndex        =   76
            Top             =   480
            Width           =   7575
            Begin VB.VScrollBar VScrollSlides 
               Height          =   1155
               Left            =   7260
               TabIndex        =   81
               Top             =   0
               Visible         =   0   'False
               Width           =   255
            End
            Begin VB.PictureBox picSlidesEntry 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000A&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   120
               ScaleHeight     =   375
               ScaleWidth      =   6975
               TabIndex        =   77
               Top             =   120
               Width           =   6975
               Begin VB.TextBox txtSlide 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Index           =   0
                  Left            =   2280
                  RightToLeft     =   -1  'True
                  TabIndex        =   80
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   2415
               End
               Begin VB.CommandButton cmdSlideReset 
                  Height          =   315
                  Index           =   0
                  Left            =   1440
                  RightToLeft     =   -1  'True
                  Style           =   1  'Graphical
                  TabIndex        =   79
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   495
               End
               Begin VB.CommandButton cmdSlide 
                  Height          =   315
                  Index           =   0
                  Left            =   5040
                  RightToLeft     =   -1  'True
                  TabIndex        =   78
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   855
               End
            End
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            Caption         =   "סלייד"
            Height          =   255
            Left            =   5400
            RightToLeft     =   -1  'True
            TabIndex        =   84
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            Caption         =   "צביעה"
            Height          =   255
            Left            =   3600
            RightToLeft     =   -1  'True
            TabIndex        =   83
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Caption         =   "איפוס"
            Height          =   255
            Left            =   1560
            RightToLeft     =   -1  'True
            TabIndex        =   82
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "הערות"
         Height          =   2175
         Left            =   -74520
         RightToLeft     =   -1  'True
         TabIndex        =   73
         Top             =   5520
         Width           =   7815
         Begin VB.TextBox txtColors 
            Alignment       =   1  'Right Justify
            Height          =   1575
            Left            =   360
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   74
            Top             =   360
            Width           =   7095
         End
      End
      Begin VB.PictureBox picBlockReembedding 
         Height          =   3015
         Left            =   -74400
         ScaleHeight     =   2955
         ScaleWidth      =   7515
         TabIndex        =   66
         Top             =   1440
         Width           =   7575
         Begin VB.VScrollBar VScrollBlocksReembedding 
            Height          =   2955
            LargeChange     =   20
            Left            =   7260
            SmallChange     =   4
            TabIndex        =   72
            Top             =   0
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.PictureBox picBlockReembeddingEntry 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   2895
            Left            =   120
            ScaleHeight     =   2895
            ScaleWidth      =   6975
            TabIndex        =   67
            Top             =   120
            Width           =   6975
            Begin VB.ComboBox cmbReembeddingReason 
               Height          =   315
               Index           =   0
               ItemData        =   "frmAdditionalActions.frx":00A8
               Left            =   3480
               List            =   "frmAdditionalActions.frx":00AA
               RightToLeft     =   -1  'True
               TabIndex        =   70
               Top             =   0
               Visible         =   0   'False
               Width           =   1695
            End
            Begin VB.CheckBox chkBlockReembeding 
               Alignment       =   1  'Right Justify
               Height          =   315
               Index           =   0
               Left            =   5280
               RightToLeft     =   -1  'True
               TabIndex        =   69
               Top             =   0
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.TextBox txtReembeddingDetails 
               Alignment       =   1  'Right Justify
               Height          =   315
               Index           =   0
               Left            =   0
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   68
               Top             =   0
               Visible         =   0   'False
               Width           =   3255
            End
            Begin VB.Label lblSampleReembedding 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   0
               Left            =   6360
               RightToLeft     =   -1  'True
               TabIndex        =   71
               Top             =   0
               Visible         =   0   'False
               Width           =   495
            End
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "הערות"
         Height          =   2775
         Left            =   -74400
         RightToLeft     =   -1  'True
         TabIndex        =   64
         Top             =   4920
         Width           =   7575
         Begin VB.TextBox txtReembedding 
            Alignment       =   1  'Right Justify
            Height          =   1935
            Left            =   480
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   65
            Top             =   480
            Width           =   6615
         End
      End
      Begin VB.PictureBox picShowSamples 
         Height          =   3015
         Left            =   -74400
         ScaleHeight     =   2955
         ScaleWidth      =   7515
         TabIndex        =   60
         Top             =   1440
         Width           =   7575
         Begin VB.VScrollBar VScrollShowSamples 
            Height          =   2955
            Left            =   7260
            TabIndex        =   63
            Top             =   0
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.PictureBox picShowSamplesEntry 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   120
            ScaleHeight     =   375
            ScaleWidth      =   6975
            TabIndex        =   61
            Top             =   120
            Width           =   6975
            Begin VB.CheckBox chkShowSample 
               Alignment       =   1  'Right Justify
               Caption         =   "Check1"
               Height          =   315
               Index           =   0
               Left            =   5400
               RightToLeft     =   -1  'True
               TabIndex        =   62
               Top             =   0
               Visible         =   0   'False
               Width           =   1095
            End
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "הערות"
         Height          =   2775
         Left            =   -74400
         RightToLeft     =   -1  'True
         TabIndex        =   58
         Top             =   4920
         Width           =   7575
         Begin VB.TextBox txtShowOriginalSample 
            Alignment       =   1  'Right Justify
            Height          =   1935
            Left            =   480
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   59
            Top             =   480
            Width           =   6615
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "הערות"
         Height          =   2775
         Left            =   -74400
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   4920
         Width           =   7575
         Begin VB.TextBox txtSlidesFromArchive 
            Alignment       =   1  'Right Justify
            Height          =   1935
            Left            =   480
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   57
            Top             =   480
            Width           =   6615
         End
      End
      Begin VB.CommandButton cmdOKColors 
         Enabled         =   0   'False
         Height          =   615
         Left            =   -66480
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   7080
         Width           =   615
      End
      Begin VB.CommandButton cmdOKAdvisors 
         Enabled         =   0   'False
         Height          =   615
         Left            =   240
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   4200
         Width           =   615
      End
      Begin VB.CommandButton cmdOKReEmbedding 
         Enabled         =   0   'False
         Height          =   615
         Left            =   -66480
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   7080
         Width           =   615
      End
      Begin VB.CommandButton cmdOKSlidesFromArchive 
         Enabled         =   0   'False
         Height          =   615
         Left            =   -66480
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   7080
         Width           =   615
      End
      Begin VB.CommandButton cmdOKShowOriginalSample 
         Enabled         =   0   'False
         Height          =   615
         Left            =   -66480
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   7080
         Width           =   615
      End
      Begin VB.TextBox txtRequest 
         Height          =   360
         Left            =   -73440
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   1080
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.TextBox txtCreatedBy 
         Height          =   360
         Left            =   -73440
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   1440
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.TextBox txtCreaedOn 
         Height          =   360
         Left            =   -73440
         TabIndex        =   48
         Top             =   1800
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.TextBox txtRequestRemarks 
         Height          =   1965
         Left            =   -73440
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   47
         Top             =   2160
         Visible         =   0   'False
         Width           =   6975
      End
      Begin VB.Frame Frame6 
         Caption         =   "בחירת סליידים"
         Height          =   3735
         Left            =   -74400
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   960
         Width           =   7575
         Begin VB.ListBox lstSlidesFromArchive 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2700
            Left            =   1080
            MultiSelect     =   1  'Simple
            RightToLeft     =   -1  'True
            TabIndex        =   45
            Top             =   360
            Width           =   2895
         End
         Begin VB.CommandButton cmdDeleteSlidesFromArchive 
            Height          =   615
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   44
            Top             =   2920
            Width           =   615
         End
         Begin VB.CommandButton cmdAddToSlidesFromArchive 
            Height          =   615
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   2160
            Width           =   615
         End
         Begin VB.Label lblStainSlides 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   17.25
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   46
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "בחירת יועץ"
         Height          =   2895
         Left            =   5880
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   1080
         Width           =   3855
         Begin VB.ListBox lstAdvisors 
            Height          =   1020
            ItemData        =   "frmAdditionalActions.frx":00AC
            Left            =   240
            List            =   "frmAdditionalActions.frx":00AE
            RightToLeft     =   -1  'True
            TabIndex        =   40
            Top             =   480
            Width           =   3495
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "בחירת ישויות דינאמיות"
         Height          =   2895
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   1080
         Width           =   4215
         Begin VB.ListBox lstSelectedObjects 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1740
            Left            =   1080
            MultiSelect     =   1  'Simple
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   480
            Width           =   2895
         End
         Begin VB.CommandButton cmdDeleteEntity 
            Height          =   615
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   2085
            Width           =   615
         End
         Begin VB.CommandButton cmdSelectEntity 
            Height          =   615
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   1200
            Width           =   615
         End
         Begin VB.Label lblAdvisorCount 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   17.25
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   480
            Width           =   615
         End
      End
      Begin VB.CommandButton cmdDrag 
         BackColor       =   &H8000000A&
         DragMode        =   1  'Automatic
         Height          =   240
         Left            =   -75000
         MousePointer    =   7  'Size N S
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   33
         Tag             =   "cmdDrag"
         Top             =   3360
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Frame frmNewSlides 
         BackColor       =   &H8000000A&
         Caption         =   "סליידים חדשים (סה""כ מ-0 בלוקים)"
         Height          =   2670
         Left            =   -74520
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   720
         Width           =   7815
         Begin VB.PictureBox picBlocks 
            BackColor       =   &H8000000A&
            Height          =   2055
            Left            =   120
            ScaleHeight     =   1995
            ScaleWidth      =   7515
            TabIndex        =   19
            Top             =   480
            Width           =   7575
            Begin VB.VScrollBar VScrollBlocks 
               Height          =   1995
               LargeChange     =   20
               Left            =   7260
               SmallChange     =   4
               TabIndex        =   27
               Top             =   0
               Visible         =   0   'False
               Width           =   255
            End
            Begin VB.PictureBox picBlockEntry 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000A&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   1935
               Left            =   120
               ScaleHeight     =   1935
               ScaleWidth      =   6975
               TabIndex        =   21
               Top             =   120
               Width           =   6975
               Begin VB.CheckBox chkBlockMicrotom 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Index           =   0
                  Left            =   1200
                  RightToLeft     =   -1  'True
                  TabIndex        =   25
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   255
               End
               Begin VB.CommandButton cmdBlockReset 
                  Height          =   315
                  Index           =   0
                  Left            =   120
                  RightToLeft     =   -1  'True
                  Style           =   1  'Graphical
                  TabIndex        =   24
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   495
               End
               Begin VB.ComboBox CmbBlockEntry 
                  Height          =   315
                  Index           =   0
                  Left            =   1920
                  TabIndex        =   23
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   2295
               End
               Begin VB.CommandButton cmdBlock 
                  Height          =   315
                  Index           =   0
                  Left            =   4560
                  Style           =   1  'Graphical
                  TabIndex        =   22
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   780
               End
               Begin VB.Label lblBlock 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000004&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "label6"
                  ForeColor       =   &H80000008&
                  Height          =   315
                  Index           =   0
                  Left            =   5520
                  RightToLeft     =   -1  'True
                  TabIndex        =   26
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   375
               End
            End
            Begin VB.CheckBox chkAddBlockToSample 
               Alignment       =   1  'Right Justify
               Caption         =   "Check1"
               Height          =   375
               Index           =   0
               Left            =   6720
               RightToLeft     =   -1  'True
               TabIndex        =   20
               Top             =   120
               Width           =   255
            End
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "מיקרוטום"
            Height          =   255
            Left            =   1080
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "צנצנת"
            Height          =   255
            Left            =   5760
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "איפוס"
            Height          =   255
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "צביעות"
            Height          =   255
            Left            =   3120
            RightToLeft     =   -1  'True
            TabIndex        =   29
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "בלוק"
            Height          =   255
            Left            =   4920
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.CommandButton cmdExistingLetters 
         Caption         =   "מכתבים שנשלחו"
         Height          =   495
         Left            =   7800
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   4080
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Frame Frame5 
         Caption         =   "התיעצויות קודמות"
         Height          =   2775
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   4920
         Width           =   9495
         Begin VB.PictureBox picAdvisor 
            Height          =   1935
            Left            =   120
            ScaleHeight     =   1875
            ScaleWidth      =   9195
            TabIndex        =   3
            Top             =   600
            Width           =   9255
            Begin VB.VScrollBar VScrollAdvisor 
               Height          =   1875
               LargeChange     =   20
               Left            =   8880
               SmallChange     =   20
               TabIndex        =   12
               Top             =   0
               Visible         =   0   'False
               Width           =   255
            End
            Begin VB.PictureBox picAdvisorEntry 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   120
               ScaleHeight     =   375
               ScaleWidth      =   8775
               TabIndex        =   4
               Top             =   120
               Width           =   8775
               Begin VB.TextBox txtAdvisorRemarks 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Index           =   0
                  Left            =   960
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   8
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   2655
               End
               Begin VB.CommandButton cmdAdvisorLetter 
                  Height          =   315
                  Index           =   0
                  Left            =   405
                  RightToLeft     =   -1  'True
                  Style           =   1  'Graphical
                  TabIndex        =   7
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   420
               End
               Begin VB.CheckBox chkAdvisorReturn 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Index           =   0
                  Left            =   3840
                  RightToLeft     =   -1  'True
                  TabIndex        =   6
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   255
               End
               Begin VB.CommandButton cmdAdvisorEntryOK 
                  Height          =   315
                  Index           =   0
                  Left            =   0
                  RightToLeft     =   -1  'True
                  Style           =   1  'Graphical
                  TabIndex        =   5
                  ToolTipText     =   "עדכון הערות / הגעת תשובה"
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   350
               End
               Begin VB.Label lblRequestId 
                  Alignment       =   1  'Right Justify
                  Caption         =   "lblRequestId"
                  Height          =   315
                  Index           =   0
                  Left            =   4320
                  RightToLeft     =   -1  'True
                  TabIndex        =   11
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   135
               End
               Begin VB.Label lblDate 
                  Alignment       =   1  'Right Justify
                  Caption         =   "lblDate"
                  Height          =   315
                  Index           =   0
                  Left            =   6480
                  TabIndex        =   10
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   2295
               End
               Begin VB.Label lblAdvisor 
                  Alignment       =   1  'Right Justify
                  Caption         =   "lblAdvisor"
                  Height          =   315
                  Index           =   0
                  Left            =   4560
                  RightToLeft     =   -1  'True
                  TabIndex        =   9
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   1815
               End
            End
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "תאריך יצירת הבקשה"
            Height          =   255
            Left            =   7080
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "יועץ"
            Height          =   255
            Left            =   6000
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            Caption         =   "הגיעה תשובה"
            Height          =   495
            Left            =   3960
            RightToLeft     =   -1  'True
            TabIndex        =   14
            Top             =   165
            Width           =   615
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            Caption         =   "הערות"
            Height          =   255
            Left            =   3000
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   360
            Width           =   855
         End
      End
      Begin MSComDlg.CommonDialog cd 
         Left            =   8760
         Top             =   4080
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid grid 
         Height          =   7095
         Left            =   -74880
         TabIndex        =   41
         Top             =   720
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   12515
         _Version        =   393216
         RightToLeft     =   -1  'True
         FocusRect       =   2
         SelectionMode   =   1
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "צנצנת"
         Height          =   255
         Left            =   -68040
         RightToLeft     =   -1  'True
         TabIndex        =   93
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "בלוק"
         Height          =   255
         Left            =   -68760
         RightToLeft     =   -1  'True
         TabIndex        =   92
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "סיבה להעמדה"
         Height          =   255
         Left            =   -70560
         RightToLeft     =   -1  'True
         TabIndex        =   91
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "צנצנת"
         Height          =   255
         Left            =   -68400
         RightToLeft     =   -1  'True
         TabIndex        =   90
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "Request"
         Height          =   255
         Left            =   -74400
         TabIndex        =   89
         Top             =   1120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "Created By"
         Height          =   255
         Left            =   -74400
         TabIndex        =   88
         Top             =   1480
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "Created On"
         Height          =   255
         Left            =   -74400
         TabIndex        =   87
         Top             =   1840
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "Remarks"
         Height          =   255
         Left            =   -74280
         TabIndex        =   86
         Top             =   2200
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "פירוט"
         Height          =   255
         Left            =   -72480
         RightToLeft     =   -1  'True
         TabIndex        =   85
         Top             =   1080
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmAdditionalActions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private isOldSdg As Boolean

Private connection As ADODB.connection
Private rsSdg As Recordset
'Private strOperatorId As String
Private dicMolecularStains As New Dictionary
Private dicSpecialStains As New Dictionary
Private dicImonohistochemistryStains As New Dictionary
Private dicHistochemistryStains As New Dictionary
Private dicOtherStainOptions As New Dictionary
Private WorkFolder As String
Private ProcessXML As LSSERVICEPROVIDERLib.NautilusProcessXML
Private NtlsUser As LSSERVICEPROVIDERLib.NautilusUser
Private NewCellBlockID As String
Private SampleID As String
Private SampleName As String
Private cellBlockRequestDataIdDic As New Dictionary
Private cellBlockRequestIdDic As New Dictionary
 
'to define if the colors OK Button is enabled:
Private iUpdateColorsBlocks As Integer
Private iUpdateColorsSlides As Integer

'to define if the re embedding OK Button is enabled:
Private iUpdateReEmbedding As Integer

'to define if the Add Block OK Button is enabled:
Private iUpdateAddBlock As Integer
 
'to define if the Show Original Sample OK Button is enabled:
Private iUpdateShowSample As Integer

Private rsExtraRequests As Recordset

Private sdg_log As SdgLog.CreateLog

'the destination folder for saving the letters for advisors:
Public strLettersFolder As String
Public ConsultStatus As Boolean
'Private Const LETTERS_FOLDER = "\\limserver\Nautilus Shared\dev result entry\letters\"
Private isCito As Boolean

'___________________________
'PAT - 002
Private isPAP As Boolean
Private IsPapLbc As Boolean
Private testCode As String
'Private papLbcSampleName As String
'_________________________________________________
'pat002
Private Const PapLbcAliquotWF = "PAPS LBC Aliquot"
Private Const PAP_LBC_TEST_CODE_MEDICAL = "81460"
Private Const PAP_LBC_TEST_CODE = "81490"
Private Const PAP_SMEAR_HEADER = "PAP Smear"
Private Const PAP_LBC_HEADER = "PAP LBC"
'__________________________________________________
'___________________________


'holds the templates for creation of slides group
'(the data of the slides to create)
'key - the name of the indicating color (selecting this color will result in the slide list)
'item - a list of slides to create, holds:
'  key - number of slide
'  item - a TemplateSlide object (holds color & layers)
Private dicTemplateSlides As New Dictionary

Private Const MARK_SELECTED = &HC0FFFF

Private selectedSampleName As String


Public Sub Initialize(Con As connection, rs As Recordset, NtlsUser_ As _
    LSSERVICEPROVIDERLib.NautilusUser, sdg_log_ As SdgLog.CreateLog, WorkFolder_ As _
    String, ProcessXML_ As LSSERVICEPROVIDERLib.NautilusProcessXML)
32860 On Error GoTo ERR_Initialize
32870     Call LoadPics
          Dim rsIsCellBlock As Recordset
32880     NewCellBlockID = ""
32890     Call cellBlockRequestDataIdDic.RemoveAll
32900     Call cellBlockRequestIdDic.RemoveAll
32910     Set connection = Con
32920     Set rsSdg = rs
32930     Set NtlsUser = NtlsUser_
          'strOperatorId = strOperatorId_
32940     Set sdg_log = sdg_log_
32950     iUpdateColorsBlocks = 0
32960     iUpdateColorsSlides = 0
32970     iUpdateReEmbedding = 0
32980     iUpdateAddBlock = 1
32990     iUpdateShowSample = 0
33000     If UCase(Left(rsSdg("name"), 1)) = "C" Then
      '       cmdOKAddBlock.Enabled = True
33010        cmdOKAddBlock.Visible = True
33020        isCito = True
33030        selectedSampleName = "-1"
33040     Else
33050         isCito = False
33060         cmdOKAddBlock.Enabled = False
33070         cmdOKAddBlock.Visible = False
33080     End If
      '____________________
      'PAT - 002
33090     isPAP = False
33100     IsPapLbc = False
33110     If UCase(Left(rsSdg("name"), 1)) = "P" Then
33120         isPAP = True
33130         Call InitIsPapLbc(rsSdg("sdg_id"))
33140     End If
      '_____________________
33150     Call InitPhrases
33160     Call InitImgList
33170     Call InitColors
33180     Call DisplayTree
33190     Call PresentBlockList
33200     Call InitAdvisorList
33210     Call InitExtraRequestsHistory
33220     Set ProcessXML = ProcessXML_
33230     WorkFolder = WorkFolder_
33240     SSTab1.Tab = 0
33250     Call InitAdvisorsRequestsList(nte(rsSdg("external_reference")))
33260     Call InitTemplateStains
          
33270     SSTab1.TabEnabled(2) = False
33280 ConsultStatus = False
            Dim i As Integer
            
33290       i = 1
33300       While i < frmAdditionalActions.chkAdvisorReturn.Count
33310           ConsultStatus = ConsultStatus Or frmAdditionalActions.chkAdvisorReturn(i).Enabled
33320           i = i + 1
33330       Wend
      '    Call DisableFieldsForAuthorizedRequest
      '    Call InitAccessToOldLetters(rsSdg("name"))
33340     Exit Sub
ERR_Initialize:
33350 MsgBox "ERR_Initialize in frmAdditionalActions " & vbCrLf & _
    Err.Description
End Sub
'_____________________________________
'PAT - 002
Private Sub InitIsPapLbc(sdgId As String)

33360     On Error GoTo ERR_InitIsPapLbc
          Dim rsSample As Recordset
          Dim sql As String
          
33370     IsPapLbc = False
33380     sql = " select  su.u_test_code, s.name sample_name  "
33390     sql = sql & " from"
33400     sql = sql & "  lims_sys.sample s , lims_sys.sample_user su"
33410     sql = sql & " where "
33420     sql = sql & "         s.sdg_id =" & sdgId
33430     sql = sql & "     and su.sample_id=s.sample_id"
33440     sql = sql & "     and s.status<>'X' "
33450     Set rsSample = connection.Execute(sql)
33460     If Not rsSample.EOF Then
          'assuming there is only one sample EVER!!
33470         testCode = nte(rsSample("u_test_code"))
      '        papLbcSampleName = "-1"
33480         If testCode = PAP_LBC_TEST_CODE Or testCode = _
    PAP_LBC_TEST_CODE_MEDICAL Then
33490             IsPapLbc = True
      '            papLbcSampleName = nte(rsSample("sample_name"))
33500         End If
33510     End If
33520     Exit Sub
ERR_InitIsPapLbc:
33530 MsgBox "ERR_InitIsPapLbc" & vbCrLf & Err.Description
End Sub
'____________________________________
 

Private Sub InitImgList()
33540 On Error GoTo ERR_InitImgList
      '    Call ImgLst.ListImages.Add(0, "sdg", LoadPicture("Resource\sdg.ico"))
33550     Call ImgLst.ListImages.Clear

33560     Call ImgLst.ListImages.Add(, "sdgA", LoadPicture("Resource\sdga.ico"))
33570     Call ImgLst.ListImages.Add(, "sdgC", LoadPicture("Resource\sdgc.ico"))
33580     Call ImgLst.ListImages.Add(, "sdgI", LoadPicture("Resource\sdgi.ico"))
33590     Call ImgLst.ListImages.Add(, "sdgP", LoadPicture("Resource\sdgp.ico"))
33600     Call ImgLst.ListImages.Add(, "sdgR", LoadPicture("Resource\sdgr.ico"))
33610     Call ImgLst.ListImages.Add(, "sdgS", LoadPicture("Resource\sdgs.ico"))
33620     Call ImgLst.ListImages.Add(, "sdgU", LoadPicture("Resource\sdgu.ico"))
33630     Call ImgLst.ListImages.Add(, "sdgV", LoadPicture("Resource\sdgv.ico"))
33640     Call ImgLst.ListImages.Add(, "sdgW", LoadPicture("Resource\sdgw.ico"))
33650     Call ImgLst.ListImages.Add(, "sdgX", LoadPicture("Resource\sdgx.ico"))
33660     Call ImgLst.ListImages.Add(, "sampleA", _
    LoadPicture("Resource\samplea.ico"))
33670     Call ImgLst.ListImages.Add(, "sampleC", _
    LoadPicture("Resource\samplec.ico"))
33680     Call ImgLst.ListImages.Add(, "sampleI", _
    LoadPicture("Resource\samplei.ico"))
33690     Call ImgLst.ListImages.Add(, "sampleP", _
    LoadPicture("Resource\samplep.ico"))
33700     Call ImgLst.ListImages.Add(, "sampleR", _
    LoadPicture("Resource\sampler.ico"))
33710     Call ImgLst.ListImages.Add(, "sampleS", _
    LoadPicture("Resource\samples.ico"))
33720     Call ImgLst.ListImages.Add(, "sampleU", _
    LoadPicture("Resource\sampleu.ico"))
33730     Call ImgLst.ListImages.Add(, "sampleV", _
    LoadPicture("Resource\samplev.ico"))
33740     Call ImgLst.ListImages.Add(, "sampleW", _
    LoadPicture("Resource\samplew.ico"))
33750     Call ImgLst.ListImages.Add(, "sampleX", _
    LoadPicture("Resource\samplex.ico"))
33760     Call ImgLst.ListImages.Add(, "aliquotA", _
    LoadPicture("Resource\aliquota.ico"))
33770     Call ImgLst.ListImages.Add(, "aliquotC", _
    LoadPicture("Resource\aliquotc.ico"))
33780     Call ImgLst.ListImages.Add(, "aliquotI", _
    LoadPicture("Resource\aliquoti.ico"))
33790     Call ImgLst.ListImages.Add(, "aliquotP", _
    LoadPicture("Resource\aliquotp.ico"))
33800     Call ImgLst.ListImages.Add(, "aliquotR", _
    LoadPicture("Resource\aliquotr.ico"))
33810     Call ImgLst.ListImages.Add(, "aliquotS", _
    LoadPicture("Resource\aliquots.ico"))
33820     Call ImgLst.ListImages.Add(, "aliquotU", _
    LoadPicture("Resource\aliquotu.ico"))
33830     Call ImgLst.ListImages.Add(, "aliquotV", _
    LoadPicture("Resource\aliquotv.ico"))
33840     Call ImgLst.ListImages.Add(, "aliquotW", _
    LoadPicture("Resource\aliquotw.ico"))
33850     Call ImgLst.ListImages.Add(, "aliquotX", _
    LoadPicture("Resource\aliquotx.ico"))
          'Call ImgLst.ListImages.Add(4, "Diseases", LoadPicture("Resource\UnitSuitePassed.ico"))
33860     Exit Sub
ERR_InitImgList:
33870 MsgBox "ERR_InitImgList" & vbCrLf & Err.Description
End Sub

Private Sub DisplayTree()
33880 On Error GoTo ERR_DisplayTree
          Dim n As Node
33890     Call tree.Nodes.Clear
33900     Set n = tree.Nodes.Add(, , "sdg:" & rsSdg("sdg_id"), rsSdg("name") & _
    " ", "sdg" + nte(rsSdg("status")))
33910     n.Expanded = True
      '    n.Image = LoadPicture("Resource\sdg" & rsSdg("STATUS") & ".ico")

33920     Call DisplaySamples(rsSdg)
      '    Set rsSamples = connection.Execute '                    ( '                    " select name, sample_id from lims_sys.sample " & '                    " where sdg_id = " & rsSdg("sdg_id") '                    )
      '
      '    While Not rsSamples.EOF
      '        Set n = tree.Nodes.Add("sdg:" & rsSdg("sdg_id"), tvwChild, '                "sample:" & rsSamples("sample_id"), rsSamples("name"), "sample")
      '        n.Expanded = True
      '
      '        rsSamples.MoveNext
      '    Wend

33930 Exit Sub
ERR_DisplayTree:
33940 MsgBox "ERR_DisplayTree" & vbCrLf & Err.Description
End Sub


Private Sub DisplaySamples(rsSdg As Recordset)
          Dim n As Node
          Dim rsSamples As Recordset
          Dim s As String
          
33950     Set rsSamples = _
    connection.Execute(" select name, s.sample_id, s.status, " & _
    " su.u_organ, su.U_TOPOGRAPHY, su.u_material " & _
    " from lims_sys.sample s, lims_sys.sample_user su " & " where sdg_id = " & _
    rsSdg("sdg_id") & " and s.sample_id = su.sample_id " & _
    " and s.status not in ('X') " & " order by su.u_order, s.sample_id ")
          
33960     While Not rsSamples.EOF
33970         s = rsSamples("name") & " (" & rsSamples("u_organ") & ") (" & _
    rsSamples("U_TOPOGRAPHY") & ")"
          
33980         If nte(rsSamples("u_material")) = "T" Then
33990             s = s & " +"
34000             Call AddToSampleList(Mid(rsSamples("name"), 12))

34010         Else
34020             s = s & " -"
34030         End If
      '_______________________________________
      'Pat - 002
      'MsgBox "DisplaySamples- IsPapLbck" & IsPapLbc
34040         If IsPapLbc Then
34050             Call AddToBlocksList(rsSamples("name"))
34060         End If
      '________________________________________
34070         Set n = tree.Nodes.Add("sdg:" & rsSdg("sdg_id"), tvwChild, _
    "sample:" & rsSamples("sample_id"), s, "sample" & rsSamples("status"))
34080         n.Expanded = True

34090         Call DisplayBlocks(rsSamples)
              
              'Call AddToSampleList(Mid(rsSamples("name"), 12))
              
34100         rsSamples.MoveNext
34110     Wend
          
End Sub

Private Sub DisplayBlocks(rsSamples As Recordset)
          Dim n As Node
          Dim rsBlocks As Recordset
          
34120     Set rsBlocks = _
    connection.Execute(" select a.name, a.aliquot_id, a.status, au.u_num_of_tissues,au.u_is_cell_block, au.u_color_type  " _
    & " from lims_sys.aliquot a, lims_sys.aliquot_user au " & _
    " where a.sample_id = " & rsSamples("sample_id") & _
    " and a.aliquot_id = au.aliquot_id " & " and a.aliquot_id not in " & " ( " & _
    " select child_aliquot_id from lims_sys.aliquot_formulation " & _
    " where child_aliquot_id = a.aliquot_id " & " ) " & _
    " and a.status not in ('X') " & " order by aliquot_id ")

34130     While Not rsBlocks.EOF
34140         Set n = tree.Nodes.Add("sample:" & rsSamples("sample_id"), _
    tvwChild, "block:" & rsBlocks("aliquot_id"), rsBlocks("name") & " (" & _
    nte(rsBlocks("u_num_of_tissues")) & nte(rsBlocks("u_color_type")) & ")", _
    "aliquot" & rsBlocks("status"))
34150         n.Expanded = True
              
34160         Call DisplaySlides(rsBlocks)
                                'PAT - 002 (not ispap)
34170         If Not isCito And Not isPAP Then
34180             Call AddToBlocksList(Mid(rsBlocks("name"), 12))
34190             cmdOKAddBlock.Visible = False
34200         ElseIf nte(rsBlocks("u_is_cell_block")) = "T" Then
34210             Call AddToBlocksList(Mid(rsBlocks("name"), 12))
34220             cmdOKAddBlock.Enabled = False
34230             cmdOKAddBlock.Visible = True
34240             cmdOKAddBlock.Caption = " CELL - BLOCK קיים"
34250         End If
      '        'allow creation of new slides / new ReEmbedding results
      '        'only id status is not authorized:
      '        If rsSdg("status") = "A" Then
      '            picBlocks.ToolTipText = "The request is already authorized"
      '            picBlockReembedding.ToolTipText = "The request is already authorized"
      '        Else
      '            Call AddToBlocksList(Mid(rsBlocks("name"), 12))
      '        End If
                      
34260         rsBlocks.MoveNext
34270     Wend

End Sub

Private Sub DisplaySlides(rsBlocks As Recordset)
          Dim n As Node
          Dim rsSlides As Recordset

34280     Set rsSlides = _
    connection.Execute(" select a.name, a.aliquot_id, a.status, au.u_color_type " & _
    " from lims_sys.aliquot a, lims_sys.aliquot_user au " & _
    " where a.aliquot_id = au.aliquot_id " & " and a.aliquot_id in " & " ( " & _
    "   select child_aliquot_id from lims_sys.aliquot_formulation " & _
    "   where child_aliquot_id = a.aliquot_id " & "   and parent_aliquot_id = " & _
    rsBlocks("aliquot_id") & " ) " & " and a.status not in ('X') " & _
    " order by aliquot_id ")

34290     While Not rsSlides.EOF
34300         Set n = tree.Nodes.Add("block:" & rsBlocks("aliquot_id"), _
    tvwChild, "slide:" & rsSlides("aliquot_id"), rsSlides("name") & " (" & _
    nte(rsSlides("u_color_type")) & ")", "aliquot" & rsSlides("status"))
34310         n.Expanded = True
              
              'call here to display slides....
34320         If nte(rsSlides("u_color_type")) = "רזרבה" Or _
    nte(rsSlides("u_color_type")) = "Reserve" Then
34330               Call AddToSlideList(Mid(rsSlides("name"), 12))
34340         End If
              
34350         rsSlides.MoveNext
34360     Wend
End Sub


Private Function nte(e As Variant) As Variant
34370     nte = IIf(IsNull(e), "", e)
End Function


'Private Sub chkAddBlockToSample_Click(Index As Integer)
'
'    If chkAddBlockToSample(Index).Value = 1 Then
'        iUpdateAddBlock = iUpdateAddBlock + 1
'        cmdOKAddBlock.Enabled = True
'    Else
'        iUpdateAddBlock = iUpdateAddBlock - 1
'        If iUpdateAddBlock = 0 Then
'            cmdOKAddBlock.Enabled = False
'        End If
'    End If
'End Sub
Private Sub chkBlockMicrotom_Click(index As Integer)
34380 On Error GoTo ERR_chkBlockMicrotom_Click

34390     BlockAddSlideList.d(BlockAddSlideList.iGuiGroup + index - _
    1).iMicrotom = chkBlockMicrotom(index)
          

34400     Exit Sub
ERR_chkBlockMicrotom_Click:
34410 MsgBox "ERR_chkBlockMicrotom_Click" & vbCrLf & Err.Description
End Sub



Private Sub chkBlockReembeding_Click(index As Integer)
34420 On Error GoTo ERR_chkBlockReembeding_Click
          
34430     If chkBlockReembeding(index).value = 1 Then
34440         iUpdateReEmbedding = iUpdateReEmbedding + 1
34450         cmdOKReEmbedding.Enabled = True
34460     Else
34470         iUpdateReEmbedding = iUpdateReEmbedding - 1
34480         If iUpdateReEmbedding = 0 Then
34490             cmdOKReEmbedding.Enabled = False
34500         End If
34510     End If
          
34520     BlockReembeddingList.d(BlockReembeddingList.iGuiGroup + index - _
    1).iReembedding = chkBlockReembeding(index).value
                                                                  
34530     Exit Sub
ERR_chkBlockReembeding_Click:
34540 MsgBox "ERR_chkBlockReembeding_Click" & vbCrLf & Err.Description
End Sub

'Private Sub chkBlockSerial_Click(Index As Integer)
'    If chkBlockSerial(Index).Value = 1 Then
'        CmbBlockEntry(Index).Enabled = False
'        If CmbBlockEntry(Index).ListCount = 0 Then
'            Call IncrementColorsBlocks
'        End If
'    Else
'        CmbBlockEntry(Index).Enabled = True
'        If CmbBlockEntry(Index).ListCount = 0 Then
'            Call DecrementColorsBlocks
'        End If
'    End If
'End Sub


Private Sub chkShowSample_Click(index As Integer)
34550     If chkShowSample(index).value = 1 Then
34560         iUpdateShowSample = iUpdateShowSample + 1
34570         cmdOKShowOriginalSample.Enabled = True
34580     Else
34590         iUpdateShowSample = iUpdateShowSample - 1
34600         If iUpdateShowSample = 0 Then
34610             cmdOKShowOriginalSample.Enabled = False
34620         End If
34630     End If
End Sub





Private Sub cmbReembeddingReason_Click(index As Integer)
34640 On Error GoTo ERR_cmbReembeddingReason_Click

34650     BlockReembeddingList.d(BlockReembeddingList.iGuiGroup + index - _
    1).strReason = cmbReembeddingReason(index).Text
34660     Exit Sub
ERR_cmbReembeddingReason_Click:
34670 MsgBox "ERR_cmbReembeddingReason_Click" & vbCrLf & Err.Description
End Sub

Private Sub cmbReembeddingReason_GotFocus(index As Integer)
34680     Call zLang.Hebrew
End Sub

Private Sub cmdAddToSlidesFromArchive_Click()
34690 On Error GoTo ERR_cmdAddToSlidesFromArchive_Click

          Dim sText As String
          Dim sKey As String
          Dim sKeyId As String
          Dim skeyName As String
          Dim i As Integer
          Dim dicSlides As New Dictionary
          
          'check if no item was selected:
34700     sText = tree.SelectedItem.Text
34710     sText = Mid(sText, 1, InStr(1, sText, " ") - 1)
34720     If sText = "" Then Exit Sub
          
          'check if item already in the list:
34730     If ExistInList(sText, lstSlidesFromArchive) > -1 Then Exit Sub
          
34740     sKey = tree.SelectedItem.Key
34750     i = InStr(1, sKey, ":")
34760     skeyName = Mid(sKey, 1, i - 1)
34770     sKeyId = Mid(sKey, i + 1)
          
34780     Call dicSlides.RemoveAll
          
34790     Select Case skeyName
              Case "slide"
34800             Call dicSlides.Add(sText, "")
      '            Call lstSlidesFromArchive.AddItem(sText)
      '            lblStainSlides = lblStainSlides + 1
34810         Case "block"
34820             Call GetSlidesForBlock(sKeyId, dicSlides)
34830         Case "sample"
34840             Call GetSlidesForSample(sKeyId, dicSlides)
34850         Case "sdg"
34860             Call GetSlidesForSDG(sKeyId, dicSlides)
34870     End Select
          
34880     For i = 0 To dicSlides.Count - 1
34890         If ExistInList(CStr(dicSlides.Keys(i)), lstSlidesFromArchive) = _
    -1 Then
34900             Call lstSlidesFromArchive.AddItem(CStr(dicSlides.Keys(i)))
34910             lblStainSlides = lblStainSlides + 1
34920         End If
34930     Next i
          
34940     If dicSlides.Count = 0 Then
34950         Call lstSlidesFromArchive.AddItem(sText)
34960     End If
          
          'check if the item is a slide:
      '    If InStr(1, sKey, "slide") <> 0 Then
      '        Call lstSlidesFromArchive.AddItem(sText)
      '    End If
          
           
34970     If lstSlidesFromArchive.ListCount > 0 Then
34980         cmdOKSlidesFromArchive.Enabled = True
34990     End If
          
35000     Exit Sub
ERR_cmdAddToSlidesFromArchive_Click:
35010 MsgBox "ERR_cmdAddToSlidesFromArchive_Click" & vbCrLf & Err.Description
End Sub


Private Sub cmdAdvisorEntryOK_Click(index As Integer)
35020 On Error GoTo ERR_cmdAdvisorEntryOK_Click

          Dim sql As String
          
          'update the remark for the advisor request tab:
35030     sql = " update lims_sys.u_extra_request r "
35040     sql = sql & " set r.DESCRIPTION = '" & txtAdvisorRemarks(index).Text _
    & "'"
35050     sql = sql & " where r.U_EXTRA_REQUEST_ID = '" & _
    lblRequestId(index).Caption & "'"
35060     Call connection.Execute(sql)

          'check if need to update the return of the material
          'from the advisor: update both status and execution date
35070     If chkAdvisorReturn(index).value = 1 And _
    chkAdvisorReturn(index).Enabled = True Then
35080         sql = " update lims_sys.u_extra_request_data_user rdu "
35090         sql = sql & " set rdu.U_STATUS = 'L', "
35100         sql = sql & "     rdu.U_LAB_ON = to_char(sysdate) "
35110         sql = sql & " where rdu.U_EXTRA_REQUEST_ID = '" & _
    lblRequestId(index).Caption & "'"
35120         Call connection.Execute(sql)
              
35130         chkAdvisorReturn(index).Enabled = False
          
       
                     
35140     End If

  'check all advisors return. if all ".enabled=false" then chage sdg consult to false.
35150     If chkAdvisorReturn.Count > 1 Then
              Dim i As Integer
             
           ' Dim sql As String
35160       ConsultStatus = False
            
35170       i = 1
35180       While i < frmAdditionalActions.chkAdvisorReturn.Count
35190           ConsultStatus = ConsultStatus Or frmAdditionalActions.chkAdvisorReturn(i).Enabled
35200           i = i + 1
35210       Wend
35220        sql = " update lims_sys.sdg_user "
35230        sql = sql & " set sdg_user.U_ISCONSULT ='" & IIf(ConsultStatus, "T", "F") & "' "
35240        sql = sql & "  where sdg_user.sdg_id = " & rsSdg("sdg_id")
35250        connection.Execute (sql)
'1550         chkCon.value = IIf(consultStatus, 1, 0)
'MsgBox 4
35260     End If
          
35270     Call InitExtraRequestsHistory

35280     Exit Sub
ERR_cmdAdvisorEntryOK_Click:
35290 MsgBox "ERR_cmdAdvisorEntryOK_Click" & vbCrLf & Err.Description
End Sub

'open Ms Word to show the letter sent to the advisor
Private Sub cmdAdvisorLetter_Click(index As Integer)
35300 On Error GoTo ERR_cmdAdvisorLetter_Click

          Dim wapp As Word.Application
          Dim wdoc As Word.Document
          Dim strFileName As String
          
35310     strFileName = strLettersFolder & Replace(rsSdg("NAME"), "/", "_") & _
    "_" & lblRequestId(index) & ".doc"
                        
35320     Set wapp = New Word.Application
35330     Set wdoc = wapp.Documents.Open(strFileName)
35340     wdoc.Windows(1).Visible = True

35350     Exit Sub
ERR_cmdAdvisorLetter_Click:
35360 MsgBox "ERR_cmdAdvisorLetter_Click" & vbCrLf & Err.Description
End Sub

'adding slides to this block - temp
Private Sub cmdBlock_Click(index As Integer)
35370 On Error GoTo ERR_cmdBlock_Click
          
35380     Call frmColors.Initialize(index, 0, cmdBlock(index).Caption, _
    dicMolecularStains, dicSpecialStains, dicImonohistochemistryStains, _
    dicHistochemistryStains, dicOtherStainOptions)
35390     frmColors.Show vbModal
35400     Call RefreshBlockColorList(index)
          
35410     Exit Sub
ERR_cmdBlock_Click:
35420 MsgBox "ERR_cmdBlock_Click" & vbCrLf & Err.Description
End Sub


Private Sub RefreshBlockColorList(iBlock As Integer)
35430 On Error GoTo ERR_RefreshBlockList

          Dim d As Dictionary
          Dim i As Integer
          
35440     Set d = BlockAddSlideList.d(BlockAddSlideList.iGuiGroup + iBlock - _
    1).dicColors

35450     Call d.RemoveAll
          
35460     For i = 0 To CmbBlockEntry(iBlock).ListCount - 1
35470         Call d.Add(d.Count, CmbBlockEntry(iBlock).list(i))
35480     Next i

      '    Set BlockAddSlideList.d(BlockAddSlideList.iGuiGroup + iBlock - 1).dicColors = d
          
35490 Exit Sub
ERR_RefreshBlockList:
35500 MsgBox "ERR_RefreshBlockList" & vbCrLf & Err.Description
End Sub


Private Sub AddToSampleList(strName As String)
35510 On Error GoTo ERR_AddToSampleList
          Dim i As Integer

35520     i = chkShowSample.Count

          '---------------------------
          'show samples to doctor TAB:
          '---------------------------
35530     Load chkShowSample(i)
35540     With chkShowSample(i)
35550         .Top = 500 * (i - 1) + 10
35560         .Visible = True
35570         .value = 0
35580         .Caption = strName
35590     End With

35600     picShowSamplesEntry.Height = chkShowSample(i).Top + _
    chkShowSample(i).Height + 200
      '    picShowSamplesEntry.Height = picShowSamplesEntry.Height + 1.6 * chkShowSample(i).Height
          
35610     VScrollShowSamples.Max = (picShowSamplesEntry.Height - _
    picShowSamples.Height) / ScaleHeight * 100
35620     If VScrollShowSamples.Max < 0 Then
35630         VScrollShowSamples.Visible = False
35640     Else
35650         VScrollShowSamples.Visible = True
35660     End If


          '---------------------------
          'add block to sample TAB:
      '    ---------------------------
      'If UCase(Left(rsSdg("name"), 1)) = "C" Then
      '       SSTab1.TabEnabled(2) = True
      '    End If


      '    Load chkAddBlockToSample(i)
      '    With chkAddBlockToSample(i)
      '        .Top = 500 * (i - 1) + 10
      '        .Visible = True
      '        .Value = 0
      '        .Caption = strName
      '    End With
      '
      '    picAddBlockEntry.Height = picAddBlockEntry.Height + 1.6 * chkAddBlockToSample(i).Height
      '
      '    VScrollAddBlocks.Max = (picAddBlockEntry.Height - picAddBlocks.Height) / ScaleHeight * 100
      '    If VScrollAddBlocks.Max < 0 Then
      '        VScrollAddBlocks.Visible = False
      '    Else
      '        VScrollAddBlocks.Visible = True
      '    End If


35670     Exit Sub
ERR_AddToSampleList:
35680 MsgBox "ERR_AddToSampleList" & vbCrLf & Err.Description
End Sub


Private Sub AddToSlideList(strName As String)
35690 On Error GoTo ERR_AddToSlideList
          Dim i As Integer
          Dim s As String
           
35700     i = cmdSlide.Count
          
35710     Load cmdSlide(i)
35720     With cmdSlide(i)
35730         .Caption = strName
35740         .Visible = True
35750         .Top = 500 * (i - 1) + 10
35760     End With

35770     Load txtSlide(i)
35780     With txtSlide(i)
35790         .Visible = True
35800         .Top = 500 * (i - 1) + 10
35810     End With

35820     Load cmdSlideReset(i)
35830     With cmdSlideReset(i)
35840         .Visible = True
35850         .Top = 500 * (i - 1) + 10
35860     End With


35870     s = frmExistingSlides.Caption
35880     s = Replace(s, CStr(i - 1), CStr(i))
35890     frmExistingSlides.Caption = s

35900     picSlidesEntry.Height = cmdSlide(i).Top + cmdSlide(i).Height + 200
          'picSlidesEntry.Height = picSlidesEntry.Height + 1.6 * cmdSlide(i).Height
          
35910     VScrollSlides.Max = (picSlidesEntry.Height - picSlides.Height) / _
    ScaleHeight * 100
35920     If VScrollSlides.Max < 0 Then
35930         VScrollSlides.Visible = False
35940     Else
35950         VScrollSlides.Visible = True
35960     End If

35970     Exit Sub
ERR_AddToSlideList:
35980 MsgBox "ERR_AddToSlideList" & vbCrLf & Err.Description
End Sub


Private Sub AddToBlocksList(strName As String)
35990 On Error GoTo ERR_AddToBlocksList

          Dim i As Integer
          Dim k As Integer
          Dim s As String
          Dim bas As BlockAddSlide
          Dim bre As BlockReembedding
          'i = cmdBlock.Count
36000     i = chkBlockReembeding.Count

          
          '-------------------------------
          'initialize the colors tab:
          '-------------------------------
          
36010     k = InStr(1, strName, ".", vbTextCompare)
36020     If Mid(strName, k + 1) = "1" Then
36030         s = Mid(strName, 1, k - 1)
36040     End If

          
36050     Set bas = New BlockAddSlide
36060     bas.strSample = s
36070     bas.strName = strName
36080     bas.iMicrotom = 0
      '______________________________
      'PAT - 002
36090     If IsPapLbc Then
36100         bas.strSample = Mid(strName, k + 1)
36110         bas.strName = "N/A"
36120     End If
      '______________________________

36130     Call BlockAddSlideList.d.Add(BlockAddSlideList.d.Count + 1, bas)

          '-------------------------------
          'initialize the reembedding tab:
          '-------------------------------

36140     Set bre = New BlockReembedding
36150     bre.strSample = s
36160     bre.strName = strName
36170     bre.iReembedding = 0
36180     bre.strReason = ""
36190     bre.strDetails = ""
36200     Call BlockReembeddingList.d.Add(BlockReembeddingList.d.Count + 1, bre)
          
36210     Exit Sub
ERR_AddToBlocksList:
36220 MsgBox "ERR_AddToBlocksList" & vbCrLf & Err.Description
End Sub


Private Sub PresentBlockList()
36230 On Error GoTo ERR_PresentBlockList

          
          Dim i As Integer
          Dim k As Integer
          
          
          'show the list of blocks to add slides from:
36240     BlockAddSlideList.iGuiGroup = 1
          
36250     For i = 1 To BlockAddSlideList.MAX_LINES
36260         If BlockAddSlideList.d.Exists(i) Then
                  
36270             Load lblBlock(i)
36280             With lblBlock(i)
36290                 .Top = 500 * (i - 1)
36300                 .Caption = BlockAddSlideList.d(i).strSample
36310                 .Visible = .Caption <> ""
36320             End With
                            
36330             Load cmdBlock(i)
36340             With cmdBlock(i)
36350                 .Top = 500 * (i - 1)
36360                 .Caption = BlockAddSlideList.d(i).strName
36370                 .Visible = True
36380             End With
                  
36390             Load CmbBlockEntry(i)
36400             With CmbBlockEntry(i)
36410                 .Top = 500 * (i - 1)
36420                 .Visible = True
36430             End With
                  
36440             Load chkBlockMicrotom(i)
36450             With chkBlockMicrotom(i)
36460                 .Top = 500 * (i - 1)
36470                 .Visible = True
36480                 .value = 0
36490             End With
              
36500             Load cmdBlockReset(i)
36510             With cmdBlockReset(i)
36520                 .Top = 500 * (i - 1)
36530                 .Visible = True
36540             End With
                  
36550         End If
36560     Next i
          
36570     VScrollBlocks.Max = (510 * (BlockAddSlideList.d.Count) - _
    picBlocks.Height) / 12570 * 100
      '    VScrollBlocks.Max = (510 * (BlockAddSlideList.d.Count) - picBlocks.Height) / ScaleHeight * 100
36580     If VScrollBlocks.Max < 0 Or BlockAddSlideList.d.Count <= _
    BlockAddSlideList.MAX_LINES Then
36590         VScrollBlocks.Visible = False
36600     Else
36610         VScrollBlocks.Visible = True
36620     End If
          'BlockAddSlideList.iScrollvalue = VScrollBlocks.Value



          

          'show the list of blocks for re embedding:
          
36630     BlockReembeddingList.iGuiGroup = 1
          
36640     For i = 1 To BlockReembeddingList.MAX_LINES
36650         If BlockReembeddingList.d.Exists(i) Then
              
36660             Load chkBlockReembeding(i)
36670             With chkBlockReembeding(i)
36680                 .Top = 500 * (i - 1)
36690                 .Caption = BlockReembeddingList.d(i).strName
36700                 .Visible = True
36710             End With
              
36720             Load lblSampleReembedding(i)
36730             With lblSampleReembedding(i)
36740                 .Top = 500 * (i - 1)
36750                 .Caption = BlockReembeddingList.d(i).strSample
36760                 .Visible = .Caption <> ""
36770             End With
              
36780             Load cmbReembeddingReason(i)
36790             With cmbReembeddingReason(i)
36800                 .Top = 500 * (i - 1)
36810                 .Visible = True
36820                 For k = 0 To cmbReembeddingReason(0).ListCount - 1
36830                     cmbReembeddingReason(i).list(k) = _
    cmbReembeddingReason(0).list(k)
36840                 Next k
36850             End With
              
36860             Load txtReembeddingDetails(i)
36870             With txtReembeddingDetails(i)
36880                 .Top = 500 * (i - 1)
36890                 .Visible = True
36900             End With
                  
36910         End If
36920     Next i
36930     VScrollBlocksReembedding.Max = (510 * (BlockReembeddingList.d.Count) _
    - picBlockReembedding.Height) / 12570 * 100
      '    VScrollBlocksReembedding.Max = (510 * (BlockReembeddingList.d.Count) - picBlockReembedding.Height) / ScaleHeight * 100
36940     If VScrollBlocksReembedding.Max < 0 Or BlockReembeddingList.d.Count _
    <= BlockReembeddingList.MAX_LINES Then
36950         VScrollBlocksReembedding.Visible = False
36960     Else
36970         VScrollBlocksReembedding.Visible = True
36980     End If
          'BlockReembeddingList.iScrollvalue = VScrollBlocksReembedding.Value


36990     Exit Sub
ERR_PresentBlockList:
37000     MsgBox "ERR_PresentBlockList" & vbCrLf & Err.Description
End Sub



Private Sub AddToAdvisorList(rs As Recordset)
37010 On Error GoTo ERR_AddToAdvisorList

          Dim i As Integer
          Dim s As String
          
37020     i = lblDate.Count
          
37030     Load lblDate(i)
37040     With lblDate(i)
37050         .Top = 500 * (i - 1) + 10
37060         .Caption = nte(rs("u_created_on"))
37070         .Visible = True
37080     End With

37090     Load lblAdvisor(i)
37100     With lblAdvisor(i)
37110         .Top = 500 * (i - 1) + 10
37120         .Caption = nte(rs("advisor"))
37130         .Visible = True
37140     End With

37150     Load chkAdvisorReturn(i)
37160     With chkAdvisorReturn(i)
37170         .Top = 500 * (i - 1) + 10
37180         s = nte(rs("status"))
37190         If s = "V" Or s = "P" Then
37200             .value = 0
37210             .Enabled = True
37220         Else
37230             .value = 1
37240             .Enabled = False
37250         End If
37260         .Visible = True
37270     End With

37280     Load txtAdvisorRemarks(i)
37290     With txtAdvisorRemarks(i)
37300         .Top = 500 * (i - 1) + 10
37310         .Text = nte(rs("description"))
37320         .Visible = True
37330     End With

37340     Load lblRequestId(i)
37350     With lblRequestId(i)
37360         .Top = 500 * (i - 1) + 10
37370         .Caption = nte(rs("u_extra_request_id"))
37380     End With

37390     Load cmdAdvisorLetter(i)
37400     With cmdAdvisorLetter(i)
37410         .Top = 500 * (i - 1) + 10
              '.Text = nte(rs("description"))
37420         .Picture = LoadPicture("Resource\address.ico")
              
37430         If ExistLetterToAdvisor(Replace(rsSdg("name"), "/", "_") & "_" & _
    lblRequestId(i).Caption & ".doc") = True Then
37440             .Visible = True
37450         Else
37460             .Visible = False
37470         End If
37480     End With

37490     Load cmdAdvisorEntryOK(i)
37500     With cmdAdvisorEntryOK(i)
37510         .Top = 500 * (i - 1) + 10
37520         .Visible = True
37530         .Enabled = True
37540         .Picture = LoadPicture("Resource\data type boolean.ico")
37550     End With



37560     picAdvisorEntry.Height = lblDate(i).Top + lblDate(i).Height + 200
      '    picBlockReembeddingEntry.Height = picBlockReembeddingEntry.Height + 1.6 * chkBlockReembeding(i).Height
          
37570     VScrollAdvisor.Max = (picAdvisorEntry.Height - picAdvisor.Height) / _
    ScaleHeight * 100
37580     If VScrollAdvisor.Max < 0 Then
37590         VScrollAdvisor.Visible = False
37600     Else
37610         VScrollAdvisor.Visible = True
37620     End If

37630     Exit Sub
ERR_AddToAdvisorList:
37640 MsgBox "ERR_AddToAdvisorList" & vbCrLf & Err.Description
End Sub


Private Sub cmdBlockReset_Click(index As Integer)
37650 On Error GoTo ERR_cmdBlockReset_Click

37660     If CmbBlockEntry(index).ListCount > 0 Then
37670         CmbBlockEntry(index).Clear
37680         CmbBlockEntry(index).BackColor = vbWhite
37690         BlockAddSlideList.d(BlockAddSlideList.iGuiGroup + index - _
    1).dicColors.RemoveAll
              
37700         Call DecrementColorsBlocks
              'If b = False Then Call DecrementColorsBlocks
37710     End If
          
37720     chkBlockMicrotom(index).value = 0
37730     BlockAddSlideList.d(BlockAddSlideList.iGuiGroup + index - _
    1).iMicrotom = False
              
37740     Exit Sub
ERR_cmdBlockReset_Click:
37750 MsgBox "ERR_cmdBlockReset_Click" & vbCrLf & Err.Description
End Sub


'if no items are selected, all is deleted
'else - only selected are deleted
Private Sub cmdDeleteEntity_Click()
37760 On Error GoTo ERR_cmdDeleteEntity_Click
          
37770     If lstSelectedObjects.SelCount = 0 Then
37780         lstSelectedObjects.Clear
37790         lblAdvisorCount = 0
37800     Else
              Dim i As Integer
              Dim d As New Dictionary
          
              'read items NOT to be deleted:
37810         For i = 0 To lstSelectedObjects.ListCount - 1
37820             If lstSelectedObjects.Selected(i) = False Then
37830                 Call d.Add(lstSelectedObjects.list(i), "")
37840             End If
37850         Next i
              
37860         lstSelectedObjects.Clear
37870         For i = 0 To d.Count - 1
37880             Call lstSelectedObjects.AddItem(d.Keys(i))
37890         Next i
37900         lblAdvisorCount = d.Count
37910     End If
          
37920     If lstSelectedObjects.ListCount = 0 Then
37930         cmdOKAdvisors.Enabled = False
37940     End If
          
      '    Dim i As Integer
      '
      '    For i = 0 To lstSelectedObjects.ListCount - 1
      '        If lstSelectedObjects.Selected(i) Then
      '            MsgBox lstSelectedObjects.List(i)
      '           ' lstSelectedObjects.RemoveItem (i)
      '        End If
      '    Next i
              
37950     Exit Sub
ERR_cmdDeleteEntity_Click:
37960 MsgBox "ERR_cmdDeleteEntity_Click" & vbCrLf & Err.Description
End Sub

'if no items are selected, all is deleted
'else - only selected are deleted
Private Sub cmdDeleteSlidesFromArchive_Click()
37970 On Error GoTo ERR_cmdDeleteSlidesFromArchive_Click

37980     If lstSlidesFromArchive.SelCount = 0 Then
37990         lstSlidesFromArchive.Clear
38000         lblStainSlides = 0
38010     Else
              Dim i As Integer
              Dim d As New Dictionary
          
              'read items NOT to be deleted:
38020         For i = 0 To lstSlidesFromArchive.ListCount - 1
38030             If lstSlidesFromArchive.Selected(i) = False Then
38040                 Call d.Add(lstSlidesFromArchive.list(i), "")
38050             End If
38060         Next i
              
38070         lstSlidesFromArchive.Clear
38080         For i = 0 To d.Count - 1
38090             Call lstSlidesFromArchive.AddItem(d.Keys(i))
38100         Next i
38110         lblStainSlides = d.Count
38120     End If
          
38130     If lstSlidesFromArchive.ListCount = 0 Then
38140         cmdOKSlidesFromArchive.Enabled = False
38150     End If

      '    lstSlidesFromArchive.Clear
      '    lblStainSlides = 0
      '    cmdOKSlidesFromArchive.Enabled = False
38160     Exit Sub
ERR_cmdDeleteSlidesFromArchive_Click:
38170 MsgBox "ERR_cmdDeleteSlidesFromArchive_Click" & vbCrLf & Err.Description
End Sub

'Private Sub cmdFirst_Click()
'On Error GoTo ERR_cmdFirst_Click
'
'    Call rsExtraRequests.MoveFirst
''    Call SetExtraRequestInfo
'    txtRequestNumber.Text = 1
'    cmdPrev.Enabled = False
'    cmdNext.Enabled = True
'
'    Exit Sub
'ERR_cmdFirst_Click:
'MsgBox "ERR_cmdFirst_Click" & vbCrLf & Err.description
'End Sub

'Private Sub cmdLast_Click()
'On Error GoTo ERR_cmdLast_Click
'
'    Call rsExtraRequests.MoveLast
''    Call SetExtraRequestInfo
'    txtRequestNumber.Text = rsExtraRequests.RecordCount
'    cmdNext.Enabled = False
'    cmdPrev.Enabled = True
'
'    Exit Sub
'ERR_cmdLast_Click:
'MsgBox "ERR_cmdLast_Click" & vbCrLf & Err.description
'End Sub

'Private Sub cmdNext_Click()
'On Error GoTo ERR_cmdNext_Click
'
'    If Not rsExtraRequests.EOF Then
'        rsExtraRequests.MoveNext
'
' '       Call SetExtraRequestInfo
'
'        txtRequestNumber.Text = txtRequestNumber.Text + 1
'
'        cmdPrev.Enabled = True
'    End If
'
'    If rsExtraRequests.RecordCount = txtRequestNumber.Text Then
'        cmdNext.Enabled = False
'    End If
'
'    Exit Sub
'ERR_cmdNext_Click:
'MsgBox "ERR_cmdNext_Click" & vbCrLf & Err.description
'End Sub

Private Sub cmdOKAddBlock_Click()
38180 On Error GoTo ERR_cmdOKAddBlock_Click
          Dim i As Integer
          Dim temp As String
          

       '   tree.
38190     If selectedSampleName = "-1" Then
38200         MsgBox "Please select a sample from the tree on the left"
38210         Exit Sub
38220     End If
              
38230     If MsgBox("Are you  sure you  want to create a CELLBLOCK ?", _
    vbYesNoCancel + vbDefaultButton3, "CELLBLOCK IS ABOUT TO BE CREATED!") <> vbYes _
    Then
38240         Exit Sub
38250     End If
      '(else)
          
          
          Dim rsSamples As Recordset
38260     Set rsSamples = _
    connection.Execute("select s.name,s.sample_id  from lims_sys.sample s " & _
    " where s.name = '" & selectedSampleName & "' " & _
    " and status in ('V','S','P','C','I') ")
38270     If rsSamples.EOF Then
38280         MsgBox "cannot find a sample in valid status, block not added"
38290         Exit Sub
38300     End If
          
38310     NewCellBlockID = TriggerSampleEvent("Add Cell Block", _
    rsSamples("sample_id"))
          
38320     temp = EnterNewRequest("Creat Cell Block", rsSdg("name"))
38330     Call cellBlockRequestIdDic.Add(temp, Null)
38340     temp = _
    EnterNewRequestData(CStr(cellBlockRequestIdDic.Keys(cellBlockRequestIdDic.Count _
    - 1)), rsSamples("name"), "Sample", "Creat Cell Block", _
    "Make a Cell Block out of a Citology Sample")
38350     Call cellBlockRequestDataIdDic.Add(temp, Null)

          'reset the Add Block tab
      '    For i = 1 To chkAddBlockToSample.Count - 1
      '        chkAddBlockToSample(i).Value = 0
      '    Next i
          'txtAddBlock.Text = ""
38360     cmdOKAddBlock.Enabled = False
38370     cmdOKAddBlock.Caption = " CELL - BLOCK קיים"
      '    iUpdateAddBlock = iUpdateAddBlock + 1
      '
      '    Call InitExtraRequestsHistory
      '    Call DisplayTree
      '    Call PresentBlockList
         
              Dim rsBlocks As Recordset
              Dim n As Node
              
38380         Set rsBlocks = _
    connection.Execute(" select a.name, a.aliquot_id, a.status, au.u_num_of_tissues,au.u_is_cell_block " _
    & " from lims_sys.aliquot a, lims_sys.aliquot_user au " & _
    " where a.sample_id = " & rsSamples("sample_id") & _
    " and a.aliquot_id = au.aliquot_id " & " and a.aliquot_id = " & NewCellBlockID)
38390         Set n = tree.Nodes.Add("sample:" & rsSamples("sample_id"), _
    tvwChild, "block:" & NewCellBlockID, rsBlocks("name") & " (" & _
    nte(rsBlocks("u_num_of_tissues")) & ")", "aliquot" & rsBlocks("status"))
38400         n.Expanded = True
      '        2) add cell block to dictionary VV
38410         Call AddToBlocksList(Mid(rsBlocks("name"), 12))
      '        3) display cell block in block list
38420         i = lblBlock.Count
38430         Load lblBlock(i)
38440         With lblBlock(i)
38450             .Top = 500 * (i - 1)
38460             .Caption = BlockAddSlideList.d(i).strSample
38470             .Visible = .Caption <> ""
38480         End With
                        
38490         Load cmdBlock(i)
38500         With cmdBlock(i)
38510             .Top = 500 * (i - 1)
38520             .Caption = BlockAddSlideList.d(i).strName
38530             .Visible = True
38540         End With
              
38550         Load CmbBlockEntry(i)
38560         With CmbBlockEntry(i)
38570             .Top = 500 * (i - 1)
38580             .Visible = True
38590         End With
              
38600         Load chkBlockMicrotom(i)
38610         With chkBlockMicrotom(i)
38620             .Top = 500 * (i - 1)
38630             .Visible = True
38640             .value = 0
38650         End With
          
38660         Load cmdBlockReset(i)
38670         With cmdBlockReset(i)
38680             .Top = 500 * (i - 1)
38690             .Visible = True
38700         End With
                  


          
38710     MsgBox "CELLBLOCK CREATED!"


38720     Exit Sub
ERR_cmdOKAddBlock_Click:
38730 MsgBox "ERR_cmdOKAddBlock_Click" & vbCrLf & Err.Description
End Sub
Private Sub tree_KeyDown(KeyCode As Integer, Shift As Integer)
          Dim strVer As String

38740     If KeyCode = vbKeyF10 And Shift = 1 Then
38750         strVer = "One Software Technologies (O.S.T) Ltd."
38760         MsgBox strVer & vbCrLf & "DEBUG MODE!!!!", vbInformation, _
    "Nautilus - Project Properties"
38770         cmdOKAddBlock.Enabled = True
               Dim tmpname As String
          Dim sql As String
          Dim rst As Recordset

38780     tmpname = Trim(getNextStr(tree.SelectedItem, " "))
38790         If Not InStr(tmpname, ".") > 1 Then
38800             selectedSampleName = "-1"
38810         Else
38820             selectedSampleName = getNextStr(tmpname, ".")
38830             selectedSampleName = selectedSampleName & "." & _
    getNextStr(tmpname & ".", ".")
38840         End If
38850     End If
End Sub





'Private Sub cmdExistingLetters_Click()
'On Error GoTo ERR_cmdExistingLetters
'
'    Dim wapp As Word.Application
'    Dim wdoc As Word.Document
'    Dim strFileName As String
'    Dim strPrefix As String
'
'    strPrefix = Replace(rsSdg("NAME"), "/", "_") & "*"
'
'    'show only letters of this request:
'    cd.InitDir = LETTERS_FOLDER
'    cd.Filter = "doc"
'    cd.FileName = strPrefix
'    cd.ShowOpen
'
'    strFileName = cd.FileName
'    If strFileName = strPrefix Then Exit Sub
'
'    Set wapp = New Word.Application
'    Set wdoc = wapp.Documents.Open(strFileName)
'    wdoc.Windows(1).Visible = True
'
'    Exit Sub
'ERR_cmdExistingLetters:
'MsgBox "ERR_cmdExistingLetters" & vbCrLf & Err.Description
'End Sub

Private Sub cmdOKAdvisors_Click()
38860 On Error GoTo ERR_cmdOKAdvisors_Click
          Dim i As Integer
          Dim strEntities As String
          Dim strExtraRequestId As String
'          Dim wapp As Word.Application
'          Dim wdoc As Word.Document
'          Dim wtbl As Word.Table
'          Dim fs As New FileSystemObject
'          Dim strDestinationFile As String
          

38870     strExtraRequestId = EnterNewRequest("Send to Consultant", "")

38880     For i = 0 To lstSelectedObjects.ListCount - 1
38890         Call EnterNewRequestData(strExtraRequestId, _
    lstSelectedObjects.list(i), GetEntityType(lstSelectedObjects.list(i)), _
    lstAdvisors.Text, "")
              
38900         strEntities = strEntities & i + 1 & ". " & _
    lstSelectedObjects.list(i) & vbCr
38910     Next i


38920     Call sdg_log.InsertLog(rsSdg("sdg_id"), "EXTRA.CREATED", "Send " & _
    lstSelectedObjects.ListCount & " item(s) to consultant")

          'produce and print a request letter from the selected information:----------------
'38450     Set wapp = New Word.Application

          'use a pre defined template for this doc:
          'Set wdoc = wapp.Documents.Add("letter2")
'38460     strDestinationFile = strLettersFolder & Replace(rsSdg("NAME"), "/", _
'    "_") & "_" & strExtraRequestId & ".doc"
'38470     Call fs.CopyFile("C:\LetterToAdvisor.doc", strDestinationFile)

'38480     Set wdoc = wapp.Documents.Open(strDestinationFile)

          'replace the field in the template with the real value (sender name):
'38490     Call InsertFieldValue("#send_to#", lstAdvisors.Text, wdoc, wapp)
'38500     Call InsertFieldValue("#date#", Date, wdoc, wapp)
'38510     Call InsertFieldValue("#doctor#", _
'    GetOperatorName(NtlsUser.GetOperatorId) & ", " & _
'    GetOperatorRole(NtlsUser.GetOperatorId), wdoc, wapp)
      '    Call InsertFieldValue("#list#", strEntities, wdoc, wapp)
      '    Call InsertFieldValue("#remarks#", Replace(txtAdvice.Text, vbCrLf, vbCr), wdoc, wapp)

          
          ' save the file in the destination letters folder assigned for that matter
          ' befor opening it for the client to complete
          ' (so he only needs to "SAVE", and not "SAVE AS" where finished)
          'Call wdoc.SaveAs(strDestinationFile)
'38520     Call wdoc.Save

'38530     wdoc.Windows(1).Visible = True
      '    Set wtbl = wdoc.Tables.Add(wdoc.Range(0, 0), 3, 2)

      '    Call wdoc.SaveAs("C:\1.doc")
          
          'if we wonna print this doc:
      '    Call wdoc.PrintOut
          
      '    Call wdoc.Close
          '----------------------------------------------------------------------------------

38930     Call InitAdvisorsRequestsList(nte(rsSdg("external_reference")))
38940     Call InitExtraRequestsHistory

          'display a message:
38950     MsgBox "The request was saved successfully"
          
          'clean the fields of this tab:
38960     lstSelectedObjects.Clear
          'txtAdvice.Text = ""
38970     lblAdvisorCount.Caption = 0
          
38980     cmdOKAdvisors.Enabled = False
          
      '    cmdExistingLetters.Visible = True

38990     Exit Sub
ERR_cmdOKAdvisors_Click:
39000 MsgBox "ERR_cmdOKAdvisors_Click" & vbCrLf & Err.Description
End Sub

Private Sub cmdOKColors_Click()
39010 On Error GoTo ERR_cmdOKColors_Click
          
39020     Me.MousePointer = 11
          
          ' Ordering additional slides of old sdg's causes a problem,
          ' since aliquotes of old sdg's were made using different WF.
          ' In such cases, orders of additional slides\colors will be made manually.
39030     isOldSdg = CDate(rsSdg("created_on")) < CDate("01/01/2006")
          
39040     If iUpdateColorsBlocks Then
39050         If IsPapLbc Then
39060             Call ApproveNewSlidesForSample
39070         Else
39080             Call ApproveNewSlides
                  
39090          End If
39100     End If
39110     If iUpdateColorsSlides Then
39120             Call ApproveExistingSlides
39130     End If
          
39140     If isOldSdg Then
          
39150         MsgBox _
    ".לדגימה זו, בנוסף להזמנה האלקטרונית, יש לצרף הזמנה ידנית ולהעבירה לביצוע במעבדה" _
    & vbCrLf & " (PAT-01-FB-014-A/B/C/D טופס)"
39160     End If
          
39170     Call ResetColorTab
39180     Call InitExtraRequestsHistory
           
          Dim i As Integer
          
          'remove all reseves from before reload
39190     If cmdSlide.Count <> 0 Then
39200         For i = cmdSlide.Count - 1 To 1 Step -1
39210             Unload cmdSlide(i)
39220             Unload txtSlide(i)
39230             Unload cmdSlideReset(i)
39240         Next i
39250     End If
39260     Me.MousePointer = 0
          
          'Call UpdateSecretaryRequest
              
39270     MsgBox "The request was saved successfully"
          
39280     Call DisplayTree

39290     Exit Sub
ERR_cmdOKColors_Click:
39300 MsgBox "ERR_cmdOKColors_Click" & vbCrLf & Err.Description
End Sub

'Private Sub UpdateAliquotStatus( AliquotID As String, chrNewStatus As String)
'On Error GoTo ERR_UpdateAliquotStatus
'
'    Dim sql As String
'
'
'
'        sql = " update lims_sys.Aliquot a"
'        sql = sql & " set a.status='" & UCase(chrNewStatus) & "'"
'        sql = sql & " where a.Aliquot_id='" & AliquotID & "'"
'
'        Call connection.Execute(sql)
'        AliquotID = ""
''        Clipboard.SetText (sql)
''        MsgBox sql
'
'    Exit Sub
'ERR_UpdateAliquotStatus:
'MsgBox "ERR_UpdateAliquotStatus" & vbCrLf & Err.Description
'End Sub
'if a secretary request,
'change the sdg patholog to Secretary:
Private Sub UpdateSecretaryRequest()
39310 On Error GoTo ERR_UpdateSecretaryRequest
          
          Dim sql As String

39320     If rsSdg("status") = "A" And NtlsUser.GetRoleId <> "63" Then
              
39330         sql = " update lims_sys.sdg_user du"
39340         sql = sql & " set du.U_PATHOLOG='113'"
39350         sql = sql & " where du.SDG_ID='" & rsSdg("sdg_id") & "'"
              
39360         Call connection.Execute(sql)
              
39370     End If

39380     Exit Sub
ERR_UpdateSecretaryRequest:
39390 MsgBox "ERR_UpdateSecretaryRequest" & vbCrLf & Err.Description
End Sub


Private Sub ApproveNewSlides()
39400 On Error GoTo ERR_ApproveNewSlides
          Dim i As Integer
          Dim j As Integer
          Dim k As Integer
          Dim iSlidesForColor As Integer
          Dim iCount As Integer
          Dim strExtraRequestId As String
          Dim strExtraRequestDataID As String
          Dim SlideAliquotID As Long
          Dim strBlockName As String
          Dim strBlockId As String
          Dim strColorName As String
          Dim strColorCount As String
          Dim strShowMicrotom As String
          Dim dColors As Dictionary
          
          'a string of all color groups for the slides;
          'written as part of the description for the sdg-log:
          Dim strColorGroups As String
          'insert data to the extra_request table:
39410     If isOldSdg Then
39420         strExtraRequestId = EnterNewRequest("Add Manual Slide", _
    txtColors.Text)
39430     Else
39440         strExtraRequestId = EnterNewRequest("Add Slide", txtColors.Text)
39450     End If
39460     If Not cellBlockRequestIdDic.Exists(strExtraRequestId) Then
39470         Call cellBlockRequestIdDic.Add(strExtraRequestId, Null)
39480     End If
          'go through all the blocks to find which are selected to add slides to:
      '    For i = 1 To cmdBlock.Count - 1

39490     For i = 1 To BlockAddSlideList.d.Count
39500         If BlockAddSlideList.d(i).dicColors.Count > 0 Then
              'If CmbBlockEntry(i).ListCount > 0 Then
39510             strBlockName = rsSdg("name") & "." & _
    BlockAddSlideList.d(i).strName 'cmdBlock(i).Caption
39520             strBlockId = GetAliquotId(strBlockName)
                  
39530             strShowMicrotom = "F"
39540             If BlockAddSlideList.d(i).iMicrotom = 1 Then
                  'If chkBlockMicrotom(i).Value = 1 Then
39550                 strShowMicrotom = "T"
39560             End If
              
                  'go through the items (requested slides) in the list of this block:
39570             Set dColors = BlockAddSlideList.d(i).dicColors
                  'For j = 0 To CmbBlockEntry(i).ListCount - 1
                 
39580             For j = 0 To dColors.Count - 1
39590                 Call ParseBlockColor(CStr(dColors.Items(j)), strColorName, _
    strColorCount)
                      'Call ParseBlockColor(CmbBlockEntry(i).list(j), strColorName, strColorCount)
39600                 iSlidesForColor = CInt(strColorCount)
                      
                      'loop for the number of requested slides
                      'of this color
                   
39610                 For k = 1 To iSlidesForColor
                          
                          'the color indicates a special set of slides is needed:
39620                     If dicTemplateSlides.Exists(strColorName) Then
                            
                              Dim d As Dictionary
                              Dim n As Integer
                              Dim ts As TemplateSlide
                       
                              'get the pre defined collection of slides
                              'indicated by this color name:
39630                         Set d = dicTemplateSlides(strColorName)
                              
39640                         For n = 0 To d.Count - 1
39650                             Set ts = d.Items(n)
                                  
39660                         Call EnterNewSlideRequest(strExtraRequestId, _
    strBlockName, strBlockId, ts.GetColor, ts.GetLayers, strShowMicrotom, _
    strColorGroups, iCount)
39670                         Next n
                              
                          'create only this one slide:
39680                     Else
                                        
39690                         Call EnterNewSlideRequest(strExtraRequestId, _
    strBlockName, strBlockId, strColorName, "", strShowMicrotom, strColorGroups, _
    iCount)
39700                     End If
                      
39710                 Next k
39720             Next j
                  
                  'update the block station to 4:
                  'Call UpdateAliquotTrace(strBlockId, "Cleen up")
                  
      '            Call UpdateSlides4Cassette(strBlockId, strBlockName)
39730         End If
39740     Next i

39750     strColorGroups = Mid(strColorGroups, 2)
39760     strColorGroups = "(" & strColorGroups & ")"

39770     If isOldSdg Then
39780         Call sdg_log.InsertLog(rsSdg("sdg_id"), "EXTRA.MANUAL_CREATED", _
    "Add# " & iCount & " slide(s) " & strColorGroups)
39790     Else
39800         Call sdg_log.InsertLog(rsSdg("sdg_id"), "EXTRA.CREATED", "Add# " _
    & iCount & " slide(s) " & strColorGroups)

39810     End If
39820     Exit Sub
ERR_ApproveNewSlides:
39830 MsgBox "ERR_ApproveNewSlides" & vbCrLf & Err.Description
End Sub
Private Function FixQuotes(ByVal stringWithSingleQuotes As String)
39840     FixQuotes = Replace(stringWithSingleQuotes, "'", "''")
End Function
'PAT 002
Private Sub ApproveNewSlidesForSample()
39850 On Error GoTo ERR_ApproveNewSlidesForSample
          Dim i As Integer
          Dim j As Integer
          Dim k As Integer
          Dim iSlidesForColor As Integer
          Dim iCount As Integer
          Dim strExtraRequestId As String
          Dim strExtraRequestDataID As String
          Dim SlideAliquotID As Long
          Dim SampleName As String
          Dim SampleID As String
          Dim strColorName As String
          Dim strColorCount As String
          Dim strShowMicrotom As String
          Dim dColors As Dictionary
          
          'a string of all color groups for the slides;
          'written as part of the description for the sdg-log:
          Dim strColorGroups As String
      '    MsgBox 2
          'insert data to the extra_request table:
          '(save the remarks)
39860     strExtraRequestId = EnterNewRequest("Add Slide", txtColors.Text)
             
          'go through all the blocks to find which are selected to add slides to:
      '    For i = 1 To cmdBlock.Count - 1
39870     For i = 1 To BlockAddSlideList.d.Count
39880         If BlockAddSlideList.d(i).dicColors.Count > 0 Then
              'If CmbBlockEntry(i).ListCount > 0 Then
39890             SampleName = rsSdg("name") & "." & _
    BlockAddSlideList.d(i).strSample
39900             SampleID = GetSampleId(SampleName)
                  
39910             strShowMicrotom = "F"
39920             If BlockAddSlideList.d(i).iMicrotom = 1 Then
                  'If chkBlockMicrotom(i).Value = 1 Then
39930                 strShowMicrotom = "T"
39940             End If
              
                  'go through the items (requested slides) in the list of this Sample:
39950             Set dColors = BlockAddSlideList.d(i).dicColors
                  'For j = 0 To CmbBlockEntry(i).ListCount - 1
39960             For j = 0 To dColors.Count - 1
39970                 Call ParseBlockColor(CStr(dColors.Items(j)), strColorName, _
    strColorCount)
                      'Call ParseBlockColor(CmbBlockEntry(i).list(j), strColorName, strColorCount)
39980                 iSlidesForColor = CInt(strColorCount)
                      
                      'loop for the number of requested slides
                      'of this color
39990                 For k = 1 To iSlidesForColor
                          
                          'the color indicates a special set of slides is needed:
40000                     If dicTemplateSlides.Exists(strColorName) Then
                            
                              Dim d As Dictionary
                              Dim n As Integer
                              Dim ts As TemplateSlide
                       
                              'get the pre defined collection of slides
                              'indicated by this color name:
40010                         Set d = dicTemplateSlides(strColorName)
                              
40020                         For n = 0 To d.Count - 1
40030                             Set ts = d.Items(n)
      'TODO PAT 002 update EnterNewSlideRequestForSample
40040                         Call _
    EnterNewSlideRequestForSample(strExtraRequestId, SampleName, SampleID, _
    ts.GetColor, ts.GetLayers, strShowMicrotom, strColorGroups, iCount)
40050                         Next n
                              
                          'create only this one slide:
40060                     Else
                                        
40070                         Call _
    EnterNewSlideRequestForSample(strExtraRequestId, SampleName, SampleID, _
    strColorName, "", strShowMicrotom, strColorGroups, iCount)
40080                     End If
                      
40090                 Next k
40100             Next j
                  
                  'update the block station to 4:
                  'Call UpdateAliquotTrace(sampleId, "Cleen up")
                  
      '            Call UpdateSlides4Cassette(sampleId, sampleName)
40110         End If
40120     Next i

40130     strColorGroups = Mid(strColorGroups, 2)
40140     strColorGroups = "(" & strColorGroups & ")"
          
40150     If isOldSdg Then
40160         Call sdg_log.InsertLog(rsSdg("sdg_id"), "EXTRA.MANUAL_CREATED", _
    "Add* " & iCount & " slide(s) " & strColorGroups)
40170     Else
40180         Call sdg_log.InsertLog(rsSdg("sdg_id"), "EXTRA.CREATED", "Add* " _
    & iCount & " slide(s) " & strColorGroups)
40190     End If

40200     Exit Sub
ERR_ApproveNewSlidesForSample:
40210 MsgBox "ERR_ApproveNewSlidesForSample" & vbCrLf & Err.Description
40220 Err.Raise Err.Number

End Sub

'1. enter a new extra request data for a requested slide
'2. create the requested slide itself
Private Sub EnterNewSlideRequest(ByVal strExtraRequestId As String, ByVal _
    strBlockName As String, ByVal strBlockId As String, ByVal strColorName As _
    String, ByVal strLayers As String, ByVal strShowMicrotom As String, ByRef _
    strColorGroups As String, ByRef iCount As Integer)
40230 On Error GoTo ERR_EnterNewSlideRequest
          Dim strExtraRequestDataID As String
          Dim SlideAliquotID As Long
      '    MsgBox 3
40240     strExtraRequestDataID = EnterNewRequestData(strExtraRequestId, _
    strBlockName, "Block", strColorName, strShowMicrotom)
40250      If Not cellBlockRequestDataIdDic.Exists(strExtraRequestDataID) Then
40260         Call cellBlockRequestDataIdDic.Add(strExtraRequestDataID, Null)
40270      End If
40280     SlideAliquotID = ExistsCanceledSlide(strBlockId)
          
          'update an existing slide of status X:
40290     If SlideAliquotID <> 0 Then
              
40300         Call UnCancelAliquot(CStr(SlideAliquotID))
          
              'update the slide color:
40310         Call UpdateSlideColor(CStr(SlideAliquotID), strColorName)
40320         Call UpdateSlideLayers(CStr(SlideAliquotID), strLayers)
              
              'update the extra request data for the slide number:
40330         Call UpdateDesc(strExtraRequestDataID, CStr(SlideAliquotID))
40340     Else
                                                           
40350         SlideAliquotID = TriggerSlideEvent("Add Slide", strBlockId)
40360         If SlideAliquotID <> 0 And IsNumeric(SlideAliquotID) Then
40370             SlideAliquotID = GetMaxSlide(SlideAliquotID)
                  
                  'update the slide color:
40380             Call UpdateSlideColor(CStr(SlideAliquotID), strColorName)
40390             Call UpdateSlideLayers(CStr(SlideAliquotID), strLayers)
                  
                  'update new slide name before calling UpdateDesc(),
                  'for it updates the request to also hold the slide name:
40400             Call UpdateSlides4Cassette(strBlockId, strBlockName)
                  
                  'update the extra request data for the slide number:
40410             Call UpdateDesc(strExtraRequestDataID, CStr(SlideAliquotID))
40420         End If
                                   
40430     End If
          
40440     strColorGroups = "," & GetColorGroup(strColorName) & strColorGroups
                                   
40450     iCount = iCount + 1
          
          'update the block station to 4:
40460     Call UpdateAliquotTrace(strBlockId, "Cleen up")

40470     Exit Sub
ERR_EnterNewSlideRequest:
40480 MsgBox "ERR_EnterNewSlideRequest" & vbCrLf & Err.Description
End Sub

'PAT 002

'1. enter a new extra request data for a requested slide
'2. create the requested slide itself
Private Sub EnterNewSlideRequestForSample(ByVal strExtraRequestId As String, _
    ByVal strSampleName As String, ByVal strSampleId As String, ByVal strColorName _
    As String, ByVal strLayers As String, ByVal strShowMicrotom As String, ByRef _
    strColorGroups As String, ByRef iCount As Integer)
40490 On Error GoTo ERR_EnterNewSlideRequestForSample
          Dim strExtraRequestDataID As String
          Dim SlideAliquotID As Long
          Dim EntityType  As String
40500     EntityType = "Sample"
          'Pat 002
40510     If IsPapLbc Then EntityType = "LBC"
40520     strExtraRequestDataID = EnterNewRequestData(strExtraRequestId, _
    strSampleName, EntityType, strColorName, strShowMicrotom)
40530     SlideAliquotID = ExistsCanceledSlide(strSampleId)
          
          'update an existing slide of status X:
40540     If SlideAliquotID <> 0 Then
              
40550         Call UnCancelAliquot(CStr(SlideAliquotID))
          
              'update the slide color:
40560         Call UpdateSlideColor(CStr(SlideAliquotID), strColorName)
40570         Call UpdateSlideLayers(CStr(SlideAliquotID), strLayers)
              
              'update the extra request data for the slide number:
40580         Call UpdateDesc(strExtraRequestDataID, CStr(SlideAliquotID))
40590     Else
                                                           
40600         SlideAliquotID = TriggerSampleEvent("Add Empty Aliquot", _
    strSampleId)
40610         If SlideAliquotID <> 0 And IsNumeric(SlideAliquotID) Then
40620             SlideAliquotID = GetMaxSlideForSample(strSampleId)
                  
                  'update the slide color:
40630             Call UpdateSlideColor(CStr(SlideAliquotID), strColorName)
40640             Call UpdateSlideLayers(CStr(SlideAliquotID), strLayers)
                  
                  'update new slide name before calling UpdateDesc(),
                  'for it updates the request to also hold the slide name:
40650             Call UpdateSlides4Sample(strSampleId, strSampleName)
                  
                  'update the extra request data for the slide number:
40660             Call UpdateDesc(strExtraRequestDataID, CStr(SlideAliquotID))
40670         End If
                                   
40680     End If
          
40690     strColorGroups = "," & GetColorGroup(strColorName) & strColorGroups
                                   
40700     iCount = iCount + 1
          
          'update the Sample station to 4:
          'canceled for pat 002
      '    Call UpdateAliquotTrace(strSampleId, "Cleen up")

40710     Exit Sub
ERR_EnterNewSlideRequestForSample:
40720 MsgBox "ERR_EnterNewSlideRequestForSample" & vbCrLf & Err.Description
40730 Err.Raise Err.Number
End Sub

Private Sub UnCancelResult(strResultId As String)
40740 On Error GoTo ERR_UnCancelResult

          Dim rs As Recordset
          Dim sql As String
          
40750     sql = " select r.OLD_STATUS"
40760     sql = sql & " from lims_sys.result r"
40770     sql = sql & " where r.RESULT_ID='" & strResultId & "'"
          
40780     Set rs = connection.Execute(sql)
40790     If rs.EOF Then Exit Sub
          
40800     Call UpdateResultStatus(strResultId, Right(nte(rs("OLD_STATUS")), 1))

40810     Exit Sub
ERR_UnCancelResult:
40820 MsgBox "ERR_UnCancelResult" & vbCrLf & Err.Description
End Sub


Private Sub UnCancelTest(strTestId As String)
40830 On Error GoTo ERR_UnCancelTest

          Dim rs As Recordset
          Dim sql As String
          
          'test:
40840     sql = " select t.OLD_STATUS"
40850     sql = sql & " from lims_sys.test t"
40860     sql = sql & " where t.TEST_ID='" & strTestId & "'"
          
40870     Set rs = connection.Execute(sql)
40880     If rs.EOF Then Exit Sub
          
40890     Call UpdateTestStatus(strTestId, Right(nte(rs("OLD_STATUS")), 1))


          'related results:
40900     sql = " select r.RESULT_ID"
40910     sql = sql & " from lims_sys.result r"
40920     sql = sql & " where r.TEST_ID='" & strTestId & "'"
40930     sql = sql & " and   r.STATUS='X'"
          
40940     Set rs = connection.Execute(sql)
          
40950     While Not rs.EOF
40960         Call UnCancelResult(nte(rs("RESULT_ID")))
40970         rs.MoveNext
40980     Wend

40990     Exit Sub
ERR_UnCancelTest:
41000 MsgBox "ERR_UnCancelTest" & vbCrLf & Err.Description
End Sub


Private Sub UnCancelAliquot(strAliquotId As String)
41010 On Error GoTo ERR_UnCancelAliquot

          Dim rs As Recordset
          Dim sql As String
          
          'current aliquot:
41020     sql = " select a.OLD_STATUS"
41030     sql = sql & " from lims_sys.aliquot a"
41040     sql = sql & " where a.ALIQUOT_ID='" & strAliquotId & "'"
          
41050     Set rs = connection.Execute(sql)
41060     If rs.EOF Then Exit Sub
          
41070     Call UpdateAliquotStatus(strAliquotId, Right(nte(rs("OLD_STATUS")), _
    1))
          
          
          'tests:
41080     sql = " select t.TEST_ID"
41090     sql = sql & " from lims_sys.test t"
41100     sql = sql & " where t.ALIQUOT_ID='" & strAliquotId & "'"
41110     sql = sql & " and   t.STATUS='X'"
          
41120     Set rs = connection.Execute(sql)
          
41130     While Not rs.EOF
41140         Call UnCancelTest(nte(rs("TEST_ID")))
41150         rs.MoveNext
41160     Wend
          
          
          'child aliquots:
41170     sql = " select af.CHILD_ALIQUOT_ID"
41180     sql = sql & " from lims_sys.aliquot_formulation af,"
41190     sql = sql & "      lims_sys.aliquot a"
41200     sql = sql & " where a.ALIQUOT_ID=af.CHILD_ALIQUOT_ID"
41210     sql = sql & " and   af.PARENT_ALIQUOT_ID='" & strAliquotId & "'"
41220     sql = sql & " and   a.STATUS='X'"
          
41230     Set rs = connection.Execute(sql)
          
41240     While Not rs.EOF
41250         Call UnCancelAliquot(nte(rs("CHILD_ALIQUOT_ID")))
41260         rs.MoveNext
41270     Wend
          
41280     Exit Sub
ERR_UnCancelAliquot:
41290 MsgBox "ERR_UnCancelAliquot" & vbCrLf & Err.Description
End Sub


Private Function ExistsCanceledSlide(strBlockId As String) As Long
41300 On Error GoTo ERR_ExistsCanceledSlide
          Dim rs As Recordset
          Dim sql As String
          
41310     ExistsCanceledSlide = 0
          
41320     sql = " select a.ALIQUOT_ID"
41330     sql = sql & " from lims_sys.aliquot a"
41340     sql = sql & " where a.STATUS='X'"
41350     sql = sql & " and exists"
41360     sql = sql & " ("
41370     sql = sql & "   select 1 "
41380     sql = sql & "   from lims_sys.aliquot_formulation af"
41390     sql = sql & "   where af.CHILD_ALIQUOT_ID=a.ALIQUOT_ID"
41400     sql = sql & "   and   af.PARENT_ALIQUOT_ID='" & strBlockId & "'"
41410     sql = sql & " )"
41420     sql = sql & " order by a.ALIQUOT_ID"
          
41430     Set rs = connection.Execute(sql)
          
41440     If Not rs.EOF Then
41450         ExistsCanceledSlide = rs("ALIQUOT_ID")
41460     End If

41470     Exit Function
ERR_ExistsCanceledSlide:
41480 MsgBox "ERR_ExistsCanceledSlide" & vbCrLf & Err.Description
End Function
'PAT 002
Private Function ExistsCanceledSlideForSample(strSampleId As String) As Long
41490 On Error GoTo ERR_ExistsCanceledSlideForSample
          Dim rs As Recordset
          Dim sql As String
          
41500     ExistsCanceledSlideForSample = 0
          
41510     sql = "  select a.ALIQUOT_ID"
41520     sql = sql & "  from lims_sys.aliquot a"
41530     sql = sql & "  where a.sample_id= " & strSampleId & " "
41540     sql = sql & "  and a.STATUS='X'"
41550     sql = sql & "  and not exists"
41560     sql = sql & "  ("
41570     sql = sql & "    select 1 "
41580     sql = sql & "    from lims_sys.aliquot_formulation af"
41590     sql = sql & "    where af.Parent_ALIQUOT_ID=a.ALIQUOT_ID"
41600     sql = sql & "  )"
41610     sql = sql & "  order by a.ALIQUOT_ID"
          
41620     Set rs = connection.Execute(sql)
          
41630     If Not rs.EOF Then
41640         ExistsCanceledSlideForSample = rs("ALIQUOT_ID")
41650     End If

41660     Exit Function
ERR_ExistsCanceledSlideForSample:
41670 MsgBox "ERR_ExistsCanceledSlideForSample" & vbCrLf & Err.Description
End Function
'generic code to enter a new request
Private Function EnterNewRequest(name As String, Description As String) As _
    String
41680 On Error GoTo ERR_EnterNewRequest
          Dim iExtraRequest As Recordset
          Dim sql As String

          'insert data to the extra_request table:
41690     Set iExtraRequest = _
    connection.Execute("select lims.sq_u_extra_request.nextval from dual")
41700     sql = " insert into lims_sys.u_extra_request"
41710     sql = sql & _
    " (u_extra_request_id, name, description, version, version_status)"
41720     sql = sql & " values"
41730     sql = sql & " (" & iExtraRequest(0) & ", '" & name & ";" & _
    iExtraRequest(0) & "', '" & FixQuotes(Description) & "','1','A')"
41740     connection.Execute (sql)

41750     sql = " insert into lims_sys.u_extra_request_user"
41760     sql = sql & _
    " (u_extra_request_id, u_sdg_id, u_created_on, u_created_by, u_status)"
41770     sql = sql & " values"
41780     sql = sql & " (" & iExtraRequest(0) & ", " & rsSdg("sdg_id") & _
    ", to_char(sysdate), " & NtlsUser.GetOperatorId & ",'V')"
41790     connection.Execute (sql)

41800     EnterNewRequest = iExtraRequest(0)

41810     Exit Function
ERR_EnterNewRequest:
41820 MsgBox "ERR_EnterNewRequest" & vbCrLf & Err.Description
41830 Err.Raise Err.Number
End Function

'generic code to enter new request's details
Private Function EnterNewRequestData(ExtraRequestId As String, name As String, _
    EntityType As String, RequestDetails As String, strDescription As String) As _
    String
41840 On Error GoTo ERR_EnterNewRequestData
          Dim iExtraRequestData As Recordset
          Dim sql As String
          
41850     Set iExtraRequestData = _
    connection.Execute("select lims.sq_u_extra_request_data.nextval from dual")
41860     sql = " insert into lims_sys.u_extra_request_data"
41870     sql = sql & _
    " (u_extra_request_data_id, name, description, version, version_status)"
41880     sql = sql & " values"
41890     sql = sql & " (" & iExtraRequestData(0) & ", '" & name & ";" & _
    iExtraRequestData(0) & "','" & FixQuotes(strDescription) & "','1','A')"
41900     connection.Execute (sql)
          
41910     sql = " insert into lims_sys.u_extra_request_data_user"
41920     sql = sql & _
    " (u_extra_request_data_id, u_extra_request_id, u_entity_type, u_request_details, u_status)"
41930     sql = sql & " values"
41940     sql = sql & " (" & iExtraRequestData(0) & "," & ExtraRequestId & ", '" _
    & EntityType & "', '" & FixQuotes(RequestDetails) & "', 'V')"
41950     connection.Execute (sql)
          
41960     EnterNewRequestData = iExtraRequestData(0)
          
41970     Exit Function
ERR_EnterNewRequestData:
41980 MsgBox "ERR_EnterNewRequestData" & vbCrLf & Err.Description
41990 Err.Raise Err.Number
End Function


Private Sub ApproveExistingSlides()
42000 On Error GoTo ERR_ApproveExistingSlides
          Dim i As Integer
          
          Dim iCount As Integer
          Dim strSlideName As String
          Dim strSlideId As String
          Dim strBlockId As String
          Dim strRequestId As String
          'hold unique list of blocks to update aliquot trace:
      '    Dim dicBlocksToUpdateTrace As New Dictionary
          
          'a string of all color groups for the slides;
          'written as part of the description for the sdg-log:
          Dim strColorGroups As String
          Dim strExtraRequestDataID As String

          'insert data to the extra_request table:
42010     If isOldSdg Then
42020         strRequestId = EnterNewRequest("Color Manual Slide", _
    txtColors.Text)
42030     Else
42040         strRequestId = EnterNewRequest("Color Slide", txtColors.Text)
42050     End If
              
42060      If Not cellBlockRequestIdDic.Exists(strRequestId) Then
42070         Call cellBlockRequestIdDic.Add(strRequestId, Null)
42080     End If
          'insert data to the extra_request_data table:
42090     For i = 1 To cmdSlide.Count - 1
42100         If txtSlide(i).Text <> "" Then
42110             strSlideName = rsSdg("name") & "." & cmdSlide(i).Caption
              
42120             strExtraRequestDataID = EnterNewRequestData(strRequestId, _
    strSlideName, "Reserve-Slide", txtSlide(i).Text, "")
42130             If Not _
    cellBlockRequestDataIdDic.Exists(strExtraRequestDataID) Then
42140                 Call cellBlockRequestDataIdDic.Add(strExtraRequestDataID, _
    Null)
42150             End If
                  

42160             strColorGroups = "," & GetColorGroup(txtSlide(i).Text) & _
    strColorGroups
                  
42170             strSlideId = GetAliquotId(strSlideName)
42180             strBlockId = GetBlockId(strSlideId)
                  
42190             Call UpdateSlideColor(strSlideId, txtSlide(i).Text)
                  
      '            If Not dicBlocksToUpdateTrace.Exists(strBlockId) Then
      '                Call dicBlocksToUpdateTrace.Add(strBlockId, "")
      '            End If
                  
                  'Call UpdateReserveSlide(strSlideName, txtSlide(i).Text)
                  
      '            Call UpdateAliquotTrace(GetBlockId(GetAliquotId(strSlideName)), "Cleen up")
                  
42200             iCount = iCount + 1
42210         End If
42220     Next i
          
          'update the station once for every relevant block:
      '    For i = 0 To dicBlocksToUpdateTrace.Count - 1
      '        Call UpdateAliquotTrace(CStr(dicBlocksToUpdateTrace.Keys(i)), "Cleen up")
      '    Next i

42230     strColorGroups = Mid(strColorGroups, 2)
42240     strColorGroups = "(" & strColorGroups & ")"
          
42250     If isOldSdg Then
42260         Call sdg_log.InsertLog(rsSdg("sdg_id"), "EXTRA.MANUAL_CREATED", _
    "Color " & iCount & " slide(s) " & strColorGroups)
42270     Else
42280         Call sdg_log.InsertLog(rsSdg("sdg_id"), "EXTRA.CREATED", "Color " _
    & iCount & " slide(s) " & strColorGroups)
42290     End If

42300     Exit Sub
ERR_ApproveExistingSlides:
42310 MsgBox "ERR_ApproveExistingSlides" & vbCrLf & Err.Description
End Sub

Private Sub cmdOKReEmbedding_Click()
42320 On Error GoTo ERR_cmdOKReEmbedding_Click
          Dim i As Integer
          Dim iCount As Integer
          Dim strExtraRequestId As String
          Dim strBlockName As String
          Dim bre As BlockReembedding
          
42330     strExtraRequestId = EnterNewRequest("Re Embedding", _
    txtReembedding.Text)

      '    For i = 1 To chkBlockReembeding.Count - 1
42340     For i = 1 To BlockReembeddingList.d.Count
42350         If BlockReembeddingList.d(i).iReembedding = 1 Then
42360             strBlockName = rsSdg("name") & "." & _
    BlockReembeddingList.d(i).strName
              
42370             Call EnterNewRequestData(strExtraRequestId, strBlockName, _
    "Block", BlockReembeddingList.d(i).strReason, _
    BlockReembeddingList.d(i).strDetails)
42380             iCount = iCount + 1
                  
42390             Call UpdateAliquotTrace(GetAliquotId(strBlockName), _
    "Result Entry")
42400         End If
42410     Next i
          
          'reset memory list:
42420     For i = 1 To BlockReembeddingList.d.Count
42430         Set bre = BlockReembeddingList.d(i)
42440         bre.iReembedding = 0
42450         bre.strReason = ""
42460         bre.strDetails = ""
42470     Next i
          
          'reset the re embedding tab:
42480     For i = 1 To chkBlockReembeding.Count - 1
42490         chkBlockReembeding(i).value = 0
42500         cmbReembeddingReason(i).Text = ""
42510         txtReembeddingDetails(i).Text = ""
42520     Next i
42530     txtReembedding.Text = ""
42540     iUpdateReEmbedding = 0
42550     cmdOKReEmbedding.Enabled = False
          
42560     Call InitExtraRequestsHistory
          
42570     Call sdg_log.InsertLog(rsSdg("sdg_id"), "EXTRA.CREATED", _
    "Re Embedding " & iCount & " block(s)")
          
42580     MsgBox "The request was saved successfully"
          
42590     Exit Sub
ERR_cmdOKReEmbedding_Click:
42600 MsgBox "ERR_cmdOKReEmbedding_Click" & vbCrLf & Err.Description
End Sub

Private Sub cmdOKShowOriginalSample_Click()
42610 On Error GoTo ERR_cmdOKShowOriginalSample_Click
          Dim i As Integer
          Dim iCount As Integer
          Dim strExtraRequestId As String
          Dim strSampleName As String
          
42620     strExtraRequestId = EnterNewRequest("Present Original Sample", _
    txtShowOriginalSample.Text)

42630     For i = 1 To chkShowSample.Count - 1
42640         If chkShowSample(i).value = 1 Then
42650             strSampleName = rsSdg("name") & "." & chkShowSample(i).Caption
              
42660             Call EnterNewRequestData(strExtraRequestId, strSampleName, _
    "Sample", "", "")
              'ביטול הקריאה לפונק' לפי בקשת מכבי- פתולוגיה , בבקשות חוזרות->מסירה נוספת , לא לעדכן את
              'ALIQUOT_STATION להיות 1
              ' Call UpdateAliquotesForSample(strSampleName)
              ''''''''''''
42670             iCount = iCount + 1
42680         End If
42690     Next i
          
          'reset the Add Block tab
42700     For i = 1 To chkShowSample.Count - 1
42710         chkShowSample(i).value = 0
42720     Next i
42730     txtShowOriginalSample.Text = ""
42740     cmdOKShowOriginalSample.Enabled = False
42750     iUpdateShowSample = 0
          
42760     Call InitExtraRequestsHistory
          
42770     Call sdg_log.InsertLog(rsSdg("sdg_id"), "EXTRA.CREATED", "Present " & _
    iCount & " original sample(s)")
          
42780     MsgBox "The request was saved successfully"
          
42790     Exit Sub
ERR_cmdOKShowOriginalSample_Click:
42800 MsgBox "ERR_cmdOKShowOriginalSample_Click" & vbCrLf & Err.Description
End Sub

Private Sub cmdOKSlidesFromArchive_Click()
42810 On Error GoTo ERR_cmdOKSlidesFromArchive_Click
          Dim i As Integer
          Dim strExtraRequestId As String
          Dim strRequestDetails As String
          Dim strName As String
          Dim strEntityType As String
          
42820     strExtraRequestId = EnterNewRequest("Slides From Archive", _
    txtSlidesFromArchive.Text)
          
42830     For i = 0 To lstSlidesFromArchive.ListCount - 1
42840         strRequestDetails = ""
42850         strName = lstSlidesFromArchive.list(i)
42860         strEntityType = GetEntityType(strName)
              
              'signify that the name is not of a slide, because the SDG
              'is an old one and slides were not created back then:
42870         If strEntityType <> "Slide" Then
42880             strRequestDetails = strEntityType & " (No List of Slides)"
42890         End If
          
42900         Call EnterNewRequestData(strExtraRequestId, _
    lstSlidesFromArchive.list(i), "Stained-Slide", strRequestDetails, "")
42910     Next i


          
42920     Call sdg_log.InsertLog(rsSdg("sdg_id"), "EXTRA.CREATED", "Get " & _
    lstSlidesFromArchive.ListCount & " slide(s) from archive")
          
          
42930     Call InitExtraRequestsHistory
          
          
          'display a message:
42940     MsgBox "The request was saved successfully"
          
          'clean the fields of this tab:
42950     lstSlidesFromArchive.Clear
42960     txtSlidesFromArchive.Text = ""
42970     lblStainSlides.Caption = 0
          
42980     cmdOKSlidesFromArchive.Enabled = False

42990     Exit Sub
ERR_cmdOKSlidesFromArchive_Click:
43000 MsgBox "ERR_cmdOKSlidesFromArchive_Click" & vbCrLf
End Sub

'Private Sub cmdPrev_Click()
'On Error GoTo ERR_cmdPrev_Click
'
'    If Not rsExtraRequests.BOF Then
'        rsExtraRequests.MovePrevious
'
' '       Call SetExtraRequestInfo
'
'        txtRequestNumber.Text = txtRequestNumber.Text - 1
'
'        cmdNext.Enabled = True
'    End If
'
'    If txtRequestNumber.Text = 1 Then
'        cmdPrev.Enabled = False
'    End If
'
'    Exit Sub
'ERR_cmdPrev_Click:
'MsgBox "ERR_cmdPrev_Click" & vbCrLf & Err.description
'End Sub

Private Sub cmdSelectEntity_Click()
43010 On Error GoTo ERR_cmdSelectEntity_Click
          Dim sText As String
          Dim sKey As String
          
          'check if no item was selected:
43020     sText = tree.SelectedItem.Text
43030     sText = Mid(sText, 1, InStr(1, sText, " ") - 1)
43040     If sText = "" Then Exit Sub
          
          'check if item already in the list:
43050     If ExistInList(sText, lstSelectedObjects) > -1 Then Exit Sub
          
          'check if the item is not a block / slide:
43060     sKey = tree.SelectedItem.Key
43070     If InStr(1, sKey, "block") = 0 And InStr(1, sKey, "slide") = 0 Then _
    Exit Sub
        
43080     Call lstSelectedObjects.AddItem(sText)
43090     lblAdvisorCount = lblAdvisorCount + 1
43100     cmdOKAdvisors.Enabled = True
          
43110     Exit Sub
ERR_cmdSelectEntity_Click:
43120 MsgBox "ERR_cmdSelectEntity_Click" & vbCrLf & Err.Description
End Sub


Private Sub InitAdvisorList()
43130 On Error GoTo ERR_InitAdvisorList
          Dim rs As Recordset
          Dim sql As String
          Dim s As String
          
43140     sql = "   select o.name, ou.U_DEGREE"
43150     sql = sql & "   from lims_sys.operator o,"
43160     sql = sql & "        lims_sys.operator_user ou,"
43170     sql = sql & "        lims_sys.operator_role r, "
43180     sql = sql & "        lims_sys.lims_role lr"
43190     sql = sql & "   where o.OPERATOR_ID=r.OPERATOR_ID"
43200     sql = sql & "   and   r.ROLE_ID=lr.ROLE_ID"
43210     sql = sql & "   and   o.OPERATOR_ID=ou.OPERATOR_ID"
43220     sql = sql & "   and   lr.NAME='External Advisor'"
          
43230     Set rs = connection.Execute(sql)

43240     While Not rs.EOF
43250         s = nte(rs("u_degree"))
43260         If s <> "" Then
43270             s = s & " "
43280         End If
43290         s = s & nte(rs("name"))
          
43300         Call lstAdvisors.AddItem(s)
43310         rs.MoveNext
43320     Wend
              
43330     If lstAdvisors.ListCount > 0 Then
43340         lstAdvisors.Selected(0) = True
43350     End If

43360     Exit Sub
ERR_InitAdvisorList:
43370 MsgBox "GoTo ERR_InitAdvisorList" & vbCrLf & Err.Description
End Sub






Private Sub cmdSlide_Click(index As Integer)
43380     Call frmColors.Initialize(0, index, cmdSlide(index).Caption, _
    dicMolecularStains, dicSpecialStains, dicImonohistochemistryStains, _
    dicHistochemistryStains, dicOtherStainOptions)
43390     frmColors.Show vbModal
          
          
          'special stains. indicates a change of all the slides
          'of that block to some color (CKMNF116 or S100)
43400     Select Case txtSlide(index).Text
              Case "SLN / Reserve / CKMNF116"
43410             Call _
    ChangeReserveSlidesForBlock(GetBlockName(cmdSlide(index).Caption), "CKMNF116")
              
43420         Case "SLN / Reserve / S-100"
43430             Call _
    ChangeReserveSlidesForBlock(GetBlockName(cmdSlide(index).Caption), "S-100")
43440     End Select
End Sub

Private Sub cmdSlideReset_Click(index As Integer)
43450     If txtSlide(index).Text = "" Then Exit Sub
          
43460     txtSlide(index).Text = ""
43470     Call DecrementColorsSlides
End Sub

'change all reserve slides for this block to the new stain
'strBlockNumber (in) - the identifier of the block name
'in the SDG, like 1.1
Private Sub ChangeReserveSlidesForBlock(strBlockNumber As String, strNewStain _
    As String)
43480 On Error GoTo ERR_ChangeReserveSlidesForBlock

          Dim i As Integer
          
43490     For i = 1 To cmdSlide.Count - 1
43500         If GetBlockName(cmdSlide(i).Caption) = strBlockNumber Then
43510             txtSlide(i).Text = strNewStain
43520         End If
43530     Next i

43540     Exit Sub
ERR_ChangeReserveSlidesForBlock:
43550 MsgBox "ERR_ChangeReserveSlidesForBlock" & vbCrLf & Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
43560 On Error GoTo ERR_Form_Unload
          'erase classes to prevent remains of unreferenced memory
          Dim i As Integer
      '     If isCito And NewCellBlockID <> "" Then
      '        If MsgBox("? לדגימה. האם לאשר  CELL - BLOCK נוסף ? ", vbYesNoCancel) <> vbYes Then
      '            Call UpdateAliquotStatus(NewCellBlockID, "X")
      '            NewCellBlockID = ""
      '            For i = 0 To cellBlockRequestDataIdDic.Count - 1
      '                Call UpdateExtraRequestDataStatus(CStr(cellBlockRequestDataIdDic.Keys(i)), "X")
      '            Next i
      '            For i = 0 To cellBlockRequestIdDic.Count - 1
      '                Call UpdateExtraRequestStatus(CStr(cellBlockRequestIdDic.Keys(i)), "X")
      '            Next i
      ''            Call DeleteExtraRequest(cellBlockRequestId, cellBlockRequestDataId)
      '        End If
      '    End If
          
43570     isCito = False
43580     Call dicTemplateSlides.RemoveAll
43590     Call BlockAddSlideList.d.RemoveAll
43600     Call BlockReembeddingList.d.RemoveAll
          
43610     Exit Sub
ERR_Form_Unload:
43620 MsgBox "ERR_Form_Unload" & vbCrLf & Err.Description
End Sub

'Private Sub DeleteExtraRequest(cellBlockRequestIdDicAs String, cellBlockRequestDataIdDicAs String)
'On Error GoTo Err_DeleteExtraRequest
'Dim sql As String
'
'    sql = ("delete from lims_sys.u_extra_request_data_user where U_EXTRA_REQUEST_DATA_ID=" & cellBlockRequestDataId)
'    connection.Execute (sql)
'    sql = ("delete from lims_sys.u_extra_request_data where U_EXTRA_REQUEST_DATA_ID=" & cellBlockRequestDataId)
'    connection.Execute (sql)
'
'If cellBlockRequestIdDic<> "" Then
'    sql = ("delete from lims_sys.u_extra_request_user where U_EXTRA_REQUEST_ID=" & cellBlockRequestId)
'    connection.Execute (sql)
'    sql = ("delete from lims_sys.u_extra_request where U_EXTRA_REQUEST_ID=" & cellBlockRequestId)
'    connection.Execute (sql)
'
'Exit Sub
'Err_DeleteExtraRequest:
' MsgBox "Err_DeleteExtraRequest : " & vbCrLf & Err.Description & vbCrLf & ' "Sql ran : " & sql
' End Sub
' cancel of an extra request
' allow if extra request status is 'V' or 'P'
' for the most advanced entity in the request
Private Sub grid_KeyDown(KeyCode As Integer, Shift As Integer)
43630 On Error GoTo ERR_grid_KeyDown
          Dim strExtraRequestDataID As String
          Dim strExtraRequestId As String
          Dim strStatus As String
          Dim strAction As String
          Dim mbr As VbMsgBoxResult
          Dim iCount As Integer
          Dim iBegin As Integer
          Dim iEnd   As Integer
          Dim i As Integer
          Dim iTemp As Integer
          Dim d As New Dictionary


43640     If KeyCode <> vbKeyDelete Then
43650         Exit Sub
43660     End If
          
          
          
43670     iBegin = grid.row
43680     iEnd = grid.RowSel
          
43690     If iBegin > iEnd Then
43700         iTemp = iBegin
43710         iBegin = iEnd
43720         iEnd = iTemp
43730     End If
          
          'the selected rows to cancel:
          
43740     For i = iBegin To iEnd
          'For i = grid.row To grid.RowSel
              
43750         If CanCancelExtraRequestData(grid.TextMatrix(i, 0)) = False Then
43760             MsgBox "Can not cancel entities that are already executed", _
    vbExclamation
43770             Exit Sub
43780         End If
              
43790         Call d.Add(grid.TextMatrix(i, 0), "")
              
43800     Next i
          
      '    iCount = iEnd - iBegin + 1
          
43810     mbr = MsgBox("Are you sure you want to cancel the " & d.Count & _
    " selected item(s)?", vbYesNo + vbQuestion)
43820     If mbr = vbNo Then Exit Sub
          
43830     For i = d.Count - 1 To 0 Step -1
43840         Call CancelExtraRequestData(CStr(d.Keys(i)))
43850     Next i
              
      '    For i = grid.row To grid.RowSel
      '        Call CancelExtraRequestData(grid.TextMatrix(i, 0))
      '    Next i
          
43860     Call InitAdvisorsRequestsList(nte(rsSdg("external_reference")))
43870     Call InitExtraRequestsHistory
          

      '    strAction = grid.TextMatrix(grid.row, 2)
      '    strExtraRequestDataID = grid.TextMatrix(grid.row, 0)
      '    strExtraRequestID = CanCancelExtraRequest(strExtraRequestDataID)
      '
      '    If strExtraRequestID = "" Then
      '        MsgBox ("The extra request can not be cancled," & vbCrLf & '                "for some of it's sub-items are already executed")
      '        Exit Sub
      '    End If
      '
      '    mbr = MsgBox("Are you sure you want to cancel the extra request" & vbCrLf & '                 "(the action cancels all brother-items of the chosen item)?", '                 vbYesNo + vbQuestion)
      '    If mbr = vbNo Then Exit Sub
      '
      '    Call CancelExtraRequest(strExtraRequestID, strAction)
      '
      '    Call InitAdvisorsRequestsList(nte(rsSdg("external_reference")))
      '    Call InitExtraRequestsHistory
          
43880     Exit Sub
ERR_grid_KeyDown:
43890 MsgBox "ERR_grid_KeyDown" & vbCrLf & Err.Description
End Sub


'if all the entities under the extra request are canceled,
'cancel the extra request as well:
Private Sub UpdateExtraRequestParentStatus(strExtraRequestId As String)
43900 On Error GoTo ERR_UpdateExtraRequestParentStatus
          
          Dim sql As String

43910     sql = "  update lims_sys.u_extra_request_user eru"
43920     sql = sql & "  set eru.u_status = 'X'"
43930     sql = sql & "  where eru.U_EXTRA_REQUEST_ID='" & strExtraRequestId & _
    "'"
43940     sql = sql & "  and"
43950     sql = sql & "  ("
43960     sql = sql & "    select count(1) "
43970     sql = sql & "    from lims_sys.u_extra_request_data_user erdu"
43980     sql = sql & "    where erdu.U_EXTRA_REQUEST_ID=eru.U_EXTRA_REQUEST_ID"
43990     sql = sql & "    and   erdu.U_STATUS <> 'X'"
44000     sql = sql & "  )=0"
          
44010     Call connection.Execute(sql)

44020     Exit Sub
ERR_UpdateExtraRequestParentStatus:
44030 MsgBox "ERR_UpdateExtraRequestParentStatus" & vbCrLf & Err.Description
End Sub


Private Sub UpdateSlideNameInExtraRequestData(strOldSlideName As String, _
    strNewSlideName As String)
44040 On Error GoTo ERR_UpdateSlideNameInExtraRequestData

          Dim sql As String
          
          'Add Slide request:
44050     sql = " update lims_sys.u_extra_request_data_user rdu"
44060     sql = sql & " set   rdu.U_SLIDE_NAME='" & strNewSlideName & "'"
44070     sql = sql & " where rdu.U_SLIDE_NAME='" & strOldSlideName & "'"

44080     Call connection.Execute(sql)
          
          
          'Color Slide / Send to Consultant / Slides From Archive request:
44090     sql = " update lims_sys.u_extra_request_data rd"
44100     sql = sql & " set rd.NAME = '" & strNewSlideName & _
    "' || ';' || rd.U_EXTRA_REQUEST_DATA_ID"
44110     sql = sql & " where rd.NAME like '" & strOldSlideName & ";%'"

44120     Call connection.Execute(sql)

44130     Exit Sub
ERR_UpdateSlideNameInExtraRequestData:
44140 MsgBox "ERR_UpdateSlideNameInExtraRequestData" & vbCrLf & Err.Description
End Sub


'if the slide to cancel is NOT the last one (biggest id) for the block,
'a shift in the slides data is needed;
'each slide, starting at the one to be canceled to the last non canceled slide
'gets the data of the next slide;
'the slide names are updated in the relevant extra requests;
Private Sub UpdateSlidesData(strSlideId As String, strBlockId As String)
44150 On Error GoTo ERR_UpdateSlideNames

          Dim d As New Dictionary
          Dim dStopList As New Dictionary
          Dim rs As Recordset
          Dim sql As String
          Dim strMaxSlideName As String
          Dim i As Integer
           
44160     Call dStopList.Add("name", "name")
44170     Call dStopList.Add("status", "status")
44180     Call dStopList.Add("old_status", "old_status")
44190     Call dStopList.Add("sdg_id", "sdg_id")
44200     Call dStopList.Add("sample_id", "sample_id")
44210     Call dStopList.Add("aliquot_id", "aliquot_id")
44220     Call dStopList.Add("test_id", "test_id")
44230     Call dStopList.Add("result_id", "result_id")
44240     Call dStopList.Add("workflow_node_id", "workflow_node_id")
44250     Call dStopList.Add("sdg_template_id", "sdg_template_id")
44260     Call dStopList.Add("sample_template_id", "sample_template_id")
44270     Call dStopList.Add("aliquot_template_id", "aliquot_template_id")
44280     Call dStopList.Add("test_template_id", "test_template_id")
44290     Call dStopList.Add("result_template_id", "result_template_id")
          
           
44300     sql = " select a.NAME, a.ALIQUOT_ID"
44310     sql = sql & " from lims_sys.aliquot a"
44320     sql = sql & " where a.ALIQUOT_ID >= '" & strSlideId & "'"
44330     sql = sql & " and   a.STATUS <> 'X'"
          'sql = sql & " and  (a.STATUS <> 'X' or a.ALIQUOT_ID = '" & strSlideId & "')"
44340     sql = sql & " and   exists"
44350     sql = sql & " ("
44360     sql = sql & "    select af.CHILD_ALIQUOT_ID"
44370     sql = sql & "    from lims_sys.aliquot_formulation af"
44380     sql = sql & "    where af.PARENT_ALIQUOT_ID = '" & strBlockId & "'"
44390     sql = sql & "    and   af.CHILD_ALIQUOT_ID = a.ALIQUOT_ID  "
44400     sql = sql & " )"
44410     sql = sql & " order by a.ALIQUOT_ID"

44420     Set rs = connection.Execute(sql)
          
44430     If rs.EOF Then Exit Sub
          
44440     While Not rs.EOF
44450         Call d.Add(nte(rs("ALIQUOT_ID")), nte(rs("NAME")))
44460         rs.MoveNext
44470     Wend
          
44480     For i = 0 To d.Count - 2

44490         Set rs = connection.Execute(" select * from lims_sys.aliquot " & _
    " where aliquot_id = '" & d.Keys(i + 1) & "'")
                                          
44500         Call UpdateRecordById("aliquot", "aliquot_id", CStr(d.Keys(i)), _
    CStr(d.Keys(i + 1)), rs, dStopList)
          
44510         Set rs = _
    connection.Execute(" select * from lims_sys.aliquot_user " & _
    " where aliquot_id = '" & d.Keys(i + 1) & "'")
                                          
44520         Call UpdateRecordById("aliquot_user", "aliquot_id", _
    CStr(d.Keys(i)), CStr(d.Keys(i + 1)), rs, dStopList)
                                     
44530         Call UpdateSlideNameInExtraRequestData(CStr(d.Items(i + 1)), _
    CStr(d.Items(i)))
44540     Next i

44550     Call UpdateAliquotStatus(CStr(d.Keys(d.Count - 1)), "X")

44560     Exit Sub
ERR_UpdateSlideNames:
44570 MsgBox "ERR_UpdateSlideNames" & vbCrLf & Err.Description
End Sub


Private Function CancelExtraRequestData(strExtraRequestDataID As String) As _
    Boolean
44580 On Error GoTo ERR_CancelExtraRequestData

          Dim rs As Recordset
          Dim sql As String
          Dim strLogDesc As String
          Dim strEntityName As String
          Dim strAction As String

44590     CancelExtraRequestData = False

44600     sql = "  select er.NAME ACTION,"
44610     sql = sql & "         erd.U_EXTRA_REQUEST_DATA_ID, "
44620     sql = sql & "         erd.NAME, "
44630     sql = sql & "      erdu.U_REQUEST_DETAILS,"
44640     sql = sql & "      erdu.U_SLIDE_NAME,"
44650     sql = sql & "      erdu.U_DESC,"
44660     sql = sql & "      eru.U_CREATED_BY,"
44670     sql = sql & "      eru.U_CREATED_ON, "
44680     sql = sql & "      eru.U_EXTRA_REQUEST_ID "
44690     sql = sql & "  from lims_sys.u_extra_request er,"
44700     sql = sql & "        lims_sys.u_extra_request_user eru,"
44710     sql = sql & "       lims_sys.u_extra_request_data erd,"
44720     sql = sql & "       lims_sys.u_extra_request_data_user erdu"
44730     sql = sql & "  where eru.U_EXTRA_REQUEST_ID=er.U_EXTRA_REQUEST_ID"
44740     sql = sql & _
    "  and   erd.U_EXTRA_REQUEST_DATA_ID=erdu.U_EXTRA_REQUEST_DATA_ID"
44750     sql = sql & "  and   eru.U_EXTRA_REQUEST_ID=erdu.U_EXTRA_REQUEST_ID"
44760     sql = sql & "  and   erdu.U_EXTRA_REQUEST_DATA_ID='" & _
    strExtraRequestDataID & "'"
          
44770     Set rs = connection.Execute(sql)
          
44780     If rs.EOF Then
44790         Exit Function
44800     End If
          
44810     strAction = CleanSemicolon(nte(rs("ACTION")))
          
          'holds the data on which aliquots should get their
          'aliquot trace data reversed:
          Dim dicReverseTrace As New Dictionary
          
44820     strLogDesc = strAction & ":"

44830     strEntityName = CleanSemicolon(rs("NAME"))

44840     Call UpdateExtraRequestDataStatus(strExtraRequestDataID, "X")
            
44850     If strAction = "Add Slide" Then
              
              'relate to all the brothers slide from the current one
              'to the last, excluding canceled ones:
              '1. update the data of slides (1 shift down)
              '2. set the last slide status to X
44860         Call UpdateSlidesData(GetAliquotId(nte(rs("U_SLIDE_NAME"))), _
    GetAliquotId(strEntityName))
              
              'Call UpdateAliquotStatus(GetAliquotId(nte(rs("U_SLIDE_NAME"))), "X")
              'Call UpdateAliquotStatus(nte(rs("U_DESC")), "X")
44870     End If
          
44880     If strAction = "Color Slide" Then
44890         Call UpdateSlideColor(GetAliquotId(strEntityName), "רזרבה")
44900     End If
          
44910     strLogDesc = strLogDesc & " (" & nte(rs("NAME")) & " - " & _
    nte(rs("U_REQUEST_DETAILS")) & ") "
          
44920     If Not dicReverseTrace.Exists(strEntityName) Then
44930         Call dicReverseTrace.Add(strEntityName, strAction)
44940     End If

44950     Call ComputeAliquotToReverseTrace(dicReverseTrace)
          
44960     Call sdg_log.InsertLog(rsSdg("sdg_id"), "EXTRA.CANCEL", strLogDesc)

44970     Call UpdateExtraRequestParentStatus(nte(rs("U_EXTRA_REQUEST_ID")))

44980     CancelExtraRequestData = True

44990     Exit Function
ERR_CancelExtraRequestData:
45000 MsgBox "ERR_CancelExtraRequestData" & vbCrLf & Err.Description
End Function


'cancel the extra request
'for the "Add Slide" action cancel the creared slides
'for the "Color Slide" action return the slides color to Reserve
'for the relevant aliquots return the aliquot station trace to the former state
Private Sub CancelExtraRequest(strExtraRequestId As String, strAction As String)
45010 On Error GoTo ERR_CancelExtraRequest
          Dim rs As Recordset
          Dim sql As String
          Dim strLogDesc As String
          Dim strEntityName As String
          
          'holds the data on which aliquots should get their
          'aliquot trace data reversed:
          Dim dicReverseTrace As New Dictionary
          
45020     strLogDesc = strAction & ":"
          
45030     sql = " select erd.U_EXTRA_REQUEST_DATA_ID, "
45040     sql = sql & "        erd.NAME, "
45050     sql = sql & "     erdu.U_REQUEST_DETAILS,"
45060     sql = sql & "     erdu.U_SLIDE_NAME,"
          'sql = sql & "     erdu.U_DESC,"
45070     sql = sql & "     eru.U_CREATED_BY,"
45080     sql = sql & "     eru.U_CREATED_ON "
45090     sql = sql & " from lims_sys.u_extra_request_user eru,"
45100     sql = sql & "      lims_sys.u_extra_request_data erd,"
45110     sql = sql & "      lims_sys.u_extra_request_data_user erdu"
45120     sql = sql & _
    " where erd.U_EXTRA_REQUEST_DATA_ID=erdu.U_EXTRA_REQUEST_DATA_ID"
45130     sql = sql & " and   eru.U_EXTRA_REQUEST_ID=erdu.U_EXTRA_REQUEST_ID"
45140     sql = sql & " and   erdu.U_EXTRA_REQUEST_ID='" & strExtraRequestId & _
    "'"
45150     sql = sql & " order by erd.U_EXTRA_REQUEST_DATA_ID"
          
45160     Set rs = connection.Execute(sql)
          
45170     If Not rs.EOF Then
45180         Call UpdateExtraRequestStatus(strExtraRequestId, "X")
45190     End If
          
45200     While Not rs.EOF
45210         strEntityName = CleanSemicolon(rs("NAME"))
          
45220         Call _
    UpdateExtraRequestDataStatus(nte(rs("U_EXTRA_REQUEST_DATA_ID")), "X")
                
45230         If strAction = "Add Slide" Then
45240             Call _
    UpdateAliquotStatus(GetAliquotId(nte(rs("U_SLIDE_NAME"))), "X")
                  'Call UpdateAliquotStatus(nte(rs("U_DESC")), "X")
45250         End If
              
45260         If strAction = "Color Slide" Then
45270             Call UpdateSlideColor(GetAliquotId(strEntityName), "רזרבה")
45280         End If
              
45290         strLogDesc = strLogDesc & " (" & nte(rs("NAME")) & " - " & _
    nte(rs("U_REQUEST_DETAILS")) & ") "
              
45300         If Not dicReverseTrace.Exists(strEntityName) Then
45310             Call dicReverseTrace.Add(strEntityName, strAction)
45320         End If
              
45330         rs.MoveNext
45340     Wend
          
45350     Call ComputeAliquotToReverseTrace(dicReverseTrace)
          
45360     Call sdg_log.InsertLog(rsSdg("sdg_id"), "EXTRA.CANCEL", strLogDesc)

45370     Exit Sub
ERR_CancelExtraRequest:
45380 MsgBox "ERR_CancelExtraRequest" & vbCrLf & Err.Description
End Sub


'decide which aliquots shoud have their stations reversed one step
'dicReverseTrace (in) - holds the u_extra_request_data.name (entity name)
'                       against the u_extra_request.name (action)
'according to the action we know:
'1. what was the change in the aliquot stations trace
'2. the relation between the entity name to the aliquots we should reverse
Private Sub ComputeAliquotToReverseTrace(dicReverseTrace As Dictionary)
45390 On Error GoTo ERR_ReverseAliquotTrace
          Dim i As Integer
          Dim strName As String
          Dim strAction As String

45400     For i = 0 To dicReverseTrace.Count - 1

45410         strName = CStr(dicReverseTrace.Keys(i))
45420         strAction = CStr(dicReverseTrace.Items(i))

45430         Select Case strAction
                  Case "Re Embedding", "Add Slide"
45440                 Call ReverseAliquotTrace(GetAliquotId(strName))
                      
45450             Case "Color Slide"
                      'Call ReverseAliquotTrace(GetBlockId(GetAliquotId(strName)))
              
45460             Case "Present Original Sample"
                      Dim rs As Recordset
                      Dim sql As String
                      
45470                 sql = " select a.ALIQUOT_ID"
45480                 sql = sql & " from lims_sys.aliquot a,"
45490                 sql = sql & "      lims_sys.sample s"
45500                 sql = sql & " where a.SAMPLE_ID=s.SAMPLE_ID "
45510                 sql = sql & " and s.name='" & strName & "'"
45520                 sql = sql & " and exists "
45530                 sql = sql & " ("
45540                 sql = sql & "    select 1"
45550                 sql = sql & "    from lims_sys.aliquot_formulation af"
45560                 sql = sql & _
    "    where af.parent_aliquot_id = a.aliquot_id"
45570                 sql = sql & " )"
                      
45580                 Set rs = connection.Execute(sql)
                      
45590                 While Not rs.EOF
45600                     Call ReverseAliquotTrace(nte(rs("ALIQUOT_ID")))
45610                     rs.MoveNext
45620                 Wend
45630         End Select
45640     Next i

45650     Exit Sub
ERR_ReverseAliquotTrace:
45660 MsgBox "ERR_ReverseAliquotTrace" & vbCrLf & Err.Description
End Sub

'change the aliquot_station, old_aliquot_station to be what they
'were one step back:
Private Sub ReverseAliquotTrace(strAliquotId As String)
45670 On Error GoTo ERR_ReverseAliquotTrace
          Dim rs As Recordset
          Dim sql As String
          Dim strStation As String
          Dim strOldStation As String

45680     sql = " select au.U_ALIQUOT_STATION, au.U_OLD_ALIQUOT_STATION"
45690     sql = sql & " from lims_sys.aliquot_user au"
45700     sql = sql & " where au.ALIQUOT_ID='" & strAliquotId & "'"

45710     Set rs = connection.Execute(sql)
45720     If rs.EOF Then Exit Sub
          
45730     strStation = nte(rs("U_ALIQUOT_STATION"))
45740     strOldStation = nte(rs("U_OLD_ALIQUOT_STATION"))
          
45750     If strStation = "" Then Exit Sub
          
45760     If strOldStation = "" Then
45770         strStation = ""
45780         strOldStation = ""
45790     Else
45800         strStation = Right(strOldStation, 1)
45810         strOldStation = Mid(strOldStation, 1, Len(strOldStation) - 1)
45820     End If
          
45830     sql = " update lims_sys.aliquot_user au"
45840     sql = sql & " set au.U_ALIQUOT_STATION='" & strStation & "',"
45850     sql = sql & "     au.U_OLD_ALIQUOT_STATION='" & strOldStation & "'"
45860     sql = sql & " where au.ALIQUOT_ID='" & strAliquotId & "'"
          
45870     Call connection.Execute(sql)
          
45880 Exit Sub
ERR_ReverseAliquotTrace:
45890 MsgBox "ERR_ReverseAliquotTrace" & vbCrLf & Err.Description
End Sub


Private Sub UpdateResultStatus(strResultId As String, strNewStatus As String)
45900 On Error GoTo ERR_UpdateResultStatus
          Dim sql As String
          
45910     sql = " update lims_sys.result r "
45920     sql = sql & " set r.status = '" & strNewStatus & "' "
45930     sql = sql & " where r.result_id = '" & strResultId & "' "
          
          
          
45940     Call connection.Execute(sql)

45950     Exit Sub
ERR_UpdateResultStatus:
45960 MsgBox "ERR_UpdateResultStatus" & vbCrLf & Err.Description
End Sub

Private Sub UpdateTestStatus(strTestId As String, strNewStatus As String)
45970 On Error GoTo ERR_UpdateTestStatus
          Dim sql As String
          
45980     sql = " update lims_sys.test t "
45990     sql = sql & " set t.status = '" & strNewStatus & "' "
46000     sql = sql & " where t.test_id = '" & strTestId & "' "
          
46010     Call connection.Execute(sql)

46020     Exit Sub
ERR_UpdateTestStatus:
46030 MsgBox "ERR_UpdateTestStatus" & vbCrLf & Err.Description
End Sub

Private Sub UpdateAliquotStatus(strAliquotId As String, strNewStatus As String)
46040 On Error GoTo ERR_UpdateAliquotStatus
          Dim sql As String
          
46050     sql = " update lims_sys.aliquot a "
46060     sql = sql & " set a.status = '" & strNewStatus & "' "
46070     sql = sql & " where a.aliquot_id = '" & strAliquotId & "' "
          
46080     Call connection.Execute(sql)

46090     Exit Sub
ERR_UpdateAliquotStatus:
46100 MsgBox "ERR_UpdateAliquotStatus" & vbCrLf & Err.Description
End Sub


'Private Sub CancelSlide(rsRequest As Recordset)
'On Error GoTo ERR_CancelSlide
'    Dim rs As Recordset
'    Dim sql As String
'
'    sql = " select a1.name, a1.aliquot_id, au.U_COLOR_TYPE "
'    sql = sql & " from lims_sys.aliquot a1,"
'    sql = sql & "      lims_sys.aliquot_user au"
'    sql = sql & " where a1.aliquot_id = au.ALIQUOT_ID"
'    sql = sql & " and   a1.created_on >= to_date('" & nte(rsRequest("u_created_on")) & "', 'dd/mm/yyyy hh24:mi:ss')"
'    sql = sql & " and   a1.created_by = '" & rsRequest("u_created_by") & "'"
'    sql = sql & " and   au.U_COLOR_TYPE = '" & rsRequest("U_REQUEST_DETAILS") & "'"
'    sql = sql & " and au.ALIQUOT_ID="
'    sql = sql & " ("
'    sql = sql & "   select af.CHILD_ALIQUOT_ID"
'    sql = sql & "   from lims_sys.aliquot_formulation af,"
'    sql = sql & "        lims_sys.aliquot a"
'    sql = sql & "   where af.CHILD_ALIQUOT_ID=au.ALIQUOT_ID"
'    sql = sql & "   and   af.PARENT_ALIQUOT_ID=a.ALIQUOT_ID"
'    sql = sql & "   and   a.name='" & CleanSemicolon(nte(rsRequest("name"))) & "'"
'    sql = sql & " )"
'    sql = sql & " and a1.status <> 'X'"
'    sql = sql & " order by a1.aliquot_id"
'
'    Set rs = connection.Execute(sql)
'
'    'delete 1st slide only: if there was a later extra request
'    'with the same color - it's slide should NOT be canceled
'    If Not rs.EOF Then
'        Call UpdateAliquotStatus(nte(rs("aliquot_id")), "X")
'
'        'MsgBox rs("name") & vbCrLf & rs("aliquot_id") & vbCrLf & rs("u_color_type")
'    End If
'
'    Exit Sub
'ERR_CancelSlide:
'MsgBox "ERR_CancelSlide" & vbCrLf & Err.Description
'End Sub



Private Sub UpdateExtraRequestStatus(strExtraRequestId As String, strNewStatus _
    As String)
46110 On Error GoTo ERR_UpdateExtraRequestStatus
          Dim sql As String
          
46120     If strExtraRequestId = "" Then Exit Sub

46130     sql = " update lims_sys.u_extra_request_user"
46140     sql = sql & " set u_status = '" & strNewStatus & "'"
46150     sql = sql & " where u_extra_request_id = '" & strExtraRequestId & "'"

46160     Call connection.Execute(sql)

46170     Exit Sub
ERR_UpdateExtraRequestStatus:
46180 MsgBox "ERR_UpdateExtraRequestStatus" & vbCrLf & Err.Description
End Sub

Private Sub UpdateExtraRequestDataStatus(strExtraRequestDataID As String, _
    strNewStatus As String)
46190 On Error GoTo ERR_UpdateExtraRequestStatus
          Dim sql As String
          
46200     If strExtraRequestDataID = "" Then Exit Sub

46210     sql = " update lims_sys.u_extra_request_data_user"
46220     sql = sql & " set u_status = '" & strNewStatus & "'"
46230     sql = sql & " where u_extra_request_data_id = '" & _
    strExtraRequestDataID & "'"

46240     Call connection.Execute(sql)

46250     Exit Sub
ERR_UpdateExtraRequestStatus:
46260 MsgBox "ERR_UpdateExtraRequestStatus" & vbCrLf & Err.Description
End Sub

'get the extra request id if could cancel this extra request, otherwise an empty string
'if the some of the items in the request passed status 'P', cancel is impossible
Private Function CanCancelExtraRequest(strExtraRequestDataID As String) As _
    String
46270 On Error GoTo ERR_CanCancelExtraRequest
          Dim rs As Recordset
          Dim sql As String
          Dim strExtraRequestId As String
          
46280     strExtraRequestId = GetExtraRequestId(strExtraRequestDataID)

46290     sql = " select 1"
46300     sql = sql & " from lims_sys.u_extra_request_data_user erdu"
46310     sql = sql & " where erdu.U_STATUS not in ('V','P')"
46320     sql = sql & " and erdu.U_EXTRA_REQUEST_ID='" & strExtraRequestId & "'"
          
46330     Set rs = connection.Execute(sql)

46340     If rs.EOF Then
46350         CanCancelExtraRequest = strExtraRequestId
46360     End If

46370     Exit Function
ERR_CanCancelExtraRequest:
46380 MsgBox "ERR_CanCancelExtraRequest" & vbCrLf & Err.Description
End Function


Private Function CanCancelExtraRequestData(strExtraRequestDataID As String) As _
    Boolean
46390 On Error GoTo ERR_CanCancelExtraRequestData
          
          Dim rs As Recordset
          Dim sql As String

46400     CanCancelExtraRequestData = False

46410     sql = " select 1"
46420     sql = sql & " from lims_sys.u_extra_request_data_user erdu"
46430     sql = sql & " where erdu.U_STATUS in ('V','P')"
46440     sql = sql & " and erdu.U_EXTRA_REQUEST_DATA_ID='" & _
    strExtraRequestDataID & "'"

46450     Set rs = connection.Execute(sql)
          
46460     If Not rs.EOF Then
46470         CanCancelExtraRequestData = True
46480     End If

46490     Exit Function
ERR_CanCancelExtraRequestData:
46500 MsgBox "ERR_CanCancelExtraRequestData" & vbCrLf & Err.Description
End Function




Private Function GetExtraRequestId(strExtraRequestDataID) As String
46510 On Error GoTo ERR_GetExtraRequestId
          Dim rs As Recordset
          Dim sql As String

46520     sql = " select erdu.U_EXTRA_REQUEST_ID"
46530     sql = sql & " from lims_sys.u_extra_request_data_user erdu"
46540     sql = sql & " where erdu.U_EXTRA_REQUEST_DATA_ID='" & _
    strExtraRequestDataID & "'"

46550     Set rs = connection.Execute(sql)
          
46560     If Not rs.EOF Then
46570         GetExtraRequestId = nte(rs("U_EXTRA_REQUEST_ID"))
46580     End If

46590     Exit Function
ERR_GetExtraRequestId:
46600 MsgBox "ERR_GetExtraRequestId" & vbCrLf & Err.Description
End Function




Private Sub SSTab1_DragDrop(Source As Control, X As Single, Y As Single)
46610 On Error GoTo ERR_SSTab1_DragDrop
      '    If Y < picBlocks.Top + 1000 Then Exit Sub
       '   If Y > picSlides.Top + picSlides.Height - 200 Then Exit Sub
          
46620     If Source.Tag = "cmdDrag" Then
              Dim move As Long
              Dim frmNewSlidesHeight As Double
              Dim picBlocksHeight As Double
              Dim VScrollBlocksHeight As Double
              Dim frmExistingSlidesTop As Double
              Dim frmExistingSlidesHeight As Double
              Dim picSlidesHeight As Double
              Dim VScrollSlidesHeight As Double
              Dim VScrollBlocksMax As Double
              Dim VScrollSlidesMax As Double
              
              'save old components values:
46630         frmNewSlidesHeight = frmNewSlides.Height
46640         picBlocksHeight = picBlocks.Height
46650         VScrollBlocksHeight = VScrollBlocks.Height
46660         frmExistingSlidesTop = frmExistingSlides.Top
46670         frmExistingSlidesHeight = frmExistingSlides.Height
46680         picSlidesHeight = picSlides.Height
46690         VScrollSlidesHeight = VScrollSlides.Height
46700         VScrollBlocksMax = VScrollBlocks.Max
46710         VScrollSlidesMax = VScrollSlides.Max
              
46720         move = Y - cmdDrag.Top
              
46730         If picSlides.Height - move < 20 Then Exit Sub
46740         If picBlocks.Height + move < 20 Then Exit Sub
                      
46750         frmNewSlides.Height = Y - frmNewSlides.Top
46760         picBlocks.Height = picBlocks.Height + move
46770         VScrollBlocks.Height = VScrollBlocks.Height + move
              
46780         frmExistingSlides.Top = frmExistingSlides.Top + move
46790         frmExistingSlides.Height = frmExistingSlides.Height - move
46800         picSlides.Height = picSlides.Height - move
46810         VScrollSlides.Height = VScrollSlides.Height - move
                      
46820         VScrollBlocks.Max = (picBlockEntry.Height - picBlocks.Height) / _
    ScaleHeight * 100
46830         If VScrollBlocks.Max < 0 Then
46840             VScrollBlocks.Visible = False
46850         Else
46860             VScrollBlocks.Visible = True
46870         End If
              
46880         VScrollSlides.Max = (picSlidesEntry.Height - picSlides.Height) / _
    ScaleHeight * 100
46890         If VScrollSlides.Max < 0 Then
46900             VScrollSlides.Visible = False
46910         Else
46920             VScrollSlides.Visible = True
46930         End If
              
46940         cmdDrag.Top = Y
46950     End If
          
46960     Exit Sub
ERR_SSTab1_DragDrop:
      'MsgBox "ERR_SSTab1_DragDrop" & vbCrLf & Err.Description
      'return to old components values:
46970 frmNewSlides.Height = frmNewSlidesHeight
46980 picBlocks.Height = picBlocksHeight
46990 VScrollBlocks.Height = VScrollBlocksHeight
47000 frmExistingSlides.Top = frmExistingSlidesTop
47010 frmExistingSlides.Height = frmExistingSlidesHeight
47020 picSlides.Height = picSlidesHeight
47030 VScrollSlides.Height = VScrollSlidesHeight
47040 VScrollBlocks.Max = VScrollBlocksMax
47050 VScrollSlides.Max = VScrollSlidesMax
End Sub





Private Sub txtAdvice_GotFocus()
'    Call zLang.Hebrew
End Sub




Private Sub tree_NodeClick(ByVal Node As MSComctlLib.Node)

47060     If Not isCito Then Exit Sub
          Dim temp As Integer, i As Integer
          Dim tmpname As String
          Dim sql As String
          Dim rst As Recordset
47070     cmdOKAddBlock.Enabled = False

47080     tmpname = Trim(getNextStr(tree.SelectedItem, " "))
47090     If Not InStr(tmpname, ".") > 1 Then
47100         selectedSampleName = "-1"
47110     Else
47120         selectedSampleName = getNextStr(tmpname, ".")
47130         selectedSampleName = selectedSampleName & "." & _
    getNextStr(tmpname & ".", ".")
              'check if sample came in a jar
              
47140         sql = " select  1"
47150         sql = sql & "  from "
47160         sql = sql & "    lims_sys.aliquot  a,"
47170         sql = sql & "    lims_sys.test t,"
47180         sql = sql & "    lims_sys.result r"
47190         sql = sql & "  where "
47200         sql = sql & "     a.name='" & selectedSampleName & "'||'.1'"
47210         sql = sql & "     and a.ALIQUOT_ID=t.ALIQUOT_ID"
47220         sql = sql & "     and t.TEST_ID=r.TEST_ID"
47230         sql = sql & "     and r.name='Arrived as Slides'"
47240         sql = sql & "     and r.ORIGINAL_RESULT ='F'"
47250         sql = sql & "  and not exists "
47260         sql = sql & "  (select  1"
47270         sql = sql & "   from "
47280         sql = sql & "         lims_sys.aliquot  acb,"
47290         sql = sql & "             lims_sys.aliquot_user  acbu"
47300         sql = sql & "   where acb.SAMPLE_ID=a.SAMPLE_ID"
47310         sql = sql & "         and acb.ALIQUOT_ID=acbu.ALIQUOT_ID"
47320         sql = sql & "         and acbu.U_IS_CELL_BLOCK='T'"
47330         sql = sql & "   )"
              
              
47340         Set rst = connection.Execute(sql)
47350         If Not rst.EOF Then
47360             cmdOKAddBlock.Enabled = True
47370             cmdOKAddBlock.Caption = "יצירת CELL - BLOCK"
47380         Else
47390             selectedSampleName = "-1"
47400         End If
              
                  
47410     End If
          
47420     temp = tree.SelectedItem.index
47430     For i = 1 To tree.Nodes.Count
47440        tree.Nodes.Item(i).Bold = False
47450     Next i
47460    tree.Nodes.Item(temp).Bold = True

End Sub

Private Sub txtAdvisorRemarks_DblClick(index As Integer)
47470     Call frmRemarks.Initialize(txtAdvisorRemarks(index).Text, False, _
    "הערות")
47480     Call frmRemarks.Show(vbModal)
47490     txtAdvisorRemarks(index).Text = frmRemarks.txt
End Sub

Private Sub txtAdvisorRemarks_GotFocus(index As Integer)
47500     Call zLang.Hebrew
End Sub

Private Sub txtReembeddingDetails_Change(index As Integer)
47510 On Error GoTo ERR_txtReembeddingDetails_Change
          
47520     BlockReembeddingList.d(BlockReembeddingList.iGuiGroup + index - _
    1).strDetails = txtReembeddingDetails(index).Text
47530     Exit Sub
ERR_txtReembeddingDetails_Change:
47540 MsgBox "ERR_txtReembeddingDetails_Change" & vbCrLf & Err.Description
End Sub

Private Sub txtReembeddingDetails_DblClick(index As Integer)
47550     Call frmRemarks.Initialize(txtReembeddingDetails(index).Text, False, _
    "הערות")
47560     Call frmRemarks.Show(vbModal)
47570     txtReembeddingDetails(index).Text = frmRemarks.txt
End Sub

Private Sub txtReembeddingDetails_GotFocus(index As Integer)
47580     Call zLang.Hebrew
End Sub

'Private Sub txtSlidesFromArchive_Change()
'    If lstSlidesFromArchive.ListCount > 0 Then Exit Sub
'
'    If txtSlidesFromArchive.Text <> "" Then
'        cmdOKSlidesFromArchive.Enabled = True
'    Else
'        cmdOKSlidesFromArchive.Enabled = False
'    End If
'End Sub

Private Sub txtSlidesFromArchive_GotFocus()
47590     Call zLang.Hebrew
End Sub

Private Sub VScrollAdvisor_Change()
47600     picAdvisorEntry.Top = -(VScrollAdvisor.value / 100) * ScaleHeight + 50
End Sub

Private Sub VScrollBlocks_Change()
47610 On Error GoTo ERR_VScrollBlocks_Change

47620      BlockAddSlideList.iGuiGroup = BlockAddSlideList.d.Count * _
    VScrollBlocks.value / VScrollBlocks.Max - BlockAddSlideList.MAX_LINES / 2
47630      Call RefreshBlockAddSlide
      '    picBlockEntry.Top = -(VScrollBlocks.Value / 100) * ScaleHeight + 50
          
47640     Exit Sub
ERR_VScrollBlocks_Change:
47650 MsgBox "ERR_VScrollBlocks_Change" & vbCrLf & Err.Description
End Sub

Private Sub RefreshBlockAddSlide()
47660 On Error GoTo ERR_RefreshBlockAddSlide

          Dim i As Integer
          Dim k As Integer
          Dim iBegin As Integer
          Dim iEnd As Integer
          Dim iDelta As Integer
          Dim d As Dictionary
          
47670     iBegin = BlockAddSlideList.iGuiGroup
47680     iEnd = iBegin + BlockAddSlideList.MAX_LINES - 1
47690     iDelta = iEnd - BlockAddSlideList.d.Count
          
47700     If iDelta >= 0 Then
47710         iEnd = iEnd - iDelta
47720         iBegin = iBegin - iDelta
47730         BlockAddSlideList.iGuiGroup = BlockAddSlideList.iGuiGroup - iDelta
      '        VScrollBlocks.Value = VScrollBlocks.Max
47740     Else
47750         If iBegin <= 1 Then
47760             BlockAddSlideList.iGuiGroup = 1
47770             iBegin = 1
47780             iEnd = iBegin + BlockAddSlideList.MAX_LINES - 1
      '            VScrollBlocks.Value = 0
47790         End If
47800     End If
          
          'update the gui data for these rows:
47810     For i = iBegin To iEnd
47820         lblBlock(i - BlockAddSlideList.iGuiGroup + 1).Caption = _
    BlockAddSlideList.d(i).strSample
47830         lblBlock(i - BlockAddSlideList.iGuiGroup + 1).Visible = _
    BlockAddSlideList.d(i).strSample <> ""
47840         cmdBlock(i - BlockAddSlideList.iGuiGroup + 1).Caption = _
    BlockAddSlideList.d(i).strName
47850         chkBlockMicrotom(i - BlockAddSlideList.iGuiGroup + 1).value = _
    BlockAddSlideList.d(i).iMicrotom
47860         With CmbBlockEntry(i - BlockAddSlideList.iGuiGroup + 1)
47870             Set d = BlockAddSlideList.d(i).dicColors
47880             .Clear
47890             For k = 0 To d.Count - 1
47900                 .AddItem CStr(d.Items(k))
47910             Next k
47920             If d.Count > 0 Then
47930                 .BackColor = MARK_SELECTED
47940             Else
47950                 .BackColor = vbWhite
47960             End If
47970         End With
              
              'chck(i - list.nGuiGroup + 1).Caption = list.dx(i).strName 'i
47980     Next i

47990     Exit Sub
ERR_RefreshBlockAddSlide:
48000 MsgBox "ERR_RefreshBlockAddSlide" & vbCrLf & Err.Description
End Sub

'get the lists of colors
Private Sub InitColors()
48010 On Error GoTo ERR_InitColors
          Dim phrase As Recordset
          Dim sql As String
      '
      '    Set phrase = connection.Execute("select phrase_name from lims_sys.phrase_entry " & '        "where phrase_id = (select phrase_id from lims_sys.phrase_header where " & '        "name = 'Pathology Molecular Stains') " & '        "order by order_number")

48020     dicMolecularStains.RemoveAll
      '    While Not phrase.EOF
      '        Call dicMolecularStains.Add(CStr(phrase("PHRASE_NAME").Value), dicMolecularStains.Count)
      '        phrase.MoveNext
      '    Wend
48030      sql = " select U_stain,u_order,u_part_code "
48040     sql = sql & " from lims_sys.u_parts_user "
48050     sql = sql & " where u_part_type='H' "
48060     sql = sql & " and U_stain is not null "
48070     sql = sql & "order by u_order "
          
48080     Set phrase = connection.Execute(sql)

48090     dicHistochemistryStains.RemoveAll
48100     While Not phrase.EOF
48110         Call dicHistochemistryStains.Add(CStr(phrase("U_STAIN").value), _
    dicHistochemistryStains.Count)
48120         phrase.MoveNext
48130     Wend
          
         
48140     sql = " select U_stain,u_order,u_part_code "
48150     sql = sql & " from lims_sys.u_parts_user "
48160     sql = sql & "where u_part_type='I' "
48170     sql = sql & "order by u_order "

48180     Set phrase = connection.Execute(sql)

48190     dicImonohistochemistryStains.RemoveAll
48200     While Not phrase.EOF
48210         Call _
    dicImonohistochemistryStains.Add(CStr(phrase("U_STAIN").value), _
    dicImonohistochemistryStains.Count)
48220         phrase.MoveNext
48230     Wend
          
48240     sql = " select U_stain,u_order,u_part_code "
48250     sql = sql & " from lims_sys.u_parts_user "
48260     sql = sql & "where u_part_type='O' "
48270     sql = sql & "order by u_order "

48280     Set phrase = connection.Execute(sql)

48290     dicOtherStainOptions.RemoveAll
48300     While Not phrase.EOF
48310         Call dicOtherStainOptions.Add(CStr(phrase("U_STAIN").value), _
    dicImonohistochemistryStains.Count)
48320         phrase.MoveNext
48330     Wend
          
          
48340     Exit Sub
ERR_InitColors:
48350 MsgBox "ERR_InitColors on line #" & Erl & vbCrLf & Err.Description
End Sub

Private Sub VScrollBlocksReembedding_Change()
48360 On Error GoTo ERR_VScrollBlocksReembedding_Change
           
48370      BlockReembeddingList.iGuiGroup = BlockReembeddingList.d.Count * _
    VScrollBlocksReembedding.value / VScrollBlocksReembedding.Max - _
    BlockReembeddingList.MAX_LINES / 2
48380      Call RefreshBlockReembedding
      '    picBlockReembeddingEntry.Top = -(VScrollBlocksReembedding.Value / 100) * ScaleHeight + 50

48390     Exit Sub
ERR_VScrollBlocksReembedding_Change:
48400 MsgBox "ERR_VScrollBlocksReembedding_Change" & vbCrLf & Err.Description
End Sub


Private Sub RefreshBlockReembedding()
48410 On Error GoTo ERR_RefreshBlockReembedding

          Dim i As Integer
          Dim k As Integer
          Dim iBegin As Integer
          Dim iEnd As Integer
          Dim iDelta As Integer
          Dim d As Dictionary
          
48420     iBegin = BlockReembeddingList.iGuiGroup
48430     iEnd = iBegin + BlockReembeddingList.MAX_LINES - 1
48440     iDelta = iEnd - BlockReembeddingList.d.Count
          
48450     If iDelta >= 0 Then
48460         iEnd = iEnd - iDelta
48470         iBegin = iBegin - iDelta
48480         BlockReembeddingList.iGuiGroup = BlockReembeddingList.iGuiGroup - _
    iDelta
      '        VScrollBlocksReembedding.Value = VScrollBlocksReembedding.Max
48490     Else
48500         If iBegin <= 1 Then
48510             BlockReembeddingList.iGuiGroup = 1
48520             iBegin = 1
48530             iEnd = iBegin + BlockReembeddingList.MAX_LINES - 1
      '            VScrollBlocksReembedding.Value = 0
48540         End If
48550     End If
          
          'update the gui data for these rows:
48560     For i = iBegin To iEnd
48570         lblSampleReembedding(i - BlockReembeddingList.iGuiGroup + _
    1).Caption = BlockReembeddingList.d(i).strSample
48580         lblSampleReembedding(i - BlockReembeddingList.iGuiGroup + _
    1).Visible = BlockReembeddingList.d(i).strSample <> ""
48590         chkBlockReembeding(i - BlockReembeddingList.iGuiGroup + 1).value _
    = BlockReembeddingList.d(i).iReembedding
48600         chkBlockReembeding(i - BlockReembeddingList.iGuiGroup + _
    1).Caption = BlockReembeddingList.d(i).strName
48610         cmbReembeddingReason(i - BlockReembeddingList.iGuiGroup + 1).Text _
    = BlockReembeddingList.d(i).strReason
48620         txtReembeddingDetails(i - BlockReembeddingList.iGuiGroup + _
    1).Text = BlockReembeddingList.d(i).strDetails
                                                  
48630     Next i

48640 Exit Sub
ERR_RefreshBlockReembedding:
48650 MsgBox "ERR_RefreshBlockReembedding" & vbCrLf & Err.Description
End Sub


'if exists - return the index of this item in the list
'otherwise - return -1
Private Function ExistInList(strName As String, list As ListBox) As Integer
          Dim i As Integer
          
48660     ExistInList = -1
          
48670     For i = 0 To list.ListCount - 1
48680         If list.list(i) = strName Then
48690             ExistInList = i
48700             Exit For
48710         End If
48720     Next i
End Function

Private Sub VScrollShowSamples_Change()
48730     picShowSamplesEntry.Top = -(VScrollShowSamples.value / 100) * _
    ScaleHeight + 50
End Sub

Private Sub VScrollSlides_Change()
48740     picSlidesEntry.Top = -(VScrollSlides.value / 100) * ScaleHeight + 50
End Sub

Public Sub IncrementColorsBlocks()
48750     iUpdateColorsBlocks = iUpdateColorsBlocks + 1
48760     cmdOKColors.Enabled = True
      '    MsgBox iUpdateColorsBlocks
End Sub

Public Sub DecrementColorsBlocks()
48770     If iUpdateColorsBlocks = 0 Then Exit Sub
          
48780     iUpdateColorsBlocks = iUpdateColorsBlocks - 1
48790     If iUpdateColorsBlocks = 0 And iUpdateColorsSlides = 0 Then
48800         cmdOKColors.Enabled = False
48810     End If
      '    MsgBox iUpdateColorsBlocks
End Sub

Public Sub IncrementColorsSlides()
48820     iUpdateColorsSlides = iUpdateColorsSlides + 1
48830     cmdOKColors.Enabled = True
      '    MsgBox iUpdateColorsSlides
End Sub

Public Sub DecrementColorsSlides()
48840     If iUpdateColorsSlides = 0 Then Exit Sub
          
48850     iUpdateColorsSlides = iUpdateColorsSlides - 1
48860     If iUpdateColorsSlides = 0 And iUpdateColorsBlocks = 0 Then
48870         cmdOKColors.Enabled = False
48880     End If
      '    MsgBox iUpdateColorsSlides
End Sub

Private Sub ResetColorTab()
48890 On Error GoTo ERR_ResetColorTab
          Dim i As Integer
          Dim bas As BlockAddSlide
       
48900     cmdOKColors.Enabled = False
          
48910     For i = 1 To cmdBlock.Count - 1
48920         CmbBlockEntry(i).BackColor = vbWhite
48930         CmbBlockEntry(i).Clear
48940         chkBlockMicrotom(i).value = 0
              'chkBlockSerial(i).Value = 0
48950     Next i
          
          'reset memory list:
48960     For i = 1 To BlockAddSlideList.d.Count
48970         Set bas = BlockAddSlideList.d(i)
48980         Call bas.dicColors.RemoveAll
48990         bas.iMicrotom = 0
49000     Next i
          
49010     For i = 0 To cmdSlide.Count - 1
49020         txtSlide(i).Text = ""
49030     Next i
          
49040     txtColors.Text = ""
          
49050     iUpdateColorsBlocks = 0
49060     iUpdateColorsSlides = 0
          
49070     Exit Sub
ERR_ResetColorTab:
49080 MsgBox "ERR_ResetColorTab" & vbCrLf & Err.Description
End Sub

Private Function GetEntityType(strEntityName As String) As String
          Dim i As Integer
          
49090     i = GetNumOfDots(strEntityName)
          
49100     Select Case i
              Case 0
49110             GetEntityType = "SDG"
49120         Case 1
49130             GetEntityType = "Sample"
49140         Case 2
49150             GetEntityType = "Block"
49160         Case 3
49170             GetEntityType = "Slide"
49180     End Select
End Function


Private Function GetNumOfDots(str As String) As Integer
          Dim MainStr As String
          Dim s As String
          Dim i As Integer
          
49190     i = -1
49200     MainStr = str
49210     While MainStr <> ""
49220         s = getNextStr(MainStr, ".")
49230         i = i + 1
49240     Wend
          
49250     GetNumOfDots = i
End Function

Private Function getNextStr(ByRef s As String, c As String)
          Dim p
          Dim res
49260     p = InStr(1, s, c)
49270     If (p = 0) Then
49280         res = s
49290         s = ""
49300         getNextStr = res
49310     Else
49320         res = Mid$(s, 1, p - 1)
49330         s = Mid$(s, p + Len(c), Len(s))
49340         getNextStr = res
49350     End If
End Function

Private Function GetOperatorRole(strOperatorId As String) As String
49360 On Error GoTo ERR_GetOperatorRole
          Dim sql As String
          Dim rs As Recordset
          
49370     sql = "  select r.NAME"
49380     sql = sql & "  from lims_sys.operator o,lims_sys.lims_role r "
49390     sql = sql & "  where o.OPERATOR_ID=" & strOperatorId
49400     sql = sql & "  and o.ROLE_ID=r.ROLE_ID"
49410     Set rs = connection.Execute(sql)
          
49420     GetOperatorRole = nte(rs("NAME"))

49430     Exit Function
ERR_GetOperatorRole:
49440 MsgBox "ERR_GetOperatorRole" & vbCrLf & Err.Description
End Function


Private Function GetOperatorName(strOperatorId As String) As String
49450 On Error GoTo ERR_GetOperatorName
          Dim sql As String
          Dim rs As Recordset
          
49460     sql = " select o.NAME "
49470     sql = sql & " from lims_sys.operator o,lims_sys.operator_user ou "
49480     sql = sql & " where o.OPERATOR_ID=" & strOperatorId
49490     sql = sql & " and o.OPERATOR_ID=ou.OPERATOR_ID"
49500     Set rs = connection.Execute(sql)
          
49510     GetOperatorName = nte(rs("NAME"))

49520     Exit Function
ERR_GetOperatorName:
49530 MsgBox "ERR_GetOperatorName" & vbCrLf & Err.Description
End Function

Private Sub InsertFieldValue(TextToReplace As String, ReplacingText As String, _
    wdoc As Word.Document, wapp As Word.Application)
      '    Dim docApp As Object 'Word.Application
          Dim mstErrMsg As String
          Dim mloErrNbr As Long

          Dim lstReplacingText As String
          Dim lstReplacingChunk As String
49540     On Error GoTo Fin
          
49550     lstReplacingText = Trim(ReplacingText)
          
49560     While Len(lstReplacingText) > 200
49570         lstReplacingChunk = Mid(lstReplacingText, 1, 200) & TextToReplace
49580         Call InsertFieldChunk(TextToReplace, lstReplacingChunk, wdoc, _
    wapp)
49590         lstReplacingText = Mid(lstReplacingText, 201)
49600     Wend
49610     Call InsertFieldChunk(TextToReplace, lstReplacingText, wdoc, wapp)
          
49620 Exit Sub
Fin:
49630     mstErrMsg = mstErrMsg & vbCrLf & "SUB: InsertFieldValue" & vbCrLf & _
    "Error Number:" & Err.Number & vbCrLf & "Description" & Err.Description & _
    vbCrLf & "TextToReplace = " & TextToReplace & vbCrLf & "ReplacingText = " & _
    ReplacingText & vbCrLf & "Replacing Chunk = " & lstReplacingChunk
49640     mloErrNbr = Err.Number
49650     Err.Raise Err.Number
      '    Select Case Err.Number
      '    Case 462   ' Word Closed
      '    Case Else
      '        MsgBox mstErrMsg, vbCritical + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, "Class: Report"
      '    End Select
End Sub

Private Sub InsertFieldChunk(TextToReplace As String, ReplacingText As String, _
    wdoc As Word.Document, wapp As Word.Application)

          Dim i As Integer
          Dim mstErrMsg As String
          Dim mloErrNbr As Long

49660     On Error GoTo Fin
      '    Set docApp = Doc.Application
49670     For i = 1 To 3
49680         Select Case i
              Case 1
49690             wdoc.ActiveWindow.ActivePane.View.SeekView = _
    wdSeekCurrentPageHeader
49700         Case 2
49710             wdoc.ActiveWindow.ActivePane.View.SeekView = _
    wdSeekCurrentPageFooter
49720         Case 3
49730             wdoc.ActiveWindow.ActivePane.View.SeekView = _
    wdSeekMainDocument
49740         End Select
49750         wapp.Selection.Find.ClearFormatting
49760         wapp.Selection.Find.Replacement.ClearFormatting
49770         With wapp.Selection.Find
49780             .Text = TextToReplace
49790             .Replacement.Text = ReplacingText
49800             .Forward = True
49810             .Wrap = wdFindContinue
49820             .Format = False
49830             .MatchCase = False
49840             .MatchWholeWord = False
49850             .MatchKashida = False
49860             .MatchDiacritics = False
49870             .MatchAlefHamza = False
49880             .MatchControl = False
49890             .MatchWildcards = False
49900             .MatchSoundsLike = False
49910             .MatchAllWordForms = False
49920         End With
49930         wapp.Selection.Find.Execute Replace:=wdReplaceAll
49940     Next i
49950     Exit Sub
Fin:
49960     mstErrMsg = mstErrMsg & vbCrLf & "SUB: InsertFieldChunk" & vbCrLf & _
    "Error Number:" & Err.Number & vbCrLf & "Description" & Err.Description & _
    vbCrLf & "TextToReplace = " & TextToReplace & vbCrLf & "ReplacingText = " & _
    ReplacingText
49970     mloErrNbr = Err.Number
49980     Err.Raise Err.Number

End Sub


Private Sub InitExtraRequestsHistory()
49990 On Error GoTo ERR_InitExtraRequestsHistory
          Dim sql As String
          Dim rs As Recordset
          Dim iRows As Integer
          Dim i As Integer

50000     Call InitializeGrid

50010     sql = _
    "  select rd.U_EXTRA_REQUEST_DATA_ID ID, rd.NAME ENTITY_NAME, r.NAME ACTION, "
50020     sql = sql & "         rdu.U_REQUEST_DETAILS,"
50030     sql = sql & "         pe.PHRASE_DESCRIPTION, "
50040     sql = sql & "         o.NAME, ru.U_CREATED_ON, r.DESCRIPTION   "
50050     sql = sql & "  from lims_sys.u_extra_request_data rd, "
50060     sql = sql & "          lims_sys.u_extra_request_data_user rdu, "
50070     sql = sql & "          lims_sys.u_extra_request r,"
50080     sql = sql & "       lims_sys.u_extra_request_user ru,"
50090     sql = sql & "       lims_sys.operator o, lims_sys.sdg d, "
50100     sql = sql & _
    "       lims_sys.phrase_header ph, lims_sys.phrase_entry pe  "
50110     sql = sql & _
    "  where rd.U_EXTRA_REQUEST_DATA_ID=rdu.U_EXTRA_REQUEST_DATA_ID"
50120     sql = sql & "  and r.U_EXTRA_REQUEST_ID=rdu.U_EXTRA_REQUEST_ID"
50130     sql = sql & "  and r.U_EXTRA_REQUEST_ID=ru.U_EXTRA_REQUEST_ID"
50140     sql = sql & "  and o.OPERATOR_ID=ru.U_CREATED_BY"
50150     sql = sql & "  and ru.U_SDG_ID=d.sdg_id "
          
          'change this to name in patholab
        '  sql = sql & "  and d.external_reference='" & nte(rsSdg("external_reference")) & "' "
50160     sql = sql & "  and substr(d.name,1,10) ='" & nte(rsSdg("name")) & "' "
          
          
      '    sql = sql & "  and ru.U_SDG_ID=" & rsSdg("sdg_id")
          'sql = sql & "  and ru.U_EXTERNAL_REFERENCE='" & nte(rsSdg("external_reference")) & "' "
50170     sql = sql & "  and pe.PHRASE_ID=ph.PHRASE_ID"
50180     sql = sql & "  and ph.NAME='Extra Request Status'"
50190     sql = sql & "  and pe.PHRASE_NAME=rdu.U_STATUS"
50200     sql = sql & "  and rdu.U_STATUS <> 'X'"
50210     sql = sql & "  order by rd.U_EXTRA_REQUEST_DATA_ID"

50220     Set rs = connection.Execute(sql)

50230     iRows = 1
          
50240     While Not rs.EOF
50250         iRows = iRows + 1
          
50260         grid.Rows = iRows
50270         grid.col = 0
50280         grid.row = grid.Rows - 1
              
50290         For i = 0 To rs.Fields.Count - 1
50300             grid.col = i
50310             grid.CellAlignment = vbLeftJustify
50320             grid.Text = CleanSemicolon(nte(rs.Fields(i).value))
                  
                  
                  
                              
50330             If rs.Fields(i).name = "PHRASE_DESCRIPTION" Then
50340                 grid.Text = grid.Text & AddSlideInTrayText(rs)
50350             End If
                  
      '            If rs.Fields(i).name = "DESCRIPTION" Then
      '                If nte(rs.Fields(i).Value) <> "" Then
      '
      '                End If
      '            End If
50360         Next i
                         
      '        If ExistRemark(grid.TextMatrix(grid.row, 0)) Then
      '            grid.col = 0
      '            grid.CellFontBold = True
      '        End If
                         
50370         rs.MoveNext
50380     Wend

50390     Exit Sub
ERR_InitExtraRequestsHistory:
50400 MsgBox "ERR_InitExtraRequestsHistory" & vbCrLf & Err.Description
End Sub


Private Function AddSlideInTrayText(rs As Recordset) As String
50410 On Error GoTo ERR_AddSlideInTrayText
          Dim strAction As String
          Dim strSlideId As String
          Dim sql As String
          Dim rsStation As Recordset
          Dim rsSlide As Recordset
          
50420     strAction = CleanSemicolon(rs("ACTION"))

50430     Select Case strAction
              Case "Add Slide"
                  
50440             sql = " select rdu.U_SLIDE_NAME "
                  'sql = " select rdu.U_DESC"
50450             sql = sql & " from lims_sys.u_extra_request_data_user rdu"
50460             sql = sql & " where rdu.U_EXTRA_REQUEST_DATA_ID='" & rs("ID") _
    & "'"
50470             Set rsSlide = connection.Execute(sql)
50480             If Not rsSlide.EOF Then
50490                 strSlideId = GetAliquotId(nte(rsSlide("U_SLIDE_NAME")))
                      'strSlideId = nte(rsSlide("U_DESC"))
50500             End If

50510         Case "Color Slide"
                           
50520             strSlideId = _
    GetAliquotId(CleanSemicolon(nte(rs("ENTITY_NAME"))))
                           
50530         Case Else
50540             Exit Function
50550     End Select

50560     If strSlideId = "" Then Exit Function
          
50570     sql = " select au.U_ALIQUOT_STATION, au.U_OLD_ALIQUOT_STATION"
50580     sql = sql & " from lims_sys.aliquot_user au"
50590     sql = sql & " where au.ALIQUOT_ID='" & strSlideId & "'"
50600     Set rsStation = connection.Execute(sql)

50610     If Not rsStation.EOF Then
50620         If InStr(1, nte(rsStation("U_ALIQUOT_STATION")), "6") > 0 Or _
    InStr(1, nte(rsStation("U_OLD_ALIQUOT_STATION")), "6") > 0 Then
50630               AddSlideInTrayText = " (T)"
50640         End If
50650     End If

50660     Exit Function
ERR_AddSlideInTrayText:
50670 MsgBox "ERR_AddSlideInTrayText" & vbCrLf & Err.Description
End Function

Private Function ExistRemark(strExtraRequestDataID As String) As Boolean
50680 On Error GoTo ERR_ExistRemark
          Dim rs As Recordset
          Dim sql As String
          
50690     sql = " select r.DESCRIPTION "
50700     sql = sql & " from lims_sys.u_extra_request_data rd, "
50710     sql = sql & "      lims_sys.u_extra_request_data_user rdu, "
50720     sql = sql & "      lims_sys.u_extra_request r"
50730     sql = sql & _
    "  where rd.U_EXTRA_REQUEST_DATA_ID=rdu.U_EXTRA_REQUEST_DATA_ID"
50740     sql = sql & "  and   r.U_EXTRA_REQUEST_ID=rdu.U_EXTRA_REQUEST_ID"
50750     sql = sql & "  and   rd.U_EXTRA_REQUEST_DATA_ID=" & _
    strExtraRequestDataID

50760     Set rs = connection.Execute(sql)
          
50770     If nte(rs("DESCRIPTION")) = "" Then
50780         ExistRemark = False
50790     Else
50800         ExistRemark = True
50810     End If

50820     Exit Function
ERR_ExistRemark:
50830 MsgBox "ERR_ExistRemark" & vbCrLf & Err.Description
End Function


'Private Sub InitExtraRequestsHistory()
'On Error GoTo ERR_InitExtraRequestsHistory
'    Dim sql As String
'
'    Call InitializeGrid
'
'    sql = "   select r.U_EXTRA_REQUEST_ID, r.NAME ACTION,"
'    sql = sql & "          o.NAME CREATED_BY, ru.U_CREATED_ON,"
'    sql = sql & "       r.DESCRIPTION   "
'    sql = sql & "  from    lims_sys.u_extra_request r,"
'    sql = sql & "       lims_sys.u_extra_request_user ru,"
'    sql = sql & "       lims_sys.operator o "
'    sql = sql & "  where r.U_EXTRA_REQUEST_ID=ru.U_EXTRA_REQUEST_ID"
'    sql = sql & "  and   o.OPERATOR_ID=ru.U_CREATED_BY"
'    sql = sql & "  and   ru.U_SDG_ID = " & rsSdg("sdg_id")
'    sql = sql & "  order by ru.U_CREATED_ON"
'
'    Set rsExtraRequests = connection.Execute(sql)
'
'    If Not rsExtraRequests.EOF Then
' '       Call SetExtraRequestInfo
' '       txtRequestNumber.Text = 1
' '       txtTotalExtraRequests.Text = rsExtraRequests.RecordCount
'        If rsExtraRequests.RecordCount = 1 Then
'            cmdNext.Enabled = False
'            cmdFirst.Enabled = False
'            cmdLast.Enabled = False
'        End If
'    Else
'        cmdNext.Enabled = False
'        cmdFirst.Enabled = False
'        cmdLast.Enabled = False
'    End If
'
'
'    cmdPrev.Enabled = False
'
'    Exit Sub
'ERR_InitExtraRequestsHistory:
'MsgBox "ERR_InitExtraRequestsHistory" & vbCrLf & Err.description
'End Sub

Private Sub InitRequestDatils(strRequestId As String)
50840 On Error GoTo ERR_InitRequestDatils
          Dim rs As Recordset
          Dim sql As String
          Dim iRows As Integer
          Dim i As Integer
          
50850     sql = "   select rdu.U_ENTITY_TYPE, rd.NAME ENTITY_NAME,"
50860     sql = sql & "         rdu.U_REQUEST_DETAILS"
50870     sql = sql & "  from lims_sys.u_extra_request_data rd, "
50880     sql = sql & "       lims_sys.u_extra_request_data_user rdu  "
50890     sql = sql & _
    "  where rd.U_EXTRA_REQUEST_DATA_ID=rdu.U_EXTRA_REQUEST_DATA_ID"
50900     sql = sql & "  and rdu.U_EXTRA_REQUEST_ID=" & strRequestId
50910     sql = sql & "  order by rd.U_EXTRA_REQUEST_DATA_ID"

50920     Set rs = connection.Execute(sql)

50930     iRows = 1
          
50940     While Not rs.EOF
50950         iRows = iRows + 1
          
50960         grid.Rows = iRows
50970         grid.col = 0
50980         grid.row = grid.Rows - 1
              
50990         For i = 0 To rs.Fields.Count - 1
51000             grid.col = i
51010             grid.CellAlignment = vbLeftJustify
51020             grid.Text = CleanSemicolon(nte(rs.Fields(i).value))
51030         Next i
                         
51040         rs.MoveNext
51050     Wend
          

51060     Exit Sub
ERR_InitRequestDatils:
51070 MsgBox "ERR_InitRequestDatils" & vbCrLf & Err.Description
End Sub

Private Sub InitExtraRequestsGrid()
51080 On Error GoTo ERR_InitGrid

           
51090     Exit Sub
ERR_InitGrid:
51100 MsgBox "ERR_InitGrid" & vbCrLf & Err.Description
End Sub


'Private Function SetExtraRequestInfo()
'    txtRequest.Text = CleanSemicolon(nte(rsExtraRequests("ACTION")))
'    txtCreatedBy.Text = CleanSemicolon(nte(rsExtraRequests("CREATED_BY")))
'    txtCreaedOn.Text = CleanSemicolon(nte(rsExtraRequests("U_CREATED_ON")))
'    txtRequestRemarks.Text = CleanSemicolon(nte(rsExtraRequests("DESCRIPTION")))
'
'    Call InitRequestDatils(rsExtraRequests("U_EXTRA_REQUEST_ID"))
'End Function

Private Function CleanSemicolon(str As String) As String
          Dim i As Integer
          
51110     CleanSemicolon = str
          
51120     i = InStr(1, str, ";")
51130     If i = 0 Then Exit Function
          
51140     CleanSemicolon = Left(str, i - 1)
End Function

Private Sub InitializeGrid()
51150 On Error GoTo ERR_InitializeGrid

          Dim X As Integer
          Dim s As String
              
51160     grid.Clear
51170     grid.RightToLeft = False
              
51180     grid.AllowBigSelection = False
51190     grid.AllowUserResizing = flexResizeNone
51200     grid.Enabled = True
          
51210     grid.ScrollBars = flexScrollBarBoth
51220     grid.SelectionMode = flexSelectionFree
51230     grid.AllowUserResizing = flexResizeBoth

51240     grid.Rows = 2
51250     grid.Cols = 8
51260     grid.FixedRows = 1
51270     grid.FixedCols = 1

51280     grid.row = 0
51290     grid.RowHeight(X) = 400
      '    For X = 1 To grid.Rows - 1
      '        grid.row = X
      '        grid.RowHeight(X) = 600
      '    Next X

51300     grid.ColWidth(0) = 650
51310     For X = 1 To grid.Cols - 1
51320         grid.col = X
51330         grid.ColWidth(X) = 1500
51340     Next X
          
          'set the text for the COLUMN HEADERS:
51350     grid.row = 0
51360     grid.col = 0
      '    grid.CellAlignment = vbLeftJustify
      '    grid.Text = "Entity Type"

      '    grid.Col = grid.Col + 1
51370     grid.CellAlignment = vbLeftJustify
      '    grid.Text = "Entity Number"

51380     grid.col = grid.col + 1
51390     grid.CellAlignment = vbLeftJustify
51400     grid.Text = "Entity Name"
          
51410     grid.col = grid.col + 1
51420     grid.CellAlignment = vbLeftJustify
51430     grid.Text = "Action"
          
51440     grid.col = grid.col + 1
51450     grid.CellAlignment = vbLeftJustify
51460     grid.Text = "Details"
          
51470     grid.col = grid.col + 1
51480     grid.CellAlignment = vbLeftJustify
51490     grid.Text = "Status"
          
51500     grid.col = grid.col + 1
51510     grid.CellAlignment = vbLeftJustify
51520     grid.Text = "Created By"
          
51530     grid.col = grid.col + 1
51540     grid.CellAlignment = vbLeftJustify
51550     grid.Text = "Created On"
          
51560     grid.col = grid.col + 1
51570     grid.CellAlignment = vbLeftJustify
51580     grid.Text = "Remarks"

51590     Exit Sub
ERR_InitializeGrid:
51600 MsgBox "ERR_InitializeGrid" & vbCrLf & Err.Description
End Sub

Private Sub InitPhrases()
51610 On Error GoTo ERR_InitPhrases
          Dim rs As Recordset
          
          'Init the Patholog combo
51620     Set rs = _
    connection.Execute("select phrase_description, phrase_name from lims_sys.phrase_entry " _
    & "where phrase_id = (select phrase_id from lims_sys.phrase_header where " & _
    "name = 'Embedding') " & " and phrase_name <> 'embedding 1' " & _
    "order by order_number")

51630     Do Until rs.EOF
51640         cmbReembeddingReason(0).AddItem (rs("phrase_name"))

51650         rs.MoveNext
51660     Loop
          
51670     Exit Sub
ERR_InitPhrases:
51680 MsgBox "ERR_InitPhrases" & vbCrLf & Err.Description
End Sub


'get a list of all slides for this Block
'strBlockId (in) -
'dicSlide (out) - the list of slides
Private Sub GetSlidesForBlock(strBlockId As String, dicSlides As Dictionary)
51690 On Error GoTo ERR_GetSlidesForBlock
          Dim rs As Recordset
          Dim sql As String

51700     sql = " select a.NAME"
51710     sql = sql & " from lims_sys.aliquot a"
51720     sql = sql & " where exists"
51730     sql = sql & " ("
51740     sql = sql & "    select aliquot_id"
51750     sql = sql & "    from lims_sys.aliquot_formulation"
51760     sql = sql & "    where child_aliquot_id = a.ALIQUOT_ID "
51770     sql = sql & "    and parent_aliquot_id = " & strBlockId
51780     sql = sql & " )"
51790     sql = sql & " and a.STATUS not in ('X')"
51800     sql = sql & " order by a.ALIQUOT_ID"
          
51810     Set rs = connection.Execute(sql)
          
51820     While Not rs.EOF
51830         Call dicSlides.Add(nte(rs("NAME")), "")
51840         rs.MoveNext
51850     Wend

51860     Exit Sub
ERR_GetSlidesForBlock:
51870 MsgBox "ERR_GetSlidesForBlock" & vbCrLf & Err.Description
End Sub

'get a list of all slides for this Sample
'strSampleId (in) -
'dicSlide (out) - the list of slides
Private Sub GetSlidesForSample(strSampleId As String, dicSlides As Dictionary)
51880 On Error GoTo ERR_GetSlidesForSample
          Dim rs As Recordset
          Dim sql As String

51890     sql = " select a.NAME"
51900     sql = sql & " from lims_sys.aliquot a"
51910     sql = sql & " where exists"
51920     sql = sql & " ("
51930     sql = sql & "    select aliquot_id"
51940     sql = sql & "    from lims_sys.aliquot_formulation"
51950     sql = sql & "    where child_aliquot_id = a.ALIQUOT_ID "
51960     sql = sql & " )"
51970     sql = sql & " and a.SAMPLE_ID = " & strSampleId
51980     sql = sql & " and a.STATUS not in ('X')"
51990     sql = sql & " order by a.ALIQUOT_ID"
          
52000     Set rs = connection.Execute(sql)
          
52010     While Not rs.EOF
52020         Call dicSlides.Add(nte(rs("NAME")), "")
52030         rs.MoveNext
52040     Wend

52050     Exit Sub
ERR_GetSlidesForSample:
52060 MsgBox "ERR_GetSlidesForSample" & vbCrLf & Err.Description
End Sub

'get a list of all slides for this SDG
'strSdgId (in) -
'dicSlide (out) - the list of slides
Private Sub GetSlidesForSDG(strSdgId As String, dicSlides As Dictionary)
52070 On Error GoTo ERR_GetSlidesForSDG
          Dim rs As Recordset
          Dim sql As String

52080     sql = " select a.NAME"
52090     sql = sql & " from lims_sys.aliquot a, "
52100     sql = sql & "      lims_sys.sample s"
52110     sql = sql & " where exists"
52120     sql = sql & " ("
52130     sql = sql & "    select aliquot_id"
52140     sql = sql & "    from lims_sys.aliquot_formulation"
52150     sql = sql & "    where child_aliquot_id = a.ALIQUOT_ID "
52160     sql = sql & " )"
52170     sql = sql & " and a.SAMPLE_ID = s.SAMPLE_ID"
52180     sql = sql & " and s.SDG_ID = " & strSdgId
52190     sql = sql & " and a.STATUS not in ('X')"
52200     sql = sql & " order by a.ALIQUOT_ID"
          
52210     Set rs = connection.Execute(sql)
          
52220     While Not rs.EOF
52230         Call dicSlides.Add(nte(rs("NAME")), "")
52240         rs.MoveNext
52250     Wend

52260     Exit Sub
ERR_GetSlidesForSDG:
52270 MsgBox "ERR_GetSlidesForSDG" & vbCrLf & Err.Description
End Sub


'get the shortcut for the color group of this color
'an empty string is returned if this is not on of the known colors
Private Function GetColorGroup(strColor As String) As String
52280 On Error GoTo ERR_GetColorGroup
          Dim rs As Recordset
          Dim sql As String
          Dim s As String
          
      '    sql = " select ph.NAME "
      '    sql = sql & " from lims_sys.phrase_header ph,"
      '    sql = sql & "      lims_sys.phrase_entry pe"
      '    sql = sql & " where pe.PHRASE_ID=ph.PHRASE_ID"
      '    sql = sql & " and pe.PHRASE_NAME='" & strColor & "' "
      '    sql = sql & " and ph.NAME in ('Pathology Molecular Stains',"
      '    sql = sql & "                 'Pathology Special Stains',"
      '    sql = sql & "                 'Pathology Other Stain Options',"
      '    sql = sql & "                 'Pathology Imonohistochemistry stains')"

52290     sql = " select u_part_type "
52300     sql = sql & " from lims_sys.u_parts_user pu "
52310     sql = sql & "where "
52320     sql = sql & " U_stain ='" & strColor & "' "

52330     Set rs = connection.Execute(sql)
          
52340     If rs.EOF = True Then Exit Function
         
52350     s = nte(rs("u_part_type"))
      '
      '    Select Case s
      '        Case "Pathology Molecular Stains"
      '            s = "Mol"
      '        Case "Pathology Special Stains"
      '            s = "S"
      '        Case "Pathology Imonohistochemistry stains"
      '            s = "IHC"
      '        Case "Pathology Other Stain Options"
      '            s = "O"
      '    End Select

52360     GetColorGroup = s

52370     Exit Function
ERR_GetColorGroup:
52380 MsgBox "ERR_GetColorGroup" & vbCrLf & Err.Description
End Function


Private Sub grid_DblClick()
52390 On Error GoTo ERR_grid_DblClick

52400     Call frmRemarks.Initialize(grid.Text, True, "")
52410     Call frmRemarks.Show(vbModal)


52420     Exit Sub
ERR_grid_DblClick:
52430 MsgBox "ERR_grid_DblClick" & vbCrLf & Err.Description
End Sub

'a click on a cell shows the remark for the request
'containing this entity
'Private Sub grid_Click()
'On Error GoTo ERR_grid_Click
'    Dim strRequestDataId As String
'    Dim sql As String
'    Dim rs As Recordset
'
'    strRequestDataId = grid.TextMatrix(grid.row, 0)
'
'    sql = " select r.DESCRIPTION "
'    sql = sql & " from lims_sys.u_extra_request_data rd, "
'    sql = sql & "      lims_sys.u_extra_request_data_user rdu, "
'    sql = sql & "      lims_sys.u_extra_request r"
'    sql = sql & "  where rd.U_EXTRA_REQUEST_DATA_ID=rdu.U_EXTRA_REQUEST_DATA_ID"
'    sql = sql & "  and   r.U_EXTRA_REQUEST_ID=rdu.U_EXTRA_REQUEST_ID"
'    sql = sql & "  and   rd.U_EXTRA_REQUEST_DATA_ID=" & strRequestDataId
'
'    Set rs = connection.Execute(sql)
'
'    If rs.EOF = True Then Exit Sub
'
'    Call frmRemarks.Initialize(nte(rs("DESCRIPTION")), True)
'    Call frmRemarks.Show(vbModal)
'
''    Call frmRemarks.Initialize(connection, strRequestDataId)
''    Call frmRemarks.Show(vbModal)
'
'    Exit Sub
'ERR_grid_Click:
'MsgBox "MsgBox" & vbCrLf & Err.Description
'End Sub

'update aliqout station - to the block above this slide
'update the color of this slide to the chosen color



Private Sub UpdateReserveSlide(strSlideName As String, strNewColor As String)
52440 On Error GoTo ERR_UpdateReserveSlide
          Dim sql As String
          Dim strSlideId As String
          Dim strBlockId As String
          Dim rs As Recordset
          
52450     strSlideId = GetAliquotId(strSlideName)
          
52460     sql = "  update lims_sys.aliquot_user au "
52470     sql = sql & "  set au.U_COLOR_TYPE='" & FixQuotes(strNewColor) & "'"
52480     sql = sql & "  where au.ALIQUOT_ID=" & strSlideId
52490     Call connection.Execute(sql)
          
52500     sql = " select a.ALIQUOT_ID "
52510     sql = sql & " from lims_sys.aliquot a"
52520     sql = sql & " where exists"
52530     sql = sql & " ("
52540     sql = sql & "   select 1 "
52550     sql = sql & "   from lims_sys.aliquot_formulation af"
52560     sql = sql & "   where af.PARENT_ALIQUOT_ID=a.ALIQUOT_ID"
52570     sql = sql & "   and   af.CHILD_ALIQUOT_ID=" & strSlideId
52580     sql = sql & " )"
52590     Set rs = connection.Execute(sql)
52600     strBlockId = nte(rs("ALIQUOT_ID"))

52610     Call UpdateAliquotTrace(strBlockId, "Cleen up")

52620     Exit Sub
ERR_UpdateReserveSlide:
52630 MsgBox "ERR_UpdateReserveSlide" & vbCrLf & Err.Description
End Sub

'set station 1 ( "number histology") for all the aliquots of that sample
Private Sub UpdateAliquotesForSample(strSampleName As String)
52640 On Error GoTo ERR_UpdateAliquotesForSample
          Dim rs As Recordset
          Dim sql As String

52650     sql = " select a.ALIQUOT_ID"
52660     sql = sql & " from lims_sys.sample s, lims_sys.aliquot a"
52670     sql = sql & " where a.SAMPLE_ID=s.SAMPLE_ID"
52680     sql = sql & " and s.NAME='" & strSampleName & "'"
52690     sql = sql & " and exists "
52700     sql = sql & " ("
52710     sql = sql & "    select 1"
52720     sql = sql & "    from lims_sys.aliquot_formulation af"
52730     sql = sql & "    where af.parent_aliquot_id = a.aliquot_id"
52740     sql = sql & " )"

52750     Set rs = connection.Execute(sql)

52760     While Not rs.EOF
52770         Call UpdateAliquotTrace(nte(rs("ALIQUOT_ID")), "number histology")
              
52780         rs.MoveNext
52790     Wend

52800     Exit Sub
ERR_UpdateAliquotesForSample:
52810 MsgBox "ERR_UpdateAliquotesForSample" & vbCrLf & Err.Description
End Sub


'update the aliquot record to show that it
'was in this station.
'if it already was here, do not update
Private Sub UpdateAliquotTrace(strAliquotId As String, strStationName As String)
52820 On Error GoTo REE_UpdateAliquotTrace

          Dim rs As Recordset
          Dim strSQL As String
          Dim strStationNumber As String
      '    Dim strOldStation
          
          'get this station's name from the phrase:
52830     Set rs = _
    connection.Execute("select phrase_name from lims_sys.phrase_entry " & _
    "where phrase_description = '" & strStationName & "' and " & _
    "phrase_id = (select phrase_id from lims_sys.phrase_header where " & _
    "name = 'AliquotStationTrace') " & "order by order_number")
          
52840     strStationNumber = rs("phrase_name")
          
          'check if this aliquot was in this station:
      '    strSql = "select u_old_aliquot_station " & "from lims_sys.aliquot_user " & "where aliquot_id = " & strAliquotId
                   
      '    Set rs = Con.Execute(strSql)
      '    strOldStation = nte(rs("u_old_aliquot_station"))
          
          'update pass through this station if needed:
      '    If InStr(1, strOldStation, StrStationName, vbTextCompare) = 0 Then
         
52850         strSQL = " update lims_sys.aliquot_user set " & _
    " u_old_aliquot_station = u_old_aliquot_station || u_aliquot_station , " & _
    " u_aliquot_station = '" & strStationNumber & "' " & " where aliquot_id = " & _
    strAliquotId
         
         '     strSql = "update lims_sys.aliquot_user " & "set u_aliquot_station = '" & StrStationName & "', " & "u_old_aliquot_station = u_old_aliquot_station || '" & StrStationName & "' " & "where aliquot_id = " & strAliquotId
                       
52860         Call connection.Execute(strSQL)
      '    End If
                   
52870     Exit Sub
REE_UpdateAliquotTrace:
52880 MsgBox "REE_UpdateAliquotTrace" & vbCrLf & Err.Description
End Sub


Private Function GetAliquotId(strAliquotName As String) As String
52890 On Error GoTo ERR_GetAliquotId
          Dim rs As Recordset
          Dim sql As String
          
52900     sql = "  select a.ALIQUOT_ID"
52910     sql = sql & "  from lims_sys.aliquot a"
52920     sql = sql & "  where a.NAME='" & strAliquotName & "'"

52930     Set rs = connection.Execute(sql)

52940     If rs.EOF Then Exit Function
          
52950     GetAliquotId = nte(rs("ALIQUOT_ID"))

52960     Exit Function
ERR_GetAliquotId:
52970 MsgBox "ERR_GetAliquotId" & vbCrLf & Err.Description
End Function
'PAT 002
Private Function GetSampleId(strSampleName As String) As String
52980 On Error GoTo ERR_GetSampleId
          Dim rs As Recordset
          Dim sql As String
          
52990     sql = "  select s.Sample_ID"
53000     sql = sql & "  from lims_sys.Sample s"
53010     sql = sql & "  where s.NAME='" & strSampleName & "'"
53020     sql = sql & "     and s.status<>'X' "

53030     Set rs = connection.Execute(sql)

53040     If rs.EOF Then Exit Function
          
53050     GetSampleId = nte(rs("Sample_ID"))

53060     Exit Function
ERR_GetSampleId:
53070 MsgBox "ERR_GetSampleId" & vbCrLf & Err.Description
End Function
Private Function GetMaxSlide(ParentID As Long) As Long
53080     On Error GoTo ErrEnd
          Dim strSQL As String
          Dim SlideRec As ADODB.Recordset

53090     GetMaxSlide = 0
53100     strSQL = "select max(a.aliquot_id) " & "from lims_sys.aliquot a " & _
    "where a.aliquot_id in " & _
    "(select child_aliquot_id from lims_sys.aliquot_formulation " & _
    "where aliquot_formulation.parent_aliquot_id = '" & ParentID & "') " & _
    "order by a.aliquot_id"
53110     Set SlideRec = connection.Execute(strSQL)

53120     If Not SlideRec.EOF Then
53130         GetMaxSlide = SlideRec(0)
53140     End If
53150     SlideRec.Close
53160     Exit Function
ErrEnd:
53170     MsgBox "GetMaxSlide... " & vbCrLf & "Parent Aliquot ID = " & ParentID _
    & vbCrLf & Err.Description
End Function
'PAT 002
Private Function GetMaxSlideForSample(SampleID As String) As Long
53180     On Error GoTo ErrEnd
          Dim sql As String
          Dim SlideRec As ADODB.Recordset

53190     GetMaxSlideForSample = 0
53200     sql = " SELECT   MAX (a.aliquot_id)"
53210     sql = sql & "     FROM lims_sys.aliquot a"
53220     sql = sql & "    WHERE a.sample_id = " & SampleID & " "
53230     sql = sql & "      AND NOT EXISTS (SELECT 1"
53240     sql = sql & _
    "                        FROM lims_sys.aliquot_formulation af"
53250     sql = sql & _
    "                       WHERE af.parent_aliquot_id = a.aliquot_id)"
53260     Set SlideRec = connection.Execute(sql)

53270     If Not SlideRec.EOF Then
53280         GetMaxSlideForSample = SlideRec(0)
53290     End If
53300     SlideRec.Close
53310     Exit Function
ErrEnd:
53320     MsgBox "GetMaxSlideForSample... " & vbCrLf & "Parent Sample ID = " & _
    SampleID & vbCrLf & Err.Description
End Function
Private Function TriggerSampleEvent(EventName As String, SampleID As String) As _
    Long
53330     On Error GoTo ErrEnd
          Dim doc As New DOMDocument
          Dim res As New DOMDocument
          Dim xmlLogin As IXMLDOMElement
          Dim xmlSdg As IXMLDOMElement
          Dim e As IXMLDOMElement
          Dim element As IXMLDOMElement
          Dim FileName As String
          Dim RetError As String

53340     Set e = doc.createElement("lims-request")
53350     Call doc.appendChild(e)
53360     Set xmlLogin = doc.createElement("login-request")
53370     Call e.appendChild(xmlLogin)
53380     Set xmlSdg = doc.createElement("SAMPLE")
53390     Call xmlLogin.appendChild(xmlSdg)
53400     Set element = doc.createElement("find-by-id")
53410     element.Text = SampleID
53420     Call xmlSdg.appendChild(element)
53430     Set element = doc.createElement("fire-event")
53440     element.Text = EventName
53450     Call xmlSdg.appendChild(element)

      '    doc.Save ("c:\Sampledoc.xml")

53460     If Trim(WorkFolder) <> "" Then
53470         FileName = "ResultEntry_AdditionalChanges_" & EventName & "_" & _
    SampleID & "_DOC3"
53480         Call xmlManager.SaveXmlFile(doc, FileName)
53490     End If

53500     RetError = ProcessXML.ProcessXMLWithResponse(doc, res)
53510     If Trim(RetError) <> "" Then
53520         MsgBox _
    "Error occurred while trying process xml file. (TriggerSampleEvent) " & vbCrLf _
    & "Sample ID: " & SampleID & vbCrLf & "Event Name: " & EventName & vbCrLf & _
    "Error: " & RetError, vbCritical, "Nautilus - ResultEntry_AdditionalChanges"
53530     End If

      '    res.Save ("c:\Sampleres.xml")

53540     If Trim(WorkFolder) <> "" Then
53550         FileName = "ResultEntry_" & EventName & "_" & SampleID & "_RES3"
53560         Call xmlManager.SaveXmlFile(res, FileName)
53570     End If

53580     TriggerSampleEvent = res.SelectSingleNode("//ALIQUOT_ID").Text
53590     Exit Function

ErrEnd:
53600     MsgBox "TriggerSampleEvent... " & vbCrLf & "Sample ID = " & SampleID _
    & vbCrLf & "Event Name = " & EventName & vbCrLf & Err.Description
End Function

Private Function TriggerSlideEvent(EventName As String, AliquotID As String) As _
    Long
53610     On Error GoTo ErrEnd
          Dim doc As New DOMDocument
          Dim res As New DOMDocument
          Dim xmlLogin As IXMLDOMElement
          Dim xmlSdg As IXMLDOMElement
          Dim e As IXMLDOMElement
          Dim element As IXMLDOMElement
          Dim FileName As String
          Dim RetError As String

53620     Set e = doc.createElement("lims-request")
53630     Call doc.appendChild(e)
53640     Set xmlLogin = doc.createElement("login-request")
53650     Call e.appendChild(xmlLogin)
53660     Set xmlSdg = doc.createElement("ALIQUOT")
53670     Call xmlLogin.appendChild(xmlSdg)
53680     Set element = doc.createElement("find-by-id")
53690     element.Text = AliquotID
53700     Call xmlSdg.appendChild(element)
53710     Set element = doc.createElement("fire-event")
53720     element.Text = EventName
53730     Call xmlSdg.appendChild(element)

      '    doc.Save ("c:\slidedoc.xml")

53740     If Trim(WorkFolder) <> "" Then
53750         FileName = "ResultEntry_" & EventName & "_" & AliquotID & "_DOC4"
53760         Call xmlManager.SaveXmlFile(doc, FileName)
53770     End If

53780     RetError = ProcessXML.ProcessXMLWithResponse(doc, res)
53790     If Trim(RetError) <> "" Then
53800         MsgBox _
    "Error occurred while trying process xml file. (TriggerSlideEvent) " & vbCrLf & _
    "Aliquot ID: " & AliquotID & vbCrLf & "Event Name: " & EventName & vbCrLf & _
    "Error: " & RetError, vbCritical, "Nautilus - Result Entry"
53810     End If

      '    res.Save ("c:\slideres.xml")

53820     If Trim(WorkFolder) <> "" Then
53830         FileName = "ResultEntry_" & EventName & "_" & AliquotID & "_RES4"
53840         Call xmlManager.SaveXmlFile(res, FileName)
53850     End If

53860     TriggerSlideEvent = res.SelectSingleNode("//ALIQUOT_ID").Text
53870     Exit Function

ErrEnd:
53880     MsgBox "TriggerSlideEvent... " & vbCrLf & "Aliquot ID = " & AliquotID _
    & vbCrLf & "Event Name = " & EventName & vbCrLf & Err.Description
End Function

'fix the names of slides for this cassette:
Private Sub UpdateSlides4Cassette(ParentAliquotID As String, CassetteName As _
    String)
53890     On Error GoTo ErrEnd
          Dim i As Integer
          Dim strSQL As String
          Dim SlideRs As ADODB.Recordset
          Dim NewSlideName As String
          Dim strSlideId As String

53900     strSQL = "select a.aliquot_id " & "from lims_sys.aliquot a " & _
    "where a.aliquot_id in " & "(select child_aliquot_id " & _
    "from lims_sys.aliquot_formulation " & _
    "where aliquot_formulation.parent_aliquot_id = '" & ParentAliquotID & "') " & _
    "order by a.aliquot_id"

53910     Set SlideRs = connection.Execute(strSQL)

53920     i = 0
53930     While Not SlideRs.EOF
53940         i = i + 1
53950         NewSlideName = CassetteName & "." & i
              
53960         strSlideId = nte(SlideRs(0))
              
53970         Call connection.Execute("update lims_sys.aliquot " & _
    "set name = '" & NewSlideName & "' " & "where aliquot_id = " & SlideRs(0))
53980         SlideRs.MoveNext
53990     Wend
54000     SlideRs.Close
54010     Exit Sub
ErrEnd:
54020     MsgBox "UpdateSlides4Cassette... " & vbCrLf & "Parent Aliquot ID = " _
    & ParentAliquotID & vbCrLf & "CassetteName = " & CassetteName & vbCrLf & _
    Err.Description
End Sub

Private Sub UpdateSlides4Sample(ParentSampleID As String, SampleName As String)
54030     On Error GoTo ErrEnd
          Dim i As Integer
          Dim sql As String
          Dim SlideRs As ADODB.Recordset
          Dim NewSlideName As String
          Dim strSlideId As String

54040     sql = " SELECT   a.aliquot_id"
54050     sql = sql & "     FROM lims_sys.aliquot a"
54060     sql = sql & "    WHERE a.sample_id=" & ParentSampleID & " "
54070     sql = sql & "      AND NOT EXISTS (SELECT 1"
54080     sql = sql & _
    "                        FROM lims_sys.aliquot_formulation af"
54090     sql = sql & _
    "                       WHERE af.parent_aliquot_id = a.aliquot_id)"
54100     sql = sql & "  order by a.ALIQUOT_ID"

54110     Set SlideRs = connection.Execute(sql)

54120     i = 0
54130     While Not SlideRs.EOF
54140         i = i + 1
54150         NewSlideName = SampleName & "." & i
              
54160         strSlideId = nte(SlideRs(0))
              
54170         Call connection.Execute("update lims_sys.aliquot " & _
    "set name = '" & NewSlideName & "' " & "where aliquot_id = " & SlideRs(0))
54180         SlideRs.MoveNext
54190     Wend
54200     SlideRs.Close
54210     Exit Sub
ErrEnd:
54220     MsgBox "UpdateSlides4Sample... " & vbCrLf & "Parent Sample ID = " & _
    ParentSampleID & vbCrLf & "SampleName = " & SampleName & vbCrLf & _
    Err.Description
End Sub
Private Sub UpdateDesc(strExtraRequestDataID As String, strSlideId As String)
54230 On Error GoTo ERR_UpdateDesc

          Dim sql As String
          
54240     sql = " update lims_sys.u_extra_request_data_user rdu "
54250     sql = sql & " set rdu.u_desc = '" & strSlideId & "', "
54260     sql = sql & "     rdu.u_slide_name =  "
54270     sql = sql & "  ( "
54280     sql = sql & "       select a.name "
54290     sql = sql & "       from lims_sys.aliquot a "
54300     sql = sql & "       where a.aliquot_id = '" & strSlideId & "' "
54310     sql = sql & "  ) "
54320     sql = sql & " where rdu.u_extra_request_data_id='" & _
    strExtraRequestDataID & "' "
          
54330     Call connection.Execute(sql)

54340     Exit Sub
ERR_UpdateDesc:
54350 MsgBox "ERR_UpdateDesc" & vbCrLf & Err.Description
End Sub

Private Sub UpdateSlideColor(strSlideId As String, strColor As String)
54360 On Error GoTo ERR_UpdateSlideColor
          Dim sql As String

54370     sql = " update lims_sys.aliquot_user au"
54380     sql = sql & " set au.U_COLOR_TYPE='" & strColor & "'"
54390     sql = sql & " where au.ALIQUOT_ID=" & strSlideId
          
54400     Call connection.Execute(sql)
              
54410     Exit Sub
ERR_UpdateSlideColor:
54420 MsgBox "ERR_UpdateSlideColor" & vbCrLf & Err.Description
End Sub

Private Sub UpdateSlideLayers(strSlideId As String, strLayers As String)
54430 On Error GoTo ERR_UpdateSlideLayers
          Dim sql As String

54440     sql = " update lims_sys.aliquot_user au"
54450     sql = sql & " set au.U_FORMATTED_COLOR='" & strLayers & "'"
54460     sql = sql & " where au.ALIQUOT_ID=" & strSlideId
          
54470     Call connection.Execute(sql)
              
54480     Exit Sub
ERR_UpdateSlideLayers:
54490 MsgBox "ERR_UpdateSlideLayers" & vbCrLf & Err.Description
End Sub


Private Function GetBlockId(strSlideId As String) As String
54500 On Error GoTo ERR_GetBlockId
          Dim rs As Recordset
          Dim sql As String

54510     sql = " select af.PARENT_ALIQUOT_ID"
54520     sql = sql & " from lims_sys.aliquot a, "
54530     sql = sql & "      lims_sys.aliquot_formulation af"
54540     sql = sql & " where a.ALIQUOT_ID=af.CHILD_ALIQUOT_ID"
54550     sql = sql & " and   a.ALIQUOT_ID=" & strSlideId

54560     Set rs = connection.Execute(sql)
          
54570     If Not rs.EOF Then
54580         GetBlockId = rs(0)
54590     End If

54600     Exit Function
ERR_GetBlockId:
54610 MsgBox "ERR_GetBlockId" & vbCrLf & Err.Description
End Function


Private Sub LoadPics()
54620 On Error GoTo ERR_LoadPics

54630     cmdOKAdvisors.Picture = LoadPicture("Resource\Tick.ico")
54640     cmdOKColors.Picture = LoadPicture("Resource\Tick.ico")
54650     cmdOKReEmbedding.Picture = LoadPicture("Resource\Tick.ico")
54660     cmdOKShowOriginalSample.Picture = LoadPicture("Resource\Tick.ico")
54670     cmdOKSlidesFromArchive.Picture = LoadPicture("Resource\Tick.ico")
54680     cmdBlockReset(0).Picture = LoadPicture("Resource\cancel_overlay.ico")
54690     cmdSlideReset(0).Picture = LoadPicture("Resource\cancel_overlay.ico")
54700     cmdDeleteSlidesFromArchive.Picture = _
    LoadPicture("Resource\cancel_overlay.ico")
54710     cmdDeleteEntity.Picture = LoadPicture("Resource\cancel_overlay.ico")
54720     cmdAddToSlidesFromArchive.Picture = _
    LoadPicture("Resource\move Right.ico")
54730     cmdSelectEntity.Picture = LoadPicture("Resource\move Right.ico")
      '    cmdExistingLetters.Picture = LoadPicture("Resource\Address.ico")
          
54740     Exit Sub
ERR_LoadPics:
54750 MsgBox "ERR_LoadPics" & vbCrLf & Err.Description
End Sub


Private Sub InitAdvisorsRequestsList(strExternalReference As String)
54760 On Error GoTo ERR_InitAdvisorsRequestsList
          Dim rs As Recordset
          Dim sql As String
          Dim i As Integer
          
          'unload any previously loaded rows:
54770     For i = 1 To lblDate.Count - 1
54780         Unload lblDate(i)
54790         Unload lblAdvisor(i)
54800         Unload lblRequestId(i)
54810         Unload chkAdvisorReturn(i)
54820         Unload txtAdvisorRemarks(i)
54830         Unload cmdAdvisorLetter(i)
54840         Unload cmdAdvisorEntryOK(i)
54850     Next i
          
54860     sql = " select r.U_EXTRA_REQUEST_ID,"
54870     sql = sql & "        ru.U_CREATED_ON,"
54880     sql = sql & "     r.DESCRIPTION,"
54890     sql = sql & "        (select rdu.U_REQUEST_DETAILS"
54900     sql = sql & "     from lims_sys.u_extra_request_data_user rdu"
54910     sql = sql & "     where rdu.U_EXTRA_REQUEST_ID=r.U_EXTRA_REQUEST_ID"
54920     sql = sql & "     and rownum=1) ADVISOR,"
54930     sql = sql & "        (select rdu.U_STATUS"
54940     sql = sql & "     from lims_sys.u_extra_request_data_user rdu"
54950     sql = sql & "     where rdu.U_EXTRA_REQUEST_ID=r.U_EXTRA_REQUEST_ID"
54960     sql = sql & "     and rownum=1) STATUS        "
54970     sql = sql & " from  lims_sys.u_extra_request r,"
54980     sql = sql & "       lims_sys.u_extra_request_user ru, "
54990     sql = sql & "       lims_sys.sdg d "
55000     sql = sql & " where r.U_EXTRA_REQUEST_ID=ru.U_EXTRA_REQUEST_ID"
55010     sql = sql & " and   d.sdg_id = ru.U_SDG_ID "
55020     sql = sql & " and   d.external_reference = '" & strExternalReference _
    & "' "
          'sql = sql & " and   ru.U_SDG_ID='" & strSdgId & "'"
          'sql = sql & " and   ru.U_EXTERNAL_REFERENCE='" & strExternalReference & "'"
55030     sql = sql & " and   r.NAME like 'Send to Consultant;%' "
55040     sql = sql & " and   (ru.U_STATUS <> 'X' or ru.U_STATUS is null) "
55050     sql = sql & " order by r.U_EXTRA_REQUEST_ID"
          
55060     Set rs = connection.Execute(sql)
          
        
          
          'add the rows to the screen:
55070     While Not rs.EOF
55080         Call AddToAdvisorList(rs)
55090         rs.MoveNext
55100     Wend

55110     Exit Sub
ERR_InitAdvisorsRequestsList:
55120 MsgBox "ERR_InitAdvisorsRequestsList" & vbCrLf & Err.Description
End Sub


'check in the letters folder to see if the letter exists
'for this extra request
Private Function ExistLetterToAdvisor(strLetterName As String) As Boolean
55130 On Error GoTo ERR_ExistLetterToAdvisor

          Dim fs As New FileSystemObject

55140     If fs.FileExists(strLettersFolder & strLetterName) = True Then
55150         ExistLetterToAdvisor = True
55160     Else
55170         ExistLetterToAdvisor = False
55180     End If

55190     Exit Function
ERR_ExistLetterToAdvisor:
55200 MsgBox "ERR_ExistLetterToAdvisor" & vbCrLf & Err.Description
End Function


Private Sub DisableFieldsForAuthorizedRequest()
55210 On Error GoTo ERR_DisableFieldsForAuthorizedRequest
          
55220     If rsSdg("status") = "A" Then
55230         Call DisableField(picBlocks)
55240         Call DisableField(picBlockReembedding)
      '        picBlocks.Enabled = False
      '        picBlocks.ToolTipText = ""
          
55250     End If

55260     Exit Sub
ERR_DisableFieldsForAuthorizedRequest:
55270 MsgBox "ERR_DisableFieldsForAuthorizedRequest" & vbCrLf & Err.Description
End Sub

Private Sub DisableField(o As Object)
55280 On Error GoTo ERR_DisableField

55290     o.Enabled = False
55300     o.ToolTipText = "Access is denied - Request is already authorized"
          
55310     Exit Sub
ERR_DisableField:
55320 MsgBox "ERR_DisableField" & vbCrLf & Err.Description
End Sub


'initialise the lists of special stains
'where one stain name stands for a group of slides:
Private Sub InitTemplateStains()
55330 On Error GoTo ERR_InitTemplateStains
          
          Dim i As Integer
          Dim rs As Recordset
          Dim d As Dictionary
          Dim ts As TemplateSlide
          
55340     Call dicTemplateSlides.RemoveAll

          'get the names of the phrases
          'each holding a set of slides to be the result
          'of a special new slide request:
55350     Set rs = _
    connection.Execute("select phrase_name from lims_sys.phrase_entry " & _
    "where phrase_id = (select phrase_id from lims_sys.phrase_header where " & _
    "name = 'Phrase Names for Slide Templates') " & "order by order_number")
                  
55360     While Not rs.EOF
          
55370         Call dicTemplateSlides.Add(nte(rs("phrase_name")), Null)
55380         rs.MoveNext
          
55390     Wend
                  
                  
          'get the keys (phrase names) for the special color names
      '    Call dicTemplateSlides.Add("SLN / CKMNF116", Null)
      '    Call dicTemplateSlides.Add("SLN / S-100", Null)
      '    Call dicTemplateSlides.Add("SLN / Reserve", Null)
          
55400     For i = 0 To dicTemplateSlides.Count - 1
          
55410         Set rs = _
    connection.Execute("select phrase_name, phrase_description, phrase_info from lims_sys.phrase_entry " _
    & "where phrase_id = (select phrase_id from lims_sys.phrase_header where " & _
    "name = '" & CStr(dicTemplateSlides.Keys(i)) & "') " & "order by order_number")
          
55420         Set d = New Dictionary
              
              'build the desired set of slides to be created
              'for the current color name:
55430         While Not rs.EOF
55440             Set ts = New TemplateSlide
55450             Call ts.Initialize(nte(rs("phrase_description")), _
    nte(rs("phrase_info")))
55460             Call d.Add(nte(rs("PHRASE_NAME")), ts)
55470             rs.MoveNext
55480         Wend
       
              'hold the current list of slides against the current color name
55490         Set dicTemplateSlides(CStr(dicTemplateSlides.Keys(i))) = d
          
55500     Next i


55510     Exit Sub
ERR_InitTemplateStains:
55520 MsgBox "ERR_InitTemplateStains" & vbCrLf & Err.Description
End Sub


'when the color name is formatted this way: color_name  #count
's(in)       - the original string
'sColor(out) - then color name
'sCount(out) - the count
Private Sub ParseBlockColor(s As String, sColor As String, sCount As String)
55530 On Error GoTo ERR_ParseBlockColor
          Dim i As Integer
          
55540     i = InStr(1, s, "#")
55550     If i <> 0 Then
55560         sColor = Mid(s, 1, i - 3)
55570         sCount = Mid(s, i + 1)
55580     Else
55590         sColor = s
55600         sCount = "0"
55610     End If

55620     Exit Sub
ERR_ParseBlockColor:
55630 MsgBox "ERR_ParseBlockColor" & vbCrLf & Err.Description
End Sub


Private Function GetBlockName(strSlideName As String) As String
55640 On Error GoTo ERR_GetBlockName

          Dim strBlockName As String
          Dim i As Integer
          
55650     i = InStr(1, strSlideName, ".")
55660     i = InStr(i + 1, strSlideName, ".")
          
55670     strBlockName = Left(strSlideName, i - 1)
55680     GetBlockName = strBlockName

55690     Exit Function
ERR_GetBlockName:
55700 MsgBox "ERR_GetBlockName" & vbCrLf & Err.Description
End Function


'updates the relevant table by the given id
'to the values held by the recordset
'strIdFieldName - the id field name for this table
'strIdTarget    - the id field value for the record to be modified
'strIdSource    - the id field value for the record to read from
Private Sub UpdateRecordById(strTableName As String, strIdFieldName As String, _
    strIdTarget As String, strIdSource As String, rs As Recordset, dicStopList As _
    Dictionary)
                                  
55710 On Error GoTo ERR_UpdateRecordById
          
          Dim i As Integer
          Dim strFieldName As String
          Dim varFieldValue As Variant
          Dim sql As String
              
55720     For i = 0 To rs.Fields.Count - 1
55730         strFieldName = rs.Fields(i).name
55740         varFieldValue = rs.Fields(i).value
              
55750         If dicStopList.Exists(LCase(strFieldName)) = False Then
                      
                  'copy the data
                  'directly from the data base:
55760             sql = " update lims_sys." & strTableName
55770             sql = sql & " set " & strFieldName & " = "
55780             sql = sql & " ( "
55790             sql = sql & "    select " & strFieldName
55800             sql = sql & "    from lims_sys." & strTableName
55810             sql = sql & "    where " & strIdFieldName & " = " & strIdSource
55820             sql = sql & " ) "
55830             sql = sql & " where " & strIdFieldName & " = " & strIdTarget
                  
55840             connection.Execute (sql)

55850         End If
              
55860     Next i

55870 Exit Sub
ERR_UpdateRecordById:
55880 MsgBox "ERR_UpdateTableById" & vbCrLf & Err.Description
End Sub


'allow browsing the old advise letters if there
'are already some for the request:
'Private Sub InitAccessToOldLetters(strSdgName As String)
'On Error GoTo ERR_InitAccessToOldLetters
'
'    Dim fs As New FileSystemObject
'
'    Dim files, file
'    Dim s As String
'
'    Dim i As Integer
'
'    Set files = fs.GetFolder(LETTERS_FOLDER).files
'
'    cmdExistingLetters.Visible = False
'
'    For Each file In files
'        s = file.name
'        If InStr(1, s, Replace(strSdgName, "/", "_")) <> 0 Then
'            cmdExistingLetters.Visible = True
'        End If
'    Next
'
'    Exit Sub
'ERR_InitAccessToOldLetters:
'MsgBox "ERR_InitAccessToOldLetters" & vbCrLf & Err.Description
'End Sub
