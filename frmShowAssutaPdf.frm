VERSION 5.00
Object = "{C416B16F-676E-4818-9EFF-634C747EE47A}#1.0#0"; "DisplayPdf.ocx"
Begin VB.Form frmShowAssutaPdf 
   Caption         =   "Form1"
   ClientHeight    =   10410
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12885
   LinkTopic       =   "Form1"
   ScaleHeight     =   10410
   ScaleWidth      =   12885
   StartUpPosition =   3  'Windows Default
   Begin DisplayPdf.DisplayPdfCtrl DisplayPdfCtrl1 
      Height          =   10455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   18441
   End
End
Attribute VB_Name = "frmShowAssutaPdf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public assutaMacase As String
Public assutaPdfPath As String
Public IsRead As Boolean
Public IsReadDescription As String

Public Sub Init()

55890  DisplayPdfCtrl1.Init
55900     DisplayPdfCtrl1.assutaMacase = assutaMacase
55910     DisplayPdfCtrl1.assutaPdfPath = assutaPdfPath
55920     DisplayPdfCtrl1.ShowPDF
55930     DisplayPdfCtrl1.VisibleCheckButtons = False
55940     IsRead = DisplayPdfCtrl1.IsRead
55950     IsReadDescription = DisplayPdfCtrl1.IsReadDescription

End Sub

