VERSION 5.00
Begin VB.Form frmResponseLetter 
   Caption         =   "Response Letter"
   ClientHeight    =   9315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   ScaleHeight     =   9315
   ScaleWidth      =   10905
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmResponseLetter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private pdfPath As String
Private Const RESPONSE_LETTER = "\\pdf-storage\pdf\#sdg_id#.PDF"
Private pdf As AcroPDFLibCtl.AcroPDF
'Private AcroPDF As AcroPDF

Public Function Initialize(strSdgId_ As String) As String
29610 On Error GoTo ERR_Initialize
          'If pdf Is Nothing Then
29620         Call LoadPDFControl
          'End If

29630     pdfPath = strSdgId_
      '    AcroPDF.LoadFile ("\\limserver\pdf\" & strSdgId & ".PDF")
29640     Call pdf.LoadFile(pdfPath)
          
29650     Exit Function
ERR_Initialize:
29660 Initialize = Err.Number
End Function

Private Sub ShowPDF()
29670 On Error GoTo ERR_ShowPDF
      '  MsgBox strSdgId
      '    AcroPDF.LoadFile ("\\limserver\pdf\" & strSdgId & ".PDF")
29680     Call pdf.LoadFile(pdfPath)

29690     Exit Sub
ERR_ShowPDF:
      'MsgBox "ERR_ShowPDF" & vbCrLf & Err.Description & vbCrLf & "Wasn't able to load PDF file"

      'Unload Me
      'frmResponseLetter.Hide
End Sub

Private Sub Form_Load()
'MsgBox 1
'    Set AcroPDF = Controls.Add("AcroPDFLibCtl.AcroPDF.1", "AcroPDF", True)
'    With AcroPDF
'        .Height = 9135
'        .Left = 120
'        .TabIndex = 1
'        .Top = 120
'        .Width = 10695
'    End With
'MsgBox 1


'    ShowPDF
End Sub

Private Sub LoadPDFControl()
29700 On Error GoTo ERR_LoadPDFControl
              
29710     Set pdf = Controls.Add("AcroPDF.PDF.1", "Test")
29720     pdf.Top = 0
29730     pdf.Height = 9720
29740     pdf.Left = 0
29750     pdf.Width = 10900
29760     pdf.Visible = True
          
29770     Exit Sub
ERR_LoadPDFControl:

End Sub
