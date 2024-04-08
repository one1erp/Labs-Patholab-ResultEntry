VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form frmColors 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "בחירת צביעות"
   ClientHeight    =   8775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   15345
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbNumOfSlides 
      Height          =   315
      Left            =   1440
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ListBox lstColorType 
      Height          =   780
      Left            =   11880
      TabIndex        =   3
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton cmdClose 
      Height          =   840
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   840
   End
   Begin MSFlexGridLib.MSFlexGrid gridColors 
      Height          =   7575
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   13361
      _Version        =   393216
      Rows            =   10
      Cols            =   7
      FixedRows       =   0
      FixedCols       =   0
      ScrollBars      =   2
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
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000A&
      Caption         =   "סוגי צביעות"
      Height          =   255
      Left            =   13920
      TabIndex        =   0
      Top             =   165
      Width           =   1215
   End
End
Attribute VB_Name = "frmColors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private iBlock As Integer
Private iSlide As Integer
Private iSlideColorRow As Integer
Private iSlideColorCol As Integer
Private dicMolecularStains As New Dictionary
Private dicSpecialStains As New Dictionary
Private dicImonohistochemistryStains As New Dictionary
Private dicHistochemistryStains As New Dictionary
Private dicOtherStainOptions As New Dictionary
Private isGridInitialized As Boolean

'holds the coloring entries for the block
'key - color name
'item - number of desired slides
Private dicBlockColors As New Dictionary

'Private Const GRID_ROWS = 20
Private nGridRows As Integer
Private Const GRID_COLS = 10

Private Const MARK_SELECTED = &HC0FFFF


Public Sub Initialize(iBlock_ As Integer, iSlide_ As Integer, strName As String, _
    dicMolecularStains_ As Dictionary, dicSpecialStains_ As Dictionary, _
    dicImonohistochemistryStains_ As Dictionary, dicHistochemistryStains_ As _
    Dictionary, dicOtherStainOptions_ As Dictionary)
29780 On Error GoTo ERR_Initialize
          
29790     iBlock = iBlock_
29800     iSlide = iSlide_
          
29810     Call InitBlockColors(iBlock)
                            
29820     isGridInitialized = False
          
29830     cmdClose.Picture = LoadPicture("Resource\Tick.ico")
                            
29840     Set dicMolecularStains = dicMolecularStains_
29850     Set dicSpecialStains = dicSpecialStains_
29860     Set dicImonohistochemistryStains = dicImonohistochemistryStains_
29870     Set dicHistochemistryStains = dicHistochemistryStains_
29880     Set dicOtherStainOptions = dicOtherStainOptions_
          
29890     Call InitColorList
          
29900     Call InitGrid
29910     isGridInitialized = True
          
      '    cmbColorType.Text = cmbColorType.list(0)
29920     Call InitGridColors(lstColorType.Text)
      '    Call InitGridColors(cmbColorType.Text)
          
          
29930     If iSlide = 0 Then
29940         frmColors.Caption = "Stain Selection for block " & strName
29950     Else
29960         frmColors.Caption = "Stain Selection for slide " & strName
29970     End If
          
29980     Call InitComboNumOfSlides
          
29990     Exit Sub
ERR_Initialize:
30000 MsgBox "ERR_initialize" & vbCrLf & Err.Description
End Sub


Private Sub cmbNumOfSlides_Click()
30010 On Error GoTo ERR_cmbNumOfSlides_Change

          'write the selection to the grid cell:
30020     gridColors.TextMatrix(gridColors.row, gridColors.col + 1) = _
    cmbNumOfSlides.Text

          'replace the number of slides:
30030     dicBlockColors(gridColors.Text) = cmbNumOfSlides.Text

30040     Exit Sub
ERR_cmbNumOfSlides_Change:
30050 MsgBox "ERR_cmbNumOfSlides_Change" & vbCrLf & Err.Description
End Sub

'Private Sub cmbColorType_Click()
'    Call InitGridColors(cmbColorType.Text)
'End Sub

Private Sub CmdClose_Click()
30060 On Error GoTo ERR_cmdClose_Click
           
30070     Call RefreshBlockCombo
30080     Me.Hide
          
30090     Exit Sub
ERR_cmdClose_Click:
30100 MsgBox "ERR_cmdClose_Click" & vbCrLf & Err.Description
End Sub


'update the block for which this form was opened
'to the selected colors / counts
Private Sub RefreshBlockCombo()
30110 On Error GoTo ERR_RefreshBlockCombo
30120     If iBlock = 0 Then Exit Sub
          
          Dim i As Integer
          
30130     Call frmAdditionalActions.CmbBlockEntry(iBlock).Clear
          
30140     For i = 0 To dicBlockColors.Count - 1
30150         Call _
    frmAdditionalActions.CmbBlockEntry(iBlock).AddItem(dicBlockColors.Keys(i) & _
    "  #" & dicBlockColors.Items(i))
30160     Next i

30170     Exit Sub
ERR_RefreshBlockCombo:
30180 MsgBox "ERR_RefreshBlockCombo" & vbCrLf & Err.Description
End Sub


Private Sub Form_Unload(Cancel As Integer)
30190     Call RefreshBlockCombo
End Sub

Private Sub gridColors_Click()
30200 On Error GoTo ERR_gridColors_Click
      '    frmAdditionalActions.Text(iBlock).Text = frmAdditionalActions.Text(iBlock).Text & gridColors.Text & ", "
          Dim n As Integer
          Dim iRow As Integer
          Dim iCol As Integer

30210     If gridColors.Text = "" Then Exit Sub
          
          'not a color column:
      '    If gridColors.col Mod 2 <> 0 Then Exit Sub
              
              
          'for a block:
30220     If iSlide = 0 Then
      '        Dim strColor As String
      '        If gridColors.col Mod 2 <> 0 Then
      '            gridColors.col = gridColors.col - 1
      '            strColor = gridColors.TextMatrix(gridColors.row, gridColors.col - 1)
      '        Else
      '            strColor = gridColors.TextMatrix(gridColors.row, gridColors.col)
      '        End If
              
              'if this color already exists, remove from the list;
              'else, add it to the list
              'n = ExistInCombo(gridColors.Text)
              
              'If n > -1 Then
              
              'for clicking the combo column:
30230         If gridColors.col Mod 2 <> 0 Then
30240             gridColors.col = gridColors.col - 1
30250             Call ShowNumOfSlidesCombo(True)
30260             Exit Sub
30270         End If
              
                      
              'for clicking a color column:
              'what happens when selecting (else) / unselecting (if) a color
30280         If dicBlockColors.Exists(gridColors.Text) Then
                  'Call frmAdditionalActions.CmbBlockEntry(iBlock).RemoveItem(n)
30290             gridColors.CellBackColor = vbWhite
                  
30300             Call ShowNumOfSlidesCombo(False)
30310             Call dicBlockColors.Remove(gridColors.Text)
                  
                  'If frmAdditionalActions.CmbBlockEntry(iBlock).ListCount = 0 Then
30320             If dicBlockColors.Count = 0 Then
30330                 frmAdditionalActions.CmbBlockEntry(iBlock).BackColor = _
    vbWhite
30340                 Call frmAdditionalActions.DecrementColorsBlocks
30350             End If
30360         Else
                  'Call frmAdditionalActions.CmbBlockEntry(iBlock).AddItem(gridColors.Text & "  #1")
                 '       gridColors.TextMatrix(gridColors.row, gridColors.col + 1))
                 ' Call frmAdditionalActions.CmbBlockEntry(iBlock).AddItem(gridColors.Text)
30370             gridColors.CellBackColor = MARK_SELECTED
                  
30380             Call ShowNumOfSlidesCombo(True)
30390             Call dicBlockColors.Add(gridColors.Text, "1")
                  
30400             If frmAdditionalActions.CmbBlockEntry(iBlock).BackColor <> _
    MARK_SELECTED Then
30410                 frmAdditionalActions.CmbBlockEntry(iBlock).BackColor = _
    MARK_SELECTED
30420                 Call frmAdditionalActions.IncrementColorsBlocks
30430             End If
                  
                  
30440         End If
          
          'for a slide:
          'what happens where selecting (else) / unselecting (if) a color:
30450     Else
30460         If gridColors.Text = frmAdditionalActions.txtSlide(iSlide).Text _
    Then
30470             frmAdditionalActions.txtSlide(iSlide).Text = ""
30480             Call frmAdditionalActions.DecrementColorsSlides
30490             gridColors.CellBackColor = vbWhite
30500         Else
30510             If frmAdditionalActions.txtSlide(iSlide).Text = "" Then
30520                 Call frmAdditionalActions.IncrementColorsSlides
30530             Else
30540                 iRow = gridColors.row
30550                 iCol = gridColors.col
                      
30560                 gridColors.row = iSlideColorRow
30570                 gridColors.col = iSlideColorCol
30580                 gridColors.CellBackColor = vbWhite
                      
30590                 gridColors.row = iRow
30600                 gridColors.col = iCol
30610                 gridColors.CellBackColor = MARK_SELECTED
30620             End If
                  
30630             iSlideColorRow = gridColors.row
30640             iSlideColorCol = gridColors.col
30650             frmAdditionalActions.txtSlide(iSlide).Text = gridColors.Text
30660             gridColors.CellBackColor = MARK_SELECTED
30670         End If
30680     End If
          
          
          
30690     Exit Sub
ERR_gridColors_Click:
30700 MsgBox "ERR_gridColors_Click" & vbCrLf & Err.Description
End Sub


'01.11.2006: show the list to select number
'of slides for this stain:
'grid row & col should be of the color cell to the left of the couter cell
Private Sub ShowNumOfSlidesCombo(b As Boolean)
30710 On Error GoTo ERR_ShowNumOfSlidesCombo
          
30720     Call InitComboNumOfSlides
          
30730     cmbNumOfSlides.Left = gridColors.CellLeft + gridColors.Left + _
    gridColors.CellWidth
30740     cmbNumOfSlides.Top = gridColors.CellTop + gridColors.Top + _
    (gridColors.CellHeight - cmbNumOfSlides.Height) / 2
      '    cmbNumOfSlides.Height = gridColors.CellHeight
30750     cmbNumOfSlides.Width = gridColors.Width / (gridColors.Cols * 2)
30760     cmbNumOfSlides.Visible = b
          
30770     If b = True Then
30780         If gridColors.TextMatrix(gridColors.row, gridColors.col + 1) <> _
    "" Then
30790             cmbNumOfSlides.Text = gridColors.TextMatrix(gridColors.row, _
    gridColors.col + 1)
30800         Else
30810             gridColors.TextMatrix(gridColors.row, gridColors.col + 1) = _
    cmbNumOfSlides.Text
30820         End If
30830     Else
30840         gridColors.TextMatrix(gridColors.row, gridColors.col + 1) = ""
30850     End If
          
      '    If b = True Then
      '        If gridColors.TextMatrix(gridColors.row, gridColors.col + 1) <> "" Then
      '            cmbNumOfSlides.Text = gridColors.TextMatrix(gridColors.row, gridColors.col)
      '        Else
      '            gridColors.TextMatrix(gridColors.row, gridColors.col + 1) = cmbNumOfSlides.Text
      '        End If
      '    Else
      ''        If gridColors.col Mod 2 <> 0 Then
      '            gridColors.TextMatrix(gridColors.row, gridColors.col) = ""
      ''        Else
      ''            gridColors.TextMatrix(gridColors.row, gridColors.col + 1) = ""
      ''        End If
      '    End If
          

30860     Exit Sub
ERR_ShowNumOfSlidesCombo:
30870 MsgBox "ERR_ShowNumOfSlidesCombo" & vbCrLf & Err.Description
End Sub


'if exists - return the index of this item in the list
'otherwise - return -1
Private Function ExistInCombo(strColor As String) As Integer
30880 On Error GoTo ERR_ExistInCombo
          Dim i As Integer
          Dim strSelectedColor As String
          
30890     ExistInCombo = -1
          
30900     For i = 0 To frmAdditionalActions.CmbBlockEntry(iBlock).ListCount - 1
30910         strSelectedColor = _
    frmAdditionalActions.CmbBlockEntry(iBlock).list(i)
30920         strSelectedColor = Mid(strSelectedColor, 1, InStr(1, _
    strSelectedColor, "#") - 3)
              
30930         If strSelectedColor = strColor Then
30940             ExistInCombo = i
30950             Exit For
30960         End If
30970     Next i
          
30980     Exit Function
ERR_ExistInCombo:
30990 MsgBox "ERR_ExistInCombo" & vbCrLf & Err.Description
End Function


Private Sub InitGrid()
31000 On Error GoTo ERR_InitGrid

          Dim i As Integer
          Dim j As Integer

31010     nGridRows = ComputNumberOfRows

31020     cmbNumOfSlides.Visible = False

31030     gridColors.Rows = nGridRows 'GRID_ROWS
31040     gridColors.Cols = GRID_COLS
              
31050     For i = 0 To gridColors.Rows - 1
31060         gridColors.row = i
31070         gridColors.RowHeight(i) = gridColors.Height / gridColors.Rows
              
31080         For j = 0 To gridColors.Cols - 1
31090             gridColors.ColWidth(j) = gridColors.Width / gridColors.Cols
                  
31100             If j Mod 2 <> 0 Then
31110                 gridColors.ColWidth(j) = gridColors.ColWidth(j) - _
    gridColors.ColWidth(j) / 2
31120             Else
31130                 gridColors.ColWidth(j) = gridColors.ColWidth(j) + _
    gridColors.ColWidth(j) / 2
31140             End If
              
              
31150             gridColors.col = j
                  
                  
31160             gridColors.ColAlignment(j) = flexAlignLeftCenter
               
                  'gridColors.CellAlignment = vbLeftJustify
                  'gridColors.CellAlignment = flexAlignLeftCenter
                  
      '            gridColors.Text = "Color"
                  
31170         Next j
31180     Next i

         ' gridColors.CellAlignment = flexAlignCenterCenter

31190     Exit Sub
ERR_InitGrid:
31200 MsgBox "ERR_InitGrid" & vbCrLf & Err.Description
End Sub

Private Sub InitGridColors(strColorType As String)
31210 On Error GoTo ERR_InitGridColors

          Dim i As Integer
          Dim iRow As Integer
          Dim iCol As Integer
          Dim s As String
          Dim k As Integer

          Dim dicColors As New Dictionary
          
31220     cmbNumOfSlides.Visible = False
          
31230     gridColors.Clear

31240     Select Case strColorType
      '        Case "מולקולרית"
      '            Set dicColors = dicMolecularStains
              Case "אימונוהיסטוכימיה"
31250             Set dicColors = dicImonohistochemistryStains
31260         Case "היסטוכימיה"
31270             Set dicColors = dicHistochemistryStains
      '        Case "מיוחדות"
      '            Set dicColors = dicSpecialStains
31280         Case "אחר"
31290             Set dicColors = dicOtherStainOptions
31300     End Select

31310     iRow = -1
31320     iCol = 0
          
31330     For i = 0 To dicColors.Count - 1
31340         iRow = iRow + 1
31350         If iRow = nGridRows Then 'GRID_ROWS Then
31360             iRow = 0
31370             iCol = iCol + 2
                  'iCol = iCol + 1
31380         End If
              
              
31390         gridColors.row = iRow
31400         gridColors.col = iCol
31410         gridColors.Text = dicColors.Keys(i)
      '        gridColors.TextMatrix(iRow, iCol) = dicColors.Keys(i)
              
              'check if the color is already selected for the slide / block:
31420         If iSlide = 0 Then
31430             If dicBlockColors.Exists(dicColors.Keys(i)) Then
31440                 gridColors.CellBackColor = MARK_SELECTED
                      
31450                 gridColors.TextMatrix(iRow, iCol + 1) = _
    dicBlockColors(dicColors.Keys(i))
31460             End If
31470         Else
31480             If frmAdditionalActions.txtSlide(iSlide).Text = _
    dicColors.Keys(i) Then
31490                 gridColors.CellBackColor = MARK_SELECTED
                  
31500                 iSlideColorRow = iRow
31510                 iSlideColorCol = iCol
31520             End If
31530         End If
              
31540     Next i

      '    If iSlide = 0 Then
      '        'read the selected colors of the relevant block;
      '        'for each color, get it's index in the current list of colors (color type) presented;
      '        'from the index, compute the row & col in the grid to mark it:
      '        For i = 0 To frmAdditionalActions.CmbBlockEntry(iBlock).ListCount - 1
      '            s = frmAdditionalActions.CmbBlockEntry(iBlock).list(i)
      '
      '            'cut the color counter
      '            s = Mid(s, 1, InStr(1, s, "#") - 3)
      '
      '            If dicColors.Exists(s) Then
      '                k = dicColors(s)
      '                gridColors.row = k Mod GRID_ROWS
      '                gridColors.col = Devide(k, GRID_ROWS)
      '                gridColors.CellBackColor = MARK_SELECTED
      '            End If
      '        Next i
      '    Else
      '        s = frmAdditionalActions.txtSlide(iSlide).Text
      '        If dicColors.Exists(s) Then
      '            k = dicColors(s)
      '            gridColors.row = k Mod GRID_ROWS
      '            gridColors.col = Devide(k, GRID_ROWS)
      '            gridColors.CellBackColor = MARK_SELECTED
      '
      '            iSlideColorRow = gridColors.row
      '            iSlideColorCol = gridColors.col
      '        End If
      '    End If

31550     Exit Sub
ERR_InitGridColors:
31560 MsgBox "ERR_InitGridColors" & vbCrLf & Err.Description
End Sub

'devide two integers; get the result rounder
'to the smaller integer (like in C...):
Private Function Devide(X As Integer, Y As Integer) As Integer
          Dim d As Double
          Dim n As Integer
          
31570     d = X / Y
31580     n = X / Y
          
31590     If d < n Then
31600         n = n - 1
31610     End If
          
31620     Devide = n
End Function

Private Sub InitComboNumOfSlides()
31630 On Error GoTo ERR_InitComboNumOfSlides
          Dim i As Integer

31640     Call cmbNumOfSlides.Clear

31650     For i = 1 To 10
31660         cmbNumOfSlides.list(cmbNumOfSlides.ListCount) = CStr(i)
31670     Next i
          
31680     cmbNumOfSlides.Text = cmbNumOfSlides.list(0)

31690     Exit Sub
ERR_InitComboNumOfSlides:
31700 MsgBox "ERR_InitComboNumOfSlides" & vbCrLf & Err.Description
End Sub

Private Sub InitBlockColors(iBlock As Integer)
31710 On Error GoTo ERR_InitBlockColors
          Dim sColor As String
          Dim sCount As String
          Dim i As Integer
          Dim s As String

31720     If iBlock = 0 Then Exit Sub

31730     Call dicBlockColors.RemoveAll
          
31740     For i = 0 To frmAdditionalActions.CmbBlockEntry(iBlock).ListCount - 1
31750         s = frmAdditionalActions.CmbBlockEntry(iBlock).list(i)
              
              'cut the color string to color / count:
31760         Call ParseBlockColor(s, sColor, sCount)
31770         Call dicBlockColors.Add(sColor, sCount)
          
31780     Next i

31790     Exit Sub
ERR_InitBlockColors:
31800 MsgBox "ERR_InitBlockColors" & vbCrLf & Err.Description
End Sub


'when the color name is formatted this way: color_name  #count
's(in)       - the original string
'sColor(out) - then color name
'sCount(out) - the count
Private Sub ParseBlockColor(s As String, sColor As String, sCount As String)
31810 On Error GoTo ERR_ParseBlockColor
          Dim i As Integer
          
31820     i = InStr(1, s, "#")
31830     If i <> 0 Then
31840         sColor = Mid(s, 1, i - 3)
31850         sCount = Mid(s, i + 1)
31860     Else
31870         sColor = s
31880         sCount = "0"
31890     End If

31900     Exit Sub
ERR_ParseBlockColor:
31910 MsgBox "ERR_ParseBlockColor" & vbCrLf & Err.Description
End Sub

Private Sub InitColorList()
31920     Call lstColorType.Clear
      '    Call lstColorType.AddItem("מיוחדות")
31930     Call lstColorType.AddItem("אימונוהיסטוכימיה")
31940     Call lstColorType.AddItem("היסטוכימיה")
      '    Call lstColorType.AddItem("מולקולרית")
31950     Call lstColorType.AddItem("אחר")
31960     lstColorType.Text = lstColorType.list(0)
End Sub

Private Sub lstColorType_Click()
31970     If Not isGridInitialized Then Exit Sub
          
          'lstColorType is defined in the form
31980     Call InitGridColors(lstColorType.Text)
End Sub


Private Function ComputNumberOfRows() As Integer
31990 On Error GoTo ERR_ComputNumberOfRows
          Dim i As Integer
          Dim nMax As Integer
          Dim nTemp As Integer
          Dim dSizes As New Dictionary
          
32000     dSizes(0) = dicMolecularStains.Count
32010     dSizes(1) = dicImonohistochemistryStains.Count
32020     dSizes(2) = dicSpecialStains.Count
32030     dSizes(3) = dicOtherStainOptions.Count
          
32040     nMax = 0
32050     For i = 0 To dSizes.Count - 1
32060         nTemp = dSizes(i)
32070         If nTemp > nMax Then
32080             nMax = nTemp
32090         End If
32100     Next i
          
32110     For i = 1 To nMax
32120         If i * GRID_COLS / 2 >= nMax Then
32130             Exit For
32140         End If
32150     Next i

32160     ComputNumberOfRows = i

32170     Exit Function
ERR_ComputNumberOfRows:
32180 MsgBox "ERR_ComputNumberOfRows" & vbCrLf & Err.Description
End Function
