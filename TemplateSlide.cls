VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TemplateSlide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private strColor As String
Private strLayers As String

Public Sub Initialize(strColor_ As String, strLayers_ As String)
10        strColor = strColor_
20        strLayers = strLayers_
End Sub

Public Function GetColor() As String
30        GetColor = strColor
End Function

Public Function GetLayers() As String
40        GetLayers = strLayers
End Function
