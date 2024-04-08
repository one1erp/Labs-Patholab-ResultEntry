VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WFResultEntryCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements LSEXT.IWorkflowExtension

Private aConnection As New ADODB.Connection
Private NtlsCon As LSSERVICEPROVIDERLib.NautilusDBConnection
Private sp As LSSERVICEPROVIDERLib.NautilusServiceProvider

Private Sub IWorkflowExtension_Execute(Parameters As LSEXT.IExtensionParametersEx)
    Dim RecordRs As ADODB.Recordset
    Dim aDatasource As String
    Dim strSQL As String

    Set sp = Parameters("SERVICE_PROVIDER")

    Set NtlsCon = sp.QueryServiceProvider("DBConnection")
    aDatasource = NtlsCon.GetADOConnectionString
    aConnection.Open aDatasource
    aConnection.CursorLocation = adUseClient

    Set RecordRs = Parameters("RECORDS")

    If Not RecordRs.EOF Then
        Call OpenResultEntryForm(nte(RecordRs("NAME")))
    End If

    aConnection.Close
End Sub

Private Sub OpenResultEntryForm(RequestNBR As String)
    Dim strParameters As String
    Dim ParamRec As ADODB.Recordset
    Dim frs As New FrmResultEntry

    If Trim(RequestNBR) = "" Then Exit Sub

     Load frs
     frs.ResultEntryCtrl.RunFromWindow = True

     Call frs.ResultEntryCtrl.IExtensionWindow_SetServiceProvider(sp)
     frs.ResultEntryCtrl.IExtensionWindow_Internationalise

     Set ParamRec = aConnection.Execute("select PARAMETER2 from lims_sys.command " & _
         "where name = 'Result Entry'")

     If Not ParamRec.EOF Then
         strParameters = Trim(nte(ParamRec("PARAMETER2")))
         Call frs.ResultEntryCtrl.IExtensionWindow_SetParameters(strParameters)
     End If

     ParamRec.Close

     frs.ResultEntryCtrl.IExtensionWindow_PreDisplay
     frs.ResultEntryCtrl.IExtensionWindow_GetButtons
     frs.ResultEntryCtrl.IExtensionWindow_Setup

     frs.ResultEntryCtrl.InitiateSdg (RequestNBR)
     frs.Show vbModal

     frs.ResultEntryCtrl.IExtensionWindow_CloseQuery
     Unload frs
     Set frs = Nothing
End Sub

Private Function nte(e As Variant) As Variant
    nte = IIf(IsNull(e), "", e)
End Function



