VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSDGCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Implements LSEXT.IGenericExtension

Option Explicit

Private Sub IGenericExtension_Execute(Parameters As LSEXT.IExtensionParametersEx)
10    On Error GoTo ERR_IGenericExtension_Execute

          Dim i As Integer
          Dim XmlDoc As New DOMDocument
          Dim Xmlres As New DOMDocument
          Dim sp As LSSERVICEPROVIDERLib.NautilusServiceProvider
          Dim Errs As IXMLDOMNodeList
          Dim constr As String
          
      '    Dim SdgRS As ADODB.Recordset
          
20        Set sp = Parameters("SERVICE_PROVIDER")
          
30         Set NtlsCon = sp.QueryServiceProvider("DBConnection")
40         Set aConnection = Nothing
50         Set aConnection = New ADODB.Connection
60        constr = "Provider=OraOLEDB.Oracle" & _
              ";Data Source=" & NtlsCon.GetServerDetails & _
              ";User ID=" & NtlsCon.GetUsername & _
              ";Password=" & NtlsCon.GetPassword
70                     If NtlsCon.GetServerIsProxy Then
80          constr = "Provider=OraOLEDB.Oracle;Data Source=" & _
            NtlsCon.GetServerDetails & ";User id=/;Persist Security Info=True;"
90        End If
100       aConnection.Open constr
110       aConnection.CursorLocation = adUseServer
120        aConnection.Execute "SET ROLE LIMS_USER"
130        Call ConnectSameSession(CDbl(NtlsCon.GetSessionId))
           
140        Set ProcessXML = sp.QueryServiceProvider("ProcessXML")
150        sdg_id = -1
160        sdg_name = -1
            
           'Call InitializeStopListFields
           
170        Call frmMain.Initialize(sp)
180        Call frmMain.Show(vbModal)
                       
      '     Set SdgRS = Parameters("RECORDS")
      '     SdgRS.MoveLast
      '     If Not SdgRS.EOF Then
      '       sdg_id = SdgRS("SDG_ID")
      '       sdg_name = SdgRS("NAME")
      '       Call runSDGCopy
      '     End If
      '     SdgRS.Close

190       Exit Sub
ERR_IGenericExtension_Execute:
200   MsgBox "Error on line:" & Erl & " in IGenericExtension_Execute" & vbCrLf & Err.Description
End Sub

Private Sub ConnectSameSession(ByVal aSessionID)
          Dim aProc As New ADODB.Command
          Dim aSession As New ADODB.Parameter
          
180       aProc.ActiveConnection = aConnection
190       aProc.CommandText = "lims.lims_env.connect_same_session"
200       aProc.CommandType = adCmdStoredProc

210       aSession.Type = adDouble
220       aSession.Direction = adParamInput
230       aSession.Value = aSessionID
240       aProc.Parameters.Append aSession

250       aProc.Execute
260       Set aSession = Nothing
270       Set aProc = Nothing
End Sub


Private Sub InitializeStopListFields()
280       Set dicStopListFields = New Dictionary
          
290       Call dicStopListFields.Add("name", "name")
300       Call dicStopListFields.Add("sdg_id", "sdg_id")
310       Call dicStopListFields.Add("sample_id", "sample_id")
320       Call dicStopListFields.Add("aliquot_id", "aliquot_id")
330       Call dicStopListFields.Add("test_id", "test_id")
340       Call dicStopListFields.Add("result_id", "result_id")
350       Call dicStopListFields.Add("status", "status")
360       Call dicStopListFields.Add("authorised_on", "authorised_on")
370       Call dicStopListFields.Add("authorised_by", "authorised_by")
380       Call dicStopListFields.Add("workflow_node_id", "workflow_node_id")
390       Call dicStopListFields.Add("sdg_template_id", "sdg_template_id")
400       Call dicStopListFields.Add("sample_template_id", "sample_template_id")
410       Call dicStopListFields.Add("aliquot_template_id", "aliquot_template_id")
420       Call dicStopListFields.Add("test_template_id", "test_template_id")
430       Call dicStopListFields.Add("result_template_id", "result_template_id")
          
      '--------------------

440       Call dicStopListFields.Add("old_status", "old_status")
450       Call dicStopListFields.Add("completed_on", "completed_on")
460       Call dicStopListFields.Add("completed_by", "completed_by")
470       Call dicStopListFields.Add("previous_sample", "previous_sample")
480       Call dicStopListFields.Add("plate_aliquot_type", "plate_aliquot_type")
490       Call dicStopListFields.Add("container_type_id", "container_type_id")
End Sub
