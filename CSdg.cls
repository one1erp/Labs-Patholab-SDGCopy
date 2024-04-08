VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSdg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public strName As String
Public strId As String
Public strExternalReference As String
Public strWorkflowName As String

Public dicSamples As Dictionary


Private Sub Class_Initialize()
'    strId = CStr(sdg_id)
'    strName = sdg_name
    
'    Set dicSamples = New Dictionary
    
'If DEFINE_DEBUG Then MsgBox "sdg: " & strId & " " & strName
End Sub

Public Sub Initialize(sdgName As String, sdgId As String)
13630     strId = sdgId
13640     strName = sdgName
          
13650     Set dicSamples = New Dictionary
          
13660 If DEFINE_DEBUG Then MsgBox "sdg: " & strId & " " & strName

13670     Call GetExternalReferenceAndWorkflowName
13680     Call FillSampleList
End Sub

Private Sub GetExternalReferenceAndWorkflowName()
13690     On Error GoTo ERR_GetExternalReferenceAndWorkflowName
              
          Dim sql As String
          Dim rs As Recordset
              
13700     sql = " select w.NAME, d.EXTERNAL_REFERENCE from"
13710     sql = sql & " lims_sys.workflow_node wn, lims_sys.workflow w , lims_sys.sdg d"
13720     sql = sql & " where d.WORKFLOW_NODE_ID = wn.WORKFLOW_NODE_ID"
13730     sql = sql & " and wn.WORKFLOW_ID=w.WORKFLOW_ID"
13740     sql = sql & " and d.SDG_ID =  " & strId
          
13750     Set rs = aConnection.Execute(sql)
          
13760     If Not rs.EOF Then
13770         strExternalReference = nte(rs("EXTERNAL_REFERENCE"))
13780         strWorkflowName = nte(rs("NAME"))
13790     End If
              
13800     Exit Sub
ERR_GetExternalReferenceAndWorkflowName:
13810 MsgBox "Error on line:" & Erl & " in GetExternalReferenceAndWorkflowName" & vbCrLf & Err.Description
End Sub


Private Sub FillSampleList()
          Dim rs As Recordset
          Dim sample As CSample
          
13820     Set rs = aConnection.Execute _
                       (" select name, sample_id " & _
                        " from lims_sys.sample " & _
                        " where sdg_id = " & strId & _
                        " order by sample_id ")
                        
13830     While Not rs.EOF
13840         Set sample = New CSample
13850         Call sample.Initialize(nte(rs("name")), nte(rs("sample_id")))
              
13860         Call dicSamples.Add(dicSamples.Count, sample)
              
13870         rs.MoveNext
13880     Wend
End Sub
