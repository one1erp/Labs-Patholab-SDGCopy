VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public strName As String
Public strId As String
Public strWorkflowName As String

Public dicBlocks As Dictionary


Private Sub Class_Initialize()
'    Set dicBlocks = New Dictionary
End Sub

Public Sub Initialize(sampleName As String, sampleId As String)
13350     strName = sampleName
13360     strId = sampleId
          
13370     Set dicBlocks = New Dictionary
          
13380 If DEFINE_DEBUG Then MsgBox "Sample: " & strName & " " & strId

13390     Call InitWorkflowName
13400     Call FillBlockList
End Sub


Private Sub InitWorkflowName()
13410     On Error GoTo ERR_InitWorkflowName
              
          Dim sql As String
          Dim rs As Recordset
              
              
13420     sql = "   select w.NAME  from"
13430     sql = sql & "  lims_sys.workflow_node wn, lims_sys.workflow w , lims_sys.sample s"
13440     sql = sql & "  where s.WORKFLOW_NODE_ID = wn.WORKFLOW_NODE_ID"
13450     sql = sql & "  and wn.WORKFLOW_ID=w.WORKFLOW_ID"
13460     sql = sql & "  and s.sample_id = " & strId
          
13470     Set rs = aConnection.Execute(sql)
          
13480     If Not rs.EOF Then
13490         strWorkflowName = nte(rs("NAME"))
13500     End If
              
13510     Exit Sub
ERR_InitWorkflowName:
13520 MsgBox "Error on line:" & Erl & " in InitWorkflowName" & vbCrLf & Err.Description
End Sub

Private Sub FillBlockList()
13530 On Error GoTo ERR_FillBlockList

          Dim block As CBlock
          Dim rs As Recordset
          
13540     Set rs = aConnection.Execute _
                (" select a.name, a.aliquot_id " & _
                 " from lims_sys.aliquot a " & _
                 " where sample_id = " & strId & _
                 " and not exists " & _
                 " (select aliquot_id from lims_sys.aliquot_formulation " & _
                 " where child_aliquot_id = a.aliquot_id )" & _
                 " order by a.aliquot_id " _
                 )
                 
13550     While Not rs.EOF
13560         Set block = New CBlock
13570         Call block.Initialize(nte(rs("name")), nte(rs("aliquot_id")))
              
13580         Call dicBlocks.Add(dicBlocks.Count, block)
13590         rs.MoveNext
13600     Wend
          
13610     Exit Sub
ERR_FillBlockList:
13620 MsgBox "FillBlockList" & vbCrLf & Err.Description
End Sub
