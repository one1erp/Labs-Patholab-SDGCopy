VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public strName As String
Public strId As String
Public strWorkflowName As String

Public Sub Initialize(resultName As String, resultId As String)
12330     strName = resultName
12340     strId = resultId
          
12350 If DEFINE_DEBUG Then MsgBox "Result: " & strName & " " & strId
          
12360     Call InitWorkflowName
End Sub


Private Sub InitWorkflowName()
12370     On Error GoTo ERR_InitWorkflowName
              
          Dim sql As String
          Dim rs As Recordset
              
          'get the name of the event that created this result:
12380     sql = " select name from lims_sys.workflow_node"
12390     sql = sql & " where workflow_node_id = "
12400     sql = sql & " ("
12410     sql = sql & "    select wn.PARENT_ID"
12420     sql = sql & "    from  lims_sys.workflow_node wn, lims_sys.result r"
12430     sql = sql & "    where r.WORKFLOW_NODE_ID = wn.WORKFLOW_NODE_ID"
12440     sql = sql & "    and r.result_id = " & strId
12450     sql = sql & " )"
          
12460     Set rs = aConnection.Execute(sql)
          
12470     If Not rs.EOF Then
12480         strWorkflowName = nte(rs("NAME"))
12490     End If
              
12500     Exit Sub
ERR_InitWorkflowName:
12510 MsgBox "Error on line:" & Erl & " in InitWorkflowName" & vbCrLf & Err.Description
End Sub
