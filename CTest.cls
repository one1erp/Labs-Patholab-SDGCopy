VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public strName As String
Public strId As String
Public strWorkflowName As String

Public dicResults As Dictionary


Private Sub Class_Initialize()
'    Set dicResults = New Dictionary
End Sub


Public Sub Initialize(testName As String, testId As String)
12520     Set dicResults = New Dictionary
          
12530     strName = testName
12540     strId = testId
          
12550 If DEFINE_DEBUG Then MsgBox "Test: " & strName & " " & strId
          
12560     Call InitWorkflowName
12570     Call FillResultList
End Sub


Private Sub InitWorkflowName()
12580     On Error GoTo ERR_InitWorkflowName
              
          Dim sql As String
          Dim rs As Recordset
              
12590     sql = "   select w.NAME  from"
12600     sql = sql & "  lims_sys.workflow_node wn, lims_sys.workflow w , lims_sys.test t"
12610     sql = sql & "  where t.WORKFLOW_NODE_ID = wn.WORKFLOW_NODE_ID"
12620     sql = sql & "  and wn.WORKFLOW_ID=w.WORKFLOW_ID"
12630     sql = sql & "  and t.test_id = " & strId
          
12640     Set rs = aConnection.Execute(sql)
          
12650     If Not rs.EOF Then
12660         strWorkflowName = nte(rs("NAME"))
12670     End If
              
12680     Exit Sub
ERR_InitWorkflowName:
12690 MsgBox "Error on line:" & Erl & " in InitWorkflowName" & vbCrLf & Err.Description
End Sub



Private Sub FillResultList()
12700 On Error GoTo ERR_FillResultList
          
          Dim result As CResult
          Dim rs As Recordset
          
12710     Set rs = aConnection.Execute( _
                    " select name, result_id " & _
                    " from lims_sys.result " & _
                    " where test_id = " & strId & _
                    " order by result_id " _
                    )
                    
12720     While Not rs.EOF
12730         Set result = New CResult
12740         Call result.Initialize(nte(rs("name")), nte(rs("result_id")))
              
12750         Call dicResults.Add(dicResults.Count, result)
12760         rs.MoveNext
12770     Wend
                     
12780     Exit Sub
ERR_FillResultList:
12790 MsgBox "Error on line:" & Erl & " in FillResultList" & vbCrLf & Err.Description
End Sub
