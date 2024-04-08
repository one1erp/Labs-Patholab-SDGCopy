VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public strName As String
Public strId As String
Public strWorkflowName As String

Public dicSlides As Dictionary
Public dicTests As Dictionary

Private Sub Class_Initialize()
'    Set dicSlides = New Dictionary
'    Set dicTests = New Dictionary
End Sub


Public Sub Initialize(blockName As String, blockId As String)
12950     strName = blockName
12960     strId = blockId
          
12970     Set dicSlides = New Dictionary
12980     Set Me.dicTests = New Dictionary

12990 If DEFINE_DEBUG Then MsgBox "Block: " & strName & " " & strId

13000     Call InitWorkflowName
13010     Call FillSlidesList
13020     Call FillTestsList
End Sub

Private Sub InitWorkflowName()
13030     On Error GoTo ERR_InitWorkflowName
              
          Dim sql As String
          Dim rs As Recordset
              
13040     sql = "   select w.NAME  from"
13050     sql = sql & "  lims_sys.workflow_node wn, lims_sys.workflow w , lims_sys.aliquot a"
13060     sql = sql & "  where a.WORKFLOW_NODE_ID = wn.WORKFLOW_NODE_ID"
13070     sql = sql & "  and wn.WORKFLOW_ID=w.WORKFLOW_ID"
13080     sql = sql & "  and a.aliquot_id = " & strId
          
13090     Set rs = aConnection.Execute(sql)
          
13100     If Not rs.EOF Then
13110         strWorkflowName = nte(rs("NAME"))
13120     End If
              
13130     Exit Sub
ERR_InitWorkflowName:
13140 MsgBox "Error on line:" & Erl & " in InitWorkflowName" & vbCrLf & Err.Description
End Sub


Private Sub FillSlidesList()
13150 On Error GoTo ERR_FillSlidesList

          Dim slide As CSlide
          Dim rs As Recordset
          
13160     Set rs = aConnection.Execute( _
                   " select a.name, a.aliquot_id " & _
                   " from lims_sys.aliquot a " & _
                   " where a.aliquot_id in " & _
                   " ( " & _
                   " select child_aliquot_id from lims_sys.aliquot_formulation " & _
                   " where parent_aliquot_id = " & strId & _
                   " ) " & _
                   " order by a.aliquot_id " _
                   )
          
13170     While Not rs.EOF
13180         Set slide = New CSlide
13190         Call slide.Initialize(nte(rs("name")), nte(rs("aliquot_id")))
              
13200         Call dicSlides.Add(dicSlides.Count, slide)
              
13210         rs.MoveNext
13220     Wend
                   
13230     Exit Sub
ERR_FillSlidesList:
13240 MsgBox "Error on line:" & Erl & " in FillSlidesList" & vbCrLf & Err.Description
End Sub


Private Sub FillTestsList()
13250 On Error GoTo ERR_FillTestsList

          Dim test As CTest
          Dim rs As Recordset
          
13260     Set rs = aConnection.Execute( _
                   " select name, test_id " & _
                   " from lims_sys.test " & _
                   " where aliquot_id = " & strId & _
                   " order by test_id " _
                   )
          
13270     While Not rs.EOF
13280         Set test = New CTest
13290         Call test.Initialize(nte(rs("name")), nte(rs("test_id")))
              
13300         Call Me.dicTests.Add(dicTests.Count, test)
13310         rs.MoveNext
13320     Wend
          
13330     Exit Sub
ERR_FillTestsList:
13340 MsgBox "Error on line:" & Erl & " in FillTestsList" & vbCrLf & Err.Description
End Sub









