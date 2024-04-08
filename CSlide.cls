VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSlide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public strName As String
Public strId As String

Public dicTests As Dictionary


Private Sub Class_Initialize()
'    Set dicTests = New Dictionary
End Sub


Public Sub Initialize(slideName As String, slideId As String)
12800     Set Me.dicTests = New Dictionary
          
12810     strName = slideName
12820     strId = slideId
          
12830 If DEFINE_DEBUG Then MsgBox "Slide: " & strName & " " & strId
          
12840     Call FillTestList
End Sub

Private Sub FillTestList()
12850 On Error GoTo ERR_FillTestsList

          Dim test As CTest
          Dim rs As Recordset
          
12860     Set rs = aConnection.Execute( _
                   " select name, test_id " & _
                   " from lims_sys.test " & _
                   " where aliquot_id = " & strId & _
                   " order by test_id " _
                   )
          
12870     While Not rs.EOF
12880         Set test = New CTest
12890         Call test.Initialize(nte(rs("name")), nte(rs("test_id")))
              
12900         Call Me.dicTests.Add(dicTests.Count, test)
12910         rs.MoveNext
12920     Wend
          
12930     Exit Sub
ERR_FillTestsList:
12940 MsgBox "Error on line:" & Erl & " in FillTestsList" & vbCrLf & Err.Description

End Sub
