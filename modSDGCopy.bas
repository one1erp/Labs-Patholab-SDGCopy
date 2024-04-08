Attribute VB_Name = "modSDGCopy"
Option Explicit
 
Const BLOCK_SIZE = 1024

Private sdg_log As New SdgLog.CreateLog
Private sdg_log_desc As String

Public Function SetOldName(name As String) As String
          Dim sql As String
          Dim maxName As String
          Dim maxNameRs As ADODB.Recordset
          Dim i, j As Double
          Dim s As String
          
500       maxName = ""
510       sql = "Select max(name) as max_name from LIMS_SYS.sdg where name like '" & name & "%'"
520       Set maxNameRs = aConnection.Execute(sql)
530       If Not maxNameRs.EOF Then
540         maxName = maxNameRs("max_name")
550           j = InStr(1, maxName, "V")
560           If j > 0 Then
570             s = Mid(maxName, j + 1, 255)
580             If checkNum(s) Then
590               i = s + 1
600             Else
610               i = 1
620             End If
630             VersionName = "V" & i
640             maxName = Mid(maxName, 1, j - 1) & "V" & i
650           Else
660             VersionName = "V1"
670             maxName = maxName & "V1"
680           End If
690       End If
700       SetOldName = maxName
710       maxNameRs.Close
End Function

Public Function CheckFieldOK(fieldName As String) As Boolean
720     If fieldName = "STATUS" Or _
          fieldName = "OLD_STATUS" Or _
          fieldName = "WORKFLOW_NODE_ID" Or _
          fieldName = "PREVIOUS_SAMPLE" Or _
          fieldName = "PLATE_ALIQUOT_TYPE" Or _
          fieldName = "CONTAINER_TYPE_ID" Or _
          fieldName = "SAMPLE_TEMPLATE_ID" Or _
          fieldName = "ALIQUOT_TEMPLATE_ID" Then
730           CheckFieldOK = False
740     Else
750       If Len(fieldName) > 3 Then
760         If Mid(fieldName, Len(fieldName) - 2, 3) = "_BY" _
              Or Mid(fieldName, Len(fieldName) - 2, 3) = "_ON" Then
770             CheckFieldOK = False
780           Else
790             CheckFieldOK = True
800         End If
810       Else
820         CheckFieldOK = True
830       End If
840     End If
End Function

'changes the value in a link field
'when writing the xml file the NAME should be taken instead of the ID
Private Function ManipulateLinkField(strFieldName As String, _
                                     vFieldValue As Variant, _
                                     strTableName As String) As String

850   On Error GoTo ERR_ManipulateLinkField
          
          Dim rs As Recordset
          Dim sql As String
          Dim strPointedTableName As String
          Dim strFieldValue As String
          
860       ManipulateLinkField = ""
          
870       If Mid(strFieldName, 1, 2) = "U_" Or Mid(strFieldName, 1, 2) = "u_" Then
880           strTableName = strTableName & "_user"
890       End If
          
      '   If isUserTable = True Then
      '       strTableName = strTableName & "_user"
      '   End If
          
900       strFieldValue = nte(vFieldValue)
          
910       If strFieldValue = "" Then
920           Exit Function
930       End If
          
          'get the table this field is linking to:
940       sql = sql & " select l.database_name "
950       sql = sql & " from lims_sys.schema_field f,  "
960       sql = sql & " lims_sys.schema_table t, lims_sys.schema_table l "
970       sql = sql & " where f.schema_table_id = t.schema_table_id "
980       sql = sql & " and upper(f.database_name) = upper('" & strFieldName & "') "
990       sql = sql & " and upper(t.database_name) = upper('" & strTableName & "') "
1000      sql = sql & " and l.schema_table_id = f.lookup_schema_table_id "

1010      Set rs = aConnection.Execute(sql)

          'this in NOT a link field
1020      If rs.EOF Then
1030          ManipulateLinkField = strFieldValue
1040          Exit Function
1050      End If

1060      strPointedTableName = rs("database_name")
          
          'get the value of the NAME column
          'in the refered table:
          'sql2 = sql2 & " select name from lims_sys." & strPointedTableName & _
          'sql2 = sql2 & " where " & strTableName & "_id = " & strFieldValue
          
1070      Set rs = aConnection.Execute(" select name from lims_sys." & strPointedTableName & _
                                       " where " & strPointedTableName & "_id = " & strFieldValue)
          
1080      ManipulateLinkField = nte(rs("name"))
          
1090  Exit Function
ERR_ManipulateLinkField:
1100  MsgBox "Error on line:" & Erl & " in ManipulateLinkField" & vbCrLf & Err.Description
End Function

Public Function TranslateFieldID(aField As ADODB.Field) As String
        Dim i As Integer
        Dim s As String
        Dim sql As String
        Dim rs As ADODB.Recordset
        Dim fieldName As String
        Dim fieldVal As String
1110    fieldName = aField.name
1120    fieldVal = CheckFieldValue(aField)
1130    sql = ""
1140    If fieldVal <> "" Then
1150      Select Case fieldName
            Case "GROUP_ID"
1160          sql = "select name from lims_sys.lims_group where group_id = " & fieldVal
1170        Case "U_PATIENT", "CLIENT_ID"
1180          sql = "select name from lims_sys.client where client_id = " & fieldVal
1190        Case "U_REFERRING_PHYSICIAN", "U_IMPLEMENTING_PHYSICIAN", "SUPPLIER_ID"
1200          sql = "select name from lims_sys.supplier where supplier_id = " & fieldVal
1210        Case "U_COLLECTION_STATION" ' U_CLINIC
1220          sql = "select name from lims_sys.U_CLINIC where u_clinic_id = " & fieldVal
1230        Case "LOCATION_ID"
1240          sql = "select name from lims_sys.LOCATION where location_id = " & fieldVal
1250        Case "PRODUCT_ID"
1260          sql = "select name from lims_sys.product where product_id = " & fieldVal
1270        Case "OPERATOR_ID" ' OPERATOR
1280          sql = "select name from lims_sys.operator where operator_id = " & fieldVal
1290        Case "INSPECTION_PLAN_ID"
1300          sql = "select name from lims_sys.inspection_plan where inspection_plan_id = " & fieldVal
1310        Case "SDG_TEMPLATE_ID", "TEST_TEMPLATE_ID", "RESULT_TEMPLATE_ID", "COFA_TEMPLATE_ID", "REVIEW_TEMPLATE_ID", "STOCK_TEMPLATE_ID"
1320          i = InStr(1, fieldName, "_TEMPLATE_ID")
1330          s = Mid(fieldName, 1, i - 1)
1340          sql = "select name from lims_sys." & s & "_TEMPLATE where " & fieldName & " = " & fieldVal
1350        Case "UNIT_ID"
1360          sql = "select name from lims_sys.UNIT where CHEMICAL_ID = " & fieldVal
1370        Case "CHEMICAL_ID"
1380          sql = "select name from lims_sys.CHEMICAL where CHEMICAL_ID = " & fieldVal
1390        Case "STOCK_TYPE_ID"
1400          sql = "select name from lims_sys.STOCK_TYPE where STOCK_TYPE_ID = " & fieldVal
1410        Case "PLATE_ID"
1420          sql = "select name from lims_sys.PLATE where PLATE_ID = " & fieldVal
1430        Case "INSTRUMENT_ID"
1440          sql = "select name from lims_sys.INSTRUMENT where INSTRUMENT_ID = " & fieldVal
1450        Case Else
1460          TranslateFieldID = fieldVal
1470      End Select
1480      If sql <> "" Then
1490        Set rs = aConnection.Execute(sql)
1500        TranslateFieldID = CheckFieldValue(rs.Fields(0))
1510        If TranslateFieldID = "" Then
1520          TranslateFieldID = fieldVal
1530        End If
1540      End If
1550    Else
1560      TranslateFieldID = ""
1570    End If
End Function

Public Function CreateResults()
        Dim LoadEl As IXMLDOMElement
        Dim ResEntry As IXMLDOMElement
        
        Dim result_id As Double
        
        Dim sql As String
        Dim rs As ADODB.Recordset
        Dim s As String
        Dim name As String
        Dim i As Integer
        Dim j As Integer
        Dim results As IXMLDOMNodeList
        Dim resultsRes As IXMLDOMNodeList
1580    Set results = XMLCopy.selectNodes("//RESULT")
1590    Set LoadEl = Nothing

1600    sql = "select r.name, r.result_id from lims_sys.result r " _
          & "where test_id = " & main_test_id

      '  sql = "select r.name, r.result_id from result r, test t, sample s, aliquot a, sdg " _
      '    & "where test_id = " & main_test_id & " " _
      '    & "and r.test_id = t.test_id " _
      '    & "and t.aliquot_id = a.aliquot_id " _
      '    & "and a.sample_id = s.sample_id " _
      '    & "and s.sdg_id = sdg.sdg_id

          
1610    Set rs = aConnection.Execute(sql)

1620    If Not rs.EOF Then
1630      While Not rs.EOF
1640        name = rs("name")
1650        Set resultsRes = XmlDoc.selectNodes("//RESULT")
1660        For i = 0 To results.length - 1
1670          For j = 0 To resultsRes.length - 1
1680            If resultsRes(j).selectSingleNode("NAME").Text = name And _
                  results(i).selectSingleNode("NAME").Text = name Then
1690                result_id = rs("result_id")
1700                If Not results(i).selectSingleNode("ORIGINAL_RESULT") Is Nothing Then
1710                    s = results(i).selectSingleNode("ORIGINAL_RESULT").Text
1720                    Set XMLResRequest = XmlresVals.createElement("result-request")
1730                    Call XMLResLims.appendChild(XMLResRequest)
1740                    If LoadEl Is Nothing Then
1750                        Set LoadEl = XMLCopy.createElement("load")
1760                        Call LoadEl.setAttribute("entity", "SDG")
1770                        Call LoadEl.setAttribute("id", main_sdg_id)
1780                        Call XMLResRequest.appendChild(LoadEl)
1790                    End If
1800                    Set ResEntry = XMLCopy.createElement("result-entry")
1810                    Call ResEntry.setAttribute("result-id", result_id)
1820                    Call ResEntry.setAttribute("original-result", s)
1830                    Call LoadEl.appendChild(ResEntry)
1840                End If
1850            End If
1860          Next j
1870        Next i
1880        Call rs.MoveNext
1890      Wend
1900    End If
1910    Call rs.Close
End Function

Public Function RenameNewNameSDG(name As String)
        Dim sql As String
1920    sql = "update lims_sys.SDG set name = '" & name & "' where sdg_id = " & main_sdg_id
1930    aConnection.Execute (sql)
End Function

Public Function RenameNewNameSample(name As String)
        Dim sql As String
1940    sql = "update lims_sys.SAMPLE set name = '" & name & "' where sample_id = " & main_sample_id
1950    aConnection.Execute (sql)
End Function

Public Function RenameNewNameAliquot(name As String)
        Dim sql As String
1960    sql = "update lims_sys.ALIQUOT set name = '" & name & "' where aliquot_id = " & main_aliquot_id
1970    aConnection.Execute (sql)
End Function

' level = 0 - sdg (son - sample)
' level = 1 - sample (son - aliquot)
' level = 2 - aliquot (son - test)
Public Function getWFSiblingCountx(wf_id As Double, level As Integer) As Integer
        Dim res As Integer
        Dim rs As ADODB.Recordset
        Dim sql As String
        Dim name1 As String
        Dim name2 As String
1980    Select Case level
          Case 0:
1990          sql = "select count(*) as count_s from lims_sys.workflow_node " _
                & "where workflow_id = " & wf_id & " " _
                & "and (name = 'Sample' or name = 'SubTree' and template is not null)"
2000      Case 1:
2010          sql = "select count(*) as count_s from lims_sys.workflow_node " _
                & "where workflow_id = " & wf_id & " " _
                & "and (name = 'Aliquot' or name = 'SubTree' and template is not null)"
2020      Case 2:
2030          sql = "select count(*) as count_s from lims_sys.workflow_node " _
                & "where workflow_id = " & wf_id & " " _
                & "and (name = 'Test' or name = 'SubTree' and template is not null)"
2040      Case 3:
2050          sql = "select count(*) as count_s from lims_sys.workflow_node " _
                & "where workflow_id = " & wf_id & " " _
                & "and (name = 'Result' or name = 'SubTree' and template is not null)"
2060    End Select
2070    Set rs = aConnection.Execute(sql)
2080    If rs.EOF Then
2090      res = 0
2100    Else
2110      res = rs("count_s")
2120    End If
2130    Call rs.Close
2140    getWFSiblingCountx = res
End Function

' levelID :
' 0 - sdg
' 1 - sample
' 2 - aliquot
' 3 - test
' 4 - result
Public Sub copyXML(levelID As Integer)
        Dim resultValNodeL As IXMLDOMNodeList
        Dim resultValNode As IXMLDOMNode
        Dim xmltemp As IXMLDOMElement
        Dim sql As String
        
        
2150    Call XMLCopy.loadXML(XmlDoc.xml)
        
2160    Set Xmlres = Nothing
2170    Set Xmlres = New DOMDocument
             
2180      If levelID < 1 Then
2190          Set resultValNodeL = XmlDoc.selectNodes("//SAMPLE")
2200          Set resultValNode = Nothing
2210          If Not resultValNodeL Is Nothing Then
2220            Set resultValNode = resultValNodeL(0)
2230          End If
2240          While Not resultValNode Is Nothing
2250            Call resultValNode.parentNode.removeChild(resultValNode)
2260            Set resultValNodeL = XmlDoc.selectNodes("//SAMPLE")
2270            If Not resultValNodeL Is Nothing Then
2280              Set resultValNode = resultValNodeL(0)
2290            End If
2300          Wend
2310      End If
          
2320      If levelID < 2 Then
2330          Set resultValNodeL = XmlDoc.selectNodes("//ALIQUOT")
2340          Set resultValNode = Nothing
2350          If Not resultValNodeL Is Nothing Then
2360            Set resultValNode = resultValNodeL(0)
2370          End If
2380          While Not resultValNode Is Nothing
2390            Call resultValNode.parentNode.removeChild(resultValNode)
2400            Set resultValNodeL = XmlDoc.selectNodes("//ALIQUOT")
2410            If Not resultValNodeL Is Nothing Then
2420              Set resultValNode = resultValNodeL(0)
2430            End If
2440          Wend
2450      End If
          
2460      If levelID < 3 Then
2470          Set resultValNodeL = XmlDoc.selectNodes("//TEST")
2480          Set resultValNode = Nothing
2490          If Not resultValNodeL Is Nothing Then
2500            Set resultValNode = resultValNodeL(0)
2510          End If
2520          While Not resultValNode Is Nothing
2530            Call resultValNode.parentNode.removeChild(resultValNode)
2540            Set resultValNodeL = XmlDoc.selectNodes("//TEST")
2550            If Not resultValNodeL Is Nothing Then
2560              Set resultValNode = resultValNodeL(0)
2570            End If
2580          Wend
2590      End If
          
2600      If levelID < 4 Then
2610          Set resultValNodeL = XmlDoc.selectNodes("//RESULT")
2620          Set resultValNode = Nothing
2630          If Not resultValNodeL Is Nothing Then
2640            Set resultValNodeL = XmlDoc.selectNodes("//RESULT")
2650            Set resultValNode = resultValNodeL(0)
2660          End If
              
2670          While Not resultValNode Is Nothing
2680            Call resultValNode.parentNode.removeChild(resultValNode)
2690            Set resultValNodeL = XmlDoc.selectNodes("//RESULT")
2700            If Not resultValNodeL Is Nothing Then
2710              Set resultValNode = resultValNodeL(0)
2720            End If
2730          Wend
2740      End If
          
2750      Set resultValNodeL = XmlDoc.selectNodes("//ORIGINAL_RESULT")
2760      Set resultValNode = Nothing
2770      If Not resultValNodeL Is Nothing Then
2780        Set resultValNode = resultValNodeL(0)
2790      End If
          
2800      While Not resultValNode Is Nothing
2810        Call resultValNode.parentNode.removeChild(resultValNode)
2820        Set resultValNodeL = XmlDoc.selectNodes("//ORIGINAL_RESULT")
2830        If Not resultValNodeL Is Nothing Then
2840          Set resultValNode = resultValNodeL(0)
2850        End If
2860      Wend
          
      '    Set resultValNodeL = XmlDoc.selectNodes("//NAME")
      '    Set resultValNode = Nothing
      '    If Not resultValNodeL Is Nothing Then
      '      Set resultValNode = resultValNodeL(0)
      '    End If
      '    While Not resultValNode Is Nothing
      '      Call resultValNode.parentNode.removeChild(resultValNode)
      '      Set resultValNodeL = XmlDoc.selectNodes("//NAME")
      '      If Not resultValNodeL Is Nothing Then
      '        Set resultValNode = resultValNodeL(0)
      '      End If
      '    Wend

2870      Set Xmlres = Nothing
2880      Set Xmlres = New DOMDocument
          
2890      Select Case levelID
            Case 0:
2900          sql = "Update lims_sys.sdg " _
                & "set name = '" & SetOldName(sdg_name) & "' " _
                & "where sdg_id = " & sdg_id
2910          aConnection.Execute (sql)
              
      '        XmlDoc.Save ("c:\sdgdoc.xml")
      '        XMLCopy.Save ("c:\sdgcopy.xml")
2920          Call ProcessXML.ProcessXMLWithResponse(XmlDoc, Xmlres)
      '        Xmlres.Save ("c:\sdgRes.xml")
      '        MsgBox Xmlres.xml
2930          main_sdg_id = Xmlres.selectNodes("//SDG").Item(0).selectSingleNode("SDG_ID").Text
2940          max_sample_count = Xmlres.selectNodes("//SAMPLE").length
2950          sample_count = 0
2960          Set xmlSamples = Xmlres
2970          Call RenameNewNameSDG(sdg_name)
2980        Case 1:
2990          sql = "Update lims_sys.sample " _
                & "set name = name || '" & VersionName & "' " _
                & "where sdg_id = " & sdg_id
              
3000          If Not gboSampleNameChanged Then
3010              aConnection.Execute (sql)
3020              gboSampleNameChanged = True
3030          End If
              
      '        sample_name = OriginSDG("aliquot_name")
      '        XmlDoc.Save ("c:\sampledoc.xml")
      '        XMLCopy.Save ("c:\samplecopy.xml")
3040          Call ProcessXML.ProcessXMLWithResponse(XmlDoc, Xmlres)
      '        Xmlres.Save ("c:\sampleres.xml")

      '        MsgBox Xmlres.xml
3050          main_sample_id = Xmlres.selectNodes("//SAMPLE").Item(0).selectSingleNode("SAMPLE_ID").Text
3060          max_aliquot_count = Xmlres.selectNodes("//ALIQUOT").length
3070          aliquot_count = 0
3080          Set xmlAliquots = Xmlres
3090          Call RenameNewNameSample(sample_name)
3100        Case 2:
3110          sql = "Update lims_sys.aliquot " _
                & "set name = name || '" & VersionName & "' " _
                & "where sample_id in ( " _
                & "select sample_id from lims_sys.sample where sdg_id = " & sdg_id & " )"
              
3120          If Not gboAliquotNameChanged Then
3130              aConnection.Execute (sql)
3140              gboAliquotNameChanged = True
3150          End If
              
      '        XmlDoc.Save ("c:\aliquotdoc.xml")
      '        XMLCopy.Save ("c:\aliquotcopy.xml")
3160          Call ProcessXML.ProcessXMLWithResponse(XmlDoc, Xmlres)
      '        Xmlres.Save ("c:\aliquotres.xml")

      '        MsgBox Xmlres.xml
3170          main_aliquot_id = Xmlres.selectNodes("//ALIQUOT").Item(0).selectSingleNode("ALIQUOT_ID").Text
3180          max_test_count = Xmlres.selectNodes("//TEST").length
3190          test_count = 0
3200          Set xmlTests = Xmlres
3210          Call RenameNewNameAliquot(aliquot_name)
3220        Case 3:
      '        XmlDoc.Save ("c:\testdoc.xml")
      '        XMLCopy.Save ("c:\testcopy.xml")
3230          Call ProcessXML.ProcessXMLWithResponse(XmlDoc, Xmlres)
      '        Xmlres.Save ("c:\testres.xml")
      '        MsgBox Xmlres.xml
3240          main_test_id = Xmlres.selectNodes("//TEST").Item(0).selectSingleNode("TEST_ID").Text
3250          max_result_count = Xmlres.selectNodes("//RESULT").length
3260          result_count = 0
3270          Set xmlresults = Xmlres
      '        xmlresults.Save ("c:\testres.xml")
3280        Case 4:
      '        XMLCopy.Save ("c:\resultcopy.xml")
3290          Call CreateResults
      '        Set XmlDoc = XMLCopy
      '        XmlDoc.Save ("c:\resultDoc.xml")
      '        Call ProcessXML.ProcessXMLWithResponse(XmlDoc, Xmlres)
      '        Xmlres.Save ("c:\resultres.xml")
      '        MsgBox Xmlres.xml
      '        Xmlres.Save ("c:\tempres.xml")
3300      End Select
3310      Set XmlDoc = Nothing
3320      Set XmlDoc = New DOMDocument
3330      Set XmlELims = XmlDoc.createElement("lims-request")
3340      Call XmlDoc.appendChild(XmlELims)
3350      Set XmlLoginReq = XmlDoc.createElement("login-request")
3360      Call XmlELims.appendChild(XmlLoginReq)
End Sub

'performs the actual copy.
'will replace the original function: runSDGCopy
Public Function CreateSdgCopy()
          Dim sdgOriginal As New CSdg
          Dim sdgCopy As New CSdg
          Dim s As String
          
3370      If Not isSdgValidForRevision(CStr(sdg_id)) Then
3380          MsgBox " SDG status must be Authorized or Rejected "
3390          Exit Function
3400      End If
          
3410      Call sdgOriginal.Initialize(sdg_name, CStr(sdg_id))
3420      s = LoginNewSdg("Login SDG", sdgOriginal.strExternalReference, sdgOriginal.strWorkflowName)
          
3430      Call sdgCopy.Initialize(sdgOriginal.strExternalReference, s)
3440      Call UpdateSdgLevel(sdgOriginal, sdgCopy)
3450      Call UpdateSampleLevel(sdgOriginal, sdgCopy)
End Function


Private Sub UpdateSdgLevel(sdgOriginal As CSdg, sdgCopy As CSdg)
3460  On Error GoTo ERR_UpdateSdgLevel

          Dim rs As Recordset
          Dim sql As String

          'decide on the right name for the old sdg
          '(there might have already been some revisions)
3470      Set rs = aConnection.Execute(" select count(*) from lims_sys.sdg " & _
              " where external_reference = '" & sdgOriginal.strExternalReference & "' " & _
              " and instr(name, 'V', 1) > 0 ")
              
3480      iRevisionNum = ntz(rs(0))
3490      iRevisionNum = iRevisionNum + 1
          
          'change the names in memory:
3500      sdgCopy.strName = sdgOriginal.strName
3510      sdgOriginal.strName = sdgOriginal.strName & "V" & CStr(iRevisionNum)
          
          
          'change the names and status in DB:
3520      aConnection.Execute (" update lims_sys.sdg set name = '" & sdgOriginal.strName & "' " & _
                               " where sdg_id = " & sdgOriginal.strId)
          
3530      aConnection.Execute (" update lims_sys.sdg set name = '" & sdgCopy.strName & "', " & _
                               " status = 'V' " & _
                               " where sdg_id = " & sdgCopy.strId)
          
          
          'update all the NEW SDG fields in the DB according to sdgOriginal
3540      Set rs = aConnection.Execute(" select * from lims_sys.sdg " & _
                                       " where sdg_id = " & sdgOriginal.strId)
                                       
3550      Call UpdateRecordById("sdg", "sdg_id", sdgCopy.strId, sdgOriginal.strId, rs)
                                       
                                       
3560      Set rs = aConnection.Execute(" select * from lims_sys.sdg_user " & _
                                       " where sdg_id = " & sdgOriginal.strId)
                                       
3570      Call UpdateRecordById("sdg_user", "sdg_id", sdgCopy.strId, sdgOriginal.strId, rs)
          
          
          'update revision cause:
3580      sql = "update lims_sys.sdg_user set u_revision_cause = '" & revision_cause & "' " _
            & "where sdg_id = " & sdg_id
3590      aConnection.Execute (sql)
          
          
          'update who is the last SDG update and who isn't:
3600      aConnection.Execute (" update lims_sys.sdg_user " & _
                               " set u_is_last_update = 'F' " & _
                               " where sdg_id = " & sdgOriginal.strId)
          
3610      aConnection.Execute (" update lims_sys.sdg_user " & _
                               " set u_is_last_update = 'T' " & _
                               " where sdg_id = " & sdgCopy.strId)
          
          
          'enter a record to the sdg log table:
3620      Set sdg_log.con = aConnection
3630      sdg_log.Session = CDbl(NtlsCon.GetSessionId)
3640      sdg_log_desc = ""
3650      Call sdg_log.InsertLog(CLng(sdg_id), "REV.UPD", sdg_log_desc)
          
3660  Exit Sub
ERR_UpdateSdgLevel:
3670  MsgBox "Error on line:" & Erl & " in UpdateSdgLevel on line " & Erl & vbCrLf & Err.Description
End Sub


Private Sub UpdateSampleLevel(sdgOriginal As CSdg, sdgCopy As CSdg)
3680  On Error GoTo ERR_UpdateSampleLevel

          Dim i As Integer
          Dim sampleCopy As CSample
          Dim sampleOriginal As CSample
          Dim rs As Recordset
          
3690      For i = 0 To sdgOriginal.dicSamples.Count - 1
3700          Set sampleOriginal = sdgOriginal.dicSamples(i)
          
3710          If Not sdgCopy.dicSamples.Exists(i) Then
                  'create the missing sample for the new sdg:
3720              Set sampleCopy = New CSample
3730              sampleCopy.strId = _
                     LoginNewSample("Login Sample", "sample_" & CStr(i), _
                                     sampleOriginal.strWorkflowName, sdgCopy.strId)
                  
                  
3740              aConnection.Execute (" update lims_sys.sample " & _
                                       " set sdg_id = " & sdgCopy.strId & _
                                       " where sample_id = " & sampleCopy.strId)
                  
3750              Call sampleCopy.Initialize("sample_" & CStr(i), sampleCopy.strId)
                  
                  'add the sample to memory in location i using CSample.initialize():
3760              Set sdgCopy.dicSamples(i) = sampleCopy
                  
3770          Else
3780              Set sampleCopy = sdgCopy.dicSamples(i)
3790          End If
              
              'change the names in memory:
3800          sampleCopy.strName = sampleOriginal.strName
3810          sampleOriginal.strName = sampleOriginal.strName & "V" & CStr(iRevisionNum)
              
              
              'change the names and status in DB:
3820          aConnection.Execute (" update lims_sys.sample set name = '" & sampleOriginal.strName & "' " & _
                                   " where sample_id = " & sampleOriginal.strId)
              
3830          aConnection.Execute (" update lims_sys.sample set name = '" & sampleCopy.strName & "', " & _
                                   " status = 'V' " & _
                                   " where sample_id = " & sampleCopy.strId)
                  
              
              
              'update all relevant fields in DB:
3840          Set rs = aConnection.Execute(" select * from lims_sys.sample " & _
                                           " where sample_id = " & sampleOriginal.strId)
                                           
3850          Call UpdateRecordById("sample", "sample_id", _
                                     sampleCopy.strId, sampleOriginal.strId, rs)
                                           
                                           
3860          Set rs = aConnection.Execute(" select * from lims_sys.sample_user " & _
                                           " where sample_id = " & sampleOriginal.strId)
                                           
3870          Call UpdateRecordById("sample_user", "sample_id", _
                                     sampleCopy.strId, sampleOriginal.strId, rs)
              
              'update all blocks for this sample:
3880          Call UpdateBlocksForSample(sampleOriginal, sampleCopy)
3890      Next i
          
          
3900  Exit Sub
ERR_UpdateSampleLevel:
3910  MsgBox "Error on line:" & Erl & " in UpdateSampleLevel" & vbCrLf & Err.Description
End Sub


Private Sub UpdateBlocksForSample(sampleOriginal As CSample, sampleCopy As CSample)
3920  On Error GoTo ERR_UpdateBlocksForSample

          Dim i As Integer
          Dim blockCopy As CBlock
          Dim blockOriginal As CBlock
          Dim rs As Recordset
          
3930      For i = 0 To sampleOriginal.dicBlocks.Count - 1
3940          Set blockOriginal = sampleOriginal.dicBlocks(i)
          
3950          If Not sampleCopy.dicBlocks.Exists(i) Then
3960              Set blockCopy = New CBlock
                  
                  'change original block's name before logging in a new block,
                  'for its name will be the same (Unique Constraint):
3970              aConnection.Execute (" update lims_sys.aliquot " & _
                    " set name = '" & blockOriginal.strName & "V" & CStr(iRevisionNum) & "' " & _
                    " where aliquot_id = " & blockOriginal.strId)
                  
3980              blockCopy.strId = _
                     LoginNewBlock("Login Block", "block__" & CStr(i), _
                                    blockOriginal.strWorkflowName, sampleCopy.strId)
                  
                  
3990              aConnection.Execute (" update lims_sys.aliquot " & _
                                       " set sample_id = " & sampleCopy.strId & _
                                       " where aliquot_id = " & blockCopy.strId)
                  
4000              Call blockCopy.Initialize("block_" & CStr(i), blockCopy.strId)
                  
4010              Set sampleCopy.dicBlocks(i) = blockCopy
4020          Else
4030              Set blockCopy = sampleCopy.dicBlocks(i)
4040          End If
              
              'change the names in memory:
4050          blockCopy.strName = blockOriginal.strName
4060          blockOriginal.strName = blockOriginal.strName & "V" & CStr(iRevisionNum)
              
              'change the names and status in DB:
4070          aConnection.Execute (" update lims_sys.aliquot set name = '" & blockOriginal.strName & "' " & _
                                   " where aliquot_id = " & blockOriginal.strId)
              
4080          aConnection.Execute (" update lims_sys.aliquot set name = '" & blockCopy.strName & "', " & _
                                   " status = 'V' " & _
                                   " where aliquot_id = " & blockCopy.strId)
                  
              
              'update all relevant fields in DB:
4090          Set rs = aConnection.Execute(" select * from lims_sys.aliquot " & _
                                           " where aliquot_id = " & blockOriginal.strId)
                                           
4100          Call UpdateRecordById("aliquot", "aliquot_id", _
                                     blockCopy.strId, blockOriginal.strId, rs)
                                           
                                           
4110          Set rs = aConnection.Execute(" select * from lims_sys.aliquot_user " & _
                                           " where aliquot_id = " & blockOriginal.strId)
                                           
4120          Call UpdateRecordById("aliquot_user", "aliquot_id", _
                                     blockCopy.strId, blockOriginal.strId, rs)
                   
                   
              'update all slides & tests for this block:
4130          Call UpdateTestsForBlock(blockOriginal, blockCopy)
4140          Call UpdateSlidesForBlock(blockOriginal, blockCopy)
4150      Next i
              
4160      Exit Sub
ERR_UpdateBlocksForSample:
4170  MsgBox "Error on line:" & Erl & " in UpdateBlocksForSample" & vbCrLf & Err.Description
End Sub


Private Sub UpdateTestsForBlock(blockOriginal As CBlock, blockCopy As CBlock)
4180  On Error GoTo ERR_UpdateTestsForBlock
          
          Dim i As Integer
          Dim testOriginal As CTest
          Dim testCopy As CTest
          Dim rs As Recordset
          
4190  If DEFINE_DEBUG Then MsgBox "update tests for block" & vbCrLf & _
                                  "block original: " & blockOriginal.dicTests.Count & vbCrLf & _
                                  "block copy: " & blockCopy.dicTests.Count
          
4200      For i = 0 To blockOriginal.dicTests.Count - 1
4210          Set testOriginal = blockOriginal.dicTests(i)

4220          If Not blockCopy.dicTests.Exists(i) Then
4230              Set testCopy = New CTest
                  
4240              Call LoginNewTest("Login Test", "test__" & CStr(i), _
                                    testOriginal.strWorkflowName, blockCopy.strId)

4250              testCopy.strId = GetMaxTest(blockCopy.strId)

4260              aConnection.Execute (" update lims_sys.test " & _
                                       " set aliquot_id = " & blockCopy.strId & _
                                       " where test_id = " & testCopy.strId)

4270              Call testCopy.Initialize("test_" & CStr(i), testCopy.strId)

4280              Set blockCopy.dicTests(i) = testCopy
4290          Else
4300              Set testCopy = blockCopy.dicTests(i)
4310          End If
          
          'never login more tests than those created by creating the block:
      '    For i = 0 To blockCopy.dicTests.Count - 1
      '        Set testOriginal = blockOriginal.dicTests(i)
      '        Set testCopy = blockCopy.dicTests(i)
          
              'change the names in memory:
4320          testCopy.strName = testOriginal.strName
              
              'change the names and status in DB:
4330          aConnection.Execute (" update lims_sys.test set name = '" & testCopy.strName & "', " & _
                                   " status = 'V' " & _
                                   " where test_id = " & testCopy.strId)
              
              'update all relevant fields in DB:
4340          Set rs = aConnection.Execute(" select * from lims_sys.test " & _
                                           " where test_id = " & testOriginal.strId)
                                           
4350          Call UpdateRecordById("test", "test_id", _
                                     testCopy.strId, testOriginal.strId, rs)
                                           
                                           
4360          Set rs = aConnection.Execute(" select * from lims_sys.test_user " & _
                                           " where test_id = " & testOriginal.strId)
                                           
4370          Call UpdateRecordById("test_user", "test_id", _
                                     testCopy.strId, testOriginal.strId, rs)
          
              'update the results of each test:
4380          Call UpdateResultsForTest(testOriginal, testCopy)
4390      Next i
          
4400      Exit Sub
ERR_UpdateTestsForBlock:
4410  MsgBox "Error on line:" & Erl & " in UpdateTestsForBlock" & vbCrLf & Err.Description
End Sub


Private Sub UpdateSlidesForBlock(blockOriginal As CBlock, blockCopy As CBlock)
4420  On Error GoTo ERR_UpdateSlidesForBlock
          
          Dim i As Integer
          Dim slideOriginal As CSlide
          Dim slideCopy As CSlide
          Dim rs As Recordset
          
4430      For i = 0 To blockOriginal.dicSlides.Count - 1
4440          Set slideOriginal = blockOriginal.dicSlides(i)
          
4450          If Not blockCopy.dicSlides.Exists(i) Then
4460              Set slideCopy = New CSlide
                  
4470              Call LoginNewSlide("Add Slide", blockCopy.strId)
4480              slideCopy.strId = GetMaxSlide(blockCopy.strId)
                  
      '            aConnection.Execute (" update lims_sys.aliquot " & _
                                       " set aliquot_id = " & blockCopy.strId & _
                                       " where aliquot_id = " & slideCopy.strId)
                  
4490              Call slideCopy.Initialize("slide_" & CStr(i), slideCopy.strId)
                  
4500              Set blockCopy.dicSlides(i) = slideCopy
4510          Else
4520              Set slideCopy = blockCopy.dicSlides(i)
4530          End If
          
              'change the names in memory:
4540          slideCopy.strName = slideOriginal.strName
4550          slideOriginal.strName = slideOriginal.strName & "V" & CStr(iRevisionNum)
              
              'change the names and status in DB:
4560          aConnection.Execute (" update lims_sys.aliquot set name = '" & slideOriginal.strName & "' " & _
                                   " where aliquot_id = " & slideOriginal.strId)
              
4570          aConnection.Execute (" update lims_sys.aliquot set name = '" & slideCopy.strName & "', " & _
                                   " status = 'V' " & _
                                   " where aliquot_id = " & slideCopy.strId)
              
              'update all relevant fields in DB:
4580          Set rs = aConnection.Execute(" select * from lims_sys.aliquot " & _
                                           " where aliquot_id = " & slideOriginal.strId)
                                           
4590          Call UpdateRecordById("aliquot", "aliquot_id", _
                                     slideCopy.strId, slideOriginal.strId, rs)
                                           
                                           
4600          Set rs = aConnection.Execute(" select * from lims_sys.aliquot_user " & _
                                           " where aliquot_id = " & slideOriginal.strId)
                                           
4610          Call UpdateRecordById("aliquot_user", "aliquot_id", _
                                     slideCopy.strId, slideOriginal.strId, rs)
          
              'update the results of each test:
4620          Call UpdateTestsForSlide(slideOriginal, slideCopy)
4630      Next i
          
4640      Exit Sub
ERR_UpdateSlidesForBlock:
4650  MsgBox "Error on line:" & Erl & " in UpdateSlidesForBlock" & vbCrLf & Err.Description
End Sub

Private Sub UpdateTestsForSlide(slideOriginal As CSlide, slideCopy As CSlide)
4660  On Error GoTo ERR_UpdateTestsForSlide
          
          Dim i As Integer
          Dim testOriginal As CTest
          Dim testCopy As CTest
          Dim rs As Recordset
          
4670  If DEFINE_DEBUG Then MsgBox "update tests for slide" & vbCrLf & _
                                  "slide original: " & slideOriginal.dicTests.Count & vbCrLf & _
                                  "slide copy: " & slideCopy.dicTests.Count
                                  
4680      For i = 0 To slideOriginal.dicTests.Count - 1
4690          Set testOriginal = slideOriginal.dicTests(i)

4700          If Not slideCopy.dicTests.Exists(i) Then
4710              Set testCopy = New CTest
                  
4720              Call LoginNewTest("Login Test", "test__" & CStr(i), _
                                    testOriginal.strWorkflowName, slideCopy.strId)

4730              testCopy.strId = GetMaxTest(slideCopy.strId)
                  
4740              aConnection.Execute (" update lims_sys.test " & _
                                       " set aliquot_id = " & slideCopy.strId & _
                                       " where test_id = " & testCopy.strId)

4750              Call testCopy.Initialize("test_" & CStr(i), testCopy.strId)

4760              Set slideCopy.dicTests(i) = testCopy
4770          Else
4780              Set testCopy = slideCopy.dicTests(i)
4790          End If
          
          'never login more tests than those created by creating the slide:
      '    For i = 0 To slideCopy.dicTests.Count - 1
      '        Set testOriginal = slideOriginal.dicTests(i)
      '        Set testCopy = slideCopy.dicTests(i)
              
              'change the names in memory:
4800          testCopy.strName = testOriginal.strName
              
              'change the names and status in DB:
4810          aConnection.Execute (" update lims_sys.test set name = '" & testCopy.strName & "', " & _
                                   " status = 'V' " & _
                                   " where test_id = " & testCopy.strId)
              
              'update all relevant fields in DB:
4820          Set rs = aConnection.Execute(" select * from lims_sys.test " & _
                                           " where test_id = " & testOriginal.strId)
                                           
4830          Call UpdateRecordById("test", "test_id", _
                                     testCopy.strId, testOriginal.strId, rs)
                                           
                                           
4840          Set rs = aConnection.Execute(" select * from lims_sys.test_user " & _
                                           " where test_id = " & testOriginal.strId)
                                           
4850          Call UpdateRecordById("test_user", "test_id", _
                                     testCopy.strId, testOriginal.strId, rs)
          
              'update the results of each test:
4860          Call UpdateResultsForTest(testOriginal, testCopy)
4870      Next i
          
4880      Exit Sub
ERR_UpdateTestsForSlide:
4890  MsgBox "Error on line:" & Erl & " in UpdateTestsForSlide" & vbCrLf & Err.Description
End Sub


Private Sub UpdateResultsForTest(testOriginal As CTest, testCopy As CTest)
4900  On Error GoTo ERR_UpdateResultsForTest
          
          Dim i As Integer
          Dim resultOriginal As CResult
          Dim resultcopy As CResult
          Dim rs As Recordset
          
4910      For i = 0 To testOriginal.dicResults.Count - 1
4920          Set resultOriginal = testOriginal.dicResults(i)
                      
4930          If Not testCopy.dicResults.Exists(i) Then
4940              Set resultcopy = New CResult
                  
4950              resultcopy.strId = LoginNewResult(resultOriginal.strWorkflowName, _
                                                     testCopy.strId)

                     'LoginNewResult("Login Result", "result__" & CStr(i), _
                                    resultOriginal.strWorkflowName, testCopy.strId)
                  
                  
4960              aConnection.Execute (" update lims_sys.result " & _
                                       " set test_id = " & testCopy.strId & _
                                       " where result_id = " & resultcopy.strId)
                  
4970              Call resultcopy.Initialize("result_" & CStr(i), resultcopy.strId)
                  
4980              Set testCopy.dicResults(i) = resultcopy
4990          Else
5000              Set resultcopy = testCopy.dicResults(i)
5010          End If
          
              'change the names in memory:
5020          resultcopy.strName = resultOriginal.strName
              
              'change the names and status in DB:
5030          aConnection.Execute (" update lims_sys.result set name = '" & resultcopy.strName & "', " & _
                                   " status = 'V' " & _
                                   " where result_id = " & resultcopy.strId)
              
              'update all relevant fields in DB:
5040          Set rs = aConnection.Execute(" select * from lims_sys.result " & _
                                           " where result_id = " & resultOriginal.strId)
                                           
5050          Call UpdateRecordById("result", "result_id", _
                                     resultcopy.strId, resultOriginal.strId, rs)
                                           
                                           
5060          Set rs = aConnection.Execute(" select * from lims_sys.result_user " & _
                                           " where result_id = " & resultOriginal.strId)
                                           
5070          Call UpdateRecordById("result_user", "result_id", _
                                     resultcopy.strId, resultOriginal.strId, rs)
          
          
              'change the status in DB:
5080          aConnection.Execute (" update lims_sys.result set " & _
                                   " status = 'C' " & _
                                   " where result_id = " & resultcopy.strId)
                                   
              
5090          Call UpdateRTFResultForResult(resultOriginal.strId, resultcopy.strId)
5100      Next i
          
5110      Exit Sub
ERR_UpdateResultsForTest:
5120  MsgBox "Error on line:" & Erl & " in UpdateResultsForTest" & vbCrLf & Err.Description
End Sub

'update the record for the results that contain FREE TEXT
'and have a record in the rtf_result table
Private Sub UpdateRTFResultForResult(strIdOriginal As String, strIdCopy As String)
5130  On Error GoTo UpdateRTFResultForResult
          
          Dim sql As String
          Dim rs As Recordset
          
          'check if this result has a record in the rtf_result table:
5140      sql = " select * from lims_sys.rtf_result where rtf_result_id = " & strIdOriginal
5150      Set rs = aConnection.Execute(sql)
5160      If rs.EOF Then
5170          Exit Sub
5180      End If
          
          'insert a new record with the ID of the copied result:
5190      sql = " insert into lims_sys.rtf_result"
5200      sql = sql & " (rtf_result_id)"
5210      sql = sql & " values (" & strIdCopy & ")"
5220      aConnection.Execute (sql)
          
          'update the result value to be the same as the original result:
5230      sql = " update lims_sys.rtf_result set rtf_text = "
5240      sql = sql & " ("
5250      sql = sql & "    select rtf_text from lims_sys.rtf_result "
5260      sql = sql & "    where rtf_result_id = " & strIdOriginal
5270      sql = sql & " )"
5280      sql = sql & " where  rtf_result_id = " & strIdCopy
5290      aConnection.Execute (sql)
          
5300      Exit Sub
UpdateRTFResultForResult:
5310  MsgBox "Error on line:" & Erl & vbCrLf & "UpdateRTFResultForResult" & vbCrLf & Err.Description
End Sub


'updates the relevant table by the given id
'to the values held by the recordset
'strIdFieldName - the id field name for this table
'strIdTarget    - the id field value for the record to be modified
'strIdSource    - the id field value for the record to read from
Private Sub UpdateRecordById(strTableName As String, strIdFieldName As String, _
                             strIdTarget As String, strIdSource As String, _
                             rs As Recordset)
                                  
5320  On Error GoTo ERR_UpdateRecordById
          
          Dim i As Integer
          Dim strFieldName As String
          Dim varFieldValue As Variant
          Dim sql As String
          
         ' sql = ""
         ' sql = sql & " update lims_sys." & strTableName & " set "
          
5330      For i = 0 To rs.Fields.Count - 1
5340          strFieldName = rs.Fields(i).name
5350          varFieldValue = rs.Fields(i).Value
              
5360          If shouldCopyField(strFieldName, strTableName, varFieldValue) = True Then
                  'If GetFieldType(strFieldName, strTableName) <> "D" Then
        '          If shouldCopyDirectly(strFieldName, strTableName, varFieldValue) = False Then

        '              sql = " update lims_sys." & strTableName
        '              sql = sql & " set " & strFieldName & " = '" & rs.Fields(i) & "'"
        '              sql = sql & " where " & strIdFieldName & " = " & strIdTarget
        '          Else
                      'copy the data
                      'directly from the data base:
5370                  sql = " update lims_sys." & strTableName
5380                  sql = sql & " set " & strFieldName & " = "
5390                  sql = sql & " ( "
5400                  sql = sql & "    select " & strFieldName
5410                  sql = sql & "    from lims_sys." & strTableName
5420                  sql = sql & "    where " & strIdFieldName & " = " & strIdSource
5430                  sql = sql & " ) "
5440                  sql = sql & " where " & strIdFieldName & " = " & strIdTarget
        '          End If
                  
5450              aConnection.Execute (sql)
         '         sql = sql & " " & strFieldName & " = '" & varFieldValue & "',"
5460          End If
5470      Next i
          
         ' sql = Left(sql, Len(sql) - 1)
         ' sql = sql & " where " & strIdFieldName & " = " & strId

      'MsgBox sql

         ' aConnection.Execute (sql)

5480  Exit Sub
ERR_UpdateRecordById:
5490  MsgBox "Error on line:" & Erl & " in UpdateTableById" & vbCrLf & Err.Description
End Sub

'decides weather this field should be copied or not:
'1. is its value NULL?
'2. is it a STOP-LIST field?
Private Function shouldCopyField(strFieldName As String, strTableName As String, _
                                 varFieldValue As Variant) As Boolean
          
5500  On Error GoTo ERR_shouldCopyField

          Dim rs As Recordset
          
5510      shouldCopyField = False
          
5520      If IsNull(varFieldValue) Then Exit Function
          
5530      If dicStopListFields.Exists(LCase(strFieldName)) Then Exit Function
          
5540      shouldCopyField = True
                 
5550      Exit Function
ERR_shouldCopyField:
5560  MsgBox "Error on line:" & Erl & " in shouldCopyField" & vbCrLf & Err.Description
End Function

'decides if the field should be copied directly
'from the DB and NOT from the memory:
Private Function shouldCopyDirectly(strFieldName As String, _
                                    strTableName As String, _
                                    varFieldValue As Variant) As Boolean

          Dim strFieldType As String
5570      strFieldType = GetFieldType(strFieldName, strTableName)
              
5580      shouldCopyDirectly = False
          
5590      If strFieldType = "D" Then
5600          shouldCopyDirectly = True
5610      Else
5620          If strFieldType = "T" And FieldContainBreaks(CStr(varFieldValue)) Then
5630              shouldCopyDirectly = True
5640          End If
5650      End If
End Function


'checks for fields like result.original_result, result.formated result;
'such fields are NOT copied correctly from the memory
Private Function FieldContainBreaks(strFieldValue As String) As Boolean
5660  On Error GoTo ERR_FieldContainBreaks
          
5670      FieldContainBreaks = False
          
5680      If InStr(1, strFieldValue, "<", vbTextCompare) > 0 Then
5690          FieldContainBreaks = True
5700      End If
          
5710      Exit Function
ERR_FieldContainBreaks:
5720  MsgBox "Error on line:" & Erl & " in FieldContainBreaks" & vbCrLf & Err.Description
End Function

'return the type of field:
Private Function GetFieldType(strFieldName As String, strTableName As String) As String
5730  On Error GoTo ERR_GetFieldType
          
          Dim sql As String
          Dim rs As Recordset
          
5740      sql = " select datatype "
5750      sql = sql & " from lims_sys.schema_field"
5760      sql = sql & " where database_name = '" & UCase(strFieldName) & "'"
5770      sql = sql & " and schema_table_id = "
5780      sql = sql & " ("
5790      sql = sql & "   select schema_table_id"
5800      sql = sql & "   from lims_sys.schema_table"
5810      sql = sql & "   where database_name = '" & UCase(strTableName) & "'"
5820      sql = sql & " )"
              
5830      Set rs = aConnection.Execute(sql)
5840      If Not rs.EOF Then
5850          GetFieldType = nte(rs(0))
5860      End If
              
5870      Exit Function
ERR_GetFieldType:
5880  MsgBox "Error on line:" & Erl & " in GetFieldType" & vbCrLf & Err.Description
End Function

'the OLD VERSION of creating the request tree copy;
'to be replaced at 04.2006 by "CreateSdgCopy()"
Public Function runSDGCopy()
        Dim Errs As IXMLDOMNodeList
        
        Dim sample_workflow_name As String
        Dim aliquot_workflow_name As String
        Dim test_workflow_name As String
        Dim result_workflow_name As String
            
        Dim xmltemp As IXMLDOMElement
        Dim first_time_sdg As Boolean
        Dim sdg_on As Boolean
        Dim sdg_user_on As Boolean
        Dim sample_on As Boolean
        Dim sample_user_on As Boolean
        Dim aliquot_on As Boolean
        Dim aliquot_user_on As Boolean
        Dim test_on As Boolean
        Dim test_user_on As Boolean
        Dim result_on As Boolean
        Dim result_user_on As Boolean
        Dim CurrFieldName As String
        Dim CurrFieldValue As Variant
        Dim r_calculation_id As String
        Dim r_result_type As String
        Dim i As Integer
        Dim tmpj As Double
        Dim ind As Double
        Dim tempS As String
        Dim s As String
        Dim sql As String
        Dim sqlWhere As String
        Dim OriginSDG As ADODB.Recordset
        Dim log_sdg_id As Long

5890    old_sdg_id = -1

5900    gboSampleNameChanged = False
5910    gboAliquotNameChanged = False
        
5920    sample_count = 0
5930    aliquot_count = 0
5940    test_count = 0
5950    result_count = 0
        
5960    max_sample_count = 0
5970    max_aliquot_count = 0
5980    max_test_count = 0
5990    max_result_count = 0
        
6000    first_time_sdg = True
        
6010    Set XmlDoc = Nothing
6020    Set XmlDoc = New DOMDocument
6030    Set Xmlres = Nothing
6040    Set Xmlres = New DOMDocument
6050    Set XmlELims = XmlDoc.createElement("lims-request")
6060    Call XmlDoc.appendChild(XmlELims)
6070    Set XmlLoginReq = XmlDoc.createElement("login-request")
6080    Call XmlELims.appendChild(XmlLoginReq)
6090    Set XMLESDG = XmlDoc.createElement("SDG")
6100    Call XmlLoginReq.appendChild(XMLESDG)
        
6110    Set XmlresVals = Nothing
6120    Set XmlresVals = New DOMDocument
6130    Set XMLResLims = XmlresVals.createElement("lims-request")
6140    Call XmlresVals.appendChild(XMLResLims)

6150    Set sdg_log.con = aConnection
6160    sdg_log.Session = CDbl(NtlsCon.GetSessionId)

6170    i = InStr(1, sdg_name, "V")
6180    If i > 0 Then
6190          frmMain.edtBarCode.BackColor = vbRed
6200          MsgBox "Cannot do revision more then once on the same SDG !"
6210          frmMain.edtBarCode.BackColor = vbWhite
6220          Call frmMain.edtBarCode.SetFocus
6230    Else
6240        sql = "select sdg_w.name as sdg_workflow_name, sdg_w.workflow_id as sdg_workflow_id, sdg.*, 'x' as xx1, sdg_u.*, 'x' as xx2, " _
              & "sa_w.name as sa_workflow_name, sa_w.workflow_id as sa_workflow_id, sa.name as sample_name, sa.*, 'x' as xx3, sa_u.*, 'x' as xx4, " _
              & "a_w.name as a_workflow_name, a_w.workflow_id as a_workflow_id, a.name as aliquot_name, a.*, 'x' as xx5, a_u.*, 'x' as xx6, " _
              & "t_w.name as t_workflow_name, t_w.workflow_id as t_workflow_id, t.name as test_name, t.*, 'x' as xx7, t_u.*, 'x' as xx8, " _
              & "r_w.name as r_workflow_name, r_w.workflow_id as r_workflow_id, r_t.calculation_id as r_calculation_id, r_t.result_type as r_result_type, r.*, 'x' as xx9, r_u.*, 'x' as xx10 " _
              & "from lims_sys.sdg, lims_sys.sample sa, lims_sys.aliquot a, lims_sys.test t, lims_sys.result r, " _
              & "lims_sys.workflow_node sdg_n, lims_sys.workflow_node sa_n, lims_sys.workflow_node a_n, lims_sys.workflow_node t_n, lims_sys.workflow_node r_n, " _
              & "lims_sys.workflow sdg_w, lims_sys.workflow sa_w, lims_sys.workflow a_w, lims_sys.workflow t_w, lims_sys.workflow r_w, " _
              & "lims_sys.sdg_user sdg_u, lims_sys.sample_user sa_u, lims_sys.aliquot_user a_u, lims_sys.test_user t_u, lims_sys.result_user r_u, lims_sys.result_template r_t "
             
6250         sqlWhere = "where sdg.sdg_id = sa.sdg_id " _
              & "and sa.sample_id = a.sample_id (+) " _
              & "and a.aliquot_id = t.aliquot_id (+) " _
              & "and t.test_id = r.test_id (+) " _
              & "and not exists (select 1 from lims_sys.aliquot_formulation where child_aliquot_id = a.aliquot_id) " _
              & "and sdg.workflow_node_id = sdg_n.workflow_node_id (+) and sa.workflow_node_id = sa_n.workflow_node_id (+) " _
              & "and a.workflow_node_id = a_n.workflow_node_id (+) and t.workflow_node_id = t_n.workflow_node_id (+) and r.workflow_node_id = r_n.workflow_node_id (+) " _
              & "and sdg_n.workflow_id = sdg_w.workflow_id (+) and sa_n.workflow_id = sa_w.workflow_id (+) and a_n.workflow_id = a_w.workflow_id (+) and t_n.workflow_id = t_w.workflow_id (+) and r_n.workflow_id = r_w.workflow_id (+) " _
              & "and sdg.sdg_id = sdg_u.sdg_id (+) " _
              & "and sa.sample_id = sa_u.sample_id (+) " _
              & "and a.aliquot_id = a_u.aliquot_id (+) " _
              & "and t.test_id = t_u.test_id (+) " _
              & "and r.result_id = r_u.result_id (+) " _
              & "and r.result_template_id = r_t.result_template_id (+) " _
              & "and sdg.status in ('A', 'R') " _
              & "and (sa.status <> 'X' or sa.status is null) and (a.status <> 'X' or a.status is null) and (t.status <> 'X' or t.status is null) and (r.status <> 'X' or r.status is null) " _
              & "and sdg.sdg_id = " & sdg_id & " order by sdg.sdg_id, sa.sample_id, a.aliquot_id, t.test_id, r.result_id "
            
      '        & "and (sa.status <> 'X' or sa.status is null) and (a.status <> 'X' or a.status is null) and (t.status <> 'X' or t.status is null) and (r.status <> 'X' or r.status is null) " _




6260        Set OriginSDG = aConnection.Execute(sql & sqlWhere)
      '      MsgBox OriginSDG.RecordCount
6270        old_sample_id = -1
6280        old_aliquot_id = -1
6290        old_test_id = -1
6300        old_result_id = -1
6310        If OriginSDG.EOF Then
6320          frmMain.edtBarCode.BackColor = vbRed
6330          MsgBox "SDG should be authorised/rejected before it is revisioned !"
6340          frmMain.edtBarCode.BackColor = vbWhite
6350          Call frmMain.edtBarCode.SetFocus
6360        Else
      '        sql = "update lims_sys.sdg_user set u_revision_cause = '" & revision_cause & "' " _
      '          & "where sdg_id = " & sdg_id
              'aConnection.Execute (sql)
6370            While Not OriginSDG.EOF
6380              sdg_on = True
6390              sdg_user_on = False
6400              sample_on = False
6410              sample_user_on = False
6420              aliquot_on = False
6430              aliquot_user_on = False
6440              test_on = False
6450              test_user_on = False
6460              result_on = False
6470              result_user_on = False
6480              For i = 0 To OriginSDG.Fields.Count - 1
6490                CurrFieldName = OriginSDG.Fields(i).name
6500                CurrFieldValue = OriginSDG.Fields(i).Value
                     
6510                If IsNull(OriginSDG.Fields(i)) Then
6520                  If UCase(CurrFieldName) = ("SDG_WORKFLOW_NAME") Or _
                         UCase(CurrFieldName) = ("SA_WORKFLOW_NAME") Or _
                         UCase(CurrFieldName) = ("A_WORKFLOW_NAME") Or _
                         UCase(CurrFieldName) = ("T_WORKFLOW_NAME") Or _
                         UCase(CurrFieldName) = ("R_WORKFLOW_NAME") Then
6530                      Exit For
6540                  End If
6550                End If
                    
6560                Select Case CurrFieldName
                      Case "XX1"
6570                    sdg_on = False
6580                    sdg_user_on = True
6590                  Case "XX2"
6600                    If old_sdg_id <> sdg_id Then 'once
6610                      Call copyXML(0)
6620                      old_sdg_id = sdg_id
6630                    End If
6640                    sdg_user_on = False
6650                    sample_on = True
6660                    first_time_sdg = False
6670                  Case "XX3"
6680                    sample_on = False
6690                    sample_user_on = True
6700                  Case "XX4"
6710                    If old_sample_id <> sample_id Then
6720                      Call copyXML(1)
6730                    End If
6740                    sample_user_on = False
6750                    aliquot_on = True
6760                  Case "XX5"
6770                    aliquot_on = False
6780                    aliquot_user_on = True
6790                  Case "XX6"
6800                    If old_aliquot_id <> aliquot_id Then
6810                      Call copyXML(2)
6820                    End If
6830                    aliquot_user_on = False
6840                    test_on = True
6850                  Case "XX7"
6860                    test_on = False
6870                    test_user_on = True
6880                  Case "XX8"
6890                    If old_test_id <> test_id Then
6900                      Call copyXML(3)
6910                    End If
6920                    test_user_on = False
6930                    result_on = True
6940                  Case "XX9"
6950                    result_on = False
6960                    result_user_on = True
6970                  Case "XX10"
6980                    result_user_on = False
6990                    Call copyXML(4)
7000                  Case Else
                     
7010                  If (sdg_on Or sdg_user_on) And first_time_sdg Then
7020                    If CurrFieldName = "SDG_WORKFLOW_NAME" Then
      '                    sample_count = 0
      '                   max_sample_count = getWFSiblingCount(OriginSDG("SDG_WORKFLOW_ID"), 0)
7030                      Set XmlECreate = XmlDoc.createElement("create-by-workflow")
7040                      Set XmlEWF = XmlDoc.createElement("workflow-name")
7050                      XmlEWF.Text = OriginSDG.Fields(i)
7060                      Call XmlECreate.appendChild(XmlEWF)
7070                      Call XMLESDG.appendChild(XmlECreate)
7080                      Set xmltemp = XmlDoc.createElement("STATUS")
7090                      xmltemp.Text = "V"
7100                      Call XMLESDG.appendChild(xmltemp)
7110                    ElseIf CheckFieldOK(CurrFieldName) And _
                              (sdg_on And (CurrFieldName <> "SDG_ID" And _
                              CurrFieldName <> "SDG_WORKFLOW_ID" And CurrFieldName <> "EVENTS") Or _
                              (sdg_user_on And Mid(CurrFieldName, 1, 2) = "U_")) Then
7120                      Set xmltemp = XmlDoc.createElement(CurrFieldName)
7130                      xmltemp.Text = ManipulateLinkField(CurrFieldName, CurrFieldValue, "sdg")
                          'xmltemp.Text = TranslateFieldID(OriginSDG.Fields(i))
7140                      If xmltemp.Text <> "" Then
7150                        Call XMLESDG.appendChild(xmltemp)
7160                      End If
7170                    End If
7180                  ElseIf (sample_on Or sample_user_on And CurrFieldName <> "SAMPLE_ID") Then
7190                    If CurrFieldName = "SA_WORKFLOW_NAME" Then
7200                      sample_workflow_name = OriginSDG.Fields(i)
7210                    ElseIf CurrFieldName = "SAMPLE_ID" Then
7220                      old_sample_id = sample_id
7230                      sample_id = OriginSDG.Fields(i)
7240                      If old_sample_id <> sample_id Then
7250                        Set XMLESample = XmlDoc.createElement("SAMPLE")
                            
7260                        sample_count = sample_count + 1
7270                        If sample_count > max_sample_count Then
                              
7280                          Set XMLESDG = XmlDoc.createElement("SDG")
7290                          Set XmlEFind = XmlDoc.createElement("find-by-id")
7300                          XmlEFind.Text = main_sdg_id
7310                          Call XmlLoginReq.appendChild(XMLESDG)
7320                          Call XMLESDG.appendChild(XmlEFind)
7330                          Call XMLESDG.appendChild(XMLESample)
                              
7340                          Set XmlECreate = XmlDoc.createElement("create-by-workflow")
7350                          Set XmlEWF = XmlDoc.createElement("workflow-name")
7360                          XmlEWF.Text = sample_workflow_name
7370                          Call XmlECreate.appendChild(XmlEWF)
7380                          Call XMLESample.appendChild(XmlECreate)
                              
7390                          Set xmltemp = XmlDoc.createElement("STATUS")
7400                          xmltemp.Text = "V"
7410                          Call XMLESample.appendChild(xmltemp)
7420                        Else
7430                          Call XmlLoginReq.appendChild(XMLESample)
7440                          tempS = xmlSamples.selectNodes("//SAMPLE").Item(sample_count - 1).selectSingleNode("SAMPLE_ID").Text
7450                          Set XmlECreate = XmlDoc.createElement("find-by-id")
7460                          XmlECreate.Text = tempS
7470                          Call XMLESample.appendChild(XmlECreate)
7480                        End If
7490                      End If
      '                    If old_sample_id = -1 Then
      '                      old_sample_id = sample_id
      '                    End If
7500                    ElseIf CurrFieldName = "SAMPLE_NAME" Then
7510                      sample_name = OriginSDG.Fields(i)
7520                    ElseIf old_sample_id <> sample_id And sample_count > max_sample_count Then
7530                      If CheckFieldOK(CurrFieldName) And (sample_on And (CurrFieldName <> "EVENTS") And (CurrFieldName <> "SDG_ID") And (CurrFieldName <> "SA_WORKFLOW_ID") _
                            Or (sample_user_on And Mid(CurrFieldName, 1, 2) = "U_")) Then
7540                        Set xmltemp = XmlDoc.createElement(CurrFieldName)
7550                        xmltemp.Text = ManipulateLinkField(CurrFieldName, CurrFieldValue, "sample")
                            'xmltemp.Text = TranslateFieldID(OriginSDG.Fields(i))
7560                        If xmltemp.Text <> "" Then
7570                          Call XMLESample.appendChild(xmltemp)
7580                        End If
7590                      End If
7600                    End If
7610                  ElseIf (aliquot_on Or aliquot_user_on And CurrFieldName <> "ALIQUOT_ID") Then
7620                    If CurrFieldName = "A_WORKFLOW_NAME" Then
7630                      aliquot_workflow_name = OriginSDG.Fields(i)
7640                    ElseIf CurrFieldName = "ALIQUOT_ID" Then
7650                      old_aliquot_id = aliquot_id
7660                      aliquot_id = OriginSDG.Fields(i)
7670                      If old_aliquot_id <> aliquot_id Then
7680                        Set XmlEAliquot = XmlDoc.createElement("ALIQUOT")
                            
7690                        aliquot_count = aliquot_count + 1
7700                        If aliquot_count > max_aliquot_count Then
7710                          Set XMLESample = XmlDoc.createElement("SAMPLE")
7720                          Set XmlEFind = XmlDoc.createElement("find-by-id")
7730                          XmlEFind.Text = main_sample_id
7740                          Call XmlLoginReq.appendChild(XMLESample)
7750                          Call XMLESample.appendChild(XmlEFind)
7760                          Call XMLESample.appendChild(XmlEAliquot)
                              
7770                          Set XmlECreate = XmlDoc.createElement("create-by-workflow")
7780                          Set XmlEWF = XmlDoc.createElement("workflow-name")
7790                          XmlEWF.Text = aliquot_workflow_name
7800                          Call XmlECreate.appendChild(XmlEWF)
7810                          Call XmlEAliquot.appendChild(XmlECreate)
7820                          Set xmltemp = XmlDoc.createElement("STATUS")
7830                          xmltemp.Text = "V"
7840                          Call XmlEAliquot.appendChild(xmltemp)
7850                        Else
7860                          Call XmlLoginReq.appendChild(XmlEAliquot)
7870                          tempS = xmlAliquots.selectNodes("//ALIQUOT").Item(aliquot_count - 1).selectSingleNode("ALIQUOT_ID").Text
7880                          Set XmlECreate = XmlDoc.createElement("find-by-id")
7890                          XmlECreate.Text = tempS
7900                          Call XmlEAliquot.appendChild(XmlECreate)
7910                        End If
7920                      End If
7930                    ElseIf CurrFieldName = "ALIQUOT_NAME" Then
7940                      aliquot_name = OriginSDG.Fields(i)
7950                    ElseIf old_aliquot_id <> aliquot_id And (aliquot_count > max_aliquot_count Or CurrFieldName = "NAME") Then
7960                      If CheckFieldOK(CurrFieldName) And (aliquot_on And (CurrFieldName <> "EVENTS") And (CurrFieldName <> "SAMPLE_ID") And (CurrFieldName <> "A_WORKFLOW_ID") _
                            Or (aliquot_user_on And Mid(CurrFieldName, 1, 2) = "U_")) Then
7970                          Set xmltemp = XmlDoc.createElement(CurrFieldName)
7980                          xmltemp.Text = ManipulateLinkField(CurrFieldName, CurrFieldValue, "aliquot")
                              'xmltemp.Text = TranslateFieldID(OriginSDG.Fields(i))
7990                          If xmltemp.Text <> "" Then
8000                            Call XmlEAliquot.appendChild(xmltemp)
8010                          End If
8020                      End If
8030                    End If
8040                  ElseIf (test_on Or test_user_on And CurrFieldName <> "TEST_ID") Then
8050                    If CurrFieldName = "T_WORKFLOW_NAME" Then
8060                      test_workflow_name = OriginSDG.Fields(i)
8070                    ElseIf CurrFieldName = "TEST_ID" Then
8080                      old_test_id = test_id
8090                      test_id = OriginSDG.Fields(i)
8100                      If old_test_id <> test_id Then
8110                        Set XmlETest = XmlDoc.createElement("TEST")
8120                        test_count = test_count + 1
                            
8130                        If test_count > max_test_count Then
8140                          Set XmlEAliquot = XmlDoc.createElement("ALIQUOT")
8150                          Set XmlEFind = XmlDoc.createElement("find-by-id")
8160                          XmlEFind.Text = main_aliquot_id
8170                          Call XmlLoginReq.appendChild(XmlEAliquot)
8180                          Call XmlEAliquot.appendChild(XmlEFind)
8190                          Call XmlEAliquot.appendChild(XmlETest)
                              
8200                          Set XmlECreate = XmlDoc.createElement("create-by-workflow")
8210                          Set XmlEWF = XmlDoc.createElement("workflow-name")
8220                          XmlEWF.Text = test_workflow_name
8230                          Call XmlECreate.appendChild(XmlEWF)
8240                          Call XmlETest.appendChild(XmlECreate)
8250                          Set xmltemp = XmlDoc.createElement("STATUS")
8260                          xmltemp.Text = "V"
8270                          Call XmlETest.appendChild(xmltemp)
8280                        Else
8290                          Call XmlLoginReq.appendChild(XmlETest)
8300                          tempS = xmlTests.selectNodes("//TEST").Item(test_count - 1).selectSingleNode("TEST_ID").Text
8310                          Set XmlECreate = XmlDoc.createElement("find-by-id")
8320                          XmlECreate.Text = tempS
8330                          Call XmlETest.appendChild(XmlECreate)
8340                        End If
8350                      End If
8360                    ElseIf CurrFieldName = "TEST_NAME" Then
8370                      test_name = OriginSDG.Fields(i)
8380                    ElseIf old_test_id <> test_id And test_count > max_test_count Then
8390                      If CheckFieldOK(CurrFieldName) And (test_on And (CurrFieldName <> "EVENTS") And (CurrFieldName <> "ALIQUOT_ID") And (CurrFieldName <> "T_WORKFLOW_ID") _
                            Or (aliquot_user_on And Mid(CurrFieldName, 1, 2) = "U_")) Then
8400                          Set xmltemp = XmlDoc.createElement(CurrFieldName)
8410                          xmltemp.Text = ManipulateLinkField(CurrFieldName, CurrFieldValue, "test")
                              'xmltemp.Text = TranslateFieldID(OriginSDG.Fields(i))
8420                          If CurrFieldName <> "NAME" And xmltemp.Text <> "" Then
8430                            Call XmlETest.appendChild(xmltemp)
8440                          End If
8450                      End If
8460                    End If
8470                  ElseIf (result_on Or result_user_on And CurrFieldName <> "EVENTS" And CurrFieldName <> "RESULT_ID") Then
8480                    If CurrFieldName = "R_WORKFLOW_NAME" Then
8490                      result_workflow_name = OriginSDG.Fields(i)
8500                    ElseIf CurrFieldName = "R_CALCULATION_ID" Then
8510                      r_calculation_id = CheckFieldValue(OriginSDG.Fields(i))
8520                    ElseIf CurrFieldName = "R_RESULT_TYPE" Then
8530                      r_result_type = CheckFieldValue(OriginSDG.Fields(i))
8540                    ElseIf r_calculation_id = "" Then
8550                        If CurrFieldName = "RESULT_ID" Then
8560                          old_result_id = result_id
8570                          result_id = OriginSDG.Fields(i)
8580                          If old_result_id <> result_id Then ' Always
8590                            Set XMLEResult = XmlDoc.createElement("RESULT")
8600                            Call XmlLoginReq.appendChild(XMLEResult)
                                
8610                            result_count = result_count + 1
8620                            If result_count > max_result_count Then
8630                              Set XmlETest = XmlDoc.createElement("TEST")
8640                              Set XmlEFind = XmlDoc.createElement("find-by-id")
8650                              XmlEFind.Text = main_test_id
8660                              Call XmlLoginReq.appendChild(XmlETest)
8670                              Call XmlETest.appendChild(XmlEFind)
8680                              Call XmlETest.appendChild(XMLEResult)
                                  
8690                              Set XmlECreate = XmlDoc.createElement("create-by-workflow")
8700                              Set XmlEWF = XmlDoc.createElement("workflow-name")
8710                              XmlEWF.Text = result_workflow_name
8720                              Call XmlECreate.appendChild(XmlEWF)
8730                              Call XMLEResult.appendChild(XmlECreate)
8740                              Set xmltemp = XmlDoc.createElement("STATUS")
8750                              xmltemp.Text = "V"
8760                              Call XMLEResult.appendChild(xmltemp)
8770                            Else
8780                              Call XmlLoginReq.appendChild(XMLEResult)
8790                              tempS = xmlresults.selectNodes("//RESULT").Item(result_count - 1).selectSingleNode("RESULT_ID").Text
8800                              Set XmlECreate = XmlDoc.createElement("find-by-id")
8810                              XmlECreate.Text = tempS
8820                              Call XMLEResult.appendChild(XmlECreate)
8830                            End If
8840                          End If
8850                        ElseIf (CurrFieldName = "ORIGINAL_RESULT") Or (CurrFieldName = "NAME") Then
      '                      ElseIf result_count > max_result_count Then
8860                          Set xmltemp = XmlDoc.createElement(CurrFieldName)
8870                          xmltemp.Text = ManipulateLinkField(CurrFieldName, CurrFieldValue, "result")
                              'xmltemp.Text = TranslateFieldID(OriginSDG.Fields(i))
8880                          If xmltemp.Text = "" Then
8890                            If r_result_type = "B" Then
8900                              xmltemp.Text = "F"
8910                            ElseIf r_result_type = "N" Then
8920                              xmltemp.Text = "0"
8930                            End If
8940                          End If
8950                          Call XMLEResult.appendChild(xmltemp)
8960                        End If
8970                      End If
8980                  End If
8990               End Select
9000              Next i
9010              OriginSDG.MoveNext
9020            Wend
            
          '    aConnection.BeginTrans
              
      '        sample_name = OriginSDG("sample_name")
                
          '    aConnection.CommitTrans
              
9030          Set Errs = Xmlres.getElementsByTagName("error")
9040          If Errs.length > 0 Then
9050            MsgBox "Error : " & Errs.Item(0).Text
9060          Else
          '      aConnection.BeginTrans
      '          Call RenameNewNames
      '          Call CreateResults
          '      aConnection.CommitTrans
9070          End If
                      
9080          Set Xmlres = Nothing
9090          Set Xmlres = New DOMDocument
      '        XmlresVals.Save ("c:\resvals.xml")
9100          Call ProcessXML.ProcessXMLWithResponse(XmlresVals, Xmlres)
      '        Xmlres.Save ("c:\xmlres.xml")
      '        MsgBox 1
      '        MsgBox sdg_id & " - " & frmMain.edtBarCode.Text
9110          Call CopyRTF(CStr(sdg_id), frmMain.edtBarCode.Text)
              
9120          sql = "update lims_sys.sdg_user set u_revision_cause = '" & revision_cause & "' " _
                & "where sdg_id = (select sdg_id from lims_sys.sdg where name = '" & frmMain.edtBarCode.Text & "') "
9130          aConnection.Execute (sql)

9140          sdg_log_desc = ""
9150          log_sdg_id = sdg_id
9160          Call sdg_log.InsertLog(log_sdg_id, "REV.UPD", sdg_log_desc)

      '        Set Errs = Xmlres.getElementsByTagName("error")
      '        If Errs.length > 0 Then
      '          MsgBox "Error : " & Errs.Item(0).Text
      '        End If
      '        Xmlres.Save ("c:\sdgerr.xml")
              
9170        End If
9180        OriginSDG.Close
9190    End If
End Function

'declare
'  cursor cc is select client_id from client;
'  savei number(16); _
'  i number(16); _
'  y rowid; _
'begin _
'  savei:= 9999; _
'  for x in cc loop _
'    begin _
'      If savei <> 9999 Then _
'        delete from client where client_id = savei; _
'        commit; _
'      End If; _
'      savei:= x.i; _
'    exception _
'      when others then rollback; _
'    End; _
'  end loop; _
'End;

Private Sub CopyRTF(oldid As String, newname As String)
          Dim OldresultRst As ADODB.Recordset
          Dim NewresultRst As ADODB.Recordset
          Dim RtfResult As New ADODB.Recordset
          Dim mStream As New ADODB.Stream
          Dim ClobRst As New ADODB.Recordset
9200      mStream.Type = adTypeText
          Dim sql As String
          Dim ResTmp As String
          Dim ResID As String

      '    sql = "select nr.result_id nrid, olr.result_id orid " & _
              "from sdg nd, sample os, sample ns, aliquot oa, aliquot na, " & _
              "test ot, test nt, result olr, result nr, rtf_result " & _
              "where olr.test_id = ot.test_id and " & _
              "ot.aliquot_id = oa.aliquot_id and " & _
              "oa.sample_id = os.sample_id and " & _
              "nr.test_id = nt.test_id and " & _
              "nt.aliquot_id = na.aliquot_id and " & _
              "na.sample_id = ns.sample_id and " & _
              "ns.sdg_id = nd.sdg_id and " & _
              "os.sdg_id = " & oldid & " and " & _
              "nd.name = '" & newname & "' and " & _
              "olr.workflow_node_id = nr.workflow_node_id and " & _
              "olr.result_id = rtf_result.rtf_result_id"

9210      sql = "select r.result_id rid, r.result_template_id tempid " & _
                "from lims_sys.sample s, lims_sys.aliquot a, lims_sys.test t, " & _
                "lims_sys.result r, lims_sys.rtf_result rtf " & _
                "where s.sdg_id = " & oldid & " and " & _
                "a.sample_id = s.sample_id and " & _
                "t.aliquot_id = a.aliquot_id and " & _
                "r.test_id = t.test_id and " & _
                "r.result_id = rtf.rtf_result_id " & _
                "order by r.result_template_id"
      '    MsgBox sql
9220      Set OldresultRst = aConnection.Execute(sql)

9230      sql = "select r.result_id rid, r.result_template_id tempid " & _
                "from lims_sys.sdg d, lims_sys.sample s, lims_sys.aliquot a, " & _
                "lims_sys.test t, lims_sys.result r " & _
                "where d.name = '" & newname & "' and s.sdg_id = d.sdg_id and " & _
                "a.sample_id = s.sample_id and " & _
                "t.aliquot_id = a.aliquot_id and " & _
                "r.test_id = t.test_id " & _
                "order by r.result_template_id"
      '    MsgBox sql
9240      Set NewresultRst = aConnection.Execute(sql)

9250      While (Not OldresultRst.EOF)

9260          While OldresultRst("TEMPID") <> NewresultRst("TEMPID")
9270              NewresultRst.MoveNext
9280          Wend

      '        MsgBox "oldid: " & OldresultRst("RID") & " newid: " & NewresultRst("RID")
9290          ResID = OldresultRst("RID")
9300          Call RtfResult.Open _
                  ("select rtf_text from lims_sys.rtf_result where rtf_result_id = " & _
                  ResID, aConnection, adOpenStatic, adLockOptimistic)
9310          ResTmp = ""
9320          If Not RtfResult.EOF Then
      '        MsgBox 2
9330              ResTmp = ReadClob(RtfResult("RTF_TEXT"))
9340              If ResTmp <> "" Then
      '            MsgBox 3
9350                  Call SaveResultRTF(NewresultRst("RID"), ResTmp)
9360              End If
9370          End If
9380          RtfResult.Close

      '        Call ClobRst.Open("select rtf_text from lims_sys.rtf_result " & _
      '            "where rtf_result_id = " & resultRst("ORID"), aConnection)
      '        mStream.WriteText ClobRst(0).Value
      '        Call aConnection.Execute("insert into lims_sys.rtf_result (rtf_result_id) " & _
      '            "values (" & resultRst("NRID") & ")")
      '        Call ClobRst.Close
      '        Call ClobRst.Open("select rtf_text from lims_sys.rtf_result " & _
      '            "where rtf_result_id = " & resultRst("NRID"), aConnection)
      '        ClobRst(0).Value = mStream.ReadText
      '        ClobRst.Update

9390          OldresultRst.MoveNext

      '        NewresultRst.MoveNext
9400      Wend
End Sub

Private Function ReadClob(fld As ADODB.Field) As String
          Dim Data As String, Temp As Variant
9410      Do
9420          Temp = fld.GetChunk(BLOCK_SIZE)
9430          If IsNull(Temp) Then Exit Do
9440          Data = Data & Temp
9450      Loop While Len(Temp) = BLOCK_SIZE
9460      ReadClob = Data
End Function

Private Sub SaveResultRTF(RtfResultId As String, ResultVal As String)
          Dim RtfResult As New ADODB.Recordset
          Dim ResSTR As String

9470      ResSTR = "select rtf_result_id from lims_sys.rtf_result " & _
                   "where rtf_result_id = " & RtfResultId

9480      Set RtfResult = aConnection.Execute(ResSTR)

9490      If RtfResult.EOF Then
9500          ResSTR = "insert into lims_sys.rtf_result (rtf_result_id) values ('" & _
                        RtfResultId & "')"
9510          Call aConnection.Execute(ResSTR)
9520      End If
9530      RtfResult.Close

9540      Call RtfResult.Open("select rtf_text from lims_sys.rtf_result where rtf_result_id = " & RtfResultId, aConnection, adOpenStatic, adLockOptimistic)

9550      Call RtfResult("RTF_TEXT").AppendChunk(ResultVal)
End Sub

Private Function checkNum(s As String) As Boolean
          Dim i As Long
9560      On Error GoTo ErrNum
          
9570      i = CLng(s)
9580      checkNum = True
9590  Exit Function
ErrNum:
9600      checkNum = False
End Function

Public Function CheckFieldValue(f As ADODB.Field) As Variant
9610      If IsNull(f) Then
9620          CheckFieldValue = ""
9630      Else
9640          CheckFieldValue = Trim(f.Value)
9650      End If
End Function

Public Function CheckApostrophe(s As String)
9660      CheckApostrophe = Replace(s, "'", "''")
End Function

Public Function nte(e As Variant) As Variant
9670      nte = IIf(IsNull(e), "", e)
End Function

Private Function ntz(e As Variant) As Variant
9680      ntz = IIf(IsNull(e), 0, e)
End Function


Private Function LoginNewSdg(EventName As String, _
                             strExternalReference As String, _
                             strWorkflowName As String) As String
                                   
9690  On Error GoTo ERR_LoginNewSdg

          Dim doc As New DOMDocument
          Dim res As New DOMDocument
          Dim xmlLogin As IXMLDOMElement
          Dim xmlSdg As IXMLDOMElement
          Dim xmlEL As IXMLDOMElement
          Dim e As IXMLDOMElement
          Dim element As IXMLDOMElement
          Dim XmlECreateByWorkflow As IXMLDOMElement
          Dim XmlEWFName As IXMLDOMElement
          Dim XmlExternal As IXMLDOMElement
          Dim FileName As String
          Dim RetError As String


9700      Set xmlEL = doc.createElement("lims-request")
9710      Call doc.appendChild(xmlEL)

9720      Set e = doc.createElement("login-request")
9730      Call xmlEL.appendChild(e)
          
          'Set xmlLogin = doc.createElement("login-request")
          'Call e.appendChild(xmlLogin)
9740      Set xmlSdg = doc.createElement("SDG")
9750      Call e.appendChild(xmlSdg)
          
9760      Set XmlECreateByWorkflow = doc.createElement("create-by-workflow")
9770      Call xmlSdg.appendChild(XmlECreateByWorkflow)
          
9780      Set XmlEWFName = doc.createElement("workflow-name")
9790      XmlEWFName.Text = strWorkflowName
9800      Call XmlECreateByWorkflow.appendChild(XmlEWFName)
          
9810      Set XmlExternal = doc.createElement("EXTERNAL_REFERENCE")
9820      XmlExternal.Text = strExternalReference
          
9830      Call xmlSdg.appendChild(XmlExternal)
          'XmlEWF.Text = OriginSDG.Fields(i)
          
          
          
          'Set element = doc.createElement("find-by-id")
          'element.Text = AliquotID
      'Call xmlSdg.appendChild(element)
          'Set element = doc.createElement("login-request")
          'Set element = doc.createElement("fire-event")
          'element.Text = "Login SDG"
          'Call xmlSdg.appendChild(xmlLogin)

       '   If Trim(WorkFolder) <> "" Then
             ' FileName = "C:\SdgCOPY_DOC1"
             ' Call xmlManager.SaveXmlFile(doc, FileName)
       '   End If

9840      RetError = ProcessXML.ProcessXMLWithResponse(doc, res)
9850      If Trim(RetError) <> "" Then
9860          MsgBox "Error occurred while trying process xml file. " & vbCrLf & _
                     "Event Name: " & EventName & vbCrLf & _
                     "Error: " & RetError, vbCritical, "Nautilus - Sdg Copy"
9870      End If

       '   If Trim(WorkFolder) <> "" Then
             ' FileName = "C:\SdgCopy_RES1"
             ' Call xmlManager.SaveXmlFile(res, FileName)
       '   End If

9880       LoginNewSdg = res.selectSingleNode("//SDG_ID").Text

9890      Exit Function
ERR_LoginNewSdg:
9900  MsgBox "Error on line:" & Erl & " in LoginNewSdg" & vbCrLf & Err.Description
End Function


Private Function LoginNewSample(EventName As String, _
                                strExternalReference As String, _
                                strWorkflowName As String, _
                                sdgId As String) As String
                                   
9910  On Error GoTo ERR_LoginNewSample

          Dim doc As New DOMDocument
          Dim res As New DOMDocument
          Dim xmlSdg As IXMLDOMElement
          Dim xmlFind As IXMLDOMElement
          Dim xmlLogin As IXMLDOMElement
          Dim xmlSample As IXMLDOMElement
          Dim xmlEL As IXMLDOMElement
          Dim e As IXMLDOMElement
          Dim element As IXMLDOMElement
          Dim XmlECreateByWorkflow As IXMLDOMElement
          Dim XmlEWFName As IXMLDOMElement
          Dim XmlExternal As IXMLDOMElement
          Dim FileName As String
          Dim RetError As String


9920      Set xmlEL = doc.createElement("lims-request")
9930      Call doc.appendChild(xmlEL)

9940      Set e = doc.createElement("login-request")
9950      Call xmlEL.appendChild(e)
          
          'Set xmlLogin = doc.createElement("login-request")
          'Call e.appendChild(xmlLogin)
          
          
          'SDG:
9960      Set xmlSdg = doc.createElement("SDG")
9970      Call e.appendChild(xmlSdg)
          
          'find by id:
9980      Set xmlFind = doc.createElement("find-by-id")
9990      xmlFind.Text = sdgId
10000     Call xmlSdg.appendChild(xmlFind)
          
10010     Set xmlSample = doc.createElement("SAMPLE")
10020     Call xmlSdg.appendChild(xmlSample)
          
10030     Set XmlECreateByWorkflow = doc.createElement("create-by-workflow")
10040     Call xmlSample.appendChild(XmlECreateByWorkflow)
          
10050     Set XmlEWFName = doc.createElement("workflow-name")
10060     XmlEWFName.Text = strWorkflowName
10070     Call XmlECreateByWorkflow.appendChild(XmlEWFName)
          
10080     Set XmlExternal = doc.createElement("EXTERNAL_REFERENCE")
10090     XmlExternal.Text = strExternalReference
          
10100     Call xmlSample.appendChild(XmlExternal)
          'XmlEWF.Text = OriginSDG.Fields(i)
          
          
          
          'Set element = doc.createElement("find-by-id")
          'element.Text = AliquotID
      'Call xmlSdg.appendChild(element)
          'Set element = doc.createElement("login-request")
          'Set element = doc.createElement("fire-event")
          'element.Text = "Login SDG"
          'Call xmlSdg.appendChild(xmlLogin)

       '   If Trim(WorkFolder) <> "" Then
             ' FileName = "C:\SdgCOPY_createSample_DOC1"
             ' Call xmlManager.SaveXmlFile(doc, FileName)
       '   End If

10110     RetError = ProcessXML.ProcessXMLWithResponse(doc, res)
10120     If Trim(RetError) <> "" Then
10130         MsgBox "Error occurred while trying process xml file. " & vbCrLf & _
                     "Event Name: " & EventName & vbCrLf & _
                     "Error: " & RetError, vbCritical, "Nautilus - Sdg Copy"
10140     End If

       '   If Trim(WorkFolder) <> "" Then
             ' FileName = "C:\SdgCopy_createSample_RES1"
             ' Call xmlManager.SaveXmlFile(res, FileName)
       '   End If

10150      LoginNewSample = res.selectSingleNode("//SAMPLE_ID").Text

10160     Exit Function
ERR_LoginNewSample:
10170 MsgBox "Error on line:" & Erl & " in LoginNewSample" & vbCrLf & Err.Description
End Function


Private Function LoginNewBlock(EventName As String, _
                               strExternalReference As String, _
                               strWorkflowName As String, _
                               strSampleId As String) As String
                                   
10180 On Error GoTo ERR_LoginNewBlock

          Dim doc As New DOMDocument
          Dim res As New DOMDocument
          Dim xmlLogin As IXMLDOMElement
          Dim xmlSample As IXMLDOMElement
          Dim xmlFind As IXMLDOMElement
          Dim xmlBlock As IXMLDOMElement
          Dim xmlEL As IXMLDOMElement
          Dim e As IXMLDOMElement
          Dim element As IXMLDOMElement
          Dim XmlECreateByWorkflow As IXMLDOMElement
          Dim XmlEWFName As IXMLDOMElement
          Dim XmlExternal As IXMLDOMElement
          Dim FileName As String
          Dim RetError As String

10190     Set xmlEL = doc.createElement("lims-request")
10200     Call doc.appendChild(xmlEL)

10210     Set e = doc.createElement("login-request")
10220     Call xmlEL.appendChild(e)
          
          'Set xmlLogin = doc.createElement("login-request")
          'Call e.appendChild(xmlLogin)
          
          
          'sample:
10230     Set xmlSample = doc.createElement("SAMPLE")
10240     Call e.appendChild(xmlSample)
          
          'find by id:
10250     Set xmlFind = doc.createElement("find-by-id")
10260     xmlFind.Text = strSampleId
10270     Call xmlSample.appendChild(xmlFind)
          
10280     Set xmlBlock = doc.createElement("ALIQUOT")
10290     Call xmlSample.appendChild(xmlBlock)
          
10300     Set XmlECreateByWorkflow = doc.createElement("create-by-workflow")
10310     Call xmlBlock.appendChild(XmlECreateByWorkflow)
          
10320     Set XmlEWFName = doc.createElement("workflow-name")
10330     XmlEWFName.Text = strWorkflowName
10340     Call XmlECreateByWorkflow.appendChild(XmlEWFName)
          
10350     Set XmlExternal = doc.createElement("EXTERNAL_REFERENCE")
10360     XmlExternal.Text = strExternalReference
          
10370     Call xmlBlock.appendChild(XmlExternal)
          'XmlEWF.Text = OriginSDG.Fields(i)

       '   If Trim(WorkFolder) <> "" Then
             ' FileName = "C:\SdgCOPY_createBlock_DOC1"
             ' Call xmlManager.SaveXmlFile(doc, FileName)
       '   End If

10380     RetError = ProcessXML.ProcessXMLWithResponse(doc, res)
10390     If Trim(RetError) <> "" Then
10400         MsgBox "Error occurred while trying process xml file. " & vbCrLf & _
                     "Event Name: " & EventName & vbCrLf & _
                     "Error: " & RetError, vbCritical, "Nautilus - Sdg Copy"
10410     End If

       '   If Trim(WorkFolder) <> "" Then
             ' FileName = "C:\SdgCopy_createBlock_RES1"
             ' Call xmlManager.SaveXmlFile(res, FileName)
       '   End If

10420      LoginNewBlock = res.selectSingleNode("//ALIQUOT_ID").Text

10430     Exit Function
ERR_LoginNewBlock:
10440 MsgBox "Error on line:" & Erl & " in LoginNewBlock" & vbCrLf & Err.Description
End Function


Private Function LoginNewTest(EventName As String, _
                              strExternalReference As String, _
                              strWorkflowName As String, _
                              strAliquotId As String)
                                   
10450 On Error GoTo ERR_LoginNewTest

          Dim doc As New DOMDocument
          Dim res As New DOMDocument
          Dim xmlLogin As IXMLDOMElement
          Dim xmlAliquot As IXMLDOMElement
          Dim xmlFind As IXMLDOMElement
          Dim xmlTest As IXMLDOMElement
          Dim xmlEL As IXMLDOMElement
          Dim e As IXMLDOMElement
          Dim element As IXMLDOMElement
          Dim XmlECreateByWorkflow As IXMLDOMElement
          Dim XmlEWFName As IXMLDOMElement
          Dim XmlExternal As IXMLDOMElement
          Dim FileName As String
          Dim RetError As String

10460     Set xmlEL = doc.createElement("lims-request")
10470     Call doc.appendChild(xmlEL)

10480     Set e = doc.createElement("login-request")
10490     Call xmlEL.appendChild(e)
          
          'Set xmlLogin = doc.createElement("login-request")
          'Call e.appendChild(xmlLogin)
          
          
          'sample:
10500     Set xmlAliquot = doc.createElement("ALIQUOT")
10510     Call e.appendChild(xmlAliquot)
          
          'find by id:
10520     Set xmlFind = doc.createElement("find-by-id")
10530     xmlFind.Text = strAliquotId
10540     Call xmlAliquot.appendChild(xmlFind)
          
10550     Set xmlTest = doc.createElement("TEST")
10560     Call xmlAliquot.appendChild(xmlTest)
          
10570     Set XmlECreateByWorkflow = doc.createElement("create-by-workflow")
10580     Call xmlTest.appendChild(XmlECreateByWorkflow)
          
10590     Set XmlEWFName = doc.createElement("workflow-name")
10600     XmlEWFName.Text = strWorkflowName
10610     Call XmlECreateByWorkflow.appendChild(XmlEWFName)
          
      '    Set XmlExternal = doc.createElement("EXTERNAL_REFERENCE")
      '    XmlExternal.Text = strExternalReference
      '    Call xmlTest.appendChild(XmlExternal)
          
          'XmlEWF.Text = OriginSDG.Fields(i)

       '   If Trim(WorkFolder) <> "" Then
            '  FileName = "C:\SdgCOPY_createTest_DOC1"
            '  Call xmlManager.SaveXmlFile(doc, FileName)
       '   End If

10620     RetError = ProcessXML.ProcessXMLWithResponse(doc, res)
10630     If Trim(RetError) <> "" Then
10640         MsgBox "Error occurred while trying process xml file. " & vbCrLf & _
                     "Event Name: " & EventName & vbCrLf & _
                     "Error: " & RetError, vbCritical, "Nautilus - Sdg Copy"
10650     End If

       '   If Trim(WorkFolder) <> "" Then
            '  FileName = "C:\SdgCopy_createTest_RES1"
            '  Call xmlManager.SaveXmlFile(res, FileName)
       '   End If

      '     LoginNewTest = res.selectSingleNode("//TEST_ID").Text

10660     Exit Function
ERR_LoginNewTest:
10670 MsgBox "Error on line:" & Erl & " in LoginNewTest" & vbCrLf & Err.Description
End Function


Private Function LoginNewResult(EventName As String, _
                                strTestId As String) As String
                                   
10680 On Error GoTo ERR_LoginNewResult

          Dim doc As New DOMDocument
          Dim res As New DOMDocument
          Dim xmlLogin As IXMLDOMElement
          Dim xmlTest As IXMLDOMElement
          Dim xmlFind As IXMLDOMElement
      '    Dim xmlResult As IXMLDOMElement
          Dim xmlEL As IXMLDOMElement
          Dim e As IXMLDOMElement
          Dim element As IXMLDOMElement
      '    Dim XmlECreateByWorkflow As IXMLDOMElement
      '    Dim XmlEWFName As IXMLDOMElement
      '    Dim XmlExternal As IXMLDOMElement
          Dim xmlFireEvent As IXMLDOMElement
          
          Dim FileName As String
          Dim RetError As String

10690     Set xmlEL = doc.createElement("lims-request")
10700     Call doc.appendChild(xmlEL)

10710     Set e = doc.createElement("login-request")
10720     Call xmlEL.appendChild(e)
          
10730     Set xmlTest = doc.createElement("TEST")
10740     Call e.appendChild(xmlTest)
          
          'find by id:
10750     Set xmlFind = doc.createElement("find-by-id")
10760     xmlFind.Text = strTestId
10770     Call xmlTest.appendChild(xmlFind)
          
10780     Set xmlFireEvent = doc.createElement("fire-event")
10790     xmlFireEvent.Text = EventName
10800     Call xmlTest.appendChild(xmlFireEvent)


       '   If Trim(WorkFolder) <> "" Then
       '       FileName = "C:\SdgCOPY_createResult_DOC1"
       '       Call xmlManager.SaveXmlFile(doc, FileName)
       '   End If

10810     RetError = ProcessXML.ProcessXMLWithResponse(doc, res)
10820     If Trim(RetError) <> "" Then
10830         MsgBox "Error occurred while trying process xml file. " & vbCrLf & _
                     "Event Name: " & EventName & vbCrLf & _
                     "Error: " & RetError, vbCritical, "Nautilus - Sdg Copy"
10840     End If

       '   If Trim(WorkFolder) <> "" Then
       '       FileName = "C:\SdgCopy_createResult_RES1"
       '       Call xmlManager.SaveXmlFile(res, FileName)
       '   End If

10850      LoginNewResult = res.selectSingleNode("//RESULT_ID").Text

10860     Exit Function
ERR_LoginNewResult:
10870 MsgBox "Error on line:" & Erl & " in LoginNewResult" & vbCrLf & Err.Description
End Function


Private Function LoginNewSlide(strEventName As String, strBlockId As String) 'As String
10880     On Error GoTo Err_LoginNewSlide
          
          Dim doc As New DOMDocument
          Dim res As New DOMDocument
          Dim xmlLogin As IXMLDOMElement
          Dim xmlBlock As IXMLDOMElement
          Dim e As IXMLDOMElement
          Dim element As IXMLDOMElement
          Dim FileName As String
          Dim RetError As String

10890     Set e = doc.createElement("lims-request")
10900     Call doc.appendChild(e)
10910     Set xmlLogin = doc.createElement("login-request")
10920     Call e.appendChild(xmlLogin)
10930     Set xmlBlock = doc.createElement("ALIQUOT")
10940     Call xmlLogin.appendChild(xmlBlock)
10950     Set element = doc.createElement("find-by-id")
10960     element.Text = strBlockId
10970     Call xmlBlock.appendChild(element)
10980     Set element = doc.createElement("fire-event")
10990     element.Text = strEventName
11000     Call xmlBlock.appendChild(element)

11010     If Trim(WorkFolder) <> "" Then
             ' FileName = "C:\SDgCopy_" & strEventName & "_" & strBlockId & "_DOC1"
             ' Call xmlManager.SaveXmlFile(doc, FileName)
11020     End If

11030     RetError = ProcessXML.ProcessXMLWithResponse(doc, res)
11040     If Trim(RetError) <> "" Then
11050         MsgBox "Error occurred while trying process xml file. (TriggerSlideEvent) " & vbCrLf & _
                     "Block ID: " & strBlockId & vbCrLf & _
                     "Event Name: " & strEventName & vbCrLf & _
                     "Error: " & RetError, vbCritical, "Nautilus - Sdg Copy"
11060     End If

11070     If Trim(WorkFolder) <> "" Then
             ' FileName = "SdgCopy_" & strEventName & "_" & strBlockId & "_RES1"
             ' Call xmlManager.SaveXmlFile(res, FileName)
11080     End If

          'LoginNewSlide = res.selectSingleNode("//ALIQUOT_ID").Text
11090     Exit Function

Err_LoginNewSlide:
11100     MsgBox "Error on line:" & Erl & " in LoginNewSlide" & vbCrLf & _
                  "Block ID = " & strBlockId & vbCrLf & _
                  "Event Name = " & strEventName & vbCrLf & _
                  Err.Description
End Function

'used for getting the latest slide created for this block
'after the login of a new slide (the login action doesn't return the slide id)
Private Function GetMaxSlide(ParentID As String) As String
11110 On Error GoTo Err_GetMaxSlide
          Dim strSql As String
          Dim SlideRec As ADODB.Recordset

11120     GetMaxSlide = 0
11130     strSql = "select max(a.aliquot_id) " & _
                   "from lims_sys.aliquot a " & _
                   "where a.aliquot_id in " & _
                      "(select child_aliquot_id from lims_sys.aliquot_formulation " & _
                  "where aliquot_formulation.parent_aliquot_id = '" & ParentID & "') " & _
                  "order by a.aliquot_id"
11140     Set SlideRec = aConnection.Execute(strSql)

11150     If Not SlideRec.EOF Then
11160         GetMaxSlide = SlideRec(0)
11170     End If
11180     SlideRec.Close
11190     Exit Function
          
Err_GetMaxSlide:
11200     MsgBox "Error on line:" & Erl & " in GetMaxSlide... " & vbCrLf & _
                  "Parent Aliquot ID = " & ParentID & vbCrLf & _
                  Err.Description
End Function


Private Function GetMaxTest(strAliquotId As String) As String
11210 On Error GoTo ERR_GetMaxTest
          
          Dim rs As Recordset
          Dim sql As String
          
11220     sql = "  select max(t.test_id)"
11230     sql = sql & "  from lims_sys.test t"
11240     sql = sql & "  where t.ALIQUOT_ID = " & strAliquotId
          
11250     Set rs = aConnection.Execute(sql)
11260     GetMaxTest = nte(rs(0))

11270     Exit Function
ERR_GetMaxTest:
11280 MsgBox "Error on line:" & Erl & " in GetMaxTest" & vbCrLf & Err.Description
End Function

Private Function isSdgValidForRevision(strSdgId As String) As Boolean
          Dim rs As Recordset
          Dim sql As String
          
11290     isSdgValidForRevision = False
          
11300     sql = " select status "
11310     sql = sql & " from lims_sys.sdg "
11320     sql = sql & " where sdg_id = " & strSdgId
11330     sql = sql & " and status in ('R','A') "
          
11340     Set rs = aConnection.Execute(sql)
          
11350     If Not rs.EOF Then
11360         isSdgValidForRevision = True
11370     End If
End Function


