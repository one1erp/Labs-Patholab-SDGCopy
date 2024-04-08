Attribute VB_Name = "modSDGVars"
Option Explicit


Public iRevisionNum As Integer

Public Const DEFINE_DEBUG = False
Public WorkFolder As String

Public aConnection As New ADODB.Connection
Public ProcessXML As LSSERVICEPROVIDERLib.NautilusProcessXML
Public NtlsCon As LSSERVICEPROVIDERLib.NautilusDBConnection

Public strSdgId As String
Public sdg_id As Double
Public sdg_name As String

Public sample_id As Double
Public aliquot_id As Double
Public test_id As Double
Public result_id As Double

Public old_sdg_id As Double
Public old_sample_id As Double
Public old_aliquot_id As Double
Public old_test_id As Double
Public old_result_id As Double

Public testNode As IXMLDOMNode
Public resultNode As IXMLDOMNode
Public XMLCopy As New DOMDocument
Public XmlDoc As DOMDocument
Public Xmlres As New DOMDocument
Public XmlresVals As New DOMDocument
Public XMLResLims As IXMLDOMElement
Public XMLResRequest As IXMLDOMElement
Public XmlELims As IXMLDOMElement
Public XmlLoginReq As IXMLDOMElement
Public XMLESDG As IXMLDOMElement
Public XMLESample As IXMLDOMElement
Public XmlEAliquot As IXMLDOMElement
Public XmlETest As IXMLDOMElement
Public XMLEResult As IXMLDOMElement
Public XmlECreate As IXMLDOMElement
Public XmlEFind As IXMLDOMElement
Public XmlEWF As IXMLDOMElement

Public Errs As IXMLDOMNodeList

Public VersionName As String

Public sample_count As Integer
Public aliquot_count As Integer
Public test_count As Integer
Public result_count As Integer

Public max_sample_count As Integer
Public max_aliquot_count As Integer
Public max_test_count As Integer
Public max_result_count As Integer

Public main_sdg_id As Double
Public main_sample_id As Double
Public main_aliquot_id As Double
Public main_test_id As Double

Public sample_name As String
Public aliquot_name As String
Public test_name As String
Public result_name As String
Public revision_cause As String

Public xmlSamples As DOMDocument
Public xmlAliquots As DOMDocument
Public xmlTests As DOMDocument
Public xmlresults As DOMDocument

Public all_revision_codes As New Dictionary

Public gboSampleNameChanged As Boolean
Public gboAliquotNameChanged As Boolean

'holds the names of all the fields NOT to be copied
'when creating the request-tree-copy:
Public dicStopListFields As Dictionary
