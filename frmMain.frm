VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SDG Copy"
   ClientHeight    =   5580
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7215
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   177
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbRevisionCause 
      Height          =   420
      Left            =   3000
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   4200
      Width           =   3855
   End
   Begin VB.Frame Frame2 
      Caption         =   "SDG "
      Height          =   1935
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   6855
      Begin VB.TextBox edtBarCode 
         Height          =   405
         Left            =   2880
         TabIndex        =   0
         Top             =   360
         Width           =   3735
      End
      Begin VB.TextBox edtExternalReference 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   405
         Left            =   2880
         TabIndex        =   2
         Top             =   1320
         Width           =   3735
      End
      Begin VB.TextBox edtOriginalName 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   405
         Left            =   2880
         TabIndex        =   1
         Top             =   840
         Width           =   3735
      End
      Begin VB.Label lblSDGName 
         Caption         =   "Barcode :"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   435
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "External Reference :"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1395
         Width           =   2535
      End
      Begin VB.Label lblSDGOriginalName 
         Caption         =   "Original Name :"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   915
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Client "
      Height          =   1935
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   6855
      Begin VB.TextBox edtLastName 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   405
         Left            =   2880
         TabIndex        =   4
         Top             =   840
         Width           =   3735
      End
      Begin VB.TextBox edtFirstName 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   405
         Left            =   2880
         TabIndex        =   5
         Top             =   1320
         Width           =   3735
      End
      Begin VB.TextBox edtIDCard 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   435
         Left            =   2880
         TabIndex        =   3
         Top             =   360
         Width           =   3735
      End
      Begin VB.Label Label2 
         Caption         =   "Last Name :"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   915
         Width           =   2535
      End
      Begin VB.Label Label3 
         Caption         =   "First Name :"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1395
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "ID Card :"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   450
         Width           =   2415
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   615
      Left            =   2160
      TabIndex        =   8
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Revision cause :"
      Height          =   495
      Left            =   240
      TabIndex        =   17
      Top             =   4163
      Width           =   2415
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private sp As LSSERVICEPROVIDERLib.NautilusServiceProvider



Public Sub Initialize(sp_ As LSSERVICEPROVIDERLib.NautilusServiceProvider)
11380     Set sp = sp_
End Sub

Private Function fillFields()
11390 On Error GoTo ERR_fillFields
        
        Dim sql As String
        Dim rs As ADODB.Recordset
        Dim empty_fields As Boolean
        Dim i As Integer
        Dim main_name As String
11400   empty_fields = True

11410     edtBarCode.Text = Trim(UCase(edtBarCode.Text))

11420     sdg_id = -1
11430     sdg_name = ""
11440     sql = "Select sdg.sdg_id, sdg.name, sdg.external_reference, " _
             & "c.name as id_card,  cu.u_first_name, cu.u_last_name , su.U_PATHOLAB_NUMBER, sdg.status  " _
             & "from lims_sys.sdg, lims_sys.sdg_user su, lims_sys.client c, lims_sys.client_user cu " _
             & "where sdg.sdg_id = su.sdg_id " _
             & "and su.u_patient = c.client_id " _
             & "and c.client_id = cu.client_id " _
             & "and ( sdg.name = '" & UCase(CheckApostrophe(edtBarCode.Text)) & "' " _
             & " or su.U_PATHOLAB_NUMBER ='" & UCase(CheckApostrophe(edtBarCode.Text)) & "') " _
             & " and sdg.name not like  '%V%' "
11450     Set rs = aConnection.Execute(sql)
      '       & "and (sdg.name = '" & CheckApostrophe(edtBarCode.Text) & "' " _
             & "or sdg.name like '" & CheckApostrophe(edtBarCode.Text) & "V%' ) "
11460   If Not rs.EOF Then
11470       strSdgId = nte(rs("SDG_ID"))
11480       sdg_id = rs("SDG_ID")
11490       sdg_name = rs("NAME")
11500       i = InStr(1, sdg_name, "V")
11510       If i = 0 Then
11520         main_name = sdg_name
11530       Else
11540         main_name = Mid(sdg_name, 1, i - 1)
11550       End If
'roy - don't allow barcode like B00001/16V1 but allow patholab number
11560      ' If main_name = edtBarCode.Text Then
11570         edtOriginalName.Text = sdg_name
11580         edtExternalReference.Text = CheckFieldValue(rs("EXTERNAL_REFERENCE"))
11590         edtIDCard.Text = CheckFieldValue(rs("ID_CARD"))
11600         edtFirstName.Text = CheckFieldValue(rs("U_FIRST_NAME"))
11610         edtLastName.Text = CheckFieldValue(rs("U_LAST_NAME"))
11620         empty_fields = False
'11630       Else
'11640         empty_fields = False
'11650         sdg_id = -1
'11660         edtBarCode.BackColor = vbRed
'11670         MsgBox "SDG already revised !"
'11680         edtBarCode.BackColor = vbWhite
'11690         Call edtBarCode.SetFocus
'11700       End If
11710   End If
11720   If empty_fields Then
11730     edtOriginalName.Text = ""
11740     edtExternalReference.Text = ""
11750     edtIDCard.Text = ""
11760     edtFirstName.Text = ""
11770     edtLastName.Text = ""
11780     edtBarCode.BackColor = vbRed
11790     MsgBox "Invalid barCode !"
11800     edtBarCode.BackColor = vbWhite
11810     Call edtBarCode.SetFocus
11820   End If
11830   Call rs.Close
        
11840   Exit Function
ERR_fillFields:
11850 MsgBox "Error on line:" & Erl & " in fillFields" & vbCrLf & Err.Description
End Function

Private Sub CancelButton_Click()
11860   Unload Me
End Sub

Private Sub edtBarCode_KeyPress(KeyAscii As Integer)
11870   If KeyAscii = 13 Then
11880     Call fillFields
11890   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
11900 On Error GoTo ERR_Form_KeyDown
          
          Dim strVer As String

11910     If KeyCode = vbKeyF10 And Shift = 1 Then
11920         strVer = "Name: " & App.EXEName & vbCrLf & vbCrLf & _
                       "Path: " & App.Path & vbCrLf & vbCrLf & _
                       "Version: " & "[" & App.Major & "." & App.Minor & "." & App.Revision & "]" & vbCrLf & vbCrLf & _
                       "Company: One Software Technologies (O.S.T) Ltd."
11930         MsgBox strVer, vbInformation, "Nautilus - Project Properties"
11940         Call edtBarCode.SetFocus
11950     End If
          
11960     Exit Sub
ERR_Form_KeyDown:
11970 MsgBox "Error on line:" & Erl & " in Form_KeyDown" & vbCrLf & Err.Description
End Sub

Private Sub Form_Load()
11980 On Error GoTo ERR_Form_Load

        Dim phrase As ADODB.Recordset
        Dim i As Integer
11990   Set phrase = aConnection.Execute("select phrase_description, phrase_name from lims_sys.phrase_entry " & _
            "where phrase_id = (select phrase_id from lims_sys.phrase_header where " & _
            "name = 'Revision cause') " & _
            "order by order_number")

12000   Call all_revision_codes.RemoveAll
12010   Do Until phrase.EOF
12020     Call all_revision_codes.Add(CheckFieldValue(phrase("phrase_description")), CheckFieldValue(phrase("phrase_name")))
12030     cmbRevisionCause.List(cmbRevisionCause.ListCount) = phrase("phrase_description")
12040     phrase.MoveNext
12050   Loop
12060   Call phrase.Close

12070   Call MacLab.zLang.English
12080   edtBarCode.Alignment = vbLeftJustify
12090   edtBarCode.RightToLeft = False
      '  Call edtBarCode.SetFocus

12100     Exit Sub
ERR_Form_Load:
12110 MsgBox "Error on line:" & Erl & " in Form_Load" & vbCrLf & Err.Description
End Sub

Private Sub OKButton_Click()
12120 On Error GoTo ERR_OKButton_Click
          
          Dim cg As Revision.CopyGenerator

12130     fillFields
12140     If sdg_id = -1 Then
12150         edtBarCode.BackColor = vbRed
12160         MsgBox "You should enter a valid SDG code first !"
12170         edtBarCode.BackColor = vbWhite
12180         Call edtBarCode.SetFocus
12190     ElseIf cmbRevisionCause.Text = "" Then
12200         edtBarCode.BackColor = vbRed
12210         MsgBox "You should enter a revision cause !"
12220         edtBarCode.BackColor = vbWhite
12230         Call edtBarCode.SetFocus
12240     Else
12250         revision_cause = all_revision_codes.Item(cmbRevisionCause.Text)

12260         Set cg = New Revision.CopyGenerator
12270         Call cg.Initialize(sp, strSdgId, revision_cause)
12280         Call cg.Execute
              
      '        Call CreateSdgCopy
      '        Call runSDGCopy
        MsgBox "Revision Created secessfully !" & vbCrLf & " This window will be closed now"
        
12290         Unload Me
12300     End If
          
12310     Exit Sub
ERR_OKButton_Click:
12320 MsgBox "Error on line:" & Erl & " in OKButton_Click" & vbCrLf & Err.Description
End Sub

