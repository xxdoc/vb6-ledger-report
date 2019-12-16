VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditSupplierGroup 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6840
   Icon            =   "frmAddEditSupplierGroup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   6840
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   4005
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6825
      _ExtentX        =   12039
      _ExtentY        =   7064
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboSubGroupTypeCode 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2640
         Width           =   2655
      End
      Begin VB.ComboBox cboGroupTypeCode 
         Height          =   315
         Left            =   2340
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2040
         Width           =   2715
      End
      Begin VB.ComboBox cboDataType 
         Height          =   315
         Left            =   2340
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   960
         Width           =   2715
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   6795
         _ExtentX        =   11986
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjLedgerReport.uctlTextBox txtSupplierCode 
         Height          =   435
         Left            =   2340
         TabIndex        =   1
         Top             =   1470
         Width           =   1875
         _ExtentX        =   5212
         _ExtentY        =   767
      End
      Begin VB.Label lblSubGroupTypeCode 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   360
         TabIndex        =   10
         Top             =   2640
         Width           =   1935
      End
      Begin VB.Label lblDataType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   255
         Left            =   330
         TabIndex        =   8
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label lblGroupTypeCode 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   360
         TabIndex        =   7
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label lblSupplierCode 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   330
         TabIndex        =   6
         Top             =   1560
         Width           =   1935
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   3435
         TabIndex        =   3
         Top             =   3120
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   1785
         TabIndex        =   2
         Top             =   3120
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditSupplierGroup.frx":27A2
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditSupplierGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_SupplierGroup As CSupplierGroup

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String
Private Sub cboDataType_Click()
   m_HasModify = True
End Sub

Private Sub cboDataType_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      CreateObject("WScript.Shell").SendKeys "{TAB}"
   End If
End Sub
Private Sub cboGroupTypeCode_Click()
   m_HasModify = True
End Sub
Private Sub cboSubGroupTypeCode_Click()
   m_HasModify = True
End Sub
Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      m_SupplierGroup.SUPPLIER_GROUP_ID = ID
      m_SupplierGroup.QueryFlag = 1
      If Not glbDaily.QuerySupplierGroup(m_SupplierGroup, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then
      Call m_SupplierGroup.PopulateFromRS(1, m_Rs)
      
      cboDataType.ListIndex = IDToListIndex(cboDataType, m_SupplierGroup.DATA_TYPE_ID)
      txtSupplierCode.Text = m_SupplierGroup.SUPPLIER_CODE
      cboGroupTypeCode.ListIndex = IDToListIndex(cboGroupTypeCode, m_SupplierGroup.GROUP_TYPE_CODE)
      cboSubGroupTypeCode.ListIndex = IDToListIndex(cboSubGroupTypeCode, m_SupplierGroup.SUB_GROUP_TYPE_CODE)
   End If
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call EnableForm(Me, True)
End Sub
Private Function SaveData() As Boolean
Dim IsOK As Boolean
   
   If Not VerifyCombo(lblDataType, cboDataType, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblSupplierCode, txtSupplierCode, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblGroupTypeCode, cboGroupTypeCode, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblSubGroupTypeCode, cboSubGroupTypeCode, False) Then
      Exit Function
   End If
   
'   If Not CheckUniqueNs(REAL_CREDIT_NO, txtDocumentNo.Text, ID) Then
'      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtDocumentNo.Text & " " & MapText("อยู่ในระบบแล้ว")
'      glbErrorLog.ShowUserError
'      Exit Function
'   End If
   
   'Call LoadARAmountByCust(Nothing, DocNoColl, -1, -1, txtDocumentNo.Text)
'   If DocNoColl.Count <= 0 Then
'      glbErrorLog.LocalErrorMsg = MapText("ไม่มีข้อมูลหมายเลขเอกสาร") & " " & txtDocumentNo.Text & " " & MapText("ในระบบ")
'      glbErrorLog.ShowUserError
'      Exit Function
'   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_SupplierGroup.AddEditMode = ShowMode
   m_SupplierGroup.SUPPLIER_GROUP_ID = ID
   m_SupplierGroup.SUPPLIER_CODE = txtSupplierCode.Text
   m_SupplierGroup.GROUP_TYPE_CODE = cboGroupTypeCode.ItemData(Minus2Zero(cboGroupTypeCode.ListIndex))
   m_SupplierGroup.SUB_GROUP_TYPE_CODE = cboSubGroupTypeCode.ItemData(Minus2Zero(cboSubGroupTypeCode.ListIndex))
   m_SupplierGroup.DATA_TYPE_ID = cboDataType.ItemData(Minus2Zero(cboDataType.ListIndex))
   
   Call EnableForm(Me, False)
   If Not glbDaily.AddEditSupplierGroup(m_SupplierGroup, IsOK, True, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      SaveData = False
      Call EnableForm(Me, True)
      Exit Function
   End If
   If Not IsOK Then
      Call EnableForm(Me, True)
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   Call EnableForm(Me, True)
   SaveData = True
End Function
Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadDataType(cboDataType)
      Call LoadGroupTypeData(cboGroupTypeCode)
      Call LoadSubGroupTypeData(cboSubGroupTypeCode)
      
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         ID = -1
      End If
      
      m_HasModify = False
   End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.Name
      glbErrorLog.ShowUserError
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 116 Then
'      Call cmdSearch_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 115 Then
'      Call cmdClear_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 118 Then
'      Call cmdAdd_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 117 Then
'      Call cmdDelete_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 114 Then
'      Call cmdEdit_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 121 Then
'      Call cmdPrint_Click
      KeyCode = 0
   End If
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   
   Call InitNormalLabel(lblSupplierCode, MapText("รหัสเจ้าหนี้"))
   Call txtSupplierCode.SetKeySearch("SUPPLIER_CODE")
   
   Call InitNormalLabel(lblGroupTypeCode, MapText("ประเภทกลุ่ม"))
   Call InitNormalLabel(lblSubGroupTypeCode, MapText("ประเภทกลุ่มย่อย"))
   Call InitNormalLabel(lblDataType, MapText("ประเภทข้อมูล"))
   
   Call txtSupplierCode.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   
   Call InitCombo(cboDataType)
   Call InitCombo(cboGroupTypeCode)
   Call InitCombo(cboSubGroupTypeCode)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
End Sub

Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   Call EnableForm(Me, False)
   m_HasActivate = False
   
   Set m_SupplierGroup = New CSupplierGroup
   Set m_Rs = New ADODB.Recordset
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set m_SupplierGroup = Nothing
End Sub
Private Sub txtGroupTypeCode_Change()
   m_HasModify = True
End Sub
Private Sub txtSupplierCode_Change()
   m_HasModify = True
End Sub
