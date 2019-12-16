VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmDatabaseSelect 
   ClientHeight    =   5685
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8490
   Icon            =   "frmDatabaseSelect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   8490
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   5700
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8505
      _ExtentX        =   15002
      _ExtentY        =   10054
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   540
         Top             =   4320
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   393216
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   4035
         Left            =   180
         TabIndex        =   0
         Top             =   870
         Width           =   8145
         _ExtentX        =   14367
         _ExtentY        =   7117
         _Version        =   393217
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "ImageList1"
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   8475
         _ExtentX        =   14949
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   2625
         TabIndex        =   1
         Top             =   5010
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4275
         TabIndex        =   2
         Top             =   5010
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblMotherNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   180
         TabIndex        =   4
         Top             =   1620
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmDatabaseSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_Houses As Collection
Private m_Employees As Collection

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public Database2 As Boolean
Public Database3 As Boolean
Public OKClick As Boolean
Public ID As Long
Public DBPath As String
Public CompanyName As String
Public CompanyFullName  As String

Private FileName As String
Private m_SumUnit As Double
Private m_OldPartItemID As Long
Private m_Companys As Collection
Private m_Companys3 As Collection

Private Sub LoadTreeView(Col As Collection)
Dim N As Node
Dim Np As Node
Dim C As CSCComp
Dim i As Long

   i = 0
''debug.print "====="
   For Each C In Col
      i = i + 1
      Set N = TreeView1.Nodes.Add(, tvwFirst, C.PATH & "-" & i, C.COMPNAM & " (" & C.PATH & ")")
      N.Tag = C.PATH
''debug.print C.PATH
      N.Expanded = False
   Next C
''debug.print "====="
End Sub

Private Sub LoadLocationTreeView(Col As Collection)
'Dim C As CLocation
'Dim N As Node
'Dim Np As Node
'
'      For Each C In Col
'         Set N = TreeView1.Nodes.Add(, tvwFirst, Trim(Str(C.LOCATION_ID)) & "-X", C.LOCATION_NAME & " (" & C.LOCATION_NO & ")", 1, 1)
'         N.Tag = C.LOCATION_ID
'         N.Checked = False
'
'         N.Expanded = False
'      Next C
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

'   IsOK = True
'   If Flag Then
'      Call EnableForm(Me, False)
'
'      m_HouseGroup.HOUSE_GROUP_ID = ID
'      If Not glbMaster.QueryHouseGroup(m_HouseGroup, m_Rs, ItemCount, IsOK, glbErrorLog) Then
'         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
'         Call EnableForm(Me, True)
'         Exit Sub
'      End If
'   End If
'
'   If ItemCount > 0 Then
'      Call m_HouseGroup.PopulateFromRS(1, m_Rs)
'      txtDocumentNo.Text = m_HouseGroup.HOUSE_GROUP_NO
'      txtMotherNo.Text = m_HouseGroup.HOUSE_GROUP_NAME
'      chkExtraFlag.Value = FlagToCheck(m_HouseGroup.EXTRA_FLAG)
'
'      Dim II As CHGroupItem
'      If m_HouseGroup.HGroupItems.Count > 0 Then
'         Call LoadTreeView(m_HouseGroup.HGroupItems)
'      End If
'   Else
'      Call LoadTreeView(m_HouseGroup.HGroupItems)
'   End If
'   If Not IsOK Then
'      glbErrorLog.ShowUserError
'      Call EnableForm(Me, True)
'      Exit Sub
'   End If
'
'   Call EnableForm(Me, True)
End Sub

Private Sub cboGroup_Click()
   m_HasModify = True
End Sub

Private Sub chkEnable_Click()
   m_HasModify = True
End Sub

Private Sub PopulateItem(Col As Collection)
'Dim N As Node
'Dim C As CHGroupItem
'
'   For Each C In Col
'      C.Flag = "D"
'   Next C
'
'   For Each N In TreeView1.Nodes
'      Set C = New CHGroupItem
'      C.Flag = "A"
'      C.LOCATION_ID = Val(N.Key)
'      If N.Checked Then
'         C.SELECT_FLAG = "Y"
'      Else
'         C.SELECT_FLAG = "N"
'      End If
'      Call Col.Add(C)
'      Set C = Nothing
'   Next N
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim L As Long
On Error Resume Next
    
   If ShowMode = SHOW_ADD Then
'      If Not VerifyAccessRight("DAILY_CUSTOMER_ADD") Then
'         Call EnableForm(Me, True)
'         Exit Function
'      End If
   ElseIf ShowMode = SHOW_EDIT Then
'      If Not VerifyAccessRight("DAILY_CUSTOMER_EDIT") Then
'         Call EnableForm(Me, True)
'         Exit Function
'      End If
   End If
   
'   If Not CheckUniqueNs(IMPORT_UNIQUE, txtDocumentNo.Text, ID) Then
'      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtDocumentNo.Text & " " & MapText("อยู่ในระบบแล้ว")
'      glbErrorLog.ShowUserError
'      Exit Function
'   End If
      
'   If Not m_HasModify Then
'      SaveData = True
'      Exit Function
'   End If
   DBPath = TreeView1.SelectedItem.Tag
   If Len(DBPath) > 8 Then
      DBPath = Left(DBPath, 8)
   End If
   
   If Database2 = True Then
     glbCompanyCode = DBPath
      L = InStr(1, glbParameterObj.DBFile, "Secure", vbTextCompare)
      DBPath = Mid(glbParameterObj.DBFile, 1, L - 1) & DBPath
      CompanyName = glbCompanyName & " , " & Mid(TreeView1.SelectedItem.Text, InStr(1, TreeView1.SelectedItem.Text, ".") + 1)
      glbCompanyName = CompanyName
'   Else
'      glbCompanyCode = DBPath
'      L = InStr(1, glbParameterObj.DBFile, "Secure", vbTextCompare)
'      DBPath = Mid(glbParameterObj.DBFile, 1, L - 1) & DBPath
'      CompanyName = Mid(TreeView1.SelectedItem.Text, InStr(1, TreeView1.SelectedItem.Text, ".") + 1)
'      glbCompanyName = CompanyName
   ElseIf Database3 = True Then
     glbCompanyCode = DBPath
      L = InStr(1, glbParameterObj.DBFile, "Secure", vbTextCompare)
      DBPath = Mid(glbParameterObj.DBFile, 1, L - 1) & DBPath
      CompanyName = glbCompanyName & " และ " & Mid(TreeView1.SelectedItem.Text, InStr(1, TreeView1.SelectedItem.Text, ".") + 1)
      glbCompanyName = CompanyName
   Else
      glbCompanyCode = DBPath
      L = InStr(1, glbParameterObj.DBFile, "Secure", vbTextCompare)
      DBPath = Mid(glbParameterObj.DBFile, 1, L - 1) & DBPath
      CompanyName = Mid(TreeView1.SelectedItem.Text, InStr(1, TreeView1.SelectedItem.Text, ".") + 1)
      glbCompanyName = CompanyName
     'CompanyFullName = Mid(TreeView1.SelectedItem.Text, InStr(1, TreeView1.SelectedItem.Text, ".") + 1)
   End If

  If Dir(DBPath, vbDirectory) = vbNullString Then
      Call EnableForm(Me, True)
      SaveData = False
   Else
     Call EnableForm(Me, True)
     SaveData = True
   End If
End Function

Private Sub cboBusinessGroup_Click()
   m_HasModify = True
End Sub

Private Sub cboBusinessType_Click()
   m_HasModify = True
End Sub

Private Sub cboEnterpriseType_Click()
   m_HasModify = True
End Sub
Private Sub chkCommit_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkExtraFlag_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub cboDBPath_Click()
'Dim DBFile As String
'
'   DBFile = cboDBPath.Text
'
'   Call EnableForm(Me, False)
'
'   Call glbDatabaseMngr.DisConnectDatabase
'   If Not glbDatabaseMngr.ConnectDatabase(DBFile, "", "", glbErrorLog) Then
'      Exit Sub
'   End If
'
'   Call LoadCompany(Nothing, m_Companys)
'   Call LoadTreeView(m_Companys)
'
'   Call EnableForm(Me, True)
End Sub

Private Sub cmdOK_Click()
   If Not SaveData Then
      OKClick = False
      MsgBox "ไม่มีฐานข้อมูลที่คุณเลือก", vbOKOnly, "ไม่พบข้อมูล"
   Else
      OKClick = True
      Unload Me
   End If
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)

      Call LoadCompany(Nothing, m_Companys, Database2, Database3)
      Call LoadTreeView(m_Companys)

      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         Call QueryData(False)
         Call LoadLocationTreeView(m_Companys)
      End If
      
      Call EnableForm(Me, True)
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

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   Set m_Houses = Nothing
   Set m_Employees = Nothing
   Set m_Companys = Nothing
   Set m_Companys3 = Nothing
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   pnlHeader.Caption = "กรุณาเลือกบริษัทที่ต้องการออกรายงาน"
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
'   Call InitCombo(cboDBPath)

   Call InitTreeView
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
End Sub

Private Sub InitTreeView()
   TreeView1.Font.Name = GLB_FONT
   TreeView1.Font.Size = 14
End Sub

Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   OKClick = False
   Call InitFormLayout

   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
   Set m_Houses = New Collection
   Set m_Employees = New Collection
   Set m_Companys = New Collection
   Set m_Companys3 = New Collection
End Sub

Private Sub txtDoNo_Change()
   m_HasModify = True
End Sub

Private Sub txtParentNo_Change()
   m_HasModify = True
End Sub

Private Sub txtSellBy_Change()
   m_HasModify = True
End Sub

Private Sub TreeView1_DblClick()
   Call cmdOK_Click
End Sub

Private Sub TreeView1_NodeCheck(ByVal Node As MSComctlLib.Node)
   m_HasModify = True
End Sub

Private Sub txtDocumentNo_Change()
   m_HasModify = True
End Sub

Private Sub txtPassword_Change()
   m_HasModify = True
End Sub

Private Sub txtSender_Change()
   m_HasModify = True
End Sub

Private Sub txtTotal_Change()
   m_HasModify = True
End Sub

Private Sub txtMotherNo_Change()
   m_HasModify = True
End Sub

Private Sub uctlSetupDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlTextBox1_Change()
   m_HasModify = True
End Sub

Private Sub uctlDeliveryLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlDocumentDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlHouseLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlResponseByLookup_Change()
   m_HasModify = True
End Sub
