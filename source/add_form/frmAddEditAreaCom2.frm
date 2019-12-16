VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditAreaCom2 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   8880
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   8325
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   10905
      _ExtentX        =   19235
      _ExtentY        =   14684
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjLedgerReport.uctlTextLookup uctlCustomerLookup 
         Height          =   375
         Left            =   2115
         TabIndex        =   0
         Top             =   1440
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   661
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   8955
         _ExtentX        =   15796
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin VB.Label lblCustomer 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   1440
         Width           =   1575
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4560
         TabIndex        =   2
         Top             =   2400
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   2880
         TabIndex        =   1
         Top             =   2400
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditAreaCom2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private itemCusArea  As CCommissionCustomerArea

Public ID As Long
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String

Public COMMISSION_CUS_AREA_ID As Long
Public MASTER_AREA_ID As Long
Public YEAR_ID As Long

Public m_CusInYear As Collection
Public m_Customer As Collection
Private itemCusInYear As CCommissionCustomerArea
'Private Sub cboDataType_Click()
'   m_HasModify = True
'End Sub
Private Sub cboDataType_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      CreateObject("WScript.Shell").SendKeys "{TAB}"
   End If
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
      
      itemCusArea.COMMISSION_CUS_AREA_ID = COMMISSION_CUS_AREA_ID
      If Not glbDaily.QueryCusAreaCom2(itemCusArea, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If ItemCount > 0 Then

      Call itemCusArea.PopulateFromRS(2, m_Rs)
      uctlCustomerLookup.MyTextBox.Text = itemCusArea.COMMISSION_CUS_ID

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
Dim m_Lookup As CCommissionCustomerArea
  Set m_Lookup = New CCommissionCustomerArea
  
   Set itemCusInYear = GetObject("", m_CusInYear, Trim(uctlCustomerLookup.MyTextBox.Text))
   If Not (itemCusInYear Is Nothing) Then
      glbErrorLog.LocalErrorMsg = MapText("���١��� : " & itemCusInYear.COMMISSION_CUS_ID & " " & itemCusInYear.COMMISSION_CUS_NAME & "   ��������� ࢵ : " & itemCusInYear.MASTER_AREA_NAME)
      glbErrorLog.ShowUserError
      Exit Function
   End If

   If Not (uctlCustomerLookup.MyCombo.ItemData(Minus2Zero(uctlCustomerLookup.MyCombo.ListIndex)) > 0) Then
      Exit Function
   End If

   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   itemCusArea.AddEditMode = ShowMode
   itemCusArea.MASTER_AREA_ID = MASTER_AREA_ID
   itemCusArea.YEAR_ID = YEAR_ID
   itemCusArea.COMMISSION_CUS_ID = uctlCustomerLookup.MyTextBox.Text
    itemCusArea.COMMISSION_CUS_NAME = uctlCustomerLookup.MyCombo.Text

   Call EnableForm(Me, False)
   If Not glbDaily.AddEditCusArea(itemCusArea, IsOK, True, glbErrorLog) Then
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
      
      Call LoadCustomerLookup(uctlCustomerLookup.MyCombo, m_Customer)
      Set uctlCustomerLookup.MyCollection = m_Customer
      
      Call LoadCusFromAreaNameCom(Nothing, YEAR_ID, m_CusInYear)

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
   
   Call InitNormalLabel(lblCustomer, MapText("�����١���"))

   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("¡��ԡ (ESC)"))
   Call InitMainButton(cmdOK, MapText("��ŧ (F2)"))
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
   Set m_Rs = New ADODB.Recordset
      
   Set itemCusArea = New CCommissionCustomerArea
   Set m_CusInYear = New Collection
    Set itemCusInYear = New CCommissionCustomerArea
   Set m_Customer = New Collection
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set itemCusArea = Nothing
   Set m_Customer = Nothing
   Set m_CusInYear = Nothing
   Set itemCusInYear = Nothing
End Sub

Private Sub uctlCustomerLookup_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlCustomerLookup_Change()
   m_HasModify = True
End Sub

