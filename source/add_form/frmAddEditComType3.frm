VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditComType3 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6825
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   6825
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   3045
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6825
      _ExtentX        =   12039
      _ExtentY        =   5371
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjLedgerReport.uctlTextBox uctlNumOne 
         Height          =   375
         Left            =   3240
         TabIndex        =   0
         Tag             =   "2"
         Top             =   1080
         Width           =   2055
         _ExtentX        =   2566
         _ExtentY        =   661
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   5
         Top             =   3000
         Width           =   6795
         _ExtentX        =   11986
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjLedgerReport.uctlTextBox uctlSlmPercent 
         Height          =   435
         Left            =   3240
         TabIndex        =   1
         Tag             =   "2"
         Top             =   1560
         Width           =   2115
         _ExtentX        =   5212
         _ExtentY        =   767
      End
      Begin Threed.SSPanel ppnlHeader 
         Height          =   705
         Left            =   -120
         TabIndex        =   8
         Top             =   0
         Width           =   6915
         _ExtentX        =   12197
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin VB.Label lblNumOne 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   255
         Left            =   150
         TabIndex        =   7
         Top             =   1080
         Width           =   2775
      End
      Begin VB.Label lblSlmPercent 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   270
         TabIndex        =   6
         Top             =   1560
         Width           =   2655
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   3435
         TabIndex        =   3
         Top             =   2430
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
         Top             =   2430
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditComType3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_CConditionCommission As CConditionCommission
Private m_CPromoteCommission As CCommissPromote

Public COM_ID As Long
Public YEAR_ID As Long
Public COMTYP As String

Public MASTER_COMMISS_SUB_PROMOTE_ID As Long
Public itemCountGrid As Long
Public MASTER_Commiss_ID As Long

Public DB_TYP As Long

Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public HeaderText As String

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
      
       If DB_TYP = 1 Then       ' เลือกตารางโปรโมต
                m_CPromoteCommission.MASTER_COMMISS_SUB_PROMOTE_ID = MASTER_COMMISS_SUB_PROMOTE_ID
                m_CPromoteCommission.Commiss_PROMOTE_ID = COM_ID
                m_CPromoteCommission.FROM_CMPL_DATE = -1
                m_CPromoteCommission.TO_CMPL_DATE = -1
                m_CPromoteCommission.QueryFlag = 1
         
                  If Not glbDaily.QueryPromoteType1(m_CPromoteCommission, m_Rs, ItemCount, IsOK, glbErrorLog) Then
                     glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
                     Call EnableForm(Me, True)
                     Exit Sub
                  End If
                  
                  If ItemCount > 0 Then
                 Call m_CPromoteCommission.PopulateFromRS(2, m_Rs)
                        
                        uctlNumOne.Text = m_CPromoteCommission.NUM_ONE
                      '   uctlNumTwo.Text = m_CConditionCommission.NUM_TWO
                        uctlSlmPercent.Text = m_CPromoteCommission.SLM_PERCENT
                      ' cboOperator.ListIndex = IDToListIndex(cboOperator, m_CConditionCommission.OPERATOR)
                 End If
                  
         Else
      
               m_CConditionCommission.COM_ID = COM_ID
                  m_CConditionCommission.YEAR_ID = YEAR_ID
                  m_CConditionCommission.FROM_CMPL_DATE = -1
                        m_CConditionCommission.TO_CMPL_DATE = -1
                  m_CConditionCommission.QueryFlag = 1

                        If Not glbDaily.QueryComType1(m_CConditionCommission, m_Rs, ItemCount, IsOK, glbErrorLog) Then
                           glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
                           Call EnableForm(Me, True)
                           Exit Sub
                        End If
               
               If ItemCount > 0 Then
                  Call m_CConditionCommission.PopulateFromRS(2, m_Rs)
                  
                  uctlNumOne.Text = m_CConditionCommission.NUM_ONE
             '     uctlNumTwo.Text = m_CConditionCommission.NUM_TWO
                  uctlSlmPercent.Text = m_CConditionCommission.SLM_PERCENT
             '    cboOperator.ListIndex = IDToListIndex(cboOperator, m_CConditionCommission.OPERATOR)
               End If
         End If
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
   
   If Not VerifyTextControl(lblNumOne, uctlNumOne, False) Then
      Exit Function
   End If
    

   If Not VerifyTextControl(lblSlmPercent, uctlSlmPercent, False) Then
      Exit Function
   End If
   

   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
' เคสเดิม
If DB_TYP = 1 Then       ' เลือกตารางโปรโมต
   m_CPromoteCommission.ShowMode = ShowMode
   m_CPromoteCommission.Commiss_TYP = COMTYP      ' ไม่ 2 ก็ 3 ขึ้นอยู่กับตอนแรก
   m_CPromoteCommission.MASTER_COMMISS_SUB_PROMOTE_ID = MASTER_COMMISS_SUB_PROMOTE_ID
   m_CPromoteCommission.GROUP1 = -1
   m_CPromoteCommission.NUM_ONE = uctlNumOne.Text
   m_CPromoteCommission.SLM_PERCENT = uctlSlmPercent.Text
'   m_CPromoteCommission.OPERATOR = cboOperator.ItemData(Minus2Zero(cboOperator.ListIndex))
   
   Call EnableForm(Me, False)
    
    If Not glbDaily.AddEditPromoteCommiss(m_CPromoteCommission, IsOK, True, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      SaveData = False
      Call EnableForm(Me, True)
      Exit Function
   End If
Else

   m_CConditionCommission.AddEditMode = ShowMode
   m_CConditionCommission.COMTYP = COMTYP
   m_CConditionCommission.YEAR_ID = YEAR_ID
   m_CConditionCommission.NUM_ONE = uctlNumOne.Text
   m_CConditionCommission.SLM_PERCENT = uctlSlmPercent.Text
   m_CConditionCommission.GROUP1 = -1
 '  m_CConditionCommission.OPERATOR = cboOperator.ItemData(Minus2Zero(cboOperator.ListIndex))
   
   Call EnableForm(Me, False)
    
      If Not glbDaily.AddEditConditionCommiss(m_CConditionCommission, IsOK, True, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         SaveData = False
         Call EnableForm(Me, True)
         Exit Function
      End If
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
      
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         COM_ID = -1
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
   ppnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   ppnlHeader.Caption = Me.Caption
   
   Call InitNormalLabel(lblNumOne, MapText("ระยะเวลา(วัน)"))
   Call uctlNumOne.SetTextLenType(TEXT_FLOAT, glbSetting.DOUBLE_TYPE)

 Call InitNormalLabel(lblSlmPercent, MapText("ค่าคอมมิชชั่น เก็บเงิน(%)"))
  Call uctlSlmPercent.SetTextLenType(TEXT_FLOAT, glbSetting.DOUBLE_TYPE)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   ppnlHeader.Font.Name = GLB_FONT
   ppnlHeader.Font.Bold = True
   ppnlHeader.Font.Size = 19
   
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
   
 Set m_CPromoteCommission = New CCommissPromote
 Set m_CConditionCommission = New CConditionCommission
   Set m_Rs = New ADODB.Recordset
   
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set m_CConditionCommission = Nothing
End Sub

Private Sub uctlNumOne_Change()
   m_HasModify = True
End Sub
Private Sub uctlSlmPercent_Change()
   m_HasModify = True
End Sub
