VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmAddEditPromoteIncentive 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmAddEditPromoteIncentive.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   8520
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   15028
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjLedgerReport.uctlDate uctlFromDate 
         Height          =   375
         Left            =   3840
         TabIndex        =   2
         Top             =   1800
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
      End
      Begin prjLedgerReport.uctlTextBox txtMasIncentiveNo 
         Height          =   450
         Left            =   3840
         TabIndex        =   0
         Top             =   840
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   873
      End
      Begin prjLedgerReport.uctlDate uctlToDate 
         Height          =   375
         Left            =   3840
         TabIndex        =   3
         Top             =   2280
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   120
         TabIndex        =   4
         Top             =   2880
         Width           =   11640
         _ExtentX        =   20532
         _ExtentY        =   979
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "AngsanaUPC"
            Size            =   14.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   4245
         Left            =   120
         TabIndex        =   5
         Top             =   3360
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   7488
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         TabKeyBehavior  =   1
         MethodHoldFields=   -1  'True
         AllowColumnDrag =   0   'False
         AllowEdit       =   0   'False
         BorderStyle     =   3
         GroupByBoxVisible=   0   'False
         DataMode        =   99
         HeaderFontName  =   "AngsanaUPC"
         HeaderFontBold  =   -1  'True
         HeaderFontSize  =   14.25
         HeaderFontWeight=   700
         FontSize        =   9.75
         BackColorBkg    =   16777215
         ColumnHeaderHeight=   480
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmAddEditPromoteIncentive.frx":27A2
         Column(2)       =   "frmAddEditPromoteIncentive.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditPromoteIncentive.frx":290E
         FormatStyle(2)  =   "frmAddEditPromoteIncentive.frx":2A6A
         FormatStyle(3)  =   "frmAddEditPromoteIncentive.frx":2B1A
         FormatStyle(4)  =   "frmAddEditPromoteIncentive.frx":2BCE
         FormatStyle(5)  =   "frmAddEditPromoteIncentive.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditPromoteIncentive.frx":2D5E
      End
      Begin prjLedgerReport.uctlTextBox txtMasterIncentiveDesc 
         Height          =   450
         Left            =   3840
         TabIndex        =   1
         Top             =   1320
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   794
      End
      Begin VB.Label lblMasIncentiveDesc 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2040
         TabIndex        =   16
         Top             =   1320
         Width           =   1365
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10080
         TabIndex        =   10
         Top             =   7800
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblFromDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   1680
         TabIndex        =   15
         Top             =   1800
         Width           =   1755
      End
      Begin VB.Label lblToDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   2160
         TabIndex        =   14
         Top             =   2280
         Width           =   1365
      End
      Begin VB.Label lblMasIncentiveNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   1680
         TabIndex        =   13
         Top             =   840
         Width           =   1755
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8400
         TabIndex        =   9
         Top             =   7800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPromoteIncentive.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1770
         TabIndex        =   7
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   150
         TabIndex        =   6
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPromoteIncentive.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   8
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditPromoteIncentive.frx":356A
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditPromoteIncentive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset

Private m_PromoteYear As CIncentiveMasterPromote
'Private m_Condition1 As CIncentivePromote
'Private m_Condition2 As CIncentivePromote
'Private m_Condition3 As CIncentivePromote
'Private m_Condition4 As CIncentivePromote
Private m_Condition5 As CIncentivePromote        ' incen
'Private m_Condition6 As CIncentivePromote       ' incen
'Private cm1_Rs As ADODB.Recordset
'Private cm2_Rs As ADODB.Recordset
'Private cm3_Rs As ADODB.Recordset
'Private cm4_Rs As ADODB.Recordset
Private cm5_Rs As ADODB.Recordset    '' incentive
'Private cm6_Rs As ADODB.Recordset     ' incentive


 
Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean

Public INCENTIVE_PROMOTE_ID As Long
Public INCENTIVE_TYP As String
Public STKCOD As String
Public STKDES As String
Public SLM_PERCENT As Long
Public MASTER_INCENTIVE_ID As Long
Public MASTER_INCENTIVE_NO As String
Public Flag As String

Public VALID_FROM As String
Public VALID_TO As String

Dim ItemCount As Long
'Dim itemCountGrid1 As Long
'Dim itemCountGrid2 As Long
'Dim itemCountGrid3 As Long
'Dim itemCountGrid4 As Long
Dim itemCountGrid5 As Long     ' incen
'Dim itemCountGrid6 As Long    ' incen

Private m_TableName As String
Private FileName As String
Private m_SumUnit As Double

Public Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean

                     IsOK = True
                     If Flag Then
                            Call EnableForm(Me, False)
                            
                           If MASTER_INCENTIVE_ID = 0 Then
                               If Not glbDaily.QueryMasterIncentive(m_PromoteYear, m_Rs, ItemCount, IsOK, glbErrorLog) Then
                                  glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
                                  Call EnableForm(Me, True)
                                  Exit Sub
                                End If
                                 Call m_PromoteYear.PopulateFromRS(1, m_Rs)
                                 MASTER_INCENTIVE_ID = m_PromoteYear.MASTER_INCENTIVE_ID
                            End If
                            
                            m_PromoteYear.MASTER_INCENTIVE_ID = MASTER_INCENTIVE_ID
                            m_PromoteYear.VALID_FROM = -1
                            m_PromoteYear.VALID_TO = -1
                            If Not glbDaily.QueryMasterIncentive(m_PromoteYear, m_Rs, ItemCount, IsOK, glbErrorLog) Then
                                  glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
                                  Call EnableForm(Me, True)
                                  Exit Sub
                            End If
                     End If

                     If ItemCount > 0 Then
                        Call m_PromoteYear.PopulateFromRS(1, m_Rs)
                        txtMasIncentiveNo.Text = m_PromoteYear.MASTER_INCENTIVE_NO    'ยังอยู่ในโหมด edit
                        uctlFromDate.ShowDate = m_PromoteYear.VALID_FROM
                        uctlToDate.ShowDate = m_PromoteYear.VALID_TO
                        txtMasterIncentiveDesc.Text = m_PromoteYear.MASTER_INCENTIVE_DESC
                     End If

                  If Not IsOK Then
                     glbErrorLog.ShowUserError
                     Call EnableForm(Me, True)
                     Exit Sub
                  End If
   
   Call TabStrip1_Click
   Call EnableForm(Me, True)
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean

   If Not VerifyTextControl(lblMasIncentiveNo, txtMasIncentiveNo, False) Then
      Exit Function
   End If
   If Not VerifyDate(lblFromDate, uctlFromDate, False) Then
      Exit Function
   End If
   If Not VerifyDate(lblToDate, uctlToDate, False) Then
      Exit Function
   End If
   
   If Not CheckUniqueNs(EXPORT_UNIQUE, txtMasIncentiveNo.Text, MASTER_INCENTIVE_ID) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtMasIncentiveNo.Text & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      Call txtMasIncentiveNo.SetFocus
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   

'If MASTER_INCENTIVE_ID = 0 Then
m_PromoteYear.ShowMode = ShowMode                        ' ตรงนี้ ตอนบันทึก
'Else:
' m_PromoteYear.ShowMode = SHOW_EDIT
' End If
    m_PromoteYear.MASTER_INCENTIVE_ID = MASTER_INCENTIVE_ID
    m_PromoteYear.MASTER_INCENTIVE_NO = txtMasIncentiveNo.Text
    m_PromoteYear.VALID_FROM = uctlFromDate.ShowDate
    m_PromoteYear.VALID_TO = uctlToDate.ShowDate
    m_PromoteYear.MASTER_INCENTIVE_DESC = txtMasterIncentiveDesc.Text
   
   Call EnableForm(Me, False)
   
   If Not glbDaily.AddEditMasterIncentive(m_PromoteYear, IsOK, True, glbErrorLog) Then
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

Private Sub cmdAdd_Click()    ' เมื่อกดเพิ่มแต่ละ case 3 case
Dim OKClick As Boolean
Dim IsOK As Boolean

If Not cmdAdd.Enabled Then
      Exit Sub
   End If
   
  If m_PromoteYear.MASTER_INCENTIVE_ID <= 0 Then
      glbErrorLog.LocalErrorMsg = MapText("กรุณาทำการบันทึกข้อมูลให้เรียบร้อยก่อน")
      glbErrorLog.ShowUserError
      Exit Sub
   End If
   
   Call EnableForm(Me, False)
   
   
   OKClick = False
   
If TabStrip1.SelectedItem.index = 1 Then
' frmAddEditINCENTIVE_TYPe2.MASTER_INCENTIVE_ID = MASTER_INCENTIVE_ID
' frmAddEditINCENTIVE_TYPe2.ShowMode = SHOW_ADD
'     frmAddEditINCENTIVE_TYPe2.HeaderText = MapText("เพิ่ม Commission ขาย")
'   Load frmAddEditINCENTIVE_TYPe2
'  frmAddEditINCENTIVE_TYPe2.Show 1
'
'   OKClick = frmAddEditINCENTIVE_TYPe2.OKClick
'
'
'   Unload frmAddEditINCENTIVE_TYPe2
'   Set frmAddEditINCENTIVE_TYPe2 = Nothing
'
'   If OKClick Then
'      frmAddEditConditionCommiss.ShowMode = SHOW_EDIT
'      Call QueryData(True)
'   End If
'
'
'   ElseIf TabStrip1.SelectedItem.Index = 2 Then
' frmAddEditINCENTIVE_TYPe3.MASTER_INCENTIVE_ID = MASTER_INCENTIVE_ID
' frmAddEditINCENTIVE_TYPe3.ShowMode = SHOW_ADD
' frmAddEditINCENTIVE_TYPe3.INCENTIVE_TYP = "02"
'   frmAddEditINCENTIVE_TYPe3.HeaderText = MapText("เพิ่ม Commission เก็บเงิน(1)")
'   Load frmAddEditINCENTIVE_TYPe3
'  frmAddEditINCENTIVE_TYPe3.Show 1
'
'   OKClick = frmAddEditINCENTIVE_TYPe3.OKClick
'
'   Unload frmAddEditINCENTIVE_TYPe3
'   Set frmAddEditINCENTIVE_TYPe3 = Nothing
'
'   If OKClick Then
'       frmAddEditConditionCommiss.ShowMode = SHOW_EDIT
'      Call QueryData(True)
'   End If
'
'   ElseIf TabStrip1.SelectedItem.Index = 3 Then
'  frmAddEditINCENTIVE_TYPe3.MASTER_INCENTIVE_ID = MASTER_INCENTIVE_ID
' frmAddEditINCENTIVE_TYPe3.ShowMode = SHOW_ADD
' frmAddEditINCENTIVE_TYPe3.INCENTIVE_TYP = "03"
' frmAddEditINCENTIVE_TYPe3.HeaderText = MapText("เพิ่ม Commission เก็บเงิน(2)")
'   Load frmAddEditINCENTIVE_TYPe3
'  frmAddEditINCENTIVE_TYPe3.Show 1
'
'   OKClick = frmAddEditINCENTIVE_TYPe3.OKClick
'
'   Unload frmAddEditINCENTIVE_TYPe3
'   Set frmAddEditINCENTIVE_TYPe3 = Nothing
'
'   If OKClick Then
'       frmAddEditConditionCommiss.ShowMode = SHOW_EDIT
'      Call QueryData(True)
'   End If
'
'   ElseIf TabStrip1.SelectedItem.Index = 4 Then
'    frmAddEditINCENTIVE_TYPe1.MASTER_INCENTIVE_ID = MASTER_INCENTIVE_ID
'      frmAddEditINCENTIVE_TYPe1.ShowMode = SHOW_ADD
'      frmAddEditINCENTIVE_TYPe1.HeaderText = MapText("เพิ่ม Commission เก็บเงิน(3)")
'   Load frmAddEditINCENTIVE_TYPe1
'  frmAddEditINCENTIVE_TYPe1.Show 1
'
'   OKClick = frmAddEditINCENTIVE_TYPe1.OKClick
'
'   Unload frmAddEditINCENTIVE_TYPe1
'   Set frmAddEditINCENTIVE_TYPe1 = Nothing
'
'   If OKClick Then
'       frmAddEditConditionCommiss.ShowMode = SHOW_EDIT
'      Call QueryData(True)
'   End If
'
'   ElseIf TabStrip1.SelectedItem.Index = 5 Then
      Set frmAddEditComType5.ParentForm = Me
           ' Set frmAddEditMasterFTItem.TempCollection = m_PromoteYear.Details
      Set frmAddEditComType5.TempCollection = m_PromoteYear.Details
       frmAddEditComType5.MASTER_INCENTIVE_ID = MASTER_INCENTIVE_ID
      frmAddEditComType5.ShowMode = SHOW_ADD
      frmAddEditComType5.HeaderText = MapText("เพิ่ม Incentive ")
      frmAddEditComType5.itemCountGrid = itemCountGrid5
      Set frmAddEditComType5.ParentForm = Me
      Load frmAddEditComType5
     frmAddEditComType5.Show 1
   
      OKClick = frmAddEditComType5.OKClick
   
      Unload frmAddEditComType5
      Set frmAddEditComType5 = Nothing

   End If
   
   If OKClick Then
      Call RefreshGrid
   End If

   If OKClick Then
      m_HasModify = True
   End If
End Sub

Private Sub cmdDelete_Click()
Dim ID1 As Long
Dim ID2 As Long

   If Not cmdDelete.Enabled Then
      Exit Sub
   End If
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   If Not ConfirmDelete(GridEX1.Value(4) & " - " & GridEX1.Value(5)) Then
      Exit Sub
   End If

   ID2 = GridEX1.Value(2)
   ID1 = GridEX1.Value(1)

   If TabStrip1.SelectedItem.index = 1 Then
      If ID1 <= 0 Then
         m_PromoteYear.Details.Remove (ID2)
      Else
         m_PromoteYear.Details.ITEM(ID2).Flag = "D"
      End If
      m_HasModify = True
' ElseIf TabStrip1.SelectedItem.Index = 2 Then
  End If
  
   Call RefreshGrid
End Sub

Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim IsCanLock As Boolean
Dim ID As Long
Dim OKClick As Boolean

 If Not cmdEdit.Enabled Then
      Exit Sub
   End If
      
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   ID = Val(GridEX1.Value(2))   ' INCENTIVE_PROMOTE_ID
   OKClick = False
   
   If TabStrip1.SelectedItem.index = 1 Then             'ไว้เมื่อแก้ไข
'   frmAddEditINCENTIVE_TYPe2.INCENTIVE_PROMOTE_ID = INCENTIVE_PROMOTE_ID
'  frmAddEditINCENTIVE_TYPe2.MASTER_INCENTIVE_ID = MASTER_INCENTIVE_ID
'      frmAddEditINCENTIVE_TYPe2.HeaderText = MapText("แก้ไขค่า Commission ขาย")
'         frmAddEditINCENTIVE_TYPe2.ShowMode = SHOW_EDIT
'         Load frmAddEditINCENTIVE_TYPe2
'            frmAddEditINCENTIVE_TYPe2.Show 1
'
'            OKClick = frmAddEditINCENTIVE_TYPe2.OKClick
'
'            Unload frmAddEditINCENTIVE_TYPe2
'            Set frmAddEditINCENTIVE_TYPe2 = Nothing
'
'            If OKClick Then
'                     GridEX1.ItemCount = itemCountGrid1
'                     GridEX1.Rebind
'          End If
'
'   ElseIf TabStrip1.SelectedItem.Index = 2 Then
'      frmAddEditINCENTIVE_TYPe3.INCENTIVE_PROMOTE_ID = INCENTIVE_PROMOTE_ID
'  frmAddEditINCENTIVE_TYPe3.MASTER_INCENTIVE_ID = MASTER_INCENTIVE_ID
'  frmAddEditINCENTIVE_TYPe3.INCENTIVE_TYP = "02"
'      frmAddEditINCENTIVE_TYPe3.HeaderText = MapText("แก้ไขค่า Commission ขาย")
'         frmAddEditINCENTIVE_TYPe3.ShowMode = SHOW_EDIT
'         Load frmAddEditINCENTIVE_TYPe3
'            frmAddEditINCENTIVE_TYPe3.Show 1
'
'            OKClick = frmAddEditINCENTIVE_TYPe3.OKClick
'
'            Unload frmAddEditINCENTIVE_TYPe3
'            Set frmAddEditINCENTIVE_TYPe3 = Nothing
'
'            If OKClick Then
'                     GridEX1.ItemCount = itemCountGrid2
'                     GridEX1.Rebind
'          End If
'
'   ElseIf TabStrip1.SelectedItem.Index = 3 Then
'      frmAddEditINCENTIVE_TYPe3.INCENTIVE_PROMOTE_ID = INCENTIVE_PROMOTE_ID
'  frmAddEditINCENTIVE_TYPe3.MASTER_INCENTIVE_ID = MASTER_INCENTIVE_ID
'    frmAddEditINCENTIVE_TYPe3.INCENTIVE_TYP = "03"
'      frmAddEditINCENTIVE_TYPe3.HeaderText = MapText("แก้ไขค่า Commission ขาย")
'         frmAddEditINCENTIVE_TYPe3.ShowMode = SHOW_EDIT
'         Load frmAddEditINCENTIVE_TYPe3
'            frmAddEditINCENTIVE_TYPe3.Show 1
'
'            OKClick = frmAddEditINCENTIVE_TYPe3.OKClick
'
'            Unload frmAddEditINCENTIVE_TYPe3
'            Set frmAddEditINCENTIVE_TYPe3 = Nothing
'
'            If OKClick Then
'                     GridEX1.ItemCount = itemCountGrid3
'                     GridEX1.Rebind
'          End If
'
'   ElseIf TabStrip1.SelectedItem.Index = 4 Then
'         frmAddEditINCENTIVE_TYPe1.INCENTIVE_PROMOTE_ID = INCENTIVE_PROMOTE_ID
'  frmAddEditINCENTIVE_TYPe1.MASTER_INCENTIVE_ID = MASTER_INCENTIVE_ID
'      frmAddEditINCENTIVE_TYPe1.HeaderText = MapText("แก้ไขค่า Commission ขาย")
'         frmAddEditINCENTIVE_TYPe1.ShowMode = SHOW_EDIT
'         Load frmAddEditINCENTIVE_TYPe1
'            frmAddEditINCENTIVE_TYPe1.Show 1
'
'            OKClick = frmAddEditINCENTIVE_TYPe3.OKClick
'
'            Unload frmAddEditINCENTIVE_TYPe1
'            Set frmAddEditINCENTIVE_TYPe1 = Nothing
'
'            If OKClick Then
'                     GridEX1.ItemCount = itemCountGrid4
'                     GridEX1.Rebind
'          End If
'
'   ElseIf TabStrip1.SelectedItem.Index = 5 Then
         Set frmAddEditComType5.ParentForm = Me
         Set frmAddEditComType5.TempCollection = m_PromoteYear.Details
         frmAddEditComType5.ID = ID                        ' ID ของ คอเล็คคชั้น
         frmAddEditComType5.MASTER_INCENTIVE_ID = MASTER_INCENTIVE_ID
         frmAddEditComType5.HeaderText = MapText("แก้ไขค่า Incentive")
         frmAddEditComType5.ShowMode = SHOW_EDIT
         frmAddEditComType5.itemCountGrid = itemCountGrid5
         Load frmAddEditComType5
         frmAddEditComType5.Show 1
            
            OKClick = frmAddEditComType5.OKClick
            
            Unload frmAddEditComType5
            Set frmAddEditComType5 = Nothing

   End If
         
   If OKClick Then
      Call RefreshGrid
      m_HasModify = True
   End If
End Sub
Private Sub cmdOK_Click()
Dim oMenu As cPopupMenu
Dim lMenuChosen  As Long


   Set oMenu = New cPopupMenu
   lMenuChosen = oMenu.Popup("บันทึก", "-", "บันทึกและออก")
   If lMenuChosen = 0 Then
      Exit Sub
   End If
   
   If lMenuChosen = 1 Then
      If Not SaveData Then
         Exit Sub
      End If
      
      ShowMode = SHOW_EDIT
      MASTER_INCENTIVE_ID = m_PromoteYear.MASTER_INCENTIVE_ID
      m_PromoteYear.QueryFlag = 1
      QueryData (True)
      m_HasModify = False
      OKClick = True
   ElseIf lMenuChosen = 3 Then
      If Not SaveData Then
         Exit Sub
      End If
      
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
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_PromoteYear.QueryFlag = 1
         Call QueryData(True)
         Call TabStrip1_Click
      ElseIf ShowMode = SHOW_ADD Then
         m_PromoteYear.QueryFlag = 0
         Call QueryData(False)
      End If
      
      Call EnableForm(Me, True)
      m_HasModify = False
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Static InUsed As Long
   If InUsed = 1 Then
      Exit Sub
   End If
   InUsed = 1
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
      Call cmdAdd_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 117 Then
      Call cmdDelete_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 114 Then
      Call cmdEdit_Click
      KeyCode = 0
'   ElseIf Shift = 0 And KeyCode = 121 Then
'      Call cmdSave_Click
'      KeyCode = 0
   End If
   InUsed = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
'   If cm5_Rs.State = adStateOpen Then
'      cm5_Rs.Close
'   End If
'   Set cm5_Rs = Nothing

   Set m_PromoteYear = Nothing
   'Set m_Condition5 = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   'debug.print ColIndex & " " & NewColWidth
End Sub
'
'Private Sub InitGrid1()
'Dim Col As JSColumn
'
'   GridEX1.Columns.Clear
'   GridEX1.BackColor = GLB_GRID_COLOR
'   GridEX1.ItemCount = 0
'   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
'   GridEX1.ColumnHeaderFont.Bold = True
'   GridEX1.ColumnHeaderFont.Name = GLB_FONT
'   GridEX1.TabKeyBehavior = jgexControlNavigation
'
'   Set Col = GridEX1.Columns.Add '1
'   Col.Width = 0
'   Col.Caption = "ID"
'
'   Set Col = GridEX1.Columns.Add '3
'   Col.Width = 2100
'   Col.Caption = MapText("เงื่อนไข")
'
'   Set Col = GridEX1.Columns.Add '3
'   Col.Width = 2100
'   Col.Caption = MapText("ยอดขาย(%)")
'
'   Set Col = GridEX1.Columns.Add '4
'   Col.Width = 4425 + 3240
'   Col.Caption = MapText("เปอร์เซ็นต์ที่ได้(%)")
'
'End Sub
'
'
'Private Sub InitGrid2()
'Dim Col As JSColumn
'
'   GridEX1.Columns.Clear
'   GridEX1.BackColor = GLB_GRID_COLOR
'   GridEX1.ItemCount = 0
'   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
'   GridEX1.ColumnHeaderFont.Bold = True
'   GridEX1.ColumnHeaderFont.Name = GLB_FONT
'   GridEX1.TabKeyBehavior = jgexControlNavigation
'
'   Set Col = GridEX1.Columns.Add '1
'   Col.Width = 0
'   Col.Caption = "ID"
'
''   Set Col = GridEX1.Columns.Add '2
''   Col.Width = 0
''   Col.Caption = "Real ID"
'
'   Set Col = GridEX1.Columns.Add '3
'   Col.Width = 2100
'   Col.Caption = MapText("เงื่อนไข")
'
'   Set Col = GridEX1.Columns.Add '3
'   Col.Width = 2100
'   Col.Caption = MapText("เครดิตภายใน(วัน)")
'
'   Set Col = GridEX1.Columns.Add '4
'   Col.Width = 4425 + 3240
'   Col.Caption = MapText("คิดเป็น(%)")
'
'End Sub
'
'Private Sub InitGrid3()
'Dim Col As JSColumn
'
'   GridEX1.Columns.Clear
'   GridEX1.BackColor = GLB_GRID_COLOR
'   GridEX1.ItemCount = 0
'   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
'   GridEX1.ColumnHeaderFont.Bold = True
'   GridEX1.ColumnHeaderFont.Name = GLB_FONT
'   GridEX1.TabKeyBehavior = jgexControlNavigation
'
'   Set Col = GridEX1.Columns.Add '1
'   Col.Width = 0
'   Col.Caption = "ID"
'
'   Set Col = GridEX1.Columns.Add '3
'   Col.Width = 2100
'   Col.Caption = MapText("เลขที่สินค้า")
'
'   Set Col = GridEX1.Columns.Add '4
'   Col.Width = 4425
'   Col.Caption = MapText("ชื่อสินค้า")
'
'   Set Col = GridEX1.Columns.Add '4
'   Col.Width = 3240
'   Col.Caption = MapText("คิดเป็น(%)")
'
'End Sub


Private Sub InitGrid4()
Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.ItemCount = 0
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX1.ColumnHeaderFont.Bold = True
   GridEX1.ColumnHeaderFont.Name = GLB_FONT
   GridEX1.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX1.Columns.Add '1
   Col.Width = 0
   Col.Caption = "ID"
   
   Set Col = GridEX1.Columns.Add '2
   Col.Width = 0
   Col.Caption = "Real ID"
   
   Set Col = GridEX1.Columns.Add '2
   Col.Width = 1700
   Col.Caption = MapText("กลุ่มสินค้าที่")

   Set Col = GridEX1.Columns.Add '2
   Col.Width = 2100
   Col.Caption = MapText("เลขที่สินค้า")

   Set Col = GridEX1.Columns.Add '3
   Col.Width = 4350
   Col.Caption = MapText("ชื่อสินค้า")
   
   Set Col = GridEX1.Columns.Add '4
   Col.Width = 1400
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("Pack")
   
 Set Col = GridEX1.Columns.Add '5
   Col.Width = 1400
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("ราคา")
   
Set Col = GridEX1.Columns.Add '6
   Col.Width = 2200
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("Incentive ต่อขวดและถุง")

Set Col = GridEX1.Columns.Add '6
   Col.Width = 1000
   Col.TextAlignment = jgexAlignRight
   Col.Caption = MapText("วัคซีน")

End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
'
   Call InitNormalLabel(lblMasIncentiveNo, MapText("เลขที่"))
      Call InitNormalLabel(lblMasIncentiveDesc, MapText("รายละเอียด"))
   Call InitNormalLabel(lblFromDate, MapText("เริ่มใช้วันที่"))
      Call InitNormalLabel(lblToDate, MapText("ถึงวันที่"))
      
      Call txtMasterIncentiveDesc.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
  Call txtMasIncentiveNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)

   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19


   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))

   
   Call InitGrid4
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.Name = GLB_FONT
   TabStrip1.Font.Size = 16
   
   TabStrip1.Tabs.Clear
'   TabStrip1.Tabs.Add().Caption = MapText("1.ขาย")
'   TabStrip1.Tabs.Add().Caption = MapText("2.เก็บเงิน - เครดิตธรรมดา")
'      TabStrip1.Tabs.Add().Caption = MapText("3.เก็บเงิน - เครดิตสินค้า Commodity")
'      TabStrip1.Tabs.Add().Caption = MapText("4.เก็บเงิน - ตั้งค่าสินค้า Commodity")
         TabStrip1.Tabs.Add().Caption = MapText("Incentive (พิเศษ)")
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
   Set m_PromoteYear = New CIncentiveMasterPromote
 '  Set cm5_Rs = New ADODB.Recordset
  ' Set m_Condition5 = New CIncentivePromote

'   m_HasActivate = False
'   Call EnableForm(Me, True)
End Sub

Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub

Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
   If TabStrip1.SelectedItem.index = 1 Then
      RowBuffer.RowStyle = RowBuffer.Value(7)
   End If
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"

 If TabStrip1.SelectedItem.index = 1 Then
      If m_PromoteYear.Details Is Nothing Then
         Exit Sub
      End If
   
      If RowIndex <= 0 Then
         Exit Sub
      End If
   
      Dim CR As CIncentivePromote
      If m_PromoteYear.Details.Count <= 0 Then
         Exit Sub
      End If
      Set CR = GetItem(m_PromoteYear.Details, RowIndex, RealIndex)
      If CR Is Nothing Then
         Exit Sub
      End If
 
'
'   ElseIf TabStrip1.SelectedItem.Index = 2 Then
'            If cm2_Rs Is Nothing Then
'      Exit Sub
'   End If
'   If cm2_Rs.State <> adStateOpen Then
'      Exit Sub
'   End If
'   If cm2_Rs.EOF Then
'      Exit Sub
'   End If
'      If RowIndex <= 0 Then
'         Exit Sub
'      End If
'   Call cm2_Rs.Move(RowIndex - 1, adBookmarkFirst)
'    Call m_Condition2.PopulateFromRS(2, cm2_Rs)   ' ใช้ป๊อปแบบที่ 2 เพราะจะได้ ตัวเลข NUM_ONE ออกมา
'      Values(1) = m_Condition2.INCENTIVE_PROMOTE_ID      ' values(1) เกี่ยวพันกับการลบ
'      Values(2) = "<="                                             ' operatorToText(m_Condition2.OPERATOR)
'      Values(3) = m_Condition2.NUM_ONE
'      Values(4) = m_Condition2.SLM_PERCENT
'
'ElseIf TabStrip1.SelectedItem.Index = 3 Then
'         If cm3_Rs Is Nothing Then
'      Exit Sub
'   End If
'   If cm3_Rs.State <> adStateOpen Then
'      Exit Sub
'   End If
'   If cm3_Rs.EOF Then
'      Exit Sub
'   End If
'      If RowIndex <= 0 Then
'         Exit Sub
'      End If
'   Call cm3_Rs.Move(RowIndex - 1, adBookmarkFirst)
'    Call m_Condition3.PopulateFromRS(2, cm3_Rs)   ' ใช้ป๊อปแบบที่ 2 เพราะจะได้ ตัวเลข NUM_ONE ออกมา
'      Values(1) = m_Condition3.INCENTIVE_PROMOTE_ID                    ' values(1) เกี่ยวพันกับการลบ
'            Values(2) = "<="                                           ' operatorToText(m_Condition3.OPERATOR)
'      Values(3) = m_Condition3.NUM_ONE
'      Values(4) = m_Condition3.SLM_PERCENT
'
'ElseIf TabStrip1.SelectedItem.Index = 4 Then
'      If cm4_Rs Is Nothing Then
'      Exit Sub
'   End If
'   If cm4_Rs.State <> adStateOpen Then
'      Exit Sub
'   End If
'   If cm4_Rs.EOF Then
'      Exit Sub
'   End If
'      If RowIndex <= 0 Then
'         Exit Sub
'      End If
'   Call cm4_Rs.Move(RowIndex - 1, adBookmarkFirst)
'    Call m_Condition4.PopulateFromRS(1, cm4_Rs)
'      Values(1) = m_Condition4.INCENTIVE_PROMOTE_ID
'      Values(2) = m_Condition4.STKCOD
'      Values(3) = m_Condition4.STKDES
'      Values(4) = m_Condition4.SLM_PERCENT
'
'   ElseIf TabStrip1.SelectedItem.Index = 5 Then
'         If cm5_Rs Is Nothing Then
'      Exit Sub
'   End If
'   If cm5_Rs.State <> adStateOpen Then
'      Exit Sub
'   End If
'   If cm5_Rs.EOF Then
'      Exit Sub
'   End If
'      If RowIndex <= 0 Then
'         Exit Sub
'      End If
'   Call cm5_Rs.Move(RowIndex - 1, adBookmarkFirst)
'    Call m_Condition5.PopulateFromRS(1, cm5_Rs)        ' ไปดูอีกที จะใช้ pop ไร เพราะมันติ๊กว่าเป็นวัคซีน
      Values(1) = CR.INCENTIVE_PROMOTE_ID
            Values(2) = RealIndex
      Values(3) = CR.GROUP1
      Values(4) = CR.STKCOD
      Values(5) = CR.STKDES
      Values(6) = CR.NUM_ONE           ' Pack
      Values(7) = CR.NUM_TWO        ' Price
      Values(8) = CR.SLM_PERCENT   ' Incentive
      Values(9) = CR.OPERATOR   ' Incentive
      
   End If
      
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub TabStrip1_Click()
   If TabStrip1.SelectedItem Is Nothing Then
      Exit Sub
   End If
   
   If TabStrip1.SelectedItem.index = 1 Then
      Call InitGrid4
      GridEX1.ItemCount = CountItem(m_PromoteYear.Details)
      GridEX1.Rebind
   End If
End Sub

Private Sub txtMasIncentiveNo_Change()
   m_HasModify = True
End Sub

Private Sub uctlToDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlFromDate_HasChange()
   m_HasModify = True
End Sub

'Private Sub txtDocumentNo_LostFocus()
'   If Not CheckUniqueNs(IMPORT_UNIQUE, txtDocumentNo.Text, ID) Then
'      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtDocumentNo.Text & " " & MapText("อยู่ในระบบแล้ว")
'      glbErrorLog.ShowUserError
'
'      'txtDocumentNo.SetFocus
'      txtDocumentNo.Text = ""
'      Exit Sub
'   End If
'End Sub

'Private Sub uctlDocumentDate_LostFocus()
'   If ShowMode = SHOW_ADD And uctlDocumentDate.ShowDate > 0 Then
'      If Not VerifyDateInterval(uctlDocumentDate.ShowDate) Then
'         uctlDocumentDate.SetFocus
'         Exit Sub
'      End If
'   ElseIf Not CheckUniqueNs(IMPORT_UNIQUE, txtDocumentNo.Text, ID) Then
'      txtDocumentNo.SetFocus
'      Exit Sub
'   ElseIf Not (uctlDocumentDate.ShowDate > 0) Then
'      uctlDocumentDate.SetFocus
'      Exit Sub
'   End If
'End Sub
Public Sub RefreshGrid()
If TabStrip1.SelectedItem.index = 1 Then
   GridEX1.ItemCount = CountItem(m_PromoteYear.Details)
   GridEX1.Rebind
ElseIf TabStrip1.SelectedItem.index = 2 Then
   GridEX1.ItemCount = CountItem(m_PromoteYear.CommissionExs)
   GridEX1.Rebind
End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
   SSFrame1.Width = ScaleWidth
   SSFrame1.Height = ScaleHeight
   pnlHeader.Width = ScaleWidth
   GridEX1.Width = ScaleWidth - 2 * GridEX1.Left
   GridEX1.Height = ScaleHeight - GridEX1.Top - 620
   TabStrip1.Width = GridEX1.Width
   cmdAdd.Top = ScaleHeight - 580
   cmdEdit.Top = ScaleHeight - 580
   cmdDelete.Top = ScaleHeight - 580
   cmdOK.Top = ScaleHeight - 580
   cmdExit.Top = ScaleHeight - 580
   cmdExit.Left = ScaleWidth - cmdExit.Width - 50
   cmdOK.Left = cmdExit.Left - cmdOK.Width - 50
End Sub

Private Sub GridEX1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = DUMMY_KEY Then
      Call cmdExit_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
      KeyCode = 0
   End If
End Sub

