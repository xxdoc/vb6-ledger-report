VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmAreaCommission 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmAreaCommission.frx":0000
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
      TabIndex        =   8
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
         TabIndex        =   9
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjLedgerReport.uctlDate uctlFromDate 
         Height          =   375
         Left            =   2880
         TabIndex        =   1
         Top             =   1440
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
      End
      Begin prjLedgerReport.uctlTextBox txtYearNum 
         Height          =   495
         Left            =   2880
         TabIndex        =   0
         Top             =   840
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   873
      End
      Begin prjLedgerReport.uctlDate uctlToDate 
         Height          =   375
         Left            =   2880
         TabIndex        =   2
         Top             =   1920
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   4725
         Left            =   120
         TabIndex        =   3
         Top             =   2640
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   8334
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
         Column(1)       =   "frmAreaCommission.frx":27A2
         Column(2)       =   "frmAreaCommission.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAreaCommission.frx":290E
         FormatStyle(2)  =   "frmAreaCommission.frx":2A6A
         FormatStyle(3)  =   "frmAreaCommission.frx":2B1A
         FormatStyle(4)  =   "frmAreaCommission.frx":2BCE
         FormatStyle(5)  =   "frmAreaCommission.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAreaCommission.frx":2D5E
      End
      Begin Threed.SSCommand cmdCusInEP 
         Height          =   525
         Left            =   8400
         TabIndex        =   13
         Top             =   7800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin VB.Label lblFromDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   960
         TabIndex        =   12
         Top             =   1440
         Width           =   1755
      End
      Begin VB.Label lblToDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   1320
         TabIndex        =   11
         Top             =   1920
         Width           =   1365
      End
      Begin VB.Label lblYearNum 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   960
         TabIndex        =   10
         Top             =   840
         Width           =   1755
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   10080
         TabIndex        =   7
         Top             =   7800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAreaCommission.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1770
         TabIndex        =   5
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
         TabIndex        =   4
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAreaCommission.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   6
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAreaCommission.frx":356A
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAreaCommission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'frmCommission2
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private masterArea As CCommissionCustomerArea   'CCommissionCustomerArea
Private tempMasterArea As CCommissionCustomerArea
Private m_AreaYear As CAreaYear
Private m_Rs As ADODB.Recordset
Private m_TableName As String
Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean

Public YEAR_ID As Long
Private Rs As ADODB.Recordset

Private Sub cmdAdd_Click()
Dim ItemCount As Long
Dim OKClick As Boolean
Dim IsOK As Boolean

   If YEAR_ID = 0 And txtYearNum.Text <> "" Then   ' เพื่อยังไม่ทันกดตกลง ยูสเซอร์จะได้สามารถกดเพิ่มได้
      Call SaveData
    
                  m_AreaYear.YEARNUM = txtYearNum.Text    'ยังอยู่ในโหมด edit
                  m_AreaYear.FROM_DATE = uctlFromDate.ShowDate
                  m_AreaYear.TO_DATE = uctlToDate.ShowDate
                   If Not glbDaily.QueryAreaYear(m_AreaYear, m_Rs, ItemCount, IsOK, glbErrorLog) Then
                        glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
                        Call EnableForm(Me, True)
                        Exit Sub
                     End If
             Call m_AreaYear.PopulateFromRS(1, m_Rs)
             YEAR_ID = m_AreaYear.YEAR_ID
End If

   frmAddEditAreaCom.HeaderText = MapText("เพิ่มเขตการขาย")
   frmAddEditAreaCom.ShowMode = SHOW_ADD
   frmAddEditAreaCom.YEAR_ID = YEAR_ID
  Load frmAddEditAreaCom
   frmAddEditAreaCom.Show 1
   
   OKClick = frmAddEditAreaCom.OKClick
   ShowMode = frmAddEditAreaCom.ShowMode
   Unload frmAddEditAreaCom
   Set frmAddEditAreaCom = Nothing
   
   If OKClick Then
      Call QueryData(True)
   End If
End Sub


Private Sub cmdCusInEP_Click()
      frmAddEditAreaInEP.YEAR_ID = YEAR_ID
      frmAddEditAreaInEP.ShowMode = SHOW_EDIT
      Load frmAddEditAreaInEP
      frmAddEditAreaInEP.Show 1
      
      Unload frmAddEditAreaInEP
      Set frmAddEditAreaInEP = Nothing
End Sub

Private Sub cmdDelete_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim MASTER_AREA_ID As Long


   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
  MASTER_AREA_ID = GridEX1.Value(1)

   If Not ConfirmDelete(GridEX1.Value(2)) Then
      Call glbDatabaseMngr.UnLockTable(m_TableName, MASTER_AREA_ID, IsCanLock, glbErrorLog)
      Exit Sub
   End If

   Call EnableForm(Me, False)
   
'   If Not glbDaily.DeleteCommissArea(MASTER_AREA_ID, IsOK, True, glbErrorLog) Then
'       masterArea.MASTER_AREA_ID = -1
'      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
'      Call glbDatabaseMngr.UnLockTable(m_TableName, MASTER_AREA_ID, IsCanLock, glbErrorLog)
'      Call EnableForm(Me, True)
'      Exit Sub
'   End If
   
 '   ??????????? COM_ID
      If Not glbDaily.DeleteCusAreaComFromMaster(MASTER_AREA_ID, IsOK, True, glbErrorLog) Then
       masterArea.MASTER_AREA_ID = -1
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      Call glbDatabaseMngr.UnLockTable(m_TableName, MASTER_AREA_ID, IsCanLock, glbErrorLog)
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call QueryData(True)
   
   Call glbDatabaseMngr.UnLockTable(m_TableName, MASTER_AREA_ID, IsCanLock, glbErrorLog)
   Call EnableForm(Me, True)
End Sub

Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim MASTER_AREA_ID As Long
Dim OKClick As Boolean

      
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

  MASTER_AREA_ID = Val(GridEX1.Value(1))
 ' Call glbDatabaseMngr.LockTable(m_TableName, MASTER_AREA_ID, IsCanLock, glbErrorLog)
   frmAddEditAreaCom.YEAR_ID = YEAR_ID
   frmAddEditAreaCom.MASTER_AREA_ID = MASTER_AREA_ID
   frmAddEditAreaCom.HeaderText = MapText("แก้ไขเขตการขาย")
   frmAddEditAreaCom.ShowMode = SHOW_EDIT
   Load frmAddEditAreaCom
   frmAddEditAreaCom.Show 1
   
   OKClick = frmAddEditAreaCom.OKClick
   
   Unload frmAddEditAreaCom
   Set frmAddEditAreaCom = Nothing
               
   If OKClick Then
      Call QueryData(True)
   End If
   Call glbDatabaseMngr.UnLockTable(m_TableName, MASTER_AREA_ID, IsCanLock, glbErrorLog)

End Sub

Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If

   OKClick = True
   Unload Me
End Sub

Private Sub cmdSearch_Click()
   Call QueryData(True)
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
         m_HasModify = False
         
     If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         YEAR_ID = -1
      End If

   End If
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim YearCount As Long
Dim Temp As Long

      Call EnableForm(Me, False)
'   If ShowMode = SHOW_VIEW Then
'
'      tempMasterArea.YEAR_ID = YEAR_ID
'      If Not glbDaily.QueryCusAreaCom(tempMasterArea, m_Rs, ItemCount, IsOK, glbErrorLog) Then
'         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
'         Call EnableForm(Me, True)
'         Exit Sub
'      End If
'
'     If Not IsOK Then
'         glbErrorLog.ShowUserError
'         Call EnableForm(Me, True)
'         Exit Sub
'      End If
'
'      GridEX1.ItemCount = ItemCount
'      GridEX1.Rebind
'
'ElseIf ShowMode = SHOW_EDIT Then

'       masterArea.YEAR_ID = YEAR_ID
'      If Not glbDaily.QueryCusAreaCom(masterArea, m_Rs, ItemCount, IsOK, glbErrorLog) Then
'         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
'         Call EnableForm(Me, True)
'         Exit Sub
'      End If

   
     m_AreaYear.YEAR_ID = YEAR_ID
     m_AreaYear.FROM_DATE = -1
     m_AreaYear.TO_DATE = -1
     If Not glbDaily.QueryAreaYear(m_AreaYear, Rs, YearCount, IsOK, glbErrorLog) Then
        glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
        Call EnableForm(Me, True)
        Exit Sub
     End If
     
     Call m_AreaYear.PopulateFromRS(1, Rs)
     txtYearNum.Text = m_AreaYear.YEARNUM    'ยังอยู่ในโหมด edit
     uctlFromDate.ShowDate = m_AreaYear.FROM_DATE
     uctlToDate.ShowDate = m_AreaYear.TO_DATE

      tempMasterArea.YEAR_ID = YEAR_ID
      tempMasterArea.MASTER_AREA_ID = -1
      If Not glbDaily.QueryCusAreaCom(tempMasterArea, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
      
      If Not IsOK Then
         glbErrorLog.ShowUserError
         Call EnableForm(Me, True)
         Exit Sub
      End If
   
      GridEX1.ItemCount = m_Rs.RecordCount
      GridEX1.Rebind
'End If
   Call EnableForm(Me, True)
End Sub


Private Function SaveData() As Boolean
Dim IsOK As Boolean

''   If ShowMode = SHOW_ADD Then
''      If Not VerifyAccessRight("PIG_ADJUST_ADD") Then
''         Call EnableForm(Me, True)
''         Exit Function
''      End If
''   ElseIf ShowMode = SHOW_EDIT Then
''      If Not VerifyAccessRight("PIG_ADJUST_EDIT") Then
''         Call EnableForm(Me, True)
''         Exit Function
''      End If
''   End If

   If Not VerifyTextControl(lblYearNum, txtYearNum, False) Then
      Exit Function
   End If
   If Not VerifyDate(lblFromDate, uctlFromDate, False) Then
      Exit Function
   End If
   If Not VerifyDate(lblToDate, uctlToDate, False) Then
      Exit Function
   End If
   
   If Not CheckUniqueNs(EXPORT_UNIQUE, txtYearNum.Text, YEAR_ID) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtYearNum.Text & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
If YEAR_ID = 0 Then
   m_AreaYear.AddEditMode = ShowMode                        ' ตรงนี้ ตอนบันทึก
Else:
 m_AreaYear.AddEditMode = SHOW_EDIT
 End If
  ' m_CommissYear.YEAR_ID = YEAR_ID
 m_AreaYear.YEARNUM = txtYearNum.Text
    m_AreaYear.FROM_DATE = uctlFromDate.ShowDate
   m_AreaYear.TO_DATE = uctlToDate.ShowDate
   
   Call EnableForm(Me, False)
   
   If Not glbDaily.AddEditAreaYear(m_AreaYear, IsOK, True, glbErrorLog) Then
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
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.Name
      glbErrorLog.ShowUserError
      KeyCode = 0
'   ElseIf Shift = 0 And KeyCode = 116 Then
'      Call cmdSearch_Click
'      KeyCode = 0
'   ElseIf Shift = 0 And KeyCode = 115 Then
'      Call cmdClear_Click
'      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 118 Then
      Call cmdAdd_Click
      KeyCode = 0
'   ElseIf Shift = 0 And KeyCode = 117 Then
      Call cmdDelete_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 114 Then
      Call cmdEdit_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 121 Then
'      Call cmdPrint_Click
      KeyCode = 0
   End If
End Sub

Private Sub InitGrid()
Dim Col As JSColumn
Dim fmsTemp As JSFormatStyle

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   
   GridEX1.FormatStyles.Clear
   Set fmsTemp = GridEX1.FormatStyles.Add("N")
   fmsTemp.ForeColor = GLB_ALERT_COLOR

   Set Col = GridEX1.Columns.Add '1
   Col.Width = 1785
   Col.Caption = MapText("รหัสเขต")
      
   Set Col = GridEX1.Columns.Add '2
   Col.Width = 2500
   Col.Caption = MapText("ชื่อเขตการขาย")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   
     Call InitNormalLabel(lblYearNum, MapText("เลขที่"))
   Call InitNormalLabel(lblFromDate, MapText("เริ่มใช้วันที่"))
      Call InitNormalLabel(lblToDate, MapText("ถึงวันที่"))
   
  Call txtYearNum.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   
   Call InitGrid
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
 '  cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdCusInEP.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
 '  Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdCusInEP, MapText("ลูกค้า Express"))

End Sub

Private Sub cmdExit_Click()
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   m_TableName = "COMMISSION_MASTER_AREA"
   OKClick = False
      m_HasModify = False
   m_HasActivate = False
   Set masterArea = New CCommissionCustomerArea
   Set tempMasterArea = New CCommissionCustomerArea
   Set m_AreaYear = New CAreaYear
   Set m_Rs = New ADODB.Recordset
   Set Rs = New ADODB.Recordset
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   'debug.print ColIndex & " " & NewColWidth
End Sub

Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"

   If m_Rs Is Nothing Then
      Exit Sub
   End If

   If m_Rs.State <> adStateOpen Then
      Exit Sub
   End If

   If m_Rs.EOF Then
      Exit Sub
   End If

   If RowIndex <= 0 Then
      Exit Sub
   End If

   Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
   Call tempMasterArea.PopulateFromRS(1, m_Rs)
   
   Values(1) = tempMasterArea.MASTER_AREA_ID
   Values(2) = tempMasterArea.MASTER_AREA_NAME
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub Form_Resize()
On Error Resume Next
   SSFrame1.Width = ScaleWidth
   SSFrame1.Height = ScaleHeight
   pnlHeader.Width = ScaleWidth
   GridEX1.Width = ScaleWidth - 2 * GridEX1.Left
   GridEX1.Height = ScaleHeight - GridEX1.Top - 620
   cmdAdd.Top = ScaleHeight - 580
   cmdEdit.Top = ScaleHeight - 580
   cmdDelete.Top = ScaleHeight - 580
   cmdOK.Top = ScaleHeight - 580
   cmdCusInEP.Top = ScaleHeight - 580
   cmdCusInEP.Left = ScaleWidth - cmdCusInEP.Width - cmdOK.Width - 50
   cmdOK.Left = ScaleWidth - cmdOK.Width - 50
End Sub

Private Sub txtYearNum_Change()
   m_HasModify = True
End Sub
Private Sub uctlFromDate_HasChange()
   m_HasModify = True
End Sub
Private Sub uctlToDate_HasChange()
   m_HasModify = True
End Sub
