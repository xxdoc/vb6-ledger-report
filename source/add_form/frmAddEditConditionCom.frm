VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmAddEditConditionCom 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8670
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   12060
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8670
   ScaleWidth      =   12060
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   11655
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   19035
      _ExtentX        =   33576
      _ExtentY        =   20558
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Threed.SSPanel ppnlHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   5
         Top             =   0
         Width           =   19005
         _ExtentX        =   33523
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   240
         TabIndex        =   9
         Top             =   2640
         Width           =   18555
         _ExtentX        =   32729
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
         Height          =   7005
         Left            =   240
         TabIndex        =   10
         Top             =   3195
         Width           =   18555
         _ExtentX        =   32729
         _ExtentY        =   12356
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
         Column(1)       =   "frmAddEditConditionCom.frx":0000
         Column(2)       =   "frmAddEditConditionCom.frx":00C8
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditConditionCom.frx":016C
         FormatStyle(2)  =   "frmAddEditConditionCom.frx":02C8
         FormatStyle(3)  =   "frmAddEditConditionCom.frx":0378
         FormatStyle(4)  =   "frmAddEditConditionCom.frx":042C
         FormatStyle(5)  =   "frmAddEditConditionCom.frx":0504
         ImageCount      =   0
         PrinterProperties=   "frmAddEditConditionCom.frx":05BC
      End
      Begin prjLedgerReport.uctlDate uctlFromDate 
         Height          =   375
         Left            =   2280
         TabIndex        =   11
         Top             =   1560
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
      End
      Begin prjLedgerReport.uctlDate uctlToDate 
         Height          =   375
         Left            =   2280
         TabIndex        =   12
         Top             =   2040
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
      End
      Begin prjLedgerReport.uctlTextBox txtYearNum 
         Height          =   495
         Left            =   2280
         TabIndex        =   13
         Top             =   960
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   873
      End
      Begin VB.Label lblYearNum 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   1755
      End
      Begin VB.Label lblToDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   600
         TabIndex        =   7
         Top             =   2040
         Width           =   1365
      End
      Begin VB.Label lblFromDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   240
         TabIndex        =   6
         Top             =   1560
         Width           =   1755
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3480
         TabIndex        =   2
         Top             =   10680
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   150
         TabIndex        =   0
         Top             =   10680
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1800
         TabIndex        =   1
         Top             =   10680
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   17160
         TabIndex        =   3
         Top             =   10560
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditConditionCom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset

Private m_CommissYear As CCommissYear
Private m_Condition1 As CConditionCommission
Private m_Condition2 As CConditionCommission
Private m_Condition3 As CConditionCommission
Private m_Condition4 As CConditionCommission
Private m_Condition5 As CConditionCommission        ' incen
Private m_Condition6 As CConditionCommission       ' incen
Private cm1_Rs As ADODB.Recordset
Private cm2_Rs As ADODB.Recordset
Private cm3_Rs As ADODB.Recordset
Private cm4_Rs As ADODB.Recordset
Private cm5_Rs As ADODB.Recordset    '' incentive
Private cm6_Rs As ADODB.Recordset     ' incentive


 
Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean

Public COM_ID As Long
Public COMTYP As String
Public STKCOD As String
Public STKDES As String
Public SLM_PERCENT As Long
Public YEAR_ID As Long
Public YEARNUM As String
Public Flag As String

Public FROM_DATE As String
Public TO_DATE As String

Dim ItemCount As Long
Dim itemCountGrid1 As Long
Dim itemCountGrid2 As Long
Dim itemCountGrid3 As Long
Dim itemCountGrid4 As Long
Dim itemCountGrid5 As Long     ' incen
Dim itemCountGrid6 As Long    ' incen

Private m_TableName As String
Private FileName As String
Private m_SumUnit As Double

Public Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean

   IsOK = True

      Call EnableForm(Me, False)
        '  If ShowMode = SHOW_ADD Then    'ในนี้ล่ะ ที่จะแสดงหลังเพิ่มข้อมูลย่อย
          
'                    m_Condition1.YEAR_ID = YEAR_ID
'                     If Not glbDaily.QueryConditionCom(m_Condition1, cm1_Rs, itemCountGrid1, IsOK, glbErrorLog) Then
'                        glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
'                        Call EnableForm(Me, True)
'                        Exit Sub
'                     End If
  '     End If
 '
  If ShowMode = SHOW_EDIT Then

                     m_Condition1.YEAR_ID = YEAR_ID
                    m_Condition1.COM_ID = 0
                     m_Condition1.COMTYP = "01"
                     m_Condition1.FROM_CMPL_DATE = -1
                     m_Condition1.TO_CMPL_DATE = -1
                     If Not glbDaily.QueryConditionCom(m_Condition1, cm1_Rs, itemCountGrid1, IsOK, glbErrorLog) Then
                        glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
                        Call EnableForm(Me, True)
                        Exit Sub
                     End If
                
                      m_Condition2.YEAR_ID = YEAR_ID
                      m_Condition2.COMTYP = "02"
                  m_Condition2.COM_ID = 0
                                       m_Condition2.FROM_CMPL_DATE = -1
                     m_Condition2.TO_CMPL_DATE = -1
                     If Not glbDaily.QueryConditionCom(m_Condition2, cm2_Rs, itemCountGrid2, IsOK, glbErrorLog) Then
                        glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
                        Call EnableForm(Me, True)
                        Exit Sub
                     End If
                
                      m_Condition3.YEAR_ID = YEAR_ID
                      m_Condition3.COMTYP = "03"
                                        m_Condition3.COM_ID = 0
                                                             m_Condition3.FROM_CMPL_DATE = -1
                     m_Condition3.TO_CMPL_DATE = -1
                     If Not glbDaily.QueryConditionCom(m_Condition3, cm3_Rs, itemCountGrid3, IsOK, glbErrorLog) Then
                        glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
                        Call EnableForm(Me, True)
                        Exit Sub
                     End If
                     
                      m_Condition4.YEAR_ID = YEAR_ID
                      m_Condition4.COMTYP = "04"
                                          m_Condition4.COM_ID = 0
                                          m_Condition4.FROM_CMPL_DATE = -1
                     m_Condition4.TO_CMPL_DATE = -1
                     If Not glbDaily.QueryConditionCom(m_Condition4, cm4_Rs, itemCountGrid4, IsOK, glbErrorLog) Then
                        glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
                        Call EnableForm(Me, True)
                        Exit Sub
                     End If
                     
                      m_Condition5.YEAR_ID = YEAR_ID                          ' อันเดียวก็พอ มีคอลัมป์ติ๊กว่าเป็น วัคซีนป่าว
                      m_Condition5.COMTYP = "05"
                                          m_Condition5.COM_ID = 0
                                          m_Condition5.FROM_CMPL_DATE = -1
                     m_Condition5.TO_CMPL_DATE = -1
                     m_Condition5.STKCOD = ""
                     If Not glbDaily.QueryConditionCom(m_Condition5, cm5_Rs, itemCountGrid5, IsOK, glbErrorLog) Then
                        glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
                        Call EnableForm(Me, True)
                        Exit Sub
                     End If
                     

                     
               '
               '   Dim j As Long
               '   If itemCountGrid > 0 Then  ' จนกว่าจะครบ จ.รอบx3
               '    While Not cm_Rs.EOF
               '      Call m_Condition.PopulateFromRS(1, cm_Rs)    'เพราะแต่ละเคส
               '      'ตรงนี้ไว้ใส่ข้อมูลในแต่ละเคส
               '            cm_Rs.MoveNext
               '      Wend
               '   End If
                  
                   m_CommissYear.YEAR_ID = YEAR_ID
                   If Not glbDaily.QueryCommissYear(m_CommissYear, m_Rs, ItemCount, IsOK, glbErrorLog) Then
                        glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
                        Call EnableForm(Me, True)
                        Exit Sub
                     End If
                     Call m_CommissYear.PopulateFromRS(1, m_Rs)
                     txtYearNum.Text = m_CommissYear.YEARNUM    'ยังอยู่ในโหมด edit
                     uctlFromDate.ShowDate = m_CommissYear.FROM_DATE
                     uctlToDate.ShowDate = m_CommissYear.TO_DATE

                  If Not IsOK Then
                     glbErrorLog.ShowUserError
                     Call EnableForm(Me, True)
                     Exit Sub
                  End If
   
 End If
   
   Call TabStrip1_Click
   Call EnableForm(Me, True)
End Sub

Private Sub cboGroup_Click()
   m_HasModify = True
End Sub

Private Sub chkEnable_Click()
   m_HasModify = True
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
  '    Exit Function
   End If
   

If YEAR_ID = 0 Then
   m_CommissYear.AddEditMode = ShowMode                        ' ตรงนี้ ตอนบันทึก
Else:
 m_CommissYear.AddEditMode = SHOW_EDIT
 End If
  ' m_CommissYear.YEAR_ID = YEAR_ID
 m_CommissYear.YEARNUM = txtYearNum.Text
    m_CommissYear.FROM_DATE = uctlFromDate.ShowDate
   m_CommissYear.TO_DATE = uctlToDate.ShowDate
   
   Call EnableForm(Me, False)
   
   If Not glbDaily.AddEditCommissYear(m_CommissYear, IsOK, True, glbErrorLog) Then
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

Private Sub chkCommit_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub cmdAdd_Click()    ' เมื่อกดเพิ่มแต่ละ case 3 case
Dim OKClick As Boolean
Dim IsOK As Boolean


   IsOK = True
   If Not cmdAdd.Enabled Then
      Exit Sub
   End If
   
   If YEAR_ID = 0 And txtYearNum.Text <> "" Then   ' เพื่อยังไม่ทันกดตกลง ยูสเซอร์จะได้สามารถกดเพิ่มได้
      Call SaveData
    
                  m_CommissYear.YEARNUM = txtYearNum.Text
                   If Not glbDaily.QueryCommissYear(m_CommissYear, m_Rs, ItemCount, IsOK, glbErrorLog) Then
                        glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
                        Call EnableForm(Me, True)
                        Exit Sub
                     End If
             Call m_CommissYear.PopulateFromRS(1, m_Rs)
             YEAR_ID = m_CommissYear.YEAR_ID
End If
   
   
   OKClick = False
If TabStrip1.SelectedItem.Index = 1 Then
 frmAddEditComType2.YEAR_ID = YEAR_ID
 frmAddEditComType2.ShowMode = SHOW_ADD
     frmAddEditComType2.HeaderText = MapText("เพิ่ม Commission ขาย")
   Load frmAddEditComType2
  frmAddEditComType2.Show 1
   
   OKClick = frmAddEditComType2.OKClick
   
   
   Unload frmAddEditComType2
   Set frmAddEditComType2 = Nothing
   
   If OKClick Then
      frmAddEditConditionCom.ShowMode = SHOW_EDIT
      Call QueryData(True)
   End If


   ElseIf TabStrip1.SelectedItem.Index = 2 Then
 frmAddEditComType3.YEAR_ID = YEAR_ID
 frmAddEditComType3.ShowMode = SHOW_ADD
 frmAddEditComType3.COMTYP = "02"
   frmAddEditComType3.HeaderText = MapText("เพิ่ม Commission เก็บเงิน(1)")
   Load frmAddEditComType3
  frmAddEditComType3.Show 1
   
   OKClick = frmAddEditComType3.OKClick
   
   Unload frmAddEditComType3
   Set frmAddEditComType3 = Nothing
   
   If OKClick Then
       frmAddEditConditionCom.ShowMode = SHOW_EDIT
      Call QueryData(True)
   End If

   ElseIf TabStrip1.SelectedItem.Index = 3 Then
  frmAddEditComType3.YEAR_ID = YEAR_ID
 frmAddEditComType3.ShowMode = SHOW_ADD
 frmAddEditComType3.COMTYP = "03"
 frmAddEditComType3.HeaderText = MapText("เพิ่ม Commission เก็บเงิน(2)")
   Load frmAddEditComType3
  frmAddEditComType3.Show 1
   
   OKClick = frmAddEditComType3.OKClick
   
   Unload frmAddEditComType3
   Set frmAddEditComType3 = Nothing
   
   If OKClick Then
       frmAddEditConditionCom.ShowMode = SHOW_EDIT
      Call QueryData(True)
   End If
   
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
    frmAddEditComType1.YEAR_ID = YEAR_ID
      frmAddEditComType1.ShowMode = SHOW_ADD
      frmAddEditComType1.HeaderText = MapText("เพิ่ม Commission เก็บเงิน(3)")
   Load frmAddEditComType1
  frmAddEditComType1.Show 1
   
   OKClick = frmAddEditComType1.OKClick
   
   Unload frmAddEditComType1
   Set frmAddEditComType1 = Nothing
   
   If OKClick Then
       frmAddEditConditionCom.ShowMode = SHOW_EDIT
      Call QueryData(True)
   End If
   
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
       frmAddEditComType4.YEAR_ID = YEAR_ID
      frmAddEditComType4.ShowMode = SHOW_ADD
      frmAddEditComType4.HeaderText = MapText("เพิ่ม Incentive ")
      frmAddEditComType4.itemCountGrid = itemCountGrid5
      Set frmAddEditComType4.ParentForm = Me
      Load frmAddEditComType4
     frmAddEditComType4.Show 1
   
      OKClick = frmAddEditComType4.OKClick
   
      Unload frmAddEditComType4
      Set frmAddEditComType4 = Nothing
   
   If OKClick Then
       frmAddEditConditionCom.ShowMode = SHOW_EDIT
      Call QueryData(True)
   End If
   
   End If
   
   If OKClick Then
      m_HasModify = True
   End If
End Sub

Private Sub cmdDelete_Click()
Dim IsCanLock As Boolean
Dim IsOK As Boolean

   If Not cmdDelete.Enabled Then
      Exit Sub
   End If
   
      If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
  COM_ID = GridEX1.Value(1)

   If Not ConfirmDelete(GridEX1.Value(2) & " " & GridEX1.Value(3)) Then
      Call glbDatabaseMngr.UnLockTable(m_TableName, COM_ID, IsCanLock, glbErrorLog)
      Exit Sub
   End If

   Call EnableForm(Me, False)

   
'
'   If Not VerifyGrid(GridEX1.Value(1)) Then
'      Exit Sub
'   End If
'
'   If Not ConfirmDelete(GridEX1.Value(3)) Then
'      Exit Sub
'   End If
'
'   ID2 = GridEX1.Value(2)
'   ID1 = GridEX1.Value(1)

   If TabStrip1.SelectedItem.Index = 1 Then
'      If ID1 <= 0 Then
'         m_CommissYear.ImportItems.Remove (ID2)
'      Else
'         m_CommissYear.ImportItems.ITEM(ID2).Flag = "D"
'      End If
'      GridEX1.ItemCount = itemCountGrid1
'      GridEX1.Rebind
'      m_HasModify = True
   If Not glbDaily.DeleteConditionCom(COM_ID, IsOK, True, glbErrorLog) Then
       m_Condition1.COM_ID = -1
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      Call glbDatabaseMngr.UnLockTable(m_TableName, COM_ID, IsCanLock, glbErrorLog)
      Call EnableForm(Me, True)
      Exit Sub
      End If


   ElseIf TabStrip1.SelectedItem.Index = 2 Then
'      If ID1 <= 0 Then
'         m_CommissYear.ExportItems.Remove (ID2)
'      Else
'         m_CommissYear.ExportItems.ITEM(ID2).Flag = "D"
'      End If
'      GridEX1.ItemCount = itemCountGrid2
'      GridEX1.Rebind
'      m_HasModify = True
   If Not glbDaily.DeleteConditionCom(COM_ID, IsOK, True, glbErrorLog) Then
       m_Condition2.COM_ID = -1
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      Call glbDatabaseMngr.UnLockTable(m_TableName, COM_ID, IsCanLock, glbErrorLog)
      Call EnableForm(Me, True)
      Exit Sub
       End If
   
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
      If Not glbDaily.DeleteConditionCom(COM_ID, IsOK, True, glbErrorLog) Then
       m_Condition3.COM_ID = -1
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      Call glbDatabaseMngr.UnLockTable(m_TableName, COM_ID, IsCanLock, glbErrorLog)
      Call EnableForm(Me, True)
      Exit Sub
      End If
   
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
      If Not glbDaily.DeleteConditionCom(COM_ID, IsOK, True, glbErrorLog) Then
       m_Condition4.COM_ID = -1
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      Call glbDatabaseMngr.UnLockTable(m_TableName, COM_ID, IsCanLock, glbErrorLog)
      Call EnableForm(Me, True)
      Exit Sub
      End If
   
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
             If Not glbDaily.DeleteConditionCom(COM_ID, IsOK, True, glbErrorLog) Then
             m_Condition5.COM_ID = -1
            glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
            Call glbDatabaseMngr.UnLockTable(m_TableName, COM_ID, IsCanLock, glbErrorLog)
            Call EnableForm(Me, True)
            Exit Sub
         End If
  End If
  
   Call QueryData(True)
   
   Call glbDatabaseMngr.UnLockTable(m_TableName, COM_ID, IsCanLock, glbErrorLog)
   Call EnableForm(Me, True)
End Sub

Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim IsCanLock As Boolean
Dim ID As Long
Dim OKClick As Boolean


      
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   COM_ID = Val(GridEX1.Value(1))
   OKClick = False
   
   If TabStrip1.SelectedItem.Index = 1 Then             'ไว้เมื่อแก้ไข
   frmAddEditComType2.COM_ID = COM_ID
  frmAddEditComType2.YEAR_ID = YEAR_ID
      frmAddEditComType2.HeaderText = MapText("แก้ไขค่า Commission ขาย")
         frmAddEditComType2.ShowMode = SHOW_EDIT
         Load frmAddEditComType2
            frmAddEditComType2.Show 1
            
            OKClick = frmAddEditComType2.OKClick
            
            Unload frmAddEditComType2
            Set frmAddEditComType2 = Nothing
            
            If OKClick Then
                     GridEX1.ItemCount = itemCountGrid1
                     GridEX1.Rebind
          End If
      
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      frmAddEditComType3.COM_ID = COM_ID
  frmAddEditComType3.YEAR_ID = YEAR_ID
  frmAddEditComType3.COMTYP = "02"
      frmAddEditComType3.HeaderText = MapText("แก้ไขค่า Commission ขาย")
         frmAddEditComType3.ShowMode = SHOW_EDIT
         Load frmAddEditComType3
            frmAddEditComType3.Show 1
            
            OKClick = frmAddEditComType3.OKClick
            
            Unload frmAddEditComType3
            Set frmAddEditComType3 = Nothing
            
            If OKClick Then
                     GridEX1.ItemCount = itemCountGrid2
                     GridEX1.Rebind
          End If
   
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
      frmAddEditComType3.COM_ID = COM_ID
  frmAddEditComType3.YEAR_ID = YEAR_ID
    frmAddEditComType3.COMTYP = "03"
      frmAddEditComType3.HeaderText = MapText("แก้ไขค่า Commission ขาย")
         frmAddEditComType3.ShowMode = SHOW_EDIT
         Load frmAddEditComType3
            frmAddEditComType3.Show 1
            
            OKClick = frmAddEditComType3.OKClick
            
            Unload frmAddEditComType3
            Set frmAddEditComType3 = Nothing
            
            If OKClick Then
                     GridEX1.ItemCount = itemCountGrid3
                     GridEX1.Rebind
          End If
   
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
         frmAddEditComType1.COM_ID = COM_ID
  frmAddEditComType1.YEAR_ID = YEAR_ID
      frmAddEditComType1.HeaderText = MapText("แก้ไขค่า Commission ขาย")
         frmAddEditComType1.ShowMode = SHOW_EDIT
         Load frmAddEditComType1
            frmAddEditComType1.Show 1
            
            OKClick = frmAddEditComType3.OKClick
            
            Unload frmAddEditComType1
            Set frmAddEditComType1 = Nothing
            
            If OKClick Then
                     GridEX1.ItemCount = itemCountGrid4
                     GridEX1.Rebind
          End If
   
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
         frmAddEditComType4.COM_ID = COM_ID
  frmAddEditComType4.YEAR_ID = YEAR_ID
      frmAddEditComType4.HeaderText = MapText("แก้ไขค่า Incentive")
         frmAddEditComType4.ShowMode = SHOW_EDIT
               Set frmAddEditComType4.ParentForm = Me
     frmAddEditComType4.itemCountGrid = itemCountGrid5
         Load frmAddEditComType4
            frmAddEditComType4.Show 1
            
            OKClick = frmAddEditComType4.OKClick
            
            Unload frmAddEditComType4
            Set frmAddEditComType4 = Nothing
            
            If OKClick Then
                     GridEX1.ItemCount = itemCountGrid5
                     GridEX1.Rebind
          End If
   End If
         

   
   If OKClick Then
      m_HasModify = True
   End If
   
       Call QueryData(True)
End Sub

Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub

'Private Sub cmdPictureAdd_Click()
'On Error Resume Next
'Dim strDescription As String
'
'   'edit the filter to support more image types
'   dlgAdd.Filter = "Picture Files (*.jpg, *.gif)|*.jpg;*.gif"
'   dlgAdd.DialogTitle = "Select Picture to Add to Database"
'   dlgAdd.ShowOpen
'   If dlgAdd.FileName = "" Then
'      Exit Sub
'   End If
'
'   m_HasModify = True
'End Sub

'Private Sub cmdSave_Click()
'Dim Result As Boolean
'   If Not SaveData Then
'      Exit Sub
'   End If
'
''   ShowMode = SHOW_EDIT      'คอมเม้นท์ไว้ เพราะตอน มันมีปัญหา
''   YEAR_ID = m_CommissYear.YEAR_ID
''   m_CommissYear.QueryFlag = 1
''   QueryData (True)
''   m_HasModify = False
'   OKClick = True
'   Unload Me
'
'End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      

      
      Call EnableForm(Me, False)
 '     Call LoadEmployee(uctlEmployeeLookup.MyCombo, m_Employees)
 '     Set uctlEmployeeLookup.MyCollection = m_Employees
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_CommissYear.QueryFlag = 1
         Call QueryData(True)
         Call TabStrip1_Click
      ElseIf ShowMode = SHOW_ADD Then
         m_CommissYear.QueryFlag = 0
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
   
   If cm1_Rs.State = adStateOpen Then
      cm1_Rs.Close
   End If
   Set cm1_Rs = Nothing
   
      If cm2_Rs.State = adStateOpen Then
      cm2_Rs.Close
   End If
   Set cm2_Rs = Nothing
   
      If cm3_Rs.State = adStateOpen Then
      cm3_Rs.Close
   End If
   Set cm3_Rs = Nothing
   
      If cm4_Rs.State = adStateOpen Then
      cm4_Rs.Close
   End If
   Set cm4_Rs = Nothing
   
         If cm5_Rs.State = adStateOpen Then
      cm5_Rs.Close
   End If
   Set cm5_Rs = Nothing
   
   Set m_CommissYear = Nothing
   Set m_Condition1 = Nothing
      Set m_Condition2 = Nothing
         Set m_Condition3 = Nothing
            Set m_Condition4 = Nothing
            Set m_Condition5 = Nothing
            
  ' Set m_Employees = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub InitGrid1()
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

   Set Col = GridEX1.Columns.Add '3
   Col.Width = 2100
   Col.Caption = MapText("เงื่อนไข")

   Set Col = GridEX1.Columns.Add '3
   Col.Width = 2100
   Col.Caption = MapText("ยอดขาย(%)")

   Set Col = GridEX1.Columns.Add '4
   Col.Width = 4425 + 3240
   Col.Caption = MapText("เปอร์เซ็นต์ที่ได้(%)")

End Sub


Private Sub InitGrid2()
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

'   Set Col = GridEX1.Columns.Add '2
'   Col.Width = 0
'   Col.Caption = "Real ID"

   Set Col = GridEX1.Columns.Add '3
   Col.Width = 2100
   Col.Caption = MapText("เงื่อนไข")

   Set Col = GridEX1.Columns.Add '3
   Col.Width = 2100
   Col.Caption = MapText("เครดิตภายใน(วัน)")

   Set Col = GridEX1.Columns.Add '4
   Col.Width = 4425 + 3240
   Col.Caption = MapText("คิดเป็น(%)")

End Sub

Private Sub InitGrid3()
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

   Set Col = GridEX1.Columns.Add '3
   Col.Width = 2100
   Col.Caption = MapText("เลขที่สินค้า")

   Set Col = GridEX1.Columns.Add '4
   Col.Width = 4425
   Col.Caption = MapText("ชื่อสินค้า")
   
   Set Col = GridEX1.Columns.Add '4
   Col.Width = 3240
   Col.Caption = MapText("คิดเป็น(%)")

End Sub


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
   Col.Caption = MapText("Pack")
   
 Set Col = GridEX1.Columns.Add '5
   Col.Width = 1400
   Col.Caption = MapText("ราคา")
   
Set Col = GridEX1.Columns.Add '6
   Col.Width = 2200
   Col.Caption = MapText("Incentive ต่อขวดและถุง")

Set Col = GridEX1.Columns.Add '6
   Col.Width = 1000
   Col.Caption = MapText("วัคซีน")

End Sub

Private Sub InitFormLayout()
   ppnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   ppnlHeader.Caption = Me.Caption
'
   Call InitNormalLabel(lblYearNum, MapText("เลขที่"))
   Call InitNormalLabel(lblFromDate, MapText("เริ่มใช้วันที่"))
      Call InitNormalLabel(lblToDate, MapText("ถึงวันที่"))
   
  Call txtYearNum.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)

   ppnlHeader.Font.Name = GLB_FONT
   ppnlHeader.Font.Bold = True
   ppnlHeader.Font.Size = 19


   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)

   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))

   
   Call InitGrid1
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.Name = GLB_FONT
   TabStrip1.Font.Size = 16
   
   TabStrip1.Tabs.Clear
   TabStrip1.Tabs.Add().Caption = MapText("1.ขาย")
   TabStrip1.Tabs.Add().Caption = MapText("2.เก็บเงิน - เครดิตธรรมดา")
      TabStrip1.Tabs.Add().Caption = MapText("3.เก็บเงิน - เครดิตสินค้า Commodity")
      TabStrip1.Tabs.Add().Caption = MapText("4.เก็บเงิน - ตั้งค่าสินค้า Commodity")
         TabStrip1.Tabs.Add().Caption = MapText("5.Incentive")
End Sub

'Private Sub cmdExit_Click()
'   If Not ConfirmExit(m_HasModify) Then
'      Exit Sub
'   End If
'
'   OKClick = False
'   Unload Me
'End Sub

Private Sub Form_Load()
      m_TableName = "CONDITION_COMMISSION"
   OKClick = False
      
   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
   Set cm1_Rs = New ADODB.Recordset
      Set cm2_Rs = New ADODB.Recordset
         Set cm3_Rs = New ADODB.Recordset
            Set cm4_Rs = New ADODB.Recordset
                  Set cm5_Rs = New ADODB.Recordset
   Set m_CommissYear = New CCommissYear
Set m_Condition1 = New CConditionCommission
Set m_Condition2 = New CConditionCommission
Set m_Condition3 = New CConditionCommission
Set m_Condition4 = New CConditionCommission
Set m_Condition5 = New CConditionCommission

      m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)

End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub

Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
   If TabStrip1.SelectedItem.Index = 5 Then
      RowBuffer.RowStyle = RowBuffer.Value(7)
   End If
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"


 If TabStrip1.SelectedItem.Index = 1 Then    ' หรือว่าต้องมี cm_Rs x5
   If cm1_Rs Is Nothing Then
      Exit Sub
   End If
   If cm1_Rs.State <> adStateOpen Then
      Exit Sub
   End If
   If cm1_Rs.EOF Then
      Exit Sub
   End If
      If RowIndex <= 0 Then
         Exit Sub
      End If
   Call cm1_Rs.Move(RowIndex - 1, adBookmarkFirst)
    Call m_Condition1.PopulateFromRS(2, cm1_Rs)   ' ใช้ป๊อปแบบที่ 2 เพราะจะได้ ตัวเลข NUM_ONE ออกมา
      Values(1) = m_Condition1.COM_ID                   ' values(1) เกี่ยวพันกับการลบ
      Values(2) = "<="                                               ' operatorToText(m_Condition1.OPERATOR)
      Values(3) = m_Condition1.NUM_ONE
      Values(4) = m_Condition1.SLM_PERCENT
      
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
            If cm2_Rs Is Nothing Then
      Exit Sub
   End If
   If cm2_Rs.State <> adStateOpen Then
      Exit Sub
   End If
   If cm2_Rs.EOF Then
      Exit Sub
   End If
      If RowIndex <= 0 Then
         Exit Sub
      End If
   Call cm2_Rs.Move(RowIndex - 1, adBookmarkFirst)
    Call m_Condition2.PopulateFromRS(2, cm2_Rs)   ' ใช้ป๊อปแบบที่ 2 เพราะจะได้ ตัวเลข NUM_ONE ออกมา
      Values(1) = m_Condition2.COM_ID      ' values(1) เกี่ยวพันกับการลบ
      Values(2) = "<="                                             ' operatorToText(m_Condition2.OPERATOR)
      Values(3) = m_Condition2.NUM_ONE
      Values(4) = m_Condition2.SLM_PERCENT
      
ElseIf TabStrip1.SelectedItem.Index = 3 Then
         If cm3_Rs Is Nothing Then
      Exit Sub
   End If
   If cm3_Rs.State <> adStateOpen Then
      Exit Sub
   End If
   If cm3_Rs.EOF Then
      Exit Sub
   End If
      If RowIndex <= 0 Then
         Exit Sub
      End If
   Call cm3_Rs.Move(RowIndex - 1, adBookmarkFirst)
    Call m_Condition3.PopulateFromRS(2, cm3_Rs)   ' ใช้ป๊อปแบบที่ 2 เพราะจะได้ ตัวเลข NUM_ONE ออกมา
      Values(1) = m_Condition3.COM_ID                    ' values(1) เกี่ยวพันกับการลบ
            Values(2) = "<="                                           ' operatorToText(m_Condition3.OPERATOR)
      Values(3) = m_Condition3.NUM_ONE
      Values(4) = m_Condition3.SLM_PERCENT
      
ElseIf TabStrip1.SelectedItem.Index = 4 Then
      If cm4_Rs Is Nothing Then
      Exit Sub
   End If
   If cm4_Rs.State <> adStateOpen Then
      Exit Sub
   End If
   If cm4_Rs.EOF Then
      Exit Sub
   End If
      If RowIndex <= 0 Then
         Exit Sub
      End If
   Call cm4_Rs.Move(RowIndex - 1, adBookmarkFirst)
    Call m_Condition4.PopulateFromRS(1, cm4_Rs)
      Values(1) = m_Condition4.COM_ID
      Values(2) = m_Condition4.STKCOD
      Values(3) = m_Condition4.STKDES
      Values(4) = m_Condition4.SLM_PERCENT
   
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
         If cm5_Rs Is Nothing Then
      Exit Sub
   End If
   If cm5_Rs.State <> adStateOpen Then
      Exit Sub
   End If
   If cm5_Rs.EOF Then
      Exit Sub
   End If
      If RowIndex <= 0 Then
         Exit Sub
      End If
   Call cm5_Rs.Move(RowIndex - 1, adBookmarkFirst)
    Call m_Condition5.PopulateFromRS(1, cm5_Rs)        ' ไปดูอีกที จะใช้ pop ไร เพราะมันติ๊กว่าเป็นวัคซีน
      Values(1) = m_Condition5.COM_ID
      Values(2) = m_Condition5.GROUP1
      Values(3) = m_Condition5.STKCOD
      Values(4) = m_Condition5.STKDES
       Values(5) = m_Condition5.NUM_ONE           ' Pack
        Values(6) = m_Condition5.NUM_TWO        ' Price
      Values(7) = m_Condition5.SLM_PERCENT   ' Incentive
      Values(8) = m_Condition5.OPERATOR   ' Incentive
      
   End If
      
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub


Private Sub TabStrip1_Click()
   If TabStrip1.SelectedItem.Index = 1 Then
      Call InitGrid1
      

      GridEX1.ItemCount = itemCountGrid1
      GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      Call InitGrid2
      
      GridEX1.ItemCount = itemCountGrid2
      GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
         Call InitGrid2
      
      GridEX1.ItemCount = itemCountGrid3
      GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
            Call InitGrid3
      
      GridEX1.ItemCount = itemCountGrid4
      GridEX1.Rebind
      
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
      Call InitGrid4
      
      GridEX1.ItemCount = itemCountGrid5
      GridEX1.Rebind
   
   End If
End Sub

Private Sub txtDoNo_Change()
   m_HasModify = True
End Sub

Private Sub txtDeliveryNo_Change()
   m_HasModify = True
End Sub

Private Sub txtSellBy_Change()
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

Private Sub txtTruckNo_Change()
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
Private Sub uctlDocumentDate_HasChange()
   m_HasModify = True
End Sub
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

Public Function operatorToText(O_case As String) As String
If O_case = "1" Then
      operatorToText = "<"
ElseIf O_case = "2" Then
      operatorToText = "<="
ElseIf O_case = "3" Then
      operatorToText = ">="
ElseIf O_case = "4" Then
      operatorToText = ">"
ElseIf O_case = "5" Then
      operatorToText = "="
End If
End Function




