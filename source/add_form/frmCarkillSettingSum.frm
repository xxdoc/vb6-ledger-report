VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmCarkillSettingSum 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14190
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9105
   ScaleWidth      =   14190
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   10535
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   14555
      _ExtentX        =   25665
      _ExtentY        =   18574
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   6
         Top             =   0
         Width           =   14525
         _ExtentX        =   25612
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   6300
         Left            =   520
         TabIndex        =   0
         Top             =   1200
         Width           =   13100
         _ExtentX        =   23098
         _ExtentY        =   11113
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
         Column(1)       =   "frmCarkillSettingSum.frx":0000
         Column(2)       =   "frmCarkillSettingSum.frx":00C8
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmCarkillSettingSum.frx":016C
         FormatStyle(2)  =   "frmCarkillSettingSum.frx":02C8
         FormatStyle(3)  =   "frmCarkillSettingSum.frx":0378
         FormatStyle(4)  =   "frmCarkillSettingSum.frx":042C
         FormatStyle(5)  =   "frmCarkillSettingSum.frx":0504
         ImageCount      =   0
         PrinterProperties=   "frmCarkillSettingSum.frx":05BC
      End
      Begin VB.Label lblcomboName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   15
         Left            =   1200
         TabIndex        =   7
         Top             =   1000
         Width           =   1755
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   3
         Top             =   8130
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
         TabIndex        =   1
         Top             =   8130
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1770
         TabIndex        =   2
         Top             =   8130
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   12095
         TabIndex        =   4
         Top             =   8130
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmCarkillSettingSum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'frmCommission2
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
'Private ComboSup As CComboSubGroup
Private ItemXlsSum As CXlsCarkillSum
'Private tempItemXlsSum As CComboSubGroupDe
Private Rs As ADODB.Recordset
Private m_Rs As ADODB.Recordset
Private m_TableName As String
Public OKClick As Boolean
Public ShowMode As SHOW_MODE_TYPE
Public ID As Long
Public HeaderText As String

Private Sub cmdAdd_Click()
Dim ItemCount As Long
Dim OKClick As Boolean

   Call EnableForm(Me, False)
   
  frmAddEditCarkillSum.ShowMode = SHOW_ADD
  frmAddEditCarkillSum.HeaderText = MapText("เพิ่มบรรทัดรวม")
'  frmAddEditCarkillSum.COMBO_SUB_ID = ID

  Load frmAddEditCarkillSum
  frmAddEditCarkillSum.Show 1

  OKClick = frmAddEditCarkillSum.OKClick

  Unload frmAddEditCarkillSum
  Set frmAddEditCarkillSum = Nothing

  If OKClick Then
     Call QueryData(True)
  End If
End Sub


Private Sub cmdDelete_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
'Dim COMBO_DETAIL_ID As Long

   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   ItemXlsSum.XLS_SUM_ID = GridEX1.Value(1)

   If Not ConfirmDelete(GridEX1.Value(2)) Then
      Exit Sub
   End If

   Call EnableForm(Me, False)
   If Not glbDaily.DeleteCarkillSum(ItemXlsSum, IsOK, True, glbErrorLog) Then
       ItemXlsSum.XLS_SUM_ID = -1
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      Call EnableForm(Me, True)
      Exit Sub
   End If

' '   ลบที่อยู่ใน COM_ID
'      If Not glbDaily.DeleteCusAreaComFromMaster(MASTER_AREA_ID, IsOK, True, glbErrorLog) Then
'       ItemXlsSum.MASTER_AREA_ID = -1
'      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
'      Call glbDatabaseMngr.UnLockTable(m_TableName, MASTER_AREA_ID, IsCanLock, glbErrorLog)
'      Call EnableForm(Me, True)
'      Exit Sub
'   End If
   Call QueryData(True)

   Call EnableForm(Me, True)
End Sub

Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim DETAIL_ID As Long
Dim OKClick As Boolean

   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

  DETAIL_ID = Val(GridEX1.Value(1))
 ' Call glbDatabaseMngr.LockTable(m_TableName, MASTER_AREA_ID, IsCanLock, glbErrorLog)

  frmAddEditCarkillSum.ID = DETAIL_ID
'  frmAddEditCarkillSum.COMBO_SUB_ID = ID
  frmAddEditCarkillSum.HeaderText = MapText("แก้ไขบรรทัดรวม")
  frmAddEditCarkillSum.ShowMode = SHOW_EDIT
  Load frmAddEditCarkillSum
  frmAddEditCarkillSum.Show 1

   OKClick = frmAddEditCarkillSum.OKClick

   Unload frmAddEditCarkillSum
   Set frmAddEditCarkillSum = Nothing

   If OKClick Then
      Call QueryData(True)
   End If

End Sub

Private Sub cmdOK_Click()
      OKClick = True
      Unload Me
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      
      If ShowMode = SHOW_ADD Then
         ID = -1
      Else
         Call QueryData(True)
      End If
      
   End If
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim itemCountMas As Long
Dim ItemCount As Long
Dim Temp As Long

      Call EnableForm(Me, False)
      
'      ItemXlsSum.COMBO_SUB_ID = ID
      ItemXlsSum.XLS_SUM_ID = -1

      If Not glbDaily.QueryCarkillSum(ItemXlsSum, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
      
     Call InitGrid

   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If

   GridEX1.ItemCount = ItemCount
   GridEX1.Rebind

   Call EnableForm(Me, True)
End Sub
'
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
'      Call cmdPrint_Click
'      KeyCode = 0
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
   Col.Width = 0
   Col.Caption = MapText("ID")

   Set Col = GridEX1.Columns.Add '2
   Col.Width = 1600
   Col.Caption = MapText("แถวที่")
   
   Set Col = GridEX1.Columns.Add '2
   Col.Width = 2100
   Col.Caption = MapText("ตัวบวก (1)")
   
   Set Col = GridEX1.Columns.Add '2
   Col.Width = 2100
   Col.Caption = MapText("ตัวบวก (2)")
   
   Set Col = GridEX1.Columns.Add '2
   Col.Width = 2100
   Col.Caption = MapText("ตัวบวก (3)")
   
   Set Col = GridEX1.Columns.Add '2
   Col.Width = 2100
   Col.Caption = MapText("ตัวบวก (4)")
   
   Set Col = GridEX1.Columns.Add '2
   Col.Width = 2100
   Col.Caption = MapText("ตัวบวก (5)")
   
   Set Col = GridEX1.Columns.Add '2
   Col.Width = 1100
   Col.Caption = MapText("รวมแนวนอน")
   
      GridEX1.ItemCount = 0
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   HeaderText = "ตั้งค่าสูตรบรรทัดรวม"

   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption

   Me.Picture = LoadPicture(glbParameterObj.MainPicture)

   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19

  ' cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)

  ' Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))

End Sub

Private Sub Form_Load()
   OKClick = False

    m_HasActivate = False
    Set ItemXlsSum = New CXlsCarkillSum
   Set Rs = New ADODB.Recordset
   Set m_Rs = New ADODB.Recordset

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
   Call ItemXlsSum.PopulateFromRS(1, m_Rs)

   Values(1) = ItemXlsSum.XLS_SUM_ID
   Values(2) = ItemXlsSum.SUM_ROW
   Values(3) = "(" & ItemXlsSum.OPERATOR_1 & ")" & column2txt(ItemXlsSum.P_COLUMN_1) & "(" & ItemXlsSum.P_ROW_1 & ")"
   If ItemXlsSum.OPERATOR_2 <> "" Then
      Values(4) = "(" & ItemXlsSum.OPERATOR_2 & ")" & column2txt(ItemXlsSum.P_COLUMN_2) & "(" & ItemXlsSum.P_ROW_2 & ")"
   End If
   If ItemXlsSum.OPERATOR_3 <> "" Then
      Values(5) = "(" & ItemXlsSum.OPERATOR_3 & ")" & column2txt(ItemXlsSum.P_COLUMN_3) & "(" & ItemXlsSum.P_ROW_3 & ")"
   End If
   If ItemXlsSum.OPERATOR_4 <> "" Then
      Values(6) = "(" & ItemXlsSum.OPERATOR_4 & ")" & column2txt(ItemXlsSum.P_COLUMN_4) & "(" & ItemXlsSum.P_ROW_4 & ")"
   End If
   If ItemXlsSum.OPERATOR_5 <> "" Then
      Values(7) = "(" & ItemXlsSum.OPERATOR_5 & ")" & column2txt(ItemXlsSum.P_COLUMN_5) & "(" & ItemXlsSum.P_ROW_5 & ")"
   End If
   Values(8) = ItemXlsSum.HORIZONTAL_FLAG
   
Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub txtcomboName_Change()
   m_HasModify = True
End Sub

Private Function column2txt(txtIn As String) As String
   If txtIn = "-1" Then
      column2txt = " (c.ก่อนหน้า) "
   ElseIf txtIn = "0" Then
      column2txt = " (c.ปัจจุบัน) "
   ElseIf txtIn = "+1" Then
      column2txt = " (c.ถัดไป) "
   Else
      column2txt = " (" & txtIn & ") "
   End If
End Function
