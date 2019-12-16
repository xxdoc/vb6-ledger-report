VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmAddEditComBudget 
   BackColor       =   &H80000000&
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11910
   Icon            =   "frmAddEditComBudget.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   8535
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   15055
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboManagerName 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   900
         Width           =   2715
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   9
         Top             =   -120
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   5295
         Left            =   180
         TabIndex        =   7
         Top             =   2400
         Width           =   11505
         _ExtentX        =   20294
         _ExtentY        =   9340
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
         Column(1)       =   "frmAddEditComBudget.frx":27A2
         Column(2)       =   "frmAddEditComBudget.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditComBudget.frx":290E
         FormatStyle(2)  =   "frmAddEditComBudget.frx":2A6A
         FormatStyle(3)  =   "frmAddEditComBudget.frx":2B1A
         FormatStyle(4)  =   "frmAddEditComBudget.frx":2BCE
         FormatStyle(5)  =   "frmAddEditComBudget.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditComBudget.frx":2D5E
      End
      Begin prjLedgerReport.uctlDate uctlFromDate 
         Height          =   375
         Left            =   1680
         TabIndex        =   11
         Top             =   1360
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
      End
      Begin prjLedgerReport.uctlDate uctlToDate 
         Height          =   375
         Left            =   1680
         TabIndex        =   12
         Top             =   1840
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   661
      End
      Begin VB.Label lblToDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   480
         TabIndex        =   14
         Top             =   1840
         Width           =   1155
      End
      Begin VB.Label lblFromDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   480
         TabIndex        =   13
         Top             =   1360
         Width           =   1155
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   525
         Left            =   9960
         TabIndex        =   1
         Top             =   1650
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   9960
         TabIndex        =   0
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditComBudget.frx":2F36
         ButtonStyle     =   3
      End
      Begin VB.Label lblManagerName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   0
         TabIndex        =   10
         Top             =   900
         Width           =   1575
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   4
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditComBudget.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   150
         TabIndex        =   2
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditComBudget.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1770
         TabIndex        =   3
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10095
         TabIndex        =   6
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8445
         TabIndex        =   5
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditComBudget.frx":3884
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditComBudget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private combudget As CCombudget
Private TempComBudget As CCombudget
Private m_Rs As ADODB.Recordset
Private m_TableName As String
Public ShowMode As SHOW_MODE_TYPE

Private m_ManagerName As Collection

Public HeaderText As String

Public OKClick As Boolean
Public ID As Long
Public SubID As Long
Private Sub cmdAdd_Click()
Dim ItemCount As Long
Dim OKClick As Boolean

'   If Not VerifyTextControl(lblGroupCode, txtGroupCode, False) Then
'      Exit Sub
'   End If
'
'   If Not VerifyTextControl(lblGroupName, txtGroupName, False) Then
'      Exit Sub
'   End If

   frmAddEditSubComBudget.HeaderText = MapText("เพิ่มกลุ่ม")
   frmAddEditSubComBudget.ShowMode = SHOW_ADD
   Set frmAddEditSubComBudget.TempAddEditDataCollection = combudget.ImportExportItems
   Load frmAddEditSubComBudget
   frmAddEditSubComBudget.Show 1
   OKClick = frmAddEditSubComBudget.OKClick
   
   Unload frmAddEditSubComBudget
   Set frmAddEditSubComBudget = Nothing

   If OKClick Then
   GridEX1.ItemCount = CountItem(combudget.ImportExportItems)
   GridEX1.Rebind
   End If
End Sub
Private Sub cmdDelete_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim ID1 As Long
Dim ID2 As Long

   If Not cmdDelete.Enabled Then
      Exit Sub
   End If

   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   ID1 = GridEX1.Value(1)
   ID2 = GridEX1.Value(2)
   If ID1 <= 0 Then
         combudget.ImportExportItems.Remove (ID2)
   Else
         combudget.ImportExportItems.ITEM(ID2).Flag = "D"
   End If
   
   GridEX1.ItemCount = CountItem(combudget.ImportExportItems)
   GridEX1.Rebind
   
End Sub
Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim OKClick As Boolean
Dim EnpAddress As CSupCombudget
Set EnpAddress = New CSupCombudget
      
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   SubID = Val(GridEX1.Value(2))
   frmAddEditSubComBudget.SubID = SubID
   frmAddEditSubComBudget.HeaderText = MapText("แก้ไขกลุ่ม")
   frmAddEditSubComBudget.ShowMode = SHOW_EDIT
   Set frmAddEditSubComBudget.TempAddEditDataCollection = combudget.ImportExportItems
   Load frmAddEditSubComBudget
   frmAddEditSubComBudget.Show 1

   OKClick = frmAddEditSubComBudget.OKClick

   Unload frmAddEditSubComBudget
   Set frmAddEditSubComBudget = Nothing

   If OKClick Then
      GridEX1.ItemCount = CountItem(combudget.ImportExportItems)
      GridEX1.Rebind
   End If
End Sub
Private Function SaveData() As Boolean
Dim ROW As Long
Dim Ge As CCombudget
Dim Gse As CSupCombudget
Dim ChequeDate  As Date
Dim CancelReason As String

Set Ge = New CCombudget
Set Gse = New CSupCombudget

'   If Not VerifyTextControl(lblGroupCode, txtGroupCode, False) Then
'      Exit Function
'   End If
'
'   If Not VerifyTextControl(lblGroupName, txtGroupName, False) Then
'      Exit Function
'   End If

'   If Not CheckUniqueNs(GROUP_ENTERPRISE_CODE, txtGroupCode.Text, ID) Then
'      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtGroupCode.Text & " " & MapText("อยู่ในระบบแล้ว")           ' เป็น  combo ชื่อ sale แทน
'      glbErrorLog.ShowUserError
'      Exit Function
'   End If

   If ShowMode = SHOW_ADD Then
      If m_HasModify Then
         Ge.ShowMode = SHOW_ADD
         Ge.MANAGER_ID = cboManagerName.ItemData(Minus2Zero(cboManagerName.ListIndex))
         Ge.FROM_DATE = uctlFromDate.ShowDate
         Ge.TO_DATE = uctlToDate.ShowDate
         Call Ge.AddEditData
         ID = Ge.Temp_ID + 1
         Set Ge = Nothing
      End If
   Else
      If m_HasModify Then
         Ge.ShowMode = SHOW_EDIT
         Ge.MANAGER_ID = cboManagerName.ItemData(Minus2Zero(cboManagerName.ListIndex))
         Ge.FROM_DATE = uctlFromDate.ShowDate
         Ge.TO_DATE = uctlToDate.ShowDate
         Ge.COM_BUDGET_ID = ID
         Call Ge.AddEditData
         Set Ge = Nothing
      End If
   End If

   For Each Gse In combudget.ImportExportItems
      If Gse.Flag = "A" Then
         Gse.COM_BUDGET_ID = ID
         Gse.ShowMode = SHOW_ADD
         Call Gse.AddEditData
      End If
      If Gse.Flag = "E" Then
         Gse.COM_BUDGET_ID = ID
         Gse.ShowMode = SHOW_EDIT
         Call Gse.AddEditData
      End If
      If Gse.Flag = "D" Then
         Call Gse.DeleteData
      End If
   Next Gse
   Set Gse = Nothing

   Call EnableForm(Me, True)
   SaveData = True
End Function
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
Set combudget = New CCombudget

   If Flag Then
      Call EnableForm(Me, False)
      
      combudget.COM_BUDGET_ID = ID
      combudget.QueryFlag = 1
'      If Not glbDaily.QueryComBudget(combudget, m_Rs, ItemCount, IsOK, glbErrorLog) Then
'         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
'         Call EnableForm(Me, True)
'         Exit Sub
'      End If
   End If
      
   If m_Rs.RecordCount > 0 Then
      Call combudget.PopulateFromRS(1, m_Rs)
      cboManagerName.ListIndex = IDToListIndex(cboManagerName, combudget.MANAGER_ID)
      uctlFromDate.ShowDate = combudget.FROM_DATE
      uctlToDate.ShowDate = combudget.TO_DATE
   End If
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   If ItemCount > 0 Then
      GridEX1.ItemCount = CountItem(combudget.ImportExportItems)
      GridEX1.Rebind
   End If
   
   Call EnableForm(Me, True)
End Sub
Private Sub Form_Activate()
  
   Call LoadSale(cboManagerName, m_ManagerName)
 
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh

      Call EnableForm(Me, False)
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         ID = -1
      End If
      
      m_HasModify = False
      
      Call EnableForm(Me, True)
   End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.Name
      glbErrorLog.ShowUserError
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
   Col.Width = 0
   Col.Caption = "ID"
   
   Set Col = GridEX1.Columns.Add '2
   Col.Width = 2000
   Col.Caption = MapText("ชื่อพนักงานขาย")

   Set Col = GridEX1.Columns.Add '3
   Col.Width = 2000
   Col.Caption = "เขตการขาย"
   
   Set Col = GridEX1.Columns.Add '4
   Col.Width = 2000
   Col.Caption = "งบประมาณ"
   
   GridEX1.ItemCount = 0
End Sub
Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = MapText("การจัดกลุ่ม")
   pnlHeader.Caption = Me.Caption
   
   Call InitGrid
   
   Call InitNormalLabel(lblManagerName, MapText("ชื่อ Manager"))
   Call InitNormalLabel(lblFromDate, MapText("วันที่เริ่มใช้"))
   Call InitNormalLabel(lblToDate, MapText("ถึงวันที่"))
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSearch.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdClear.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdSearch, MapText("ค้นหา (F5)"))
   Call InitMainButton(cmdClear, MapText("เคลียร์ (F4)"))
End Sub
Private Sub cmdExit_Click()
   OKClick = False
   Unload Me
End Sub
Private Sub Form_Load()
      
   m_TableName = "COM_SUB_BUDGET"
   
   Set m_Rs = New ADODB.Recordset
   Set combudget = New CCombudget
   Set TempComBudget = New CCombudget
   Set m_ManagerName = New Collection

   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set combudget = Nothing
   Set TempComBudget = Nothing
   
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
End Sub
Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   Debug.Print ColIndex & " " & NewColWidth
End Sub
Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
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
Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
   RowBuffer.RowStyle = RowBuffer.Value(2)
End Sub
Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"
   
   If combudget.ImportExportItems Is Nothing Then
      Exit Sub
   End If
   If RowIndex <= 0 Then
      Exit Sub
   End If
   If combudget.ImportExportItems.Count <= 0 Then
      Exit Sub
   End If
   Dim UD As CSupCombudget
   Set UD = GetItem(combudget.ImportExportItems, RowIndex, RealIndex)
   If UD Is Nothing Then
      Exit Sub
   End If
   Values(1) = UD.COM_SUB_BUDGET_ID
   Values(2) = UD.MANAGER_ID
   Values(3) = UD.MASTER_AREA_ID
   Values(4) = UD.BUDGET
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
   cmdExit.Top = ScaleHeight - 580
   cmdExit.Left = ScaleWidth - cmdExit.Width - 50
   cmdOK.Left = cmdExit.Left - cmdOK.Width - 50
End Sub
Private Sub uctlFromDate_HasChange()
   m_HasModify = True
End Sub
Private Sub uctlToDate_HasChange()
   m_HasModify = True
End Sub
Private Sub cboManagerName_Click()
   m_HasModify = True
End Sub

