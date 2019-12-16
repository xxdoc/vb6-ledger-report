VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmConditionCommission 
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
      Height          =   8655
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   12075
      _ExtentX        =   21299
      _ExtentY        =   15266
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboOrderType 
         Height          =   315
         Left            =   6450
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1470
         Width           =   2955
      End
      Begin VB.ComboBox cboOrderBy 
         Height          =   315
         Left            =   1905
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1470
         Width           =   2985
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   12
         Top             =   0
         Width           =   12045
         _ExtentX        =   21246
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjLedgerReport.uctlTextBox txtYear 
         Height          =   435
         Left            =   1905
         TabIndex        =   0
         Top             =   960
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   767
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   5295
         Left            =   240
         TabIndex        =   5
         Top             =   2160
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
         Column(1)       =   "frmConditionCommissionä.frx":0000
         Column(2)       =   "frmConditionCommissionä.frx":00C8
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmConditionCommissionä.frx":016C
         FormatStyle(2)  =   "frmConditionCommissionä.frx":02C8
         FormatStyle(3)  =   "frmConditionCommissionä.frx":0378
         FormatStyle(4)  =   "frmConditionCommissionä.frx":042C
         FormatStyle(5)  =   "frmConditionCommissionä.frx":0504
         ImageCount      =   0
         PrinterProperties=   "frmConditionCommissionä.frx":05BC
      End
      Begin VB.Label lblYear 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   15
         Top             =   960
         Width           =   1755
      End
      Begin VB.Label lblOrderType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4975
         TabIndex        =   14
         Top             =   1470
         Width           =   1365
      End
      Begin VB.Label lblOrderBy 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   13
         Top             =   1470
         Width           =   1755
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   10110
         TabIndex        =   3
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   525
         Left            =   10110
         TabIndex        =   4
         Top             =   1530
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
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
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10095
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmConditionCommission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'frmCommission2
Option Explicit
Private m_HasActivate As Boolean
Private m_YearSeq As CCommissYear
Private m_TempYearSeq As CCommissYear
Private m_Rs As ADODB.Recordset
Private m_TableName As String

Public OKClick As Boolean

Private Sub cmdPasswd_Click()

End Sub

Private Sub cmdAdd_Click()
Dim ItemCount As Long
Dim OKClick As Boolean


   frmAddEditConditionCommiss.HeaderText = MapText("‡æ‘Ë¡‡ß◊ËÕπ‰¢°“√®Ë“¬§Õ¡¡‘™™—Ëπ·µË≈–ª’")
   frmAddEditConditionCommiss.ShowMode = SHOW_ADD
   Load frmAddEditConditionCommiss
  frmAddEditConditionCommiss.Show 1
   
   OKClick = frmAddEditConditionCommiss.OKClick
   
   Unload frmAddEditConditionCommiss
   Set frmAddEditConditionCommiss = Nothing
   
   If OKClick Then
      Call QueryData(True)
   End If
End Sub

Private Sub cmdClear_Click()
   txtYear.Text = ""
   cboOrderBy.ListIndex = -1
   cboOrderType.ListIndex = -1
End Sub

Private Sub cmdDelete_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim YEAR_ID As Long


   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
  YEAR_ID = GridEX1.Value(1)

   If Not ConfirmDelete(GridEX1.Value(2)) Then
      Call glbDatabaseMngr.UnLockTable(m_TableName, YEAR_ID, IsCanLock, glbErrorLog)
      Exit Sub
   End If

   Call EnableForm(Me, False)
   If Not glbDaily.DeleteCommissYear(YEAR_ID, IsOK, True, glbErrorLog) Then
       m_YearSeq.YEAR_ID = -1
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      Call glbDatabaseMngr.UnLockTable(m_TableName, YEAR_ID, IsCanLock, glbErrorLog)
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
 '   ≈∫∑’ËÕ¬ŸË„π COM_ID
      If Not glbDaily.DeleteConditionComFromYear(YEAR_ID, IsOK, True, glbErrorLog) Then
       m_YearSeq.YEAR_ID = -1
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      Call glbDatabaseMngr.UnLockTable(m_TableName, YEAR_ID, IsCanLock, glbErrorLog)
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call QueryData(True)
   
   Call glbDatabaseMngr.UnLockTable(m_TableName, YEAR_ID, IsCanLock, glbErrorLog)
   Call EnableForm(Me, True)
End Sub

Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim YEAR_ID As Long
Dim OKClick As Boolean

      
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

  YEAR_ID = Val(GridEX1.Value(1))
 '  Call glbDatabaseMngr.LockTable(m_TableName, YEAR_ID, IsCanLock, glbErrorLog)
   
   frmAddEditConditionCommiss.YEAR_ID = YEAR_ID
   frmAddEditConditionCommiss.HeaderText = MapText("·°È‰¢‡ß◊ËÕπ‰¢°“√®Ë“¬§Õ¡¡‘™™—Ëπ√“¬ª’")
   frmAddEditConditionCommiss.ShowMode = SHOW_EDIT
   Load frmAddEditConditionCommiss
   frmAddEditConditionCommiss.Show 1
   
   OKClick = frmAddEditConditionCommiss.OKClick
   
   Unload frmAddEditConditionCommiss
   Set frmAddEditConditionCommiss = Nothing
               
   If OKClick Then
      Call QueryData(True)
   End If
   Call glbDatabaseMngr.UnLockTable(m_TableName, YEAR_ID, IsCanLock, glbErrorLog)

End Sub

Private Sub cmdOK_Click()
   OKClick = True
   Unload Me
End Sub

Private Sub cmdSearch_Click()
   Call QueryData(True)
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      
    Call InitYearComOrderBy(cboOrderBy)
    Call InitOrderType(cboOrderType)
      
      Call QueryData(True)
   End If
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim Temp As Long

   If Flag Then
      Call EnableForm(Me, False)
      
'      If Not VerifyAccessRight("PIG_WEEK_QUERY") Then
'         Call EnableForm(Me, True)
'         Exit Sub
'      End If
      
       m_YearSeq.YEARNUM = txtYear.Text
       m_YearSeq.OrderBy = cboOrderBy.ItemData(Minus2Zero(cboOrderBy.ListIndex))
       m_YearSeq.OrderType = cboOrderType.ItemData(Minus2Zero(cboOrderType.ListIndex))
       If Not glbDaily.QueryCommissYear(m_YearSeq, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
       End If
   End If
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   GridEX1.ItemCount = ItemCount
   GridEX1.Rebind
   
   Call EnableForm(Me, True)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.Name
      glbErrorLog.ShowUserError
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 116 Then
      Call cmdSearch_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 115 Then
      Call cmdClear_Click
      KeyCode = 0
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
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.Add '2
   Col.Width = 5000
   Col.Caption = MapText("‡≈¢∑’Ë")
      
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 2500
   Col.Caption = MapText("«—π∑’Ë‡√‘Ë¡µÈπ")

   Set Col = GridEX1.Columns.Add '3
   Col.Width = 2500
   Col.Caption = MapText("«—π∑’Ë ‘Èπ ÿ¥")
   
      Set Col = GridEX1.Columns.Add '3
   Col.Width = 1900
   Col.Caption = MapText("‡§√¥‘µ Incentive («—π)")
   
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 2500
   Col.Caption = MapText("°“√®—¥°“√ ‘π§È“")
   
    Set Col = GridEX1.Columns.Add '3
   Col.Width = 2500
   Col.Caption = MapText("°≈ÿË¡ ‘π§È“")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = MapText("‡ß◊ËÕπ‰¢°“√®Ë“¬§Õ¡¡‘™™—Ëπ")
   pnlHeader.Caption = Me.Caption
   
   Call InitGrid
   
   Call InitNormalLabel(lblYear, MapText("‡≈¢∑’Ë"))
   Call InitNormalLabel(lblOrderBy, MapText("‡√’¬ßµ“¡"))
   Call InitNormalLabel(lblOrderType, MapText("‡√’¬ß®“°"))
   
   Call txtYear.SetTextLenType(TEXT_STRING, glbSetting.NAME_LEN)
   
   Call InitCombo(cboOrderBy)
   Call InitCombo(cboOrderType)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSearch.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdClear.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("¬°‡≈‘° (ESC)"))
   Call InitMainButton(cmdOK, MapText("µ°≈ß (F2)"))
   Call InitMainButton(cmdAdd, MapText("‡æ‘Ë¡ (F7)"))
   Call InitMainButton(cmdEdit, MapText("·°È‰¢ (F3)"))
   Call InitMainButton(cmdDelete, MapText("≈∫ (F6)"))
   Call InitMainButton(cmdSearch, MapText("§ÈπÀ“ (F5)"))
   Call InitMainButton(cmdClear, MapText("‡§≈’¬√Ï (F4)"))
End Sub

Private Sub cmdExit_Click()
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   m_TableName = "COMMISSION_YEAR"
   m_HasActivate = False
   Set m_YearSeq = New CCommissYear
   Set m_TempYearSeq = New CCommissYear
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

Private Sub GridEX1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim oMenu As cPopupMenu
Dim lMenuChosen As Long
Dim TempID1 As Long
Dim TempID2 As Long
Dim Ms As CCommissYear
Dim IsOK As Boolean
Dim OKClick As Boolean
Dim YEAR_ID As Long

   If GridEX1.ItemCount <= 0 Then
         Exit Sub
   End If
   
     YEAR_ID = GridEX1.Value(1)
   
   If Button = 2 Then
      Set oMenu = New cPopupMenu
      lMenuChosen = oMenu.Popup("COPY")
      If lMenuChosen = 0 Then
         Exit Sub
      End If
      Set oMenu = Nothing
   Else
      Exit Sub
   End If
   
   Call EnableForm(Me, False)
   If lMenuChosen = 1 Then
      Set Ms = New CCommissYear
      Ms.YEAR_ID = YEAR_ID
      Call glbDaily.CopyCommissYear(Ms, IsOK, True, glbErrorLog)
      Call QueryData(True)
      Set Ms = Nothing
   End If
   
   Call EnableForm(Me, True)
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
   Call m_TempYearSeq.PopulateFromRS(1, m_Rs)
   
   Values(1) = m_TempYearSeq.YEAR_ID
   Values(2) = m_TempYearSeq.YEARNUM
   Values(3) = m_TempYearSeq.FROM_DATE
   Values(4) = m_TempYearSeq.TO_DATE
   Values(5) = m_TempYearSeq.INCEN_CR
   Values(6) = m_TempYearSeq.GOODS_MASTER_NAME
   Values(7) = m_TempYearSeq.GOODS_GROUP_NAME
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
   RowBuffer.RowStyle = RowBuffer.Value(2)
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

