VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frm0AreaCommission 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   8535
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   15055
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   7
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   6300
         Left            =   520
         TabIndex        =   0
         Top             =   1200
         Width           =   10600
         _ExtentX        =   18706
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
         Column(1)       =   "frm0AreaCommissionä.frx":0000
         Column(2)       =   "frm0AreaCommissionä.frx":00C8
         FormatStylesCount=   5
         FormatStyle(1)  =   "frm0AreaCommissionä.frx":016C
         FormatStyle(2)  =   "frm0AreaCommissionä.frx":02C8
         FormatStyle(3)  =   "frm0AreaCommissionä.frx":0378
         FormatStyle(4)  =   "frm0AreaCommissionä.frx":042C
         FormatStyle(5)  =   "frm0AreaCommissionä.frx":0504
         ImageCount      =   0
         PrinterProperties=   "frm0AreaCommissionä.frx":05BC
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   3
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
         TabIndex        =   1
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
         TabIndex        =   2
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
         TabIndex        =   5
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
         TabIndex        =   4
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
Attribute VB_Name = "frm0AreaCommission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'frmCommission2
Option Explicit
Private m_HasActivate As Boolean
Private masterArea As CCommissMasterArea
Private tempMasterArea As CCommissMasterArea
Private m_Rs As ADODB.Recordset
Private m_TableName As String

Public OKClick As Boolean

Private Sub cmdPasswd_Click()

End Sub

Private Sub cmdAdd_Click()
Dim ItemCount As Long
Dim OKClick As Boolean


  frmAddEditMasterArea.HeaderText = MapText("‡æ‘Ë¡‡¢µ°“√¢“¬")
  frmAddEditMasterArea.ShowMode = SHOW_ADD
  Load frmAddEditMasterArea
  frmAddEditMasterArea.Show 1
   
   OKClick = frmAddEditMasterArea.OKClick
   
   Unload frmAddEditMasterArea
   Set frmAddEditMasterArea = Nothing
   
   If OKClick Then
      Call QueryData(True)
   End If
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
   If Not glbDaily.DeleteCommissArea(MASTER_AREA_ID, IsOK, True, glbErrorLog) Then
       masterArea.MASTER_AREA_ID = -1
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      Call glbDatabaseMngr.UnLockTable(m_TableName, MASTER_AREA_ID, IsCanLock, glbErrorLog)
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
 '   ≈∫∑’ËÕ¬ŸË„π COM_ID
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
   
   frmAddEditAreaCom.MASTER_AREA_ID = MASTER_AREA_ID
   frmAddEditAreaCom.HeaderText = MapText("·°È‰¢¢ÈÕ¡Ÿ≈≈Ÿ°§È“„π‡¢µ°“√¢“¬")
'   frmAddEditAreaCom.ShowMode = SHOW_EDIT
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
   OKClick = True
   Unload Me
End Sub

Private Sub cmdSearch_Click()
   Call QueryData(True)
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      
      Call QueryData(True)
   End If
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim Temp As Long

   If Flag Then
      Call EnableForm(Me, False)

      If Not glbDaily.QueryCommissArea(masterArea, m_Rs, ItemCount, IsOK, glbErrorLog) Then
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
   Col.Caption = MapText("√À— ‡¢µ")
      
   Set Col = GridEX1.Columns.Add '2
   Col.Width = 2500
   Col.Caption = MapText("‡¢µ°“√¢“¬")

   GridEX1.ItemCount = 0
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = MapText("µ—Èß§Ë“‡¢µ°“√¢“¬")
   pnlHeader.Caption = Me.Caption
   
   Call InitGrid
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("¬°‡≈‘° (ESC)"))
   Call InitMainButton(cmdOK, MapText("µ°≈ß (F2)"))
   Call InitMainButton(cmdAdd, MapText("‡æ‘Ë¡ (F7)"))
   Call InitMainButton(cmdEdit, MapText("·°È‰¢ (F3)"))
   Call InitMainButton(cmdDelete, MapText("≈∫ (F6)"))

End Sub

Private Sub cmdExit_Click()
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   m_TableName = "COMMISSION_MASTER_AREA"
   OKClick = False
      
   m_HasActivate = False
   Set masterArea = New CCommissMasterArea
   Set tempMasterArea = New CCommissMasterArea
   Set m_Rs = New ADODB.Recordset
   m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   Debug.Print ColIndex & " " & NewColWidth
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



