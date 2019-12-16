VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmCarkillSetFW 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8535
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   8535
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   10535
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   8555
      _ExtentX        =   15081
      _ExtentY        =   18574
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   6
         Top             =   0
         Width           =   8525
         _ExtentX        =   15028
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   5800
         Left            =   520
         TabIndex        =   0
         Top             =   1200
         Width           =   7100
         _ExtentX        =   12515
         _ExtentY        =   10239
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
         Column(1)       =   "frmCarkillSetFW.frx":0000
         Column(2)       =   "frmCarkillSetFW.frx":00C8
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmCarkillSetFW.frx":016C
         FormatStyle(2)  =   "frmCarkillSetFW.frx":02C8
         FormatStyle(3)  =   "frmCarkillSetFW.frx":0378
         FormatStyle(4)  =   "frmCarkillSetFW.frx":042C
         FormatStyle(5)  =   "frmCarkillSetFW.frx":0504
         ImageCount      =   0
         PrinterProperties=   "frmCarkillSetFW.frx":05BC
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
         Top             =   7530
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
         Top             =   7530
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
         Top             =   7530
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   6595
         TabIndex        =   4
         Top             =   7530
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmCarkillSetFW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'frmCommission2
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
'Private ComboSup As CComboSubGroup
Private ItemXlsFW As CXlsCarkillFW
'Private tempItemXlsFW As CComboSubGroupDe
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
   
  frmAddEditCarkillFW.ShowMode = SHOW_ADD
  frmAddEditCarkillFW.HeaderText = MapText("เพิ่มบรรทัดยกไป")

  Load frmAddEditCarkillFW
  frmAddEditCarkillFW.Show 1

  OKClick = frmAddEditCarkillFW.OKClick

  Unload frmAddEditCarkillFW
  Set frmAddEditCarkillFW = Nothing

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
   ItemXlsFW.XLS_FORWARD_ID = GridEX1.Value(1)

   If Not ConfirmDelete(GridEX1.Value(2)) Then
      Exit Sub
   End If

   Call EnableForm(Me, False)
   If Not glbDaily.DeleteCarkillFW(ItemXlsFW, IsOK, True, glbErrorLog) Then
       ItemXlsFW.XLS_FORWARD_ID = -1
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      Call EnableForm(Me, True)
      Exit Sub
   End If

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

  frmAddEditCarkillFW.ID = DETAIL_ID
  frmAddEditCarkillFW.HeaderText = MapText("แก้ไขบรรทัดยกไป")
  frmAddEditCarkillFW.ShowMode = SHOW_EDIT
  Load frmAddEditCarkillFW
  frmAddEditCarkillFW.Show 1

   OKClick = frmAddEditCarkillFW.OKClick

   Unload frmAddEditCarkillFW
   Set frmAddEditCarkillFW = Nothing

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
      ItemXlsFW.XLS_FORWARD_ID = -1

      If Not glbDaily.QueryCarkillFW(ItemXlsFW, m_Rs, ItemCount, IsOK, glbErrorLog) Then
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
   Col.Width = 1800
   Col.Caption = MapText("แถวที่")
   
   Set Col = GridEX1.Columns.Add '2
   Col.Width = 1600
   Col.Caption = MapText("คำนวณจากซื้อ-จ่าย")
   
   Set Col = GridEX1.Columns.Add '2
   Col.Width = 1600
   Col.Caption = MapText("แถวหลัก")
   
    GridEX1.ItemCount = 0
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   HeaderText = "ตั้งค่าบรรทัดคงเหลือยกไป"

   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption

   Me.Picture = LoadPicture(glbParameterObj.MainPicture)

   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19

   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)

   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))

End Sub

Private Sub Form_Load()
   OKClick = False

    m_HasActivate = False
    Set ItemXlsFW = New CXlsCarkillFW
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
   Call ItemXlsFW.PopulateFromRS(1, m_Rs)

   Values(1) = ItemXlsFW.XLS_FORWARD_ID
   Values(2) = ItemXlsFW.FW_ROW
   Values(3) = ItemXlsFW.UPPER_FLAG
   Values(4) = ItemXlsFW.MAIN_FLAG

Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
