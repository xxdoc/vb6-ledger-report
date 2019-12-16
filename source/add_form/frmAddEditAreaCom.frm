VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmAddEditAreaCom 
   BackColor       =   &H80000000&
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11910
   Icon            =   "frmAddEditAreaCom.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   8535
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   15055
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboAreaName 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1080
         Width           =   3375
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   5895
         Left            =   180
         TabIndex        =   0
         Top             =   1800
         Width           =   11505
         _ExtentX        =   20294
         _ExtentY        =   10398
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         TabKeyBehavior  =   1
         MultiSelect     =   -1  'True
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
         Column(1)       =   "frmAddEditAreaCom.frx":27A2
         Column(2)       =   "frmAddEditAreaCom.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditAreaCom.frx":290E
         FormatStyle(2)  =   "frmAddEditAreaCom.frx":2A6A
         FormatStyle(3)  =   "frmAddEditAreaCom.frx":2B1A
         FormatStyle(4)  =   "frmAddEditAreaCom.frx":2BCE
         FormatStyle(5)  =   "frmAddEditAreaCom.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditAreaCom.frx":2D5E
      End
      Begin VB.Label lblAreaName 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   0
         TabIndex        =   7
         Top             =   1100
         Width           =   1575
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
         MouseIcon       =   "frmAddEditAreaCom.frx":2F36
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
         MouseIcon       =   "frmAddEditAreaCom.frx":3250
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
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   10080
         TabIndex        =   4
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditAreaCom.frx":356A
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditAreaCom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private customerArea As CCommissionCustomerArea
Private tempCustomerArea As CCommissionCustomerArea

Private masterArea As CCommissMasterArea
Public HeaderText As String

Private m_Rs As ADODB.Recordset
Private mm_Rs As ADODB.Recordset
Public OKClick As Boolean
Public YEAR_ID As Long
Public ShowMode As SHOW_MODE_TYPE

Public MASTER_AREA_ID As Long
Dim RowDelete As Long

Private Sub cmdAdd_Click()
Dim ItemCount As Long
Dim OKClick As Boolean
   
   MASTER_AREA_ID = cboAreaName.ItemData(Minus2Zero(cboAreaName.ListIndex))
      
   frmAddEditAreaCom2.HeaderText = MapText("เพิ่มข้อมูลลูกค้า")
   frmAddEditAreaCom2.ShowMode = SHOW_ADD
   frmAddEditAreaCom2.MASTER_AREA_ID = MASTER_AREA_ID
   frmAddEditAreaCom2.YEAR_ID = YEAR_ID
   Load frmAddEditAreaCom2
    frmAddEditAreaCom2.Show 1

   OKClick = frmAddEditAreaCom2.OKClick

Unload frmAddEditAreaCom2
  Set frmAddEditAreaCom2 = Nothing

   If OKClick Then
      Call QueryData(True)
   End If
End Sub
Private Sub cmdDelete_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim ID As Long
Dim ROW As Long

   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   Call EnableForm(Me, False)
   
   For ROW = 1 To GridEX1.RowCount
      If GridEX1.RowSelected(ROW) = True Then
         ID = GridEX1.GetRowData(ROW).Value(1)                         ' อันที่จะถูกลบ
         customerArea.COMMISSION_CUS_AREA_ID = ID
         
         If Not ConfirmDelete(GridEX1.GetRowData(ROW).Value(2) & " " & GridEX1.GetRowData(ROW).Value(3)) Then
            Exit Sub
         End If

         If Not glbDaily.DeleteCusAreaCom(customerArea.COMMISSION_CUS_AREA_ID, IsOK, True, glbErrorLog) Then
              customerArea.COMMISSION_CUS_AREA_ID = -1
              glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
            customerArea.COMMISSION_CUS_ID = ""
            glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
            customerArea.COMMISSION_CUS_NAME = ""
            glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
            Call EnableForm(Me, True)
            Exit Sub
         End If
         
         End If
   Next ROW

   customerArea.COMMISSION_CUS_AREA_ID = 0
   Call QueryData(True)
   Call EnableForm(Me, True)
End Sub
Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim ID As Long
Dim OKClick As Boolean
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   ID = GridEX1.Value(1)
   
   masterArea.MASTER_AREA_ID = cboAreaName.ItemData(Minus2Zero(cboAreaName.ListIndex))
      
    frmAddEditAreaCom2.COMMISSION_CUS_AREA_ID = ID
    frmAddEditAreaCom2.MASTER_AREA_ID = MASTER_AREA_ID
    frmAddEditAreaCom2.HeaderText = MapText("แก้ไขข้อมูลลูกค้า")
   frmAddEditAreaCom2.ShowMode = SHOW_EDIT
   frmAddEditAreaCom2.YEAR_ID = YEAR_ID
   Load frmAddEditAreaCom2
   frmAddEditAreaCom2.Show 1

   OKClick = frmAddEditAreaCom2.OKClick

   Unload frmAddEditAreaCom2
  Set frmAddEditAreaCom2 = Nothing

   If OKClick Then
      Call QueryData(True)
   End If

End Sub

Private Sub cmdOK_Click()
'   If Not SaveData Then
'      Exit Sub
'   End If

   OKClick = True
   Unload Me
End Sub

Private Sub cmdSearch_Click()
   Call QueryData(True)
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      
      Call LoadAreaCom(cboAreaName)
      
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         MASTER_AREA_ID = -1
      End If
      
   End If
End Sub
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim m_ItemCount As Long
Dim Temp As Long

'   If Flag Then
'      End If
                     tempCustomerArea.YEAR_ID = YEAR_ID
                    tempCustomerArea.MASTER_AREA_ID = MASTER_AREA_ID
                      tempCustomerArea.COMMISSION_CUS_AREA_ID = -1
                     If Not glbDaily.QueryCusDetailAreaCom(tempCustomerArea, m_Rs, ItemCount, IsOK, glbErrorLog) Then
                        glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
                        Call EnableForm(Me, True)
                        Exit Sub
                     End If

                   masterArea.MASTER_AREA_ID = MASTER_AREA_ID
                   If Not glbDaily.QueryCommissArea(masterArea, mm_Rs, m_ItemCount, IsOK, glbErrorLog) Then
                        glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
                        Call EnableForm(Me, True)
                        Exit Sub
                   End If
                   
                     Call masterArea.PopulateFromRS(1, mm_Rs)

                     cboAreaName.ListIndex = IDToListIndex(cboAreaName, masterArea.MASTER_AREA_ID)

                  If Not IsOK Then
                     glbErrorLog.ShowUserError
                     Call EnableForm(Me, True)
                     Exit Sub
                  End If
                  
   GridEX1.ItemCount = ItemCount
   
   GridEX1.Rebind
   'debug.print RowDelete
  
   Call EnableForm(Me, True)
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean

   If Not VerifyCombo(lblAreaName, cboAreaName, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If

   masterArea.AddEditMode = SHOW_EDIT
'   masterArea.MASTER_AREA_NAME = txtAreaName.Text
   masterArea.MASTER_AREA_ID = cboAreaName.ItemData(Minus2Zero(cboAreaName.ListIndex))
   
   Call EnableForm(Me, False)
'
'   If Not glbDaily.AddEditComMasterArea(masterArea, IsOK, True, glbErrorLog) Then
'      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
'      SaveData = False
'      Call EnableForm(Me, True)
'      Exit Function
'   End If

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
  Col.Caption = MapText("ID")
   
   Set Col = GridEX1.Columns.Add '2
   Col.Width = 2000
   Col.Caption = MapText("รหัสลูกค้า")
   
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 4000
    Col.Caption = MapText("ชื่อลูกค้า")

   GridEX1.ItemCount = 0
End Sub
Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = MapText("ข้อมูลลูกของเขตการขาย")
   pnlHeader.Caption = Me.Caption
   
   Call InitGrid
   
   Call InitNormalLabel(lblAreaName, MapText("เขตการขาย"))
   Call InitCombo(cboAreaName)

   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
'   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
'   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   
End Sub


Private Sub Form_Load()
      OKClick = False
         m_HasActivate = False
   m_HasModify = False
   
   Set customerArea = New CCommissionCustomerArea
   Set tempCustomerArea = New CCommissionCustomerArea
   Set masterArea = New CCommissMasterArea
   Set m_Rs = New ADODB.Recordset
   Set mm_Rs = New ADODB.Recordset
   
         m_HasActivate = False
   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub
Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
   RowBuffer.RowStyle = RowBuffer.Value(2)
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
   Call tempCustomerArea.PopulateFromRS(3, m_Rs)
   
   Values(1) = tempCustomerArea.COMMISSION_CUS_AREA_ID
   Values(2) = tempCustomerArea.COMMISSION_CUS_ID
   Values(3) = tempCustomerArea.COMMISSION_CUS_NAME
   RowDelete = RowIndex + 1
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
'   cmdExit.Top = ScaleHeight - 580
'   cmdExit.Left = ScaleWidth - cmdExit.Width - 50
   cmdOK.Left = ScaleWidth - cmdOK.Width - 50
End Sub

Private Sub txtAreaName_hasChange()
   m_HasModify = True
End Sub

Private Sub cboAreaName_Click()
 m_HasModify = True
End Sub
