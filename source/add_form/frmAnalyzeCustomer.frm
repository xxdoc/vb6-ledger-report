VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmAnalyzeCustomer 
   BackColor       =   &H80000000&
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11910
   Icon            =   "frmAnalyzeCustomer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Threed.SSFrame SSFrame1 
      Height          =   8535
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   15055
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboOrderBy 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1440
         Width           =   2955
      End
      Begin VB.ComboBox cboOrderType 
         Height          =   315
         Left            =   6120
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1440
         Width           =   2955
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   5655
         Left            =   180
         TabIndex        =   11
         Top             =   2040
         Width           =   11505
         _ExtentX        =   20294
         _ExtentY        =   9975
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
         Column(1)       =   "frmAnalyzeCustomer.frx":27A2
         Column(2)       =   "frmAnalyzeCustomer.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAnalyzeCustomer.frx":290E
         FormatStyle(2)  =   "frmAnalyzeCustomer.frx":2A6A
         FormatStyle(3)  =   "frmAnalyzeCustomer.frx":2B1A
         FormatStyle(4)  =   "frmAnalyzeCustomer.frx":2BCE
         FormatStyle(5)  =   "frmAnalyzeCustomer.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAnalyzeCustomer.frx":2D5E
      End
      Begin prjLedgerReport.uctlTextBox txtSaleCode 
         Height          =   435
         Left            =   1680
         TabIndex        =   0
         Top             =   840
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   767
      End
      Begin prjLedgerReport.uctlTextBox txtCustomerCode 
         Height          =   435
         Left            =   6050
         TabIndex        =   1
         Top             =   840
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   767
      End
      Begin VB.Label lblCustomerCode 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4530
         TabIndex        =   17
         Top             =   900
         Width           =   1455
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   525
         Left            =   9960
         TabIndex        =   3
         Top             =   1440
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   9960
         TabIndex        =   2
         Top             =   840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAnalyzeCustomer.frx":2F36
         ButtonStyle     =   3
      End
      Begin VB.Label lblOrderBy 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   120
         TabIndex        =   16
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label lblOrderType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4680
         TabIndex        =   15
         Top             =   1440
         Width           =   1365
      End
      Begin VB.Label lblSaleCode 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   0
         TabIndex        =   14
         Top             =   930
         Width           =   1575
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
         MouseIcon       =   "frmAnalyzeCustomer.frx":3250
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
         MouseIcon       =   "frmAnalyzeCustomer.frx":356A
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
         MouseIcon       =   "frmAnalyzeCustomer.frx":3884
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAnalyzeCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_Customer As CARTrn
Private m_TempCustomer As CARTrn
Private m_TempData As CARTrn
Private m_Rs As ADODB.Recordset
Private m_ReceiveAmounts As Collection
Private m_ReceiveAllAmounts As Collection
Private m_CnAllAmounts As Collection
Private m_AnalyzeCustomer As Collection
Private ImportExportItems As Collection
Public OKClick As Boolean
Public Ari As CARRcIt
Public Ari2 As CARRcIt
Public Apt1 As CARTrn
Public ARt2 As CARTrn
Public Ac As CAnalyzeCustomer
Public PaidBalance As Double 'ชำระแล้วยกมา
Private Sub cmdAdd_Click()
Dim ItemCount As Long
Dim OKClick As Boolean
   If OKClick Then
      Call QueryData(True)
   End If
End Sub
Private Sub cmdDelete_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim ID As Long

   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   
   Call QueryData(True)
   Call EnableForm(Me, True)
End Sub
Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim ItemCount As Long
Dim IsCanLock As Boolean
Dim ID As Long
Dim SubID As Long
Dim OKClick As Boolean
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   
   SubID = GridEX1.Value(1)
   
   frmAddEditAnalyzeCustomer.SubID = SubID
   frmAddEditAnalyzeCustomer.HeaderText = MapText("แก้ไขข้อมูลวิเคราะห์อายุหนี้")
   frmAddEditAnalyzeCustomer.ShowMode = SHOW_EDIT
   Set frmAddEditAnalyzeCustomer.TempAddEditDataCollection = m_AnalyzeCustomer
   Set frmAddEditAnalyzeCustomer.RunDataCollection = ImportExportItems
   Set frmAddEditAnalyzeCustomer.ParentForm = Me
    Load frmAddEditAnalyzeCustomer
   frmAddEditAnalyzeCustomer.Show 1

   OKClick = frmAddEditAnalyzeCustomer.OKClick

   Unload frmAddEditAnalyzeCustomer
   Set frmAddEditAnalyzeCustomer = Nothing
               
   If OKClick Then
      Call QueryData(True)
   End If

End Sub
Private Sub cmdOK_Click()
   OKClick = True
   Unload Me
End Sub
Private Sub cmdSearch_Click()
   Call QueryData(True)
End Sub
Private Sub cmdClear_Click()
   txtCustomerCode.Text = ""
   txtSaleCode.Text = ""
'   cboOrderBy.ListIndex = -1
'   cboOrderType.ListIndex = -1
End Sub
Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      
      Call LoadReceiveAmountByBill(Nothing, m_ReceiveAmounts, -1, -1)
      Call LoadReceiveAmountByBill(Nothing, m_ReceiveAllAmounts, -1, -1)
      Call LoadARCNAmountByBill(Nothing, m_CnAllAmounts, -1, -1, -1, -1)       'ต้องบวกยอดยกมาเพิ่ม ยอด CN ทั้งหมดของบิลหลัง CN แต่ เอกสาร LINK เป็นในช่วงยกมา
'      Call InitCustomerOrderBy(cboOrderBy)
'      Call InitOrderType(cboOrderType)
      
'      Call QueryData(True)
   End If
End Sub
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long
Dim Temp As Long
Dim ItemID As Long

   If Flag Then
      Call EnableForm(Me, False)
      m_Customer.CUSCOD = PatchWildCard(txtCustomerCode.Text)
      m_Customer.SLMCOD = PatchWildCard(txtSaleCode.Text)
      m_Customer.RecTypeSet = "('3', '4', '5')"
      m_Customer.RECTYP = ""
      m_Customer.OrderBy = 4
'      m_Customer.OrderBy = cboOrderBy.ItemData(Minus2Zero(cboOrderBy.ListIndex))
'      m_Customer.OrderType = cboOrderType.ItemData(Minus2Zero(cboOrderType.ListIndex))
      If Not glbDaily.QueryARTran2(m_Customer, m_Rs, ItemCount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
'   Call LoadAnalyzeCustomer(Nothing, m_AnalyzeCustomer)
   
   Set ImportExportItems = Nothing
   Set ImportExportItems = New Collection
   ItemID = 1
   
   While Not m_Rs.EOF
      Set m_TempData = New CARTrn
      Call m_TempData.PopulateFromRS(1, m_Rs)
      Set Ari = GetARRcpItem(m_ReceiveAmounts, m_TempData.DOCNUM)
      Set Ari2 = GetARRcpItemEx(m_ReceiveAllAmounts, m_TempData.DOCNUM)
      Set ARt2 = GetARTrn(m_CnAllAmounts, m_TempData.DOCNUM)
      PaidBalance = m_TempData.RCVAMT - Ari2.RCVAMT - ARt2.AMOUNT
      If Not (m_TempData.AMOUNT) > (Ari.RCVAMT + PaidBalance) Then
         ItemCount = ItemCount - 1
         m_TempData.Flag = "D"
         m_TempData.KEY_ID = ItemID
      Else
         m_TempData.Flag = "A"
         m_TempData.KEY_ID = ItemID
      End If
      Call ImportExportItems.Add(m_TempData, Str(ItemID))
      ItemID = ItemID + 1
      
      Set m_TempData = Nothing
      
      m_Rs.MoveNext
   Wend

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
   Col.Width = 4000
   Col.Caption = MapText("ลูกค้า")
   
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 2000
   Col.Caption = MapText("เลขที่ INV.")
   
   Set Col = GridEX1.Columns.Add '4
   Col.Width = 2000
    Col.Caption = MapText("วันที่ขาย")
   
   Set Col = GridEX1.Columns.Add '5
   Col.Width = ScaleWidth - 5200
   Col.Caption = MapText("ประมาณการวันที่ชำระ")
   
   GridEX1.ItemCount = 0
End Sub
Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = MapText("ข้อมูลวิเคราะห์อายุหนี้")
   pnlHeader.Caption = Me.Caption
   
   Call InitGrid
   
   Call InitNormalLabel(lblSaleCode, MapText("รหัส SALE"))
   Call InitNormalLabel(lblCustomerCode, MapText("รหัสลูกค้า"))
'   Call InitNormalLabel(lblOrderBy, MapText("เรียงตาม"))
'   Call InitNormalLabel(lblOrderType, MapText("เรียงจาก"))
   lblOrderType.Visible = False
   lblOrderBy.Visible = False
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
'   Call InitCombo(cboOrderBy)
'   Call InitCombo(cboOrderType)
   cboOrderBy.Visible = False
   cboOrderType.Visible = False
   
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
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   cmdAdd.Enabled = False
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   cmdDelete.Enabled = False
   Call InitMainButton(cmdSearch, MapText("ค้นหา (F5)"))
   Call InitMainButton(cmdClear, MapText("เคลียร์ (F4)"))
   
End Sub
Private Sub cmdExit_Click()
   OKClick = False
   Unload Me
End Sub
Private Sub Form_Load()
   
   Set m_Customer = New CARTrn
   Set m_TempCustomer = New CARTrn
   Set m_TempData = New CARTrn
   Set m_Rs = New ADODB.Recordset
   Set m_ReceiveAmounts = New Collection
   Set m_ReceiveAllAmounts = New Collection
   Set m_CnAllAmounts = New Collection
   Set m_AnalyzeCustomer = New Collection
   Set ImportExportItems = New Collection

   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Set m_Customer = Nothing
   Set m_TempCustomer = Nothing
   Set m_TempData = Nothing
   Set m_Rs = Nothing
   Set m_ReceiveAmounts = Nothing
   Set m_ReceiveAllAmounts = Nothing
   Set m_CnAllAmounts = Nothing
   Set m_AnalyzeCustomer = Nothing
   Set ImportExportItems = Nothing
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

'   If m_Rs.EOF Then
'      Exit Sub
'   End If

   If RowIndex <= 0 Then
      Exit Sub
   End If

      If ImportExportItems.Count <= 0 Then
         Exit Sub
      End If
      Set m_TempCustomer = GetItem(ImportExportItems, RowIndex, RealIndex)
      If m_TempCustomer Is Nothing Then
         Exit Sub
      End If
      
         Values(1) = m_TempCustomer.KEY_ID
         Values(2) = m_TempCustomer.CUSNAM
         Values(3) = m_TempCustomer.DOCNUM
         Values(4) = DateToStringExtEx2(m_TempCustomer.DOCDAT)
         
         Call LoadAnalyzeCustomer(Nothing, m_AnalyzeCustomer)
         Set Ac = GetAnalyzeCustomer(m_AnalyzeCustomer, Trim(m_TempCustomer.DOCNUM))
         If Ac.ANALYZE_CUSTOMER_ID = 0 Then
            Values(5) = ""
         Else
            Values(5) = DateToStringExtEx2(Ac.DATE_OF_PAYMENT)
        End If
  
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
