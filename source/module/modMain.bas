Attribute VB_Name = "modMain"
Option Explicit
Public Const ROOT_TREE = "Root"

Public Const DUMMY_KEY = 27
Public Const DUMMY_KEY2 = 16
Public GLB_GRID_COLOR As Long
Public GLB_NORMAL_COLOR As Long
Public GLB_ALERT_COLOR As Long
Public GLB_SHOW_COLOR As Long
Public GLB_FORM_COLOR As Long
Public GLB_HEAD_COLOR As Long
Public GLB_GRIDHD_COLOR As Long
Public GLB_MANDATORY_COLOR As Long

Public Enum SHOW_MODE_TYPE
   SHOW_ADD = 1
   SHOW_EDIT = 2
   SHOW_VIEW = 3
   SHOW_VIEW_ONLY = 4
End Enum

Public Enum TEXT_BOX_TYPE
   TEXT_STRING = 1
   TEXT_INTEGER = 2
   TEXT_FLOAT = 3
   TEXT_FLOAT_MONEY = 4
   TEXT_INTEGER_MONEY = 5
End Enum

Public Enum LANGUAGE_TYPE
   LANG_ENG = 1
   LANG_THAI = 2
End Enum

Public Enum UNIQUE_TYPE
   EMPCODE_UNIQUE = 1
   EMPNAME_LASTNAME_UNIQUE = 2
   TRUCK_UNIQUE = 3
   DO_PLAN_UNIQUE = 4
   DBN_UNIQUE = 5
   CUSTCODE_UNIQUE = 6
   USERGROUP_UNIQUE = 7
   USERNAME_UNIQUE = 8
   IMPORT_UNIQUE = 9
   EXPORT_UNIQUE = 10
   REPAIR_UNIQUE = 11
   REPAIR_FORMULA_UNIQUE = 12
   SUPPLIER_UNIQUE = 13
   PARTNO_UNIQUE = 14
   QUOATATION_UNIQUE = 15
   TEACHER_UNIQUE = 16
   SUBJECT_UNIQUE = 17
   FACULTY_UNIQUE = 18
   EXPENSE_UNIQUE = 19
   PO_UNIQUE = 20
   CUSTOMER_UNIQUE = 21
   REVENUE_UNIQUE = 22
   BORROW_UNIQUE = 23
   PRDFEATURE_UNIQUE = 24
   JOBPLAN_UNIQUE = 25
   
   PARTTYPE_NO = 26
   PARTTYPE_NAME = 27
   LOCATION_NO = 28
   LOCATION_NAME = 29
   PRODUCTTYPE_NO = 30
   PRODUCTTYPE_NAME = 31
   PRODUCTSTATUS_NO = 32
   PRODUCTSTATUS_NAME = 33
   HOUSE_NO = 34
   HOUSE_NAME = 35
   COUNTRY_NO = 36
   COUNTRY_NAME = 37
   CSTGRADE_NO = 38
   CSTGRADE_NAME = 39
   CSTTYPE_NO = 40
   CSTTYPE_NAME = 41
   SUPPLIERTYPE_NO = 42
   SUPPLIERYPE_NAME = 43
   SUPPLIERGRADE_NO = 44
   SUPPLIERGRADE_NAME = 45
   SUPPLIERSTATUS_NO = 46
   SUPPLIERSTATUS_NAME = 47
   POSITION_NO = 48
   UNIT_NO = 49
   UNIT_NAME = 50
   YEAR_NO = 51
   PARTGROUP_NO = 52
   PARTGROUP_NAME = 53
   LOCATION_NO_EX = 54
   REAL_CREDIT_NO = 55
   DOCUMENT_CANCEL_NO = 56
   MASTER_FT_UNIQUE = 57   ' คืออะไร
   CUS_PIG_TYPE_UNIQUE = 58
   
   PROVINCE_MAP_UNIQUE = 59
   COST_PRODUCTS_UNIQUE = 60
End Enum

Public Enum MASTER_COMMISSION_AREA
   COMMISSION_TABLE = 1
   RETURN_TABLE = 2
   COMMISSION_CHART = 3
   COMMISSION_TABLE_EX = 4
End Enum

Public Enum FIELD_TYPE
   INT_TYPE = 1
   MONEY_TYPE = 2
   DATE_TYPE = 3
   STRING_TYPE = 4
   BOOLEAN_TYPE = 5
End Enum

Public Enum FIELD_CAT
   ID_CAT = 1
   MODIFY_DATE_CAT = 2
   CREATE_DATE_CAT = 3
   MODIFY_BY_CAT = 4
   CREATE_BY_CAT = 5
   DATA_CAT = 6
   TEMP_CAT = 7
End Enum

Public Enum NUMBER_TYPE
   PO_NUMBER = 1
   OPERATE_NUMBER = 2
   BORROW_NUMBER = 3
   DEBIT_NOTE_NUMBER = 4
   'bum+
   EXPENSE_NUMBER = 5
   REPAIR_NUMBER = 6
   IMPORT_NUMBER = 7
   EXPORT_NUMBER = 8
   PLAN_NUMBER = 9
   FUEL_NUMBER1 = 10
   FUEL_NUMBER2 = 11
   BILL_NUMBER = 13
   QUOATATION_NUMBER = 14
   REVENUE_NUMBER = 15
   DO_NUMBER = 16
   RECEIPT_NUMBER = 17
   JOBPLAN_NUMBER = 18
   INVOICE_RECEIPT_NUMBER = 19
End Enum

'===================== For clear treeview =========================
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd _
    As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Const TV_FIRST As Long = &H1100
Const TVM_GETNEXTITEM As Long = (TV_FIRST + 10)
Const TVM_DELETEITEM As Long = (TV_FIRST + 1)
Const TVGN_ROOT As Long = &H0
Const WM_SETREDRAW As Long = &HB
'===================== For clear treeview =========================

Public Const PROJECT_NAME = "Mittraphap Report Management"
Public Const GLB_FONT = "JasmineUPC"
Private Const MODULE_NAME = "modMain"

Public glbErrorLog As clsErrorLog
Public glbDatabaseMngr As clsDatabaseMngr
Public glbSetting As clsGlobalSetting
Public glbParameterObj As clsParameter
Public glbUser As CUser
Public glbGroup As CGroup
'Public glbAdmin As clsAdmin
'Public glbMaster As clsMaster
Public glbDaily As clsDaily
'Public glbEnterPrise As CEnterprise
'
'Public glbLoginTracking As CLoginTracking
Public glbAccessRight As Collection
Public glbCompanyName As String
Public glbCompanyCode As String
Public glbCompanyFullName  As String
'Public glbEnterPrise As CEnterprise

Public m_DrCrFlage As Boolean

Public m_SupplierColl As Collection

Public Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Function VerifyDate(L As Label, D As uctlDate, Optional NullAllow As Boolean = False) As Boolean
Dim S As String
   If L Is Nothing Then
      S = ""
   Else
      S = L.Caption
   End If

   If Not D.VerifyDate(NullAllow) Then
      VerifyDate = False
      D.SetFocus
      Call MsgBox("กรุณากรอกข้อมูล " & " '" & S & "' " & "ให้ถูกต้องและครบถ้วน ", vbOKOnly, PROJECT_NAME)
   Else
      VerifyDate = True
   End If
End Function

Public Function VerifyTime(L As Label, T As uctlTime, Optional NullAllow As Boolean = False) As Boolean
Dim S As String
   If L Is Nothing Then
      S = ""
   Else
      S = L.Caption
   End If

   If Not T.VerifyTime(NullAllow) Then
      VerifyTime = False
      T.SetFocus
      Call MsgBox("กรุณากรอกข้อมูล " & " '" & S & "' " & "ให้ถูกต้องและครบถ้วน ", vbOKOnly, PROJECT_NAME)
   Else
      VerifyTime = True
   End If
End Function
Public Function AdjustType2Code(TempID As Long) As String
   If TempID = 1 Then
      AdjustType2Code = "E"
   ElseIf TempID = 2 Then
      AdjustType2Code = "I"
   End If
End Function
Public Function Code2AdjustType(Cd As String) As Long
   If Cd = "E" Then
      Code2AdjustType = 1
   ElseIf Cd = "I" Then
      Code2AdjustType = 2
   End If
End Function
Public Function VerifyTextData(L As Label, T As TextBox, Optional NullAllow As Boolean = False) As Boolean
Dim S As String
   If L Is Nothing Then
      S = ""
   Else
      S = L.Caption
   End If
   
   If Not NullAllow Then
      If Len(Trim(T.Text)) = 0 Then
         VerifyTextData = False
         Call MsgBox("กรุณากรอกข้อมูล " & " '" & S & "' " & "ให้ถูกต้องและครบถ้วน ", vbOKOnly, PROJECT_NAME)
         If T.Enabled Then
            T.SetFocus
         End If
         Exit Function
      End If
   End If
   
   If (T.Tag = TEXT_INTEGER) Or (T.Tag = TEXT_FLOAT) Or (T.Tag = TEXT_FLOAT_MONEY) Or (T.Tag = TEXT_INTEGER_MONEY) Then
      If Trim(T.Text) = "" Then
         If NullAllow Then
            VerifyTextData = True
            Exit Function
         End If
      End If
      If IsNumeric(Trim(T.Text)) Then
         If InStr(1, T.Text, ".") <= 0 Then
            If Val(Trim(T.Text)) < 0 Then
               VerifyTextData = False
            Else
               VerifyTextData = True
               Exit Function
            End If
         Else
            If T.Tag = TEXT_INTEGER Then
               VerifyTextData = False
            Else
               If Val(Trim(T.Text)) < 0 Then
                  VerifyTextData = False
               Else
                  VerifyTextData = True
               End If
            End If
            Exit Function
         End If
      End If
      
      VerifyTextData = False
      Call MsgBox("กรุณากรอกข้อมูล " & " '" & S & "' " & "ให้ถูกต้องและครบถ้วน ", vbOKOnly, PROJECT_NAME)
      If T.Enabled Then
         T.SetFocus
      End If
      Exit Function
   ElseIf T.Tag = TEXT_STRING Then
      If (InStr(1, T.Text, ";") > 0) Or (InStr(1, T.Text, "|") > 0) Then
         Call MsgBox("กรุณากรอกข้อมูล " & " '" & S & "' " & "ให้ถูกต้องและครบถ้วน ", vbOKOnly, PROJECT_NAME)
         T.SetFocus
         
         VerifyTextData = False
         Exit Function
      End If
      
      VerifyTextData = True
   End If
End Function

Public Function VerifyTextControl(L As Label, T As uctlTextBox, Optional NullAllow As Boolean = False) As Boolean
Dim S As String
   If L Is Nothing Then
      S = ""
   Else
      S = L.Caption
   End If
   
   If Not NullAllow Then
      If Len(Trim(T.Text)) = 0 Then
         VerifyTextControl = False
         Call MsgBox("กรุณากรอกข้อมูล " & " '" & S & "' " & "ให้ถูกต้องและครบถ้วน ", vbOKOnly, PROJECT_NAME)
         If T.Enabled Then
            T.SetFocus
         End If
         Exit Function
      End If
   End If
   
   If (T.Tag = TEXT_INTEGER) Or (T.Tag = TEXT_FLOAT) Or (T.Tag = TEXT_FLOAT_MONEY) Or (T.Tag = TEXT_INTEGER_MONEY) Then
      If Trim(T.Text) = "" Then
         If NullAllow Then
            VerifyTextControl = True
            Exit Function
         End If
      End If
      If IsNumeric(Trim(T.Text)) Then
         If InStr(1, T.Text, ".") <= 0 Then
            If Val(Trim(T.Text)) < 0 Then
               VerifyTextControl = True 'false
               Exit Function 'remove this if false
            Else
               VerifyTextControl = True
               Exit Function
            End If
         Else
            If T.Tag = TEXT_INTEGER Then
               VerifyTextControl = False
            Else
               If Val(Trim(T.Text)) < 0 Then
                  VerifyTextControl = True 'false
                  Exit Function
               Else
                  VerifyTextControl = True
                  Exit Function
               End If
            End If
'            Exit Function
         End If
      End If
      
      VerifyTextControl = False
      Call MsgBox("กรุณากรอกข้อมูล " & " '" & S & "' " & "ให้ถูกต้องและครบถ้วน ", vbOKOnly, PROJECT_NAME)
      If T.Enabled Then
         T.SetFocus
      End If
      Exit Function
   ElseIf T.Tag = TEXT_STRING Then
      If (InStr(1, T.Text, ";") > 0) Or (InStr(1, T.Text, "|") > 0) Then
         Call MsgBox("กรุณากรอกข้อมูล " & " '" & S & "' " & "ให้ถูกต้องและครบถ้วน ", vbOKOnly, PROJECT_NAME)
         T.SetFocus
         
         VerifyTextControl = False
         Exit Function
      End If
      
      VerifyTextControl = True
   End If
End Function

'Private Sub GetParentItemDesc(Acc As String, Ri As CRightItem, ReportName As String)
'   Ri.DEFAULT_VALUE = "N"
'   If Acc = "ADMIN" Then
'      Ri.RIGHT_ITEM_DESC = "ระบบข้อมูลผู้ใช้งาน"
'   ElseIf Acc = "ADMIN_GROUP" Then
'      Ri.RIGHT_ITEM_DESC = "ระบบข้อมูลกลุ่มผู้ใช้งาน"
'   ElseIf Acc = "ADMIN_GROUP_QUERY" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถเรียกค้นข้อมูลกลุ่มผู้ใช้งาน"
'   ElseIf Acc = "ADMIN_GROUP_ADD" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถเพิ่มข้อมูลกลุ่มผู้ใช้งาน"
'   ElseIf Acc = "ADMIN_GROUP_EDIT" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถแก้ไขข้อมูลกลุ่มผู้ใช้งาน"
'   ElseIf Acc = "ADMIN_GROUP_DELETE" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถลบข้อมูลกลุ่มผู้ใช้งาน"
'
'   ElseIf Acc = "ADMIN_USER" Then
'      Ri.RIGHT_ITEM_DESC = "ระบบข้อมูลผู้ใช้งาน"
'   ElseIf Acc = "ADMIN_USER_QUERY" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถเรียกค้นข้อมูลผู้ใช้งาน"
'   ElseIf Acc = "ADMIN_USER_ADD" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถเพิ่มข้อมูลผู้ใช้งาน"
'   ElseIf Acc = "ADMIN_USER_EDIT" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถแก้ไขข้อมูลผู้ใช้งาน"
'   ElseIf Acc = "ADMIN_USER_DELETE" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถลบข้อมูลผู้ใช้งาน"
'
'   ElseIf Acc = "ADMIN_REPORT" Then
'      Ri.RIGHT_ITEM_DESC = "รายงานระบบข้อมูลผู้ใช้งาน"
'   ElseIf Mid(Acc, 1, 12) = "ADMIN_REPORT" Then
'      Ri.RIGHT_ITEM_DESC = ReportName
'
'   ElseIf Acc = "MASTER" Then
'      Ri.RIGHT_ITEM_DESC = "ระบบข้อมูลหลัก"
'
'   ElseIf Acc = "MASTER_MAIN" Then
'      Ri.RIGHT_ITEM_DESC = "ระบบข้อมูลหลักส่วนกลาง"
'   ElseIf Acc = "MASTER_MAIN_QUERY" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถค้นหาข้อมูลหลักส่วนกลาง"
'   ElseIf Acc = "MASTER_MAIN_ADD" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถเพิ่มข้อมูลหลักส่วนกลาง"
'   ElseIf Acc = "MASTER_MAIN_EDIT" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถแก้ไขข้อมูลหลักส่วนกลาง"
'   ElseIf Acc = "MASTER_MAIN_DELETE" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถลบข้อมูลหลักส่วนกลาง"
'
'   ElseIf Acc = "MASTER_INVENTORY" Then
'      Ri.RIGHT_ITEM_DESC = "ระบบข้อมูลหลักคลัง"
'   ElseIf Acc = "MASTER_INVENTORY_QUERY" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถค้นหาข้อมูลหลักคลัง"
'   ElseIf Acc = "MASTER_INVENTORY_ADD" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถเพิ่มข้อมูลหลักคลัง"
'   ElseIf Acc = "MASTER_INVENTORY_EDIT" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถแก้ไขข้อมูลหลักคลัง"
'   ElseIf Acc = "MASTER_INVENTORY_DELETE" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถลบข้อมูลหลักคลัง"
'
'   ElseIf Acc = "MASTER_PIG" Then
'      Ri.RIGHT_ITEM_DESC = "ระบบข้อมูลหลักระบบบริหารสุกร"
'   ElseIf Acc = "MASTER_PIG_QUERY" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถค้นหาข้อมูลหลักระบบบริหารสุกร"
'   ElseIf Acc = "MASTER_PIG_ADD" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถเพิ่มข้อมูลหลักระบบบริหารสุกร"
'   ElseIf Acc = "MASTER_PIG_EDIT" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถแก้ไขข้อมูลหลักระบบบริหารสุกร"
'   ElseIf Acc = "MASTER_PIG_DELETE" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถลบข้อมูลหลักระบบบริหารสุกร"
'
'   ElseIf Acc = "MASTER_LEDGER" Then
'      Ri.RIGHT_ITEM_DESC = "ระบบข้อมูลหลักระบบบัญชี"
'   ElseIf Acc = "MASTER_LEDGER_QUERY" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถค้นหาข้อมูลหลักระบบบัญชี"
'   ElseIf Acc = "MASTER_LEDGER_ADD" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถเพิ่มข้อมูลหลักระบบบัญชี"
'   ElseIf Acc = "MASTER_LEDGER_EDIT" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถแก้ไขข้อมูลหลักระบบบัญชี"
'   ElseIf Acc = "MASTER_LEDGER_DELETE" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถลบข้อมูลหลักระบบบัญชี"
'
'   ElseIf Acc = "MASTER_INVENTORY" Then
'      Ri.RIGHT_ITEM_DESC = "ระบบข้อมูลหลักระบบคลัง"
'   ElseIf Acc = "MASTER_PRODUCTION" Then
'      Ri.RIGHT_ITEM_DESC = "ระบบข้อมูลหลักระบบการผลิต"
'   ElseIf Acc = "MASTER_LEDGER" Then
'      Ri.RIGHT_ITEM_DESC = "ระบบข้อมูลหลักระบบบัญชี"
'
'   ElseIf Acc = "MAIN" Then
'      Ri.RIGHT_ITEM_DESC = "ระบบข้อมูลส่วนกลาง"
'
'   ElseIf Acc = "MAIN_CUSTOMER" Then
'      Ri.RIGHT_ITEM_DESC = "ระบบข้อมูลลูกค้า"
'   ElseIf Acc = "MAIN_CUSTOMER_QUERY" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถค้นหาข้อมูลลูกค้า"
'   ElseIf Acc = "MAIN_CUSTOMER_ADD" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถเพิ่มข้อมูลลูกค้า"
'   ElseIf Acc = "MAIN_CUSTOMER_EDIT" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถแก้ไขข้อมูลลูกค้า"
'   ElseIf Acc = "MAIN_CUSTOMER_DELETE" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถลบข้อมูลลูกค้า"
'
'   ElseIf Acc = "MAIN_SUPPLIER" Then
'      Ri.RIGHT_ITEM_DESC = "ระบบข้อมูลซัพพลายเออร์"
'   ElseIf Acc = "MAIN_SUPPLIER_QUERY" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถค้นหาข้อมูลซัพพลายเออร์"
'   ElseIf Acc = "MAIN_SUPPLIER_ADD" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถเพิ่มข้อมูลซัพพลายเออร์"
'   ElseIf Acc = "MAIN_SUPPLIER_EDIT" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถแก้ไขข้อมูลซัพพลายเออร์"
'   ElseIf Acc = "MAIN_SUPPLIER_DELETE" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถลบข้อมูลซัพพลายเออร์"
'
'   ElseIf Acc = "MAIN_ENTERPRISE" Then
'      Ri.RIGHT_ITEM_DESC = "ระบบข้อมูลบริษัท"
'   ElseIf Acc = "MAIN_ENTERPRISE_QUERY" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถค้นหาข้อมูลบริษัท"
'   ElseIf Acc = "MAIN_ENTERPRISE_ADD" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถเพิ่มข้อมูลบริษัท"
'   ElseIf Acc = "MAIN_ENTERPRISE_EDIT" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถแก้ไขข้อมูลบริษัท"
'   ElseIf Acc = "MAIN_ENTERPRISE_DELETE" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถลบข้อมูลบริษัท"
'
'   ElseIf Acc = "MAIN_EMPLOYEE" Then
'      Ri.RIGHT_ITEM_DESC = "ระบบข้อมูลพนักงาน"
'   ElseIf Acc = "MAIN_EMPLOYEE_QUERY" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถค้นหาข้อมูลพนักงาน"
'   ElseIf Acc = "MAIN_EMPLOYEE_ADD" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถเพิ่มข้อมูลพนักงาน"
'   ElseIf Acc = "MAIN_EMPLOYEE_EDIT" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถแก้ไขข้อมูลพนักงาน"
'   ElseIf Acc = "MAIN_EMPLOYEE_DELETE" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถลบข้อมูลพนักงาน"
'
'   ElseIf Acc = "MAIN_REPORT" Then
'      Ri.RIGHT_ITEM_DESC = "รายงานระบบข้อมูลส่วนกลาง"
'   ElseIf Mid(Acc, 1, 11) = "MAIN_REPORT" Then
'      Ri.RIGHT_ITEM_DESC = ReportName
'
'   ElseIf Acc = "INVENTORY" Then
'      Ri.RIGHT_ITEM_DESC = "ระบบข้อมูลคลัง"
'
'   ElseIf Acc = "INVENTORY_PART" Then
'      Ri.RIGHT_ITEM_DESC = "ระบบข้อมูลวัตถุดิบ"
'   ElseIf Acc = "INVENTORY_PART_QUERY" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถค้นหาข้อมูลวัตถุดิบ"
'   ElseIf Acc = "INVENTORY_PART_ADD" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถเพิ่มข้อมูลวัตถุดิบ"
'   ElseIf Acc = "INVENTORY_PART_EDIT" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถแก้ไขข้อมูลวัตถุดิบ"
'   ElseIf Acc = "INVENTORY_PART_DELETE" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถลบข้อมูลวัตถุดิบ"
'
'   ElseIf Acc = "INVENTORY_IMPORT" Then
'      Ri.RIGHT_ITEM_DESC = "ระบบข้อมูลการนำเข้าวัตถุดิบ"
'   ElseIf Acc = "INVENTORY_IMPORT_QUERY" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถค้นหาข้อมูลการนำเข้าวัตถุดิบ"
'   ElseIf Acc = "INVENTORY_IMPORT_ADD" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถเพิ่มข้อมูลการนำเข้าวัตถุดิบ"
'   ElseIf Acc = "INVENTORY_IMPORT_EDIT" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถแก้ไขข้อมูลการนำเข้าวัตถุดิบ"
'   ElseIf Acc = "INVENTORY_IMPORT_DELETE" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถลบข้อมูลการนำเข้าวัตถุดิบ"
'
'   ElseIf Acc = "INVENTORY_EXPORT" Then
'      Ri.RIGHT_ITEM_DESC = "ระบบข้อมูลการเบิกจ่ายวัตถุดิบ"
'   ElseIf Acc = "INVENTORY_EXPORT_QUERY" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถค้นหาข้อมูลการเบิกจ่ายวัตถุดิบ"
'   ElseIf Acc = "INVENTORY_EXPORT_ADD" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถเพิ่มข้อมูลการเบิกจ่ายวัตถุดิบ"
'   ElseIf Acc = "INVENTORY_EXPORT_EDIT" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถแก้ไขข้อมูลการเบิกจ่ายวัตถุดิบ"
'   ElseIf Acc = "INVENTORY_EXPORT_DELETE" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถลบข้อมูลการเบิกจ่ายวัตถุดิบ"
'
'   ElseIf Acc = "INVENTORY_TRANSFER" Then
'      Ri.RIGHT_ITEM_DESC = "ระบบข้อมูลการโอนวัตถุดิบ"
'   ElseIf Acc = "INVENTORY_TRANSFER_QUERY" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถค้นหาข้อมูลการโอนวัตถุดิบ"
'   ElseIf Acc = "INVENTORY_TRANSFER_ADD" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถเพิ่มข้อมูลการโอนวัตถุดิบ"
'   ElseIf Acc = "INVENTORY_TRANSFER_EDIT" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถแก้ไขข้อมูลการโอนวัตถุดิบ"
'   ElseIf Acc = "INVENTORY_TRANSFER_DELETE" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถลบข้อมูลการโอนวัตถุดิบ"
'
'   ElseIf Acc = "INVENTORY_ADJUST" Then
'      Ri.RIGHT_ITEM_DESC = "ระบบข้อมูลการปรับยอดวัตถุดิบ"
'   ElseIf Acc = "INVENTORY_ADJUST_QUERY" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถค้นหาข้อมูลการปรับยอดวัตถุดิบ"
'   ElseIf Acc = "INVENTORY_ADJUST_ADD" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถเพิ่มข้อมูลการปรับยอดวัตถุดิบ"
'   ElseIf Acc = "INVENTORY_ADJUST_EDIT" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถแก้ไขข้อมูลการปรับยอดวัตถุดิบ"
'   ElseIf Acc = "INVENTORY_ADJUST_DELETE" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถลบข้อมูลการปรับยอดวัตถุดิบ"
'
'   ElseIf Acc = "INVENTORY_REPORT" Then
'      Ri.RIGHT_ITEM_DESC = "รายงานระบบข้อมูลคลัง"
'   ElseIf Mid(Acc, 1, 16) = "INVENTORY_REPORT" Then
'      Ri.RIGHT_ITEM_DESC = ReportName
'
'
'   ElseIf Acc = "PIG" Then
'      Ri.RIGHT_ITEM_DESC = "ระบบข้อมูลบริหารสุกร"
'
'   ElseIf Acc = "PIG_WEEK" Then
'      Ri.RIGHT_ITEM_DESC = "ระบบข้อมูลสัปดาห์เกิดสุกร"
'   ElseIf Acc = "PIG_WEEK_QUERY" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถค้นหาข้อมูลสัปดาห์เกิดสุกร"
'   ElseIf Acc = "PIG_WEEK_ADD" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถเพิ่มข้อมูลสัปดาห์เกิดสุกร"
'   ElseIf Acc = "PIG_WEEK_EDIT" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถแก้ไขข้อมูลสัปดาห์เกิดสุกร"
'   ElseIf Acc = "PIG_WEEK_DELETE" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถลบข้อมูลสัปดาห์เกิดสุกร"
'
'   ElseIf Acc = "PIG_IMPORT" Then
'      Ri.RIGHT_ITEM_DESC = "ระบบข้อมูลการนำเข้าสุกร"
'   ElseIf Acc = "PIG_IMPORT_QUERY" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถค้นหาข้อมูลการนำเข้าสุกร"
'   ElseIf Acc = "PIG_IMPORT_ADD" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถเพิ่มข้อมูลการนำเข้าสุกร"
'   ElseIf Acc = "PIG_IMPORT_EDIT" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถแก้ไขข้อมูลการนำเข้าสุกร"
'   ElseIf Acc = "PIG_IMPORT_DELETE" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถลบข้อมูลการนำเข้าสุกร"
'
'   ElseIf Acc = "PIG_BIRTH" Then
'      Ri.RIGHT_ITEM_DESC = "ระบบข้อมูลสุกรคลอด"
'   ElseIf Acc = "PIG_BIRTH_QUERY" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถค้นหาข้อมูลสุกรคลอด"
'   ElseIf Acc = "PIG_BIRTH_ADD" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถเพิ่มข้อมูลสุกรคลอด"
'   ElseIf Acc = "PIG_BIRTH_EDIT" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถแก้ไขข้อมูลสุกรคลอด"
'   ElseIf Acc = "PIG_BIRTH_DELETE" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถลบข้อมูลสุกรคลอด"
'
'   ElseIf Acc = "PIG_TRANSFER" Then
'      Ri.RIGHT_ITEM_DESC = "ระบบข้อมูลการโอนสุกร"
'   ElseIf Acc = "PIG_TRANSFER_QUERY" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถค้นหาข้อมูลการโอนสุกร"
'   ElseIf Acc = "PIG_TRANSFER_ADD" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถเพิ่มข้อมูลการโอนสุกร"
'   ElseIf Acc = "PIG_TRANSFER_EDIT" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถแก้ไขข้อมูลการโอนสุกร"
'   ElseIf Acc = "PIG_TRANSFER_DELETE" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถลบข้อมูลการโอนสุกร"
'
'   ElseIf Acc = "PIG_ADJUST" Then
'      Ri.RIGHT_ITEM_DESC = "ระบบข้อมูลการปรับยอดสุกร"
'   ElseIf Acc = "PIG_ADJUST_QUERY" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถค้นหาข้อมูลการปรับยอดสุกร"
'   ElseIf Acc = "PIG_ADJUST_ADD" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถเพิ่มข้อมูลการปรับยอดสุกร"
'   ElseIf Acc = "PIG_ADJUST_EDIT" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถแก้ไขข้อมูลการปรับยอดสุกร"
'   ElseIf Acc = "PIG_ADJUST_DELETE" Then
'      Ri.RIGHT_ITEM_DESC = "สามารถลบข้อมูลการปรับยอดสุกร"
'
'   ElseIf Acc = "PIG_REPORT" Then
'      Ri.RIGHT_ITEM_DESC = "รายงานระบบข้อมูลบริหารสุกร"
'   ElseIf Mid(Acc, 1, 10) = "PIG_REPORT" Then
'      Ri.RIGHT_ITEM_DESC = ReportName
'
'   Else
'      Ri.RIGHT_ITEM_DESC = ""
'   End If
'End Sub
'
'Private Function GetParentKey(Acc As String, TopFlag As Boolean) As String
'Dim I As Long
'Dim j As Long
'
'   For I = 1 To Len(Acc)
'      If Mid(Acc, I, 1) = "_" Then
'         j = I
'      End If
'   Next I
'
'   If j > 1 Then
'      GetParentKey = Mid(Acc, 1, j - 1)
'      TopFlag = False
'   Else
'      GetParentKey = ""
'      TopFlag = True
'   End If
'End Function
'
'Private Function CreatePermissionNode(Acc As String, ParentID As Long, ReportName As String) As Boolean
'Dim ParentKey As String
'Dim TopFlag As Boolean
'Dim TempParentID As Long
'Dim CreateFlag As Boolean
'Dim Ri As CRightItem
'Dim TempRs As ADODB.Recordset
'Dim iCount As Long
'
'   'Create node here
'   Set Ri = New CRightItem
'   Set TempRs = New ADODB.Recordset
'   TempParentID = 0
'
'   Ri.RIGHT_ID = -1
'   Ri.RIGHT_ITEM_NAME = Acc
'   Call Ri.QueryData(1, TempRs, iCount)
'   If TempRs.EOF Then
'      ParentKey = GetParentKey(Acc, TopFlag)
'      If Not TopFlag Then
'         Call CreatePermissionNode(ParentKey, TempParentID, ReportName)
'         Ri.PARENT_ID = TempParentID
'      End If
'
'      Ri.AddEditMode = SHOW_ADD
'      Call GetParentItemDesc(Acc, Ri, ReportName)
'      Call Ri.AddEditData
'      ParentID = Ri.RIGHT_ID
'   Else
'      Call Ri.PopulateFromRS(1, TempRs)
'      ParentID = Ri.RIGHT_ID
'   End If
'
'   If TempRs.State = adStateOpen Then
'      TempRs.Close
'   End If
'   Set TempRs = Nothing
'   Set Ri = Nothing
'End Function

Public Function VerifyAccessRight(Acc As String, Optional ReportName As String = "") As Boolean
'Dim R As CGroupRight
'Dim iCount As Long
'Dim TempParentID As Long
'Dim FoundFlag As Boolean
'
'   If glbUser.REAL_USER_ID = 0 Then
'      VerifyAccessRight = True
'      Exit Function
'   End If
'
'   Call glbDaily.StartTransaction
'   Call CreatePermissionNode(Acc, TempParentID, ReportName)
'   Call glbDaily.CommitTransaction
'
'   FoundFlag = False
'   For Each R In glbAccessRight
'      If R.RIGHT_ITEM_NAME = Acc Then
'         FoundFlag = True
'         If R.RIGHT_STATUS = "Y" Then
'            VerifyAccessRight = True
'            Exit For
'         Else
'            VerifyAccessRight = False
'            Exit For
'         End If
'      End If
'   Next R
'
'   If FoundFlag And (Not VerifyAccessRight) Then
'      VerifyAccessRight = False
'      glbErrorLog.LocalErrorMsg = "ไม่สามารถใช้งานโปรแกรมส่วนนี้ได้เนื่องจากมีสิทธ์ไม่พอเพียง -> " & Acc
'      glbErrorLog.ShowUserError
'   Else
'      VerifyAccessRight = True
'   End If
End Function

Public Function VerifyCombo(L As Label, C As ComboBox, Optional NullAllow As Boolean = False) As Boolean
Dim S As String
   If L Is Nothing Then
      S = ""
   Else
      S = L.Caption
   End If
   
   If Not NullAllow Then
      If Len(C.Text) = 0 Then
         VerifyCombo = False
         Call MsgBox("กรุณากรอกข้อมูล " & " '" & S & "' " & "ให้ถูกต้องและครบถ้วน ", vbOKOnly, PROJECT_NAME)
         If C.Enabled And C.Visible Then
            C.SetFocus
         End If
         Exit Function
      End If
   End If
   
   VerifyCombo = True
End Function

Public Function VerifyComboEx(C As ComboBox, Optional NullAllow As Boolean = False) As Boolean
Dim S As String
   
   If Not NullAllow Then
      If Len(C.Text) = 0 Then
         VerifyComboEx = False
         Call MsgBox("กรุณากรอกข้อมูล " & " '" & S & "' " & "ให้ถูกต้องและครบถ้วน ", vbOKOnly, PROJECT_NAME)
         If C.Enabled Then
            C.SetFocus
         End If
         Exit Function
      End If
   End If
   
   VerifyComboEx = True
End Function

Public Function VerifyItem(C As Collection, T As Object, Idx As Long) As Boolean
Dim i As Long
Dim Count As Long

   If C.Count <= 0 Then
      VerifyItem = True
      Exit Function
   End If
   
   For i = 1 To C.Count
      If C.ITEM(i).CURRENT_FLAG = "Y" Then
         Count = Count + 1
      End If
   Next i
   
   If Count <> 1 Then
      Call MsgBox("กรุณาเลือกข้อมูลให้มีค่าปัจจุบัน 1 รายการ", vbOKOnly, PROJECT_NAME)
   
      T.Tabs.ITEM(Idx).Selected = True
      VerifyItem = False
      Exit Function
   End If
   
   VerifyItem = True
End Function

Public Sub SetTextLenType(T As TextBox, TT As TEXT_BOX_TYPE, L As Long)
   If TT = TEXT_FLOAT_MONEY Or TT = TEXT_INTEGER_MONEY Then
      T.Alignment = 1
   End If
   
   T.Tag = TT
   T.MaxLength = L
End Sub

Public Function ChangeQuote(StrQ As String) As String
   ChangeQuote = Replace(StrQ, "'", "''")
End Function

Public Function NVLI(Value As Variant, i As Long) As Long
On Error Resume Next

   If IsNull(Value) Then
      NVLI = i
   Else
      NVLI = Value
   End If
End Function
Public Function NVLB(Value As Variant, i As Long) As Long
On Error Resume Next

   If IsNull(Value) Then
      NVLB = i
   Else
      NVLB = Value
   End If
End Function

Public Function NVLD(Value As Variant, i As Double) As Double
On Error Resume Next

   If IsNull(Value) Then
      NVLD = i
   Else
      NVLD = Value
   End If
End Function

Public Function NVLS(Value As Variant, S As String) As String
On Error Resume Next

   If IsNull(Value) Then
      NVLS = S
'   ElseIf IsEmpty(Value) Then
'      NVLS = S
   Else
      NVLS = Trim(Replace(Value, vbCrLf, ""))
   End If
End Function


Public Function EmptyToString(Value As String, S As String) As String
On Error Resume Next

   If Value = "" Then
      EmptyToString = S
   Else
      EmptyToString = Value
   End If
End Function

Public Function CryptString(strInput As String, strKey As String, action As Boolean)
Dim i As Integer, C As Integer
Dim strOutput As String

If Len(strKey) Then
    For i = 1 To Len(strInput)
        C = Asc(Mid$(strInput, i, 1))
        If action Then
            C = C + Asc(Mid$(strKey, (i Mod Len(strKey)) + 1, 1))
        Else: C = C - Asc(Mid$(strKey, (i Mod Len(strKey)) + 1, 1))
        End If
        strOutput = strOutput & Chr$(C And &HFF)
    Next i
Else
    strOutput = strInput
End If
CryptString = strOutput
End Function

Public Function EncryptText(PText As String) As String
   EncryptText = CryptString(PText, "GENETICOTHELLO", True)
End Function

Public Function DecryptText(CText As String) As String
   DecryptText = CryptString(CText, "GENETICOTHELLO", False)
End Function

Public Function EnableForm(Frm As Form, En As Boolean)
   If Frm Is Nothing Then
      Exit Function
   End If
   
   Frm.Enabled = En
   If En Then
      Screen.MousePointer = vbArrow
   Else
      Frm.Refresh
      DoEvents
      Screen.MousePointer = 11
   End If
End Function

Public Function IntToShortThaiMonth(M As Long) As String
   
   If M = 1 Then
         IntToShortThaiMonth = "ม.ค."
   ElseIf M = 2 Then
         IntToShortThaiMonth = "ก.พ."
   ElseIf M = 3 Then
         IntToShortThaiMonth = "มี.ค."
   ElseIf M = 4 Then
         IntToShortThaiMonth = "เม.ย."
   ElseIf M = 5 Then
         IntToShortThaiMonth = "พ.ค."
   ElseIf M = 6 Then
         IntToShortThaiMonth = "มิ.ย."
   ElseIf M = 7 Then
         IntToShortThaiMonth = "ก.ค."
   ElseIf M = 8 Then
         IntToShortThaiMonth = "ส.ค."
   ElseIf M = 9 Then
         IntToShortThaiMonth = "ก.ย."
   ElseIf M = 10 Then
         IntToShortThaiMonth = "ต.ค."
   ElseIf M = 11 Then
         IntToShortThaiMonth = "พ.ย."
   ElseIf M = 12 Then
         IntToShortThaiMonth = "ธ.ค."
   Else
         IntToShortThaiMonth = ""
   End If
End Function
Public Function IntToThaiMonth(M As Long) As String
   If glbParameterObj Is Nothing Then
      Exit Function
   End If
   
   If M = 1 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonth = "มกราคม"
      Else
         IntToThaiMonth = "January"
      End If
   ElseIf M = 2 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonth = "กุมภาพันธ์"
      Else
         IntToThaiMonth = "February"
      End If
      
   ElseIf M = 3 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonth = "มีนาคม"
      Else
         IntToThaiMonth = "March"
      End If
      
   ElseIf M = 4 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonth = "เมษายน"
      Else
         IntToThaiMonth = "April"
      End If
      
   ElseIf M = 5 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonth = "พฤษภาคม"
      Else
         IntToThaiMonth = "May"
      End If
      
   ElseIf M = 6 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonth = "มิถุนายน"
      Else
         IntToThaiMonth = "June"
      End If
      
   ElseIf M = 7 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonth = "กรกฎาคม"
      Else
         IntToThaiMonth = "July"
      End If
      
   ElseIf M = 8 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonth = "สิงหาคม"
      Else
         IntToThaiMonth = "August"
      End If
      
   ElseIf M = 9 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonth = "กันยายน"
      Else
         IntToThaiMonth = "September"
      End If
      
   ElseIf M = 10 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonth = "ตุลาคม"
      Else
         IntToThaiMonth = "October"
      End If
      
   ElseIf M = 11 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonth = "พฤศจิกายน"
      Else
         IntToThaiMonth = "November"
      End If
      
   ElseIf M = 12 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonth = "ธันวาคม"
      Else
         IntToThaiMonth = "December"
      End If
   Else
      IntToThaiMonth = ""
   End If
End Function

Public Function DateToStringMonthYearExt(D As Date) As String
   If D < 0 Then
      DateToStringMonthYearExt = ""
      Exit Function
   End If
   
   DateToStringMonthYearExt = " " & IntToThaiMonth(Month(D)) & " " & Format(Year(D) + 543, "0000")
End Function

Public Function DateToStringExt(D As Date) As String
   If D < 0 Then
      DateToStringExt = "-"
      Exit Function
   Else
      DateToStringExt = Day(D) & " " & IntToThaiMonth(Month(D)) & " " & Format(Year(D) + 543, "0000")
   End If
End Function

Public Function DateToStringExtEx(D As Date) As String
   If D < 0 Then
      DateToStringExtEx = ""
      Exit Function
   End If
   
   DateToStringExtEx = Day(D) & " " & IntToThaiMonth(Month(D)) & " " & Format(Year(D) + 543, "0000") & _
                     " " & Format(Hour(D), "00") & ":" & Format(Minute(D), "00") & ":" & Format(Second(D), "00")
End Function

Public Function DateToStringIntEx2(D As Date, Minute As Long, Second As Long) As String
   DateToStringIntEx2 = Format(Year(D), "0000") & "-" & Format(Month(D), "00") & "-" & Format(Day(D), "00") & " " & _
   Format(Minute, "00") & ":" & Format(Second, "00") & ":00"
End Function

Public Function DateToStringExtEx2(D As Date) As String
   If D > 0 Then
      DateToStringExtEx2 = Format(Day(D), "00") & "/" & Format(Month(D), "00") & "/" & Format(Year(D) + 543, "0000")
   Else
      DateToStringExtEx2 = ""
   End If
End Function

Public Function DateToStringExtEx3(D As Date) As String
   If D > 0 Then
      DateToStringExtEx3 = Format(Day(D), "00") & "/" & Format(Month(D), "00") & "/" & Format(Year(D) + 543, "0000")
      DateToStringExtEx3 = DateToStringExtEx3 & " " & Format(Hour(D), "00") & ":" & Format(Minute(D), "00") & ":" & Format(Second(D), "00")
   Else
      DateToStringExtEx3 = ""
   End If
End Function

Public Function DateToStringIntEx3(D As Date) As String
   DateToStringIntEx3 = Format(Year(D), "0000") & "-" & Format(Month(D), "00") & "-" & Format(Day(D), "00")
End Function

Public Function InternalDateToStringEx4(D As String) As String
Dim T As Date
   T = InternalDateToDate(D)
   If T > 0 Then
      InternalDateToStringEx4 = Format(Day(T), "00") & "/" & Format(Month(T), "00") & "/" & Format(Year(T) + 543, "0000")
   Else
      InternalDateToStringEx4 = ""
   End If
End Function

Public Function DateToStringInt(D As Date) As String
   If D = -1 Then
      DateToStringInt = "99999999"
   ElseIf D = -2 Then
      DateToStringInt = "00000000"
   Else
      DateToStringInt = Format(Year(D), "0000") & Format(Month(D), "00") & Format(Day(D), "00")
   End If
End Function
Public Function DateToStringIntEndMonth(D As Date) As String
   DateToStringIntEndMonth = Format(Year(D), "0000") & "-" & Format(Month(D), "00") & "-31" & _
                     " 00:00:00"
End Function

Public Function DateToStringIntEx(D As Date) As String
   DateToStringIntEx = Format(Year(D), "0000") & "-" & Format(Month(D), "00") & "-" & Format(Day(D), "00") & _
                     " 23:59:59"
End Function

Public Function DateToStringIntHi(D As Date) As String
   If D > 0 Then
      DateToStringIntHi = Format(Year(D), "0000") & Format(Month(D), "00") & Format(Day(D), "00")
   Else
      DateToStringIntHi = "99999999"
   End If
End Function
Public Function DateToString(D As Date) As String
   If D > 0 Then
      DateToString = Format(Day(D), "00") & "/" & Format(Month(D), "00") & "/" & Format(Year(D) + 543, "0000")
   Else
      DateToString = "99999999"
   End If
End Function

Public Function DateToStringIntLow(D As Date) As String
   If D = -1 Then
      DateToStringIntLow = "99999999"
   ElseIf D = -2 Then
      DateToStringIntLow = "00000000"
   Else
      DateToStringIntLow = Format(Year(D), "0000") & Format(Month(D), "00") & Format(Day(D), "00")
   End If
End Function

Public Function InternalDateToDate(IntDate As String) As Date
Dim DStr As Long
Dim D As Long
Dim MStr As String
Dim M As Long
Dim YStr As String
Dim Y As Long

Dim HHStr As Long
Dim HH As Long
Dim MMStr As String
Dim MM As Long
Dim SSStr As String
Dim SS As Long

   If (IntDate = "") Or (IntDate = "99999999") Then
      InternalDateToDate = -1
      Exit Function
   End If
   
   If (IntDate = "") Or (IntDate = "00000000") Then
      InternalDateToDate = -2
      Exit Function
   End If
   
   If Len(IntDate) < 8 Then
      InternalDateToDate = Now
      Exit Function
   End If
   
   YStr = Mid(IntDate, 1, 4)
   MStr = Mid(IntDate, 5, 2)
   DStr = Mid(IntDate, 7, 2)
         
   HH = 0
   MM = 0
   SS = 0
   
   Y = Val(YStr)
   M = Val(MStr)
   D = Val(DStr)
   
   InternalDateToDate = DateSerial(Y, M, D) + TimeSerial(HH, MM, SS)
End Function

Public Function InternalDateToDateEx(IntDate As String) As Date
Dim DStr As Long
Dim D As Long
Dim MStr As String
Dim M As Long
Dim YStr As String
Dim Y As Long

Dim HHStr As Long
Dim HH As Long
Dim MMStr As String
Dim MM As Long
Dim SSStr As String
Dim SS As Long

   If (IntDate = "") Or (IntDate = "9999-99-99 99:99:99") Then
      InternalDateToDateEx = -1
      Exit Function
   End If
   
   If (IntDate = "") Or (IntDate = "0000-00-00 00:00:00") Then
      InternalDateToDateEx = -1
      Exit Function
   End If
   
   If Len(IntDate) < 8 Then
      InternalDateToDateEx = Now
      Exit Function
   End If
   
   YStr = Mid(IntDate, 1, 4)
   MStr = Mid(IntDate, 5, 2)
   DStr = Mid(IntDate, 7, 2)
   
'   If Not IsNumeric(YStr) Then
'      Exit Function
'   End If
'
'   If Not IsNumeric(MStr) Then
'      Exit Function
'   End If
'
'   If Not IsNumeric(DStr) Then
'      Exit Function
'   End If
   
   HHStr = "00"
   MMStr = "00"
   SSStr = "00"
   
'   If Not IsNumeric(HHStr) Then
'      Exit Function
'   End If
'
'   If Not IsNumeric(MMStr) Then
'      Exit Function
'   End If
'
'   If Not IsNumeric(SSStr) Then
'      Exit Function
'   End If
   
   HH = Val(HHStr)
   MM = Val(MMStr)
   SS = Val(SSStr)
   
   Y = Val(YStr)
   M = Val(MStr)
   D = Val(DStr)
   
   InternalDateToDateEx = Y & "/" & M & "/" & D
End Function

Public Function InternalDateToDateEx2(IntDate As String) As Date
Dim DStr As Long
Dim D As Long
Dim MStr As String
Dim M As Long
Dim YStr As String
Dim Y As Long

Dim HHStr As Long
Dim HH As Long
Dim MMStr As String
Dim MM As Long
Dim SSStr As String
Dim SS As Long

   If (IntDate = "") Or (IntDate = "9999-99-99 99:99:99") Then
      InternalDateToDateEx2 = -1
      Exit Function
   End If
   
   If (IntDate = "") Or (IntDate = "0000-00-00 00:00:00") Then
      InternalDateToDateEx2 = -1
      Exit Function
   End If
   
   If Len(IntDate) < 10 Then
      InternalDateToDateEx2 = Now
      Exit Function
   End If
   
   YStr = Mid(IntDate, 1, 4)
   MStr = Mid(IntDate, 6, 2)
   DStr = Mid(IntDate, 9, 2)
      
   HHStr = "00"
   MMStr = "00"
   SSStr = "00"
      
   HH = Val(HHStr)
   MM = Val(MMStr)
   SS = Val(SSStr)
   
   Y = Val(YStr)
   M = Val(MStr)
   D = Val(DStr)
   
   InternalDateToDateEx2 = DateSerial(Y, M, D) + TimeSerial(HH, MM, SS)
End Function

Public Function ReFormatDate(DStr As String) As String
Dim YYYY As String
Dim MM As String
Dim dd As String

   YYYY = Mid(DStr, 5, 4)
   MM = Mid(DStr, 3, 2)
   dd = Mid(DStr, 1, 2)
   
   ReFormatDate = YYYY & MM & dd
End Function
Public Function ReFormatDate2(DStr As String) As String
Dim YYYY As String
Dim MM As String
Dim dd As String

   YYYY = Mid(DStr, 1, 4)
   MM = Mid(DStr, 5, 2)
   dd = Mid(DStr, 7, 2)
   
   ReFormatDate2 = dd & "/" & MM & "/" & YYYY
End Function
Public Sub InitTextBox(T As TextBox, Msg As String, Optional Password As String = "")
   T.PasswordChar = Password
   T.FontSize = 12
   T.FontName = "MS Sans Serif"
   T.Text = Msg
   T.BackColor = GLB_GRID_COLOR
   'T.FontBold = True
End Sub

Public Sub ShowTotalLabel(L As Label, Value As Long)
   L.Caption = "รวม = " & Value
End Sub

Public Sub ClearTreeView(ByVal tvHwnd As Long)
Dim lNodeHandle As Long

    'Turn off redrawing on the Treeview for more speed improvements
    SendMessageLong tvHwnd, WM_SETREDRAW, False, 0

    Do
        lNodeHandle = SendMessageLong(tvHwnd, TVM_GETNEXTITEM, TVGN_ROOT, 0)
         If lNodeHandle > 0 Then
            SendMessageLong tvHwnd, TVM_DELETEITEM, 0, lNodeHandle
         Else
            Exit Do
         End If
    Loop

    'Turn on redrawing on the Treeview
    SendMessageLong tvHwnd, WM_SETREDRAW, True, 0
End Sub

Public Sub InitCombo(C As ComboBox)
   C.FontSize = 12
   C.FontName = "MS Sans Serif"
   C.BackColor = GLB_GRID_COLOR
End Sub

Public Function VerifyGrid(S As String) As Boolean
   If S = "" Then
      VerifyGrid = False
      glbErrorLog.LocalErrorMsg = "กรุณาเลือกข้อมูลที่ต้องการก่อน"
      glbErrorLog.ShowUserError
   Else
      VerifyGrid = True
   End If
End Function

Public Function ConfirmDelete(S As String) As Boolean
   glbErrorLog.LocalErrorMsg = "ท่านต้องการจะลบข้อมูล " & S & "' ใช่หรือไม่"
   If glbErrorLog.AskMessage = vbNo Then
      ConfirmDelete = False
      Exit Function
   Else
      ConfirmDelete = True
   End If
End Function

Public Function ConfirmSaveExist(S1 As String, S2 As String) As Boolean
   glbErrorLog.LocalErrorMsg = "ท่านต้องการจะบันทึกข้อมูล " & S1 & " แทนที่ข้อมูล " & S2 & "' หรือไม่"
   If glbErrorLog.AskMessage = vbNo Then
      ConfirmSaveExist = False
      Exit Function
   Else
      ConfirmSaveExist = True
   End If
End Function

Public Function ConfirmSaveRecord(S As String) As Boolean
   glbErrorLog.LocalErrorMsg = "ท่านต้องการจะบันทึกข้อมูล " & S & "' หรือไม่"
   If glbErrorLog.AskMessage = vbNo Then
      ConfirmSaveRecord = False
      Exit Function
   Else
      ConfirmSaveRecord = True
   End If
End Function

Public Sub InitFormHeader(L As Label, Caption As String)
   L.Caption = Caption
   L.FontBold = True
   L.FontSize = 20
   L.FontName = GLB_FONT
   L.Alignment = 2
   L.ForeColor = RGB(0, 10, 0)
End Sub

Public Sub InitDialogHeader(L As Label, Caption As String)
   L.Caption = Caption
   L.FontBold = True
   L.FontSize = 16
   L.FontName = GLB_FONT
   L.Alignment = 2
End Sub

Public Sub InitNormalLabel(L As Label, Caption As String, Optional Color As Long = 0)
   L.Caption = Caption
   L.FontBold = False
   L.FontSize = 14
   L.FontBold = True
   L.FontName = GLB_FONT
   L.BackStyle = 0
   L.ForeColor = Color
End Sub

Public Sub InitOption(O As OptionButton, Caption As String)
   O.Caption = Caption
   O.FontSize = 14
   O.FontBold = True
   O.FontName = GLB_FONT
   O.BackColor = GLB_FORM_COLOR
End Sub

Public Sub InitOptionEx(O As SSOption, Caption As String)
   O.Caption = Caption
   O.Font.Size = 14
   O.Font.Bold = True
   O.Font.Name = GLB_FONT
   O.BackColor = GLB_FORM_COLOR
   O.BackStyle = ssTransparent
End Sub

Public Sub InitCheckBox(C As SSCheck, Caption As String)
   C.Caption = Caption
   C.FontSize = 14
   C.FontBold = True
   C.FontName = GLB_FONT
   C.BackColor = GLB_FORM_COLOR
   C.BackStyle = ssTransparent
   C.TripleState = True
End Sub

Public Sub InitMainButton(B As SSCommand, Caption As String, Optional Color As Double = &HFFFFFF)
   B.Caption = Caption
   B.Font.Bold = True
   B.Font.Size = 14
   B.Font.Name = GLB_FONT
   B.Font3D = ssInsetLight
   B.BackColor = RGB(255, 255, 255)
   B.ButtonStyle = ssWin95 '= ssActiveBorders
   B.MousePointer = ssCustom
   B.MouseIcon = LoadPicture(glbParameterObj.ButtonCursor)
End Sub

Public Sub InitHeaderFooter(H As SSPanel, F As SSPanel)
'   H.PICTURE = LoadPicture("D:\Picture\WINPricing100\header.gif")
   If Not (F Is Nothing) Then
'      F.PICTURE = LoadPicture("D:\Picture\WINPricing100\footer.gif")
   End If
End Sub

Public Sub InitMainButtonOld(B As CommandButton, Caption As String, Optional Color As Double = &HFFFFFF)
   B.Caption = Caption
   B.Font.Bold = True
   B.Font.Size = 14
   B.Font.Name = GLB_FONT
   B.BackColor = GLB_FORM_COLOR
End Sub

Public Sub SetSelect(T As TextBox)
   T.SelStart = 0
   T.SelLength = Len(T.Text)
End Sub

Public Sub InitDialogButton(B As CommandButton, Caption As String)
   B.Caption = Caption
   B.FontBold = True
   B.FontSize = 14
   B.FontName = GLB_FONT
   
   B.BackColor = &HFFFFFF
End Sub

Public Sub ReleaseAll()
   Set glbErrorLog = Nothing
   Set glbDatabaseMngr = Nothing
   Set glbParameterObj = Nothing
   Set glbUser = Nothing
   Set glbGroup = Nothing
   Set m_SupplierColl = Nothing
'   Set glbHR = Nothing
'   Set glbAccessRight = Nothing
End Sub

Public Sub SetEnableDisableTextBox(T As TextBox, En As Boolean)
   If En Then
      T.Enabled = True
      T.BackColor = GLB_GRID_COLOR
   Else
      T.Enabled = False
      T.BackColor = &H8000000F
   End If
End Sub

Public Sub SetEnableDisableComboBox(T As ComboBox, En As Boolean)
   If En Then
      T.Enabled = True
      T.BackColor = GLB_GRID_COLOR
   Else
      T.Enabled = False
      T.BackColor = &H8000000F
   End If
End Sub

Public Sub SetEnableDisableButton(B As SSCommand, En As Boolean)
   If En Then
      B.Enabled = True
      B.BackColor = GLB_GRID_COLOR
   Else
      B.Enabled = False
      B.BackColor = &H8000000F
   End If
End Sub

Public Function ConfirmExit(HasEdit As Boolean) As Boolean
   If Not HasEdit Then
      ConfirmExit = True
   Else
      glbErrorLog.LocalErrorMsg = "ท่านต้องการจะออกจากโปรแกรมโดยไม่มีการบันทึกข้อมูลใช่หรือไม่"
      If glbErrorLog.AskMessage = vbYes Then
         ConfirmExit = True
      Else
         ConfirmExit = False
      End If
   End If
End Function

Public Function DuplicateData() As Boolean
'   If Not HasEdit Then
'      DuplicateData = True
'   Else
      glbErrorLog.LocalErrorMsg = "IV และ สินค้ารหัสนี้มีในระบบแล้ว ไม่สามารถบันทึกข้อมูลซ้ำได้"
      If glbErrorLog.AskMessage = vbYes Then
         DuplicateData = True
      Else
         DuplicateData = False
      End If
'   End If
End Function


Public Function ThaiBaht(ByVal pamt As Double) As String
Dim valstr As String, vLen As Integer, vno As Integer, syslge As String
Dim i As Integer, j As Integer, v As Integer
Dim wnumber(10) As String, wdigit(10) As String, spcdg(5) As String
Dim vword(20) As String

 If pamt <= 0# Then
   ThaiBaht = ""
   Exit Function
 End If
 valstr = Trim(Format$(pamt, "##########0.00"))
 vLen = Len(valstr) - 3
 For i = 1 To 20
     vword(i) = ""
 Next i
wnumber(1) = "หนึ่ง": wnumber(2) = "สอง": wnumber(3) = "สาม": wnumber(4) = "สี่"
wnumber(5) = "ห้า": wnumber(6) = "หก": wnumber(7) = "เจ็ด": wnumber(8) = "แปด"
wnumber(9) = "เก้า": wdigit(1) = "บาท": wdigit(2) = "สิบ": wdigit(3) = "ร้อย": wdigit(4) = "พัน"
wdigit(5) = "หมื่น": wdigit(6) = "แสน": wdigit(7) = "ล้าน": spcdg(1) = "สตางค์": spcdg(2) = "เอ็ด"
spcdg(3) = "ยี่": spcdg(4) = "ถ้วน"
For i = 1 To vLen
    vno = Int(Val(Mid$(valstr, i, 1)))
    If vno = 0 Then
        vword(i) = ""
        If (vLen - i + 1) = 7 Then
            vword(i) = wdigit(7)             '--ล้าน
        End If
    Else
        If (vLen - i + 1) > 7 Then
            j = vLen - i - 5               '--เกินหลักล้าน
        Else
            j = vLen - i + 1               '--หลักแสน
        End If
        vword(i) = wnumber(vno) + wdigit(j) '-30ถึง90
        If vno = 1 And j = 2 Then
            vword(i) = wdigit(2)             '--สิบ
        End If
        If vno = 2 And j = 2 Then
            vword(i) = spcdg(3) + wdigit(j)  '--ยี่สิบ
        End If
        If j = 1 Then                       ' สิยเอ็ค -->เก้าสิบเอ็ด
            vword(i) = wnumber(vno)
            If vno = 1 And vLen > 1 Then
                If Mid$(valstr, i - 1, 1) <> "0" Then
                    vword(i) = spcdg(2)
                End If
            End If
        End If
        If j = 7 Then         '-แก้บักกรณี 11,111,111.00 สิบเอ็ด
            vword(i) = wnumber(vno) + wdigit(j)   '-ล้าน
            If vno = 1 And vLen > 7 Then
                If Mid$(valstr, i - 1, 1) <> "0" Then
                    vword(i) = spcdg(2) + wdigit(j)
                End If
            End If
        End If
    End If
Next i
    
If Int(pamt) > 0 Then
       vword(vLen) = vword(vLen) + wdigit(1)
End If
 '--------------ทศนิยม --------------
valstr = Mid$(valstr, vLen + 2, 2)
vLen = Len(valstr)
For i = 1 To vLen
    vno = Int(Val(Mid$(valstr, i, 1)))
    If vno = 0 Then
           vword(i + 10) = ""
    Else
           j = vLen - i + 1
           vword(i + 10) = wnumber(vno) + wdigit(j)
        If vno = 1 And j = 2 Then
              vword(i + 10) = wdigit(2)
        End If
        If vno = 2 And j = 2 Then
              vword(i + 10) = spcdg(3) + wdigit(j)
        End If
        If j = 1 Then
            If vno = 1 And Int(Val(Mid$(valstr, i - 1, 1))) <> 0 Then
                 vword(i + 10) = spcdg(2)
            Else
                 vword(i + 10) = wnumber(vno)
            End If
        End If
    End If
Next i
If pamt <> 0 Then
    If Val(valstr) = 0 Then
        vword(13) = spcdg(4)
    Else
        vword(13) = spcdg(1)
    End If
End If

 '*** เผื่อใช้กรณียาวมาก และต้องการตัดประโยค
 valstr = ""
 For i = 1 To 20
    'IF LEN(valstr) < 70 AND LEN(valstr + vword(i)) > 70 Then
    '   valstr = valstr + REPLICATE(" ",70 - LEN(valstr))
    'END IF
    valstr = valstr + vword(i)
 Next i
 'valstr='('+valstr+')'
 ThaiBaht = (valstr)
End Function

Public Function WildCard(WStr As String, SubLen As Long, NewStr As String) As Boolean
Dim Tmp As String
   Tmp = Trim(WStr)
   If Tmp = "" Then
      WildCard = False
      Exit Function
   End If
   
   If Mid(Tmp, Len(Tmp)) = "%" Then
      SubLen = Len(Tmp) - 1
      NewStr = Mid(Tmp, 1, SubLen)
      
      WildCard = True
   Else
      WildCard = False
   End If
End Function

Public Function FormatString(S As String, Patch As String, L As Long) As String
Dim Temp As String
Dim Start As Long
Dim i As Long
Dim j As Long

   Temp = Space(L)
   Call Replace(Temp, " ", Patch)
   j = 0
   Start = (L - Len(S)) \ 2
   
   For i = 1 To L
      If i < Start Then
         Mid(Temp, i) = Patch
      Else
         If i > Start + Len(S) Then
            Mid(Temp, i) = Patch
         Else
            j = j + 1
            Mid(Temp, i) = Mid(S, j)
         End If
      End If
   Next i
   
   FormatString = Temp
End Function
Public Function FormatNumber(N As Variant, Optional ZeroString As String = "0.00") As String
Dim T As Double

   If IsNull(N) Then
      T = 0
   Else
      T = Val(N)
   End If
   
   If T = 0 Then
      FormatNumber = ZeroString
   ElseIf T > 0 Then
      FormatNumber = Format(T, "#,##0.00")
   ElseIf T < 0 Then
      FormatNumber = "(" & Format(-1 * T, "#,##0.00") & ")"
   End If
End Function

Public Function FormatNumberInt(N As Variant, Optional ZeroString As String = "0") As String
Dim T As Double

   If IsNull(N) Then
      T = 0
   Else
      T = N
   End If
   
   If T = 0 Then
      FormatNumberInt = ZeroString
   ElseIf T > 0 Then
      FormatNumberInt = Format(T, "#,##0")
   ElseIf T < 0 Then
      FormatNumberInt = "(" & Format(-1 * T, "#,##0") & ")"
   End If
End Function
Public Function FormatNumberToNull(N As Variant, Optional DecimalPoint As Long = 2, Optional Quat As Boolean = True, Optional ZeroString As String = "") As String
Dim T As Double
Dim TempStr As String
Dim i As Long

   TempStr = "."
   For i = 1 To DecimalPoint
      TempStr = TempStr & "0"
   Next i
   If DecimalPoint = 0 Then
       TempStr = ""
   End If
   
   If IsNull(N) Then
      T = 0
   Else
      T = N
   End If
   
   If T = 0 Then
      If ZeroString = "0" Then
         FormatNumberToNull = ZeroString & TempStr
      Else
         FormatNumberToNull = ZeroString
      End If
   ElseIf Quat Then
      FormatNumberToNull = Format(T, "#,##0" & TempStr)
   Else
      FormatNumberToNull = Format(T, "0" & TempStr)
   End If
End Function

Public Function ReverseFormatNumber(N As String) As Double
   ReverseFormatNumber = Val(Replace(N, ",", ""))
End Function

Public Function IDToListIndex(Cbo As ComboBox, ID As Long) As Long
Dim i As Long
Dim Temp As String

   IDToListIndex = -1
   For i = 0 To Cbo.ListCount - 1
      If InStr(Cbo.ItemData(i), ":") <= 0 Then
         Temp = Cbo.ItemData(i)
      Else
         Temp = Mid(Cbo.ItemData(i), 1, InStr(Cbo.ItemData(i), ":") - 1)
      End If
      If Temp = ID Then
         IDToListIndex = i
      End If
   Next i
End Function

Public Sub Main()
On Error GoTo ErrorHandler
Dim i As Long

   GLB_GRID_COLOR = RGB(255, 255, 250)
   GLB_NORMAL_COLOR = RGB(0, 0, 0)
   GLB_ALERT_COLOR = RGB(255, 0, 0)
   GLB_FORM_COLOR = RGB(180, 200, 200)
   GLB_HEAD_COLOR = GLB_FORM_COLOR
   GLB_GRIDHD_COLOR = RGB(149, 194, 240)
   GLB_SHOW_COLOR = RGB(0, 0, 240)
   GLB_MANDATORY_COLOR = RGB(0, 0, 255)

   Set glbSetting = New clsGlobalSetting
   Set glbParameterObj = New clsParameter
   Set glbUser = New CUser
   Set glbGroup = New CGroup
   Set m_SupplierColl = New Collection
   
   Set glbErrorLog = New clsErrorLog
   glbErrorLog.DayKeepLog = 10
   glbErrorLog.LogFileMode = LOG_CURRENT_DATE
   
   glbErrorLog.ModuleName = MODULE_NAME
   glbErrorLog.RoutineName = "Main"
   glbErrorLog.MsgBoxTitle = PROJECT_NAME
   
   If App.PrevInstance = True Then
      glbErrorLog.LocalErrorMsg = "โปรแกรมเดิมได้ถูกรันก่อนหน้านี้แล้ว"
      glbErrorLog.ShowUserError

      Set glbErrorLog = Nothing
      Exit Sub
   End If
   
   Load frmSplash
   frmSplash.Show 0
   frmSplash.Refresh
   
   Set glbDatabaseMngr = New clsDatabaseMngr
   If Not glbDatabaseMngr.ConnectDatabase(glbParameterObj.DBFile, glbParameterObj.UserName, glbParameterObj.Password, glbErrorLog) Then
'      frmDBSetting.UserName = glbParameterObj.UserName
'      frmDBSetting.Password = glbParameterObj.Password
'      frmDBSetting.FileDb = glbParameterObj.DBFile
'      frmDBSetting.Header = " ไม่สามารถเชื่อต่อฐานข้อมูลได้ "
'
'      Load frmDBSetting
'      frmDBSetting.Show 1
'      If frmDBSetting.OKClick Then
'         glbParameterObj.UserName = frmDBSetting.UserName
'         glbParameterObj.Password = frmDBSetting.Password
'         glbParameterObj.DBFile = frmDBSetting.FileDb
'      Else
'         Unload frmDBSetting
'         Set frmDBSetting = Nothing
'
'         Unload frmSplash
'         Set frmSplash = Nothing
'
'         Call ReleaseAll
'         End
'      End If
'      Unload frmDBSetting
'      Set frmDBSetting = Nothing
   End If
   
   If Not glbDatabaseMngr.ConnectLegacyDatabase(glbParameterObj.DBConfigFile, glbParameterObj.UserName, glbParameterObj.Password, glbErrorLog) Then
      'debug.print "Error"
   End If
   glbParameterObj.ComputerName = glbDatabaseMngr.GetComputerName
   
'   If Not glbDatabaseMngr.ConnectAgentServer(glbParameterObj.LicenseIP, glbParameterObj.LicensePort, glbErrorLog) Then
'      frmAgentSetting.Port = glbParameterObj.LicensePort
'      frmAgentSetting.IP = glbParameterObj.LicenseIP
'      frmAgentSetting.Header = " ไม่สามารถเชื่อมต่อกับไลเซนส์เซิร์ฟเวอร์ได้ "
'
'      Load frmAgentSetting
'      frmAgentSetting.Show 1
'
'      If frmAgentSetting.OKClick Then
'         glbParameterObj.LicenseIP = frmAgentSetting.IP
'         glbParameterObj.LicensePort = frmAgentSetting.Port
'      Else
'         Unload frmAgentSetting
'         Set frmAgentSetting = Nothing
'
'         Unload frmSplash
'         Set frmSplash = Nothing
'
'         Call ReleaseAll
'         End
'      End If
'      Unload frmAgentSetting
'      Set frmAgentSetting = Nothing
'   End If
   Unload frmSplash
   Set frmSplash = Nothing
   
   Set glbDaily = New clsDaily
'   Set glbAdmin = New clsAdmin
'   Set glbMaster = New clsMaster
'   Set glbLegacy = New clsLegacy
'   Set glbLoginTracking = New CLoginTracking
'   Set glbEnterPrise = New CEnterprise
   Set glbAccessRight = New Collection
   
   Load frmLedgerReportMain
   frmLedgerReportMain.Show

   Exit Sub
   
ErrorHandler:
   If glbErrorLog Is Nothing Then
      MsgBox Err.DESCRIPTION
   Else
      glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   End If
   
End Sub

Public Sub InitOrderType(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("น้อยไปมาก"))
   C.ItemData(1) = 1
   
   C.AddItem (MapText("มากไปน้อย"))
   C.ItemData(2) = 2
End Sub

Public Function GetItem(Col As Collection, Idx As Long, RealIndex As Long) As Object
Dim i As Long
Dim Count As Long

   Count = 0
   For i = 1 To Col.Count
      If Col.ITEM(i).Flag <> "D" Then
         Count = Count + 1
      End If
      If Count = Idx Then
         RealIndex = i
         Set GetItem = Col.ITEM(i)
         Exit Function
      End If
   Next i
   
   Set GetItem = Nothing
End Function

Public Function CountItem(Col As Collection) As Long
Dim i As Long
Dim Count As Long

   Count = 0
   For i = 1 To Col.Count
      If Col.ITEM(i).Flag <> "D" Then
         Count = Count + 1
      End If
   Next i
   
   CountItem = Count
End Function

Public Function VSP_CalTable(ByVal pRaw As String, ByVal pWidth As Long, ByRef pPer() As Long) As String
On Error GoTo ErrorHandler
Dim strTemp As String
Dim i As Long
Dim Count As Long
Dim iPer As Long
Dim tPer As Long
Dim TOTAL As Long
Dim Prefix() As String
Dim Value() As Long
Dim iTemp As Long
   
   pRaw = Trim$(pRaw)
   If Len(pRaw) <= 0 Then
      VSP_CalTable = ""
      Exit Function
   End If
   Count = 0
   iPer = 1
   TOTAL = 0
   strTemp = ""
   While iPer <= Len(pRaw)
      If Val(Mid$(pRaw, iPer, 1)) <= 0 Then
         strTemp = strTemp & Mid$(pRaw, iPer, 1)
      Else
         Count = Count + 1
         ReDim Preserve Prefix(Count)
         ReDim Preserve Value(Count)
         Prefix(Count) = strTemp
         tPer = InStr(iPer, pRaw, "|")
         If tPer <= 0 Then tPer = InStr(iPer, pRaw, ";")

         Value(Count) = Val(Mid$(pRaw, iPer, tPer - iPer))
         TOTAL = TOTAL + Value(Count)
         iPer = tPer
         strTemp = ""
      End If
      iPer = iPer + 1
   Wend
   strTemp = ""
   ReDim pPer(Count)
   For i = 1 To Count - 1
      iTemp = CLng((Value(i) * pWidth) / TOTAL)
      strTemp = strTemp & Trim$(Prefix(i)) & Trim$(Str$(iTemp)) & "|"
      If i = 1 Then
         pPer(i - 1) = iTemp
      Else
         pPer(i - 1) = pPer(i - 2) + iTemp
      End If
   Next i
   strTemp = strTemp & Trim$(Prefix(i)) & CLng(((Value(i) * pWidth) / TOTAL)) & ";"
   If i > 1 Then
      iTemp = CLng((Value(i) * pWidth) / TOTAL)
      pPer(i - 1) = pPer(i - 2) + iTemp
   End If
   VSP_CalTable = strTemp

   Exit Function
ErrorHandler:
   glbErrorLog.LocalErrorMsg = "Runtime error."
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Function

Public Function Check2Flag(A As Long) As String
   If A = ssCBChecked Then
      Check2Flag = "Y"
   Else
      Check2Flag = "N"
   End If
End Function

Public Function CheckUniqueNs(UnqType As UNIQUE_TYPE, KEY As String, ID As Long, Optional TempID As Long = -1, Optional dbType As Long = 1) As Boolean
On Error GoTo ErrorHandler
Dim TableName As String
Dim FieldName1 As String
Dim FieldName2 As String
Dim Flag As Boolean
Dim Count As Long

   CheckUniqueNs = False
'
'   TEACHER_UNIQUE = 16
'   SUBJECT_UNIQUE = 17
'   FACULTY_UNIQUE = 18
   
   Flag = False
   If UnqType = TEACHER_UNIQUE Then
      TableName = "TEACHER"
      FieldName1 = "TEACHER_CODE"
      FieldName2 = "TEACHER_ID"
      Flag = True
   ElseIf UnqType = USERGROUP_UNIQUE Then
      TableName = "USER_GROUP"
      FieldName1 = "GROUP_NAME"
      FieldName2 = "GROUP_ID"
      Flag = True
   ElseIf UnqType = SUBJECT_UNIQUE Then
      TableName = "SUBJECT"
      FieldName1 = "SUBJECT_CODE"
      FieldName2 = "SUBJECT_ID"
      Flag = True
   ElseIf UnqType = PRDFEATURE_UNIQUE Then
      TableName = "PRDFEATURE_NAME"
      FieldName1 = "PRODUCT_CODE"
      FieldName2 = "PRDFEATURE_NAME_ID"
      Flag = True
   ElseIf UnqType = FACULTY_UNIQUE Then
      TableName = "FACULTY"
      FieldName1 = "FACULTY_CODE"
      FieldName2 = "FACULTY_ID"
      Flag = True
   ElseIf UnqType = DBN_UNIQUE Then
      TableName = "BILL"
      FieldName1 = "BILL_NO"
      FieldName2 = "BILL_ID"
      Flag = True
   ElseIf UnqType = EMPCODE_UNIQUE Then
      TableName = "EMPLOYEE"
      FieldName1 = "EMP_CODE"
      FieldName2 = "EMP_ID"
      Flag = True
   ElseIf UnqType = USERNAME_UNIQUE Then
      TableName = "USER_ACCOUNT"
      FieldName1 = "USER_NAME"
      FieldName2 = "USER_ID"
      Flag = True
   ElseIf UnqType = REPAIR_UNIQUE Then
      TableName = "REPAIR_DATA"
      FieldName1 = "REPAIR_NUM"
      FieldName2 = "REPAIR_ID"
      Flag = True
   ElseIf UnqType = IMPORT_UNIQUE Then
      TableName = "INVENTORY_DOC"
      FieldName1 = "DOCUMENT_NO"
      FieldName2 = "INVENTORY_DOC_ID"
      Flag = True
   ElseIf UnqType = EXPORT_UNIQUE Then
      TableName = "INVENTORY_DOC"
      FieldName1 = "DOCUMENT_NO"
      FieldName2 = "INVENTORY_DOC_ID"
      Flag = True
   ElseIf UnqType = REPAIR_FORMULA_UNIQUE Then
      TableName = "REPAIR_FORMULA"
      FieldName1 = "FORMULA_CODE"
      FieldName2 = "FORMULA_ID"
      Flag = True
   ElseIf UnqType = SUPPLIER_UNIQUE Then
      TableName = "SUPPLIER"
      FieldName1 = "SUPPLIER_CODE"
      FieldName2 = "SUPPLIER_ID"
      Flag = True
   ElseIf UnqType = PARTNO_UNIQUE Then
      TableName = "PART_ITEM"
      FieldName1 = "PART_NO"
      FieldName2 = "PART_ITEM_ID"
      Flag = True
   ElseIf UnqType = QUOATATION_UNIQUE Then
      TableName = "QUOATATION"
      FieldName1 = "QUOATATION_NO"
      FieldName2 = "QUOATATION_ID"
      Flag = True
   ElseIf UnqType = EXPENSE_UNIQUE Then
      TableName = "EXPENSE_GROUP"
      FieldName1 = "GROUP_NO"
      FieldName2 = "EXPENSE_GROUP_ID"
      Flag = True
   ElseIf UnqType = REVENUE_UNIQUE Then
      TableName = "REVENUE_GROUP"
      FieldName1 = "GROUP_NO"
      FieldName2 = "REVENUE_GROUP_ID"
      Flag = True
   ElseIf UnqType = PO_UNIQUE Then
      TableName = "PURCHASE_ORDER"
      FieldName1 = "PO_NO"
      FieldName2 = "PO_ID"
      Flag = True
   ElseIf UnqType = CUSTOMER_UNIQUE Then
      TableName = "PATIENT"
      FieldName1 = "PATIENT_CODE"
      FieldName2 = "PATIENT_ID"
      Flag = True
   ElseIf UnqType = BORROW_UNIQUE Then
      TableName = "EMP_RECEIVABLE"
      FieldName1 = "BORROW_NO"
      FieldName2 = "EMP_RECEIVABLE_ID"
      Flag = True
   ElseIf UnqType = TRUCK_UNIQUE Then
      TableName = "RESOURCE"
      FieldName1 = "RESOURCE_NO"
      FieldName2 = "RESOURCE_ID"
      Flag = True
   ElseIf UnqType = JOBPLAN_UNIQUE Then
      TableName = "JOB_PLAN"
      FieldName1 = "PLAN_NO"
      FieldName2 = "JOB_PLAN_ID"
      Flag = True
   ElseIf UnqType = PARTTYPE_NO Then
      TableName = "PART_TYPE"
      FieldName1 = "PART_TYPE_NO"
      FieldName2 = "PART_TYPE_ID"
      Flag = True
   ElseIf UnqType = PARTTYPE_NAME Then
      TableName = "PART_TYPE"
      FieldName1 = "PART_TYPE_NAME"
      FieldName2 = "PART_TYPE_ID"
      Flag = True
   ElseIf UnqType = LOCATION_NO Then
      TableName = "LOCATION"
      FieldName1 = "LOCATION_NO"
      FieldName2 = "LOCATION_ID"
      Flag = True
   ElseIf UnqType = LOCATION_NO_EX Then
      TableName = "LOCATION"
      FieldName1 = "LOCATION_NO"
      FieldName2 = "LOCATION_TYPE"
      Flag = True
   ElseIf UnqType = LOCATION_NAME Then
      TableName = "LOCATION"
      FieldName1 = "LOCATION_NAME"
      FieldName2 = "LOCATION_ID"
      Flag = True
   ElseIf UnqType = PRODUCTTYPE_NO Then
      TableName = "PRODUCT_TYPE"
      FieldName1 = "PRODUCT_TYPE_NO"
      FieldName2 = "PRODUCT_TYPE_ID"
      Flag = True
   ElseIf UnqType = PRODUCTTYPE_NAME Then
      TableName = "PRODUCT_TYPE"
      FieldName1 = "PRODUCT_TYPE_NAME"
      FieldName2 = "PRODUCT_TYPE_ID"
      Flag = True
   ElseIf UnqType = PRODUCTSTATUS_NO Then
      TableName = "PRODUCT_STATUS"
      FieldName1 = "PRODUCT_STATUS_NO"
      FieldName2 = "PRODUCT_STATUS_ID"
      Flag = True
   ElseIf UnqType = PRODUCTSTATUS_NAME Then
      TableName = "PRODUCT_STATUS"
      FieldName1 = "PRODUCT_STATUS_NAME"
      FieldName2 = "PRODUCT_STATUS_ID"
      Flag = True
   ElseIf UnqType = HOUSE_NO Then
      TableName = "HOUSE"
      FieldName1 = "HOUSE_NO"
      FieldName2 = "HOUSE_ID"
      Flag = True
   ElseIf UnqType = HOUSE_NAME Then
      TableName = "HOUSE"
      FieldName1 = "HOUSE_NAME"
      FieldName2 = "HOUSE_ID"
      Flag = True
   ElseIf UnqType = COUNTRY_NO Then
      TableName = "COUNTRY"
      FieldName1 = "COUNTRY_NO"
      FieldName2 = "COUNTRY_ID"
      Flag = True
   ElseIf UnqType = COUNTRY_NAME Then
      TableName = "COUNTRY"
      FieldName1 = "COUNTRY_NAME"
      FieldName2 = "COUNTRY_ID"
      Flag = True
   ElseIf UnqType = CSTGRADE_NO Then
      TableName = "CUSTOMER_GRADE"
      FieldName1 = "CSTGRADE_NO"
      FieldName2 = "CSTGRADE_ID"
      Flag = True
   ElseIf UnqType = CSTGRADE_NAME Then
      TableName = "CUSTOMER_GRADE"
      FieldName1 = "CSTGRADE_NAME"
      FieldName2 = "CSTGRADE_ID"
      Flag = True
   ElseIf UnqType = CSTTYPE_NO Then
      TableName = "CUSTOMER_TYPE"
      FieldName1 = "CSTTYPE_NO"
      FieldName2 = "CSTTYPE_ID"
      Flag = True
   ElseIf UnqType = CSTTYPE_NAME Then
      TableName = "CUSTOMER_TYPE"
      FieldName1 = "CSTTYPE_NAME"
      FieldName2 = "CSTTYPE_ID"
      Flag = True
   ElseIf UnqType = CUSTCODE_UNIQUE Then
      TableName = "CUSTOMER"
      FieldName1 = "CUSTOMER_CODE"
      FieldName2 = "CUSTOMER_ID"
      Flag = True
   ElseIf UnqType = SUPPLIERGRADE_NO Then
      TableName = "SUPPLIER_GRADE"
      FieldName1 = "SUPPLIER_GRADE_NO"
      FieldName2 = "SUPPLIER_GRADE_ID"
      Flag = True
   ElseIf UnqType = SUPPLIERGRADE_NAME Then
      TableName = "SUPPLIER_GRADE"
      FieldName1 = "SUPPLIER_GRADE_NAME"
      FieldName2 = "SUPPLIER_GRADE_ID"
      Flag = True
   ElseIf UnqType = SUPPLIERTYPE_NO Then
      TableName = "SUPPLIER_TYPE"
      FieldName1 = "SUPPLIER_TYPE_NO"
      FieldName2 = "SUPPLIER_TYPE_ID"
      Flag = True
   ElseIf UnqType = SUPPLIERYPE_NAME Then
      TableName = "SUPPLIER_TYPE"
      FieldName1 = "SUPPLIER_TYPE_NAME"
      FieldName2 = "SUPPLIER_TYPE_ID"
      Flag = True
   ElseIf UnqType = SUPPLIERSTATUS_NO Then
      TableName = "SUPPLIER_STATUS"
      FieldName1 = "SUPPLIER_STATUS_NO"
      FieldName2 = "SUPPLIER_STATUS_ID"
      Flag = True
   ElseIf UnqType = SUPPLIERSTATUS_NAME Then
      TableName = "SUPPLIER_STATUS"
      FieldName1 = "SUPPLIER_STATUS_NAME"
      FieldName2 = "SUPPLIER_STATUS_ID"
      Flag = True
   ElseIf UnqType = POSITION_NO Then
      TableName = "EMP_POSITION"
      FieldName1 = "POSITION_NAME"
      FieldName2 = "POSITION_ID"
      Flag = True
   ElseIf UnqType = UNIT_NO Then
      TableName = "UNIT"
      FieldName1 = "UNIT_NO"
      FieldName2 = "UNIT_ID"
      Flag = True
   ElseIf UnqType = UNIT_NAME Then
      TableName = "UNIT"
      FieldName1 = "UNIT_NAME"
      FieldName2 = "UNIT_ID"
      Flag = True
   ElseIf UnqType = YEAR_NO Then
      TableName = "YEAR_SEQ"
      FieldName1 = "YEAR_NO"
      FieldName2 = "YEAR_SEQ_ID"
      Flag = True
   ElseIf UnqType = PARTGROUP_NO Then
      TableName = "PART_GROUP"
      FieldName1 = "PART_GROUP_NO"
      FieldName2 = "PART_GROUP_ID"
      Flag = True
   ElseIf UnqType = PARTGROUP_NAME Then
      TableName = "PART_GROUP"
      FieldName1 = "PART_GROUP_NAME"
      FieldName2 = "PART_GROUP_ID"
      Flag = True
   ElseIf UnqType = DO_PLAN_UNIQUE Then
      TableName = "BILLING_DOC"
      FieldName1 = "DOCUMENT_NO"
      FieldName2 = "BILLING_DOC_ID"
      Flag = True
   ElseIf UnqType = REAL_CREDIT_NO Then
      TableName = "REAL_CREDIT"
      FieldName1 = "DOCUMENT_NO"
      FieldName2 = "ID"
      Flag = True
   ElseIf UnqType = DOCUMENT_CANCEL_NO Then
      TableName = "DOCUMENT_CANCEL"
      FieldName1 = "DOCUMENT_NO"
      FieldName2 = "ID"
      Flag = True
  ElseIf UnqType = MASTER_FT_UNIQUE Then
      TableName = "MASTER_FROMTO2"
      FieldName1 = "MASTER_FROMTO_NO"
      FieldName2 = "MASTER_FROMTO_ID"
      Flag = True
   ElseIf UnqType = CUS_PIG_TYPE_UNIQUE Then
      TableName = "CUS_PIG_TYPE"
      FieldName1 = "CUS_PIG_TYPE_CODE"
      FieldName2 = "CUS_PIG_TYPE_YEAR"
      Flag = True
   ElseIf UnqType = PROVINCE_MAP_UNIQUE Then
      TableName = "PROVINCE_MAP"
      FieldName1 = "KEY_SEARCH"
      FieldName2 = "KEY_ID"
      Flag = True
   ElseIf UnqType = COST_PRODUCTS_UNIQUE Then
      TableName = "COST_PRODUCTS"
      FieldName1 = "PRODUCT_CODE"
      FieldName2 = "PRODUCT_ID"
      Flag = True
   End If
   If Flag Then
      Count = glbDatabaseMngr.CountRecord(TableName, FieldName1, FieldName2, KEY, ID, glbErrorLog, dbType)
      If Count <> 0 Then
         CheckUniqueNs = False
      Else
         CheckUniqueNs = True
      End If
   End If
      
   Exit Function
ErrorHandler:
   glbErrorLog.LocalErrorMsg = "Runtime error."
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   
   CheckUniqueNs = False
End Function

Public Function Check2FlagConvert(A As Long) As String
   If A = 1 Then
      Check2FlagConvert = "N"
   Else
      Check2FlagConvert = "Y"
   End If
End Function

Public Function FlagToCheck(F As String) As Long
   If F = "Y" Then
      FlagToCheck = 1
   Else
      FlagToCheck = 0
   End If
End Function

Public Function Minus2Zero(A As Double) As Long
   If A < 0 Then
      Minus2Zero = 0
   Else
      Minus2Zero = A
   End If
End Function

Public Function Zero2One(A As Double) As Long
   If A = 0 Then
      Zero2One = 1
   Else
      Zero2One = A
   End If
End Function

Public Function Minus2Flag(A As Double) As String
   If A < 0 Then
      Minus2Flag = "Y"
   Else
      Minus2Flag = "N"
   End If
End Function

Public Function AdjustPage(Vsp As VSPrinter, Header As String, Body As String, Offset As Long, Optional TestFlag As Boolean = False, Optional SpaceCount As Long) As Boolean
Dim TempStr As String

   TempStr = Header & Body
   Vsp.CalcTable = TempStr
   
   If (Vsp.Y1 + Offset - SpaceCount) > (Vsp.PageHeight - Vsp.MarginBottom) Then
      If Not TestFlag Then
         Vsp.NewPage
      End If
      AdjustPage = True
   Else
      AdjustPage = False
   End If
End Function

Public Function PatchTable(Vsp As VSPrinter, Header As String, Body As String, Offset As Long, Optional EnableFlag As Boolean = True, Optional SpaceCount As Long = 0) As Boolean
Dim TempStr As String
   
   If Not EnableFlag Then
      PatchTable = True
      Exit Function
   End If
   
   TempStr = Header & Body
   Vsp.CalcTable = TempStr
   
   While Not AdjustPage(Vsp, Header, Body, Offset, True, SpaceCount)
      Call Vsp.AddTable(Header, "", Body)
   Wend
End Function

Public Sub PatchDB()
Dim p As CPatch

   Set p = New CPatch
      
'   If Not p.IsPatch("1_0_12_9") Then
'      Call p.Patch_1_0_12_9
'   End If
      
   Set p = Nothing
End Sub

Public Function MyDiffEx(ByVal D1 As Double, ByVal D2 As Double) As Double
   If D2 = 0 Then
      MyDiffEx = 0
   Else
      MyDiffEx = D1 / D2
   End If
End Function

Public Function MyDiff(ByVal D1 As Double, ByVal D2 As Double) As Double
   If D2 = 0 Then
      MyDiff = 0
   Else
      MyDiff = CDbl(Format(D1 / D2, "0.00"))
   End If
End Function

'Public Sub CheckMemo(TriggerCode As Long)
'Dim M As CMemo
'Dim TempRs As ADODB.Recordset
'Dim ItemCount As Long
'
'   Set M = New CMemo
'   Set TempRs = New ADODB.Recordset
'
'   M.MEMO_ID = -1
'   M.MEMO_STATUS = "N"
'   M.ASSIGN_TO = glbUser.REAL_USER_ID
'   M.FROM_DATE = Now
'   M.TO_DATE = DateAdd("H", 1, M.FROM_DATE)
'   M.TRIGGER_CODE = TriggerCode
'   Call M.QueryData2(TempRs, ItemCount)
'
'   If ItemCount > 0 Then
'      glbErrorLog.LocalErrorMsg = "มีรายการแจ้งเตือนที่ถึงกำหนดแล้ว ท่านต้องการจะดูรายการหรือไม่ ?"
'      If glbErrorLog.AskMessage = vbYes Then
'         frmMemo.MemoStatus = "N"
'         frmMemo.HeaderText = "ตรวจสอบรายการแจ้งเตือน"
'         Load frmMemo
'         frmMemo.Show 1
'
'         Unload frmMemo
'         Set frmMemo = Nothing
'      End If
'   End If
'
'   If TempRs.State = adStateOpen Then
'      TempRs.Close
'   End If
'   Set TempRs = Nothing
'   Set M = Nothing
'End Sub
'
'Public Sub PatchDB()
'Dim p As CPatch
'
'   Set p = New CPatch
'
'   If Not p.IsPatch("3_0_12_19") Then
'      Call p.Patch_3_0_12_19
'   End If
'
'   If Not p.IsPatch("3_0_12_20") Then
'      Call p.Patch_3_0_12_20
'   End If
'
'   If Not p.IsPatch("3_0_12_21") Then
'      Call p.Patch_3_0_12_21
'   End If
'
'   If Not p.IsPatch("3_0_12_22") Then
'      Call p.Patch_3_0_12_22
'   End If
'
'   If Not p.IsPatch("3_0_12_23") Then
'      Call p.Patch_3_0_12_23
'   End If
'
'   Set p = Nothing
'End Sub
'
'Public Function DOType2Flag(DoType As Long) As String
'   If DoType = 1 Then
'      DOType2Flag = "N"
'   ElseIf DoType = 2 Then
'      DOType2Flag = "Y"
'   Else
'      DOType2Flag = ""
'
'   End If
'End Function

Public Function PackAddress(Rs As ADODB.Recordset) As String
Dim AddressStr As String

   AddressStr = ""
   
   If NVLS(Rs("HOME_NO1"), "") <> "" Then
      AddressStr = AddressStr & NVLS(Rs("HOME_NO1"), "") & " "
   End If

   If NVLS(Rs("MOO1"), "") <> "" Then
      AddressStr = AddressStr & "หมู่." & NVLS(Rs("MOO1"), "") & " "
   End If

   If NVLS(Rs("SOI1"), "") <> "" Then
      AddressStr = AddressStr & "ซอย." & NVLS(Rs("SOI1"), "") & " "
   End If

   If NVLS(Rs("ROAD1"), "") <> "" Then
      AddressStr = AddressStr & "ถ." & NVLS(Rs("ROAD1"), "") & " "
   End If

   If NVLS(Rs("KWANG1"), "") <> "" Then
      AddressStr = AddressStr & "แขวง" & NVLS(Rs("KWANG1"), "") & " "
   End If

   If NVLS(Rs("KHATE1"), "") <> "" Then
      AddressStr = AddressStr & "เขต" & NVLS(Rs("KHATE1"), "") & " "
   End If

   If NVLS(Rs("PROVINCE"), "") <> "" Then
      AddressStr = AddressStr & "จ." & NVLS(Rs("PROVINCE"), "") & " "
   End If

   If NVLS(Rs("ZIPCODE1"), "") <> "" Then
      AddressStr = AddressStr & " " & NVLS(Rs("ZIPCODE1"), "") & " "
   End If

   PackAddress = AddressStr
End Function

Public Function MapText(Msg As String) As String
   MapText = Msg
End Function

Public Function SetReportConfig(Vsp As VSPrinter, ReportClassName As String) As Boolean
Dim i As Long
Dim Count As Long
Dim Rp As CReportConfig
Dim TempRs As ADODB.Recordset
Dim Rps As Collection
Dim iCount As Long

   If Rps Is Nothing Then
      Set TempRs = New ADODB.Recordset
      
      Set Rps = New Collection
      Set Rp = New CReportConfig
      
      Rp.REPORT_CONFIG_ID = -1
      Call Rp.QueryData(TempRs, iCount)
      Set Rp = Nothing
      
      While Not TempRs.EOF
         Set Rp = New CReportConfig
         
         Call Rp.PopulateFromRS(1, TempRs)
         Call Rps.Add(Rp)
         
         Set Rp = Nothing
         TempRs.MoveNext
      Wend
      
      Set Rp = Nothing
      If TempRs.State = adStateOpen Then
         TempRs.Close
      End If
      Set TempRs = Nothing
   End If
   
   SetReportConfig = False
   For Each Rp In Rps
      If (Rp.REPORT_KEY = ReportClassName) And (LCase(Rp.COMPUTER_NAME) = LCase(glbParameterObj.ComputerName)) Then
         Vsp.PaperSize = Rp.PAPER_SIZE
         Vsp.ORIENTATION = Rp.ORIENTATION
         Vsp.MarginBottom = Rp.MARGIN_BOTTOM * 567
         Vsp.MarginFooter = Rp.MARGIN_FOOTER * 567
         Vsp.MarginHeader = Rp.MARGIN_HEADER * 567
         Vsp.MarginLeft = Rp.MARGIN_LEFT * 567
         Vsp.MarginRight = Rp.MARGIN_RIGHT * 567
         Vsp.MarginTop = Rp.MARGIN_TOP * 567
'         Vsp.FontName = Rp.FONT_NAME
'         Vsp.FontSize = Rp.FONT_SIZE
         Vsp.MarginLeft = Rp.MARGIN_LEFT * 567
         Vsp.MarginRight = Rp.MARGIN_RIGHT * 567
         If Rp.PAPER_HEIGHT > 0 Then
            Vsp.PaperWidth = Rp.PAPER_HEIGHT * 567
         End If
         If Rp.PAPER_WIDTH > 0 Then
            Vsp.PaperHeight = Rp.PAPER_HEIGHT * 567
         End If
         
         SetReportConfig = True
         Exit Function
      End If
   Next Rp
   Set Rps = Nothing
End Function


Public Function LastDayOfMonth(ByVal ValidDate As Date) As Byte
Dim LastDay As Byte
   LastDay = DatePart("d", DateAdd("d", -1, DateAdd("m", 1, DateAdd("d", -DatePart("d", ValidDate) + 1, ValidDate))))
   LastDayOfMonth = LastDay
End Function

Public Sub GetFirstLastDate(D As Date, Optional FD As Date = -1, Optional Ld As Date = -1)
Dim MM As Long
Dim DD1 As Long
Dim DD2 As Long
Dim YYYY As Long

   MM = Month(D)
   DD1 = 1
   DD2 = LastDayOfMonth(D)
   YYYY = Year(D)
   
   If FD > -1 Then
         FD = DateSerial(YYYY, MM, DD1)
   End If
   If Ld > -1 Then
         Ld = DateSerial(YYYY, MM, DD2)
   End If
End Sub

Public Sub StartExportFile(Vsp As VSPrinter)
   Vsp.ExportFile = ""
   Vsp.ExportFile = glbParameterObj.ReportFile
   Vsp.ExportFormat = vpxPlainHTML
End Sub

Public Sub CloseExportFile(Vsp As VSPrinter)
   Vsp.ExportFile = ""
End Sub

Public Sub InitOrientation(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (ID2Orientation(orLandscape))
   C.ItemData(1) = orLandscape

   C.AddItem (ID2Orientation(orPortrait))
   C.ItemData(2) = orPortrait
End Sub

Public Sub InitPaperSize(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (ID2PaperSize(pprA4))
   C.ItemData(1) = pprA4

   C.AddItem (ID2PaperSize(pprLetter))
   C.ItemData(2) = pprLetter

   C.AddItem (ID2PaperSize(pprFanfoldUS))
   C.ItemData(3) = pprFanfoldUS
End Sub

Public Sub InitFontName(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("AngsanaUPC")
   C.ItemData(1) = 1
End Sub

Public Function ID2Orientation(TempID As OrientationSettings) As String
   If TempID = orLandscape Then
      ID2Orientation = "แนวนอน"
   Else
      ID2Orientation = "แนวตั้ง"
   End If
End Function

Public Function ID2PaperSize(TempID As PaperSizeSettings) As String
   If TempID = pprA4 Then
      ID2PaperSize = "A4"
   ElseIf TempID = pprLetter Then
      ID2PaperSize = "Letter"
   ElseIf TempID = pprFanfoldUS Then
      ID2PaperSize = "Us standard"
   Else
      ID2PaperSize = "A4"
   End If
End Function

Public Function GetNextID(OldID As Long, Col As Collection) As Long
Dim O As Object
Dim i As Long

   i = 0
   For Each O In Col
      i = i + 1
      If (i > OldID) And (O.Flag <> "D") Then
         GetNextID = i
         Exit Function
      End If
   Next O
   GetNextID = OldID
End Function
Public Function GetObject(ClassName As String, m_TempCol As Collection, TempKey As String, Optional SetNew As Boolean = True) As Object
On Error Resume Next
Dim Ei As Object
   
   Set Ei = m_TempCol(TempKey)
   If Ei Is Nothing Then
      If SetNew Then
         Set GetObject = GetNewClass(ClassName)
      End If
   Else
      Set GetObject = Ei
   End If
End Function
Public Function GetNewClass(ClassName As String) As Object
   If ClassName = "CRealCredit" Then
      Static m_CRealCredit As CRealCredit
      If m_CRealCredit Is Nothing Then
         Set m_CRealCredit = New CRealCredit
      End If
      Set GetNewClass = m_CRealCredit
   ElseIf ClassName = "CAPMas" Then
      Static m_CAPMas As CAPMas
      If m_CAPMas Is Nothing Then
         Set m_CAPMas = New CAPMas
      End If
      Set GetNewClass = m_CAPMas
   ElseIf ClassName = "CARRcIt" Then
      Static m_CARRcIt As CARRcIt
      If m_CARRcIt Is Nothing Then
         Set m_CARRcIt = New CARRcIt
      End If
      Set GetNewClass = m_CARRcIt
   ElseIf ClassName = "CStcrd" Then
      Static m_CStcrd As CStcrd
      If m_CStcrd Is Nothing Then
         Set m_CStcrd = New CStcrd
      End If
      Set GetNewClass = m_CStcrd
   ElseIf ClassName = "CARTrn" Then
      Static m_CARTrn As CARTrn
      If m_CARTrn Is Nothing Then
         Set m_CARTrn = New CARTrn
      End If
      Set GetNewClass = m_CARTrn
   ElseIf ClassName = "CGroupType" Then
      Static m_CGroupType As CGroupType
      If m_CGroupType Is Nothing Then
         Set m_CGroupType = New CGroupType
      End If
      Set GetNewClass = m_CGroupType
    ElseIf ClassName = "CPromotionConfig" Then
      Static m_CPromotionConfig As CPromotionConfig
      If m_CPromotionConfig Is Nothing Then
         Set m_CPromotionConfig = New CPromotionConfig
      End If
      Set GetNewClass = m_CPromotionConfig
   ElseIf ClassName = "CPromotionPayCustom" Then
      Static m_CPromotionPayCustom As CPromotionPayCustom
      If m_CPromotionPayCustom Is Nothing Then
         Set m_CPromotionPayCustom = New CPromotionPayCustom
      End If
      Set GetNewClass = m_CPromotionPayCustom
   ElseIf ClassName = "CPromotionYear" Then
      Static m_CPromotionYear As CPromotionYear
      If m_CPromotionYear Is Nothing Then
         Set m_CPromotionYear = New CPromotionYear
      End If
      Set GetNewClass = m_CPromotionYear
   End If

End Function
Public Sub InitThaiMonth(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("มกราคม"))
   C.ItemData(1) = 1

   C.AddItem (MapText("กุมภาพันธ์"))
   C.ItemData(2) = 2

   C.AddItem (MapText("มีนาคม"))
   C.ItemData(3) = 3

   C.AddItem (MapText("เมษายน"))
   C.ItemData(4) = 4

   C.AddItem (MapText("พฤษภาคม"))
   C.ItemData(5) = 5
   
   C.AddItem (MapText("มิถุนายน"))
   C.ItemData(6) = 6

   C.AddItem (MapText("กรกฎาคม"))
   C.ItemData(7) = 7
   
   C.AddItem (MapText("สิงหาคม"))
   C.ItemData(8) = 8
   
   C.AddItem (MapText("กันยายน"))
   C.ItemData(9) = 9
   
   C.AddItem (MapText("ตุลาคม"))
   C.ItemData(10) = 10
   
   C.AddItem (MapText("พฤษศจิกายน"))
   C.ItemData(11) = 11

   C.AddItem (MapText(" ธันวาคม"))
   C.ItemData(12) = 12
End Sub
Public Function PatchWildCard(T As String) As String
   If Len(Trim(T)) <> 0 Then
      PatchWildCard = T & "%"
   Else
      PatchWildCard = T
   End If
End Function
Public Function PatchWildCard2(T As String) As String
   If Len(Trim(T)) <> 0 Then
      PatchWildCard2 = "%" & T & "%"
   Else
      PatchWildCard2 = T
   End If
End Function

Public Sub PopulateInternalField(ShowMode As SHOW_MODE_TYPE, O As Object)
Dim Tf As CTableField
Dim TempID As Long
Dim InternalDate As String

   For Each Tf In O.m_FieldList
      If Tf.FieldCat = ID_CAT Then
         If ShowMode = SHOW_ADD Then
            Call glbDatabaseMngr.GetSeqID(O.SequenceName, TempID, glbErrorLog)
            Call Tf.SetValue(TempID)
         End If
      ElseIf Tf.FieldCat = CREATE_DATE_CAT Then
         If ShowMode = SHOW_ADD Then
            Call glbDatabaseMngr.GetServerDateTime(InternalDate, glbErrorLog)
            Call Tf.SetValue(InternalDateToDate(InternalDate))
         End If
      ElseIf Tf.FieldCat = MODIFY_DATE_CAT Then
         'If ShowMode = SHOW_EDIT Then
            Call glbDatabaseMngr.GetServerDateTime(InternalDate, glbErrorLog)
            Call Tf.SetValue(InternalDateToDate(InternalDate))
         'End If
      ElseIf Tf.FieldCat = CREATE_BY_CAT Then
         If ShowMode = SHOW_ADD Then
            Call Tf.SetValue(glbUser.USER_ID)
         End If
      ElseIf Tf.FieldCat = MODIFY_BY_CAT Then
         'If ShowMode = SHOW_EDIT Then
            Call Tf.SetValue(glbUser.USER_ID)
         'End If
      End If
   Next Tf
End Sub


Public Function GenerateUpdateSQL(O As Object) As String
Dim Tf As CTableField
Dim SQL As String
Dim Sep As String
Dim TempKeyName As String
Dim TempKeyVal As Long

   SQL = "UPDATE " & O.TableName & " SET" & vbCrLf
   For Each Tf In O.m_FieldList
      If Tf.FieldCat <> TEMP_CAT And Tf.FieldCat <> CREATE_DATE_CAT And Tf.FieldCat <> CREATE_BY_CAT Then
         If Tf.FieldCat = ID_CAT Then
            TempKeyName = Tf.FieldName
            TempKeyVal = Tf.GetValue
         Else
            If Tf.FieldCat = MODIFY_BY_CAT Then
               Sep = "" & vbCrLf
            Else
               Sep = ", " & vbCrLf
            End If
            
            SQL = SQL & Tf.FieldName & " = " & Tf.TransformToSQLString & Sep
         End If
      End If
   Next Tf
      
   SQL = SQL & "WHERE " & TempKeyName & " = " & TempKeyVal
   
   GenerateUpdateSQL = SQL
End Function

Public Function GenerateInsertSQL(O As Object) As String





Dim Tf As CTableField
Dim SQL As String
Dim Sep As String

   SQL = "INSERT INTO " & O.TableName & vbCrLf & " (" & vbCrLf
   For Each Tf In O.m_FieldList
      If Tf.FieldCat <> TEMP_CAT Then
         If Tf.FieldCat = MODIFY_BY_CAT Then
            Sep = "" & vbCrLf & ") " & vbCrLf & "VALUES " & vbCrLf & "(" & vbCrLf
         Else
            Sep = ", " & vbCrLf
         End If
         
         SQL = SQL & Tf.FieldName & Sep
      End If
   Next Tf
   
   For Each Tf In O.m_FieldList
      If Tf.FieldCat <> TEMP_CAT Then
         If Tf.FieldCat = MODIFY_BY_CAT Then
            Sep = "" & vbCrLf & ")"
         Else
            Sep = ", " & vbCrLf
         End If
''debug.print "---" & Tf.FieldName
         SQL = SQL & Tf.TransformToSQLString & Sep
''debug.print "---" & Tf.GetValue
      End If
   Next Tf
   
   GenerateInsertSQL = SQL
End Function
Public Function GenerateSearchLike(StartWith As String, SearchIn As String, SubLen As Long, NewStr As String) As String
    Dim WhereStr As String
    Dim StartStringNo As Long
    Dim i As Long
    StartStringNo = 1
    WhereStr = " " & StartWith & "((SUBSTR(" & SearchIn & "," & StartStringNo & "," & StartStringNo + SubLen - 1 & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
    For i = 2 To 50
        StartStringNo = StartStringNo + 1
        WhereStr = WhereStr & " OR " & "(SUBSTR(" & SearchIn & "," & StartStringNo & "," & StartStringNo + SubLen - 1 & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
    Next i
    WhereStr = WhereStr & ")"
    GenerateSearchLike = WhereStr
End Function
