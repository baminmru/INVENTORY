VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDB2TDS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Обмен данными с ТДС"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Зарузка из файла"
      Height          =   1335
      Left            =   120
      TabIndex        =   8
      Top             =   4080
      Width           =   4935
      Begin VB.CommandButton cmdPath 
         Caption         =   "..."
         Height          =   375
         Left            =   2880
         TabIndex        =   11
         Top             =   720
         Width           =   450
      End
      Begin VB.CommandButton cmdManualLoad 
         Caption         =   "Ручная загрузка"
         Height          =   375
         Left            =   3360
         TabIndex        =   10
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtPath 
         Height          =   405
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label Label2 
         Caption         =   "Файл базы данных"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.TextBox txtLog 
      Height          =   1695
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Top             =   2280
      Width           =   4935
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Visible         =   0   'False
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Отмена"
      Height          =   495
      Left            =   3480
      TabIndex        =   4
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Обмен данными"
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton cmdInvI 
      Caption         =   "..."
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txtINV 
      Height          =   405
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   4335
   End
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   4920
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblStatus 
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   4815
   End
   Begin VB.Label Label1 
      Caption         =   "Инвентаризация"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmDB2TDS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Cnn As cConnection
Private rs As cRecordset
Private irs As ADODB.Recordset
Private W As Writer

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public Function GetUserTempPath() As String
  Dim sTempPath As String
  sTempPath = Space(1024)
  sTempPath = Replace(sTempPath, " ", "\")
  Call GetTempPath(Len(sTempPath), sTempPath)
  sTempPath = Replace(sTempPath, "" + Chr(0), "")
  Dim i  As Integer
  For i = 1 To 1024
  sTempPath = Replace(sTempPath, "\\", "\")
  Next
  GetUserTempPath = sTempPath
End Function

Private Sub cmdInvI_Click()
Dim id As String
Dim brief As String
If Manager.GetObjectListDialogEx3(id, brief, "", "INV_INV") Then
    txtINV.Text = brief
    txtINV.Tag = id
End If

End Sub
Private Sub Log(v As String)
  If W Is Nothing Then
      Set W = New Writer
      Dim temppath As String
      Dim fname As String
      temppath = GetSetting("MTZ", "CONFIG", "TEMPPATH", GetUserTempPath())
      fname = temppath & "\SGS_TDS_" & Year(Date) & "_" & Month(Date) & "_" & Day(Date) & "_" & Hour(Now) & "_" & Minute(Now) & "_" & Second(Now) & ".log"
      W.SetFilePath fname
      W.putBuf "Start log:" & Now
  End If
  
  lblStatus = v
  txtLog = txtLog & vbCrLf & v
  W.putBuf v
  If Err.Number <> 0 Then
    txtLog = txtLog & vbCrLf & Err.Description
    W.putBuf Err.Description
    Err.Clear
  End If
  DoEvents
End Sub


Private Sub cmdManualLoad_Click()
  Dim fname As String
  fname = txtPath.Text
 Log "Есть ли файл с БД " & fname
    If FileExists(fname) Then
        Log "file " & fname & " OK"
        If Not ProcessTDS(fname) Then
          MsgBox "Не удалось загрузить данные из терминала.", vbCritical + vbOKOnly, "Ошибка"
          Log "Не удалось загрузить данные из терминала."
        Else
          MsgBox "Загружены данные из терминала.", vbOKOnly, "Загрузка"
          Log "Загружены данные из терминала."
        End If
    Else
      Log "Файл <" & fname & "> не обнаружен"
    End If
End Sub

Private Sub cmdPath_Click()
  
  On Error GoTo bye
  Dim fn As String
  cdlg.CancelError = True
  cdlg.Filter = "База данных ТДС |*.db3"
  cdlg.DefaultExt = "db3"
  cdlg.flags = cdlOFNPathMustExist + cdlOFNHideReadOnly + cdlOFNFileMustExist
  cdlg.ShowOpen
  txtPath = cdlg.fileName
bye:
End Sub

Private Sub Command1_Click()
   On Error GoTo bye
    Dim fname As String
    Dim temppath As String
    temppath = GetSetting("MTZ", "CONFIG", "TEMPPATH", GetUserTempPath())
    fname = temppath & "\TDS2OMK_" & Year(Date) & Month(Date) & Day(Date) & Hour(Now) & Minute(Now) & Second(Now) & ".db3"
    Log "Соединение с устройством"
  
     If RapiConnect Then
        DoEvents
        Log "Получение файла"
        Log ">Copy" & "/OMK.db3 to " & fname
        RAPICopyCEFileToPC "/OMK.db3", fname
        Log ">Copy OK"
        DoEvents
        On Error Resume Next
        Log "Disconnect"
        RapiDisconnect
        Log ">Disconnect OK"
     Else
        Log "Не найдено устройство"
     End If
     On Error GoTo bye
    Log "Есть ли файл с БД " & fname
    If FileExists(fname) Then
        Log "file " & fname & " OK"
        If Not ProcessTDS(fname) Then
          MsgBox "Не удалось загрузить данные из терминала.", vbCritical + vbOKOnly, "Ошибка"
          Log "Не удалось загрузить данные из терминала."
        Else
          MsgBox "Загружены данные из терминала.", vbOKOnly, "Загрузка"
          Log "Загружены данные из терминала."
        End If
    Else
      Log "Файл <" & fname & "> не обнаружен"
    End If
    
    If txtINV.Tag <> "" Then
        fname = temppath & "\OMK2TDS_" & Year(Date) & Month(Date) & Day(Date) & Hour(Now) & Minute(Now) & Second(Now) & ".db3"
        Log "Выгрузка данных в файл " & fname
        BuildMobileDB fname
        If FileExists(fname) Then
        Log "Соединение с устройством"
        If RapiConnect Then
            Log "Удаление файла на усройстве"
            CeDeleteFile "/OMK.db3"
            Log "Кпирование данных на устройство"
            RAPICopyPCFileToCE fname, "/OMK.db3"
            Log "Отключение от устройства"
            RapiDisconnect
            Log "Устройство отключено"
        Else
            MsgBox "Вставте  Терминал сбора данных в кредл.", vbInformation, "Ошибка передачи данных"
            Log "Вставте  Терминал сбора данных в кредл."
            Exit Sub
        End If
        MsgBox "Данные переданы на терминал", vbOKOnly, "Выгрузка данных"
        Log "Данные переданы на терминал"
        Else
          Log "Ошибка создания файла данных"
          MsgBox "", vbInformation, "Ошибка создания файла данных"
        End If
    End If
    Exit Sub
bye:
    Log Err.Description
    MsgBox "Ошибка при попытке обмена с терминалом сбора данных. Проверьие наличие ActiveSync."
        
End Sub

Private Sub Command2_Click()
Me.Hide
End Sub


Private Sub BuildMobileDB(ByVal fileName As String)
  On Error Resume Next
  
  Log "Create new sqlite connection"
  Set Cnn = New cConnection 'instantiate the Connection-Object
  If FileExists(fileName) Then
    Log "Kill old file on PC"
    Kill fileName
  End If
  
  Log "Create new base for mobile"
  Cnn.CreateNewDB fileName
  Log "Create new base OK"
  Log "Query for Record count"
  Dim cnt As Long
  Set irs = Session.GetData("select count(*) cnt " & _
                "from v_rpt_inventory_bad where inv_instanceid='" & txtINV.Tag & "'")
                
  Log "Select " & irs!cnt & " OS"
  cnt = irs!cnt
  pb.Min = 0
  pb.Max = cnt
  pb.Value = 0
  pb.Visible = True
  
  Log "Query for Records"
  Set irs = Session.GetData("select " & _
        "INVOS_INFO_ShortName " & _
        ",INVOS_INFO_Name " & _
        ",INVOS_INFO_CardNum " & _
        ",INVOS_PLACE_ComplNumber " & _
        ",INVOS_INFO_info " & _
        ",replace(COALESCE(cast(INVOS_PLACE_theowner as varchar),''),';', '') INVOS_PLACE_theowner " & _
        ",INVOS_INFO_INVNum " & _
        ",INV_INSTANCEID " & _
        ",InstanceID " & _
        ",VisibleCode " & _
        ",0 Changed " & _
        "from v_rpt_inventory_bad where inv_instanceid='" & txtINV.Tag & "'")
  
  Log "Create INV"
  CreateTableFromADORs Cnn, "INV", irs
   pb.Visible = False

  Log "Create INV_IDX"
  Cnn.Execute " create index INV_IDX on INV(VisibleCode)"
  pb.Value = 0
  
   Set irs = Session.GetData("select count(*) cnt " & _
                "from INVD_OWNER")
                
  Log "Select " & irs!cnt & " owners"
  cnt = irs!cnt
  pb.Min = 0
  pb.Max = cnt
  pb.Value = 0
  pb.Visible = True
  
  If Session.IsMSSQL Then
    Set irs = Session.GetData("select isnull(FamiliName,'?') +' ' + isnull(Name,'?') +' ' +isnull(SurName,'?')+ ' ' Name from INVD_OWNER order by FamiliName,Name,Surname ")
  End If
  If Session.IsPOSTGRESQL Then
   Set irs = Session.GetData("select  COALESCE(cast(FamiliName as varchar),'?') ||' ' ||  COALESCE(cast(Name as varchar),'?') ||' ' || COALESCE(cast(SurName as varchar),'?') || ' ' as Name from INVD_OWNER order by FamiliName,Name,Surname ")
  End If
  
  Log "Create OWNERS"
  CreateTableFromADORs Cnn, "OWNERS", irs
  pb.Visible = False
 
  Log "Create T"
  Cnn.Execute "Create Table T(shCode Text, Status Text, CheckTime datetime, INVID TEXT, OSID TEXT)"
  Log "Create T_IDX"
  Cnn.Execute " create index T_IDX on T(shCode)"
  Log "Create B"
  Cnn.Execute "Create Table B(shCode Text,  CheckTime datetime)"
  Log "Create B_IDX"
  Cnn.Execute " create index B_IDX on B(shCode)"
  Log "Create RENT"
  Cnn.Execute "Create Table RENT(shCode Text, INFO TEXT, CheckTime datetime)"
  Log "Create RENT_IDX"
  Cnn.Execute " create index RENT_IDX on RENT(shCode)"
  Log "Create REP"
  Cnn.Execute "Create Table REP(shCode Text,  INFO TEXT, CheckTime datetime)"
  Log "Create REP_IDX"
  Cnn.Execute " create index REP_IDX on REP(shCode)"
  Log "Create EXPL"
  Cnn.Execute "Create Table EXPL(shCode Text,   CheckTime datetime)"
  Log "Create EXPL_IDX"
  Cnn.Execute " create index EXPL_IDX on EXPL(shCode)"
  Log "Create RPT"
  Cnn.Execute "Create Table PRT(shCode Text,   CheckTime datetime)"
  Log "Create RPT_IDX"
  Cnn.Execute " create index PRT_IDX on PRT(shCode)"
  Log "Create U"
  Cnn.Execute "Create Table U (PLACE Text,NAME Text,INFO Text)"
  Log "Database for mobile created"
End Sub


Private Sub CreateTableFromADORs(CnnDst As cConnection, ByVal TableName As String, adoRs As Object)
Dim rs As ADODB.Recordset, Fld As ADODB.FIELD, v, B() As Byte
Dim i As Long, SQLiteFTs() As dhRichClient3.FIELDTYPE, FDesc() As String, Cmd As cCommand

On Error Resume Next
  Set rs = adoRs 'first we do the cast
  If rs Is Nothing Then Exit Sub
  If rs.fields.Count = 0 Then Exit Sub
  
  ReDim SQLiteFTs(rs.fields.Count - 1) 'redim the FieldType-Array
  ReDim FDesc(rs.fields.Count - 1) 'redim the Field-Description-Array
  
  If Left$(TableName, 1) <> "[" Then TableName = "[" & TableName & "]"
  
  'now scan for the Fieldnames and FieldTypes
  For i = 0 To rs.fields.Count - 1
  
    FDesc(i) = " [" & rs.fields(i).Name & "] " & _
               GetSQLiteFieldType(rs.fields(i).Type, SQLiteFTs(i))
  Next i
  
On Error GoTo RollBack
  CnnDst.BeginTrans
    'first we try to create the appropriate table on the SQLiteCnn
    CnnDst.Execute "Create Table " & TableName & " (" & Join$(FDesc, ",") & ")"
    
    'now the insert-loop
    If Not rs.EOF Then
      For i = 0 To UBound(FDesc)
        FDesc(i) = "?" 'prepare InsertParam-PlaceHolders
      Next i
      
      Set Cmd = CnnDst.CreateCommand("Insert Into " & TableName & " Values(" & Join$(FDesc, ",") & ")")
      
      Do Until rs.EOF
        pb.Value = pb.Value + 1
        For i = 0 To UBound(SQLiteFTs)
          With rs.fields(i)
            v = .Value
            If IsNull(v) Then
              Cmd.SetNull i + 1
            Else
              Select Case SQLiteFTs(i)
                Case SQLite_TEXT: Cmd.SetText i + 1, CStr(v)
                Case SQLite_INTEGER: Cmd.SetInt32 i + 1, v
                Case SQLite_DOUBLE: Cmd.SetDouble i + 1, v
                Case VB_Boolean_AutoConverted: Cmd.SetBoolean i + 1, v
                Case VB_DATE_AutoConverted: Cmd.SetDate i + 1, v
                Case VB_ShortDate_AutoConverted: Cmd.SetShortDate i + 1, v
                Case VB_Time_AutoConverted: Cmd.SetTime i + 1, v
                Case SQLite_BLOB: B = v: Cmd.SetBlob i + 1, B
              End Select
            End If
          End With
        Next i
        Cmd.Execute
        
        rs.MoveNext
      Loop
    End If
  CnnDst.CommitTrans
Exit Sub
RollBack:
  CnnDst.RollbackTrans
End Sub

Private Function GetSQLiteFieldType(ByVal DataType As ADODB.DataTypeEnum, ByRef SQLiteFT As dhRichClient3.FIELDTYPE) As String
  Select Case DataType
    Case adBoolean
      GetSQLiteFieldType = "BIT": SQLiteFT = VB_Boolean_AutoConverted
    Case adInteger, adBigInt, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
      GetSQLiteFieldType = "INTEGER": SQLiteFT = SQLite_INTEGER
    Case adDate, adDBTimeStamp
      GetSQLiteFieldType = "DATE": SQLiteFT = VB_DATE_AutoConverted
    Case adDBDate
      GetSQLiteFieldType = "SHORTDATE": SQLiteFT = VB_ShortDate_AutoConverted
    Case adDBTime
      GetSQLiteFieldType = "TIME": SQLiteFT = VB_Time_AutoConverted
    Case adDouble, adSingle, adCurrency, adNumeric, adVarNumeric, adDecimal
      GetSQLiteFieldType = "REAL": SQLiteFT = SQLite_DOUBLE
    Case adBinary, adVarBinary, adLongVarBinary
      GetSQLiteFieldType = "BLOB": SQLiteFT = SQLite_BLOB
    Case Else 'adChar, adWChar, adVarChar, adVarWChar, adBSTR, adGUID
      GetSQLiteFieldType = "TEXT": SQLiteFT = SQLite_TEXT
  End Select
End Function

Private Function ProcessTDS(ByVal fname As String) As Boolean
  On Error GoTo bye
  ProcessTDS = False
  Log "Open as database " & fname
  Set Cnn = New cConnection 'instantiate the Connection-Object
  If Not FileExists(fname) Then
    Log "Process data. File not found " & fname
    Exit Function
  End If
  Cnn.OpenDB fname
  Log "DB Open ok"

  
  
  Log "Process Scanned Codes"
  Set rs = Cnn.OpenRecordset("select count(*) from T")
  Dim cnt As Long
  cnt = rs.fields(0).Value
  Log "Found " & cnt & " barcodes"
  
  If cnt = 0 Then
    cnt = 1
  End If
  
  pb.Min = 0
  pb.Max = cnt
  pb.Value = 0
  pb.Visible = True
  
  Log "update scancodes"
  Set rs = Cnn.OpenRecordset("select * from T")
  
  Dim os As INV_OS.Application
  Dim osi As INV_OS.INVOS_INFO
  Dim inv As INV_INV.Application
  
  While Not rs.EOF
    pb.Value = pb.Value + 1
    Log rs.fields(0).Value & " " & rs.fields(1).Value
    Set inv = Manager.GetInstanceObject(rs.fields(3).Value)
    If Not inv Is Nothing Then
        Log "INV found: " & inv.brief
        
        Session.GetData ("DELETE FROM INVI_DONE WHERE instanceid='" & inv.id & "' and  THEOS NOT IN (SELECT INVOS_INFOID FROM INVOS_INFO)")
        Session.GetData ("DELETE FROM INVI_OBJ WHERE instanceid='" & inv.id & "' and  THEOS NOT IN (SELECT INVOS_INFOID FROM INVOS_INFO)")


        Set os = Manager.GetInstanceObject(rs.fields(4).Value)
        If Not os Is Nothing Then
          Log "OS found: " & os.INVOS_INFO.Item(1).Name & " Code:" & os.INVOS_CODE.Item(1).VisibleCode
          If os.INVOS_INFO.Count > 0 Then
          
           
            Set osi = os.INVOS_INFO.Item(1)
            Log "os.INVOS_INFO exists"
            Log "clean INV_DONE"
            Session.GetData ("DELETE FROM INVI_DONE WHERE instanceid='" & inv.id & "' and TheOs='" & osi.id & "'")
            ' check for duplicate row
            
            Dim jj As Integer
'            Log "clean INV_DONE"
'again_jj:
'            inv.INVI_DONE.Refresh
'            For jj = 1 To inv.INVI_DONE.Count
'
'              'Log jj & "(" & inv.INVI_DONE.Count & ")"
'
'              If inv.INVI_DONE.Item(jj).TheOS Is Nothing Then
'
'                inv.INVI_DONE.Delete jj
'                GoTo again_jj
'              Else
'                If inv.INVI_DONE.Item(jj).TheOS.id = osi.id Then
'                  inv.INVI_DONE.Delete jj
'                  GoTo again_jj
'                End If
'              End If
'            Next
            
            
            Log "Add to INV_DONE"
            inv.INVI_DONE.Refresh
            With inv.INVI_DONE.Add
                Log "Set OS"
                Set .TheOS = osi
                Log "Set CheckDate"
                .CheckDate = rs.fields(2).Value
                Log "Set Status"
                Set .OSStatus = FindStatusByName(rs.fields(1).Value)
                Log "Saving"
                .save
            End With
          End If
          
          ' kill perv data
          Dim kk As Integer
again_kk:
          
          Log "clean INVOS_INV"
          
          Session.GetData ("DELETE FROM INVOS_INV WHERE INSTANCEID='" & os.id & "' AND  INVENTORY NOT IN ( SELECT INVI_defID FROM INVI_def)")
          Session.GetData "DELETE FROM INVOS_INV WHERE INSTANCEID='" & os.id & "' AND  INVENTORY ='" & inv.invi_DEF.Item(1).id & "'"
          
'          For kk = 1 To os.INVOS_INV.Count
'            If os.INVOS_INV.Item(kk).Inventory Is Nothing Then
'               os.INVOS_INV.Delete kk
'               GoTo again_kk
'            Else
'              If os.INVOS_INV.Item(kk).Inventory.id = inv.invi_DEF.Item(1).id Then
'                os.INVOS_INV.Delete kk
'                GoTo again_kk
'              End If
'            End If
'          Next
'
           Log "ADD INVOS_INV"
           
           os.INVOS_INV.Refresh
           With os.INVOS_INV.Add
              Set .OSStatus = FindStatusByName(rs.fields(1).Value)
              .InvDate = rs.fields(2).Value
              Set .Inventory = inv.invi_DEF.Item(1)
              .save
           End With
           
          Manager.FreeInstanceObject os.id
          Set os = Nothing
        End If
    End If
  
    rs.MoveNext
  Wend
  pb.Visible = False
  
  
  
  Log "Process INV changed"
  Set rs = Cnn.OpenRecordset("select count(*) from INV where changed = 1")
  cnt = rs.fields(0).Value
  If cnt = 0 Then
  cnt = 1
  End If
  pb.Min = 0
  pb.Max = cnt
  pb.Value = 0
  pb.Visible = True
  
  Set rs = Cnn.OpenRecordset("select instanceid,invos_info_info,invos_place_theowner from INV where changed =1")
  

  
  While Not rs.EOF
    pb.Value = pb.Value + 1
    
    Set os = Manager.GetInstanceObject(rs.fields("instanceid").Value)
    If Not os Is Nothing Then
      If os.INVOS_INFO.Count > 0 Then
          Set osi = os.INVOS_INFO.Item(1)
          With osi
              .Info = rs.fields("invos_info_info").Value
              .save
          End With
      End If
      
      If os.INVOS_PLACE.Count = 0 Then
        os.INVOS_PLACE.Add
      End If
      With os.INVOS_PLACE.Item(1)
          Set .TheOwner = FindOwnerByName(rs.fields("invos_place_theowner").Value)
          .save
      End With
      
    End If
    Manager.FreeInstanceObject os.id
    Set os = Nothing
    rs.MoveNext
  Wend
  pb.Visible = False
  
  
  Log "Process REP"
  Set rs = Cnn.OpenRecordset("select count(*) from REP")
  cnt = rs.fields(0).Value
  
    If cnt = 0 Then
  cnt = 1
  End If
  pb.Min = 0
  pb.Max = cnt
  pb.Value = 0
  pb.Visible = True
  
  Set rs = Cnn.OpenRecordset("select *  from REP")

  
  While Not rs.EOF
    pb.Value = pb.Value + 1
    
    Set os = Code2OS(rs.fields("shCode").Value)
    If Not os Is Nothing Then
      os.StatusID = "{8E6E78D2-82AA-4913-B08C-1230A8C8B4A9}"
      With os.INVOS_REPAIR.Add
        .Info = rs.fields("INFO").Value
        .save
      End With
      
    End If
    Manager.FreeInstanceObject os.id
    Set os = Nothing
    rs.MoveNext
  Wend
  pb.Visible = False
  
  
  Log "Process RENT"
  Set rs = Cnn.OpenRecordset("select count(*) from RENT")
  cnt = rs.fields(0).Value
  If cnt = 0 Then
  cnt = 1
  End If
  pb.Min = 0
  pb.Max = cnt
  pb.Value = 0
  pb.Visible = True
  
  Set rs = Cnn.OpenRecordset("select *  from RENT")
  

  While Not rs.EOF
    pb.Value = pb.Value + 1
    
    Set os = Code2OS(rs.fields("shCode").Value)
    If Not os Is Nothing Then
      os.StatusID = "{2AA78799-2880-4541-99E0-3C8750AC33E6}"
      With os.INVOS_RENT.Add
        .DocNumber = rs.fields("INFO").Value
        .save
      End With
    End If
    Manager.FreeInstanceObject os.id
    Set os = Nothing
    rs.MoveNext
  Wend
  pb.Visible = False
  
  
  Log "Process EXPL"
  Set rs = Cnn.OpenRecordset("select count(*) from EXPL")
  cnt = rs.fields(0).Value
  If cnt = 0 Then
  cnt = 1
  End If
  pb.Min = 0
  pb.Max = cnt
  pb.Value = 0
  pb.Visible = True
  
  Set rs = Cnn.OpenRecordset("select *  from EXPL")
  

  While Not rs.EOF
    pb.Value = pb.Value + 1
    
    Set os = Code2OS(rs.fields("shCode").Value)
    If Not os Is Nothing Then
    os.StatusID = "{8AD15E54-CF87-4FCF-8A1E-A85336E23C73}"
    End If
    Manager.FreeInstanceObject os.id
    Set os = Nothing
    rs.MoveNext
  Wend
  pb.Visible = False
  
  
  Set inv = Nothing
  Set rs = Cnn.OpenRecordset("select INV_INSTANCEID from INV")
  If Not rs.EOF Then
      Set inv = Manager.GetInstanceObject(rs!inv_instanceid)
  End If
  
  
  If Not inv Is Nothing Then
      Log "Process B"
      Set rs = Cnn.OpenRecordset("select count(*) from B")
      cnt = rs.fields(0).Value
      If cnt = 0 Then
      cnt = 1
      End If
      pb.Min = 0
      pb.Max = cnt
      pb.Value = 0
      pb.Visible = True
      
      Set rs = Cnn.OpenRecordset("select *  from B")
      
   
      While Not rs.EOF
        pb.Value = pb.Value + 1
again_BAD:
          For jj = 1 To inv.INVI_BAD.Count
            If inv.INVI_BAD.Item(jj).ShCode = rs.fields("shCode").Value Then
              inv.INVI_BAD.Delete jj
              GoTo again_BAD
            End If
          Next
        With inv.INVI_BAD.Add
          .ShCode = rs.fields("shCode").Value
          .save
        End With
        rs.MoveNext
      Wend
      pb.Visible = False
      
      
      Log "Process U"
      Set rs = Cnn.OpenRecordset("select count(*) from U")
      cnt = rs.fields(0).Value
      If cnt = 0 Then
      cnt = 1
      End If
      pb.Min = 0
      pb.Max = cnt
      pb.Value = 0
      pb.Visible = True
      
      Set rs = Cnn.OpenRecordset("select *  from U")
    
      While Not rs.EOF
        pb.Value = pb.Value + 1
        
again_UNK:
          For jj = 1 To inv.INVI_UNK.Count
            If inv.INVI_UNK.Item(jj).TheRoom = rs.fields("PLACE").Value And inv.INVI_UNK.Item(jj).Info = "" & rs.fields("Name").Value & " " & rs.fields("Info").Value Then
              inv.INVI_UNK.Delete jj
              GoTo again_UNK
            End If
          Next
        
        With inv.INVI_UNK.Add
            .TheRoom = rs.fields("PLACE").Value
            .Info = "" & rs.fields("Name").Value & " " & rs.fields("Info").Value
            .save
        End With
       
        rs.MoveNext
      Wend
      pb.Visible = False
      
      
      Log "Process PRT"
      Set rs = Cnn.OpenRecordset("select count(*) from PRT")
      cnt = rs.fields(0).Value
      If cnt = 0 Then
      cnt = 1
      End If
      pb.Min = 0
      pb.Max = cnt
      pb.Value = 0
      pb.Visible = True
      
      Set rs = Cnn.OpenRecordset("select *  from PRT")
      
      
      While Not rs.EOF
        pb.Value = pb.Value + 1
        
again_CHNG:
          For jj = 1 To inv.INVI_CHNG.Count
            If inv.INVI_CHNG.Item(jj).ShCode = rs.fields("shCode").Value Then
              inv.INVI_CHNG.Delete jj
              GoTo again_CHNG
            End If
          Next
        
        With inv.INVI_CHNG.Add
          .ShCode = rs.fields("shCode").Value
          .save
        End With
       
        
        rs.MoveNext
      Wend
      pb.Visible = False
  
  End If
 
  ProcessTDS = True
  Log "Free objects in memory"
  Manager.FreeAllInstanses
  Log "Load OK"
  Exit Function
bye:
  Log "Process TDS data error"
  Log Err.Description
  ProcessTDS = False
End Function

Private Function Code2OS(ByVal Code As String) As INV_OS.Application
  Dim rs As ADODB.Recordset
  Dim id As String
  Dim os As INV_OS.Application
  Set rs = Session.GetData("select instanceid from invos_code where visiblecode='" & Code & "'")
  If Not rs.EOF Then
    id = rs!InstanceID
    Set os = Manager.GetInstanceObject(id)
    Set Code2OS = os
  End If
  Exit Function

End Function

Private Function FindStatusByName(ByVal Name As String) As INVD_OSSTATUS
    Dim id As String
    Dim rs As ADODB.Recordset
    Dim dic As INV_DIC.Application
    Dim OSS As INV_DIC.INVD_OSSTATUS
    Set rs = Manager.ListInstances("", "INV_DIC")
    If Not rs.EOF Then
      id = rs!InstanceID
    Else
      id = CreateGUID2
      Manager.NewInstance id, "INV_DIC", "Справочник"
    End If
    Set dic = Manager.GetInstanceObject(id)
    
    Log "find status " & Name
    Set rs = Session.GetData("select * from INVD_OSSTATUS where name='" & Name & "'")
    If rs.EOF Then
        Log "Add new status"
        Set OSS = dic.INVD_OSSTATUS.Add
        
        OSS.Name = Name
        OSS.save
        
    Else
        Set OSS = dic.INVD_OSSTATUS.Item(rs!invd_OSSTATUSid)
    End If
    Log "Return status"
    Set FindStatusByName = OSS
    

End Function


Private Function FindOwnerByName(ByVal Name As String) As INVD_OWNER
    Dim id As String
    Dim rs As ADODB.Recordset
    Dim dic As INV_DIC.Application
    Dim OSS As INV_DIC.INVD_OWNER
    Set rs = Manager.ListInstances("", "INV_DIC")
    If Not rs.EOF Then
      id = rs!InstanceID
    Else
      id = CreateGUID2
      Manager.NewInstance id, "INV_DIC", "Справочник"
    End If
    Set dic = Manager.GetInstanceObject(id)
    Dim arr() As String
    arr = Split(Name, " ")
    
     
    If Session.IsMSSQL Then
      Set rs = Session.GetData("select  * from INVD_owner where isnull(FamiliName,'?') +' ' + isnull(Name,'?') +' ' +isnull(SurName,'?')+ ' ' ='" & Trim(Name) & " '")
    End If
    If Session.IsPOSTGRESQL Then
      Set rs = Session.GetData("select  * from INVD_owner where  COALESCE(cast(FamiliName as varchar),'?') ||' ' ||  COALESCE(cast(Name as varchar),'?') ||' ' || COALESCE(cast(SurName as varchar),'?') || ' ' ='" & Trim(Name) & " '")
    End If
    
    If rs.EOF Then
        Set OSS = dic.INVD_OWNER.Add
        If UBound(arr) >= 0 Then OSS.FamiliName = arr(0)
        If UBound(arr) >= 1 Then OSS.Name = arr(1)
        If UBound(arr) >= 2 Then OSS.SurName = arr(2)
        OSS.save
    Else
        Set OSS = dic.INVD_OWNER.Item(rs!INVD_OWNERid)
    End If
    Set FindOwnerByName = OSS
    

End Function

