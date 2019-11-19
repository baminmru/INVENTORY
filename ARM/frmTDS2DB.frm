VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTDS2DB 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Получение данных"
   ClientHeight    =   1095
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ProgressBar pb 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Получить данные"
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "frmTDS2DB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Cnn As cConnection
Private rs As cRecordset
Private irs As ADODB.Recordset
Private Sub CancelButton_Click()
Me.Hide
End Sub

Private Sub OKButton_Click()
    Dim fName As String
   
        fName = App.path & "\OMK.db3"
        Label1 = "Соединение с устройством"
        If RapiConnect Then
          
            
            
            Label1 = "Получение файла"
            RAPICopyCEFileToPC "/OMK.db3", fName
            
            If Not FileExists(fName) Then
                MsgBox "", vbInformation, "Ошибка создания файла данных"
                RapiDisconnect
                Label1 = ""
                Exit Sub
            End If
            Label1 = "Загрузка данных"
            LoadTDSData fName
            Label1 = "Удаление данных"
            CeDeleteFile "/OMK.db3"
            
            Label1 = "Отключение от устройства"
            RapiDisconnect
        Else
            MsgBox "Вставте  Терминал сбора данных в кредл.", vbInformation, "Ошибка передачи данных"
            Label1 = ""
            Exit Sub
        End If
        MsgBox "Данные инвентаризации загружены", vbOKOnly, "Загрузка данных"
        Label1 = ""
        Me.Hide
   

Me.Hide
End Sub


Private Sub LoadTDSData(ByVal fName As String)
  Set Cnn = New cConnection 'instantiate the Connection-Object
  If Not FileExists(fName) Then
    Exit Sub
  End If
  Cnn.OpenDB fName
  Set rs = Cnn.OpenRecordset("select count(*) from T")
  
  
  Dim cnt As Long
  
  cnt = rs.fields(0).Value
  
  pb.Min = 0
  pb.Max = cnt
  pb.Value = 0
  pb.Visible = True
  
  Set rs = Cnn.OpenRecordset("select * from T")
  
  Dim os As INV_OS.Application
  Dim osi As INV_OS.INVOS_INFO
  Dim inv As INV_INV.Application
  
  While Not rs.EOF
    pb.Value = pb.Value + 1
    Label1 = rs.fields(0).Value + " " + rs.fields(1).Value
    '"Create Table T(shCode Text, Status Text, CheckTime datetime, INVID TEXT, OSID TEXT)"
    
    Set inv = Manager.GetInstanceObject(rs.fields(3).Value)
    Set os = Manager.GetInstanceObject(rs.fields(4).Value)
    Set osi = os.INVOS_INFO.Item(1)
    With inv.INVI_DONE.Add
        Set .TheOS = osi
        .CheckDate = rs.fields(2).Value
        Set .OSStatus = FindStatusByName(rs.fields(1).Value)
        .save
    End With
     With os.INVOS_INV.Add
        Set .OSStatus = FindStatusByName(rs.fields(1).Value)
        .InvDate = rs.fields(2).Value
        Set .Inventory = inv.invi_DEF.Item(1)
        .save
     End With
    
    ' найти ОС по коду
    ' найти инвентаризацию
    ' найти в справочнике состояние по имени
    ' вставить запись в инвентаризацию
    
    'MsgBox Label1
    rs.MoveNext
  Wend
  pb.Visible = False
  
  
End Sub


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
    
    Set rs = Session.GetData("select * from INVD_OSSTATUS where name='" & Name & "'")
    If rs.EOF Then
        Set OSS = dic.INVD_OSSTATUS.Add
        OSS.Name = Name
        OSS.save
    Else
        Set OSS = dic.INVD_OSSTATUS.Item(rs!invd_OSSTATUSid)
    End If
    Set FindStatusByName = OSS
    

End Function
