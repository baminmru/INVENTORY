VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Главное окно"
   ClientHeight    =   6705
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   8760
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer MenuTimer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2355
      Top             =   840
   End
   Begin VB.Timer Timer2 
      Interval        =   60000
      Left            =   1665
      Top             =   855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   1080
      Top             =   840
   End
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   240
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuJRNL 
      Caption         =   "Журналы"
      Begin VB.Menu mnuINV_INV 
         Caption         =   "Инвентаризация"
         Begin VB.Menu mnuAllINV_INV 
            Caption         =   "Инвентаризация - все состояния"
         End
         Begin VB.Menu mnuINV_INV_1 
            Caption         =   "Инвентаризация :Утверждена"
         End
         Begin VB.Menu mnuINV_INV_2 
            Caption         =   "Инвентаризация :Инвентаризация завершена"
         End
         Begin VB.Menu mnuINV_INV_3 
            Caption         =   "Инвентаризация :Оформляется"
         End
         Begin VB.Menu mnuINV_INV_4 
            Caption         =   "Инвентаризация :Идет инвентаризация"
         End
      End
      Begin VB.Menu mnuINV_OS 
         Caption         =   "Карточка основного средства"
         Begin VB.Menu mnuAllINV_OS 
            Caption         =   "Карточка основного средства - все состояния"
         End
         Begin VB.Menu mnuINV_OS_1 
            Caption         =   "Карточка основного средства :В ремонте"
         End
         Begin VB.Menu mnuINV_OS_2 
            Caption         =   "Карточка основного средства :В лизинге"
         End
         Begin VB.Menu mnuINV_OS_3 
            Caption         =   "Карточка основного средства :Оформляется"
         End
         Begin VB.Menu mnuINV_OS_4 
            Caption         =   "Карточка основного средства :В аренде"
         End
         Begin VB.Menu mnuINV_OS_5 
            Caption         =   "Карточка основного средства :На модернизации"
         End
         Begin VB.Menu mnuINV_OS_6 
            Caption         =   "Карточка основного средства :В эксплуатации"
         End
         Begin VB.Menu mnuINV_OS_7 
            Caption         =   "Карточка основного средства :Списано"
         End
         Begin VB.Menu mnuINV_OS_8 
            Caption         =   "Карточка основного средства :На консервации"
         End
         Begin VB.Menu mnuINV_OS_6a 
            Caption         =   "Карточка основного средства :Прошли инвентаризацию"
         End
         Begin VB.Menu mnuINV_OS_7a 
            Caption         =   "Карточка основного средства :Не прошли инвентаризацию"
         End
      End
      Begin VB.Menu mnuINVF 
         Caption         =   "Загрузка данных"
      End
      Begin VB.Menu mnuLog 
         Caption         =   "Протокол"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Выход"
      End
   End
   Begin VB.Menu mnuOperation 
      Caption         =   "Операции"
      Begin VB.Menu mnuLoadOS 
         Caption         =   "Загрузка данных ОС"
      End
      Begin VB.Menu mnuLoadMat 
         Caption         =   "Загрузка данных  по материалам"
      End
      Begin VB.Menu mnuLoadTech 
         Caption         =   "Загрузка технической информации"
      End
      Begin VB.Menu mnuLoad2TDS 
         Caption         =   "Обмен данными с  ТДС"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "Отчеты"
      Begin VB.Menu mnuPrintSHCode 
         Caption         =   "Печать штрихкодов"
      End
      Begin VB.Menu mnuRptShCode52 
         Caption         =   "Печать штрихкодов средние"
      End
      Begin VB.Menu mnuRtSH2 
         Caption         =   "Печать штрихкодов малые"
      End
      Begin VB.Menu mnuRptInv 
         Caption         =   "Справка об инвентаризации"
      End
      Begin VB.Menu mnuRptBadInv 
         Caption         =   "Справка об отсутствии ОС"
      End
      Begin VB.Menu mnuRptVed 
         Caption         =   "Сличительная ведомость"
      End
      Begin VB.Menu mnuRptUnknown 
         Caption         =   "Неучтенные объекты"
      End
      Begin VB.Menu mnuRptOS 
         Caption         =   "Отчет по ОС"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Настройки"
      Begin VB.Menu mnuINV_DIC 
         Caption         =   "Справочник"
      End
      Begin VB.Menu mnuPrinters 
         Caption         =   "Принтера"
      End
      Begin VB.Menu mnuOptMat 
         Caption         =   "Нумерация материалов"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLoadPers 
         Caption         =   "Загрузка справочника владельцев (Excel)"
      End
      Begin VB.Menu mnuPortalCFG 
         Caption         =   "Параметры для соединения с порталом"
      End
      Begin VB.Menu mnuLoadPortal 
         Caption         =   "Загрузка и актуализация владельцев (Портал)"
      End
      Begin VB.Menu mnuCleadOwners 
         Caption         =   "Убрать лишние записи из справочника владельцев"
      End
      Begin VB.Menu mnuSep8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMaintainSetup 
         Caption         =   "Настройка обслуживания базы"
      End
      Begin VB.Menu mnuVacuum 
         Caption         =   "Обслуживание базы"
      End
   End
   Begin VB.Menu mnuWin 
      Caption         =   "Окно"
      WindowList      =   -1  'True
      Begin VB.Menu mnuAbout 
         Caption         =   "О программе"
      End
      Begin VB.Menu mnuCascade 
         Caption         =   "Каскад"
      End
      Begin VB.Menu mnuTileVert 
         Caption         =   "Разложить вертикально"
      End
      Begin VB.Menu mnuTileHor 
         Caption         =   "Разложить горизонтально"
      End
      Begin VB.Menu mnuArrangeIcon 
         Caption         =   "Разложить иконки"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Dim frmFind As Form
Dim frmFindFT As Form

Dim inTimer1 As Boolean
Dim inTimer2 As Boolean
Dim OnLoad As Boolean
Dim DelayedCommand As String

Private WithEvents osRPt As MTZReportHelper.WordHelper
Attribute osRPt.VB_VarHelpID = -1
Private ObjectToReport As INV_OS.Application


Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)

Dim WithEvents jfmnuAllINV_INV As frmJournalShow
Attribute jfmnuAllINV_INV.VB_VarHelpID = -1

Dim WithEvents jfmnuINV_INV_1 As frmJournalShow
Attribute jfmnuINV_INV_1.VB_VarHelpID = -1

Dim WithEvents jfmnuINV_INV_2 As frmJournalShow
Attribute jfmnuINV_INV_2.VB_VarHelpID = -1

Dim WithEvents jfmnuINV_INV_3 As frmJournalShow
Attribute jfmnuINV_INV_3.VB_VarHelpID = -1

Dim WithEvents jfmnuINV_INV_4 As frmJournalShow
Attribute jfmnuINV_INV_4.VB_VarHelpID = -1

Dim WithEvents jfmnuINVF As frmJournalShow
Attribute jfmnuINVF.VB_VarHelpID = -1

Dim WithEvents jfmnuAllINV_OS As frmJournalShow
Attribute jfmnuAllINV_OS.VB_VarHelpID = -1

Dim WithEvents jfmnuINV_OS_1 As frmJournalShow
Attribute jfmnuINV_OS_1.VB_VarHelpID = -1

Dim WithEvents jfmnuINV_OS_2 As frmJournalShow
Attribute jfmnuINV_OS_2.VB_VarHelpID = -1

Dim WithEvents jfmnuINV_OS_3 As frmJournalShow
Attribute jfmnuINV_OS_3.VB_VarHelpID = -1

Dim WithEvents jfmnuINV_OS_4 As frmJournalShow
Attribute jfmnuINV_OS_4.VB_VarHelpID = -1


Dim WithEvents jfmnuINV_OS_5 As frmJournalShow
Attribute jfmnuINV_OS_5.VB_VarHelpID = -1

Dim WithEvents jfmnuINV_OS_6 As frmJournalShow
Attribute jfmnuINV_OS_6.VB_VarHelpID = -1
Dim WithEvents jfmnuINV_OS_7 As frmJournalShow
Attribute jfmnuINV_OS_7.VB_VarHelpID = -1
Dim WithEvents jfmnuINV_OS_8 As frmJournalShow
Attribute jfmnuINV_OS_8.VB_VarHelpID = -1

Dim WithEvents jfmnuINV_OS_OK As frmJournalShow
Attribute jfmnuINV_OS_OK.VB_VarHelpID = -1
Dim WithEvents jfmnuINV_OS_BAD As frmJournalShow
Attribute jfmnuINV_OS_BAD.VB_VarHelpID = -1


Dim WithEvents jfmnuINV_NUM As frmJournalShow
Attribute jfmnuINV_NUM.VB_VarHelpID = -1

Public RptInvOK As ReportShow
Public RptInvBAD As ReportShow
Public RptSHCODE As ReportShow
Public RptSLICH As ReportShow
Public RptInvUnk As ReportShow



Public Sub On_Load()
   Me.Caption = App.FileDescription & " (" & Site & "\" & MyRole.Name & "\" & MyUser.brief & ")"
   On Error Resume Next
   'If command$ <> "DEBUG" Then
     Dim c As Control
     For Each c In Me.Controls
      If TypeName(c) = "Menu" Then
         
        If CheckMenu(c.Name) = RoleMenuStatus_Hidden Then
          c.Visible = False
        Else
          frmSplash.lblWarning = "Инициализация меню: " & c.Caption
          DoEvents
        End If
      End If
     Next
  'End If
   Manager.FreeAllInstanses
End Sub

Private Sub jfmnuAllINV_INV_OnRun(ByVal RowIndex As Long, usedefaut As Boolean, Refesh As Boolean)
 usedefaut = False
  Dim f As frmInvService
  Set f = New frmInvService
  Dim id As String
  id = jfmnuAllINV_INV.jv.RowInstanceID(RowIndex)
  f.id = id
  f.Show vbModal
  If f.OK Then
   
    ExcludeObjects id, f.txtINV.Tag, f.cmbExclude.ListIndex
  End If
  Unload f
  Set f = Nothing
End Sub

Private Sub jfmnuAllINV_OS_OnRun(ByVal RowIndex As Long, usedefaut As Boolean, Refesh As Boolean)
If MsgBox("Показать протокол работы с документом?", vbQuestion + vbYesNo, "Уточнение задачи") = vbYes Then
  usedefaut = False
  Dim f As frmLog
  Set f = New frmLog
  Set f.Item = Manager.GetInstanceObject(jfmnuAllINV_OS.jv.RowInstanceID(jfmnuAllINV_OS.jv.Row))
  If Not f.Item Is Nothing Then
    f.Show
  Else
    Set f = Nothing
    usedefaut = True
  End If
  
End If
End Sub

Private Sub jfmnuINV_INV_1_OnRun(ByVal RowIndex As Long, usedefaut As Boolean, Refesh As Boolean)
  usedefaut = False
  Dim f As frmInvService
  Set f = New frmInvService
  Dim id As String
  id = jfmnuINV_INV_1.jv.RowInstanceID(RowIndex)
  f.id = id

  f.Show vbModal
  If f.OK Then

    ExcludeObjects id, f.txtINV.Tag, f.cmbExclude.ListIndex
  End If
  Unload f
  Set f = Nothing
End Sub

Private Sub jfmnuINV_INV_3_OnRun(ByVal RowIndex As Long, usedefaut As Boolean, Refesh As Boolean)
  usedefaut = False
  Dim f As frmInvService
  Set f = New frmInvService
  Dim id As String
  id = jfmnuINV_INV_3.jv.RowInstanceID(RowIndex)
  f.id = id
  f.Show vbModal
  If f.OK Then
    ExcludeObjects id, f.txtINV.Tag, f.cmbExclude.ListIndex
  End If
  Unload f
  Set f = Nothing

End Sub

Private Sub jfmnuINV_OS_6_OnRun(ByVal RowIndex As Long, usedefaut As Boolean, Refesh As Boolean)
    
    Dim os As INV_OS.Application
    Dim ostype As INV_DIC.INVD_OSTYPE
    Dim id As String
    id = jfmnuINV_OS_6.jv.RowInstanceID(RowIndex)
    Set os = Manager.GetInstanceObject(id)
    Set ostype = os.INVOS_INFO.Item(1).ostype
    If ostype.ShowTech = Boolean_Da Then
        If MsgBox("Загрузить технические данные для данного компьютера ?", vbYesNo + vbQuestion, "Подтвердите") = vbYes Then
         usedefaut = False
         LoadTech os
        Else
         usedefaut = True
        End If
    Else
        usedefaut = True
    End If
    
    
    
End Sub

Private Sub LoadTech(os As INV_OS.Application)
    Dim s As String
    Dim f As frmGetFile
    Dim ff As Integer
    Dim arr() As String
    Dim sPath As String
    sPath = os.INVOS_INFO.Item(1).TechFilePath
    
    If sPath = "" Then
    
      Set f = New frmGetFile
      f.Show vbModal
      If f.OK Then
        sPath = f.txtPath
      End If
      Unload f
      Set f = Nothing
    End If
    
    
    Dim md5 As String
    If sPath <> "" Then
        Dim ccc As CMD5
        Set ccc = New CMD5
        md5 = ccc.FileMD5(sPath)
        Set ccc = Nothing
        If IsFileLoaded(sPath, md5, "ТЕХНИЧЕСКАЯ ИНФОРМАЦИЯ") Then
          'MsgBox "Файл уже загружен в систему"
          Exit Sub
        End If
    
    
        ff = FreeFile
        On Error GoTo bye
        Open sPath For Input As #ff
        os.INVOS_INFO.Item(1).TechFilePath = sPath
        os.INVOS_INFO.Item(1).save
        s = input(LOF(ff), ff)
        Close #ff
        arr = Split(s, vbCrLf)
        Dim i As Integer
        For i = LBound(arr) To UBound(arr)
            If Left(arr(i), 1) = "[" Then
                If Not fProg Is Nothing Then
                  fProg.NextVal os.brief & "->" & arr(i)
                End If
                If arr(i) = "[Info]" Then
                 
                    Call LoadOSInfo(arr, i, os)
                End If
                If arr(i) = "[Computer]" Then
                 Call LoadOSComputer(arr, i, os)
                End If
                If arr(i) = "[Current_Config]" Then
                 Call LoadOSCurCFG(arr, i, os)
                End If
                If arr(i) = "[Windows_Soft]" Then
                 Call LoadOSWinSoft(arr, i, os)
                End If
                If arr(i) = "[Windows_Devices]" Then
                 Call LoadOSWinDev(arr, i, os)
                End If
                
                If arr(i) = "[Config_changes]" Then
                 Call LoadCC(arr, i, os)
                End If
                 If arr(i) = "[Hardware]" Then
                 Call LoadHard(arr, i, os)
                End If
                
            End If
        Next
        RegisterFile sPath, md5, "ТЕХНИЧЕСКАЯ ИНФОРМАЦИЯ"
    End If
    
bye:
    If Not fProg Is Nothing Then
            fProg.NextVal os.brief & "->" & Err.Description
            Err.Clear
            If MsgBox("Удалить данные о пути к файлу с Тех. информацией?", vbQuestion + vbYesNo, "Ошибка загрузки данных") = vbYes Then
              os.INVOS_INFO.Item(1).TechFilePath = ""
              os.INVOS_INFO.Item(1).save
            End If
    End If

End Sub

Private Sub LoadOSInfo(arr() As String, StartIdx As Integer, os As INV_OS.Application)
    If os.INVOS_TECH.Count = 1 Then
        os.INVOS_TECH.Delete 1
    End If
    Dim bbb() As String
    Dim i As Integer
    With os.INVOS_TECH.Add
        For i = StartIdx + 1 To UBound(arr)
         If Left(arr(i), 1) = "[" Then
          .save
          Exit For
         End If
         bbb = Split(arr(i), "=")
         Select Case bbb(0)
         Case "MAC_Addr"
            .MAC_Addr = bbb(1)
         Case "Current_User_Name"
         .Current_User_Name = bbb(1)
         Case "Computer_Name"
         .Computer_Name = bbb(1)
         Case "IP_Addr"
         .IP_Addr = bbb(1)
         Case "System"
         .System = bbb(1)
         Case "Record_Date"
         .Record_Date = bbb(1)
         End Select
         
        Next
    End With
End Sub

Private Sub LoadOSComputer(arr() As String, StartIdx As Integer, os As INV_OS.Application)
 Dim bbb() As String
    Dim i As Integer
        For i = StartIdx + 1 To UBound(arr)
         If Left(arr(i), 1) = "[" Then
          Exit For
         End If
         bbb = Split(arr(i), "=")
            With os.INVOS_TECH.Item(1).INVOS_COMP.Add
            .Name = bbb(0)
            .ParamValue = bbb(1)
            .save
            End With
        Next
    
End Sub

Private Sub LoadOSCurCFG(arr() As String, StartIdx As Integer, os As INV_OS.Application)
Dim bbb() As String
    Dim i As Integer
        For i = StartIdx + 1 To UBound(arr)
         If Left(arr(i), 1) = "[" Then
          Exit For
         End If
         bbb = Split(arr(i), "=")
            With os.INVOS_TECH.Item(1).INVOS_CURCFG.Add
            .Name = bbb(0)
            .ParamValue = bbb(1)
            .save
            End With
        Next
End Sub

Private Sub LoadOSWinSoft(arr() As String, StartIdx As Integer, os As INV_OS.Application)
Dim bbb() As String
    Dim i As Integer
        For i = StartIdx + 1 To UBound(arr)
         If Left(arr(i), 1) = "[" Then
          Exit For
         End If
         bbb = Split(arr(i), "=")
            With os.INVOS_TECH.Item(1).INVOS_WINSOFT.Add
            .Name = bbb(0)
            .ParamValue = bbb(1)
            .save
            End With
        Next
End Sub


Private Sub LoadOSWinDev(arr() As String, StartIdx As Integer, os As INV_OS.Application)
Dim bbb() As String
    Dim i As Integer
        For i = StartIdx + 1 To UBound(arr)
         If Left(arr(i), 1) = "[" Then
          Exit For
         End If
         bbb = Split(arr(i), "=")
            With os.INVOS_TECH.Item(1).INVOS_DEVICES.Add
            .Name = bbb(0)
            .ParamValue = bbb(1)
            .save
            End With
        Next
End Sub

Private Sub LoadCC(arr() As String, StartIdx As Integer, os As INV_OS.Application)
On Error Resume Next
Dim bbb() As String
Dim ccc() As String
    Dim i As Integer
        For i = StartIdx + 1 To UBound(arr)
         If Left(arr(i), 1) = "[" Then
          Exit For
         End If
         bbb = Split(arr(i), "=")
         ccc = Split(bbb(0), " ")
            With os.INVOS_TECH.Item(1).INVOS_CFGCH.Add
            .ChDate = CDate(ccc(0) & " " & ccc(1))
            .ChNum = ccc(2)
            .TheValue = bbb(1)
            .save
            End With
        Next
End Sub

Private Sub LoadHard(arr() As String, StartIdx As Integer, os As INV_OS.Application)
Dim bbb() As String
    Dim i As Integer
        For i = StartIdx + 1 To UBound(arr)
         If Left(arr(i), 1) = "[" Then
          Exit For
         End If
         If arr(i) <> "" Then
         bbb = Split(arr(i), "=")
            With os.INVOS_TECH.Item(1).INVOS_HARD.Add
            .Name = bbb(0)
            .ParamValue = bbb(1)
            .save
            End With
        End If
        Next
End Sub

Private Sub jfmnuINV_OS_7_OnRun(ByVal RowIndex As Long, usedefaut As Boolean, Refesh As Boolean)
usedefaut = False

If MsgBox("Вернуть в эксплуатацию все списанные Основные средства ?", vbYesNo, "Подтвердите операцию") = vbYes Then
  Session.GetData ("update instance set status='{8AD15E54-CF87-4FCF-8A1E-A85336E23C73}' where instanceid in (" & _
  "select instanceid from v_autoinvos_info where intsancestatusid='{166D4978-0C4C-4575-8192-B251AC113781}' and invos_info_ismaterial_val=0)")
  Refesh = True
End If

If MsgBox("Вернуть в эксплуатацию все списанные Материалы ?", vbYesNo, "Подтвердите операцию") = vbYes Then
  Session.GetData ("update instance set status='{8AD15E54-CF87-4FCF-8A1E-A85336E23C73}' where instanceid in (" & _
  "select instanceid from v_autoinvos_info where intsancestatusid='{166D4978-0C4C-4575-8192-B251AC113781}' and invos_info_ismaterial_val=-1)")
  Refesh = True
End If


End Sub

Private Sub jfmnuINV_OS_BAD_OnInit(bAdd As Boolean, bEdit As Boolean, bRun As Boolean, bDel As Boolean, bFilter As Boolean)
bAdd = False
bDel = False
bFilter = False

End Sub

Private Sub jfmnuINV_OS_OK_OnInit(bAdd As Boolean, bEdit As Boolean, bRun As Boolean, bDel As Boolean, bFilter As Boolean)
bAdd = False
bDel = False
bFilter = False
End Sub

Private Sub jfmnuINVF_OnInit(bAdd As Boolean, bEdit As Boolean, bRun As Boolean, bDel As Boolean, bFilter As Boolean)
bAdd = False
bDel = False
End Sub

Private Sub MDIForm_Load()
On_Load
End Sub

Private Sub mdiForm_Unload(Cancel As Integer)
On Error Resume Next

' whait for finalize timer loops
inTimer1 = True
Me.Timer1.Enabled = False

inTimer2 = True
Me.Timer2.Enabled = False


Timer1.Enabled = False
Timer2.Enabled = False

On Error Resume Next

' unload all dynamically created journals and reports
UnloadObjects

If Not frmFind Is Nothing Then
  Unload frmFind
End If
Set frmFind = Nothing

If Not frmFindFT Is Nothing Then
  Unload frmFindFT
End If
Set frmFindFT = Nothing



Dim f As Form
For Each f In Forms
  If f.MDIChild = True Then
    On Error Resume Next
    'Call f.Controls.Item(0).object.Init(Nothing, Nothing, False, Nothing)
    Unload f
  End If
Next

  For Each f In Forms
      On Error Resume Next
      Debug.Print f.Name
  Next
  
  
  Set MyRole = Nothing
  Set MyUser = Nothing
  Set usr = Nothing


  Session.Logout
  Set Session = Nothing
  Manager.CloseClass
  Set Manager = Nothing

  If Command$ <> "DEBUG" Then
   TerminateProcess GetCurrentProcess, 0
  'Else
  ' End
  End If
End Sub


Private Sub mnuAbout_Click()
frmAbout.Show vbModal, Me
End Sub



Private Sub mnuArrangeIcon_Click()
  Me.Arrange vbArrangeIcons
End Sub

Private Sub mnuCascade_Click()
  Me.Arrange vbCascade
End Sub

Private Sub mnuCleadOwners_Click()
  Dim mc As String
  Dim rs As ADODB.Recordset
  Dim dic As INV_DIC.Application
  Dim i As Long
  Dim j As Long
  Dim okINI As String
  Dim okOwner As INVD_OWNER
  If MsgBox("Удалить лишние записи из справочника владельцев ?", vbQuestion + vbYesNo, "Подтверждение") = vbYes Then
  
  
    mc = Me.Caption
    On Error Resume Next
      
    Dim id As String
    Set rs = Manager.ListInstances("", "INV_DIC")
    If Not rs.EOF Then
      Set dic = Manager.GetInstanceObject(rs!InstanceID)
    Else
    Exit Sub
    End If
    
    Set rs = Session.GetData("select lower(FamiliName) as FamiliName,lower(Name) as Name , lower(SurName) as SurName from INVD_owner  group by lower(FamiliName),lower(Name),lower(SurName)  having count(*) >1 order by lower(FamiliName),lower(Name),lower(SurName) ")
    While Not rs.EOF
      Set okOwner = Nothing
      For i = 1 To dic.INVD_OWNER.Count
      
        With dic.INVD_OWNER.Item(i)
          If LCase(.FamiliName) = LCase("" & rs!FamiliName) And LCase(.Name) = LCase("" & rs!Name) And LCase(.SurName) = LCase("" & rs!SurName) Then
            Set okOwner = dic.INVD_OWNER.Item(i)
            Me.Caption = mc & " Ждите. Удаление лишних записей для:" & okOwner.brief
            DoEvents
            Exit For
          End If
        End With
        
      Next
      
      If Not okOwner Is Nothing Then
        For i = 1 To dic.INVD_OWNER.Count
          With dic.INVD_OWNER.Item(i)
            If LCase(.FamiliName) = LCase(okOwner.FamiliName) And LCase(.Name) = LCase(okOwner.Name) And LCase(.SurName) = LCase(okOwner.SurName) Then
              If .id <> okOwner.id Then
                Session.GetData "update invos_HIST set  MatOtv ='" & okOwner.id & "' where MatOtv='" & .id & "'"
                Session.GetData "update INVOS_PLACE set  MatOtv ='" & okOwner.id & "' where MatOtv='" & .id & "'"
                Session.GetData "update INVI_DEF set  MatOtv ='" & okOwner.id & "' where MatOtv='" & .id & "'"
                Session.GetData "update invos_HIST set  TheOwner ='" & okOwner.id & "' where TheOwner='" & .id & "'"
                Session.GetData "update INVOS_PLACE set  TheOwner ='" & okOwner.id & "' where TheOwner='" & .id & "'"
                Session.GetData "update INVI_DEF set  TheOwner ='" & okOwner.id & "' where TheOwner='" & .id & "'"
                Session.GetData "delete from invd_owner where invd_ownerid ='" & .id & "'"
              End If
            End If
          End With
        Next
        dic.INVD_OWNER.Refresh
      End If
      
      
      rs.MoveNext
    Wend
    Set rs = Nothing
    
    
    If MsgBox("Пробуем объединять записи с инициалами и полным именем?", vbQuestion + vbYesNo, "Подтверждение") = vbYes Then
     Dim K As Long
     K = 0
    Set rs = Session.GetData("select lower(FamiliName) FamiliName  from INVD_owner  group by lower(FamiliName) having count(*) >1 order by lower(FamiliName)")
    While Not rs.EOF
      K = K + 1
      Set okOwner = Nothing
      For i = 1 To dic.INVD_OWNER.Count
       Set okOwner = Nothing
       If i <= dic.INVD_OWNER.Count Then
        With dic.INVD_OWNER.Item(i)
          If LCase(.FamiliName) = LCase("" & rs!FamiliName) And InStr(.Name, ".") = 0 And .Name <> "" And InStr(.SurName, ".") = 0 And .SurName <> "" Then
            Set okOwner = dic.INVD_OWNER.Item(i)
            Me.Caption = mc & " Ждите. Попытка сборки идентичных записей для: " & okOwner.brief & "(" & K & ")"
            DoEvents
          End If
        End With
      
        If Not okOwner Is Nothing Then
          okINI = Left(okOwner.Name, 1) & Left(okOwner.SurName, 1)
          
          For j = 1 To dic.INVD_OWNER.Count
            If j <= dic.INVD_OWNER.Count Then
            With dic.INVD_OWNER.Item(j)
              If LCase(.FamiliName) = LCase(okOwner.FamiliName) Then
                If .id <> okOwner.id Then
                
                
                
                  If (LCase(okINI) = Replace(LCase(Trim(.Name)), ".", "")) Or (LCase(okINI) = Replace(LCase(Trim(.Name & .SurName)), ".", "")) Then
                    Me.Caption = mc & " Ждите. Попытка сборки идентичных записей для:" & okOwner.brief & "(" & K & ")" + " & .brief + " / ""
                    DoEvents
                    Session.GetData "update invos_HIST set  MatOtv ='" & okOwner.id & "' where MatOtv='" & .id & "'"
                    Me.Caption = mc & " Ждите. Попытка сборки идентичных записей для:" & okOwner.brief & "(" & K & ")" & " + " & .brief + " -"
                    DoEvents
                    Session.GetData "update INVOS_PLACE set  MatOtv ='" & okOwner.id & "' where MatOtv='" & .id & "'"
                    Me.Caption = mc & " Ждите. Попытка сборки идентичных записей для:" & okOwner.brief & "(" & K & ")" & " + " & .brief + " \"
                    DoEvents
                    
                    Session.GetData "update INVI_DEF set  MatOtv ='" & okOwner.id & "' where MatOtv='" & .id & "'"
                    Me.Caption = mc & " Ждите. Попытка сборки идентичных записей для:" & okOwner.brief & "(" & K & ")" & " + " & .brief + " |"
                    DoEvents
                    
                    Session.GetData "update invos_HIST set  TheOwner ='" & okOwner.id & "' where TheOwner='" & .id & "'"
                    Me.Caption = mc & " Ждите. Попытка сборки идентичных записей для:" & okOwner.brief & "(" & K & ")" & " + " & .brief + " /"
                    DoEvents
                    
                    Session.GetData "update INVOS_PLACE set  TheOwner ='" & okOwner.id & "' where TheOwner='" & .id & "'"
                    Me.Caption = mc & " Ждите. Попытка сборки идентичных записей для:" & okOwner.brief & "(" & K & ")" & " + " & .brief + " -"
                    DoEvents
                    
                    Session.GetData "update INVI_DEF set  TheOwner ='" & okOwner.id & "' where TheOwner='" & .id & "'"
                    Me.Caption = mc & " Ждите. Попытка сборки идентичных записей для:" & okOwner.brief & "(" & K & ")" & " + " & .brief + " \"
                    DoEvents
                    
                    Session.GetData "delete from invd_owner where invd_ownerid ='" & .id & "'"
                    Me.Caption = mc & " Ждите. Попытка сборки идентичных записей для:" & okOwner.brief & "(" & K & ")"
                    DoEvents
                    dic.INVD_OWNER.Refresh
                  End If
                  
                  If ((LCase(okOwner.Name) = LCase(Trim(.Name))) And .SurName = "") Or _
                      ((LCase(okOwner.Name) = LCase(Trim(.SurName))) And .Name = "") Or _
                      (.Name = "" And .SurName = "") _
                  Then
                    Me.Caption = mc & " Ждите. Попытка сборки идентичных записей для:" & okOwner.brief & "(" & K & ")" + " & .brief + " / ""
                    DoEvents
                    Session.GetData "update invos_HIST set  MatOtv ='" & okOwner.id & "' where MatOtv='" & .id & "'"
                    Me.Caption = mc & " Ждите. Попытка сборки идентичных записей для:" & okOwner.brief & "(" & K & ")" & " + " & .brief + " -"
                    DoEvents
                    Session.GetData "update INVOS_PLACE set  MatOtv ='" & okOwner.id & "' where MatOtv='" & .id & "'"
                    Me.Caption = mc & " Ждите. Попытка сборки идентичных записей для:" & okOwner.brief & "(" & K & ")" & " + " & .brief + " \"
                    DoEvents
                    
                    Session.GetData "update INVI_DEF set  MatOtv ='" & okOwner.id & "' where MatOtv='" & .id & "'"
                    Me.Caption = mc & " Ждите. Попытка сборки идентичных записей для:" & okOwner.brief & "(" & K & ")" & " + " & .brief + " |"
                    DoEvents
                    
                    Session.GetData "update invos_HIST set  TheOwner ='" & okOwner.id & "' where TheOwner='" & .id & "'"
                    Me.Caption = mc & " Ждите. Попытка сборки идентичных записей для:" & okOwner.brief & "(" & K & ")" & " + " & .brief + " /"
                    DoEvents
                    
                    Session.GetData "update INVOS_PLACE set  TheOwner ='" & okOwner.id & "' where TheOwner='" & .id & "'"
                    Me.Caption = mc & " Ждите. Попытка сборки идентичных записей для:" & okOwner.brief & "(" & K & ")" & " + " & .brief + " -"
                    DoEvents
                    
                    Session.GetData "update INVI_DEF set  TheOwner ='" & okOwner.id & "' where TheOwner='" & .id & "'"
                    Me.Caption = mc & " Ждите. Попытка сборки идентичных записей для:" & okOwner.brief & "(" & K & ")" & " + " & .brief + " \"
                    DoEvents
                    
                    Session.GetData "delete from invd_owner where invd_ownerid ='" & .id & "'"
                    Me.Caption = mc & " Ждите. Попытка сборки идентичных записей для:" & okOwner.brief & "(" & K & ")"
                    DoEvents
                    dic.INVD_OWNER.Refresh
                  End If
                End If
              End If
            End With
            End If
          Next
          
        End If
        End If
      Next
      dic.INVD_OWNER.Refresh
      rs.MoveNext
    Wend
    Set rs = Nothing
    End If
    
    If MsgBox("Удалить позиции с короткими именами из справочника владельцев?", vbQuestion + vbYesNo, "Подтверждение") = vbYes Then
     For i = 1 To dic.INVD_OWNER.Count
          With dic.INVD_OWNER.Item(i)
            If InStr(.Name, ".") > 0 Or InStr(.SurName, ".") > 0 Then
                Dim OK As Boolean
                OK = True
                Me.Caption = mc & " Ждите. Проверка ссылок для:" & .brief
                DoEvents
                Set rs = Session.GetData("SELECT COUNT(*) CNT FROM  invos_HIST  where MatOtv='" & .id & "'")
                If rs!cnt > 0 Then OK = False
                Set rs = Session.GetData("SELECT COUNT(*) CNT FROM  INVOS_PLACE  where MatOtv='" & .id & "'")
                If rs!cnt > 0 Then OK = False
                Set rs = Session.GetData("SELECT COUNT(*) CNT FROM  INVI_DEF  where MatOtv='" & .id & "'")
                If rs!cnt > 0 Then OK = False
                Set rs = Session.GetData("SELECT COUNT(*) CNT FROM  invos_HIST  where TheOwner='" & .id & "'")
                If rs!cnt > 0 Then OK = False
                Set rs = Session.GetData("SELECT COUNT(*) CNT FROM  INVOS_PLACE  where TheOwner='" & .id & "'")
                If rs!cnt > 0 Then OK = False
                Set rs = Session.GetData("SELECT COUNT(*) CNT FROM  INVI_DEF  where TheOwner='" & .id & "'")
                If rs!cnt > 0 Then OK = False
                If OK Then
                  Me.Caption = mc & " Ждите. Удаление строки:" & .brief
                  DoEvents
                  Session.GetData "delete from invd_owner where invd_ownerid ='" & .id & "'"
                End If
             End If
          End With
        Next
        dic.INVD_OWNER.Refresh
    End If
    
   
    
    
   
    
    
    
    Me.Caption = mc
    MsgBox "Исправление справочника владельцев завершено"
  End If

End Sub

Private Sub mnuExit_Click()
  Unload Me
End Sub

Private Sub mnuINV_OS_6a_Click()
  Dim journal As Object
    On Error Resume Next
    If jfmnuINV_OS_OK Is Nothing Then
      Set jfmnuINV_OS_OK = New frmJournalShow
      Set journal = Manager.GetInstanceObject("{5CE3CC0B-5224-4145-B43F-6A29CC390C17}")
      Manager.LockInstanceObject journal.id
      Set jfmnuINV_OS_OK.jv.journal = journal
      AddColors jfmnuINV_OS_OK.jv
      jfmnuINV_OS_OK.jv.OpenModal = False
      jfmnuINV_OS_OK.Caption = "Карточка основного средства - Прошли инвентаризацию"
      Me.MousePointer = vbHourglass
      DoEvents
    End If
    Dim f As String
    f = "1=1"
    Dim frm As frmDates
      Set frm = New frmDates
      frm.Show vbModal
      If frm.OK Then
        If Session.IsMSSQL Then
            f = f & " and  instanceid in (select instanceid from invos_inv where invdate >=" & MakeMSSQLDate(frm.dtpFrom.Value) & " and invdate <= " & MakeMSSQLDate(frm.dtpTo.Value) & ")"
        End If
        If Session.IsPOSTGRESQL Then
            f = f & " and  instanceid in (select instanceid from invos_inv where invdate >=" & MakePGSQLDate(frm.dtpFrom.Value) & " and invdate <= " & MakePGSQLDate(frm.dtpTo.Value) & ")"
        End If
        
      Else
        If Session.IsMSSQL Then
            f = f & " and  instanceid in (select instanceid from invos_inv )"
        End If
        If Session.IsPOSTGRESQL Then
            f = f & " and  instanceid in (select instanceid from invos_inv )"
        End If
      
      End If
      
    jfmnuINV_OS_OK.jv.Filter.Add "AUTOINVOS_INFO", f
    jfmnuINV_OS_OK.jv.Refresh
    Me.MousePointer = vbNormal
    jfmnuINV_OS_OK.Show
    jfmnuINV_OS_OK.WindowState = 0
    jfmnuINV_OS_OK.ZOrder 0
End Sub

Private Sub mnuINV_OS_7a_Click()
    Dim journal As Object
    On Error Resume Next
    If jfmnuINV_OS_BAD Is Nothing Then
      Set jfmnuINV_OS_BAD = New frmJournalShow
      Set journal = Manager.GetInstanceObject("{5CE3CC0B-5224-4145-B43F-6A29CC390C17}")
      Manager.LockInstanceObject journal.id
      Set jfmnuINV_OS_BAD.jv.journal = journal
      AddColors jfmnuINV_OS_BAD.jv
      jfmnuINV_OS_BAD.jv.OpenModal = False
      jfmnuINV_OS_BAD.Caption = "Карточка основного средства - не прошли инвентаризацию"
      Me.MousePointer = vbHourglass
    End If
    DoEvents
    Dim f As String
    f = "1=1"
    Dim frm As frmNoInvFilter
    Set frm = New frmNoInvFilter
    frm.Show vbModal
    If frm.OK Then
      If Session.IsMSSQL Then
          f = f & " and  instanceid not in (select instanceid from invos_inv where invdate >=" & MakeMSSQLDate(frm.dtpFrom.Value) & " and invdate <= " & MakeMSSQLDate(frm.dtpTo.Value) & ")"
      End If
      If Session.IsPOSTGRESQL Then
          f = f & " and  instanceid not in (select instanceid from invos_inv where invdate >=" & MakePGSQLDate(frm.dtpFrom.Value) & " and invdate <= " & MakePGSQLDate(frm.dtpTo.Value) & ")"
        
      End If
      
      If frm.chkShowClosed.Value = vbUnchecked Then
        f = f & " and   intsancestatusid<>'166d4978-0c4c-4575-8192-b251ac113781'"
      End If
      
    Else
      If Session.IsMSSQL Then
          f = f & " and  instanceid not in (select instanceid from invos_inv )"
      End If
      If Session.IsPOSTGRESQL Then
          f = f & " and  instanceid not in (select instanceid from invos_inv )"
      End If
    
    End If
    
    jfmnuINV_OS_BAD.jv.Filter.Add "AUTOINVOS_INFO", f
    
    
    jfmnuINV_OS_BAD.jv.Refresh
    Me.MousePointer = vbNormal
    jfmnuINV_OS_BAD.Show
    jfmnuINV_OS_BAD.WindowState = 0
    jfmnuINV_OS_BAD.ZOrder 0
End Sub

Private Sub mnuLoad2TDS_Click()
    Dim f As frmDB2TDS
    Set f = New frmDB2TDS
    f.Show vbModal
    Set f = Nothing
End Sub



Private Sub mnuINVF_Click()
    Dim journal As Object
    On Error Resume Next
    If jfmnuINVF Is Nothing Then
      Set jfmnuINVF = New frmJournalShow
      Set journal = Manager.GetInstanceObject("{89B136F2-5AF7-437F-9AC5-F4F5EB00FAF6}")
      Manager.LockInstanceObject journal.id
      Set jfmnuINVF.jv.journal = journal
      jfmnuINVF.jv.OpenModal = False
      jfmnuINVF.Caption = "Информация о загрузке файлов"
      Me.MousePointer = vbHourglass
      DoEvents
      Dim f As String
    f = "1=1"
    Dim fltr As frmINVF
    Set fltr = New frmINVF
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
      If fltr.lblINVF_DEF_LoadDate_LE.Value = vbChecked Then
        f = f & " and INVF_DEF_LoadDate<=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVF_DEF_LoadDate_LE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVF_DEF_LoadDate_LE.Value), MakePGSQLDate(fltr.dtpINVF_DEF_LoadDate_LE.Value)))
      End If
      If fltr.lblINVF_DEF_TheUser.Value = vbChecked Then
        f = f & " and INVF_DEF_TheUser_ID='" & fltr.txtINVF_DEF_TheUser.Tag & "'"
      End If
      If fltr.lblINVF_DEF_TypeOfFile.Value = vbChecked Then
        f = f & " and INVF_DEF_TypeOfFile like '%" & fltr.txtINVF_DEF_TypeOfFile.Text & "%'"
      End If
      If fltr.lblINVF_DEF_ThePath.Value = vbChecked Then
        f = f & " and INVF_DEF_ThePath like '%" & fltr.txtINVF_DEF_ThePath.Text & "%'"
      End If
      If fltr.lblINVF_DEF_TheHash.Value = vbChecked Then
        f = f & " and INVF_DEF_TheHash like '%" & fltr.txtINVF_DEF_TheHash.Text & "%'"
      End If
      If fltr.lblINVF_DEF_LoadDate_GE.Value = vbChecked Then
        f = f & " and INVF_DEF_LoadDate>=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVF_DEF_LoadDate_GE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVF_DEF_LoadDate_GE.Value), MakePGSQLDate(fltr.dtpINVF_DEF_LoadDate_GE.Value)))
      End If
    jfmnuINVF.jv.Filter.Add "AUTOINVF_DEF", f
    End If
      jfmnuINVF.jv.Refresh
      Me.MousePointer = vbNormal
    End If
    jfmnuINVF.Show
    jfmnuINVF.WindowState = 0
    jfmnuINVF.ZOrder 0
End Sub
Private Sub jfmnuINVF_OnFilter(usedefault As Boolean)
    Dim fltr As frmINVF
    Dim f As String
    Set fltr = New frmINVF
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
    f = "1=1"
      If fltr.lblINVF_DEF_LoadDate_LE.Value = vbChecked Then
        f = f & " and INVF_DEF_LoadDate<=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVF_DEF_LoadDate_LE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVF_DEF_LoadDate_LE.Value), MakePGSQLDate(fltr.dtpINVF_DEF_LoadDate_LE.Value)))
      End If
      If fltr.lblINVF_DEF_TheUser.Value = vbChecked Then
        f = f & " and INVF_DEF_TheUser_ID='" & fltr.txtINVF_DEF_TheUser.Tag & "'"
      End If
      If fltr.lblINVF_DEF_TypeOfFile.Value = vbChecked Then
        f = f & " and INVF_DEF_TypeOfFile like '%" & fltr.txtINVF_DEF_TypeOfFile.Text & "%'"
      End If
      If fltr.lblINVF_DEF_ThePath.Value = vbChecked Then
        f = f & " and INVF_DEF_ThePath like '%" & fltr.txtINVF_DEF_ThePath.Text & "%'"
      End If
      If fltr.lblINVF_DEF_TheHash.Value = vbChecked Then
        f = f & " and INVF_DEF_TheHash like '%" & fltr.txtINVF_DEF_TheHash.Text & "%'"
      End If
      If fltr.lblINVF_DEF_LoadDate_GE.Value = vbChecked Then
        f = f & " and INVF_DEF_LoadDate>=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVF_DEF_LoadDate_GE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVF_DEF_LoadDate_GE.Value), MakePGSQLDate(fltr.dtpINVF_DEF_LoadDate_GE.Value)))
      End If
    jfmnuINVF.jv.Filter.Add "AUTOINVF_DEF", f
    End If
    Unload fltr
    usedefault = False
End Sub
Private Sub jfmnuINVF_OnAdd(usedefaut As Boolean, Refesh As Boolean)
  Dim objGui  As Object
  Dim o As Object
  Dim id As String
  id = CreateGUID2
  Manager.NewInstance id, "INVF", "Информация о загрузке файлов" & Now, Site
  Set o = Manager.GetInstanceObject(id)
  If IsDocDenied(o) Then
    MsgBox "Не разрешен доступ к документам такого типа"
    Exit Sub
  End If

  Dim g  As Object
  Set g = Manager.GetInstanceGUI(o.id)
  If Not g Is Nothing Then
    g.Show GetDocumentMode(o), o, False
  End If
  usedefaut = False
  Refesh = False
End Sub
Private Sub mnuLoadMat_Click()
 Dim f As frmGetExcelMat
    Set f = New frmGetExcelMat
    f.Show vbModal
End Sub

Private Sub mnuLoadOS_Click()
    Dim f As frmGetExcel
    Set f = New frmGetExcel
    f.Show vbModal
End Sub


Private Sub mnuLoadPers_Click()
Dim f As frmLoadPers
Set f = New frmLoadPers
f.Show vbModal
Set f = Nothing
End Sub

Private Sub mnuLoadPortal_Click()
  Dim f As frmLoadPortal
  Set f = New frmLoadPortal
  f.Show vbModal
  Set f = Nothing
End Sub

Private Sub mnuLoadTech_Click()
  Dim rs As ADODB.Recordset
  Set rs = Session.GetData("select invos_info.instanceid from invos_info join invd_ostype on invd_ostypeid= OSTYPE and invd_ostype.Showtech <>0")
  
  If fProg Is Nothing Then
    Set fProg = New frmProgress
    fProg.Show
  End If
  Dim os As INV_OS.Application
  Dim isrk As INV_OS.INVOS_SROK
  While Not rs.EOF
    Set os = Manager.GetInstanceObject(rs!InstanceID)
    
    If os.INVOS_INFO.Item(1).TechFilePath <> "" Then
      fProg.NextVal os.brief & "-->" & os.INVOS_INFO.Item(1).TechFilePath
      LoadTech os
    End If
    
    Manager.FreeAllInstanses
    rs.MoveNext
  Wend
  
  rs.Close
  Set rs = Nothing
  fProg.Hide
  Unload fProg
  Set fProg = Nothing

End Sub

Private Sub mnuLog_Click()
  Dim f As frmLog
  Set f = New frmLog
  f.Show
End Sub

Private Sub mnuMaintainSetup_Click()
  Dim f As frmSetupMaintain
  Set f = New frmSetupMaintain
  f.Show vbModal
  Set f = Nothing
End Sub

Private Sub mnuOptMat_Click()
  Dim journal As Object
    On Error Resume Next
    If jfmnuINV_NUM Is Nothing Then
      Set jfmnuINV_NUM = New frmJournalShow
      Set journal = Manager.GetInstanceObject("{D6D2D29A-BDE5-4690-9E36-F4988B15B481}")
      Manager.LockInstanceObject journal.id
      Set jfmnuINV_NUM.jv.journal = journal
      jfmnuINV_NUM.jv.OpenModal = False
      jfmnuINV_NUM.Caption = "Нумерация"
      Me.MousePointer = vbHourglass
      DoEvents
      Dim f As String
    f = "1=1"
    Dim fltr As frmINV_NUM
    Set fltr = New frmINV_NUM
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
      If fltr.lblINVN_DEF_TheNumber_GE.Value = vbChecked Then
        f = f & " and INVN_DEF_TheNumber>=" & val(fltr.txtINVN_DEF_TheNumber_GE.Text)
      End If
      If fltr.lblINVN_DEF_TheNumber_LE.Value = vbChecked Then
        f = f & " and INVN_DEF_TheNumber<=" & val(fltr.txtINVN_DEF_TheNumber_LE.Text)
      End If
    jfmnuINV_NUM.jv.Filter.Add "AUTOINVN_DEF", f
    End If
      jfmnuINV_NUM.jv.Refresh
      Me.MousePointer = vbNormal
    End If
    jfmnuINV_NUM.Show
    jfmnuINV_NUM.WindowState = 0
    jfmnuINV_NUM.ZOrder 0
End Sub

Private Sub mnuPortalCFG_Click()
Dim f As frmPortalSetup
Set f = New frmPortalSetup
f.Show vbModal
Set f = Nothing
End Sub

Private Sub mnuPrinters_Click()
Dim f As frmPrnSetup
Set f = New frmPrnSetup
f.Show vbModal
Set f = Nothing
End Sub

Private Sub mnuPrintSHCode_Click()
    Dim f As frmGetInv2
    Set f = New frmGetInv2

    f.Show vbModal
    If f.OK Then
        On Error Resume Next
        InstallFont App.path & "\code128.ttf", "Code 128", "code128.ttf"
        Set RptSHCODE = New ReportShow
        RptSHCODE.Caption = "Штрихкоды"
        
        RptSHCODE.PrinterName = GetSetting("SGS", "ITTSETTINGS", "ZPRN", "")
        RptSHCODE.ReportPath = App.path & "\shcode_BIG.rpt"
        RptSHCODE.ReportSource = "v_RPT_shcode"
        
        '
        
        Dim s As String
        s = "INV_INSTANCEID='" + f.txtINV.Tag + "'"
        If f.chkBad.Value = vbChecked Then
          s = s & " and invos_code_visiblecode in (select shcode from invi_chng where instanceid ='" + f.txtINV.Tag + "')"
        End If
        If f.txtCodeMask <> "" Then
          s = s & " and invos_code_visiblecode like '" & f.txtCodeMask & "%' "
        End If
        If f.txtMask <> "" Then
          s = s & " and invos_info_name like '" & f.txtMask & "%' "
        End If
        If f.txtMaskE(0) <> "" Then
          s = s & " and invos_info_name not like '" & f.txtMaskE(0) & "%' "
        End If
       
        If f.txtMaskE(1) <> "" Then
          s = s & " and invos_info_name not like '" & f.txtMaskE(1) & "%' "
        End If
        If f.txtMaskE(2) <> "" Then
          s = s & " and invos_info_name not like '" & f.txtMaskE(2) & "%' "
        End If
        If f.txtMaskE(3) <> "" Then
          s = s & " and invos_info_name not like '" & f.txtMaskE(3) & "%' "
        End If
        If f.txtMaskE(4) <> "" Then
          s = s & " and invos_info_name not like '" & f.txtMaskE(4) & "%' "
        End If
        If f.txtMaskE(5) <> "" Then
          s = s & " and invos_info_name not like '" & f.txtMaskE(5) & "%' "
        End If
        RptSHCODE.ReportFilter = s
        RptSHCODE.Run
    End If
End Sub

Private Sub mnuRptBadInv_Click()
 
    Dim f As frmGetInv
    Set f = New frmGetInv

    f.Show vbModal
    If f.OK Then
        Set RptInvBAD = New ReportShow
        RptInvBAD.Caption = "Справка об ОС непрошедших инвентаризацию"
        
        RptInvBAD.PrinterName = GetSetting("SGS", "ITTSETTINGS", "DOCPRN", "")
        RptInvBAD.ReportPath = App.path & "\InvBAD.rpt"
        RptInvBAD.ReportSource = "v_RPT_INVENTORY_BAD"
        
        Dim s As String
        s = " intsancestatusid<>'166d4978-0c4c-4575-8192-b251ac113781'  and INV_INSTANCEID='" + f.txtINV.Tag + "'"
        RptInvBAD.ReportFilter = s
        RptInvBAD.Run
    End If
End Sub

Private Sub mnuRptInv_Click()

    Dim f As frmGetInv
    Set f = New frmGetInv

    f.Show vbModal
    If f.OK Then
        Set RptInvOK = New ReportShow
        RptInvOK.Caption = "Справка об инвентаризации"
        
        RptInvOK.PrinterName = GetSetting("SGS", "ITTSETTINGS", "DOCPRN", "")
        RptInvOK.ReportPath = App.path & "\invok.rpt"
        RptInvOK.ReportSource = "v_RPT_INVENTORY_OK"
        
        Dim s As String
        s = "INV_INSTANCEID='" + f.txtINV.Tag + "'"
        RptInvOK.ReportFilter = s
        RptInvOK.Run
    End If


End Sub

Private Sub mnuRptOS_Click()
  Dim f As frmGetOS
  Set f = New frmGetOS
  f.Show vbModal
  If f.OK Then
  
  
   Set ObjectToReport = Manager.GetInstanceObject(f.txtINV.Tag)
         Dim fn As String
         fn = f.txtPath.Text
         Set osRPt = New MTZReportHelper.WordHelper
         osRPt.MakeDocument fn
         Set osRPt = Nothing
bye:
         Set osRPt = Nothing
         Set ObjectToReport = Nothing
  End If
End Sub


Private Sub mnuRptShCode52_Click()
 Dim f As frmGetInv2
    Set f = New frmGetInv2

    f.Show vbModal
    If f.OK Then
        On Error Resume Next
        InstallFont App.path & "\code128.ttf", "Code 128", "code128.ttf"
        Set RptSHCODE = New ReportShow
        RptSHCODE.Caption = "Штрихкоды"
        
        RptSHCODE.PrinterName = GetSetting("SGS", "ITTSETTINGS", "ZPRN", "")
        RptSHCODE.ReportPath = App.path & "\shcode_52.rpt"
        RptSHCODE.ReportSource = "v_RPT_shcode"
        
        Dim s As String
        s = "INV_INSTANCEID='" + f.txtINV.Tag + "'"
         If f.chkBad.Value = vbChecked Then
          s = s & " and invos_code_visiblecode in (select shcode from invi_chng where instanceid ='" + f.txtINV.Tag + "')"
        End If
        If f.txtCodeMask <> "" Then
          s = s & " and invos_code_visiblecode like '" & f.txtCodeMask & "%' "
        End If
        If f.txtMask <> "" Then
          s = s & " and invos_info_name like '" & f.txtMask & "%' "
        End If
        If f.txtMaskE(0) <> "" Then
          s = s & " and invos_info_name not like '" & f.txtMaskE(0) & "%' "
        End If
       
        If f.txtMaskE(1) <> "" Then
          s = s & " and invos_info_name not like '" & f.txtMaskE(1) & "%' "
        End If
        If f.txtMaskE(2) <> "" Then
          s = s & " and invos_info_name not like '" & f.txtMaskE(2) & "%' "
        End If
        If f.txtMaskE(3) <> "" Then
          s = s & " and invos_info_name not like '" & f.txtMaskE(3) & "%' "
        End If
        If f.txtMaskE(4) <> "" Then
          s = s & " and invos_info_name not like '" & f.txtMaskE(4) & "%' "
        End If
        If f.txtMaskE(5) <> "" Then
          s = s & " and invos_info_name not like '" & f.txtMaskE(5) & "%' "
        End If
        RptSHCODE.ReportFilter = s
        RptSHCODE.Run
    End If
End Sub

Private Sub mnuRptUnknown_Click()
 Dim f As frmGetInv
    Set f = New frmGetInv

    f.Show vbModal
    If f.OK Then
        Set RptInvUnk = New ReportShow
        RptInvUnk.Caption = "Неучтенные объекты"
        
        RptInvUnk.PrinterName = GetSetting("SGS", "ITTSETTINGS", "DOCPRN", "")
        RptInvUnk.ReportPath = App.path & "\InvUNK.rpt"
        RptInvUnk.ReportSource = "INVI_UNK"
        
        Dim s As String
        s = "INSTANCEID='" + f.txtINV.Tag + "'"
        RptInvUnk.ReportFilter = s
        RptInvUnk.Run
    End If
End Sub

Private Sub mnuVacuum_Click()
  DBMainteice
  MsgBox ("Обслуживание базы завершено")
End Sub

Private Sub osRPt_MakeContent()
Dim os As INV_OS.Application
Dim s As String, sp As Long, ep As Long
Dim sc As SortableCollection
Dim db As DBuffer
  Dim i As Long

Set os = ObjectToReport

osRPt.h = -1
osRPt.NextHeader

osRPt.OutStr "Сводка по: " & os.Name
osRPt.Header

If os.INVOS_INFO.Count > 0 Then
  With os.INVOS_INFO.Item(1)
    On Error Resume Next
    
    osRPt.OutStr "Название: " & .ShortName
    osRPt.Bold
    osRPt.OutStr "Полное наименование: " & .Name
    osRPt.OutStr "Статус: " & os.StatusName
    
    If Not .theorg Is Nothing Then
      osRPt.OutStr "На учете в: " & .theorg.brief
    End If
    
    If Not .ostype Is Nothing Then
      osRPt.OutStr "Группа: " & .ostype.brief
    End If
    
    If .IsMaterial = Boolean_Da Then
      osRPt.OutStr "Материал"
      osRPt.OutStr "Карточка учета: " & NoTabs(.CardNum)
      osRPt.OutStr "№ в партии: " & NoTabs(.InLineNum)
    Else
      osRPt.OutStr "Основное средство"
    End If
    
    osRPt.OutStr "Инвентарный номер: " & NoTabs(.invNum)
   
    osRPt.OutStr "Цена: " & NoTabs(.TheCost)
    
    osRPt.OutStr "Примечание: " & NoTabs(.Info)
    
  End With
End If

If os.INVOS_PLACE.Count > 0 Then
  With os.INVOS_PLACE.Item(1)
    On Error Resume Next
    
   
    If Not .TheHouse Is Nothing Then
      osRPt.OutStr "Здание:" & .TheHouse.brief
    End If
     If Not .Direction Is Nothing Then
      osRPt.OutStr "Дирекция: " & .Direction.brief
    End If
    
     If Not .Uprav Is Nothing Then
      osRPt.OutStr "Управление: " & .Uprav.brief
    End If
    
     If Not .Otdel Is Nothing Then
      osRPt.OutStr "Отдел: " & .Otdel.brief
    End If
    
    If Not .TheOwner Is Nothing Then
      osRPt.OutStr "Ответственный: " & .TheOwner.brief
    End If
    
    osRPt.OutStr "Номер комплекта: " & NoTabs(.ComplNumber)
    
    If Not .MatOtv Is Nothing Then
      osRPt.OutStr "МОЛ:" & .MatOtv.brief
    End If
    osRPt.OutStr "Комментарий: " & NoTabs(.Info)
    
  End With
End If
    
    
 'Драг. металлы
If os.INVOS_DRAG.Count > 0 Then
  osRPt.NextHeader
  osRPt.OutStr "Содержание драг. металлов"
  osRPt.Header

  sp = osRPt.wdoc.Paragraphs.Count
  s = "Драг. м."
  s = s & vbTab & "Содержание"
  osRPt.OutStr s
  osRPt.Bold
  
 

  For i = 1 To os.INVOS_DRAG.Count
      With os.INVOS_DRAG.Item(i)
        If Not .DragMet Is Nothing Then
         s = NoTabs(.DragMet.brief)
        Else
         s = "-"
        End If
        s = s & vbTab & NoTabs(.q)
      
       
      End With
      osRPt.OutStr s
    
  Next
  ep = osRPt.wdoc.Paragraphs.Count
  osRPt.MakeTable sp, ep, ep - sp + 1, 2
  osRPt.PrevHeader
End If
    
    
    
' заметки
If os.INVOS_CMNT.Count > 0 Then
  osRPt.NextHeader
  osRPt.OutStr "Заметки о состоянии"
  osRPt.Header

  sp = osRPt.wdoc.Paragraphs.Count
  s = "Дата"
  s = s & vbTab & "Зарегистрировал"
  s = s & vbTab & "Информация"
  s = s & vbTab & "Примечание"
  
  osRPt.OutStr s
  osRPt.Bold

 
  os.INVOS_CMNT.Sort = "TheDate"
  For i = 1 To os.INVOS_CMNT.Count
      With os.INVOS_CMNT.Item(i)
        s = NoTabs(.TheDate)
        s = s & vbTab & NoTabs(.TheCommenter.brief)
        s = s & vbTab & NoTabs(.Info)
        s = s & vbTab & NoTabs(.TheComment)
      End With
      osRPt.OutStr s
    
  Next
  ep = osRPt.wdoc.Paragraphs.Count
  osRPt.MakeTable sp, ep, ep - sp + 1, 4
  osRPt.PrevHeader
End If

' ремонты
If os.INVOS_REPAIR.Count > 0 Then
  osRPt.NextHeader
  osRPt.OutStr "Ремонты"
  osRPt.Header

  sp = osRPt.wdoc.Paragraphs.Count
  s = "Дата начала"
  s = s & vbTab & "Дата завершения"
  s = s & vbTab & "Вид ремонта"
  s = s & vbTab & "№ приказа"
  
  osRPt.OutStr s
  osRPt.Bold
  
 
  os.INVOS_REPAIR.Sort = "StartDate"
  For i = 1 To os.INVOS_REPAIR.Count
      With os.INVOS_REPAIR.Item(i)
        s = NoTabs(.StartDate)
        s = s & vbTab & NoTabs(.EndDate)
        s = s & vbTab & NoTabs(.Info)
        s = s & vbTab & NoTabs(.DocNumber)
      End With
      osRPt.OutStr s
    
  Next
  ep = osRPt.wdoc.Paragraphs.Count
  osRPt.MakeTable sp, ep, ep - sp + 1, 4
  osRPt.PrevHeader
End If


'модернизации
If os.INVOS_MOD.Count > 0 Then
  osRPt.NextHeader
  osRPt.OutStr "Модернизации"
  osRPt.Header

  sp = osRPt.wdoc.Paragraphs.Count
  s = "Дата начала"
  s = s & vbTab & "Дата завершения"
  s = s & vbTab & "Вид модернизации"
  s = s & vbTab & "№ приказа"
  
  osRPt.OutStr s
  osRPt.Bold
 
 
  os.INVOS_MOD.Sort = "StartDate"
  For i = 1 To os.INVOS_MOD.Count
      With os.INVOS_MOD.Item(i)
        s = NoTabs(.StartDate)
        s = s & vbTab & NoTabs(.EndDate)
        s = s & vbTab & NoTabs(.Info)
        s = s & vbTab & NoTabs(.DocNumber)
      End With
      osRPt.OutStr s
    
  Next
  ep = osRPt.wdoc.Paragraphs.Count
  osRPt.MakeTable sp, ep, ep - sp + 1, 4
  osRPt.PrevHeader
End If

'консервации
If os.INVOS_CNSRV.Count > 0 Then
  osRPt.NextHeader
  osRPt.OutStr "Консервации"
  osRPt.Header

  sp = osRPt.wdoc.Paragraphs.Count
  s = "Дата начала"
  s = s & vbTab & "Дата завершения"
  s = s & vbTab & "№ приказа"
 
  osRPt.OutStr s
  osRPt.Bold

 
  os.INVOS_CNSRV.Sort = "StartDate"
  For i = 1 To os.INVOS_CNSRV.Count
      With os.INVOS_CNSRV.Item(i)
        s = NoTabs(.StartDate)
         s = s & vbTab & NoTabs(.EndDate)
        s = s & vbTab & NoTabs(.DocNumber)
      End With
      osRPt.OutStr s
    
  Next
  ep = osRPt.wdoc.Paragraphs.Count
  osRPt.MakeTable sp, ep, ep - sp + 1, 3
  osRPt.PrevHeader
End If


'Аренда
If os.INVOS_RENT.Count > 0 Then
  osRPt.NextHeader
  osRPt.OutStr "Передача в аренду"
  osRPt.Header

  sp = osRPt.wdoc.Paragraphs.Count
  s = "Дата начала"
  s = s & vbTab & "Дата завершения"
  s = s & vbTab & "№ приказа"
  s = s & vbTab & "Договор ареды"
  s = s & vbTab & "Арендатор"
  osRPt.OutStr s
  osRPt.Bold

 
  os.INVOS_RENT.Sort = "StartDate"
  For i = 1 To os.INVOS_RENT.Count
      With os.INVOS_RENT.Item(i)
        s = NoTabs(.StartDate)
        s = s & vbTab & NoTabs(.EndDate)
        s = s & vbTab & NoTabs(.DocNumber)
        s = s & vbTab & NoTabs(.ADog)
        If Not .arendator Is Nothing Then
         s = s & vbTab & NoTabs(.arendator.brief)
        Else
         s = s & vbTab & "-"
        End If
      End With
      osRPt.OutStr s
    
  Next
  ep = osRPt.wdoc.Paragraphs.Count
  osRPt.MakeTable sp, ep, ep - sp + 1, 5
  osRPt.PrevHeader
End If

'лизинг
If os.INVOS_LIZING.Count > 0 Then
  osRPt.NextHeader
  osRPt.OutStr "Передача в лизинг"
  osRPt.Header

  sp = osRPt.wdoc.Paragraphs.Count
  s = "Дата передачи"
  s = s & vbTab & "№ приказа"
  s = s & vbTab & "Контрагент"
  osRPt.OutStr s
  osRPt.Bold

 
  os.INVOS_LIZING.Sort = "TheDate"
  For i = 1 To os.INVOS_LIZING.Count
      With os.INVOS_LIZING.Item(i)
        s = NoTabs(.TheDate)
        s = s & vbTab & NoTabs(.DocNumber)
    
        If Not .TheAgent Is Nothing Then
         s = s & vbTab & NoTabs(.TheAgent.brief)
        Else
         s = s & vbTab & "-"
        End If
      End With
      osRPt.OutStr s
    
  Next
  ep = osRPt.wdoc.Paragraphs.Count
  osRPt.MakeTable sp, ep, ep - sp + 1, 3
  osRPt.PrevHeader
End If


'Списание
If os.INVOS_OFFRULE.Count > 0 And os.StatusName = "Списано" Then
  osRPt.NextHeader
  osRPt.OutStr "Списание"
  osRPt.Header

  sp = osRPt.wdoc.Paragraphs.Count
  s = "Дата передачи"
  s = s & vbTab & "№ приказа"
  s = s & vbTab & "Причина списания"
  s = s & vbTab & "Примечание"
  osRPt.OutStr s
  osRPt.Bold

 
  os.INVOS_OFFRULE.Sort = "DocDate"
  For i = 1 To os.INVOS_OFFRULE.Count
      With os.INVOS_OFFRULE.Item(i)
        s = NoTabs(.DocDate)
        s = s & vbTab & NoTabs(.DocNumber)
        s = s & vbTab & NoTabs(.Info)
        s = s & vbTab & NoTabs(.TheComment)
      End With
      osRPt.OutStr s
    
  Next
  ep = osRPt.wdoc.Paragraphs.Count
  osRPt.MakeTable sp, ep, ep - sp + 1, 4
  osRPt.PrevHeader
End If


End Sub


Private Sub mnuRptVed_Click()
    Dim f As frmMultiInv
    Set f = New frmMultiInv

    f.Show vbModal
    If f.OK Then
        
        Dim s As String
        Dim q As String
        Dim ReportFilter As String
        s = ""
        q = ""
        Dim i As Long
        For i = 0 To f.lstInv.ListCount - 1
          If f.lstInv.Selected(i) Then
            If q <> "" Then
              q = q & ","
            End If
            q = q & "'" & f.col.Item(f.lstInv.ItemData(i)).id & "'"
          End If
        Next
        Dim dop As String, dopRoom As String
        
        
      ' build flter expression
     
     
      
      If f.lblinvi_DEF_Otdel.Value = vbChecked Then
        s = s & " and invos_place_Otdel='" & f.txtinvi_DEF_Otdel.Text & "'"
      End If
    
     
    
    
      If f.lblinvi_DEF_DIrection.Value = vbChecked Then
        s = s & " and invos_place_DIrection='" & f.txtinvi_DEF_DIrection.Text & "'"
      End If
      If f.lblinvi_DEF_Uprev.Value = vbChecked Then
        s = s & " and invos_place_Uprav='" & f.txtinvi_DEF_Uprev.Text & "'"
      End If
     
      If f.lblinvi_DEF_TheOrg.Value = vbChecked Then
        s = s & " and invos_info_TheOrg='" & f.txtinvi_DEF_TheOrg.Text & "'"
      End If
        
       
      If f.chkExcludeBroken.Value = vbChecked Then
      s = s & " and a.intsancestatusid<>'166d4978-0c4c-4575-8192-b251ac113781' "
      End If
      
      If f.lblinvi_DEF_TheFlow.Value = vbChecked Then
        dop = f.txtinvi_DEF_TheFlow.Text
      ElseIf f.lblinvi_DEF_TheRoom.Value = vbChecked Or f.lblinvi_DEF_TheWorkPlace.Value = vbChecked Then
        dop = "%"
      End If
      


     If f.lblinvi_DEF_TheRoom.Value = vbChecked Then
        If dop <> "" Then
          dop = dop & "."
        End If
        dop = dop & f.txtinvi_DEF_TheRoom.Text
      ElseIf f.lblinvi_DEF_TheWorkPlace.Value = vbChecked Then
        If dop <> "" Then
          dop = dop & ".%"
        End If
      End If


        If f.lblinvi_DEF_TheWorkPlace.Value = vbChecked Then
          If dop <> "" Then
            dop = dop & "."
          End If
          dop = dop & f.txtinvi_DEF_TheWorkPlace.Text
        Else
          dopRoom = dop
          If dop <> "" Then
            dop = dop & ".%"
          End If
        End If
        
   
        Dim zz As String
         If dop <> "" Then
          If dopRoom <> "" Then
           zz = "b.INSTANCEID in (" & q & ") " & s & " and ( invos_place_complnumber like '" & dop & "' or invos_place_complnumber='" & dopRoom & "')"
          Else
           zz = "b.INSTANCEID in (" & q & ") " & s & " and ( invos_place_complnumber like '" & dop & "')"
          End If
        Else
          zz = "b.INSTANCEID in (" & q & ") " & s
        End If
        
        Dim rs As ADODB.Recordset
        Set rs = Session.GetData( _
        vbCrLf & "select  a.invos_info_theorg, a.invos_info_ismaterial, a.invos_info_ostype, a.invos_info_name,a.invos_info_cardnum, a.invos_info_thecost, a.invos_place_matotv, a.invos_place_thehouse," & _
        vbCrLf & " a.invos_place_complnumber, a.invos_place_direction, a.invos_place_uprav, a.invos_place_otdel," & _
        vbCrLf & " a.invos_place_theowner, a.statusname,sum(was) AS was, sum(found) AS found " & _
        vbCrLf & "from(" & _
        vbCrLf & "  SELECT DISTINCT A.ID, a.invos_info_theorg, a.invos_info_ismaterial, a.invos_info_ostype, a.invos_info_name, " & _
        vbCrLf & "           a.invos_info_cardnum, a.invos_info_thecost, a.invos_place_matotv, a.invos_place_thehouse," & _
        vbCrLf & "           a.invos_place_complnumber, a.invos_place_direction, a.invos_place_uprav, a.invos_place_otdel," & _
        vbCrLf & "           a.invos_place_theowner, a.statusname,1 AS was, 0 AS found" & _
        vbCrLf & "  FROM v_autoinvos_info a" & _
        vbCrLf & "  JOIN invi_obj b ON b.theos = a.id where " & zz & _
        vbCrLf & "  union all " & _
        vbCrLf & "  SELECT DISTINCT A.ID, a.invos_info_theorg, a.invos_info_ismaterial, a.invos_info_ostype, a.invos_info_name," & _
        vbCrLf & "           a.invos_info_cardnum, a.invos_info_thecost, a.invos_place_matotv, a.invos_place_thehouse," & _
        vbCrLf & "           a.invos_place_complnumber, a.invos_place_direction, a.invos_place_uprav, a.invos_place_otdel," & _
        vbCrLf & "           a.invos_place_theowner, a.statusname,0 AS was, 1 AS found" & _
        vbCrLf & "  FROM v_autoinvos_info a" & _
        vbCrLf & "  JOIN invi_done b ON b.theos = a.id where " & zz & _
        vbCrLf & ") A group by " & _
        vbCrLf & " a.invos_info_theorg, a.invos_info_ismaterial, a.invos_info_ostype, a.invos_info_name,a.invos_info_cardnum, a.invos_info_thecost, a.invos_place_matotv, a.invos_place_thehouse," & _
        vbCrLf & " a.invos_place_complnumber, a.invos_place_direction, a.invos_place_uprav, a.invos_place_otdel," & _
        vbCrLf & " a.invos_place_theowner , a.StatusName ")

    
    
    
      If f.ExportData Then
        
          Dim ex As Object
    Dim excel As Object
    
    Set excel = CreateObject("Excel.Application")
    With excel.Workbooks.Add
    Set ex = .Worksheets.Item(1)
    End With
    
    Dim xs() As Variant
    ReDim xs(0 To 15)
    Dim j As Long
    xs(0) = "Организация"
    xs(1) = "Материал"
    xs(2) = "Тип"
    xs(3) = "Название"
    xs(4) = "№ карточки учета"
    xs(5) = "Цена"
    xs(6) = "Ответственное лицо"
    xs(7) = "Здание"
    xs(8) = "№ комплекта"
    xs(9) = "Дирекция"
    xs(10) = "Управление"
    xs(11) = "Отдел"
    xs(12) = "Владелец"
    xs(13) = "Состояние"
    xs(14) = "Числится"
    xs(15) = "По факту"
    
    Dim cap As String
    Dim q1 As Long
    Dim q2 As Long
    cap = Me.Caption
    
    ex.Range(ex.Cells(1, 1), ex.Cells(1, 16)).Value = xs
    i = 0
    While Not rs.EOF And i < 65000
     
      For j = 0 To rs.fields.Count - 1
        
        If j <> 5 And j <= 13 Then
          xs(j) = "'" & rs.fields(j)
        Else
          xs(j) = rs.fields(j)
          If j = 14 Then q1 = q1 + rs.fields(j)
          If j = 15 Then q2 = q2 + rs.fields(j)
        End If
      Next
      ex.Range(ex.Cells(i + 2, 1), ex.Cells(i + 2, 16)).Value = xs
      Me.Caption = i & ":" & rs.fields(3)
      DoEvents
      rs.MoveNext
      i = i + 1
     
    Wend
    
    
    ex.Range(ex.Cells(1, 1), ex.Cells(i + 1, 16)).Select
    
     For j = 0 To rs.fields.Count - 1
      
          xs(j) = ""
          If j = 14 Then xs(j) = q1
          If j = 15 Then xs(j) = q2
      
      Next
    xs(0) = "Итого"
    ex.Range(ex.Cells(i + 2, 1), ex.Cells(i + 2, 16)).Value = xs
   
    excel.Selection.Columns.AutoFit
    If f.optCompl.Value = True Then
    excel.Selection.Sort Key1:=ex.Range("A2"), Order1:=1, Key2:=ex.Range("J2") _
        , Order2:=1, Key3:=ex.Range("I2"), Order3:=1, Header:=0
    Else
        excel.Selection.Sort Key1:=ex.Range("A2"), Order1:=1, Key2:=ex.Range("B2") _
        , Order2:=1, Key3:=ex.Range("C2"), Order3:=1, Header:=0

    End If
'    ex.Select
    excel.Selection.AutoFormat Format:=12, Font _
            :=True, Alignment:=True, Border:=True, Pattern:=True, Width:=True
    excel.ActiveWindow.Visible = True
    excel.Visible = True
      Me.Caption = cap
      Else
        Set RptSLICH = New ReportShow
        RptSLICH.Caption = "Справка об инвентаризации"
        
        RptSLICH.PrinterName = GetSetting("SGS", "ITTSETTINGS", "DOCPRN", "")
        If f.optCompl.Value = True Then
          RptSLICH.ReportPath = App.path & "\SLICH2.rpt"
        Else
          RptSLICH.ReportPath = App.path & "\SLICH.rpt"
        End If
        RptSLICH.ReportSource = "v_RPT_SLICH"
        
        RptSLICH.EnableTree = True
        RptSLICH.RunDirectRS rs
      End If
        
    End If
End Sub

Private Sub mnuRtSH2_Click()
    Dim f As frmGetInv2
    Set f = New frmGetInv2

    f.Show vbModal
    If f.OK Then
        On Error Resume Next
        InstallFont App.path & "\code128.ttf", "Code 128", "code128.ttf"
        Set RptSHCODE = New ReportShow
        RptSHCODE.Caption = "Штрихкоды"
        
        RptSHCODE.PrinterName = GetSetting("SGS", "ITTSETTINGS", "ZPRN", "")
        RptSHCODE.ReportPath = App.path & "\shcode.rpt"
        RptSHCODE.ReportSource = "v_RPT_shcode"
        
        Dim s As String
        s = "INV_INSTANCEID='" + f.txtINV.Tag + "'"
         If f.chkBad.Value = vbChecked Then
          s = s & " and invos_code_visiblecode in (select shcode from invi_chng where instanceid ='" + f.txtINV.Tag + "')"
        End If
        If f.txtCodeMask <> "" Then
          s = s & " and invos_code_visiblecode like '" & f.txtCodeMask & "%' "
        End If
        If f.txtMask <> "" Then
          s = s & " and invos_info_name like '" & f.txtMask & "%' "
        End If
        If f.txtMaskE(0) <> "" Then
          s = s & " and invos_info_name not like '" & f.txtMaskE(0) & "%' "
        End If
       
        If f.txtMaskE(1) <> "" Then
          s = s & " and invos_info_name not like '" & f.txtMaskE(1) & "%' "
        End If
        If f.txtMaskE(2) <> "" Then
          s = s & " and invos_info_name not like '" & f.txtMaskE(2) & "%' "
        End If
        If f.txtMaskE(3) <> "" Then
          s = s & " and invos_info_name not like '" & f.txtMaskE(3) & "%' "
        End If
        If f.txtMaskE(4) <> "" Then
          s = s & " and invos_info_name not like '" & f.txtMaskE(4) & "%' "
        End If
        If f.txtMaskE(5) <> "" Then
          s = s & " and invos_info_name not like '" & f.txtMaskE(5) & "%' "
        End If
        RptSHCODE.ReportFilter = s
        RptSHCODE.Run
    End If
End Sub

Private Sub mnuTileHor_Click()
  Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuTileVert_Click()
  Me.Arrange vbTileVertical
End Sub

Private Sub Timer2_Timer()
  If inTimer2 Then Exit Sub
  inTimer2 = True
  On Error Resume Next
  Session.SessionTouch
  inTimer2 = False
End Sub





Private Function NoTabs(ByVal s As String) As String
  NoTabs = Replace(Replace(Replace(Replace(s, vbTab, " "), vbCrLf, " "), vbCr, " "), vbLf, " ")
End Function


Private Sub OpenForm(o As Object)
  Dim t As Form
  For Each t In Forms
    If t.Caption = o.Name Then
      t.WindowState = vbNormal
      t.ZOrder 0
      t.Show
      Me.MousePointer = vbNormal
      Exit Sub
    End If
  Next
  
  Dim f As frmObj
  Set f = New frmObj
  f.Init o
  f.Show
  

End Sub




Private Function SynchronizeARMDescription()
    Dim objARM As Object
    Dim objMenuItem As Menu
    Dim ObjItem As Object

    Set objARM = Manager.GetInstanceObject(ARMID)
    
    Dim i As Long
    Dim objRS As ADODB.Recordset
    Dim objEntryPoint As Object
    
    For i = 0 To Me.Controls.Count - 1
        Set ObjItem = Me.Controls(i)
        If UCase(TypeName(ObjItem)) = UCase("menu") Then
            If ObjItem.Caption <> "-" Then
              Debug.Print "Found menu " + ObjItem.Caption + "-" + ObjItem.Name
              Set objRS = Session.GetRowsEx("EntryPoints", ARMID, , "Caption='" + ObjItem.Caption + "' or Name='" & ObjItem.Name & "'")
              If objRS.EOF And objRS.BOF Then
                  Set objEntryPoint = objARM.EntryPoints.Add
                  objEntryPoint.Caption = ObjItem.Caption
                  objEntryPoint.Name = ObjItem.Name
                  objEntryPoint.AsToolbarItem = Boolean_Net
                  objEntryPoint.ActionType = 0 'MenuActionType_Nicego_ne_delat_
                  objEntryPoint.save
              Else
                  Set objEntryPoint = objARM.FindRowObject("EntryPoints", objRS!Entrypointsid)
                  If Not objEntryPoint Is Nothing Then
                    objEntryPoint.Caption = ObjItem.Caption
                    objEntryPoint.Name = ObjItem.Name
                    objEntryPoint.AsToolbarItem = Boolean_Net
                    objEntryPoint.save
                  End If
              End If
              objRS.Close
            End If
        End If
    Next
End Function


Private Sub mnuAllINV_INV_Click()
    Dim journal As Object
    On Error Resume Next
    If jfmnuAllINV_INV Is Nothing Then
      Set jfmnuAllINV_INV = New frmJournalShow
      Set journal = Manager.GetInstanceObject("{D743AE87-E934-4691-B1A3-00BAC0E83F0C}")
      Manager.LockInstanceObject journal.id
      Set jfmnuAllINV_INV.jv.journal = journal
      jfmnuAllINV_INV.jv.OpenModal = False
      jfmnuAllINV_INV.Caption = "Инвентаризация - все состояния"
      Me.MousePointer = vbHourglass
      DoEvents
      Dim f As String
    f = "1=1"
    Dim fltr As frmINV_INV
    Set fltr = New frmINV_INV
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
      If fltr.lblinvi_DEF_Building.Value = vbChecked Then
        f = f & " and invi_DEF_Building_ID='" & fltr.txtinvi_DEF_Building.Tag & "'"
      End If
      If fltr.lblinvi_DEF_TheFlow.Value = vbChecked Then
        f = f & " and invi_DEF_TheFlow like '%" & fltr.txtinvi_DEF_TheFlow.Text & "%'"
      End If
      If fltr.lblinvi_DEF_Otdel.Value = vbChecked Then
        f = f & " and invi_DEF_Otdel_ID='" & fltr.txtinvi_DEF_Otdel.Tag & "'"
      End If
      If fltr.lblinvi_DEF_TheOwner.Value = vbChecked Then
        f = f & " and invi_DEF_TheOwner_ID='" & fltr.txtinvi_DEF_TheOwner.Tag & "'"
      End If
      If fltr.lblinvi_DEF_MatOtv.Value = vbChecked Then
        f = f & " and invi_DEF_MatOtv_ID='" & fltr.txtinvi_DEF_MatOtv.Tag & "'"
      End If
      If fltr.lblinvi_DEF_Info.Value = vbChecked Then
        f = f & " and invi_DEF_Info like '%" & fltr.txtinvi_DEF_Info.Text & "%'"
      End If
      If fltr.lblinvi_DEF_TheRoom.Value = vbChecked Then
        f = f & " and invi_DEF_TheRoom like '%" & fltr.txtinvi_DEF_TheRoom.Text & "%'"
      End If
      If fltr.lblinvi_DEF_TheWorkPlace.Value = vbChecked Then
        f = f & " and invi_DEF_TheWorkPlace like '%" & fltr.txtinvi_DEF_TheWorkPlace.Text & "%'"
      End If
      If fltr.lblinvi_DEF_StartDate_LE.Value = vbChecked Then
        f = f & " and invi_DEF_StartDate<=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpinvi_DEF_StartDate_LE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpinvi_DEF_StartDate_LE.Value), MakePGSQLDate(fltr.dtpinvi_DEF_StartDate_LE.Value)))
      End If
      If fltr.lblinvi_DEF_EndDate_GE.Value = vbChecked Then
        f = f & " and invi_DEF_EndDate>=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpinvi_DEF_EndDate_GE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpinvi_DEF_EndDate_GE.Value), MakePGSQLDate(fltr.dtpinvi_DEF_EndDate_GE.Value)))
      End If
      If fltr.lblinvi_DEF_OrderNum.Value = vbChecked Then
        f = f & " and invi_DEF_OrderNum like '%" & fltr.txtinvi_DEF_OrderNum.Text & "%'"
      End If
      If fltr.lblinvi_DEF_StartDate_GE.Value = vbChecked Then
        f = f & " and invi_DEF_StartDate>=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpinvi_DEF_StartDate_GE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpinvi_DEF_StartDate_GE.Value), MakePGSQLDate(fltr.dtpinvi_DEF_StartDate_GE.Value)))
      End If
      If fltr.lblinvi_DEF_DIrection.Value = vbChecked Then
        f = f & " and invi_DEF_DIrection_ID='" & fltr.txtinvi_DEF_DIrection.Tag & "'"
      End If
      If fltr.lblinvi_DEF_Uprev.Value = vbChecked Then
        f = f & " and invi_DEF_Uprev_ID='" & fltr.txtinvi_DEF_Uprev.Tag & "'"
      End If
      If fltr.lblinvi_DEF_EndDate_LE.Value = vbChecked Then
        f = f & " and invi_DEF_EndDate<=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpinvi_DEF_EndDate_LE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpinvi_DEF_EndDate_LE.Value), MakePGSQLDate(fltr.dtpinvi_DEF_EndDate_LE.Value)))
      End If
      If fltr.lblinvi_DEF_TheOrg.Value = vbChecked Then
        f = f & " and invi_DEF_TheOrg_ID='" & fltr.txtinvi_DEF_TheOrg.Tag & "'"
      End If
    jfmnuAllINV_INV.jv.Filter.Add "AUTOinvi_DEF", f
    End If
      jfmnuAllINV_INV.jv.Refresh
      Me.MousePointer = vbNormal
    End If
    jfmnuAllINV_INV.Show
    jfmnuAllINV_INV.WindowState = 0
    jfmnuAllINV_INV.ZOrder 0
End Sub
Private Sub jfmnuAllINV_INV_OnFilter(usedefault As Boolean)
    Dim fltr As frmINV_INV
    Dim f As String
    Set fltr = New frmINV_INV
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
    f = "1=1"
      If fltr.lblinvi_DEF_Building.Value = vbChecked Then
        f = f & " and invi_DEF_Building_ID='" & fltr.txtinvi_DEF_Building.Tag & "'"
      End If
      If fltr.lblinvi_DEF_TheFlow.Value = vbChecked Then
        f = f & " and invi_DEF_TheFlow like '%" & fltr.txtinvi_DEF_TheFlow.Text & "%'"
      End If
      If fltr.lblinvi_DEF_Otdel.Value = vbChecked Then
        f = f & " and invi_DEF_Otdel_ID='" & fltr.txtinvi_DEF_Otdel.Tag & "'"
      End If
      If fltr.lblinvi_DEF_TheOwner.Value = vbChecked Then
        f = f & " and invi_DEF_TheOwner_ID='" & fltr.txtinvi_DEF_TheOwner.Tag & "'"
      End If
      If fltr.lblinvi_DEF_MatOtv.Value = vbChecked Then
        f = f & " and invi_DEF_MatOtv_ID='" & fltr.txtinvi_DEF_MatOtv.Tag & "'"
      End If
      If fltr.lblinvi_DEF_Info.Value = vbChecked Then
        f = f & " and invi_DEF_Info like '%" & fltr.txtinvi_DEF_Info.Text & "%'"
      End If
      If fltr.lblinvi_DEF_TheRoom.Value = vbChecked Then
        f = f & " and invi_DEF_TheRoom like '%" & fltr.txtinvi_DEF_TheRoom.Text & "%'"
      End If
      If fltr.lblinvi_DEF_TheWorkPlace.Value = vbChecked Then
        f = f & " and invi_DEF_TheWorkPlace like '%" & fltr.txtinvi_DEF_TheWorkPlace.Text & "%'"
      End If
      If fltr.lblinvi_DEF_StartDate_LE.Value = vbChecked Then
        f = f & " and invi_DEF_StartDate<=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpinvi_DEF_StartDate_LE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpinvi_DEF_StartDate_LE.Value), MakePGSQLDate(fltr.dtpinvi_DEF_StartDate_LE.Value)))
      End If
      If fltr.lblinvi_DEF_EndDate_GE.Value = vbChecked Then
        f = f & " and invi_DEF_EndDate>=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpinvi_DEF_EndDate_GE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpinvi_DEF_EndDate_GE.Value), MakePGSQLDate(fltr.dtpinvi_DEF_EndDate_GE.Value)))
      End If
      If fltr.lblinvi_DEF_OrderNum.Value = vbChecked Then
        f = f & " and invi_DEF_OrderNum like '%" & fltr.txtinvi_DEF_OrderNum.Text & "%'"
      End If
      If fltr.lblinvi_DEF_StartDate_GE.Value = vbChecked Then
        f = f & " and invi_DEF_StartDate>=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpinvi_DEF_StartDate_GE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpinvi_DEF_StartDate_GE.Value), MakePGSQLDate(fltr.dtpinvi_DEF_StartDate_GE.Value)))
      End If
      If fltr.lblinvi_DEF_DIrection.Value = vbChecked Then
        f = f & " and invi_DEF_DIrection_ID='" & fltr.txtinvi_DEF_DIrection.Tag & "'"
      End If
      If fltr.lblinvi_DEF_Uprev.Value = vbChecked Then
        f = f & " and invi_DEF_Uprev_ID='" & fltr.txtinvi_DEF_Uprev.Tag & "'"
      End If
      If fltr.lblinvi_DEF_EndDate_LE.Value = vbChecked Then
        f = f & " and invi_DEF_EndDate<=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpinvi_DEF_EndDate_LE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpinvi_DEF_EndDate_LE.Value), MakePGSQLDate(fltr.dtpinvi_DEF_EndDate_LE.Value)))
      End If
      If fltr.lblinvi_DEF_TheOrg.Value = vbChecked Then
        f = f & " and invi_DEF_TheOrg_ID='" & fltr.txtinvi_DEF_TheOrg.Tag & "'"
      End If
    jfmnuAllINV_INV.jv.Filter.Add "AUTOinvi_DEF", f
    End If
    Unload fltr
    usedefault = False
End Sub
Private Sub jfmnuAllINV_INV_OnAdd(usedefaut As Boolean, Refesh As Boolean)
  Dim objGui  As Object
  Dim o As Object
  Dim id As String
  id = CreateGUID2
  Manager.NewInstance id, "INV_INV", "Инвентаризация" & Now, Site
  Set o = Manager.GetInstanceObject(id)
  If IsDocDenied(o) Then
    MsgBox "Не разрешен доступ к документам такого типа"
    Exit Sub
  End If

  Dim g  As Object
  Set g = Manager.GetInstanceGUI(o.id)
  If Not g Is Nothing Then
    g.Show GetDocumentMode(o), o, False
  End If
  usedefaut = False
  Refesh = False
End Sub


Private Sub mnuINV_INV_1_Click()
    Dim journal As Object
    On Error Resume Next
    If jfmnuINV_INV_1 Is Nothing Then
      Set jfmnuINV_INV_1 = New frmJournalShow
      Set journal = Manager.GetInstanceObject("{D743AE87-E934-4691-B1A3-00BAC0E83F0C}")
      Manager.LockInstanceObject journal.id
      Set jfmnuINV_INV_1.jv.journal = journal
      jfmnuINV_INV_1.jv.OpenModal = False
      jfmnuINV_INV_1.Caption = "Инвентаризация :Утверждена"
      Me.MousePointer = vbHourglass
      DoEvents
      Dim f As String
    f = "1=1"
   f = " INTSANCEStatusID='{706FBA86-116E-4CF4-932E-32CF7DEBC573}'"
    jfmnuINV_INV_1.jv.Filter.Add "AUTOinvi_DEF", f
    Dim fltr As frmINV_INV
    Set fltr = New frmINV_INV
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
      If fltr.lblinvi_DEF_Building.Value = vbChecked Then
        f = f & " and invi_DEF_Building_ID='" & fltr.txtinvi_DEF_Building.Tag & "'"
      End If
      If fltr.lblinvi_DEF_TheFlow.Value = vbChecked Then
        f = f & " and invi_DEF_TheFlow like '%" & fltr.txtinvi_DEF_TheFlow.Text & "%'"
      End If
      If fltr.lblinvi_DEF_Otdel.Value = vbChecked Then
        f = f & " and invi_DEF_Otdel_ID='" & fltr.txtinvi_DEF_Otdel.Tag & "'"
      End If
      If fltr.lblinvi_DEF_TheOwner.Value = vbChecked Then
        f = f & " and invi_DEF_TheOwner_ID='" & fltr.txtinvi_DEF_TheOwner.Tag & "'"
      End If
      If fltr.lblinvi_DEF_MatOtv.Value = vbChecked Then
        f = f & " and invi_DEF_MatOtv_ID='" & fltr.txtinvi_DEF_MatOtv.Tag & "'"
      End If
      If fltr.lblinvi_DEF_Info.Value = vbChecked Then
        f = f & " and invi_DEF_Info like '%" & fltr.txtinvi_DEF_Info.Text & "%'"
      End If
      If fltr.lblinvi_DEF_TheRoom.Value = vbChecked Then
        f = f & " and invi_DEF_TheRoom like '%" & fltr.txtinvi_DEF_TheRoom.Text & "%'"
      End If
      If fltr.lblinvi_DEF_TheWorkPlace.Value = vbChecked Then
        f = f & " and invi_DEF_TheWorkPlace like '%" & fltr.txtinvi_DEF_TheWorkPlace.Text & "%'"
      End If
      If fltr.lblinvi_DEF_StartDate_LE.Value = vbChecked Then
        f = f & " and invi_DEF_StartDate<=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpinvi_DEF_StartDate_LE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpinvi_DEF_StartDate_LE.Value), MakePGSQLDate(fltr.dtpinvi_DEF_StartDate_LE.Value)))
      End If
      If fltr.lblinvi_DEF_EndDate_GE.Value = vbChecked Then
        f = f & " and invi_DEF_EndDate>=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpinvi_DEF_EndDate_GE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpinvi_DEF_EndDate_GE.Value), MakePGSQLDate(fltr.dtpinvi_DEF_EndDate_GE.Value)))
      End If
      If fltr.lblinvi_DEF_OrderNum.Value = vbChecked Then
        f = f & " and invi_DEF_OrderNum like '%" & fltr.txtinvi_DEF_OrderNum.Text & "%'"
      End If
      If fltr.lblinvi_DEF_StartDate_GE.Value = vbChecked Then
        f = f & " and invi_DEF_StartDate>=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpinvi_DEF_StartDate_GE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpinvi_DEF_StartDate_GE.Value), MakePGSQLDate(fltr.dtpinvi_DEF_StartDate_GE.Value)))
      End If
      If fltr.lblinvi_DEF_DIrection.Value = vbChecked Then
        f = f & " and invi_DEF_DIrection_ID='" & fltr.txtinvi_DEF_DIrection.Tag & "'"
      End If
      If fltr.lblinvi_DEF_Uprev.Value = vbChecked Then
        f = f & " and invi_DEF_Uprev_ID='" & fltr.txtinvi_DEF_Uprev.Tag & "'"
      End If
      If fltr.lblinvi_DEF_EndDate_LE.Value = vbChecked Then
        f = f & " and invi_DEF_EndDate<=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpinvi_DEF_EndDate_LE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpinvi_DEF_EndDate_LE.Value), MakePGSQLDate(fltr.dtpinvi_DEF_EndDate_LE.Value)))
      End If
      If fltr.lblinvi_DEF_TheOrg.Value = vbChecked Then
        f = f & " and invi_DEF_TheOrg_ID='" & fltr.txtinvi_DEF_TheOrg.Tag & "'"
      End If
    jfmnuINV_INV_1.jv.Filter.Add "AUTOinvi_DEF", f
    End If
      jfmnuINV_INV_1.jv.Refresh
      Me.MousePointer = vbNormal
    End If
    jfmnuINV_INV_1.Show
    jfmnuINV_INV_1.WindowState = 0
    jfmnuINV_INV_1.ZOrder 0
End Sub
Private Sub jfmnuINV_INV_1_OnFilter(usedefault As Boolean)
    Dim fltr As frmINV_INV
    Dim f As String
    Set fltr = New frmINV_INV
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
    f = "1=1"
   f = " INTSANCEStatusID='{706FBA86-116E-4CF4-932E-32CF7DEBC573}'"
      If fltr.lblinvi_DEF_Building.Value = vbChecked Then
        f = f & " and invi_DEF_Building_ID='" & fltr.txtinvi_DEF_Building.Tag & "'"
      End If
      If fltr.lblinvi_DEF_TheFlow.Value = vbChecked Then
        f = f & " and invi_DEF_TheFlow like '%" & fltr.txtinvi_DEF_TheFlow.Text & "%'"
      End If
      If fltr.lblinvi_DEF_Otdel.Value = vbChecked Then
        f = f & " and invi_DEF_Otdel_ID='" & fltr.txtinvi_DEF_Otdel.Tag & "'"
      End If
      If fltr.lblinvi_DEF_TheOwner.Value = vbChecked Then
        f = f & " and invi_DEF_TheOwner_ID='" & fltr.txtinvi_DEF_TheOwner.Tag & "'"
      End If
      If fltr.lblinvi_DEF_MatOtv.Value = vbChecked Then
        f = f & " and invi_DEF_MatOtv_ID='" & fltr.txtinvi_DEF_MatOtv.Tag & "'"
      End If
      If fltr.lblinvi_DEF_Info.Value = vbChecked Then
        f = f & " and invi_DEF_Info like '%" & fltr.txtinvi_DEF_Info.Text & "%'"
      End If
      If fltr.lblinvi_DEF_TheRoom.Value = vbChecked Then
        f = f & " and invi_DEF_TheRoom like '%" & fltr.txtinvi_DEF_TheRoom.Text & "%'"
      End If
      If fltr.lblinvi_DEF_TheWorkPlace.Value = vbChecked Then
        f = f & " and invi_DEF_TheWorkPlace like '%" & fltr.txtinvi_DEF_TheWorkPlace.Text & "%'"
      End If
      If fltr.lblinvi_DEF_StartDate_LE.Value = vbChecked Then
        f = f & " and invi_DEF_StartDate<=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpinvi_DEF_StartDate_LE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpinvi_DEF_StartDate_LE.Value), MakePGSQLDate(fltr.dtpinvi_DEF_StartDate_LE.Value)))
      End If
      If fltr.lblinvi_DEF_EndDate_GE.Value = vbChecked Then
        f = f & " and invi_DEF_EndDate>=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpinvi_DEF_EndDate_GE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpinvi_DEF_EndDate_GE.Value), MakePGSQLDate(fltr.dtpinvi_DEF_EndDate_GE.Value)))
      End If
      If fltr.lblinvi_DEF_OrderNum.Value = vbChecked Then
        f = f & " and invi_DEF_OrderNum like '%" & fltr.txtinvi_DEF_OrderNum.Text & "%'"
      End If
      If fltr.lblinvi_DEF_StartDate_GE.Value = vbChecked Then
        f = f & " and invi_DEF_StartDate>=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpinvi_DEF_StartDate_GE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpinvi_DEF_StartDate_GE.Value), MakePGSQLDate(fltr.dtpinvi_DEF_StartDate_GE.Value)))
      End If
      If fltr.lblinvi_DEF_DIrection.Value = vbChecked Then
        f = f & " and invi_DEF_DIrection_ID='" & fltr.txtinvi_DEF_DIrection.Tag & "'"
      End If
      If fltr.lblinvi_DEF_Uprev.Value = vbChecked Then
        f = f & " and invi_DEF_Uprev_ID='" & fltr.txtinvi_DEF_Uprev.Tag & "'"
      End If
      If fltr.lblinvi_DEF_EndDate_LE.Value = vbChecked Then
        f = f & " and invi_DEF_EndDate<=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpinvi_DEF_EndDate_LE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpinvi_DEF_EndDate_LE.Value), MakePGSQLDate(fltr.dtpinvi_DEF_EndDate_LE.Value)))
      End If
      If fltr.lblinvi_DEF_TheOrg.Value = vbChecked Then
        f = f & " and invi_DEF_TheOrg_ID='" & fltr.txtinvi_DEF_TheOrg.Tag & "'"
      End If
    jfmnuINV_INV_1.jv.Filter.Add "AUTOinvi_DEF", f
    End If
    Unload fltr
    usedefault = False
End Sub
Private Sub jfmnuINV_INV_1_OnClearFilter()
   jfmnuINV_INV_1.jv.Filter.Add "AUTOinvi_DEF", " INTSANCEStatusID='{706FBA86-116E-4CF4-932E-32CF7DEBC573}'"
End Sub
Private Sub jfmnuINV_INV_1_OnAdd(usedefaut As Boolean, Refesh As Boolean)
  Dim objGui  As Object
  Dim o As Object
  Dim id As String
  id = CreateGUID2
  Manager.NewInstance id, "INV_INV", "Инвентаризация" & Now, Site
  Set o = Manager.GetInstanceObject(id)
  If IsDocDenied(o) Then
    MsgBox "Не разрешен доступ к документам такого типа"
    Exit Sub
  End If

  Dim g  As Object
  Set g = Manager.GetInstanceGUI(o.id)
  If Not g Is Nothing Then
    g.Show GetDocumentMode(o), o, False
  End If
  usedefaut = False
  Refesh = False
End Sub


Private Sub mnuINV_INV_2_Click()
    Dim journal As Object
    On Error Resume Next
    If jfmnuINV_INV_2 Is Nothing Then
      Set jfmnuINV_INV_2 = New frmJournalShow
      Set journal = Manager.GetInstanceObject("{D743AE87-E934-4691-B1A3-00BAC0E83F0C}")
      Manager.LockInstanceObject journal.id
      Set jfmnuINV_INV_2.jv.journal = journal
      jfmnuINV_INV_2.jv.OpenModal = False
      jfmnuINV_INV_2.Caption = "Инвентаризация :Инвентаризация завершена"
      Me.MousePointer = vbHourglass
      DoEvents
      Dim f As String
    f = "1=1"
   f = " INTSANCEStatusID='{03A3E27E-FF6E-4325-8174-462D88422A0E}'"
    jfmnuINV_INV_2.jv.Filter.Add "AUTOinvi_DEF", f
    Dim fltr As frmINV_INV
    Set fltr = New frmINV_INV
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
      If fltr.lblinvi_DEF_Building.Value = vbChecked Then
        f = f & " and invi_DEF_Building_ID='" & fltr.txtinvi_DEF_Building.Tag & "'"
      End If
      If fltr.lblinvi_DEF_TheFlow.Value = vbChecked Then
        f = f & " and invi_DEF_TheFlow like '%" & fltr.txtinvi_DEF_TheFlow.Text & "%'"
      End If
      If fltr.lblinvi_DEF_Otdel.Value = vbChecked Then
        f = f & " and invi_DEF_Otdel_ID='" & fltr.txtinvi_DEF_Otdel.Tag & "'"
      End If
      If fltr.lblinvi_DEF_TheOwner.Value = vbChecked Then
        f = f & " and invi_DEF_TheOwner_ID='" & fltr.txtinvi_DEF_TheOwner.Tag & "'"
      End If
      If fltr.lblinvi_DEF_MatOtv.Value = vbChecked Then
        f = f & " and invi_DEF_MatOtv_ID='" & fltr.txtinvi_DEF_MatOtv.Tag & "'"
      End If
      If fltr.lblinvi_DEF_Info.Value = vbChecked Then
        f = f & " and invi_DEF_Info like '%" & fltr.txtinvi_DEF_Info.Text & "%'"
      End If
      If fltr.lblinvi_DEF_TheRoom.Value = vbChecked Then
        f = f & " and invi_DEF_TheRoom like '%" & fltr.txtinvi_DEF_TheRoom.Text & "%'"
      End If
      If fltr.lblinvi_DEF_TheWorkPlace.Value = vbChecked Then
        f = f & " and invi_DEF_TheWorkPlace like '%" & fltr.txtinvi_DEF_TheWorkPlace.Text & "%'"
      End If
      If fltr.lblinvi_DEF_StartDate_LE.Value = vbChecked Then
        f = f & " and invi_DEF_StartDate<=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpinvi_DEF_StartDate_LE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpinvi_DEF_StartDate_LE.Value), MakePGSQLDate(fltr.dtpinvi_DEF_StartDate_LE.Value)))
      End If
      If fltr.lblinvi_DEF_EndDate_GE.Value = vbChecked Then
        f = f & " and invi_DEF_EndDate>=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpinvi_DEF_EndDate_GE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpinvi_DEF_EndDate_GE.Value), MakePGSQLDate(fltr.dtpinvi_DEF_EndDate_GE.Value)))
      End If
      If fltr.lblinvi_DEF_OrderNum.Value = vbChecked Then
        f = f & " and invi_DEF_OrderNum like '%" & fltr.txtinvi_DEF_OrderNum.Text & "%'"
      End If
      If fltr.lblinvi_DEF_StartDate_GE.Value = vbChecked Then
        f = f & " and invi_DEF_StartDate>=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpinvi_DEF_StartDate_GE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpinvi_DEF_StartDate_GE.Value), MakePGSQLDate(fltr.dtpinvi_DEF_StartDate_GE.Value)))
      End If
      If fltr.lblinvi_DEF_DIrection.Value = vbChecked Then
        f = f & " and invi_DEF_DIrection_ID='" & fltr.txtinvi_DEF_DIrection.Tag & "'"
      End If
      If fltr.lblinvi_DEF_Uprev.Value = vbChecked Then
        f = f & " and invi_DEF_Uprev_ID='" & fltr.txtinvi_DEF_Uprev.Tag & "'"
      End If
      If fltr.lblinvi_DEF_EndDate_LE.Value = vbChecked Then
        f = f & " and invi_DEF_EndDate<=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpinvi_DEF_EndDate_LE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpinvi_DEF_EndDate_LE.Value), MakePGSQLDate(fltr.dtpinvi_DEF_EndDate_LE.Value)))
      End If
      If fltr.lblinvi_DEF_TheOrg.Value = vbChecked Then
        f = f & " and invi_DEF_TheOrg_ID='" & fltr.txtinvi_DEF_TheOrg.Tag & "'"
      End If
    jfmnuINV_INV_2.jv.Filter.Add "AUTOinvi_DEF", f
    End If
      jfmnuINV_INV_2.jv.Refresh
      Me.MousePointer = vbNormal
    End If
    jfmnuINV_INV_2.Show
    jfmnuINV_INV_2.WindowState = 0
    jfmnuINV_INV_2.ZOrder 0
End Sub
Private Sub jfmnuINV_INV_2_OnFilter(usedefault As Boolean)
    Dim fltr As frmINV_INV
    Dim f As String
    Set fltr = New frmINV_INV
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
    f = "1=1"
   f = " INTSANCEStatusID='{03A3E27E-FF6E-4325-8174-462D88422A0E}'"
      If fltr.lblinvi_DEF_Building.Value = vbChecked Then
        f = f & " and invi_DEF_Building_ID='" & fltr.txtinvi_DEF_Building.Tag & "'"
      End If
      If fltr.lblinvi_DEF_TheFlow.Value = vbChecked Then
        f = f & " and invi_DEF_TheFlow like '%" & fltr.txtinvi_DEF_TheFlow.Text & "%'"
      End If
      If fltr.lblinvi_DEF_Otdel.Value = vbChecked Then
        f = f & " and invi_DEF_Otdel_ID='" & fltr.txtinvi_DEF_Otdel.Tag & "'"
      End If
      If fltr.lblinvi_DEF_TheOwner.Value = vbChecked Then
        f = f & " and invi_DEF_TheOwner_ID='" & fltr.txtinvi_DEF_TheOwner.Tag & "'"
      End If
      If fltr.lblinvi_DEF_MatOtv.Value = vbChecked Then
        f = f & " and invi_DEF_MatOtv_ID='" & fltr.txtinvi_DEF_MatOtv.Tag & "'"
      End If
      If fltr.lblinvi_DEF_Info.Value = vbChecked Then
        f = f & " and invi_DEF_Info like '%" & fltr.txtinvi_DEF_Info.Text & "%'"
      End If
      If fltr.lblinvi_DEF_TheRoom.Value = vbChecked Then
        f = f & " and invi_DEF_TheRoom like '%" & fltr.txtinvi_DEF_TheRoom.Text & "%'"
      End If
      If fltr.lblinvi_DEF_TheWorkPlace.Value = vbChecked Then
        f = f & " and invi_DEF_TheWorkPlace like '%" & fltr.txtinvi_DEF_TheWorkPlace.Text & "%'"
      End If
      If fltr.lblinvi_DEF_StartDate_LE.Value = vbChecked Then
        f = f & " and invi_DEF_StartDate<=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpinvi_DEF_StartDate_LE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpinvi_DEF_StartDate_LE.Value), MakePGSQLDate(fltr.dtpinvi_DEF_StartDate_LE.Value)))
      End If
      If fltr.lblinvi_DEF_EndDate_GE.Value = vbChecked Then
        f = f & " and invi_DEF_EndDate>=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpinvi_DEF_EndDate_GE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpinvi_DEF_EndDate_GE.Value), MakePGSQLDate(fltr.dtpinvi_DEF_EndDate_GE.Value)))
      End If
      If fltr.lblinvi_DEF_OrderNum.Value = vbChecked Then
        f = f & " and invi_DEF_OrderNum like '%" & fltr.txtinvi_DEF_OrderNum.Text & "%'"
      End If
      If fltr.lblinvi_DEF_StartDate_GE.Value = vbChecked Then
        f = f & " and invi_DEF_StartDate>=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpinvi_DEF_StartDate_GE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpinvi_DEF_StartDate_GE.Value), MakePGSQLDate(fltr.dtpinvi_DEF_StartDate_GE.Value)))
      End If
      If fltr.lblinvi_DEF_DIrection.Value = vbChecked Then
        f = f & " and invi_DEF_DIrection_ID='" & fltr.txtinvi_DEF_DIrection.Tag & "'"
      End If
      If fltr.lblinvi_DEF_Uprev.Value = vbChecked Then
        f = f & " and invi_DEF_Uprev_ID='" & fltr.txtinvi_DEF_Uprev.Tag & "'"
      End If
      If fltr.lblinvi_DEF_EndDate_LE.Value = vbChecked Then
        f = f & " and invi_DEF_EndDate<=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpinvi_DEF_EndDate_LE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpinvi_DEF_EndDate_LE.Value), MakePGSQLDate(fltr.dtpinvi_DEF_EndDate_LE.Value)))
      End If
      If fltr.lblinvi_DEF_TheOrg.Value = vbChecked Then
        f = f & " and invi_DEF_TheOrg_ID='" & fltr.txtinvi_DEF_TheOrg.Tag & "'"
      End If
    jfmnuINV_INV_2.jv.Filter.Add "AUTOinvi_DEF", f
    End If
    Unload fltr
    usedefault = False
End Sub
Private Sub jfmnuINV_INV_2_OnClearFilter()
   jfmnuINV_INV_2.jv.Filter.Add "AUTOinvi_DEF", " INTSANCEStatusID='{03A3E27E-FF6E-4325-8174-462D88422A0E}'"
End Sub
Private Sub jfmnuINV_INV_2_OnAdd(usedefaut As Boolean, Refesh As Boolean)
  Dim objGui  As Object
  Dim o As Object
  Dim id As String
  id = CreateGUID2
  Manager.NewInstance id, "INV_INV", "Инвентаризация" & Now, Site
  Set o = Manager.GetInstanceObject(id)
  If IsDocDenied(o) Then
    MsgBox "Не разрешен доступ к документам такого типа"
    Exit Sub
  End If

  Dim g  As Object
  Set g = Manager.GetInstanceGUI(o.id)
  If Not g Is Nothing Then
    g.Show GetDocumentMode(o), o, False
  End If
  usedefaut = False
  Refesh = False
End Sub


Private Sub mnuINV_INV_3_Click()
    Dim journal As Object
    On Error Resume Next
    If jfmnuINV_INV_3 Is Nothing Then
      Set jfmnuINV_INV_3 = New frmJournalShow
      Set journal = Manager.GetInstanceObject("{D743AE87-E934-4691-B1A3-00BAC0E83F0C}")
      Manager.LockInstanceObject journal.id
      Set jfmnuINV_INV_3.jv.journal = journal
      jfmnuINV_INV_3.jv.OpenModal = False
      jfmnuINV_INV_3.Caption = "Инвентаризация :Оформляется"
      Me.MousePointer = vbHourglass
      DoEvents
      Dim f As String
    f = "1=1"
   f = " INTSANCEStatusID='{926A2E1C-FBF5-44A4-9536-E195AF47D32F}'"
    jfmnuINV_INV_3.jv.Filter.Add "AUTOinvi_DEF", f
    Dim fltr As frmINV_INV
    Set fltr = New frmINV_INV
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
      If fltr.lblinvi_DEF_Building.Value = vbChecked Then
        f = f & " and invi_DEF_Building_ID='" & fltr.txtinvi_DEF_Building.Tag & "'"
      End If
      If fltr.lblinvi_DEF_TheFlow.Value = vbChecked Then
        f = f & " and invi_DEF_TheFlow like '%" & fltr.txtinvi_DEF_TheFlow.Text & "%'"
      End If
      If fltr.lblinvi_DEF_Otdel.Value = vbChecked Then
        f = f & " and invi_DEF_Otdel_ID='" & fltr.txtinvi_DEF_Otdel.Tag & "'"
      End If
      If fltr.lblinvi_DEF_TheOwner.Value = vbChecked Then
        f = f & " and invi_DEF_TheOwner_ID='" & fltr.txtinvi_DEF_TheOwner.Tag & "'"
      End If
      If fltr.lblinvi_DEF_MatOtv.Value = vbChecked Then
        f = f & " and invi_DEF_MatOtv_ID='" & fltr.txtinvi_DEF_MatOtv.Tag & "'"
      End If
      If fltr.lblinvi_DEF_Info.Value = vbChecked Then
        f = f & " and invi_DEF_Info like '%" & fltr.txtinvi_DEF_Info.Text & "%'"
      End If
      If fltr.lblinvi_DEF_TheRoom.Value = vbChecked Then
        f = f & " and invi_DEF_TheRoom like '%" & fltr.txtinvi_DEF_TheRoom.Text & "%'"
      End If
      If fltr.lblinvi_DEF_TheWorkPlace.Value = vbChecked Then
        f = f & " and invi_DEF_TheWorkPlace like '%" & fltr.txtinvi_DEF_TheWorkPlace.Text & "%'"
      End If
      If fltr.lblinvi_DEF_StartDate_LE.Value = vbChecked Then
        f = f & " and invi_DEF_StartDate<=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpinvi_DEF_StartDate_LE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpinvi_DEF_StartDate_LE.Value), MakePGSQLDate(fltr.dtpinvi_DEF_StartDate_LE.Value)))
      End If
      If fltr.lblinvi_DEF_EndDate_GE.Value = vbChecked Then
        f = f & " and invi_DEF_EndDate>=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpinvi_DEF_EndDate_GE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpinvi_DEF_EndDate_GE.Value), MakePGSQLDate(fltr.dtpinvi_DEF_EndDate_GE.Value)))
      End If
      If fltr.lblinvi_DEF_OrderNum.Value = vbChecked Then
        f = f & " and invi_DEF_OrderNum like '%" & fltr.txtinvi_DEF_OrderNum.Text & "%'"
      End If
      If fltr.lblinvi_DEF_StartDate_GE.Value = vbChecked Then
        f = f & " and invi_DEF_StartDate>=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpinvi_DEF_StartDate_GE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpinvi_DEF_StartDate_GE.Value), MakePGSQLDate(fltr.dtpinvi_DEF_StartDate_GE.Value)))
      End If
      If fltr.lblinvi_DEF_DIrection.Value = vbChecked Then
        f = f & " and invi_DEF_DIrection_ID='" & fltr.txtinvi_DEF_DIrection.Tag & "'"
      End If
      If fltr.lblinvi_DEF_Uprev.Value = vbChecked Then
        f = f & " and invi_DEF_Uprev_ID='" & fltr.txtinvi_DEF_Uprev.Tag & "'"
      End If
      If fltr.lblinvi_DEF_EndDate_LE.Value = vbChecked Then
        f = f & " and invi_DEF_EndDate<=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpinvi_DEF_EndDate_LE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpinvi_DEF_EndDate_LE.Value), MakePGSQLDate(fltr.dtpinvi_DEF_EndDate_LE.Value)))
      End If
      If fltr.lblinvi_DEF_TheOrg.Value = vbChecked Then
        f = f & " and invi_DEF_TheOrg_ID='" & fltr.txtinvi_DEF_TheOrg.Tag & "'"
      End If
    jfmnuINV_INV_3.jv.Filter.Add "AUTOinvi_DEF", f
    End If
      jfmnuINV_INV_3.jv.Refresh
      Me.MousePointer = vbNormal
    End If
    jfmnuINV_INV_3.Show
    jfmnuINV_INV_3.WindowState = 0
    jfmnuINV_INV_3.ZOrder 0
End Sub
Private Sub jfmnuINV_INV_3_OnFilter(usedefault As Boolean)
    Dim fltr As frmINV_INV
    Dim f As String
    Set fltr = New frmINV_INV
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
    f = "1=1"
   f = " INTSANCEStatusID='{926A2E1C-FBF5-44A4-9536-E195AF47D32F}'"
      If fltr.lblinvi_DEF_Building.Value = vbChecked Then
        f = f & " and invi_DEF_Building_ID='" & fltr.txtinvi_DEF_Building.Tag & "'"
      End If
      If fltr.lblinvi_DEF_TheFlow.Value = vbChecked Then
        f = f & " and invi_DEF_TheFlow like '%" & fltr.txtinvi_DEF_TheFlow.Text & "%'"
      End If
      If fltr.lblinvi_DEF_Otdel.Value = vbChecked Then
        f = f & " and invi_DEF_Otdel_ID='" & fltr.txtinvi_DEF_Otdel.Tag & "'"
      End If
      If fltr.lblinvi_DEF_TheOwner.Value = vbChecked Then
        f = f & " and invi_DEF_TheOwner_ID='" & fltr.txtinvi_DEF_TheOwner.Tag & "'"
      End If
      If fltr.lblinvi_DEF_MatOtv.Value = vbChecked Then
        f = f & " and invi_DEF_MatOtv_ID='" & fltr.txtinvi_DEF_MatOtv.Tag & "'"
      End If
      If fltr.lblinvi_DEF_Info.Value = vbChecked Then
        f = f & " and invi_DEF_Info like '%" & fltr.txtinvi_DEF_Info.Text & "%'"
      End If
      If fltr.lblinvi_DEF_TheRoom.Value = vbChecked Then
        f = f & " and invi_DEF_TheRoom like '%" & fltr.txtinvi_DEF_TheRoom.Text & "%'"
      End If
      If fltr.lblinvi_DEF_TheWorkPlace.Value = vbChecked Then
        f = f & " and invi_DEF_TheWorkPlace like '%" & fltr.txtinvi_DEF_TheWorkPlace.Text & "%'"
      End If
      If fltr.lblinvi_DEF_StartDate_LE.Value = vbChecked Then
        f = f & " and invi_DEF_StartDate<=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpinvi_DEF_StartDate_LE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpinvi_DEF_StartDate_LE.Value), MakePGSQLDate(fltr.dtpinvi_DEF_StartDate_LE.Value)))
      End If
      If fltr.lblinvi_DEF_EndDate_GE.Value = vbChecked Then
        f = f & " and invi_DEF_EndDate>=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpinvi_DEF_EndDate_GE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpinvi_DEF_EndDate_GE.Value), MakePGSQLDate(fltr.dtpinvi_DEF_EndDate_GE.Value)))
      End If
      If fltr.lblinvi_DEF_OrderNum.Value = vbChecked Then
        f = f & " and invi_DEF_OrderNum like '%" & fltr.txtinvi_DEF_OrderNum.Text & "%'"
      End If
      If fltr.lblinvi_DEF_StartDate_GE.Value = vbChecked Then
        f = f & " and invi_DEF_StartDate>=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpinvi_DEF_StartDate_GE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpinvi_DEF_StartDate_GE.Value), MakePGSQLDate(fltr.dtpinvi_DEF_StartDate_GE.Value)))
      End If
      If fltr.lblinvi_DEF_DIrection.Value = vbChecked Then
        f = f & " and invi_DEF_DIrection_ID='" & fltr.txtinvi_DEF_DIrection.Tag & "'"
      End If
      If fltr.lblinvi_DEF_Uprev.Value = vbChecked Then
        f = f & " and invi_DEF_Uprev_ID='" & fltr.txtinvi_DEF_Uprev.Tag & "'"
      End If
      If fltr.lblinvi_DEF_EndDate_LE.Value = vbChecked Then
        f = f & " and invi_DEF_EndDate<=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpinvi_DEF_EndDate_LE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpinvi_DEF_EndDate_LE.Value), MakePGSQLDate(fltr.dtpinvi_DEF_EndDate_LE.Value)))
      End If
      If fltr.lblinvi_DEF_TheOrg.Value = vbChecked Then
        f = f & " and invi_DEF_TheOrg_ID='" & fltr.txtinvi_DEF_TheOrg.Tag & "'"
      End If
    jfmnuINV_INV_3.jv.Filter.Add "AUTOinvi_DEF", f
    End If
    Unload fltr
    usedefault = False
End Sub
Private Sub jfmnuINV_INV_3_OnClearFilter()
   jfmnuINV_INV_3.jv.Filter.Add "AUTOinvi_DEF", " INTSANCEStatusID='{926A2E1C-FBF5-44A4-9536-E195AF47D32F}'"
End Sub
Private Sub jfmnuINV_INV_3_OnAdd(usedefaut As Boolean, Refesh As Boolean)
  Dim objGui  As Object
  Dim o As Object
  Dim id As String
  id = CreateGUID2
  Manager.NewInstance id, "INV_INV", "Инвентаризация" & Now, Site
  Set o = Manager.GetInstanceObject(id)
  If IsDocDenied(o) Then
    MsgBox "Не разрешен доступ к документам такого типа"
    Exit Sub
  End If

  Dim g  As Object
  Set g = Manager.GetInstanceGUI(o.id)
  If Not g Is Nothing Then
    g.Show GetDocumentMode(o), o, False
  End If
  usedefaut = False
  Refesh = False
End Sub


Private Sub mnuINV_INV_4_Click()
    Dim journal As Object
    On Error Resume Next
    If jfmnuINV_INV_4 Is Nothing Then
      Set jfmnuINV_INV_4 = New frmJournalShow
      Set journal = Manager.GetInstanceObject("{D743AE87-E934-4691-B1A3-00BAC0E83F0C}")
      Manager.LockInstanceObject journal.id
      Set jfmnuINV_INV_4.jv.journal = journal
      jfmnuINV_INV_4.jv.OpenModal = False
      jfmnuINV_INV_4.Caption = "Инвентаризация :Идет инвентаризация"
      Me.MousePointer = vbHourglass
      DoEvents
      Dim f As String
    f = "1=1"
   f = " INTSANCEStatusID='{FA929BE8-0966-46CD-99FC-FFF5E25EC4D5}'"
    jfmnuINV_INV_4.jv.Filter.Add "AUTOinvi_DEF", f
    Dim fltr As frmINV_INV
    Set fltr = New frmINV_INV
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
      If fltr.lblinvi_DEF_Building.Value = vbChecked Then
        f = f & " and invi_DEF_Building_ID='" & fltr.txtinvi_DEF_Building.Tag & "'"
      End If
      If fltr.lblinvi_DEF_TheFlow.Value = vbChecked Then
        f = f & " and invi_DEF_TheFlow like '%" & fltr.txtinvi_DEF_TheFlow.Text & "%'"
      End If
      If fltr.lblinvi_DEF_Otdel.Value = vbChecked Then
        f = f & " and invi_DEF_Otdel_ID='" & fltr.txtinvi_DEF_Otdel.Tag & "'"
      End If
      If fltr.lblinvi_DEF_TheOwner.Value = vbChecked Then
        f = f & " and invi_DEF_TheOwner_ID='" & fltr.txtinvi_DEF_TheOwner.Tag & "'"
      End If
      If fltr.lblinvi_DEF_MatOtv.Value = vbChecked Then
        f = f & " and invi_DEF_MatOtv_ID='" & fltr.txtinvi_DEF_MatOtv.Tag & "'"
      End If
      If fltr.lblinvi_DEF_Info.Value = vbChecked Then
        f = f & " and invi_DEF_Info like '%" & fltr.txtinvi_DEF_Info.Text & "%'"
      End If
      If fltr.lblinvi_DEF_TheRoom.Value = vbChecked Then
        f = f & " and invi_DEF_TheRoom like '%" & fltr.txtinvi_DEF_TheRoom.Text & "%'"
      End If
      If fltr.lblinvi_DEF_TheWorkPlace.Value = vbChecked Then
        f = f & " and invi_DEF_TheWorkPlace like '%" & fltr.txtinvi_DEF_TheWorkPlace.Text & "%'"
      End If
      If fltr.lblinvi_DEF_StartDate_LE.Value = vbChecked Then
        f = f & " and invi_DEF_StartDate<=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpinvi_DEF_StartDate_LE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpinvi_DEF_StartDate_LE.Value), MakePGSQLDate(fltr.dtpinvi_DEF_StartDate_LE.Value)))
      End If
      If fltr.lblinvi_DEF_EndDate_GE.Value = vbChecked Then
        f = f & " and invi_DEF_EndDate>=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpinvi_DEF_EndDate_GE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpinvi_DEF_EndDate_GE.Value), MakePGSQLDate(fltr.dtpinvi_DEF_EndDate_GE.Value)))
      End If
      If fltr.lblinvi_DEF_OrderNum.Value = vbChecked Then
        f = f & " and invi_DEF_OrderNum like '%" & fltr.txtinvi_DEF_OrderNum.Text & "%'"
      End If
      If fltr.lblinvi_DEF_StartDate_GE.Value = vbChecked Then
        f = f & " and invi_DEF_StartDate>=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpinvi_DEF_StartDate_GE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpinvi_DEF_StartDate_GE.Value), MakePGSQLDate(fltr.dtpinvi_DEF_StartDate_GE.Value)))
      End If
      If fltr.lblinvi_DEF_DIrection.Value = vbChecked Then
        f = f & " and invi_DEF_DIrection_ID='" & fltr.txtinvi_DEF_DIrection.Tag & "'"
      End If
      If fltr.lblinvi_DEF_Uprev.Value = vbChecked Then
        f = f & " and invi_DEF_Uprev_ID='" & fltr.txtinvi_DEF_Uprev.Tag & "'"
      End If
      If fltr.lblinvi_DEF_EndDate_LE.Value = vbChecked Then
        f = f & " and invi_DEF_EndDate<=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpinvi_DEF_EndDate_LE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpinvi_DEF_EndDate_LE.Value), MakePGSQLDate(fltr.dtpinvi_DEF_EndDate_LE.Value)))
      End If
      If fltr.lblinvi_DEF_TheOrg.Value = vbChecked Then
        f = f & " and invi_DEF_TheOrg_ID='" & fltr.txtinvi_DEF_TheOrg.Tag & "'"
      End If
    jfmnuINV_INV_4.jv.Filter.Add "AUTOinvi_DEF", f
    End If
      jfmnuINV_INV_4.jv.Refresh
      Me.MousePointer = vbNormal
    End If
    jfmnuINV_INV_4.Show
    jfmnuINV_INV_4.WindowState = 0
    jfmnuINV_INV_4.ZOrder 0
End Sub
Private Sub jfmnuINV_INV_4_OnFilter(usedefault As Boolean)
    Dim fltr As frmINV_INV
    Dim f As String
    Set fltr = New frmINV_INV
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
    f = "1=1"
   f = " INTSANCEStatusID='{FA929BE8-0966-46CD-99FC-FFF5E25EC4D5}'"
      If fltr.lblinvi_DEF_Building.Value = vbChecked Then
        f = f & " and invi_DEF_Building_ID='" & fltr.txtinvi_DEF_Building.Tag & "'"
      End If
      If fltr.lblinvi_DEF_TheFlow.Value = vbChecked Then
        f = f & " and invi_DEF_TheFlow like '%" & fltr.txtinvi_DEF_TheFlow.Text & "%'"
      End If
      If fltr.lblinvi_DEF_Otdel.Value = vbChecked Then
        f = f & " and invi_DEF_Otdel_ID='" & fltr.txtinvi_DEF_Otdel.Tag & "'"
      End If
      If fltr.lblinvi_DEF_TheOwner.Value = vbChecked Then
        f = f & " and invi_DEF_TheOwner_ID='" & fltr.txtinvi_DEF_TheOwner.Tag & "'"
      End If
      If fltr.lblinvi_DEF_MatOtv.Value = vbChecked Then
        f = f & " and invi_DEF_MatOtv_ID='" & fltr.txtinvi_DEF_MatOtv.Tag & "'"
      End If
      If fltr.lblinvi_DEF_Info.Value = vbChecked Then
        f = f & " and invi_DEF_Info like '%" & fltr.txtinvi_DEF_Info.Text & "%'"
      End If
      If fltr.lblinvi_DEF_TheRoom.Value = vbChecked Then
        f = f & " and invi_DEF_TheRoom like '%" & fltr.txtinvi_DEF_TheRoom.Text & "%'"
      End If
      If fltr.lblinvi_DEF_TheWorkPlace.Value = vbChecked Then
        f = f & " and invi_DEF_TheWorkPlace like '%" & fltr.txtinvi_DEF_TheWorkPlace.Text & "%'"
      End If
      If fltr.lblinvi_DEF_StartDate_LE.Value = vbChecked Then
        f = f & " and invi_DEF_StartDate<=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpinvi_DEF_StartDate_LE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpinvi_DEF_StartDate_LE.Value), MakePGSQLDate(fltr.dtpinvi_DEF_StartDate_LE.Value)))
      End If
      If fltr.lblinvi_DEF_EndDate_GE.Value = vbChecked Then
        f = f & " and invi_DEF_EndDate>=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpinvi_DEF_EndDate_GE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpinvi_DEF_EndDate_GE.Value), MakePGSQLDate(fltr.dtpinvi_DEF_EndDate_GE.Value)))
      End If
      If fltr.lblinvi_DEF_OrderNum.Value = vbChecked Then
        f = f & " and invi_DEF_OrderNum like '%" & fltr.txtinvi_DEF_OrderNum.Text & "%'"
      End If
      If fltr.lblinvi_DEF_StartDate_GE.Value = vbChecked Then
        f = f & " and invi_DEF_StartDate>=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpinvi_DEF_StartDate_GE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpinvi_DEF_StartDate_GE.Value), MakePGSQLDate(fltr.dtpinvi_DEF_StartDate_GE.Value)))
      End If
      If fltr.lblinvi_DEF_DIrection.Value = vbChecked Then
        f = f & " and invi_DEF_DIrection_ID='" & fltr.txtinvi_DEF_DIrection.Tag & "'"
      End If
      If fltr.lblinvi_DEF_Uprev.Value = vbChecked Then
        f = f & " and invi_DEF_Uprev_ID='" & fltr.txtinvi_DEF_Uprev.Tag & "'"
      End If
      If fltr.lblinvi_DEF_EndDate_LE.Value = vbChecked Then
        f = f & " and invi_DEF_EndDate<=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpinvi_DEF_EndDate_LE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpinvi_DEF_EndDate_LE.Value), MakePGSQLDate(fltr.dtpinvi_DEF_EndDate_LE.Value)))
      End If
      If fltr.lblinvi_DEF_TheOrg.Value = vbChecked Then
        f = f & " and invi_DEF_TheOrg_ID='" & fltr.txtinvi_DEF_TheOrg.Tag & "'"
      End If
    jfmnuINV_INV_4.jv.Filter.Add "AUTOinvi_DEF", f
    End If
    Unload fltr
    usedefault = False
End Sub
Private Sub jfmnuINV_INV_4_OnClearFilter()
   jfmnuINV_INV_4.jv.Filter.Add "AUTOinvi_DEF", " INTSANCEStatusID='{FA929BE8-0966-46CD-99FC-FFF5E25EC4D5}'"
End Sub
Private Sub jfmnuINV_INV_4_OnAdd(usedefaut As Boolean, Refesh As Boolean)
  Dim objGui  As Object
  Dim o As Object
  Dim id As String
  id = CreateGUID2
  Manager.NewInstance id, "INV_INV", "Инвентаризация" & Now, Site
  Set o = Manager.GetInstanceObject(id)
  If IsDocDenied(o) Then
    MsgBox "Не разрешен доступ к документам такого типа"
    Exit Sub
  End If

  Dim g  As Object
  Set g = Manager.GetInstanceGUI(o.id)
  If Not g Is Nothing Then
    g.Show GetDocumentMode(o), o, False
  End If
  usedefaut = False
  Refesh = False
End Sub






Private Sub AddColors(ByVal j As MTZJournal.JournalView)
  j.ColorColName = "Остаточный срок ПИ" '"INVOS_INFO_SrokOI"
  j.Colors.Add vbYellow, vbBlack, 0

End Sub



Private Sub mnuAllINV_OS_Click()
    Dim journal As Object
    On Error Resume Next
    If jfmnuAllINV_OS Is Nothing Then
      Set jfmnuAllINV_OS = New frmJournalShow
      Set journal = Manager.GetInstanceObject("{5CE3CC0B-5224-4145-B43F-6A29CC390C17}")
      Manager.LockInstanceObject journal.id
      Set jfmnuAllINV_OS.jv.journal = journal
      AddColors jfmnuAllINV_OS.jv
      jfmnuAllINV_OS.jv.OpenModal = False
      jfmnuAllINV_OS.Caption = "Карточка основного средства - все состояния"
      Me.MousePointer = vbHourglass
      DoEvents
      Dim f As String
    f = "1=1"
    Dim fltr As frmINV_OS
    Set fltr = New frmINV_OS
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
      If fltr.lblINVOS_SROK_RecalcDate_LE.Value = vbChecked Then
        f = f & " and INVOS_SROK_RecalcDate<=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_SROK_RecalcDate_LE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_SROK_RecalcDate_LE.Value), MakePGSQLDate(fltr.dtpINVOS_SROK_RecalcDate_LE.Value)))
      End If
      If fltr.lblINVOS_INFO_ShortName.Value = vbChecked Then
        f = f & " and INVOS_INFO_ShortName like '%" & fltr.txtINVOS_INFO_ShortName.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_ActivateDate_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_ActivateDate>=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_INFO_ActivateDate_GE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_INFO_ActivateDate_GE.Value), MakePGSQLDate(fltr.dtpINVOS_INFO_ActivateDate_GE.Value)))
      End If
      If fltr.lblINVOS_INFO_SrokFI_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokFI>=" & val(fltr.txtINVOS_INFO_SrokFI_GE.Text)
      End If
      If fltr.lblINVOS_PLACE_TheHouse.Value = vbChecked Then
        f = f & " and INVOS_PLACE_TheHouse_ID='" & fltr.txtINVOS_PLACE_TheHouse.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_InLineNum_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_InLineNum<=" & val(fltr.txtINVOS_INFO_InLineNum_LE.Text)
      End If
      If fltr.lblINVOS_INFO_Info.Value = vbChecked Then
        f = f & " and INVOS_INFO_Info like '%" & fltr.txtINVOS_INFO_Info.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_TheOrg.Value = vbChecked Then
        f = f & " and INVOS_INFO_TheOrg_ID='" & fltr.txtINVOS_INFO_TheOrg.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_SrokPI_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokPI>=" & val(fltr.txtINVOS_INFO_SrokPI_GE.Text)
      End If
      If fltr.lblINVOS_INFO_IsMaterial.Value = vbChecked Then
        f = f & " and INVOS_INFO_IsMaterial='" & fltr.cmbINVOS_INFO_IsMaterial.Text & "'"
      End If
      If fltr.lblINVOS_PLACE_WorkPlaceNum.Value = vbChecked Then
        f = f & " and INVOS_PLACE_WorkPlaceNum like '%" & fltr.txtINVOS_PLACE_WorkPlaceNum.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_MatOtv.Value = vbChecked Then
        f = f & " and INVOS_PLACE_MatOtv_ID='" & fltr.txtINVOS_PLACE_MatOtv.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_SrokFI_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokFI<=" & val(fltr.txtINVOS_INFO_SrokFI_LE.Text)
      End If
      If fltr.lblINVOS_INFO_SrokPI_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokPI<=" & val(fltr.txtINVOS_INFO_SrokPI_LE.Text)
      End If
      If fltr.lblINVOS_PLACE_ComplNumber.Value = vbChecked Then
        f = f & " and INVOS_PLACE_ComplNumber like '%" & fltr.txtINVOS_PLACE_ComplNumber.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_TheCost_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_TheCost>=" & val(fltr.txtINVOS_INFO_TheCost_GE.Text)
      End If
      If fltr.lblINVOS_PLACE_TheOwner.Value = vbChecked Then
        f = f & " and INVOS_PLACE_TheOwner_ID='" & fltr.txtINVOS_PLACE_TheOwner.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_TechFilePath.Value = vbChecked Then
        f = f & " and INVOS_INFO_TechFilePath like '%" & fltr.txtINVOS_INFO_TechFilePath.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_DIrection.Value = vbChecked Then
        f = f & " and INVOS_PLACE_DIrection_ID='" & fltr.txtINVOS_PLACE_DIrection.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_INVNum.Value = vbChecked Then
        f = f & " and INVOS_INFO_INVNum like '%" & fltr.txtINVOS_INFO_INVNum.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_Uprav.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Uprav_ID='" & fltr.txtINVOS_PLACE_Uprav.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_TheCost_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_TheCost<=" & val(fltr.txtINVOS_INFO_TheCost_LE.Text)
      End If
      If fltr.lblINVOS_INFO_SrokOI_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokOI<=" & val(fltr.txtINVOS_INFO_SrokOI_LE.Text)
      End If
      If fltr.lblINVOS_INFO_Name.Value = vbChecked Then
        f = f & " and INVOS_INFO_Name like '%" & fltr.txtINVOS_INFO_Name.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_CardNum.Value = vbChecked Then
        f = f & " and INVOS_INFO_CardNum like '%" & fltr.txtINVOS_INFO_CardNum.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_Otdel.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Otdel_ID='" & fltr.txtINVOS_PLACE_Otdel.Tag & "'"
      End If
      If fltr.lblINVOS_PLACE_Info.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Info like '%" & fltr.txtINVOS_PLACE_Info.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_OSType.Value = vbChecked Then
        f = f & " and INVOS_INFO_OSType_ID='" & fltr.txtINVOS_INFO_OSType.Tag & "'"
      End If
      If fltr.lblINVOS_PLACE_Room.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Room like '%" & fltr.txtINVOS_PLACE_Room.Text & "%'"
      End If
      If fltr.lblINVOS_SROK_RecalcDate_GE.Value = vbChecked Then
        f = f & " and INVOS_SROK_RecalcDate>=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_SROK_RecalcDate_GE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_SROK_RecalcDate_GE.Value), MakePGSQLDate(fltr.dtpINVOS_SROK_RecalcDate_GE.Value)))
      End If
      If fltr.lblINVOS_INFO_ActivateDate_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_ActivateDate<=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_INFO_ActivateDate_LE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_INFO_ActivateDate_LE.Value), MakePGSQLDate(fltr.dtpINVOS_INFO_ActivateDate_LE.Value)))
      End If
      If fltr.lblINVOS_INFO_InLineNum_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_InLineNum>=" & val(fltr.txtINVOS_INFO_InLineNum_GE.Text)
      End If
      If fltr.lblINVOS_INFO_SrokOI_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokOI>=" & val(fltr.txtINVOS_INFO_SrokOI_GE.Text)
      End If
      If fltr.lblINVOS_CODE_VisibleCode.Value = vbChecked Then
        f = f & " and INVOS_CODE_VisibleCode like '%" & fltr.txtINVOS_CODE_VisibleCode.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_Flow.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Flow like '%" & fltr.txtINVOS_PLACE_Flow.Text & "%'"
      End If
    jfmnuAllINV_OS.jv.Filter.Add "AUTOINVOS_INFO", f
    End If
      jfmnuAllINV_OS.jv.Refresh
      Me.MousePointer = vbNormal
    End If
   
    jfmnuAllINV_OS.Show
    jfmnuAllINV_OS.WindowState = 0
    jfmnuAllINV_OS.ZOrder 0
End Sub
Private Sub jfmnuAllINV_OS_OnFilter(usedefault As Boolean)
    Dim fltr As frmINV_OS
    Dim f As String
    Set fltr = New frmINV_OS
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
    f = "1=1"
      If fltr.lblINVOS_SROK_RecalcDate_LE.Value = vbChecked Then
        f = f & " and INVOS_SROK_RecalcDate<=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_SROK_RecalcDate_LE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_SROK_RecalcDate_LE.Value), MakePGSQLDate(fltr.dtpINVOS_SROK_RecalcDate_LE.Value)))
      End If
      If fltr.lblINVOS_INFO_ShortName.Value = vbChecked Then
        f = f & " and INVOS_INFO_ShortName like '%" & fltr.txtINVOS_INFO_ShortName.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_ActivateDate_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_ActivateDate>=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_INFO_ActivateDate_GE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_INFO_ActivateDate_GE.Value), MakePGSQLDate(fltr.dtpINVOS_INFO_ActivateDate_GE.Value)))
      End If
      If fltr.lblINVOS_INFO_SrokFI_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokFI>=" & val(fltr.txtINVOS_INFO_SrokFI_GE.Text)
      End If
      If fltr.lblINVOS_PLACE_TheHouse.Value = vbChecked Then
        f = f & " and INVOS_PLACE_TheHouse_ID='" & fltr.txtINVOS_PLACE_TheHouse.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_InLineNum_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_InLineNum<=" & val(fltr.txtINVOS_INFO_InLineNum_LE.Text)
      End If
      If fltr.lblINVOS_INFO_Info.Value = vbChecked Then
        f = f & " and INVOS_INFO_Info like '%" & fltr.txtINVOS_INFO_Info.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_TheOrg.Value = vbChecked Then
        f = f & " and INVOS_INFO_TheOrg_ID='" & fltr.txtINVOS_INFO_TheOrg.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_SrokPI_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokPI>=" & val(fltr.txtINVOS_INFO_SrokPI_GE.Text)
      End If
      If fltr.lblINVOS_INFO_IsMaterial.Value = vbChecked Then
        f = f & " and INVOS_INFO_IsMaterial='" & fltr.cmbINVOS_INFO_IsMaterial.Text & "'"
      End If
      If fltr.lblINVOS_PLACE_WorkPlaceNum.Value = vbChecked Then
        f = f & " and INVOS_PLACE_WorkPlaceNum like '%" & fltr.txtINVOS_PLACE_WorkPlaceNum.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_MatOtv.Value = vbChecked Then
        f = f & " and INVOS_PLACE_MatOtv_ID='" & fltr.txtINVOS_PLACE_MatOtv.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_SrokFI_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokFI<=" & val(fltr.txtINVOS_INFO_SrokFI_LE.Text)
      End If
      If fltr.lblINVOS_INFO_SrokPI_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokPI<=" & val(fltr.txtINVOS_INFO_SrokPI_LE.Text)
      End If
      If fltr.lblINVOS_PLACE_ComplNumber.Value = vbChecked Then
        f = f & " and INVOS_PLACE_ComplNumber like '%" & fltr.txtINVOS_PLACE_ComplNumber.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_TheCost_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_TheCost>=" & val(fltr.txtINVOS_INFO_TheCost_GE.Text)
      End If
      If fltr.lblINVOS_PLACE_TheOwner.Value = vbChecked Then
        f = f & " and INVOS_PLACE_TheOwner_ID='" & fltr.txtINVOS_PLACE_TheOwner.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_TechFilePath.Value = vbChecked Then
        f = f & " and INVOS_INFO_TechFilePath like '%" & fltr.txtINVOS_INFO_TechFilePath.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_DIrection.Value = vbChecked Then
        f = f & " and INVOS_PLACE_DIrection_ID='" & fltr.txtINVOS_PLACE_DIrection.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_INVNum.Value = vbChecked Then
        f = f & " and INVOS_INFO_INVNum like '%" & fltr.txtINVOS_INFO_INVNum.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_Uprav.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Uprav_ID='" & fltr.txtINVOS_PLACE_Uprav.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_TheCost_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_TheCost<=" & val(fltr.txtINVOS_INFO_TheCost_LE.Text)
      End If
      If fltr.lblINVOS_INFO_SrokOI_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokOI<=" & val(fltr.txtINVOS_INFO_SrokOI_LE.Text)
      End If
      If fltr.lblINVOS_INFO_Name.Value = vbChecked Then
        f = f & " and INVOS_INFO_Name like '%" & fltr.txtINVOS_INFO_Name.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_CardNum.Value = vbChecked Then
        f = f & " and INVOS_INFO_CardNum like '%" & fltr.txtINVOS_INFO_CardNum.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_Otdel.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Otdel_ID='" & fltr.txtINVOS_PLACE_Otdel.Tag & "'"
      End If
      If fltr.lblINVOS_PLACE_Info.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Info like '%" & fltr.txtINVOS_PLACE_Info.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_OSType.Value = vbChecked Then
        f = f & " and INVOS_INFO_OSType_ID='" & fltr.txtINVOS_INFO_OSType.Tag & "'"
      End If
      If fltr.lblINVOS_PLACE_Room.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Room like '%" & fltr.txtINVOS_PLACE_Room.Text & "%'"
      End If
      If fltr.lblINVOS_SROK_RecalcDate_GE.Value = vbChecked Then
        f = f & " and INVOS_SROK_RecalcDate>=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_SROK_RecalcDate_GE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_SROK_RecalcDate_GE.Value), MakePGSQLDate(fltr.dtpINVOS_SROK_RecalcDate_GE.Value)))
      End If
      If fltr.lblINVOS_INFO_ActivateDate_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_ActivateDate<=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_INFO_ActivateDate_LE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_INFO_ActivateDate_LE.Value), MakePGSQLDate(fltr.dtpINVOS_INFO_ActivateDate_LE.Value)))
      End If
      If fltr.lblINVOS_INFO_InLineNum_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_InLineNum>=" & val(fltr.txtINVOS_INFO_InLineNum_GE.Text)
      End If
      If fltr.lblINVOS_INFO_SrokOI_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokOI>=" & val(fltr.txtINVOS_INFO_SrokOI_GE.Text)
      End If
      If fltr.lblINVOS_CODE_VisibleCode.Value = vbChecked Then
        f = f & " and INVOS_CODE_VisibleCode like '%" & fltr.txtINVOS_CODE_VisibleCode.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_Flow.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Flow like '%" & fltr.txtINVOS_PLACE_Flow.Text & "%'"
      End If
    jfmnuAllINV_OS.jv.Filter.Add "AUTOINVOS_INFO", f
    End If
    Unload fltr
    usedefault = False
End Sub
Private Sub jfmnuAllINV_OS_OnAdd(usedefaut As Boolean, Refesh As Boolean)
  Dim objGui  As Object
  Dim o As Object
  Dim id As String
  id = CreateGUID2
  Manager.NewInstance id, "INV_OS", "Карточка основного средства" & Now, Site
  Set o = Manager.GetInstanceObject(id)
  If IsDocDenied(o) Then
    MsgBox "Не разрешен доступ к документам такого типа"
    Exit Sub
  End If

  Dim g  As Object
  Set g = Manager.GetInstanceGUI(o.id)
  If Not g Is Nothing Then
    g.Show GetDocumentMode(o), o, False
  End If
  usedefaut = False
  Refesh = False
End Sub


Private Sub mnuINV_OS_1_Click()
    Dim journal As Object
    On Error Resume Next
    If jfmnuINV_OS_1 Is Nothing Then
      Set jfmnuINV_OS_1 = New frmJournalShow
      Set journal = Manager.GetInstanceObject("{5CE3CC0B-5224-4145-B43F-6A29CC390C17}")
      Manager.LockInstanceObject journal.id
      Set jfmnuINV_OS_1.jv.journal = journal
       AddColors jfmnuINV_OS_1.jv
      jfmnuINV_OS_1.jv.OpenModal = False
      jfmnuINV_OS_1.Caption = "Карточка основного средства :В ремонте"
      Me.MousePointer = vbHourglass
      DoEvents
      Dim f As String
    f = "1=1"
   f = " INTSANCEStatusID='{8E6E78D2-82AA-4913-B08C-1230A8C8B4A9}'"
    jfmnuINV_OS_1.jv.Filter.Add "AUTOINVOS_INFO", f
    Dim fltr As frmINV_OS
    Set fltr = New frmINV_OS
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
      If fltr.lblINVOS_INFO_IsMaterial.Value = vbChecked Then
        f = f & " and INVOS_INFO_IsMaterial='" & fltr.cmbINVOS_INFO_IsMaterial.Text & "'"
      End If
      If fltr.lblINVOS_INFO_TheCost_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_TheCost<=" & val(fltr.txtINVOS_INFO_TheCost_LE.Text)
      End If
      If fltr.lblINVOS_INFO_SrokOI_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokOI<=" & val(fltr.txtINVOS_INFO_SrokOI_LE.Text)
      End If
      If fltr.lblINVOS_PLACE_ComplNumber.Value = vbChecked Then
        f = f & " and INVOS_PLACE_ComplNumber like '%" & fltr.txtINVOS_PLACE_ComplNumber.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_DIrection.Value = vbChecked Then
        f = f & " and INVOS_PLACE_DIrection_ID='" & fltr.txtINVOS_PLACE_DIrection.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_SrokOI_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokOI>=" & val(fltr.txtINVOS_INFO_SrokOI_GE.Text)
      End If
      If fltr.lblINVOS_INFO_InLineNum_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_InLineNum<=" & val(fltr.txtINVOS_INFO_InLineNum_LE.Text)
      End If
      If fltr.lblINVOS_INFO_Info.Value = vbChecked Then
        f = f & " and INVOS_INFO_Info like '%" & fltr.txtINVOS_INFO_Info.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_Otdel.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Otdel_ID='" & fltr.txtINVOS_PLACE_Otdel.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_ShortName.Value = vbChecked Then
        f = f & " and INVOS_INFO_ShortName like '%" & fltr.txtINVOS_INFO_ShortName.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_TheOrg.Value = vbChecked Then
        f = f & " and INVOS_INFO_TheOrg_ID='" & fltr.txtINVOS_INFO_TheOrg.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_ActivateDate_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_ActivateDate>=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_INFO_ActivateDate_GE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_INFO_ActivateDate_GE.Value), MakePGSQLDate(fltr.dtpINVOS_INFO_ActivateDate_GE.Value)))
      End If
      If fltr.lblINVOS_PLACE_TheOwner.Value = vbChecked Then
        f = f & " and INVOS_PLACE_TheOwner_ID='" & fltr.txtINVOS_PLACE_TheOwner.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_Name.Value = vbChecked Then
        f = f & " and INVOS_INFO_Name like '%" & fltr.txtINVOS_INFO_Name.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_TechFilePath.Value = vbChecked Then
        f = f & " and INVOS_INFO_TechFilePath like '%" & fltr.txtINVOS_INFO_TechFilePath.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_Room.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Room like '%" & fltr.txtINVOS_PLACE_Room.Text & "%'"
      End If
      If fltr.lblINVOS_SROK_RecalcDate_GE.Value = vbChecked Then
        f = f & " and INVOS_SROK_RecalcDate>=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_SROK_RecalcDate_GE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_SROK_RecalcDate_GE.Value), MakePGSQLDate(fltr.dtpINVOS_SROK_RecalcDate_GE.Value)))
      End If
      If fltr.lblINVOS_PLACE_MatOtv.Value = vbChecked Then
        f = f & " and INVOS_PLACE_MatOtv_ID='" & fltr.txtINVOS_PLACE_MatOtv.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_SrokPI_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokPI>=" & val(fltr.txtINVOS_INFO_SrokPI_GE.Text)
      End If
      If fltr.lblINVOS_INFO_OSType.Value = vbChecked Then
        f = f & " and INVOS_INFO_OSType_ID='" & fltr.txtINVOS_INFO_OSType.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_InLineNum_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_InLineNum>=" & val(fltr.txtINVOS_INFO_InLineNum_GE.Text)
      End If
      If fltr.lblINVOS_CODE_VisibleCode.Value = vbChecked Then
        f = f & " and INVOS_CODE_VisibleCode like '%" & fltr.txtINVOS_CODE_VisibleCode.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_SrokPI_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokPI<=" & val(fltr.txtINVOS_INFO_SrokPI_LE.Text)
      End If
      If fltr.lblINVOS_PLACE_Uprav.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Uprav_ID='" & fltr.txtINVOS_PLACE_Uprav.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_INVNum.Value = vbChecked Then
        f = f & " and INVOS_INFO_INVNum like '%" & fltr.txtINVOS_INFO_INVNum.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_SrokFI_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokFI<=" & val(fltr.txtINVOS_INFO_SrokFI_LE.Text)
      End If
      If fltr.lblINVOS_INFO_ActivateDate_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_ActivateDate<=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_INFO_ActivateDate_LE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_INFO_ActivateDate_LE.Value), MakePGSQLDate(fltr.dtpINVOS_INFO_ActivateDate_LE.Value)))
      End If
      If fltr.lblINVOS_INFO_SrokFI_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokFI>=" & val(fltr.txtINVOS_INFO_SrokFI_GE.Text)
      End If
      If fltr.lblINVOS_INFO_TheCost_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_TheCost>=" & val(fltr.txtINVOS_INFO_TheCost_GE.Text)
      End If
      If fltr.lblINVOS_PLACE_TheHouse.Value = vbChecked Then
        f = f & " and INVOS_PLACE_TheHouse_ID='" & fltr.txtINVOS_PLACE_TheHouse.Tag & "'"
      End If
      If fltr.lblINVOS_SROK_RecalcDate_LE.Value = vbChecked Then
        f = f & " and INVOS_SROK_RecalcDate<=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_SROK_RecalcDate_LE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_SROK_RecalcDate_LE.Value), MakePGSQLDate(fltr.dtpINVOS_SROK_RecalcDate_LE.Value)))
      End If
      If fltr.lblINVOS_INFO_CardNum.Value = vbChecked Then
        f = f & " and INVOS_INFO_CardNum like '%" & fltr.txtINVOS_INFO_CardNum.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_WorkPlaceNum.Value = vbChecked Then
        f = f & " and INVOS_PLACE_WorkPlaceNum like '%" & fltr.txtINVOS_PLACE_WorkPlaceNum.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_Info.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Info like '%" & fltr.txtINVOS_PLACE_Info.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_Flow.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Flow like '%" & fltr.txtINVOS_PLACE_Flow.Text & "%'"
      End If
    jfmnuINV_OS_1.jv.Filter.Add "AUTOINVOS_INFO", f
    End If
      jfmnuINV_OS_1.jv.Refresh
      Me.MousePointer = vbNormal
    End If
    jfmnuINV_OS_1.Show
    jfmnuINV_OS_1.WindowState = 0
    jfmnuINV_OS_1.ZOrder 0
End Sub
Private Sub jfmnuINV_OS_1_OnFilter(usedefault As Boolean)
    Dim fltr As frmINV_OS
    Dim f As String
    Set fltr = New frmINV_OS
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
    f = "1=1"
   f = " INTSANCEStatusID='{8E6E78D2-82AA-4913-B08C-1230A8C8B4A9}'"
      If fltr.lblINVOS_INFO_IsMaterial.Value = vbChecked Then
        f = f & " and INVOS_INFO_IsMaterial='" & fltr.cmbINVOS_INFO_IsMaterial.Text & "'"
      End If
      If fltr.lblINVOS_INFO_TheCost_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_TheCost<=" & val(fltr.txtINVOS_INFO_TheCost_LE.Text)
      End If
      If fltr.lblINVOS_INFO_SrokOI_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokOI<=" & val(fltr.txtINVOS_INFO_SrokOI_LE.Text)
      End If
      If fltr.lblINVOS_PLACE_ComplNumber.Value = vbChecked Then
        f = f & " and INVOS_PLACE_ComplNumber like '%" & fltr.txtINVOS_PLACE_ComplNumber.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_DIrection.Value = vbChecked Then
        f = f & " and INVOS_PLACE_DIrection_ID='" & fltr.txtINVOS_PLACE_DIrection.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_SrokOI_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokOI>=" & val(fltr.txtINVOS_INFO_SrokOI_GE.Text)
      End If
      If fltr.lblINVOS_INFO_InLineNum_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_InLineNum<=" & val(fltr.txtINVOS_INFO_InLineNum_LE.Text)
      End If
      If fltr.lblINVOS_INFO_Info.Value = vbChecked Then
        f = f & " and INVOS_INFO_Info like '%" & fltr.txtINVOS_INFO_Info.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_Otdel.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Otdel_ID='" & fltr.txtINVOS_PLACE_Otdel.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_ShortName.Value = vbChecked Then
        f = f & " and INVOS_INFO_ShortName like '%" & fltr.txtINVOS_INFO_ShortName.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_TheOrg.Value = vbChecked Then
        f = f & " and INVOS_INFO_TheOrg_ID='" & fltr.txtINVOS_INFO_TheOrg.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_ActivateDate_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_ActivateDate>=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_INFO_ActivateDate_GE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_INFO_ActivateDate_GE.Value), MakePGSQLDate(fltr.dtpINVOS_INFO_ActivateDate_GE.Value)))
      End If
      If fltr.lblINVOS_PLACE_TheOwner.Value = vbChecked Then
        f = f & " and INVOS_PLACE_TheOwner_ID='" & fltr.txtINVOS_PLACE_TheOwner.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_Name.Value = vbChecked Then
        f = f & " and INVOS_INFO_Name like '%" & fltr.txtINVOS_INFO_Name.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_TechFilePath.Value = vbChecked Then
        f = f & " and INVOS_INFO_TechFilePath like '%" & fltr.txtINVOS_INFO_TechFilePath.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_Room.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Room like '%" & fltr.txtINVOS_PLACE_Room.Text & "%'"
      End If
      If fltr.lblINVOS_SROK_RecalcDate_GE.Value = vbChecked Then
        f = f & " and INVOS_SROK_RecalcDate>=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_SROK_RecalcDate_GE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_SROK_RecalcDate_GE.Value), MakePGSQLDate(fltr.dtpINVOS_SROK_RecalcDate_GE.Value)))
      End If
      If fltr.lblINVOS_PLACE_MatOtv.Value = vbChecked Then
        f = f & " and INVOS_PLACE_MatOtv_ID='" & fltr.txtINVOS_PLACE_MatOtv.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_SrokPI_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokPI>=" & val(fltr.txtINVOS_INFO_SrokPI_GE.Text)
      End If
      If fltr.lblINVOS_INFO_OSType.Value = vbChecked Then
        f = f & " and INVOS_INFO_OSType_ID='" & fltr.txtINVOS_INFO_OSType.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_InLineNum_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_InLineNum>=" & val(fltr.txtINVOS_INFO_InLineNum_GE.Text)
      End If
      If fltr.lblINVOS_CODE_VisibleCode.Value = vbChecked Then
        f = f & " and INVOS_CODE_VisibleCode like '%" & fltr.txtINVOS_CODE_VisibleCode.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_SrokPI_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokPI<=" & val(fltr.txtINVOS_INFO_SrokPI_LE.Text)
      End If
      If fltr.lblINVOS_PLACE_Uprav.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Uprav_ID='" & fltr.txtINVOS_PLACE_Uprav.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_INVNum.Value = vbChecked Then
        f = f & " and INVOS_INFO_INVNum like '%" & fltr.txtINVOS_INFO_INVNum.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_SrokFI_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokFI<=" & val(fltr.txtINVOS_INFO_SrokFI_LE.Text)
      End If
      If fltr.lblINVOS_INFO_ActivateDate_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_ActivateDate<=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_INFO_ActivateDate_LE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_INFO_ActivateDate_LE.Value), MakePGSQLDate(fltr.dtpINVOS_INFO_ActivateDate_LE.Value)))
      End If
      If fltr.lblINVOS_INFO_SrokFI_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokFI>=" & val(fltr.txtINVOS_INFO_SrokFI_GE.Text)
      End If
      If fltr.lblINVOS_INFO_TheCost_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_TheCost>=" & val(fltr.txtINVOS_INFO_TheCost_GE.Text)
      End If
      If fltr.lblINVOS_PLACE_TheHouse.Value = vbChecked Then
        f = f & " and INVOS_PLACE_TheHouse_ID='" & fltr.txtINVOS_PLACE_TheHouse.Tag & "'"
      End If
      If fltr.lblINVOS_SROK_RecalcDate_LE.Value = vbChecked Then
        f = f & " and INVOS_SROK_RecalcDate<=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_SROK_RecalcDate_LE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_SROK_RecalcDate_LE.Value), MakePGSQLDate(fltr.dtpINVOS_SROK_RecalcDate_LE.Value)))
      End If
      If fltr.lblINVOS_INFO_CardNum.Value = vbChecked Then
        f = f & " and INVOS_INFO_CardNum like '%" & fltr.txtINVOS_INFO_CardNum.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_WorkPlaceNum.Value = vbChecked Then
        f = f & " and INVOS_PLACE_WorkPlaceNum like '%" & fltr.txtINVOS_PLACE_WorkPlaceNum.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_Info.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Info like '%" & fltr.txtINVOS_PLACE_Info.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_Flow.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Flow like '%" & fltr.txtINVOS_PLACE_Flow.Text & "%'"
      End If
    jfmnuINV_OS_1.jv.Filter.Add "AUTOINVOS_INFO", f
    End If
    Unload fltr
    usedefault = False
End Sub
Private Sub jfmnuINV_OS_1_OnClearFilter()
   jfmnuINV_OS_1.jv.Filter.Add "AUTOINVOS_INFO", " INTSANCEStatusID='{8E6E78D2-82AA-4913-B08C-1230A8C8B4A9}'"
End Sub
Private Sub jfmnuINV_OS_1_OnAdd(usedefaut As Boolean, Refesh As Boolean)
  Dim objGui  As Object
  Dim o As Object
  Dim id As String
  id = CreateGUID2
  Manager.NewInstance id, "INV_OS", "Карточка основного средства" & Now, Site
  Set o = Manager.GetInstanceObject(id)
  If IsDocDenied(o) Then
    MsgBox "Не разрешен доступ к документам такого типа"
    Exit Sub
  End If

  Dim g  As Object
  Set g = Manager.GetInstanceGUI(o.id)
  If Not g Is Nothing Then
    g.Show GetDocumentMode(o), o, False
  End If
  usedefaut = False
  Refesh = False
End Sub


Private Sub mnuINV_OS_2_Click()
    Dim journal As Object
    On Error Resume Next
    If jfmnuINV_OS_2 Is Nothing Then
      Set jfmnuINV_OS_2 = New frmJournalShow
      Set journal = Manager.GetInstanceObject("{5CE3CC0B-5224-4145-B43F-6A29CC390C17}")
      Manager.LockInstanceObject journal.id
      Set jfmnuINV_OS_2.jv.journal = journal
      AddColors jfmnuINV_OS_2.jv
      jfmnuINV_OS_2.jv.OpenModal = False
      jfmnuINV_OS_2.Caption = "Карточка основного средства :В лизинге"
      Me.MousePointer = vbHourglass
      DoEvents
      Dim f As String
    f = "1=1"
   f = " INTSANCEStatusID='{72195EFD-2052-4539-AB55-1D7E6B3AA767}'"
    jfmnuINV_OS_2.jv.Filter.Add "AUTOINVOS_INFO", f
    Dim fltr As frmINV_OS
    Set fltr = New frmINV_OS
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
      If fltr.lblINVOS_INFO_ActivateDate_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_ActivateDate>=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_INFO_ActivateDate_GE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_INFO_ActivateDate_GE.Value), MakePGSQLDate(fltr.dtpINVOS_INFO_ActivateDate_GE.Value)))
      End If
      If fltr.lblINVOS_INFO_IsMaterial.Value = vbChecked Then
        f = f & " and INVOS_INFO_IsMaterial='" & fltr.cmbINVOS_INFO_IsMaterial.Text & "'"
      End If
      If fltr.lblINVOS_INFO_ShortName.Value = vbChecked Then
        f = f & " and INVOS_INFO_ShortName like '%" & fltr.txtINVOS_INFO_ShortName.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_INVNum.Value = vbChecked Then
        f = f & " and INVOS_INFO_INVNum like '%" & fltr.txtINVOS_INFO_INVNum.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_Uprav.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Uprav_ID='" & fltr.txtINVOS_PLACE_Uprav.Tag & "'"
      End If
      If fltr.lblINVOS_PLACE_TheOwner.Value = vbChecked Then
        f = f & " and INVOS_PLACE_TheOwner_ID='" & fltr.txtINVOS_PLACE_TheOwner.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_SrokOI_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokOI>=" & val(fltr.txtINVOS_INFO_SrokOI_GE.Text)
      End If
      If fltr.lblINVOS_PLACE_ComplNumber.Value = vbChecked Then
        f = f & " and INVOS_PLACE_ComplNumber like '%" & fltr.txtINVOS_PLACE_ComplNumber.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_TheCost_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_TheCost>=" & val(fltr.txtINVOS_INFO_TheCost_GE.Text)
      End If
      If fltr.lblINVOS_INFO_TheOrg.Value = vbChecked Then
        f = f & " and INVOS_INFO_TheOrg_ID='" & fltr.txtINVOS_INFO_TheOrg.Tag & "'"
      End If
      If fltr.lblINVOS_SROK_RecalcDate_LE.Value = vbChecked Then
        f = f & " and INVOS_SROK_RecalcDate<=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_SROK_RecalcDate_LE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_SROK_RecalcDate_LE.Value), MakePGSQLDate(fltr.dtpINVOS_SROK_RecalcDate_LE.Value)))
      End If
      If fltr.lblINVOS_INFO_SrokFI_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokFI>=" & val(fltr.txtINVOS_INFO_SrokFI_GE.Text)
      End If
      If fltr.lblINVOS_INFO_CardNum.Value = vbChecked Then
        f = f & " and INVOS_INFO_CardNum like '%" & fltr.txtINVOS_INFO_CardNum.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_TheCost_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_TheCost<=" & val(fltr.txtINVOS_INFO_TheCost_LE.Text)
      End If
      If fltr.lblINVOS_INFO_Info.Value = vbChecked Then
        f = f & " and INVOS_INFO_Info like '%" & fltr.txtINVOS_INFO_Info.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_TheHouse.Value = vbChecked Then
        f = f & " and INVOS_PLACE_TheHouse_ID='" & fltr.txtINVOS_PLACE_TheHouse.Tag & "'"
      End If
      If fltr.lblINVOS_PLACE_MatOtv.Value = vbChecked Then
        f = f & " and INVOS_PLACE_MatOtv_ID='" & fltr.txtINVOS_PLACE_MatOtv.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_SrokPI_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokPI>=" & val(fltr.txtINVOS_INFO_SrokPI_GE.Text)
      End If
      If fltr.lblINVOS_INFO_SrokOI_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokOI<=" & val(fltr.txtINVOS_INFO_SrokOI_LE.Text)
      End If
      If fltr.lblINVOS_INFO_TechFilePath.Value = vbChecked Then
        f = f & " and INVOS_INFO_TechFilePath like '%" & fltr.txtINVOS_INFO_TechFilePath.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_WorkPlaceNum.Value = vbChecked Then
        f = f & " and INVOS_PLACE_WorkPlaceNum like '%" & fltr.txtINVOS_PLACE_WorkPlaceNum.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_ActivateDate_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_ActivateDate<=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_INFO_ActivateDate_LE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_INFO_ActivateDate_LE.Value), MakePGSQLDate(fltr.dtpINVOS_INFO_ActivateDate_LE.Value)))
      End If
      If fltr.lblINVOS_CODE_VisibleCode.Value = vbChecked Then
        f = f & " and INVOS_CODE_VisibleCode like '%" & fltr.txtINVOS_CODE_VisibleCode.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_SrokFI_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokFI<=" & val(fltr.txtINVOS_INFO_SrokFI_LE.Text)
      End If
      If fltr.lblINVOS_PLACE_DIrection.Value = vbChecked Then
        f = f & " and INVOS_PLACE_DIrection_ID='" & fltr.txtINVOS_PLACE_DIrection.Tag & "'"
      End If
      If fltr.lblINVOS_PLACE_Info.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Info like '%" & fltr.txtINVOS_PLACE_Info.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_Room.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Room like '%" & fltr.txtINVOS_PLACE_Room.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_Name.Value = vbChecked Then
        f = f & " and INVOS_INFO_Name like '%" & fltr.txtINVOS_INFO_Name.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_InLineNum_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_InLineNum<=" & val(fltr.txtINVOS_INFO_InLineNum_LE.Text)
      End If
      If fltr.lblINVOS_INFO_OSType.Value = vbChecked Then
        f = f & " and INVOS_INFO_OSType_ID='" & fltr.txtINVOS_INFO_OSType.Tag & "'"
      End If
      If fltr.lblINVOS_SROK_RecalcDate_GE.Value = vbChecked Then
        f = f & " and INVOS_SROK_RecalcDate>=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_SROK_RecalcDate_GE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_SROK_RecalcDate_GE.Value), MakePGSQLDate(fltr.dtpINVOS_SROK_RecalcDate_GE.Value)))
      End If
      If fltr.lblINVOS_PLACE_Flow.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Flow like '%" & fltr.txtINVOS_PLACE_Flow.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_InLineNum_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_InLineNum>=" & val(fltr.txtINVOS_INFO_InLineNum_GE.Text)
      End If
      If fltr.lblINVOS_PLACE_Otdel.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Otdel_ID='" & fltr.txtINVOS_PLACE_Otdel.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_SrokPI_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokPI<=" & val(fltr.txtINVOS_INFO_SrokPI_LE.Text)
      End If
    jfmnuINV_OS_2.jv.Filter.Add "AUTOINVOS_INFO", f
    End If
      jfmnuINV_OS_2.jv.Refresh
      Me.MousePointer = vbNormal
    End If
    jfmnuINV_OS_2.Show
    jfmnuINV_OS_2.WindowState = 0
    jfmnuINV_OS_2.ZOrder 0
End Sub
Private Sub jfmnuINV_OS_2_OnFilter(usedefault As Boolean)
    Dim fltr As frmINV_OS
    Dim f As String
    Set fltr = New frmINV_OS
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
    f = "1=1"
   f = " INTSANCEStatusID='{72195EFD-2052-4539-AB55-1D7E6B3AA767}'"
      If fltr.lblINVOS_INFO_ActivateDate_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_ActivateDate>=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_INFO_ActivateDate_GE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_INFO_ActivateDate_GE.Value), MakePGSQLDate(fltr.dtpINVOS_INFO_ActivateDate_GE.Value)))
      End If
      If fltr.lblINVOS_INFO_IsMaterial.Value = vbChecked Then
        f = f & " and INVOS_INFO_IsMaterial='" & fltr.cmbINVOS_INFO_IsMaterial.Text & "'"
      End If
      If fltr.lblINVOS_INFO_ShortName.Value = vbChecked Then
        f = f & " and INVOS_INFO_ShortName like '%" & fltr.txtINVOS_INFO_ShortName.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_INVNum.Value = vbChecked Then
        f = f & " and INVOS_INFO_INVNum like '%" & fltr.txtINVOS_INFO_INVNum.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_Uprav.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Uprav_ID='" & fltr.txtINVOS_PLACE_Uprav.Tag & "'"
      End If
      If fltr.lblINVOS_PLACE_TheOwner.Value = vbChecked Then
        f = f & " and INVOS_PLACE_TheOwner_ID='" & fltr.txtINVOS_PLACE_TheOwner.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_SrokOI_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokOI>=" & val(fltr.txtINVOS_INFO_SrokOI_GE.Text)
      End If
      If fltr.lblINVOS_PLACE_ComplNumber.Value = vbChecked Then
        f = f & " and INVOS_PLACE_ComplNumber like '%" & fltr.txtINVOS_PLACE_ComplNumber.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_TheCost_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_TheCost>=" & val(fltr.txtINVOS_INFO_TheCost_GE.Text)
      End If
      If fltr.lblINVOS_INFO_TheOrg.Value = vbChecked Then
        f = f & " and INVOS_INFO_TheOrg_ID='" & fltr.txtINVOS_INFO_TheOrg.Tag & "'"
      End If
      If fltr.lblINVOS_SROK_RecalcDate_LE.Value = vbChecked Then
        f = f & " and INVOS_SROK_RecalcDate<=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_SROK_RecalcDate_LE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_SROK_RecalcDate_LE.Value), MakePGSQLDate(fltr.dtpINVOS_SROK_RecalcDate_LE.Value)))
      End If
      If fltr.lblINVOS_INFO_SrokFI_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokFI>=" & val(fltr.txtINVOS_INFO_SrokFI_GE.Text)
      End If
      If fltr.lblINVOS_INFO_CardNum.Value = vbChecked Then
        f = f & " and INVOS_INFO_CardNum like '%" & fltr.txtINVOS_INFO_CardNum.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_TheCost_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_TheCost<=" & val(fltr.txtINVOS_INFO_TheCost_LE.Text)
      End If
      If fltr.lblINVOS_INFO_Info.Value = vbChecked Then
        f = f & " and INVOS_INFO_Info like '%" & fltr.txtINVOS_INFO_Info.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_TheHouse.Value = vbChecked Then
        f = f & " and INVOS_PLACE_TheHouse_ID='" & fltr.txtINVOS_PLACE_TheHouse.Tag & "'"
      End If
      If fltr.lblINVOS_PLACE_MatOtv.Value = vbChecked Then
        f = f & " and INVOS_PLACE_MatOtv_ID='" & fltr.txtINVOS_PLACE_MatOtv.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_SrokPI_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokPI>=" & val(fltr.txtINVOS_INFO_SrokPI_GE.Text)
      End If
      If fltr.lblINVOS_INFO_SrokOI_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokOI<=" & val(fltr.txtINVOS_INFO_SrokOI_LE.Text)
      End If
      If fltr.lblINVOS_INFO_TechFilePath.Value = vbChecked Then
        f = f & " and INVOS_INFO_TechFilePath like '%" & fltr.txtINVOS_INFO_TechFilePath.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_WorkPlaceNum.Value = vbChecked Then
        f = f & " and INVOS_PLACE_WorkPlaceNum like '%" & fltr.txtINVOS_PLACE_WorkPlaceNum.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_ActivateDate_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_ActivateDate<=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_INFO_ActivateDate_LE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_INFO_ActivateDate_LE.Value), MakePGSQLDate(fltr.dtpINVOS_INFO_ActivateDate_LE.Value)))
      End If
      If fltr.lblINVOS_CODE_VisibleCode.Value = vbChecked Then
        f = f & " and INVOS_CODE_VisibleCode like '%" & fltr.txtINVOS_CODE_VisibleCode.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_SrokFI_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokFI<=" & val(fltr.txtINVOS_INFO_SrokFI_LE.Text)
      End If
      If fltr.lblINVOS_PLACE_DIrection.Value = vbChecked Then
        f = f & " and INVOS_PLACE_DIrection_ID='" & fltr.txtINVOS_PLACE_DIrection.Tag & "'"
      End If
      If fltr.lblINVOS_PLACE_Info.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Info like '%" & fltr.txtINVOS_PLACE_Info.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_Room.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Room like '%" & fltr.txtINVOS_PLACE_Room.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_Name.Value = vbChecked Then
        f = f & " and INVOS_INFO_Name like '%" & fltr.txtINVOS_INFO_Name.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_InLineNum_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_InLineNum<=" & val(fltr.txtINVOS_INFO_InLineNum_LE.Text)
      End If
      If fltr.lblINVOS_INFO_OSType.Value = vbChecked Then
        f = f & " and INVOS_INFO_OSType_ID='" & fltr.txtINVOS_INFO_OSType.Tag & "'"
      End If
      If fltr.lblINVOS_SROK_RecalcDate_GE.Value = vbChecked Then
        f = f & " and INVOS_SROK_RecalcDate>=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_SROK_RecalcDate_GE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_SROK_RecalcDate_GE.Value), MakePGSQLDate(fltr.dtpINVOS_SROK_RecalcDate_GE.Value)))
      End If
      If fltr.lblINVOS_PLACE_Flow.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Flow like '%" & fltr.txtINVOS_PLACE_Flow.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_InLineNum_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_InLineNum>=" & val(fltr.txtINVOS_INFO_InLineNum_GE.Text)
      End If
      If fltr.lblINVOS_PLACE_Otdel.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Otdel_ID='" & fltr.txtINVOS_PLACE_Otdel.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_SrokPI_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokPI<=" & val(fltr.txtINVOS_INFO_SrokPI_LE.Text)
      End If
    jfmnuINV_OS_2.jv.Filter.Add "AUTOINVOS_INFO", f
    End If
    Unload fltr
    usedefault = False
End Sub
Private Sub jfmnuINV_OS_2_OnClearFilter()
   jfmnuINV_OS_2.jv.Filter.Add "AUTOINVOS_INFO", " INTSANCEStatusID='{72195EFD-2052-4539-AB55-1D7E6B3AA767}'"
End Sub
Private Sub jfmnuINV_OS_2_OnAdd(usedefaut As Boolean, Refesh As Boolean)
  Dim objGui  As Object
  Dim o As Object
  Dim id As String
  id = CreateGUID2
  Manager.NewInstance id, "INV_OS", "Карточка основного средства" & Now, Site
  Set o = Manager.GetInstanceObject(id)
  If IsDocDenied(o) Then
    MsgBox "Не разрешен доступ к документам такого типа"
    Exit Sub
  End If

  Dim g  As Object
  Set g = Manager.GetInstanceGUI(o.id)
  If Not g Is Nothing Then
    g.Show GetDocumentMode(o), o, False
  End If
  usedefaut = False
  Refesh = False
End Sub


Private Sub mnuINV_OS_3_Click()
    Dim journal As Object
    On Error Resume Next
    If jfmnuINV_OS_3 Is Nothing Then
      Set jfmnuINV_OS_3 = New frmJournalShow
      Set journal = Manager.GetInstanceObject("{5CE3CC0B-5224-4145-B43F-6A29CC390C17}")
      Manager.LockInstanceObject journal.id
      Set jfmnuINV_OS_3.jv.journal = journal
      AddColors jfmnuINV_OS_3.jv
      jfmnuINV_OS_3.jv.OpenModal = False
      jfmnuINV_OS_3.Caption = "Карточка основного средства :Оформляется"
      Me.MousePointer = vbHourglass
      DoEvents
      Dim f As String
    f = "1=1"
   f = " INTSANCEStatusID='{179CB53A-CBE7-46B4-9905-22E35FAAE801}'"
    jfmnuINV_OS_3.jv.Filter.Add "AUTOINVOS_INFO", f
    Dim fltr As frmINV_OS
    Set fltr = New frmINV_OS
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
      If fltr.lblINVOS_INFO_CardNum.Value = vbChecked Then
        f = f & " and INVOS_INFO_CardNum like '%" & fltr.txtINVOS_INFO_CardNum.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_Uprav.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Uprav_ID='" & fltr.txtINVOS_PLACE_Uprav.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_Info.Value = vbChecked Then
        f = f & " and INVOS_INFO_Info like '%" & fltr.txtINVOS_INFO_Info.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_Flow.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Flow like '%" & fltr.txtINVOS_PLACE_Flow.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_OSType.Value = vbChecked Then
        f = f & " and INVOS_INFO_OSType_ID='" & fltr.txtINVOS_INFO_OSType.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_ActivateDate_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_ActivateDate<=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_INFO_ActivateDate_LE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_INFO_ActivateDate_LE.Value), MakePGSQLDate(fltr.dtpINVOS_INFO_ActivateDate_LE.Value)))
      End If
      If fltr.lblINVOS_SROK_RecalcDate_LE.Value = vbChecked Then
        f = f & " and INVOS_SROK_RecalcDate<=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_SROK_RecalcDate_LE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_SROK_RecalcDate_LE.Value), MakePGSQLDate(fltr.dtpINVOS_SROK_RecalcDate_LE.Value)))
      End If
      If fltr.lblINVOS_INFO_SrokOI_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokOI>=" & val(fltr.txtINVOS_INFO_SrokOI_GE.Text)
      End If
      If fltr.lblINVOS_INFO_TheCost_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_TheCost<=" & val(fltr.txtINVOS_INFO_TheCost_LE.Text)
      End If
      If fltr.lblINVOS_PLACE_Otdel.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Otdel_ID='" & fltr.txtINVOS_PLACE_Otdel.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_InLineNum_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_InLineNum<=" & val(fltr.txtINVOS_INFO_InLineNum_LE.Text)
      End If
      If fltr.lblINVOS_INFO_ActivateDate_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_ActivateDate>=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_INFO_ActivateDate_GE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_INFO_ActivateDate_GE.Value), MakePGSQLDate(fltr.dtpINVOS_INFO_ActivateDate_GE.Value)))
      End If
      If fltr.lblINVOS_PLACE_ComplNumber.Value = vbChecked Then
        f = f & " and INVOS_PLACE_ComplNumber like '%" & fltr.txtINVOS_PLACE_ComplNumber.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_Name.Value = vbChecked Then
        f = f & " and INVOS_INFO_Name like '%" & fltr.txtINVOS_INFO_Name.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_INVNum.Value = vbChecked Then
        f = f & " and INVOS_INFO_INVNum like '%" & fltr.txtINVOS_INFO_INVNum.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_SrokPI_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokPI<=" & val(fltr.txtINVOS_INFO_SrokPI_LE.Text)
      End If
      If fltr.lblINVOS_INFO_SrokFI_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokFI>=" & val(fltr.txtINVOS_INFO_SrokFI_GE.Text)
      End If
      If fltr.lblINVOS_INFO_TheOrg.Value = vbChecked Then
        f = f & " and INVOS_INFO_TheOrg_ID='" & fltr.txtINVOS_INFO_TheOrg.Tag & "'"
      End If
      If fltr.lblINVOS_PLACE_WorkPlaceNum.Value = vbChecked Then
        f = f & " and INVOS_PLACE_WorkPlaceNum like '%" & fltr.txtINVOS_PLACE_WorkPlaceNum.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_ShortName.Value = vbChecked Then
        f = f & " and INVOS_INFO_ShortName like '%" & fltr.txtINVOS_INFO_ShortName.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_TheCost_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_TheCost>=" & val(fltr.txtINVOS_INFO_TheCost_GE.Text)
      End If
      If fltr.lblINVOS_INFO_SrokOI_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokOI<=" & val(fltr.txtINVOS_INFO_SrokOI_LE.Text)
      End If
      If fltr.lblINVOS_INFO_InLineNum_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_InLineNum>=" & val(fltr.txtINVOS_INFO_InLineNum_GE.Text)
      End If
      If fltr.lblINVOS_SROK_RecalcDate_GE.Value = vbChecked Then
        f = f & " and INVOS_SROK_RecalcDate>=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_SROK_RecalcDate_GE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_SROK_RecalcDate_GE.Value), MakePGSQLDate(fltr.dtpINVOS_SROK_RecalcDate_GE.Value)))
      End If
      If fltr.lblINVOS_PLACE_TheHouse.Value = vbChecked Then
        f = f & " and INVOS_PLACE_TheHouse_ID='" & fltr.txtINVOS_PLACE_TheHouse.Tag & "'"
      End If
      If fltr.lblINVOS_PLACE_Info.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Info like '%" & fltr.txtINVOS_PLACE_Info.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_TechFilePath.Value = vbChecked Then
        f = f & " and INVOS_INFO_TechFilePath like '%" & fltr.txtINVOS_INFO_TechFilePath.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_MatOtv.Value = vbChecked Then
        f = f & " and INVOS_PLACE_MatOtv_ID='" & fltr.txtINVOS_PLACE_MatOtv.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_SrokPI_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokPI>=" & val(fltr.txtINVOS_INFO_SrokPI_GE.Text)
      End If
      If fltr.lblINVOS_INFO_IsMaterial.Value = vbChecked Then
        f = f & " and INVOS_INFO_IsMaterial='" & fltr.cmbINVOS_INFO_IsMaterial.Text & "'"
      End If
      If fltr.lblINVOS_CODE_VisibleCode.Value = vbChecked Then
        f = f & " and INVOS_CODE_VisibleCode like '%" & fltr.txtINVOS_CODE_VisibleCode.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_DIrection.Value = vbChecked Then
        f = f & " and INVOS_PLACE_DIrection_ID='" & fltr.txtINVOS_PLACE_DIrection.Tag & "'"
      End If
      If fltr.lblINVOS_PLACE_TheOwner.Value = vbChecked Then
        f = f & " and INVOS_PLACE_TheOwner_ID='" & fltr.txtINVOS_PLACE_TheOwner.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_SrokFI_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokFI<=" & val(fltr.txtINVOS_INFO_SrokFI_LE.Text)
      End If
      If fltr.lblINVOS_PLACE_Room.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Room like '%" & fltr.txtINVOS_PLACE_Room.Text & "%'"
      End If
    jfmnuINV_OS_3.jv.Filter.Add "AUTOINVOS_INFO", f
    End If
      jfmnuINV_OS_3.jv.Refresh
      Me.MousePointer = vbNormal
    End If
    jfmnuINV_OS_3.Show
    jfmnuINV_OS_3.WindowState = 0
    jfmnuINV_OS_3.ZOrder 0
End Sub
Private Sub jfmnuINV_OS_3_OnFilter(usedefault As Boolean)
    Dim fltr As frmINV_OS
    Dim f As String
    Set fltr = New frmINV_OS
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
    f = "1=1"
   f = " INTSANCEStatusID='{179CB53A-CBE7-46B4-9905-22E35FAAE801}'"
      If fltr.lblINVOS_INFO_CardNum.Value = vbChecked Then
        f = f & " and INVOS_INFO_CardNum like '%" & fltr.txtINVOS_INFO_CardNum.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_Uprav.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Uprav_ID='" & fltr.txtINVOS_PLACE_Uprav.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_Info.Value = vbChecked Then
        f = f & " and INVOS_INFO_Info like '%" & fltr.txtINVOS_INFO_Info.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_Flow.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Flow like '%" & fltr.txtINVOS_PLACE_Flow.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_OSType.Value = vbChecked Then
        f = f & " and INVOS_INFO_OSType_ID='" & fltr.txtINVOS_INFO_OSType.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_ActivateDate_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_ActivateDate<=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_INFO_ActivateDate_LE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_INFO_ActivateDate_LE.Value), MakePGSQLDate(fltr.dtpINVOS_INFO_ActivateDate_LE.Value)))
      End If
      If fltr.lblINVOS_SROK_RecalcDate_LE.Value = vbChecked Then
        f = f & " and INVOS_SROK_RecalcDate<=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_SROK_RecalcDate_LE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_SROK_RecalcDate_LE.Value), MakePGSQLDate(fltr.dtpINVOS_SROK_RecalcDate_LE.Value)))
      End If
      If fltr.lblINVOS_INFO_SrokOI_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokOI>=" & val(fltr.txtINVOS_INFO_SrokOI_GE.Text)
      End If
      If fltr.lblINVOS_INFO_TheCost_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_TheCost<=" & val(fltr.txtINVOS_INFO_TheCost_LE.Text)
      End If
      If fltr.lblINVOS_PLACE_Otdel.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Otdel_ID='" & fltr.txtINVOS_PLACE_Otdel.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_InLineNum_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_InLineNum<=" & val(fltr.txtINVOS_INFO_InLineNum_LE.Text)
      End If
      If fltr.lblINVOS_INFO_ActivateDate_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_ActivateDate>=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_INFO_ActivateDate_GE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_INFO_ActivateDate_GE.Value), MakePGSQLDate(fltr.dtpINVOS_INFO_ActivateDate_GE.Value)))
      End If
      If fltr.lblINVOS_PLACE_ComplNumber.Value = vbChecked Then
        f = f & " and INVOS_PLACE_ComplNumber like '%" & fltr.txtINVOS_PLACE_ComplNumber.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_Name.Value = vbChecked Then
        f = f & " and INVOS_INFO_Name like '%" & fltr.txtINVOS_INFO_Name.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_INVNum.Value = vbChecked Then
        f = f & " and INVOS_INFO_INVNum like '%" & fltr.txtINVOS_INFO_INVNum.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_SrokPI_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokPI<=" & val(fltr.txtINVOS_INFO_SrokPI_LE.Text)
      End If
      If fltr.lblINVOS_INFO_SrokFI_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokFI>=" & val(fltr.txtINVOS_INFO_SrokFI_GE.Text)
      End If
      If fltr.lblINVOS_INFO_TheOrg.Value = vbChecked Then
        f = f & " and INVOS_INFO_TheOrg_ID='" & fltr.txtINVOS_INFO_TheOrg.Tag & "'"
      End If
      If fltr.lblINVOS_PLACE_WorkPlaceNum.Value = vbChecked Then
        f = f & " and INVOS_PLACE_WorkPlaceNum like '%" & fltr.txtINVOS_PLACE_WorkPlaceNum.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_ShortName.Value = vbChecked Then
        f = f & " and INVOS_INFO_ShortName like '%" & fltr.txtINVOS_INFO_ShortName.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_TheCost_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_TheCost>=" & val(fltr.txtINVOS_INFO_TheCost_GE.Text)
      End If
      If fltr.lblINVOS_INFO_SrokOI_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokOI<=" & val(fltr.txtINVOS_INFO_SrokOI_LE.Text)
      End If
      If fltr.lblINVOS_INFO_InLineNum_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_InLineNum>=" & val(fltr.txtINVOS_INFO_InLineNum_GE.Text)
      End If
      If fltr.lblINVOS_SROK_RecalcDate_GE.Value = vbChecked Then
        f = f & " and INVOS_SROK_RecalcDate>=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_SROK_RecalcDate_GE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_SROK_RecalcDate_GE.Value), MakePGSQLDate(fltr.dtpINVOS_SROK_RecalcDate_GE.Value)))
      End If
      If fltr.lblINVOS_PLACE_TheHouse.Value = vbChecked Then
        f = f & " and INVOS_PLACE_TheHouse_ID='" & fltr.txtINVOS_PLACE_TheHouse.Tag & "'"
      End If
      If fltr.lblINVOS_PLACE_Info.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Info like '%" & fltr.txtINVOS_PLACE_Info.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_TechFilePath.Value = vbChecked Then
        f = f & " and INVOS_INFO_TechFilePath like '%" & fltr.txtINVOS_INFO_TechFilePath.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_MatOtv.Value = vbChecked Then
        f = f & " and INVOS_PLACE_MatOtv_ID='" & fltr.txtINVOS_PLACE_MatOtv.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_SrokPI_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokPI>=" & val(fltr.txtINVOS_INFO_SrokPI_GE.Text)
      End If
      If fltr.lblINVOS_INFO_IsMaterial.Value = vbChecked Then
        f = f & " and INVOS_INFO_IsMaterial='" & fltr.cmbINVOS_INFO_IsMaterial.Text & "'"
      End If
      If fltr.lblINVOS_CODE_VisibleCode.Value = vbChecked Then
        f = f & " and INVOS_CODE_VisibleCode like '%" & fltr.txtINVOS_CODE_VisibleCode.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_DIrection.Value = vbChecked Then
        f = f & " and INVOS_PLACE_DIrection_ID='" & fltr.txtINVOS_PLACE_DIrection.Tag & "'"
      End If
      If fltr.lblINVOS_PLACE_TheOwner.Value = vbChecked Then
        f = f & " and INVOS_PLACE_TheOwner_ID='" & fltr.txtINVOS_PLACE_TheOwner.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_SrokFI_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokFI<=" & val(fltr.txtINVOS_INFO_SrokFI_LE.Text)
      End If
      If fltr.lblINVOS_PLACE_Room.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Room like '%" & fltr.txtINVOS_PLACE_Room.Text & "%'"
      End If
    jfmnuINV_OS_3.jv.Filter.Add "AUTOINVOS_INFO", f
    End If
    Unload fltr
    usedefault = False
End Sub
Private Sub jfmnuINV_OS_3_OnClearFilter()
   jfmnuINV_OS_3.jv.Filter.Add "AUTOINVOS_INFO", " INTSANCEStatusID='{179CB53A-CBE7-46B4-9905-22E35FAAE801}'"
End Sub
Private Sub jfmnuINV_OS_3_OnAdd(usedefaut As Boolean, Refesh As Boolean)
  Dim objGui  As Object
  Dim o As Object
  Dim id As String
  id = CreateGUID2
  Manager.NewInstance id, "INV_OS", "Карточка основного средства" & Now, Site
  Set o = Manager.GetInstanceObject(id)
  If IsDocDenied(o) Then
    MsgBox "Не разрешен доступ к документам такого типа"
    Exit Sub
  End If

  Dim g  As Object
  Set g = Manager.GetInstanceGUI(o.id)
  If Not g Is Nothing Then
    g.Show GetDocumentMode(o), o, False
  End If
  usedefaut = False
  Refesh = False
End Sub


Private Sub mnuINV_OS_4_Click()
    Dim journal As Object
    On Error Resume Next
    If jfmnuINV_OS_4 Is Nothing Then
      Set jfmnuINV_OS_4 = New frmJournalShow
      Set journal = Manager.GetInstanceObject("{5CE3CC0B-5224-4145-B43F-6A29CC390C17}")
      Manager.LockInstanceObject journal.id
      Set jfmnuINV_OS_4.jv.journal = journal
      AddColors jfmnuINV_OS_4.jv
      jfmnuINV_OS_4.jv.OpenModal = False
      jfmnuINV_OS_4.Caption = "Карточка основного средства :В аренде"
      Me.MousePointer = vbHourglass
      DoEvents
      Dim f As String
    f = "1=1"
   f = " INTSANCEStatusID='{2AA78799-2880-4541-99E0-3C8750AC33E6}'"
    jfmnuINV_OS_4.jv.Filter.Add "AUTOINVOS_INFO", f
    Dim fltr As frmINV_OS
    Set fltr = New frmINV_OS
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
      If fltr.lblINVOS_PLACE_TheOwner.Value = vbChecked Then
        f = f & " and INVOS_PLACE_TheOwner_ID='" & fltr.txtINVOS_PLACE_TheOwner.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_InLineNum_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_InLineNum<=" & val(fltr.txtINVOS_INFO_InLineNum_LE.Text)
      End If
      If fltr.lblINVOS_INFO_Name.Value = vbChecked Then
        f = f & " and INVOS_INFO_Name like '%" & fltr.txtINVOS_INFO_Name.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_CardNum.Value = vbChecked Then
        f = f & " and INVOS_INFO_CardNum like '%" & fltr.txtINVOS_INFO_CardNum.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_Info.Value = vbChecked Then
        f = f & " and INVOS_INFO_Info like '%" & fltr.txtINVOS_INFO_Info.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_ActivateDate_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_ActivateDate>=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_INFO_ActivateDate_GE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_INFO_ActivateDate_GE.Value), MakePGSQLDate(fltr.dtpINVOS_INFO_ActivateDate_GE.Value)))
      End If
      If fltr.lblINVOS_INFO_SrokOI_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokOI>=" & val(fltr.txtINVOS_INFO_SrokOI_GE.Text)
      End If
      If fltr.lblINVOS_INFO_ActivateDate_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_ActivateDate<=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_INFO_ActivateDate_LE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_INFO_ActivateDate_LE.Value), MakePGSQLDate(fltr.dtpINVOS_INFO_ActivateDate_LE.Value)))
      End If
      If fltr.lblINVOS_INFO_ShortName.Value = vbChecked Then
        f = f & " and INVOS_INFO_ShortName like '%" & fltr.txtINVOS_INFO_ShortName.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_SrokFI_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokFI<=" & val(fltr.txtINVOS_INFO_SrokFI_LE.Text)
      End If
      If fltr.lblINVOS_SROK_RecalcDate_GE.Value = vbChecked Then
        f = f & " and INVOS_SROK_RecalcDate>=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_SROK_RecalcDate_GE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_SROK_RecalcDate_GE.Value), MakePGSQLDate(fltr.dtpINVOS_SROK_RecalcDate_GE.Value)))
      End If
      If fltr.lblINVOS_INFO_SrokFI_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokFI>=" & val(fltr.txtINVOS_INFO_SrokFI_GE.Text)
      End If
      If fltr.lblINVOS_PLACE_DIrection.Value = vbChecked Then
        f = f & " and INVOS_PLACE_DIrection_ID='" & fltr.txtINVOS_PLACE_DIrection.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_IsMaterial.Value = vbChecked Then
        f = f & " and INVOS_INFO_IsMaterial='" & fltr.cmbINVOS_INFO_IsMaterial.Text & "'"
      End If
      If fltr.lblINVOS_PLACE_Flow.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Flow like '%" & fltr.txtINVOS_PLACE_Flow.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_TheOrg.Value = vbChecked Then
        f = f & " and INVOS_INFO_TheOrg_ID='" & fltr.txtINVOS_INFO_TheOrg.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_TechFilePath.Value = vbChecked Then
        f = f & " and INVOS_INFO_TechFilePath like '%" & fltr.txtINVOS_INFO_TechFilePath.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_Room.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Room like '%" & fltr.txtINVOS_PLACE_Room.Text & "%'"
      End If
      If fltr.lblINVOS_CODE_VisibleCode.Value = vbChecked Then
        f = f & " and INVOS_CODE_VisibleCode like '%" & fltr.txtINVOS_CODE_VisibleCode.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_SrokPI_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokPI<=" & val(fltr.txtINVOS_INFO_SrokPI_LE.Text)
      End If
      If fltr.lblINVOS_PLACE_TheHouse.Value = vbChecked Then
        f = f & " and INVOS_PLACE_TheHouse_ID='" & fltr.txtINVOS_PLACE_TheHouse.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_SrokPI_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokPI>=" & val(fltr.txtINVOS_INFO_SrokPI_GE.Text)
      End If
      If fltr.lblINVOS_PLACE_MatOtv.Value = vbChecked Then
        f = f & " and INVOS_PLACE_MatOtv_ID='" & fltr.txtINVOS_PLACE_MatOtv.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_InLineNum_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_InLineNum>=" & val(fltr.txtINVOS_INFO_InLineNum_GE.Text)
      End If
      If fltr.lblINVOS_PLACE_Uprav.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Uprav_ID='" & fltr.txtINVOS_PLACE_Uprav.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_TheCost_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_TheCost>=" & val(fltr.txtINVOS_INFO_TheCost_GE.Text)
      End If
      If fltr.lblINVOS_INFO_SrokOI_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokOI<=" & val(fltr.txtINVOS_INFO_SrokOI_LE.Text)
      End If
      If fltr.lblINVOS_INFO_TheCost_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_TheCost<=" & val(fltr.txtINVOS_INFO_TheCost_LE.Text)
      End If
      If fltr.lblINVOS_SROK_RecalcDate_LE.Value = vbChecked Then
        f = f & " and INVOS_SROK_RecalcDate<=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_SROK_RecalcDate_LE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_SROK_RecalcDate_LE.Value), MakePGSQLDate(fltr.dtpINVOS_SROK_RecalcDate_LE.Value)))
      End If
      If fltr.lblINVOS_PLACE_Info.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Info like '%" & fltr.txtINVOS_PLACE_Info.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_OSType.Value = vbChecked Then
        f = f & " and INVOS_INFO_OSType_ID='" & fltr.txtINVOS_INFO_OSType.Tag & "'"
      End If
      If fltr.lblINVOS_PLACE_WorkPlaceNum.Value = vbChecked Then
        f = f & " and INVOS_PLACE_WorkPlaceNum like '%" & fltr.txtINVOS_PLACE_WorkPlaceNum.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_Otdel.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Otdel_ID='" & fltr.txtINVOS_PLACE_Otdel.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_INVNum.Value = vbChecked Then
        f = f & " and INVOS_INFO_INVNum like '%" & fltr.txtINVOS_INFO_INVNum.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_ComplNumber.Value = vbChecked Then
        f = f & " and INVOS_PLACE_ComplNumber like '%" & fltr.txtINVOS_PLACE_ComplNumber.Text & "%'"
      End If
    jfmnuINV_OS_4.jv.Filter.Add "AUTOINVOS_INFO", f
    End If
      jfmnuINV_OS_4.jv.Refresh
      Me.MousePointer = vbNormal
    End If
    jfmnuINV_OS_4.Show
    jfmnuINV_OS_4.WindowState = 0
    jfmnuINV_OS_4.ZOrder 0
End Sub
Private Sub jfmnuINV_OS_4_OnFilter(usedefault As Boolean)
    Dim fltr As frmINV_OS
    Dim f As String
    Set fltr = New frmINV_OS
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
    f = "1=1"
   f = " INTSANCEStatusID='{2AA78799-2880-4541-99E0-3C8750AC33E6}'"
      If fltr.lblINVOS_PLACE_TheOwner.Value = vbChecked Then
        f = f & " and INVOS_PLACE_TheOwner_ID='" & fltr.txtINVOS_PLACE_TheOwner.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_InLineNum_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_InLineNum<=" & val(fltr.txtINVOS_INFO_InLineNum_LE.Text)
      End If
      If fltr.lblINVOS_INFO_Name.Value = vbChecked Then
        f = f & " and INVOS_INFO_Name like '%" & fltr.txtINVOS_INFO_Name.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_CardNum.Value = vbChecked Then
        f = f & " and INVOS_INFO_CardNum like '%" & fltr.txtINVOS_INFO_CardNum.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_Info.Value = vbChecked Then
        f = f & " and INVOS_INFO_Info like '%" & fltr.txtINVOS_INFO_Info.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_ActivateDate_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_ActivateDate>=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_INFO_ActivateDate_GE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_INFO_ActivateDate_GE.Value), MakePGSQLDate(fltr.dtpINVOS_INFO_ActivateDate_GE.Value)))
      End If
      If fltr.lblINVOS_INFO_SrokOI_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokOI>=" & val(fltr.txtINVOS_INFO_SrokOI_GE.Text)
      End If
      If fltr.lblINVOS_INFO_ActivateDate_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_ActivateDate<=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_INFO_ActivateDate_LE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_INFO_ActivateDate_LE.Value), MakePGSQLDate(fltr.dtpINVOS_INFO_ActivateDate_LE.Value)))
      End If
      If fltr.lblINVOS_INFO_ShortName.Value = vbChecked Then
        f = f & " and INVOS_INFO_ShortName like '%" & fltr.txtINVOS_INFO_ShortName.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_SrokFI_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokFI<=" & val(fltr.txtINVOS_INFO_SrokFI_LE.Text)
      End If
      If fltr.lblINVOS_SROK_RecalcDate_GE.Value = vbChecked Then
        f = f & " and INVOS_SROK_RecalcDate>=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_SROK_RecalcDate_GE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_SROK_RecalcDate_GE.Value), MakePGSQLDate(fltr.dtpINVOS_SROK_RecalcDate_GE.Value)))
      End If
      If fltr.lblINVOS_INFO_SrokFI_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokFI>=" & val(fltr.txtINVOS_INFO_SrokFI_GE.Text)
      End If
      If fltr.lblINVOS_PLACE_DIrection.Value = vbChecked Then
        f = f & " and INVOS_PLACE_DIrection_ID='" & fltr.txtINVOS_PLACE_DIrection.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_IsMaterial.Value = vbChecked Then
        f = f & " and INVOS_INFO_IsMaterial='" & fltr.cmbINVOS_INFO_IsMaterial.Text & "'"
      End If
      If fltr.lblINVOS_PLACE_Flow.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Flow like '%" & fltr.txtINVOS_PLACE_Flow.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_TheOrg.Value = vbChecked Then
        f = f & " and INVOS_INFO_TheOrg_ID='" & fltr.txtINVOS_INFO_TheOrg.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_TechFilePath.Value = vbChecked Then
        f = f & " and INVOS_INFO_TechFilePath like '%" & fltr.txtINVOS_INFO_TechFilePath.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_Room.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Room like '%" & fltr.txtINVOS_PLACE_Room.Text & "%'"
      End If
      If fltr.lblINVOS_CODE_VisibleCode.Value = vbChecked Then
        f = f & " and INVOS_CODE_VisibleCode like '%" & fltr.txtINVOS_CODE_VisibleCode.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_SrokPI_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokPI<=" & val(fltr.txtINVOS_INFO_SrokPI_LE.Text)
      End If
      If fltr.lblINVOS_PLACE_TheHouse.Value = vbChecked Then
        f = f & " and INVOS_PLACE_TheHouse_ID='" & fltr.txtINVOS_PLACE_TheHouse.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_SrokPI_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokPI>=" & val(fltr.txtINVOS_INFO_SrokPI_GE.Text)
      End If
      If fltr.lblINVOS_PLACE_MatOtv.Value = vbChecked Then
        f = f & " and INVOS_PLACE_MatOtv_ID='" & fltr.txtINVOS_PLACE_MatOtv.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_InLineNum_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_InLineNum>=" & val(fltr.txtINVOS_INFO_InLineNum_GE.Text)
      End If
      If fltr.lblINVOS_PLACE_Uprav.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Uprav_ID='" & fltr.txtINVOS_PLACE_Uprav.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_TheCost_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_TheCost>=" & val(fltr.txtINVOS_INFO_TheCost_GE.Text)
      End If
      If fltr.lblINVOS_INFO_SrokOI_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokOI<=" & val(fltr.txtINVOS_INFO_SrokOI_LE.Text)
      End If
      If fltr.lblINVOS_INFO_TheCost_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_TheCost<=" & val(fltr.txtINVOS_INFO_TheCost_LE.Text)
      End If
      If fltr.lblINVOS_SROK_RecalcDate_LE.Value = vbChecked Then
        f = f & " and INVOS_SROK_RecalcDate<=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_SROK_RecalcDate_LE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_SROK_RecalcDate_LE.Value), MakePGSQLDate(fltr.dtpINVOS_SROK_RecalcDate_LE.Value)))
      End If
      If fltr.lblINVOS_PLACE_Info.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Info like '%" & fltr.txtINVOS_PLACE_Info.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_OSType.Value = vbChecked Then
        f = f & " and INVOS_INFO_OSType_ID='" & fltr.txtINVOS_INFO_OSType.Tag & "'"
      End If
      If fltr.lblINVOS_PLACE_WorkPlaceNum.Value = vbChecked Then
        f = f & " and INVOS_PLACE_WorkPlaceNum like '%" & fltr.txtINVOS_PLACE_WorkPlaceNum.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_Otdel.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Otdel_ID='" & fltr.txtINVOS_PLACE_Otdel.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_INVNum.Value = vbChecked Then
        f = f & " and INVOS_INFO_INVNum like '%" & fltr.txtINVOS_INFO_INVNum.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_ComplNumber.Value = vbChecked Then
        f = f & " and INVOS_PLACE_ComplNumber like '%" & fltr.txtINVOS_PLACE_ComplNumber.Text & "%'"
      End If
    jfmnuINV_OS_4.jv.Filter.Add "AUTOINVOS_INFO", f
    End If
    Unload fltr
    usedefault = False
End Sub
Private Sub jfmnuINV_OS_4_OnClearFilter()
   jfmnuINV_OS_4.jv.Filter.Add "AUTOINVOS_INFO", " INTSANCEStatusID='{2AA78799-2880-4541-99E0-3C8750AC33E6}'"
End Sub
Private Sub jfmnuINV_OS_4_OnAdd(usedefaut As Boolean, Refesh As Boolean)
  Dim objGui  As Object
  Dim o As Object
  Dim id As String
  id = CreateGUID2
  Manager.NewInstance id, "INV_OS", "Карточка основного средства" & Now, Site
  Set o = Manager.GetInstanceObject(id)
  If IsDocDenied(o) Then
    MsgBox "Не разрешен доступ к документам такого типа"
    Exit Sub
  End If

  Dim g  As Object
  Set g = Manager.GetInstanceGUI(o.id)
  If Not g Is Nothing Then
    g.Show GetDocumentMode(o), o, False
  End If
  usedefaut = False
  Refesh = False
End Sub


Private Sub mnuINV_OS_5_Click()
    Dim journal As Object
    On Error Resume Next
    If jfmnuINV_OS_5 Is Nothing Then
      Set jfmnuINV_OS_5 = New frmJournalShow
      Set journal = Manager.GetInstanceObject("{5CE3CC0B-5224-4145-B43F-6A29CC390C17}")
      Manager.LockInstanceObject journal.id
      Set jfmnuINV_OS_5.jv.journal = journal
      AddColors jfmnuINV_OS_5.jv
      jfmnuINV_OS_5.jv.OpenModal = False
      jfmnuINV_OS_5.Caption = "Карточка основного средства :На модернизации"
      Me.MousePointer = vbHourglass
      DoEvents
      Dim f As String
    f = "1=1"
   f = " INTSANCEStatusID='{55270A15-FA1D-4121-860B-A1B697B40A40}'"
    jfmnuINV_OS_5.jv.Filter.Add "AUTOINVOS_INFO", f
    Dim fltr As frmINV_OS
    Set fltr = New frmINV_OS
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
      If fltr.lblINVOS_SROK_RecalcDate_LE.Value = vbChecked Then
        f = f & " and INVOS_SROK_RecalcDate<=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_SROK_RecalcDate_LE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_SROK_RecalcDate_LE.Value), MakePGSQLDate(fltr.dtpINVOS_SROK_RecalcDate_LE.Value)))
      End If
      If fltr.lblINVOS_INFO_IsMaterial.Value = vbChecked Then
        f = f & " and INVOS_INFO_IsMaterial='" & fltr.cmbINVOS_INFO_IsMaterial.Text & "'"
      End If
      If fltr.lblINVOS_PLACE_MatOtv.Value = vbChecked Then
        f = f & " and INVOS_PLACE_MatOtv_ID='" & fltr.txtINVOS_PLACE_MatOtv.Tag & "'"
      End If
      If fltr.lblINVOS_PLACE_Otdel.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Otdel_ID='" & fltr.txtINVOS_PLACE_Otdel.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_SrokFI_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokFI<=" & val(fltr.txtINVOS_INFO_SrokFI_LE.Text)
      End If
      If fltr.lblINVOS_INFO_Name.Value = vbChecked Then
        f = f & " and INVOS_INFO_Name like '%" & fltr.txtINVOS_INFO_Name.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_ActivateDate_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_ActivateDate<=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_INFO_ActivateDate_LE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_INFO_ActivateDate_LE.Value), MakePGSQLDate(fltr.dtpINVOS_INFO_ActivateDate_LE.Value)))
      End If
      If fltr.lblINVOS_INFO_TheOrg.Value = vbChecked Then
        f = f & " and INVOS_INFO_TheOrg_ID='" & fltr.txtINVOS_INFO_TheOrg.Tag & "'"
      End If
      If fltr.lblINVOS_PLACE_ComplNumber.Value = vbChecked Then
        f = f & " and INVOS_PLACE_ComplNumber like '%" & fltr.txtINVOS_PLACE_ComplNumber.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_ActivateDate_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_ActivateDate>=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_INFO_ActivateDate_GE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_INFO_ActivateDate_GE.Value), MakePGSQLDate(fltr.dtpINVOS_INFO_ActivateDate_GE.Value)))
      End If
      If fltr.lblINVOS_PLACE_TheHouse.Value = vbChecked Then
        f = f & " and INVOS_PLACE_TheHouse_ID='" & fltr.txtINVOS_PLACE_TheHouse.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_TheCost_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_TheCost<=" & val(fltr.txtINVOS_INFO_TheCost_LE.Text)
      End If
      If fltr.lblINVOS_INFO_SrokPI_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokPI>=" & val(fltr.txtINVOS_INFO_SrokPI_GE.Text)
      End If
      If fltr.lblINVOS_PLACE_Room.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Room like '%" & fltr.txtINVOS_PLACE_Room.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_SrokOI_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokOI>=" & val(fltr.txtINVOS_INFO_SrokOI_GE.Text)
      End If
      If fltr.lblINVOS_PLACE_TheOwner.Value = vbChecked Then
        f = f & " and INVOS_PLACE_TheOwner_ID='" & fltr.txtINVOS_PLACE_TheOwner.Tag & "'"
      End If
      If fltr.lblINVOS_PLACE_Info.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Info like '%" & fltr.txtINVOS_PLACE_Info.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_InLineNum_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_InLineNum<=" & val(fltr.txtINVOS_INFO_InLineNum_LE.Text)
      End If
      If fltr.lblINVOS_INFO_SrokFI_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokFI>=" & val(fltr.txtINVOS_INFO_SrokFI_GE.Text)
      End If
      If fltr.lblINVOS_PLACE_Flow.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Flow like '%" & fltr.txtINVOS_PLACE_Flow.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_ShortName.Value = vbChecked Then
        f = f & " and INVOS_INFO_ShortName like '%" & fltr.txtINVOS_INFO_ShortName.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_WorkPlaceNum.Value = vbChecked Then
        f = f & " and INVOS_PLACE_WorkPlaceNum like '%" & fltr.txtINVOS_PLACE_WorkPlaceNum.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_CardNum.Value = vbChecked Then
        f = f & " and INVOS_INFO_CardNum like '%" & fltr.txtINVOS_INFO_CardNum.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_TheCost_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_TheCost>=" & val(fltr.txtINVOS_INFO_TheCost_GE.Text)
      End If
      If fltr.lblINVOS_INFO_TechFilePath.Value = vbChecked Then
        f = f & " and INVOS_INFO_TechFilePath like '%" & fltr.txtINVOS_INFO_TechFilePath.Text & "%'"
      End If
      If fltr.lblINVOS_CODE_VisibleCode.Value = vbChecked Then
        f = f & " and INVOS_CODE_VisibleCode like '%" & fltr.txtINVOS_CODE_VisibleCode.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_Info.Value = vbChecked Then
        f = f & " and INVOS_INFO_Info like '%" & fltr.txtINVOS_INFO_Info.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_SrokPI_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokPI<=" & val(fltr.txtINVOS_INFO_SrokPI_LE.Text)
      End If
      If fltr.lblINVOS_SROK_RecalcDate_GE.Value = vbChecked Then
        f = f & " and INVOS_SROK_RecalcDate>=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_SROK_RecalcDate_GE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_SROK_RecalcDate_GE.Value), MakePGSQLDate(fltr.dtpINVOS_SROK_RecalcDate_GE.Value)))
      End If
      If fltr.lblINVOS_PLACE_Uprav.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Uprav_ID='" & fltr.txtINVOS_PLACE_Uprav.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_OSType.Value = vbChecked Then
        f = f & " and INVOS_INFO_OSType_ID='" & fltr.txtINVOS_INFO_OSType.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_SrokOI_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokOI<=" & val(fltr.txtINVOS_INFO_SrokOI_LE.Text)
      End If
      If fltr.lblINVOS_INFO_InLineNum_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_InLineNum>=" & val(fltr.txtINVOS_INFO_InLineNum_GE.Text)
      End If
      If fltr.lblINVOS_PLACE_DIrection.Value = vbChecked Then
        f = f & " and INVOS_PLACE_DIrection_ID='" & fltr.txtINVOS_PLACE_DIrection.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_INVNum.Value = vbChecked Then
        f = f & " and INVOS_INFO_INVNum like '%" & fltr.txtINVOS_INFO_INVNum.Text & "%'"
      End If
    jfmnuINV_OS_5.jv.Filter.Add "AUTOINVOS_INFO", f
    End If
      jfmnuINV_OS_5.jv.Refresh
      Me.MousePointer = vbNormal
    End If
    jfmnuINV_OS_5.Show
    jfmnuINV_OS_5.WindowState = 0
    jfmnuINV_OS_5.ZOrder 0
End Sub
Private Sub jfmnuINV_OS_5_OnFilter(usedefault As Boolean)
    Dim fltr As frmINV_OS
    Dim f As String
    Set fltr = New frmINV_OS
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
    f = "1=1"
   f = " INTSANCEStatusID='{55270A15-FA1D-4121-860B-A1B697B40A40}'"
      If fltr.lblINVOS_SROK_RecalcDate_LE.Value = vbChecked Then
        f = f & " and INVOS_SROK_RecalcDate<=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_SROK_RecalcDate_LE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_SROK_RecalcDate_LE.Value), MakePGSQLDate(fltr.dtpINVOS_SROK_RecalcDate_LE.Value)))
      End If
      If fltr.lblINVOS_INFO_IsMaterial.Value = vbChecked Then
        f = f & " and INVOS_INFO_IsMaterial='" & fltr.cmbINVOS_INFO_IsMaterial.Text & "'"
      End If
      If fltr.lblINVOS_PLACE_MatOtv.Value = vbChecked Then
        f = f & " and INVOS_PLACE_MatOtv_ID='" & fltr.txtINVOS_PLACE_MatOtv.Tag & "'"
      End If
      If fltr.lblINVOS_PLACE_Otdel.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Otdel_ID='" & fltr.txtINVOS_PLACE_Otdel.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_SrokFI_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokFI<=" & val(fltr.txtINVOS_INFO_SrokFI_LE.Text)
      End If
      If fltr.lblINVOS_INFO_Name.Value = vbChecked Then
        f = f & " and INVOS_INFO_Name like '%" & fltr.txtINVOS_INFO_Name.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_ActivateDate_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_ActivateDate<=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_INFO_ActivateDate_LE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_INFO_ActivateDate_LE.Value), MakePGSQLDate(fltr.dtpINVOS_INFO_ActivateDate_LE.Value)))
      End If
      If fltr.lblINVOS_INFO_TheOrg.Value = vbChecked Then
        f = f & " and INVOS_INFO_TheOrg_ID='" & fltr.txtINVOS_INFO_TheOrg.Tag & "'"
      End If
      If fltr.lblINVOS_PLACE_ComplNumber.Value = vbChecked Then
        f = f & " and INVOS_PLACE_ComplNumber like '%" & fltr.txtINVOS_PLACE_ComplNumber.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_ActivateDate_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_ActivateDate>=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_INFO_ActivateDate_GE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_INFO_ActivateDate_GE.Value), MakePGSQLDate(fltr.dtpINVOS_INFO_ActivateDate_GE.Value)))
      End If
      If fltr.lblINVOS_PLACE_TheHouse.Value = vbChecked Then
        f = f & " and INVOS_PLACE_TheHouse_ID='" & fltr.txtINVOS_PLACE_TheHouse.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_TheCost_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_TheCost<=" & val(fltr.txtINVOS_INFO_TheCost_LE.Text)
      End If
      If fltr.lblINVOS_INFO_SrokPI_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokPI>=" & val(fltr.txtINVOS_INFO_SrokPI_GE.Text)
      End If
      If fltr.lblINVOS_PLACE_Room.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Room like '%" & fltr.txtINVOS_PLACE_Room.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_SrokOI_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokOI>=" & val(fltr.txtINVOS_INFO_SrokOI_GE.Text)
      End If
      If fltr.lblINVOS_PLACE_TheOwner.Value = vbChecked Then
        f = f & " and INVOS_PLACE_TheOwner_ID='" & fltr.txtINVOS_PLACE_TheOwner.Tag & "'"
      End If
      If fltr.lblINVOS_PLACE_Info.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Info like '%" & fltr.txtINVOS_PLACE_Info.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_InLineNum_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_InLineNum<=" & val(fltr.txtINVOS_INFO_InLineNum_LE.Text)
      End If
      If fltr.lblINVOS_INFO_SrokFI_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokFI>=" & val(fltr.txtINVOS_INFO_SrokFI_GE.Text)
      End If
      If fltr.lblINVOS_PLACE_Flow.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Flow like '%" & fltr.txtINVOS_PLACE_Flow.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_ShortName.Value = vbChecked Then
        f = f & " and INVOS_INFO_ShortName like '%" & fltr.txtINVOS_INFO_ShortName.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_WorkPlaceNum.Value = vbChecked Then
        f = f & " and INVOS_PLACE_WorkPlaceNum like '%" & fltr.txtINVOS_PLACE_WorkPlaceNum.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_CardNum.Value = vbChecked Then
        f = f & " and INVOS_INFO_CardNum like '%" & fltr.txtINVOS_INFO_CardNum.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_TheCost_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_TheCost>=" & val(fltr.txtINVOS_INFO_TheCost_GE.Text)
      End If
      If fltr.lblINVOS_INFO_TechFilePath.Value = vbChecked Then
        f = f & " and INVOS_INFO_TechFilePath like '%" & fltr.txtINVOS_INFO_TechFilePath.Text & "%'"
      End If
      If fltr.lblINVOS_CODE_VisibleCode.Value = vbChecked Then
        f = f & " and INVOS_CODE_VisibleCode like '%" & fltr.txtINVOS_CODE_VisibleCode.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_Info.Value = vbChecked Then
        f = f & " and INVOS_INFO_Info like '%" & fltr.txtINVOS_INFO_Info.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_SrokPI_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokPI<=" & val(fltr.txtINVOS_INFO_SrokPI_LE.Text)
      End If
      If fltr.lblINVOS_SROK_RecalcDate_GE.Value = vbChecked Then
        f = f & " and INVOS_SROK_RecalcDate>=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_SROK_RecalcDate_GE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_SROK_RecalcDate_GE.Value), MakePGSQLDate(fltr.dtpINVOS_SROK_RecalcDate_GE.Value)))
      End If
      If fltr.lblINVOS_PLACE_Uprav.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Uprav_ID='" & fltr.txtINVOS_PLACE_Uprav.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_OSType.Value = vbChecked Then
        f = f & " and INVOS_INFO_OSType_ID='" & fltr.txtINVOS_INFO_OSType.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_SrokOI_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokOI<=" & val(fltr.txtINVOS_INFO_SrokOI_LE.Text)
      End If
      If fltr.lblINVOS_INFO_InLineNum_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_InLineNum>=" & val(fltr.txtINVOS_INFO_InLineNum_GE.Text)
      End If
      If fltr.lblINVOS_PLACE_DIrection.Value = vbChecked Then
        f = f & " and INVOS_PLACE_DIrection_ID='" & fltr.txtINVOS_PLACE_DIrection.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_INVNum.Value = vbChecked Then
        f = f & " and INVOS_INFO_INVNum like '%" & fltr.txtINVOS_INFO_INVNum.Text & "%'"
      End If
    jfmnuINV_OS_5.jv.Filter.Add "AUTOINVOS_INFO", f
    End If
    Unload fltr
    usedefault = False
End Sub
Private Sub jfmnuINV_OS_5_OnClearFilter()
   jfmnuINV_OS_5.jv.Filter.Add "AUTOINVOS_INFO", " INTSANCEStatusID='{55270A15-FA1D-4121-860B-A1B697B40A40}'"
End Sub
Private Sub jfmnuINV_OS_5_OnAdd(usedefaut As Boolean, Refesh As Boolean)
  Dim objGui  As Object
  Dim o As Object
  Dim id As String
  id = CreateGUID2
  Manager.NewInstance id, "INV_OS", "Карточка основного средства" & Now, Site
  Set o = Manager.GetInstanceObject(id)
  If IsDocDenied(o) Then
    MsgBox "Не разрешен доступ к документам такого типа"
    Exit Sub
  End If

  Dim g  As Object
  Set g = Manager.GetInstanceGUI(o.id)
  If Not g Is Nothing Then
    g.Show GetDocumentMode(o), o, False
  End If
  usedefaut = False
  Refesh = False
End Sub


Private Sub mnuINV_OS_6_Click()
    Dim journal As Object
    On Error Resume Next
    If jfmnuINV_OS_6 Is Nothing Then
      Set jfmnuINV_OS_6 = New frmJournalShow
      Set journal = Manager.GetInstanceObject("{5CE3CC0B-5224-4145-B43F-6A29CC390C17}")
      Manager.LockInstanceObject journal.id
      Set jfmnuINV_OS_6.jv.journal = journal
      AddColors jfmnuINV_OS_6.jv
      jfmnuINV_OS_6.jv.OpenModal = False
      jfmnuINV_OS_6.Caption = "Карточка основного средства :В эксплуатации"
      Me.MousePointer = vbHourglass
      DoEvents
      Dim f As String
    f = "1=1"
   f = " INTSANCEStatusID='{8AD15E54-CF87-4FCF-8A1E-A85336E23C73}'"
    jfmnuINV_OS_6.jv.Filter.Add "AUTOINVOS_INFO", f
    Dim fltr As frmINV_OS
    Set fltr = New frmINV_OS
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
      If fltr.lblINVOS_INFO_TechFilePath.Value = vbChecked Then
        f = f & " and INVOS_INFO_TechFilePath like '%" & fltr.txtINVOS_INFO_TechFilePath.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_Uprav.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Uprav_ID='" & fltr.txtINVOS_PLACE_Uprav.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_InLineNum_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_InLineNum>=" & val(fltr.txtINVOS_INFO_InLineNum_GE.Text)
      End If
      If fltr.lblINVOS_PLACE_MatOtv.Value = vbChecked Then
        f = f & " and INVOS_PLACE_MatOtv_ID='" & fltr.txtINVOS_PLACE_MatOtv.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_IsMaterial.Value = vbChecked Then
        f = f & " and INVOS_INFO_IsMaterial='" & fltr.cmbINVOS_INFO_IsMaterial.Text & "'"
      End If
      If fltr.lblINVOS_INFO_INVNum.Value = vbChecked Then
        f = f & " and INVOS_INFO_INVNum like '%" & fltr.txtINVOS_INFO_INVNum.Text & "%'"
      End If
      If fltr.lblINVOS_SROK_RecalcDate_GE.Value = vbChecked Then
        f = f & " and INVOS_SROK_RecalcDate>=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_SROK_RecalcDate_GE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_SROK_RecalcDate_GE.Value), MakePGSQLDate(fltr.dtpINVOS_SROK_RecalcDate_GE.Value)))
      End If
      If fltr.lblINVOS_CODE_VisibleCode.Value = vbChecked Then
        f = f & " and INVOS_CODE_VisibleCode like '%" & fltr.txtINVOS_CODE_VisibleCode.Text & "%'"
      End If
      If fltr.lblINVOS_SROK_RecalcDate_LE.Value = vbChecked Then
        f = f & " and INVOS_SROK_RecalcDate<=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_SROK_RecalcDate_LE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_SROK_RecalcDate_LE.Value), MakePGSQLDate(fltr.dtpINVOS_SROK_RecalcDate_LE.Value)))
      End If
      If fltr.lblINVOS_INFO_CardNum.Value = vbChecked Then
        f = f & " and INVOS_INFO_CardNum like '%" & fltr.txtINVOS_INFO_CardNum.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_TheCost_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_TheCost>=" & val(fltr.txtINVOS_INFO_TheCost_GE.Text)
      End If
      If fltr.lblINVOS_INFO_InLineNum_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_InLineNum<=" & val(fltr.txtINVOS_INFO_InLineNum_LE.Text)
      End If
      If fltr.lblINVOS_INFO_SrokPI_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokPI>=" & val(fltr.txtINVOS_INFO_SrokPI_GE.Text)
      End If
      If fltr.lblINVOS_PLACE_TheOwner.Value = vbChecked Then
        f = f & " and INVOS_PLACE_TheOwner_ID='" & fltr.txtINVOS_PLACE_TheOwner.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_SrokPI_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokPI<=" & val(fltr.txtINVOS_INFO_SrokPI_LE.Text)
      End If
      If fltr.lblINVOS_INFO_ActivateDate_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_ActivateDate<=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_INFO_ActivateDate_LE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_INFO_ActivateDate_LE.Value), MakePGSQLDate(fltr.dtpINVOS_INFO_ActivateDate_LE.Value)))
      End If
      If fltr.lblINVOS_INFO_TheOrg.Value = vbChecked Then
        f = f & " and INVOS_INFO_TheOrg_ID='" & fltr.txtINVOS_INFO_TheOrg.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_SrokOI_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokOI<=" & val(fltr.txtINVOS_INFO_SrokOI_LE.Text)
      End If
      If fltr.lblINVOS_INFO_TheCost_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_TheCost<=" & val(fltr.txtINVOS_INFO_TheCost_LE.Text)
      End If
      If fltr.lblINVOS_PLACE_Info.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Info like '%" & fltr.txtINVOS_PLACE_Info.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_Otdel.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Otdel_ID='" & fltr.txtINVOS_PLACE_Otdel.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_SrokFI_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokFI>=" & val(fltr.txtINVOS_INFO_SrokFI_GE.Text)
      End If
      If fltr.lblINVOS_PLACE_Room.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Room like '%" & fltr.txtINVOS_PLACE_Room.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_Flow.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Flow like '%" & fltr.txtINVOS_PLACE_Flow.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_WorkPlaceNum.Value = vbChecked Then
        f = f & " and INVOS_PLACE_WorkPlaceNum like '%" & fltr.txtINVOS_PLACE_WorkPlaceNum.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_ActivateDate_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_ActivateDate>=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_INFO_ActivateDate_GE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_INFO_ActivateDate_GE.Value), MakePGSQLDate(fltr.dtpINVOS_INFO_ActivateDate_GE.Value)))
      End If
      If fltr.lblINVOS_PLACE_ComplNumber.Value = vbChecked Then
        f = f & " and INVOS_PLACE_ComplNumber like '%" & fltr.txtINVOS_PLACE_ComplNumber.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_DIrection.Value = vbChecked Then
        f = f & " and INVOS_PLACE_DIrection_ID='" & fltr.txtINVOS_PLACE_DIrection.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_SrokOI_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokOI>=" & val(fltr.txtINVOS_INFO_SrokOI_GE.Text)
      End If
      If fltr.lblINVOS_INFO_Info.Value = vbChecked Then
        f = f & " and INVOS_INFO_Info like '%" & fltr.txtINVOS_INFO_Info.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_Name.Value = vbChecked Then
        f = f & " and INVOS_INFO_Name like '%" & fltr.txtINVOS_INFO_Name.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_ShortName.Value = vbChecked Then
        f = f & " and INVOS_INFO_ShortName like '%" & fltr.txtINVOS_INFO_ShortName.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_OSType.Value = vbChecked Then
        f = f & " and INVOS_INFO_OSType_ID='" & fltr.txtINVOS_INFO_OSType.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_SrokFI_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokFI<=" & val(fltr.txtINVOS_INFO_SrokFI_LE.Text)
      End If
      If fltr.lblINVOS_PLACE_TheHouse.Value = vbChecked Then
        f = f & " and INVOS_PLACE_TheHouse_ID='" & fltr.txtINVOS_PLACE_TheHouse.Tag & "'"
      End If
    jfmnuINV_OS_6.jv.Filter.Add "AUTOINVOS_INFO", f
    End If
      jfmnuINV_OS_6.jv.Refresh
      Me.MousePointer = vbNormal
    End If
    jfmnuINV_OS_6.Show
    jfmnuINV_OS_6.WindowState = 0
    jfmnuINV_OS_6.ZOrder 0
End Sub
Private Sub jfmnuINV_OS_6_OnFilter(usedefault As Boolean)
    Dim fltr As frmINV_OS
    Dim f As String
    Set fltr = New frmINV_OS
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
    f = "1=1"
   f = " INTSANCEStatusID='{8AD15E54-CF87-4FCF-8A1E-A85336E23C73}'"
      If fltr.lblINVOS_INFO_TechFilePath.Value = vbChecked Then
        f = f & " and INVOS_INFO_TechFilePath like '%" & fltr.txtINVOS_INFO_TechFilePath.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_Uprav.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Uprav_ID='" & fltr.txtINVOS_PLACE_Uprav.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_InLineNum_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_InLineNum>=" & val(fltr.txtINVOS_INFO_InLineNum_GE.Text)
      End If
      If fltr.lblINVOS_PLACE_MatOtv.Value = vbChecked Then
        f = f & " and INVOS_PLACE_MatOtv_ID='" & fltr.txtINVOS_PLACE_MatOtv.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_IsMaterial.Value = vbChecked Then
        f = f & " and INVOS_INFO_IsMaterial='" & fltr.cmbINVOS_INFO_IsMaterial.Text & "'"
      End If
      If fltr.lblINVOS_INFO_INVNum.Value = vbChecked Then
        f = f & " and INVOS_INFO_INVNum like '%" & fltr.txtINVOS_INFO_INVNum.Text & "%'"
      End If
      If fltr.lblINVOS_SROK_RecalcDate_GE.Value = vbChecked Then
        f = f & " and INVOS_SROK_RecalcDate>=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_SROK_RecalcDate_GE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_SROK_RecalcDate_GE.Value), MakePGSQLDate(fltr.dtpINVOS_SROK_RecalcDate_GE.Value)))
      End If
      If fltr.lblINVOS_CODE_VisibleCode.Value = vbChecked Then
        f = f & " and INVOS_CODE_VisibleCode like '%" & fltr.txtINVOS_CODE_VisibleCode.Text & "%'"
      End If
      If fltr.lblINVOS_SROK_RecalcDate_LE.Value = vbChecked Then
        f = f & " and INVOS_SROK_RecalcDate<=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_SROK_RecalcDate_LE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_SROK_RecalcDate_LE.Value), MakePGSQLDate(fltr.dtpINVOS_SROK_RecalcDate_LE.Value)))
      End If
      If fltr.lblINVOS_INFO_CardNum.Value = vbChecked Then
        f = f & " and INVOS_INFO_CardNum like '%" & fltr.txtINVOS_INFO_CardNum.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_TheCost_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_TheCost>=" & val(fltr.txtINVOS_INFO_TheCost_GE.Text)
      End If
      If fltr.lblINVOS_INFO_InLineNum_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_InLineNum<=" & val(fltr.txtINVOS_INFO_InLineNum_LE.Text)
      End If
      If fltr.lblINVOS_INFO_SrokPI_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokPI>=" & val(fltr.txtINVOS_INFO_SrokPI_GE.Text)
      End If
      If fltr.lblINVOS_PLACE_TheOwner.Value = vbChecked Then
        f = f & " and INVOS_PLACE_TheOwner_ID='" & fltr.txtINVOS_PLACE_TheOwner.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_SrokPI_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokPI<=" & val(fltr.txtINVOS_INFO_SrokPI_LE.Text)
      End If
      If fltr.lblINVOS_INFO_ActivateDate_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_ActivateDate<=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_INFO_ActivateDate_LE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_INFO_ActivateDate_LE.Value), MakePGSQLDate(fltr.dtpINVOS_INFO_ActivateDate_LE.Value)))
      End If
      If fltr.lblINVOS_INFO_TheOrg.Value = vbChecked Then
        f = f & " and INVOS_INFO_TheOrg_ID='" & fltr.txtINVOS_INFO_TheOrg.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_SrokOI_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokOI<=" & val(fltr.txtINVOS_INFO_SrokOI_LE.Text)
      End If
      If fltr.lblINVOS_INFO_TheCost_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_TheCost<=" & val(fltr.txtINVOS_INFO_TheCost_LE.Text)
      End If
      If fltr.lblINVOS_PLACE_Info.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Info like '%" & fltr.txtINVOS_PLACE_Info.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_Otdel.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Otdel_ID='" & fltr.txtINVOS_PLACE_Otdel.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_SrokFI_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokFI>=" & val(fltr.txtINVOS_INFO_SrokFI_GE.Text)
      End If
      If fltr.lblINVOS_PLACE_Room.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Room like '%" & fltr.txtINVOS_PLACE_Room.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_Flow.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Flow like '%" & fltr.txtINVOS_PLACE_Flow.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_WorkPlaceNum.Value = vbChecked Then
        f = f & " and INVOS_PLACE_WorkPlaceNum like '%" & fltr.txtINVOS_PLACE_WorkPlaceNum.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_ActivateDate_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_ActivateDate>=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_INFO_ActivateDate_GE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_INFO_ActivateDate_GE.Value), MakePGSQLDate(fltr.dtpINVOS_INFO_ActivateDate_GE.Value)))
      End If
      If fltr.lblINVOS_PLACE_ComplNumber.Value = vbChecked Then
        f = f & " and INVOS_PLACE_ComplNumber like '%" & fltr.txtINVOS_PLACE_ComplNumber.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_DIrection.Value = vbChecked Then
        f = f & " and INVOS_PLACE_DIrection_ID='" & fltr.txtINVOS_PLACE_DIrection.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_SrokOI_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokOI>=" & val(fltr.txtINVOS_INFO_SrokOI_GE.Text)
      End If
      If fltr.lblINVOS_INFO_Info.Value = vbChecked Then
        f = f & " and INVOS_INFO_Info like '%" & fltr.txtINVOS_INFO_Info.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_Name.Value = vbChecked Then
        f = f & " and INVOS_INFO_Name like '%" & fltr.txtINVOS_INFO_Name.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_ShortName.Value = vbChecked Then
        f = f & " and INVOS_INFO_ShortName like '%" & fltr.txtINVOS_INFO_ShortName.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_OSType.Value = vbChecked Then
        f = f & " and INVOS_INFO_OSType_ID='" & fltr.txtINVOS_INFO_OSType.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_SrokFI_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokFI<=" & val(fltr.txtINVOS_INFO_SrokFI_LE.Text)
      End If
      If fltr.lblINVOS_PLACE_TheHouse.Value = vbChecked Then
        f = f & " and INVOS_PLACE_TheHouse_ID='" & fltr.txtINVOS_PLACE_TheHouse.Tag & "'"
      End If
    jfmnuINV_OS_6.jv.Filter.Add "AUTOINVOS_INFO", f
    End If
    Unload fltr
    usedefault = False
End Sub
Private Sub jfmnuINV_OS_6_OnClearFilter()
   jfmnuINV_OS_6.jv.Filter.Add "AUTOINVOS_INFO", " INTSANCEStatusID='{8AD15E54-CF87-4FCF-8A1E-A85336E23C73}'"
End Sub
Private Sub jfmnuINV_OS_6_OnAdd(usedefaut As Boolean, Refesh As Boolean)
  Dim objGui  As Object
  Dim o As Object
  Dim id As String
  id = CreateGUID2
  Manager.NewInstance id, "INV_OS", "Карточка основного средства" & Now, Site
  Set o = Manager.GetInstanceObject(id)
  If IsDocDenied(o) Then
    MsgBox "Не разрешен доступ к документам такого типа"
    Exit Sub
  End If

  Dim g  As Object
  Set g = Manager.GetInstanceGUI(o.id)
  If Not g Is Nothing Then
    g.Show GetDocumentMode(o), o, False
  End If
  usedefaut = False
  Refesh = False
End Sub


Private Sub mnuINV_OS_7_Click()
    Dim journal As Object
    On Error Resume Next
    If jfmnuINV_OS_7 Is Nothing Then
      Set jfmnuINV_OS_7 = New frmJournalShow
      Set journal = Manager.GetInstanceObject("{5CE3CC0B-5224-4145-B43F-6A29CC390C17}")
      Manager.LockInstanceObject journal.id
      Set jfmnuINV_OS_7.jv.journal = journal
      AddColors jfmnuINV_OS_7.jv
      jfmnuINV_OS_7.jv.OpenModal = False
      jfmnuINV_OS_7.Caption = "Карточка основного средства :Списано"
      Me.MousePointer = vbHourglass
      DoEvents
      Dim f As String
    f = "1=1"
   f = " INTSANCEStatusID='{166D4978-0C4C-4575-8192-B251AC113781}'"
    jfmnuINV_OS_7.jv.Filter.Add "AUTOINVOS_INFO", f
    Dim fltr As frmINV_OS
    Set fltr = New frmINV_OS
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
      If fltr.lblINVOS_INFO_SrokFI_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokFI<=" & val(fltr.txtINVOS_INFO_SrokFI_LE.Text)
      End If
      If fltr.lblINVOS_INFO_SrokPI_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokPI<=" & val(fltr.txtINVOS_INFO_SrokPI_LE.Text)
      End If
      If fltr.lblINVOS_PLACE_DIrection.Value = vbChecked Then
        f = f & " and INVOS_PLACE_DIrection_ID='" & fltr.txtINVOS_PLACE_DIrection.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_SrokFI_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokFI>=" & val(fltr.txtINVOS_INFO_SrokFI_GE.Text)
      End If
      If fltr.lblINVOS_CODE_VisibleCode.Value = vbChecked Then
        f = f & " and INVOS_CODE_VisibleCode like '%" & fltr.txtINVOS_CODE_VisibleCode.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_TechFilePath.Value = vbChecked Then
        f = f & " and INVOS_INFO_TechFilePath like '%" & fltr.txtINVOS_INFO_TechFilePath.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_IsMaterial.Value = vbChecked Then
        f = f & " and INVOS_INFO_IsMaterial='" & fltr.cmbINVOS_INFO_IsMaterial.Text & "'"
      End If
      If fltr.lblINVOS_INFO_TheOrg.Value = vbChecked Then
        f = f & " and INVOS_INFO_TheOrg_ID='" & fltr.txtINVOS_INFO_TheOrg.Tag & "'"
      End If
      If fltr.lblINVOS_PLACE_MatOtv.Value = vbChecked Then
        f = f & " and INVOS_PLACE_MatOtv_ID='" & fltr.txtINVOS_PLACE_MatOtv.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_SrokPI_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokPI>=" & val(fltr.txtINVOS_INFO_SrokPI_GE.Text)
      End If
      If fltr.lblINVOS_INFO_TheCost_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_TheCost<=" & val(fltr.txtINVOS_INFO_TheCost_LE.Text)
      End If
      If fltr.lblINVOS_PLACE_TheOwner.Value = vbChecked Then
        f = f & " and INVOS_PLACE_TheOwner_ID='" & fltr.txtINVOS_PLACE_TheOwner.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_SrokOI_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokOI<=" & val(fltr.txtINVOS_INFO_SrokOI_LE.Text)
      End If
      If fltr.lblINVOS_INFO_InLineNum_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_InLineNum<=" & val(fltr.txtINVOS_INFO_InLineNum_LE.Text)
      End If
      If fltr.lblINVOS_SROK_RecalcDate_LE.Value = vbChecked Then
        f = f & " and INVOS_SROK_RecalcDate<=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_SROK_RecalcDate_LE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_SROK_RecalcDate_LE.Value), MakePGSQLDate(fltr.dtpINVOS_SROK_RecalcDate_LE.Value)))
      End If
      If fltr.lblINVOS_PLACE_Otdel.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Otdel_ID='" & fltr.txtINVOS_PLACE_Otdel.Tag & "'"
      End If
      If fltr.lblINVOS_PLACE_Uprav.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Uprav_ID='" & fltr.txtINVOS_PLACE_Uprav.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_Name.Value = vbChecked Then
        f = f & " and INVOS_INFO_Name like '%" & fltr.txtINVOS_INFO_Name.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_Info.Value = vbChecked Then
        f = f & " and INVOS_INFO_Info like '%" & fltr.txtINVOS_INFO_Info.Text & "%'"
      End If
      If fltr.lblINVOS_SROK_RecalcDate_GE.Value = vbChecked Then
        f = f & " and INVOS_SROK_RecalcDate>=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_SROK_RecalcDate_GE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_SROK_RecalcDate_GE.Value), MakePGSQLDate(fltr.dtpINVOS_SROK_RecalcDate_GE.Value)))
      End If
      If fltr.lblINVOS_PLACE_Room.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Room like '%" & fltr.txtINVOS_PLACE_Room.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_TheHouse.Value = vbChecked Then
        f = f & " and INVOS_PLACE_TheHouse_ID='" & fltr.txtINVOS_PLACE_TheHouse.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_INVNum.Value = vbChecked Then
        f = f & " and INVOS_INFO_INVNum like '%" & fltr.txtINVOS_INFO_INVNum.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_WorkPlaceNum.Value = vbChecked Then
        f = f & " and INVOS_PLACE_WorkPlaceNum like '%" & fltr.txtINVOS_PLACE_WorkPlaceNum.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_InLineNum_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_InLineNum>=" & val(fltr.txtINVOS_INFO_InLineNum_GE.Text)
      End If
      If fltr.lblINVOS_INFO_ActivateDate_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_ActivateDate<=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_INFO_ActivateDate_LE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_INFO_ActivateDate_LE.Value), MakePGSQLDate(fltr.dtpINVOS_INFO_ActivateDate_LE.Value)))
      End If
      If fltr.lblINVOS_PLACE_Flow.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Flow like '%" & fltr.txtINVOS_PLACE_Flow.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_ComplNumber.Value = vbChecked Then
        f = f & " and INVOS_PLACE_ComplNumber like '%" & fltr.txtINVOS_PLACE_ComplNumber.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_TheCost_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_TheCost>=" & val(fltr.txtINVOS_INFO_TheCost_GE.Text)
      End If
      If fltr.lblINVOS_INFO_ShortName.Value = vbChecked Then
        f = f & " and INVOS_INFO_ShortName like '%" & fltr.txtINVOS_INFO_ShortName.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_ActivateDate_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_ActivateDate>=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_INFO_ActivateDate_GE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_INFO_ActivateDate_GE.Value), MakePGSQLDate(fltr.dtpINVOS_INFO_ActivateDate_GE.Value)))
      End If
      If fltr.lblINVOS_INFO_SrokOI_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokOI>=" & val(fltr.txtINVOS_INFO_SrokOI_GE.Text)
      End If
      If fltr.lblINVOS_PLACE_Info.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Info like '%" & fltr.txtINVOS_PLACE_Info.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_CardNum.Value = vbChecked Then
        f = f & " and INVOS_INFO_CardNum like '%" & fltr.txtINVOS_INFO_CardNum.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_OSType.Value = vbChecked Then
        f = f & " and INVOS_INFO_OSType_ID='" & fltr.txtINVOS_INFO_OSType.Tag & "'"
      End If
    jfmnuINV_OS_7.jv.Filter.Add "AUTOINVOS_INFO", f
    End If
      jfmnuINV_OS_7.jv.Refresh
      Me.MousePointer = vbNormal
    End If
    jfmnuINV_OS_7.Show
    jfmnuINV_OS_7.WindowState = 0
    jfmnuINV_OS_7.ZOrder 0
End Sub
Private Sub jfmnuINV_OS_7_OnFilter(usedefault As Boolean)
    Dim fltr As frmINV_OS
    Dim f As String
    Set fltr = New frmINV_OS
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
    f = "1=1"
   f = " INTSANCEStatusID='{166D4978-0C4C-4575-8192-B251AC113781}'"
      If fltr.lblINVOS_INFO_SrokFI_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokFI<=" & val(fltr.txtINVOS_INFO_SrokFI_LE.Text)
      End If
      If fltr.lblINVOS_INFO_SrokPI_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokPI<=" & val(fltr.txtINVOS_INFO_SrokPI_LE.Text)
      End If
      If fltr.lblINVOS_PLACE_DIrection.Value = vbChecked Then
        f = f & " and INVOS_PLACE_DIrection_ID='" & fltr.txtINVOS_PLACE_DIrection.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_SrokFI_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokFI>=" & val(fltr.txtINVOS_INFO_SrokFI_GE.Text)
      End If
      If fltr.lblINVOS_CODE_VisibleCode.Value = vbChecked Then
        f = f & " and INVOS_CODE_VisibleCode like '%" & fltr.txtINVOS_CODE_VisibleCode.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_TechFilePath.Value = vbChecked Then
        f = f & " and INVOS_INFO_TechFilePath like '%" & fltr.txtINVOS_INFO_TechFilePath.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_IsMaterial.Value = vbChecked Then
        f = f & " and INVOS_INFO_IsMaterial='" & fltr.cmbINVOS_INFO_IsMaterial.Text & "'"
      End If
      If fltr.lblINVOS_INFO_TheOrg.Value = vbChecked Then
        f = f & " and INVOS_INFO_TheOrg_ID='" & fltr.txtINVOS_INFO_TheOrg.Tag & "'"
      End If
      If fltr.lblINVOS_PLACE_MatOtv.Value = vbChecked Then
        f = f & " and INVOS_PLACE_MatOtv_ID='" & fltr.txtINVOS_PLACE_MatOtv.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_SrokPI_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokPI>=" & val(fltr.txtINVOS_INFO_SrokPI_GE.Text)
      End If
      If fltr.lblINVOS_INFO_TheCost_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_TheCost<=" & val(fltr.txtINVOS_INFO_TheCost_LE.Text)
      End If
      If fltr.lblINVOS_PLACE_TheOwner.Value = vbChecked Then
        f = f & " and INVOS_PLACE_TheOwner_ID='" & fltr.txtINVOS_PLACE_TheOwner.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_SrokOI_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokOI<=" & val(fltr.txtINVOS_INFO_SrokOI_LE.Text)
      End If
      If fltr.lblINVOS_INFO_InLineNum_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_InLineNum<=" & val(fltr.txtINVOS_INFO_InLineNum_LE.Text)
      End If
      If fltr.lblINVOS_SROK_RecalcDate_LE.Value = vbChecked Then
        f = f & " and INVOS_SROK_RecalcDate<=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_SROK_RecalcDate_LE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_SROK_RecalcDate_LE.Value), MakePGSQLDate(fltr.dtpINVOS_SROK_RecalcDate_LE.Value)))
      End If
      If fltr.lblINVOS_PLACE_Otdel.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Otdel_ID='" & fltr.txtINVOS_PLACE_Otdel.Tag & "'"
      End If
      If fltr.lblINVOS_PLACE_Uprav.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Uprav_ID='" & fltr.txtINVOS_PLACE_Uprav.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_Name.Value = vbChecked Then
        f = f & " and INVOS_INFO_Name like '%" & fltr.txtINVOS_INFO_Name.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_Info.Value = vbChecked Then
        f = f & " and INVOS_INFO_Info like '%" & fltr.txtINVOS_INFO_Info.Text & "%'"
      End If
      If fltr.lblINVOS_SROK_RecalcDate_GE.Value = vbChecked Then
        f = f & " and INVOS_SROK_RecalcDate>=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_SROK_RecalcDate_GE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_SROK_RecalcDate_GE.Value), MakePGSQLDate(fltr.dtpINVOS_SROK_RecalcDate_GE.Value)))
      End If
      If fltr.lblINVOS_PLACE_Room.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Room like '%" & fltr.txtINVOS_PLACE_Room.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_TheHouse.Value = vbChecked Then
        f = f & " and INVOS_PLACE_TheHouse_ID='" & fltr.txtINVOS_PLACE_TheHouse.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_INVNum.Value = vbChecked Then
        f = f & " and INVOS_INFO_INVNum like '%" & fltr.txtINVOS_INFO_INVNum.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_WorkPlaceNum.Value = vbChecked Then
        f = f & " and INVOS_PLACE_WorkPlaceNum like '%" & fltr.txtINVOS_PLACE_WorkPlaceNum.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_InLineNum_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_InLineNum>=" & val(fltr.txtINVOS_INFO_InLineNum_GE.Text)
      End If
      If fltr.lblINVOS_INFO_ActivateDate_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_ActivateDate<=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_INFO_ActivateDate_LE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_INFO_ActivateDate_LE.Value), MakePGSQLDate(fltr.dtpINVOS_INFO_ActivateDate_LE.Value)))
      End If
      If fltr.lblINVOS_PLACE_Flow.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Flow like '%" & fltr.txtINVOS_PLACE_Flow.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_ComplNumber.Value = vbChecked Then
        f = f & " and INVOS_PLACE_ComplNumber like '%" & fltr.txtINVOS_PLACE_ComplNumber.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_TheCost_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_TheCost>=" & val(fltr.txtINVOS_INFO_TheCost_GE.Text)
      End If
      If fltr.lblINVOS_INFO_ShortName.Value = vbChecked Then
        f = f & " and INVOS_INFO_ShortName like '%" & fltr.txtINVOS_INFO_ShortName.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_ActivateDate_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_ActivateDate>=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_INFO_ActivateDate_GE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_INFO_ActivateDate_GE.Value), MakePGSQLDate(fltr.dtpINVOS_INFO_ActivateDate_GE.Value)))
      End If
      If fltr.lblINVOS_INFO_SrokOI_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokOI>=" & val(fltr.txtINVOS_INFO_SrokOI_GE.Text)
      End If
      If fltr.lblINVOS_PLACE_Info.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Info like '%" & fltr.txtINVOS_PLACE_Info.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_CardNum.Value = vbChecked Then
        f = f & " and INVOS_INFO_CardNum like '%" & fltr.txtINVOS_INFO_CardNum.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_OSType.Value = vbChecked Then
        f = f & " and INVOS_INFO_OSType_ID='" & fltr.txtINVOS_INFO_OSType.Tag & "'"
      End If
    jfmnuINV_OS_7.jv.Filter.Add "AUTOINVOS_INFO", f
    End If
    Unload fltr
    usedefault = False
End Sub
Private Sub jfmnuINV_OS_7_OnClearFilter()
   jfmnuINV_OS_7.jv.Filter.Add "AUTOINVOS_INFO", " INTSANCEStatusID='{166D4978-0C4C-4575-8192-B251AC113781}'"
End Sub
Private Sub jfmnuINV_OS_7_OnAdd(usedefaut As Boolean, Refesh As Boolean)
  Dim objGui  As Object
  Dim o As Object
  Dim id As String
  id = CreateGUID2
  Manager.NewInstance id, "INV_OS", "Карточка основного средства" & Now, Site
  Set o = Manager.GetInstanceObject(id)
  If IsDocDenied(o) Then
    MsgBox "Не разрешен доступ к документам такого типа"
    Exit Sub
  End If

  Dim g  As Object
  Set g = Manager.GetInstanceGUI(o.id)
  If Not g Is Nothing Then
    g.Show GetDocumentMode(o), o, False
  End If
  usedefaut = False
  Refesh = False
End Sub


Private Sub mnuINV_OS_8_Click()
    Dim journal As Object
    On Error Resume Next
    If jfmnuINV_OS_8 Is Nothing Then
      Set jfmnuINV_OS_8 = New frmJournalShow
      Set journal = Manager.GetInstanceObject("{5CE3CC0B-5224-4145-B43F-6A29CC390C17}")
      Manager.LockInstanceObject journal.id
      Set jfmnuINV_OS_8.jv.journal = journal
      AddColors jfmnuINV_OS_8.jv
      jfmnuINV_OS_8.jv.OpenModal = False
      jfmnuINV_OS_8.Caption = "Карточка основного средства :На консервации"
      Me.MousePointer = vbHourglass
      DoEvents
      Dim f As String
    f = "1=1"
   f = " INTSANCEStatusID='{DA1E3744-00B3-4D9E-AA07-BE499D2402E4}'"
    jfmnuINV_OS_8.jv.Filter.Add "AUTOINVOS_INFO", f
    Dim fltr As frmINV_OS
    Set fltr = New frmINV_OS
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
      If fltr.lblINVOS_PLACE_Info.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Info like '%" & fltr.txtINVOS_PLACE_Info.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_ActivateDate_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_ActivateDate>=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_INFO_ActivateDate_GE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_INFO_ActivateDate_GE.Value), MakePGSQLDate(fltr.dtpINVOS_INFO_ActivateDate_GE.Value)))
      End If
      If fltr.lblINVOS_PLACE_Flow.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Flow like '%" & fltr.txtINVOS_PLACE_Flow.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_ComplNumber.Value = vbChecked Then
        f = f & " and INVOS_PLACE_ComplNumber like '%" & fltr.txtINVOS_PLACE_ComplNumber.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_InLineNum_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_InLineNum<=" & val(fltr.txtINVOS_INFO_InLineNum_LE.Text)
      End If
      If fltr.lblINVOS_INFO_TheOrg.Value = vbChecked Then
        f = f & " and INVOS_INFO_TheOrg_ID='" & fltr.txtINVOS_INFO_TheOrg.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_SrokPI_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokPI<=" & val(fltr.txtINVOS_INFO_SrokPI_LE.Text)
      End If
      If fltr.lblINVOS_INFO_OSType.Value = vbChecked Then
        f = f & " and INVOS_INFO_OSType_ID='" & fltr.txtINVOS_INFO_OSType.Tag & "'"
      End If
      If fltr.lblINVOS_PLACE_MatOtv.Value = vbChecked Then
        f = f & " and INVOS_PLACE_MatOtv_ID='" & fltr.txtINVOS_PLACE_MatOtv.Tag & "'"
      End If
      If fltr.lblINVOS_SROK_RecalcDate_GE.Value = vbChecked Then
        f = f & " and INVOS_SROK_RecalcDate>=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_SROK_RecalcDate_GE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_SROK_RecalcDate_GE.Value), MakePGSQLDate(fltr.dtpINVOS_SROK_RecalcDate_GE.Value)))
      End If
      If fltr.lblINVOS_INFO_TechFilePath.Value = vbChecked Then
        f = f & " and INVOS_INFO_TechFilePath like '%" & fltr.txtINVOS_INFO_TechFilePath.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_Uprav.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Uprav_ID='" & fltr.txtINVOS_PLACE_Uprav.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_ShortName.Value = vbChecked Then
        f = f & " and INVOS_INFO_ShortName like '%" & fltr.txtINVOS_INFO_ShortName.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_TheCost_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_TheCost<=" & val(fltr.txtINVOS_INFO_TheCost_LE.Text)
      End If
      If fltr.lblINVOS_INFO_CardNum.Value = vbChecked Then
        f = f & " and INVOS_INFO_CardNum like '%" & fltr.txtINVOS_INFO_CardNum.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_DIrection.Value = vbChecked Then
        f = f & " and INVOS_PLACE_DIrection_ID='" & fltr.txtINVOS_PLACE_DIrection.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_SrokFI_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokFI<=" & val(fltr.txtINVOS_INFO_SrokFI_LE.Text)
      End If
      If fltr.lblINVOS_INFO_TheCost_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_TheCost>=" & val(fltr.txtINVOS_INFO_TheCost_GE.Text)
      End If
      If fltr.lblINVOS_PLACE_TheHouse.Value = vbChecked Then
        f = f & " and INVOS_PLACE_TheHouse_ID='" & fltr.txtINVOS_PLACE_TheHouse.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_Info.Value = vbChecked Then
        f = f & " and INVOS_INFO_Info like '%" & fltr.txtINVOS_INFO_Info.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_SrokOI_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokOI>=" & val(fltr.txtINVOS_INFO_SrokOI_GE.Text)
      End If
      If fltr.lblINVOS_INFO_ActivateDate_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_ActivateDate<=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_INFO_ActivateDate_LE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_INFO_ActivateDate_LE.Value), MakePGSQLDate(fltr.dtpINVOS_INFO_ActivateDate_LE.Value)))
      End If
      If fltr.lblINVOS_INFO_SrokOI_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokOI<=" & val(fltr.txtINVOS_INFO_SrokOI_LE.Text)
      End If
      If fltr.lblINVOS_INFO_INVNum.Value = vbChecked Then
        f = f & " and INVOS_INFO_INVNum like '%" & fltr.txtINVOS_INFO_INVNum.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_Otdel.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Otdel_ID='" & fltr.txtINVOS_PLACE_Otdel.Tag & "'"
      End If
      If fltr.lblINVOS_CODE_VisibleCode.Value = vbChecked Then
        f = f & " and INVOS_CODE_VisibleCode like '%" & fltr.txtINVOS_CODE_VisibleCode.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_Name.Value = vbChecked Then
        f = f & " and INVOS_INFO_Name like '%" & fltr.txtINVOS_INFO_Name.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_InLineNum_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_InLineNum>=" & val(fltr.txtINVOS_INFO_InLineNum_GE.Text)
      End If
      If fltr.lblINVOS_SROK_RecalcDate_LE.Value = vbChecked Then
        f = f & " and INVOS_SROK_RecalcDate<=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_SROK_RecalcDate_LE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_SROK_RecalcDate_LE.Value), MakePGSQLDate(fltr.dtpINVOS_SROK_RecalcDate_LE.Value)))
      End If
      If fltr.lblINVOS_INFO_SrokPI_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokPI>=" & val(fltr.txtINVOS_INFO_SrokPI_GE.Text)
      End If
      If fltr.lblINVOS_INFO_SrokFI_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokFI>=" & val(fltr.txtINVOS_INFO_SrokFI_GE.Text)
      End If
      If fltr.lblINVOS_PLACE_TheOwner.Value = vbChecked Then
        f = f & " and INVOS_PLACE_TheOwner_ID='" & fltr.txtINVOS_PLACE_TheOwner.Tag & "'"
      End If
      If fltr.lblINVOS_PLACE_WorkPlaceNum.Value = vbChecked Then
        f = f & " and INVOS_PLACE_WorkPlaceNum like '%" & fltr.txtINVOS_PLACE_WorkPlaceNum.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_Room.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Room like '%" & fltr.txtINVOS_PLACE_Room.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_IsMaterial.Value = vbChecked Then
        f = f & " and INVOS_INFO_IsMaterial='" & fltr.cmbINVOS_INFO_IsMaterial.Text & "'"
      End If
    jfmnuINV_OS_8.jv.Filter.Add "AUTOINVOS_INFO", f
    End If
      jfmnuINV_OS_8.jv.Refresh
      Me.MousePointer = vbNormal
    End If
    jfmnuINV_OS_8.Show
    jfmnuINV_OS_8.WindowState = 0
    jfmnuINV_OS_8.ZOrder 0
End Sub
Private Sub jfmnuINV_OS_8_OnFilter(usedefault As Boolean)
    Dim fltr As frmINV_OS
    Dim f As String
    Set fltr = New frmINV_OS
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
    f = "1=1"
   f = " INTSANCEStatusID='{DA1E3744-00B3-4D9E-AA07-BE499D2402E4}'"
      If fltr.lblINVOS_PLACE_Info.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Info like '%" & fltr.txtINVOS_PLACE_Info.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_ActivateDate_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_ActivateDate>=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_INFO_ActivateDate_GE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_INFO_ActivateDate_GE.Value), MakePGSQLDate(fltr.dtpINVOS_INFO_ActivateDate_GE.Value)))
      End If
      If fltr.lblINVOS_PLACE_Flow.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Flow like '%" & fltr.txtINVOS_PLACE_Flow.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_ComplNumber.Value = vbChecked Then
        f = f & " and INVOS_PLACE_ComplNumber like '%" & fltr.txtINVOS_PLACE_ComplNumber.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_InLineNum_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_InLineNum<=" & val(fltr.txtINVOS_INFO_InLineNum_LE.Text)
      End If
      If fltr.lblINVOS_INFO_TheOrg.Value = vbChecked Then
        f = f & " and INVOS_INFO_TheOrg_ID='" & fltr.txtINVOS_INFO_TheOrg.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_SrokPI_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokPI<=" & val(fltr.txtINVOS_INFO_SrokPI_LE.Text)
      End If
      If fltr.lblINVOS_INFO_OSType.Value = vbChecked Then
        f = f & " and INVOS_INFO_OSType_ID='" & fltr.txtINVOS_INFO_OSType.Tag & "'"
      End If
      If fltr.lblINVOS_PLACE_MatOtv.Value = vbChecked Then
        f = f & " and INVOS_PLACE_MatOtv_ID='" & fltr.txtINVOS_PLACE_MatOtv.Tag & "'"
      End If
      If fltr.lblINVOS_SROK_RecalcDate_GE.Value = vbChecked Then
        f = f & " and INVOS_SROK_RecalcDate>=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_SROK_RecalcDate_GE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_SROK_RecalcDate_GE.Value), MakePGSQLDate(fltr.dtpINVOS_SROK_RecalcDate_GE.Value)))
      End If
      If fltr.lblINVOS_INFO_TechFilePath.Value = vbChecked Then
        f = f & " and INVOS_INFO_TechFilePath like '%" & fltr.txtINVOS_INFO_TechFilePath.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_Uprav.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Uprav_ID='" & fltr.txtINVOS_PLACE_Uprav.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_ShortName.Value = vbChecked Then
        f = f & " and INVOS_INFO_ShortName like '%" & fltr.txtINVOS_INFO_ShortName.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_TheCost_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_TheCost<=" & val(fltr.txtINVOS_INFO_TheCost_LE.Text)
      End If
      If fltr.lblINVOS_INFO_CardNum.Value = vbChecked Then
        f = f & " and INVOS_INFO_CardNum like '%" & fltr.txtINVOS_INFO_CardNum.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_DIrection.Value = vbChecked Then
        f = f & " and INVOS_PLACE_DIrection_ID='" & fltr.txtINVOS_PLACE_DIrection.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_SrokFI_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokFI<=" & val(fltr.txtINVOS_INFO_SrokFI_LE.Text)
      End If
      If fltr.lblINVOS_INFO_TheCost_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_TheCost>=" & val(fltr.txtINVOS_INFO_TheCost_GE.Text)
      End If
      If fltr.lblINVOS_PLACE_TheHouse.Value = vbChecked Then
        f = f & " and INVOS_PLACE_TheHouse_ID='" & fltr.txtINVOS_PLACE_TheHouse.Tag & "'"
      End If
      If fltr.lblINVOS_INFO_Info.Value = vbChecked Then
        f = f & " and INVOS_INFO_Info like '%" & fltr.txtINVOS_INFO_Info.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_SrokOI_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokOI>=" & val(fltr.txtINVOS_INFO_SrokOI_GE.Text)
      End If
      If fltr.lblINVOS_INFO_ActivateDate_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_ActivateDate<=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_INFO_ActivateDate_LE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_INFO_ActivateDate_LE.Value), MakePGSQLDate(fltr.dtpINVOS_INFO_ActivateDate_LE.Value)))
      End If
      If fltr.lblINVOS_INFO_SrokOI_LE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokOI<=" & val(fltr.txtINVOS_INFO_SrokOI_LE.Text)
      End If
      If fltr.lblINVOS_INFO_INVNum.Value = vbChecked Then
        f = f & " and INVOS_INFO_INVNum like '%" & fltr.txtINVOS_INFO_INVNum.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_Otdel.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Otdel_ID='" & fltr.txtINVOS_PLACE_Otdel.Tag & "'"
      End If
      If fltr.lblINVOS_CODE_VisibleCode.Value = vbChecked Then
        f = f & " and INVOS_CODE_VisibleCode like '%" & fltr.txtINVOS_CODE_VisibleCode.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_Name.Value = vbChecked Then
        f = f & " and INVOS_INFO_Name like '%" & fltr.txtINVOS_INFO_Name.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_InLineNum_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_InLineNum>=" & val(fltr.txtINVOS_INFO_InLineNum_GE.Text)
      End If
      If fltr.lblINVOS_SROK_RecalcDate_LE.Value = vbChecked Then
        f = f & " and INVOS_SROK_RecalcDate<=" & IIf(Session.IsMSSQL, MakeMSSQLDate(fltr.dtpINVOS_SROK_RecalcDate_LE.Value), IIf(Session.IsORACLE, MakeORACLEDate(fltr.dtpINVOS_SROK_RecalcDate_LE.Value), MakePGSQLDate(fltr.dtpINVOS_SROK_RecalcDate_LE.Value)))
      End If
      If fltr.lblINVOS_INFO_SrokPI_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokPI>=" & val(fltr.txtINVOS_INFO_SrokPI_GE.Text)
      End If
      If fltr.lblINVOS_INFO_SrokFI_GE.Value = vbChecked Then
        f = f & " and INVOS_INFO_SrokFI>=" & val(fltr.txtINVOS_INFO_SrokFI_GE.Text)
      End If
      If fltr.lblINVOS_PLACE_TheOwner.Value = vbChecked Then
        f = f & " and INVOS_PLACE_TheOwner_ID='" & fltr.txtINVOS_PLACE_TheOwner.Tag & "'"
      End If
      If fltr.lblINVOS_PLACE_WorkPlaceNum.Value = vbChecked Then
        f = f & " and INVOS_PLACE_WorkPlaceNum like '%" & fltr.txtINVOS_PLACE_WorkPlaceNum.Text & "%'"
      End If
      If fltr.lblINVOS_PLACE_Room.Value = vbChecked Then
        f = f & " and INVOS_PLACE_Room like '%" & fltr.txtINVOS_PLACE_Room.Text & "%'"
      End If
      If fltr.lblINVOS_INFO_IsMaterial.Value = vbChecked Then
        f = f & " and INVOS_INFO_IsMaterial='" & fltr.cmbINVOS_INFO_IsMaterial.Text & "'"
      End If
    jfmnuINV_OS_8.jv.Filter.Add "AUTOINVOS_INFO", f
    End If
    Unload fltr
    usedefault = False
End Sub
Private Sub jfmnuINV_OS_8_OnClearFilter()
   jfmnuINV_OS_8.jv.Filter.Add "AUTOINVOS_INFO", " INTSANCEStatusID='{DA1E3744-00B3-4D9E-AA07-BE499D2402E4}'"
End Sub
Private Sub jfmnuINV_OS_8_OnAdd(usedefaut As Boolean, Refesh As Boolean)
  Dim objGui  As Object
  Dim o As Object
  Dim id As String
  id = CreateGUID2
  Manager.NewInstance id, "INV_OS", "Карточка основного средства" & Now, Site
  Set o = Manager.GetInstanceObject(id)
  If IsDocDenied(o) Then
    MsgBox "Не разрешен доступ к документам такого типа"
    Exit Sub
  End If

  Dim g  As Object
  Set g = Manager.GetInstanceGUI(o.id)
  If Not g Is Nothing Then
    g.Show GetDocumentMode(o), o, False
  End If
  usedefaut = False
  Refesh = False
End Sub





Private Sub mnuINV_DIC_Click()
 Dim o As Object
 Dim rs  As ADODB.Recordset
 Dim id As String
  Set rs = Manager.ListInstances("", "INV_DIC")
  If Not rs.EOF Then
    id = rs!InstanceID
  Else
    id = CreateGUID2
    Manager.NewInstance id, "INV_DIC", "Справочник"
  End If
    Set o = Manager.GetInstanceObject(id)
    If IsDocDenied(o) Then
      MsgBox "Не разрешен доступ к документам такого типа"
      Exit Sub
    End If

    Dim g  As Object
    Set g = Manager.GetInstanceGUI(o.id)
    If Not g Is Nothing Then
      g.Show GetDocumentMode(o), o, False
    End If
  Set rs = Nothing
End Sub




Private Sub UnloadObjects()



Unload jfmnuAllINV_INV
Set jfmnuAllINV_INV = Nothing

Unload jfmnuINV_INV_1
Set jfmnuINV_INV_1 = Nothing

Unload jfmnuINV_INV_2
Set jfmnuINV_INV_2 = Nothing

Unload jfmnuINV_INV_3
Set jfmnuINV_INV_3 = Nothing

Unload jfmnuINV_INV_4
Set jfmnuINV_INV_4 = Nothing


Unload jfmnuAllINV_OS
Set jfmnuAllINV_OS = Nothing

Unload jfmnuINV_OS_1
Set jfmnuINV_OS_1 = Nothing

Unload jfmnuINV_OS_2
Set jfmnuINV_OS_2 = Nothing

Unload jfmnuINV_OS_3
Set jfmnuINV_OS_3 = Nothing

Unload jfmnuINV_OS_4
Set jfmnuINV_OS_4 = Nothing

Unload jfmnuINV_OS_5
Set jfmnuINV_OS_5 = Nothing


Unload jfmnuINV_OS_6
Set jfmnuINV_OS_6 = Nothing

Unload jfmnuINV_OS_7
Set jfmnuINV_OS_7 = Nothing

Unload jfmnuINV_OS_8
Set jfmnuINV_OS_8 = Nothing

Unload jfmnuINV_OS_OK
Set jfmnuINV_OS_OK = Nothing

Unload jfmnuINV_OS_BAD
Set jfmnuINV_OS_BAD = Nothing


Unload jfmnuINV_OS_BAD
Set jfmnuINV_NUM = Nothing

Unload jfmnuINVF
Set jfmnuINVF = Nothing

End Sub



Private Sub jfmnuINV_NUM_OnFilter(usedefault As Boolean)
    Dim fltr As frmINV_NUM
    Dim f As String
    Set fltr = New frmINV_NUM
    fltr.Show vbModal
    If fltr.OK Then
    ' build flter expression
    f = "1=1"
      If fltr.lblINVN_DEF_TheNumber_GE.Value = vbChecked Then
        f = f & " and INVN_DEF_TheNumber>=" & val(fltr.txtINVN_DEF_TheNumber_GE.Text)
      End If
      If fltr.lblINVN_DEF_TheNumber_LE.Value = vbChecked Then
        f = f & " and INVN_DEF_TheNumber<=" & val(fltr.txtINVN_DEF_TheNumber_LE.Text)
      End If
    jfmnuINV_NUM.jv.Filter.Add "AUTOINVN_DEF", f
    End If
    Unload fltr
    usedefault = False
End Sub
Private Sub jfmnuINV_NUM_OnAdd(usedefaut As Boolean, Refesh As Boolean)
  Dim objGui  As Object
  Dim o As Object
  Dim id As String
  id = CreateGUID2
  Manager.NewInstance id, "INV_NUM", "Нумерация" & Now, Site
  Set o = Manager.GetInstanceObject(id)
  If IsDocDenied(o) Then
    MsgBox "Не разрешен доступ к документам такого типа"
    Exit Sub
  End If

  Dim g  As Object
  Set g = Manager.GetInstanceGUI(o.id)
  If Not g Is Nothing Then
    g.Show GetDocumentMode(o), o, False
  End If
  usedefaut = False
  Refesh = False
End Sub

' убирает из новой инвентаризации объекты, которые уже где-то учтены
Private Sub ExcludeObjects(ByVal MyInvID As String, ByVal ExcludeInvID As String, ByVal ExcludeType As Integer, Optional ByVal Mask As String = "")
 Dim myInv As INV_INV.Application
 Dim eInv As INV_INV.Application
 Dim Obj As INV_OS.Application
 Dim i As Long
 Dim j As Long
 
 On Error Resume Next
 Set myInv = Manager.GetInstanceObject(MyInvID)
 
 Set eInv = Manager.GetInstanceObject(ExcludeInvID)
 
 
 If myInv Is Nothing Then
   Exit Sub
 End If
 
 If ExcludeType < 2 And eInv Is Nothing Then
   Exit Sub
 End If
 
 On Error Resume Next
 For i = 1 To myInv.INVI_OBJ.Count
  Dim crs As ADODB.Recordset
  If ExcludeType = 0 Then
    Set crs = Nothing
    Set crs = Session.GetData("select theos from INVI_OBJ where instanceid='" & eInv.id & "' and theOS='" & myInv.INVI_OBJ.Item(i).TheOS.id & "'")
    
    'For j = 1 To eInv.INVI_OBJ.Count
    If Not crs Is Nothing Then
      If Not crs.EOF Then
        Set myInv.INVI_OBJ.Item(i).TheOS = Nothing
        myInv.INVI_OBJ.Item(i).save
        'Exit For
      End If
    End If
    'Next
  End If
  
  If ExcludeType = 1 Then
    'For j = 1 To eInv.INVI_DONE.Count
     Set crs = Nothing
    Set crs = Session.GetData("select theos from INVI_DONE where instanceid='" & eInv.id & "' and theOS='" & myInv.INVI_OBJ.Item(i).TheOS.id & "'")
    
    'For j = 1 To eInv.INVI_OBJ.Count
    If Not crs Is Nothing Then
      If Not crs.EOF Then
      'If myInv.INVI_OBJ.Item(i).TheOS.id = eInv.INVI_DONE.Item(j).TheOS.id Then
        Set myInv.INVI_OBJ.Item(i).TheOS = Nothing
        myInv.INVI_OBJ.Item(i).save
        'Exit For
      'End If
    End If
  End If
    'Next
  End If
  
  If ExcludeType = 2 Then
      Dim osname As INVOS_INFO
      Set osname = myInv.INVI_OBJ.Item(i).TheOS
      If InStr(osname.Name, Mask) > 0 Then
        Set myInv.INVI_OBJ.Item(i).TheOS = Nothing
        myInv.INVI_OBJ.Item(i).save
      End If
  End If
 Me.Caption = myInv.brief & " Проверка строка:" & i & " из (" & myInv.INVI_OBJ.Count & ")"
 DoEvents
 Next
' For i = 1 To myInv.INVI_OBJ.Count
' If myInv.INVI_OBJ.Item(i).TheOS Is Nothing Then
' myInv.INVI_OBJ.Item(i).save
' End If
' Next
  
again:
 For i = 1 To myInv.INVI_OBJ.Count
  If myInv.INVI_OBJ.Item(i).TheOS Is Nothing Then
    myInv.INVI_OBJ.Delete i
    Me.Caption = myInv.brief & " Исключение строка:" & i & " из (" & myInv.INVI_OBJ.Count & ")"
    DoEvents
    GoTo again
  End If
 Next
 
MsgBox "Исключение объектов из инвентаризации завершено"
End Sub

Public Sub DBMainteice()
  Dim cap As String
  cap = Me.Caption
  On Error Resume Next
  If Session.IsPOSTGRESQL Then
    Me.Caption = "Сборка мусора"
    DoEvents
    Session.GetData ("vacuum full")
    Me.Caption = "Переиндексация"
    DoEvents
    
    Session.GetData ("reindex database test")
    Me.Caption = "Сбор статистики"
    DoEvents
    Session.GetData ("analize")
  End If
  
  Me.Caption = "Очистка лога"
    DoEvents
  Session.GetData ("truncate table syslog")
   
  On Error GoTo bye
  Me.Caption = "Разблокировка"
  DoEvents
  Dim v As NamedValues
  Set v = New NamedValues
  Call Session.Exec("AdminUnlockAll", v)
  Call SaveSetting("ABOL", "INVENTORY", "DBMAINTAIN", Date + 7)
  Me.Caption = cap
  Exit Sub
bye:
  MsgBox Err.Description
  Me.Caption = cap
End Sub
