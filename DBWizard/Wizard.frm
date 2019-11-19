VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmWizard 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DB Wizard"
   ClientHeight    =   5625
   ClientLeft      =   1965
   ClientTop       =   1815
   ClientWidth     =   7065
   ControlBox      =   0   'False
   Icon            =   "Wizard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   7065
   Tag             =   "10"
   Begin VB.Frame Frame 
      Caption         =   "Создание базы данных"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Index           =   1
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   7095
      Begin VB.Frame frameRight 
         Caption         =   "Параметры подключения"
         Height          =   3825
         Index           =   1
         Left            =   1455
         TabIndex        =   7
         Top             =   720
         Width           =   4155
         Begin VB.CommandButton cmdTest 
            Caption         =   "Тест"
            Height          =   375
            Index           =   1
            Left            =   960
            TabIndex        =   27
            Top             =   3240
            Width           =   2055
         End
         Begin VB.TextBox txtServer 
            Height          =   285
            Index           =   1
            Left            =   180
            TabIndex        =   13
            Top             =   480
            Width           =   3855
         End
         Begin VB.TextBox txtLogin 
            Height          =   285
            Index           =   1
            Left            =   180
            TabIndex        =   21
            Top             =   2100
            Width           =   3855
         End
         Begin VB.TextBox txtPassword 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Index           =   1
            Left            =   180
            PasswordChar    =   "*"
            TabIndex        =   25
            Top             =   2745
            Width           =   3855
         End
         Begin VB.TextBox txtDatabase 
            Height          =   285
            Index           =   1
            Left            =   180
            TabIndex        =   17
            Top             =   1125
            Width           =   3855
         End
         Begin VB.CheckBox chkIntegrated 
            Caption         =   "Интегрированная NT безопасность"
            Height          =   255
            Index           =   1
            Left            =   180
            TabIndex        =   19
            Top             =   1560
            Width           =   3855
         End
         Begin VB.Label lblServer 
            Caption         =   "SQL сервер:"
            Height          =   255
            Index           =   1
            Left            =   180
            TabIndex        =   11
            Top             =   240
            Width           =   3855
         End
         Begin VB.Label lblLogin 
            Caption         =   "SQL имя пользователя:"
            Height          =   255
            Index           =   1
            Left            =   180
            TabIndex        =   20
            Top             =   1890
            Width           =   3855
         End
         Begin VB.Label lblPassword 
            Caption         =   "SQL пароль:"
            Height          =   255
            Index           =   1
            Left            =   180
            TabIndex        =   23
            Top             =   2520
            Width           =   3855
         End
         Begin VB.Label lblDatabase 
            Caption         =   "База данных:"
            Height          =   255
            Index           =   1
            Left            =   180
            TabIndex        =   15
            Top             =   878
            Width           =   3855
         End
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "На этой вкладке задаются параметры НОВОЙ базы данных."
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   240
         TabIndex        =   30
         Top             =   240
         Width           =   6375
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Готово"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Index           =   4
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   7095
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Создание и загрузка базы данных завершена."
         Height          =   1095
         Left            =   720
         TabIndex        =   29
         Top             =   1440
         Width           =   5775
      End
      Begin VB.Label Label1 
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   18
         Top             =   2280
         Width           =   6615
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame Frame 
      Caption         =   "Информация об ошибках"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Index           =   3
      Left            =   2400
      TabIndex        =   12
      Top             =   0
      Width           =   7095
      Begin RichTextLib.RichTextBox txtErr 
         Height          =   4455
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   7858
         _Version        =   393217
         TextRTF         =   $"Wizard.frx":0442
      End
      Begin VB.Label Label1 
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   14
         Top             =   2280
         Width           =   6615
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Процесс установки"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Index           =   2
      Left            =   1320
      TabIndex        =   8
      Top             =   0
      Width           =   7095
      Begin MSComctlLib.ProgressBar pb 
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   2040
         Visible         =   0   'False
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.ProgressBar PBTop 
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Visible         =   0   'False
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblPass 
         Height          =   375
         Left            =   240
         TabIndex        =   26
         Top             =   600
         Width           =   5295
      End
      Begin VB.Label lblLines 
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   1560
         Width           =   5535
      End
      Begin VB.Label Label1 
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   2280
         Width           =   6615
      End
   End
   Begin VB.PictureBox picNav 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   7065
      TabIndex        =   0
      Top             =   5055
      Width           =   7065
      Begin VB.CommandButton cmdNav 
         Caption         =   "О программе"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Index           =   0
         Left            =   108
         MaskColor       =   &H00000000&
         TabIndex        =   5
         Tag             =   "100"
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton cmdNav 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Index           =   1
         Left            =   2280
         MaskColor       =   &H00000000&
         TabIndex        =   4
         Tag             =   "101"
         Top             =   120
         Width           =   1092
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "< &Back"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Index           =   2
         Left            =   3435
         MaskColor       =   &H00000000&
         TabIndex        =   3
         Tag             =   "102"
         Top             =   120
         Width           =   1092
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "&Next >"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Index           =   3
         Left            =   4560
         MaskColor       =   &H00000000&
         TabIndex        =   1
         Tag             =   "103"
         Top             =   120
         Width           =   1092
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "&Finish"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Index           =   4
         Left            =   5910
         MaskColor       =   &H00000000&
         TabIndex        =   2
         Tag             =   "104"
         Top             =   120
         Width           =   1092
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   108
         X2              =   7012
         Y1              =   24
         Y2              =   24
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   108
         X2              =   7012
         Y1              =   0
         Y2              =   0
      End
   End
End
Attribute VB_Name = "frmWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const NUM_STEPS = 5

Const BTN_HELP = 0
Const BTN_CANCEL = 1
Const BTN_BACK = 2
Const BTN_NEXT = 3
Const BTN_FINISH = 4

Const STEP_INTRO = 0
Const STEP_1 = 1
Const STEP_2 = 2
Const STEP_3 = 3
Const STEP_4 = 4
Const STEP_FINISH = 5

Const DIR_NONE = 0
Const DIR_BACK = 1
Const DIR_NEXT = 2


'module level vars
Dim mnCurStep       As Integer
Dim TestCore As Boolean
Dim testRef As Boolean
Dim ds As MDataSource
Private xdom As MSXML2.DOMDocument
Private e As MSXML2.IXMLDOMElement
Private GenResp As MTZGenerator.Response
Private GenPrj As MTZGenerator.ProjectHolder
Private PrevCFG As String
Dim txtLog As String

Dim mbFinishOK      As Boolean

Dim m As MTZManager.Main
Dim s As MTZSession.Session
Dim o As Object
Dim u As Object
Dim rs As ADODB.RecordSet
Dim site As String
Private Declare Function lstrlenW Lib "kernel32" _
    (ByVal lpString As Long) As Long
Private Declare Function SHGetFolderPath Lib "shfolder.dll" _
    Alias "SHGetFolderPathA" _
    (ByVal hwndOwner As Long, _
    ByVal nFolder As Long, _
    ByVal hToken As Long, _
    ByVal dwReserved As Long, _
    ByVal lpszPath As String) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Sub LoadObjects(start As Integer)
On Error Resume Next
Dim xdom As MSXML2.DOMDocument

Dim path As String

Dim drs As Object, ID As String, typename As String, name As String
Dim i As Long
i = start
path = Dir(App.path & "\LOAD" & "\*.xml")
While path <> ""
    Set xdom = New MSXML2.DOMDocument
    xdom.Load App.path & "\LOAD" & "\" & path
    ID = xdom.lastChild.firstChild.Attributes.getNamedItem("ID").nodeValue
    typename = xdom.lastChild.firstChild.Attributes.getNamedItem("TYPENAME").nodeValue
    name = typename
    
    'try if new format
    name = xdom.lastChild.firstChild.Attributes.getNamedItem("NAME").nodeValue
    i = i + 1
    lblLines.Caption = CStr(i) & ": Загрузка " + typename + " (" + path + ")"
    DoEvents
    Set drs = m.GetInstanceObject(ID)
    If drs Is Nothing Then
      m.NewInstance ID, typename, name
    End If
    If UCase(drs.ID) <> "{88DEEBA4-69B1-454A-992A-FAE3CEBFBCA1}" Then
      Set drs = m.GetInstanceObject(ID)
      If Not drs Is Nothing Then
        drs.LockResource True
        drs.AutoLoadPart = True
        'drs.WorkOffline = True
'        If chkAppend.Value = vbChecked Then
'          drs.XMLLoad xdom.lastChild, 0
'        Else
          drs.XMLLoad xdom.lastChild, 1
'        End If
        'drs.XMLLoad xdom.lastChild, 0
        drs.WorkOffline = False
        lblLines.Caption = CStr(i) & ": Сохранение " + typename + " (" + path + ")"
        DoEvents
        drs.BatchUpdate
        drs.UnLockResource
      End If
      Set xdom = Nothing
    Else
      lblLines.Caption = CStr(i) & ": Пропуск " + typename
      DoEvents
    End If
    path = Dir
Wend
lblLines.Caption = CStr(i) & " загружено"
End Sub

Private Function LoadMetaModel() As Integer
'MetaModel
Set rs = m.ListInstances(site, "MTZMetaModel")
Dim drs As Object, ID As String
 If Not rs.EOF Then
   ID = rs!InstanceID
 Else
  ID = "{88DEEBA4-69B1-454A-992A-FAE3CEBFBCA1}"
  m.NewInstance ID, "MTZMetaModel", "Спец:Метамодель"
 End If
 Set drs = m.GetInstanceObject(ID)

drs.LockResource True
drs.AutoLoadPart = False
'drs.WorkOffline = True
LoadMetaModel = 0
lblLines.Caption = "Загрузка MetaModel"
DoEvents
On Error Resume Next
  Dim xdom As MSXML2.DOMDocument
  Set xdom = New MSXML2.DOMDocument
  xdom.Load App.path & "\LOAD" & "\{88DEEBA4-69B1-454A-992A-FAE3CEBFBCA1}.xml"
  If xdom.xml <> "" Then
    'Llblines.Caption = "Loading MetaModel"
    DoEvents
'    If chkAppend.Value = vbChecked Then
'      drs.XMLLoad xdom.lastChild, 0
'    Else
      drs.XMLLoad xdom.lastChild, 1
'    End If
    drs.WorkOffline = False
    lblLines.Caption = "Сохранение MetaModel"
    DoEvents
    drs.BatchUpdate
    LoadMetaModel = 1
  End If
  Set xdom = Nothing


End Function


Private Sub chkIntegrated_Click(Index As Integer)
If Index = 0 Then
    TestCore = False
Else
    testRef = False
End If
End Sub

Private Sub cmdNav_Click(Index As Integer)
    Dim nAltStep As Integer
    Dim lHelpTopic As Long
    Dim rc As Long
    
    Select Case Index
        Case BTN_HELP
            Dim fabout As frmAbout
            Set fabout = New frmAbout
            fabout.Show vbModal
            Set fabout = Nothing
        
        Case BTN_CANCEL
            Unload Me
          
        Case BTN_BACK
            'place special cases here to jump
            'to alternate steps
            nAltStep = mnCurStep - 1
            SetStep nAltStep, DIR_BACK
          
        Case BTN_NEXT
            'place special cases here to jump
            'to alternate steps
            nAltStep = mnCurStep + 1
            SetStep nAltStep, DIR_NEXT
          
        Case BTN_FINISH
            'wizard creation code goes here
      
            Unload Me
            
        
    End Select
End Sub

Private Sub cmdTest_Click(Index As Integer)
  If Index = 0 Then
    TestCore = False
  Else
    testRef = False
  End If
  Set ds = New MDataSource
  ds.Server = txtServer(Index)
  If Index = 0 Then
    ds.DataBaseName = txtDatabase(0)
  Else
    ds.DataBaseName = "master"
  End If
  
  ds.UserName = txtLogin(Index)
  ds.Password = txtPassword(Index)
  ds.Integrated = (chkIntegrated(Index).Value = vbChecked)
  If Not ds.ServerLogIn Then
    If Index = 0 Then
      MsgBox "Не удается подключиться к базе данных CORE IMS", vbCritical, "Ошибка"
    Else
      MsgBox "Не удается подключиться к SQL Server", vbCritical, "Ошибка"
    End If
    Set ds = Nothing
    Exit Sub
  Else
    MsgBox "Соединение успешно", vbOKOnly, "Тест соединения"
  End If
  If Index = 0 Then
    TestCore = True
  Else
    testRef = True
  End If
  Set ds = Nothing
End Sub


Private Sub Form_Load()
    Dim i As Integer
    'init all vars
    mbFinishOK = False
    mnCurStep = 1
    
    For i = 1 To NUM_STEPS - 1
      Frame(i).Visible = False
    Next
    
    SetStep 1, DIR_NONE
    
End Sub

Private Sub SetStep(nStep As Integer, nDirection As Integer)
  
    Select Case nStep
        Case STEP_INTRO
           
        Case STEP_1
'           If Not TestCore Then
'                MsgBox "Протестируйте соединение с БД"
'                nStep = mnCurStep
'            End If
        Case STEP_2
           If Not testRef Then
                MsgBox "Протестируйте соединение с БД"
                nStep = mnCurStep
            End If
        Case STEP_3
      
        Case STEP_4
            mbFinishOK = True
      

        
    End Select
    
    'move to new step
    Frame(mnCurStep).Enabled = False
    Frame(nStep).Visible = True
    Frame(nStep).Left = 0
    If nStep <> mnCurStep Then
        Frame(mnCurStep).Visible = False
    End If
    Frame(nStep).Enabled = True
  
    SetCaption nStep
    SetNavBtns nStep
    
    
     Select Case nStep
        Case STEP_INTRO
           
        Case STEP_1
         
        Case STEP_2
          Install
        Case STEP_3
            'Install
        Case STEP_4
            mbFinishOK = True
      

        
    End Select
  
End Sub



Private Sub SetNavBtns(nStep As Integer)
    mnCurStep = nStep
    
    If mnCurStep = 1 Then
        cmdNav(BTN_BACK).Enabled = False
        cmdNav(BTN_NEXT).Enabled = True
    ElseIf mnCurStep = NUM_STEPS - 1 Then
        cmdNav(BTN_NEXT).Enabled = False
        cmdNav(BTN_BACK).Enabled = True
    Else
        cmdNav(BTN_BACK).Enabled = True
        cmdNav(BTN_NEXT).Enabled = True
    End If
    
    If mbFinishOK Then
        cmdNav(BTN_FINISH).Enabled = True
    Else
        cmdNav(BTN_FINISH).Enabled = False
    End If
End Sub


Private Sub Install()
    If MsgBox("Начать создание новой базы данных?", vbQuestion + vbYesNo) = vbYes Then
        cmdNav(BTN_BACK).Enabled = False
        cmdNav(BTN_NEXT).Enabled = False
        cmdNav(BTN_FINISH).Enabled = False
    
        
        lblPass = "Создание базы данных"
        MakeNewDB
        lblPass = "Доп. модификации базы данных"
        MakeUpdate
        lblPass = "Настройка параметров"
        MakeCFG
        lblPass = "Загрузка первичных данных"
        LoadData
        lblPass = "Восстановление параметров"
        RestoreCFG
        lblPass = "Операции завершены. Нажмите кнопку Next."
        
        txtErr.Text = txtLog
        cmdNav(BTN_BACK).Enabled = True
        cmdNav(BTN_NEXT).Enabled = True
        cmdNav(BTN_FINISH).Enabled = True
    
    End If
End Sub


Private Sub SetCaption(nStep As Integer)
    On Error Resume Next

   

End Sub

'=========================================================
'this sub displays an error message when the user has
'not entered enough data to continue
'=========================================================
Sub IncompleteData(nIndex As Integer)
    On Error Resume Next
    Dim sTmp As String
      
    'get the base error messagee
    sTmp = sTmp & vbCrLf
    Beep
    MsgBox sTmp, vbInformation
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Dim rc As Long
  
End Sub


Private Sub Timer1_Timer()
  On Error Resume Next: m.GetSession(site).Exec "SessionTouch", Nothing
End Sub

Private Sub txtDatabase_Change(Index As Integer)
If Index = 0 Then
    TestCore = False
Else
    testRef = False
End If
End Sub

Private Sub txtLogin_Change(Index As Integer)
If Index = 0 Then
    TestCore = False
Else
    testRef = False
End If
End Sub

Private Sub txtPassword_Change(Index As Integer)
If Index = 0 Then
    TestCore = False
Else
    testRef = False
End If
End Sub

Private Sub txtServer_Change(Index As Integer)
If Index = 0 Then
    TestCore = False
Else
    testRef = False
End If
End Sub




Private Sub MakeCore()
  Set ds = New MDataSource
  ds.Server = txtServer(0)
  ds.DataBaseName = txtDatabase(0)
  ds.UserName = txtLogin(0)
  ds.Password = txtPassword(0)
  ds.Integrated = (chkIntegrated(0).Value = vbChecked)
  If Not ds.ServerLogIn Then
    MsgBox "Не удается подключиться к Microsoft SQL Server", vbCritical
    Set ds = Nothing
    Exit Sub
  End If
  
  Set GenResp = New MTZGenerator.Response
  Set GenPrj = GenResp.Project
  GenPrj.Load App.path & "\DB\Core.xml"
  
  On Error Resume Next
  
  
  
  Dim i As Long, j As Long
  Dim blocks As Integer
  For i = 1 To GenPrj.Modules.Count
    blocks = blocks + GenPrj.Modules.Item(i).blocks.Count
    
  Next
  PBTop.Min = 0
  PBTop.Value = 0
  PBTop.Max = blocks + 1
  PBTop.Visible = True
  Dim k As Long
  k = 0
  For i = 1 To GenPrj.Modules.Count
    For j = 1 To GenPrj.Modules.Item(i).blocks.Count
      
      execBlock GenPrj.Modules.Item(i).blocks.Item(j), GenPrj.Modules.Item(i).modulename
     
      k = k + 1
      pb.Value = k
    Next
  Next
  PBTop.Visible = False

  lblLines = ""
    
  Set ds = Nothing
  Set GenResp = Nothing
  Set GenPrj = Nothing
End Sub


Private Sub MakeNewDB()
  Set ds = New MDataSource
  ds.Server = txtServer(1)
  ds.DataBaseName = "master"
  ds.UserName = txtLogin(1)
  ds.Password = txtPassword(1)
  ds.Integrated = (chkIntegrated(1).Value = vbChecked)
  If Not ds.ServerLogIn Then
    MsgBox "Не удается подключиться к Microsoft SQL Server", vbCritical
    Set ds = Nothing
    Exit Sub
  End If
  
  Set GenResp = New MTZGenerator.Response
  Set GenPrj = GenResp.Project
  GenPrj.Load App.path & "\db\all.xml"
  
  On Error Resume Next
  
  ds.Execute ("create database [" & txtDatabase(1).Text & "] COLLATE Cyrillic_General_CI_AS")
  
  If Not ds.Execute("use [" & txtDatabase(1).Text & "]") Then
    MsgBox "Не удается создать базу данных", vbCritical
    Set ds = Nothing
    Set GenResp = Nothing
    Set GenPrj = Nothing
    Exit Sub
  End If
  
  
  Dim i As Long, j As Long
  Dim blocks As Integer
  For i = 1 To GenPrj.Modules.Count
    blocks = blocks + GenPrj.Modules.Item(i).blocks.Count
    
  Next
  PBTop.Min = 0
  PBTop.Value = 0
  PBTop.Max = blocks + 1
  PBTop.Visible = True
  Dim k As Long
  k = 0
  For i = 1 To GenPrj.Modules.Count
    For j = 1 To GenPrj.Modules.Item(i).blocks.Count
   
      execBlock GenPrj.Modules.Item(i).blocks.Item(j), GenPrj.Modules.Item(i).modulename
  
      k = k + 1
       PBTop.Value = k
    Next
  Next
  PBTop.Visible = False
  lblLines = ""
  
  Set ds = Nothing
  Set GenResp = Nothing
  Set GenPrj = Nothing
End Sub


Private Sub MakeUpdate()
  Set ds = New MDataSource
  ds.Server = txtServer(1)
  ds.DataBaseName = "master"
  ds.UserName = txtLogin(1)
  ds.Password = txtPassword(1)
  ds.Integrated = (chkIntegrated(1).Value = vbChecked)
  If Not ds.ServerLogIn Then
    MsgBox "Не удается подключиться к Microsoft SQL Server", vbCritical
    Set ds = Nothing
    Exit Sub
  End If
  
  Set GenResp = New MTZGenerator.Response
  Set GenPrj = GenResp.Project
  GenPrj.Load App.path & "\db\update.xml"
  
  On Error Resume Next
  
  
  If Not ds.Execute("use [" & txtDatabase(1).Text & "]") Then
    MsgBox "Не удается обновить базу данных", vbCritical
    Set ds = Nothing
    Set GenResp = Nothing
    Set GenPrj = Nothing
    Exit Sub
  End If
  
  
  Dim i As Long, j As Long
  Dim blocks As Integer
  For i = 1 To GenPrj.Modules.Count
    blocks = blocks + GenPrj.Modules.Item(i).blocks.Count
    
  Next
  PBTop.Min = 0
  PBTop.Value = 0
  PBTop.Max = blocks + 1
  PBTop.Visible = True
  Dim k As Long
  k = 0
  For i = 1 To GenPrj.Modules.Count
    For j = 1 To GenPrj.Modules.Item(i).blocks.Count
   
      execBlock GenPrj.Modules.Item(i).blocks.Item(j), GenPrj.Modules.Item(i).modulename
  
      k = k + 1
       PBTop.Value = k
    Next
  Next
  PBTop.Visible = False
  lblLines = ""
  
  Set ds = Nothing
  Set GenResp = Nothing
  Set GenPrj = Nothing
End Sub

Private Sub execBlock(b As BlockHolder, modulename As String)
Dim s As String, lines() As String, i As Long
lines = Split(b.BlockCode, vbCrLf, , vbTextCompare)
s = ""
pb.Min = LBound(lines)
pb.Max = UBound(lines)
pb.Value = LBound(lines)
pb.Visible = True
For i = LBound(lines) To UBound(lines)
  lblLines.Caption = modulename & ". Строка " & i & " из " & UBound(lines)
  pb.Value = i
  If UCase(Trim(lines(i))) = "GO" Then
   On Error GoTo err1
   If Trim(s) <> "" Then
   ds.Execute s
   DoEvents
   End If
   s = ""
   GoTo cont
err1:
  txtLog = txtLog & vbCrLf & b.BlockName & ":" & modulename & vbCrLf & s & vbCrLf & "------------------------" & vbCrLf & Err.Description
  Debug.Print Err.Number, Err.Description
  Resume err2
err2:
   s = ""
  Else
    s = s & vbCrLf & lines(i)
  End If
cont:
Next
pb.Visible = False


End Sub

Private Sub LoadData()
  
  
  Set m = New MTZManager.Main
  Set s = m.GetSession("LOADCFG")
  If s Is Nothing Then
    Exit Sub
  End If
  s.Login "supervisor", "bami"
  If s.sessionid = "" Then
    Exit Sub
  End If

  Timer1.Enabled = True
  
  LoadObjects LoadMetaModel

  Timer1.Enabled = False
  m.GetSession(site).Logout
  Set m = Nothing

  lblLines = ""
End Sub

Private Sub MakeCFG()
    Dim frf As Integer
    Dim s As String
    frf = FreeFile
    Open App.path & "\LOADCFG\cfg.xml" For Input As #frf
    s = input(LOF(frf), frf)
    Close #frf
    
    s = Replace(s, "%SERVER%", txtServer(1))
    s = Replace(s, "%DB%", txtDatabase(1))
    
    s = Replace(s, "%USER%", txtLogin(1))
    s = Replace(s, "%PASSWORD%", txtPassword(1))
    
    If chkIntegrated(1).Value = vbChecked Then
        s = Replace(s, "%INTEGRATED%", "-1")
    Else
        s = Replace(s, "%INTEGRATED%", "0")
    End If
    
    
    Dim ttt As String
    ttt = GetDocFolder()
    
    s = Replace(s, "%IMAGES%", "c:\program files\sgs\sgs\images\")
    s = Replace(s, "%LAYOUTS%", ttt)
    s = Replace(s, "%TEMP%", GetUserTempPath())
    
    If GetSetting("MTZ", "CONFIG", "IMAGEPATH", "") = "" Then
     Call SaveSetting("MTZ", "CONFIG", "IMAGEPATH", "c:\program files\sgs\sgs\images\")
    End If
    If GetSetting("MTZ", "CONFIG", "LAYOUTS", "") = "" Then
        Call SaveSetting("MTZ", "CONFIG", "LAYOUTS", ttt)
    End If
    If GetSetting("MTZ", "CONFIG", "TEMPPATH", "") = "" Then
        Call SaveSetting("MTZ", "CONFIG", "TEMPPATH", GetUserTempPath())
    End If
    
    frf = FreeFile
    
    
    Open ttt & "\dbcfg.xml" For Output As #frf
    Print #frf, s
    Close #frf
    
    PrevCFG = MTZGetSetting("MTZ", "CONFIG", "XMLPATH", "")
    Call SaveSetting("MTZ", "CONFIG", "XMLPATH", ttt & "\dbcfg.xml")
    On Error Resume Next
    Call MTZSaveSetting("MTZ", "CONFIG", "XMLPATH", ttt & "\dbcfg.xml")
    
End Sub

Private Sub RestoreCFG()
     On Error Resume Next
     If PrevCFG <> "" Then
        Call SaveSetting("MTZ", "CONFIG", "XMLPATH", PrevCFG)
        Call MTZSaveSetting("MTZ", "CONFIG", "XMLPATH", PrevCFG)
        'Kill App.path & "\LOADCFG\dbcfg.xml"
     End If
 End Sub




Private Function GetDocFolder()
    Dim buff As String
    'Allocate buffer space
    buff = Space$(260)
    If SHGetFolderPath(Me.hWnd, &H5, -1, &H0, buff) = 0 Then
        GetDocFolder = TrimNull(buff)
    Else
        GetDocFolder = Environ("USERPROFILE")
    End If
End Function

Private Function TrimNull(startstr As String) As String
    TrimNull = Left$(startstr, lstrlenW(StrPtr(startstr)))
End Function





Public Function GetUserTempPath() As String
  Dim sTempPath As String
  sTempPath = Space(1024)
  sTempPath = Replace(sTempPath, " ", "\")
  Call GetTempPath(Len(sTempPath), sTempPath)
  GetUserTempPath = TrimNull(sTempPath)
End Function

