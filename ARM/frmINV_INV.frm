VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{1801C003-859D-471D-BF31-D4428050324B}#2.1#0"; "MTZ_PANEL.ocx"
Begin VB.Form frmINV_INV 
   Caption         =   "Фильтр для док. Инвентаризация"
   ClientHeight    =   7695
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9780
   LinkTopic       =   "Form1"
   ScaleHeight     =   7695
   ScaleWidth      =   9780
   StartUpPosition =   3  'Windows Default
   Tag             =   "Card.ICO"
   Begin MSComctlLib.TabStrip ts 
      Height          =   1500
      Left            =   6960
      TabIndex        =   2
      Top             =   240
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   2646
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Отмена"
      Height          =   330
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "Отказ от задания фильтра"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   750
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   330
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Применить фильтр"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   750
   End
   Begin MTZ_PANEL.ScrolledWindow PanelfGroup 
      Height          =   5925
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   10451
      Begin VB.TextBox txtinvi_DEF_Info 
         Height          =   1200
         Left            =   3450
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   42
         ToolTipText     =   "Примечания"
         Top             =   4635
         Width           =   3000
      End
      Begin VB.CheckBox lblinvi_DEF_Info 
         Caption         =   "Примечания:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   3450
         TabIndex        =   41
         Top             =   4305
         Width           =   3000
      End
      Begin MTZ_PANEL.DropButton cmdinvi_DEF_MatOtv 
         Height          =   300
         Left            =   6000
         TabIndex        =   40
         Tag             =   "refopen.ico"
         ToolTipText     =   "МОЛ"
         Top             =   3930
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   529
         Caption         =   ""
      End
      Begin VB.TextBox txtinvi_DEF_MatOtv 
         Height          =   300
         Left            =   3450
         Locked          =   -1  'True
         TabIndex        =   39
         ToolTipText     =   "МОЛ"
         Top             =   3930
         Width           =   2550
      End
      Begin VB.CheckBox lblinvi_DEF_MatOtv 
         Caption         =   "МОЛ:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   3450
         TabIndex        =   38
         Top             =   3600
         Width           =   3000
      End
      Begin VB.TextBox txtinvi_DEF_TheWorkPlace 
         Height          =   300
         Left            =   3450
         MaxLength       =   10
         TabIndex        =   37
         ToolTipText     =   "Рабочее место"
         Top             =   3225
         Width           =   3000
      End
      Begin VB.CheckBox lblinvi_DEF_TheWorkPlace 
         Caption         =   "Рабочее место:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   3450
         TabIndex        =   36
         Top             =   2895
         Width           =   3000
      End
      Begin VB.TextBox txtinvi_DEF_TheRoom 
         Height          =   300
         Left            =   3450
         MaxLength       =   10
         TabIndex        =   35
         ToolTipText     =   "Комната"
         Top             =   2520
         Width           =   3000
      End
      Begin VB.CheckBox lblinvi_DEF_TheRoom 
         Caption         =   "Комната:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   3450
         TabIndex        =   34
         Top             =   2190
         Width           =   3000
      End
      Begin VB.TextBox txtinvi_DEF_TheFlow 
         Height          =   300
         Left            =   3450
         MaxLength       =   5
         TabIndex        =   33
         ToolTipText     =   "Этаж"
         Top             =   1815
         Width           =   3000
      End
      Begin VB.CheckBox lblinvi_DEF_TheFlow 
         Caption         =   "Этаж:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   3480
         TabIndex        =   32
         Top             =   1485
         Width           =   3000
      End
      Begin MTZ_PANEL.DropButton cmdinvi_DEF_Building 
         Height          =   300
         Left            =   6000
         TabIndex        =   31
         Tag             =   "refopen.ico"
         ToolTipText     =   "Здание"
         Top             =   1110
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   529
         Caption         =   ""
      End
      Begin VB.TextBox txtinvi_DEF_Building 
         Height          =   300
         Left            =   3450
         Locked          =   -1  'True
         TabIndex        =   30
         ToolTipText     =   "Здание"
         Top             =   1110
         Width           =   2550
      End
      Begin VB.CheckBox lblinvi_DEF_Building 
         Caption         =   "Здание:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   3450
         TabIndex        =   29
         Top             =   780
         Width           =   3000
      End
      Begin MTZ_PANEL.DropButton cmdinvi_DEF_TheOwner 
         Height          =   300
         Left            =   6000
         TabIndex        =   28
         Tag             =   "refopen.ico"
         ToolTipText     =   "Владелец"
         Top             =   405
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   529
         Caption         =   ""
      End
      Begin VB.TextBox txtinvi_DEF_TheOwner 
         Height          =   300
         Left            =   3450
         Locked          =   -1  'True
         TabIndex        =   27
         ToolTipText     =   "Владелец"
         Top             =   405
         Width           =   2550
      End
      Begin VB.CheckBox lblinvi_DEF_TheOwner 
         Caption         =   "Владелец:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   3450
         TabIndex        =   26
         Top             =   75
         Width           =   3000
      End
      Begin MTZ_PANEL.DropButton cmdinvi_DEF_Otdel 
         Height          =   300
         Left            =   2850
         TabIndex        =   25
         Tag             =   "refopen.ico"
         ToolTipText     =   "Отдел"
         Top             =   6045
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   529
         Caption         =   ""
      End
      Begin VB.TextBox txtinvi_DEF_Otdel 
         Height          =   300
         Left            =   300
         Locked          =   -1  'True
         TabIndex        =   24
         ToolTipText     =   "Отдел"
         Top             =   6045
         Width           =   2550
      End
      Begin VB.CheckBox lblinvi_DEF_Otdel 
         Caption         =   "Отдел:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   300
         TabIndex        =   23
         Top             =   5715
         Width           =   3000
      End
      Begin MTZ_PANEL.DropButton cmdinvi_DEF_Uprev 
         Height          =   300
         Left            =   2850
         TabIndex        =   22
         Tag             =   "refopen.ico"
         ToolTipText     =   "Управление"
         Top             =   5340
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   529
         Caption         =   ""
      End
      Begin VB.TextBox txtinvi_DEF_Uprev 
         Height          =   300
         Left            =   300
         Locked          =   -1  'True
         TabIndex        =   21
         ToolTipText     =   "Управление"
         Top             =   5340
         Width           =   2550
      End
      Begin VB.CheckBox lblinvi_DEF_Uprev 
         Caption         =   "Управление:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   300
         TabIndex        =   20
         Top             =   5010
         Width           =   3000
      End
      Begin MTZ_PANEL.DropButton cmdinvi_DEF_DIrection 
         Height          =   300
         Left            =   2850
         TabIndex        =   19
         Tag             =   "refopen.ico"
         ToolTipText     =   "Дирекция"
         Top             =   4635
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   529
         Caption         =   ""
      End
      Begin VB.TextBox txtinvi_DEF_DIrection 
         Height          =   300
         Left            =   300
         Locked          =   -1  'True
         TabIndex        =   18
         ToolTipText     =   "Дирекция"
         Top             =   4635
         Width           =   2550
      End
      Begin VB.CheckBox lblinvi_DEF_DIrection 
         Caption         =   "Дирекция:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   300
         TabIndex        =   17
         Top             =   4305
         Width           =   3000
      End
      Begin MTZ_PANEL.DropButton cmdinvi_DEF_TheOrg 
         Height          =   300
         Left            =   2850
         TabIndex        =   16
         Tag             =   "refopen.ico"
         ToolTipText     =   "Юр. лицо"
         Top             =   3930
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   529
         Caption         =   ""
      End
      Begin VB.TextBox txtinvi_DEF_TheOrg 
         Height          =   300
         Left            =   300
         Locked          =   -1  'True
         TabIndex        =   15
         ToolTipText     =   "Юр. лицо"
         Top             =   3930
         Width           =   2550
      End
      Begin VB.CheckBox lblinvi_DEF_TheOrg 
         Caption         =   "Юр. лицо:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   300
         TabIndex        =   14
         Top             =   3600
         Width           =   3000
      End
      Begin MSComCtl2.DTPicker dtpinvi_DEF_EndDate_LE 
         Height          =   300
         Left            =   300
         TabIndex        =   13
         ToolTipText     =   "Дата завершения инвентраизации по"
         Top             =   3225
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   107216899
         CurrentDate     =   40148
      End
      Begin VB.CheckBox lblinvi_DEF_EndDate_LE 
         Caption         =   "Дата завершения инвентаризации по:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   300
         TabIndex        =   12
         Top             =   2895
         Width           =   3000
      End
      Begin MSComCtl2.DTPicker dtpinvi_DEF_EndDate_GE 
         Height          =   300
         Left            =   300
         TabIndex        =   11
         ToolTipText     =   "Дата завершения инвентраизации C"
         Top             =   2520
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   107216899
         CurrentDate     =   40148
      End
      Begin VB.CheckBox lblinvi_DEF_EndDate_GE 
         Caption         =   "Дата завершения инвентаризации C:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   300
         TabIndex        =   10
         Top             =   2190
         Width           =   3000
      End
      Begin MSComCtl2.DTPicker dtpinvi_DEF_StartDate_LE 
         Height          =   300
         Left            =   300
         TabIndex        =   9
         ToolTipText     =   "Дата начала инвентаризации по"
         Top             =   1815
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   107216899
         CurrentDate     =   40148
      End
      Begin VB.CheckBox lblinvi_DEF_StartDate_LE 
         Caption         =   "Дата начала инвентаризации по:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   300
         TabIndex        =   8
         Top             =   1485
         Width           =   3000
      End
      Begin MSComCtl2.DTPicker dtpinvi_DEF_StartDate_GE 
         Height          =   300
         Left            =   300
         TabIndex        =   7
         ToolTipText     =   "Дата начала инвентаризации C"
         Top             =   1110
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   107216899
         CurrentDate     =   40148
      End
      Begin VB.CheckBox lblinvi_DEF_StartDate_GE 
         Caption         =   "Дата начала инвентаризации C:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   300
         TabIndex        =   6
         Top             =   780
         Width           =   3000
      End
      Begin VB.TextBox txtinvi_DEF_OrderNum 
         Height          =   300
         Left            =   300
         MaxLength       =   30
         TabIndex        =   5
         ToolTipText     =   "Номер приказа"
         Top             =   405
         Width           =   3000
      End
      Begin VB.CheckBox lblinvi_DEF_OrderNum 
         Caption         =   "Номер приказа:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   300
         TabIndex        =   4
         Top             =   75
         Width           =   3000
      End
   End
   Begin VB.Menu mnuCtl 
      Caption         =   "mnuCtl"
      Visible         =   0   'False
      Begin VB.Menu mnuSetup 
         Caption         =   "Настройка"
      End
   End
End
Attribute VB_Name = "frmINV_INV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Item As Object
Public OK As Boolean
Private OnInit As Boolean
Public Event Changed()
Private TSCustom As MTZ_CUSTOMTAB.TabStripCustomizer







Private Sub cmdOK_Click()
    On Error Resume Next
    OK = True
    Me.Hide
End Sub
Private Sub cmdCancel_Click()
    On Error Resume Next
    OK = False
    Me.Hide
End Sub
Public Sub Init(ObjItem As Object)
 Set Item = ObjItem
 If Item Is Nothing Then Set Item = MyUser.Application
 TInit
End Sub
Private Sub Form_Load()
  On Error Resume Next
  Dim ff As Long, buf As String
  LoadFromSkin Me
  ts.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - cmdok.Height
  cmdok.Move Me.ScaleWidth - 110 * Screen.TwipsPerPixelX, Me.ScaleHeight - cmdok.Height, cmdok.Width, cmdok.Height
  cmdcancel.Move Me.ScaleWidth - 55 * Screen.TwipsPerPixelX, Me.ScaleHeight - cmdok.Height, cmdcancel.Width, cmdok.Height
  If Item Is Nothing Then Init MyUser.Application
End Sub
Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  Set Item = Nothing
  Set TSCustom = Nothing
  SaveToSkin Me
  Exit Sub
bye:
  MsgBox Err.Description, vbOKOnly
  Cancel = -1
End Sub
Private Sub Form_Resize()
 If Me.WindowState = 1 Then Exit Sub
 On Error Resume Next
  ts.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - cmdok.Height
  cmdok.Move Me.ScaleWidth - 110 * Screen.TwipsPerPixelX, Me.ScaleHeight - cmdok.Height, cmdok.Width, cmdok.Height
  cmdcancel.Move Me.ScaleWidth - 55 * Screen.TwipsPerPixelX, Me.ScaleHeight - cmdok.Height, cmdcancel.Width, cmdok.Height
  ts_click
End Sub
Private Sub mnuSetup_Click()
TSCustom.Setup ts
End Sub
Private Sub ts_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 And Shift = 0 Then
    PopupMenu mnuCtl
  End If
End Sub
Private Sub ts_click()
  On Error Resume Next
  panelfGroup.Visible = False

   Select Case ts.SelectedItem.Key
   Case "fGroup"
     With panelfGroup
     .Top = ts.ClientTop
     .Left = ts.ClientLeft
     .Width = ts.ClientWidth
     .Height = ts.ClientHeight
     .Visible = True
     .ZOrder 0
     End With
     End Select
End Sub
Private Sub TInit()
  On Error Resume Next
  Dim ff As Long, buf As String

ts.Tabs.Item(1).Caption = "Описание"
ts.Tabs.Item(1).Key = "fGroup"
PanelfGroupInit
  Set TSCustom = New MTZ_CUSTOMTAB.TabStripCustomizer
  TSCustom.Init ts, "INV_INV", "fctlINV_INV"
  TSCustom.SetupFromRegistry ts
 ts_click
End Sub


Private Sub Changing()
If OnInit Then Exit Sub
 RaiseEvent Changed
End Sub
Private Sub txtinvi_DEF_OrderNum_Change()
  Changing
End Sub
Private Sub dtpinvi_DEF_StartDate_GE_Change()
  Changing
End Sub
Private Sub dtpinvi_DEF_StartDate_LE_Change()
  Changing
End Sub
Private Sub dtpinvi_DEF_EndDate_GE_Change()
  Changing
End Sub
Private Sub dtpinvi_DEF_EndDate_LE_Change()
  Changing
End Sub
Private Sub txtinvi_DEF_TheOrg_Change()
  Changing
End Sub
Private Sub cmdinvi_DEF_TheOrg_Click()
  On Error Resume Next
        Dim id As String, brief As String
        If Item.Application.Manager.GetReferenceDialogEx2("INVD_ORG", id, brief) Then
          txtinvi_DEF_TheOrg.Tag = Left(id, 38)
          txtinvi_DEF_TheOrg = brief
        End If
End Sub
Private Sub txtinvi_DEF_DIrection_Change()
  Changing
End Sub
Private Sub cmdinvi_DEF_DIrection_Click()
  On Error Resume Next
        Dim id As String, brief As String
        If Item.Application.Manager.GetReferenceDialogEx2("INVD_DIR", id, brief) Then
          txtinvi_DEF_DIrection.Tag = Left(id, 38)
          txtinvi_DEF_DIrection = brief
        End If
End Sub
Private Sub txtinvi_DEF_Uprev_Change()
  Changing
End Sub
Private Sub cmdinvi_DEF_Uprev_Click()
  On Error Resume Next
        Dim id As String, brief As String
        If Item.Application.Manager.GetReferenceDialogEx2("INVD_UPR", id, brief) Then
          txtinvi_DEF_Uprev.Tag = Left(id, 38)
          txtinvi_DEF_Uprev = brief
        End If
End Sub
Private Sub txtinvi_DEF_Otdel_Change()
  Changing
End Sub
Private Sub cmdinvi_DEF_Otdel_Click()
  On Error Resume Next
        Dim id As String, brief As String
        If Item.Application.Manager.GetReferenceDialogEx2("INVD_OTDEL", id, brief) Then
          txtinvi_DEF_Otdel.Tag = Left(id, 38)
          txtinvi_DEF_Otdel = brief
        End If
End Sub
Private Sub txtinvi_DEF_TheOwner_Change()
  Changing
End Sub
Private Sub cmdinvi_DEF_TheOwner_Click()
  On Error Resume Next
        Dim id As String, brief As String
        If Item.Application.Manager.GetReferenceDialogEx2("INVD_OWNER", id, brief) Then
          txtinvi_DEF_TheOwner.Tag = Left(id, 38)
          txtinvi_DEF_TheOwner = brief
        End If
End Sub
Private Sub txtinvi_DEF_Building_Change()
  Changing
End Sub
Private Sub cmdinvi_DEF_Building_Click()
  On Error Resume Next
        Dim id As String, brief As String
        If Item.Application.Manager.GetReferenceDialogEx2("INVD_BLD", id, brief) Then
          txtinvi_DEF_Building.Tag = Left(id, 38)
          txtinvi_DEF_Building = brief
        End If
End Sub
Private Sub txtinvi_DEF_TheFlow_Change()
  Changing
End Sub
Private Sub txtinvi_DEF_TheRoom_Change()
  Changing
End Sub
Private Sub txtinvi_DEF_TheWorkPlace_Change()
  Changing
End Sub
Private Sub txtinvi_DEF_MatOtv_Change()
  Changing
End Sub
Private Sub cmdinvi_DEF_MatOtv_Click()
  On Error Resume Next
        Dim id As String, brief As String
        If Item.Application.Manager.GetReferenceDialogEx2("INVD_OWNER", id, brief) Then
          txtinvi_DEF_MatOtv.Tag = Left(id, 38)
          txtinvi_DEF_MatOtv = brief
        End If
End Sub
Private Sub txtinvi_DEF_Info_Change()
  Changing
End Sub
Private Sub PanelfGroupInit()
OnInit = True
Dim iii As Long ' for combo only

txtinvi_DEF_OrderNum = ""
dtpinvi_DEF_StartDate_GE = Date
dtpinvi_DEF_StartDate_LE = Date
dtpinvi_DEF_EndDate_GE = Date
dtpinvi_DEF_EndDate_LE = Date
  txtinvi_DEF_TheOrg.Tag = ""
  txtinvi_DEF_TheOrg = ""
 LoadBtnPictures cmdinvi_DEF_TheOrg, cmdinvi_DEF_TheOrg.Tag
  cmdinvi_DEF_TheOrg.RemoveAllMenu
  txtinvi_DEF_DIrection.Tag = ""
  txtinvi_DEF_DIrection = ""
 LoadBtnPictures cmdinvi_DEF_DIrection, cmdinvi_DEF_DIrection.Tag
  cmdinvi_DEF_DIrection.RemoveAllMenu
  txtinvi_DEF_Uprev.Tag = ""
  txtinvi_DEF_Uprev = ""
 LoadBtnPictures cmdinvi_DEF_Uprev, cmdinvi_DEF_Uprev.Tag
  cmdinvi_DEF_Uprev.RemoveAllMenu
  txtinvi_DEF_Otdel.Tag = ""
  txtinvi_DEF_Otdel = ""
 LoadBtnPictures cmdinvi_DEF_Otdel, cmdinvi_DEF_Otdel.Tag
  cmdinvi_DEF_Otdel.RemoveAllMenu
  txtinvi_DEF_TheOwner.Tag = ""
  txtinvi_DEF_TheOwner = ""
 LoadBtnPictures cmdinvi_DEF_TheOwner, cmdinvi_DEF_TheOwner.Tag
  cmdinvi_DEF_TheOwner.RemoveAllMenu
  txtinvi_DEF_Building.Tag = ""
  txtinvi_DEF_Building = ""
 LoadBtnPictures cmdinvi_DEF_Building, cmdinvi_DEF_Building.Tag
  cmdinvi_DEF_Building.RemoveAllMenu
txtinvi_DEF_TheFlow = ""
txtinvi_DEF_TheRoom = ""
txtinvi_DEF_TheWorkPlace = ""
  txtinvi_DEF_MatOtv.Tag = ""
  txtinvi_DEF_MatOtv = ""
 LoadBtnPictures cmdinvi_DEF_MatOtv, cmdinvi_DEF_MatOtv.Tag
  cmdinvi_DEF_MatOtv.RemoveAllMenu
OnInit = False
End Sub



