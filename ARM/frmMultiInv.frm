VERSION 5.00
Object = "{1801C003-859D-471D-BF31-D4428050324B}#2.1#0"; "MTZ_PANEL.ocx"
Begin VB.Form frmMultiInv 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Данные для отчета"
   ClientHeight    =   6585
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   9165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   9165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExport 
      Caption         =   "Экспорт"
      Height          =   375
      Left            =   4200
      TabIndex        =   26
      Top             =   6120
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Caption         =   "Вариант"
      Height          =   1575
      Left            =   120
      TabIndex        =   4
      Top             =   4440
      Width           =   4215
      Begin VB.OptionButton optCompl 
         Caption         =   "По комплектам"
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Value           =   -1  'True
         Width           =   3855
      End
      Begin VB.OptionButton optType 
         Caption         =   "По типам"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Дополнительное условие"
      Height          =   5895
      Left            =   4440
      TabIndex        =   3
      Top             =   120
      Width           =   4575
      Begin VB.CheckBox chkExcludeBroken 
         Caption         =   "Исключить списанные"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   5520
         Width           =   3015
      End
      Begin VB.CheckBox lblinvi_DEF_TheFlow 
         Caption         =   "Этаж:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   150
         TabIndex        =   24
         Top             =   360
         Width           =   3000
      End
      Begin VB.TextBox txtinvi_DEF_TheFlow 
         Height          =   300
         Left            =   120
         MaxLength       =   5
         TabIndex        =   23
         ToolTipText     =   "Этаж"
         Top             =   690
         Width           =   3000
      End
      Begin VB.CheckBox lblinvi_DEF_TheRoom 
         Caption         =   "Комната:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   120
         TabIndex        =   22
         Top             =   1065
         Width           =   3000
      End
      Begin VB.TextBox txtinvi_DEF_TheRoom 
         Height          =   300
         Left            =   120
         MaxLength       =   10
         TabIndex        =   21
         ToolTipText     =   "Комната"
         Top             =   1395
         Width           =   3000
      End
      Begin VB.CheckBox lblinvi_DEF_TheWorkPlace 
         Caption         =   "Рабочее место:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   120
         TabIndex        =   20
         Top             =   1770
         Width           =   3000
      End
      Begin VB.TextBox txtinvi_DEF_TheWorkPlace 
         Height          =   300
         Left            =   120
         MaxLength       =   10
         TabIndex        =   19
         ToolTipText     =   "Рабочее место"
         Top             =   2100
         Width           =   3000
      End
      Begin VB.CheckBox lblinvi_DEF_TheOrg 
         Caption         =   "Юр. лицо:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   120
         TabIndex        =   18
         Top             =   2520
         Width           =   3000
      End
      Begin VB.TextBox txtinvi_DEF_TheOrg 
         Height          =   300
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   17
         ToolTipText     =   "Юр. лицо"
         Top             =   2850
         Width           =   2550
      End
      Begin VB.CheckBox lblinvi_DEF_DIrection 
         Caption         =   "Дирекция:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   120
         TabIndex        =   15
         Top             =   3225
         Width           =   3000
      End
      Begin VB.TextBox txtinvi_DEF_DIrection 
         Height          =   300
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   14
         ToolTipText     =   "Дирекция"
         Top             =   3555
         Width           =   2550
      End
      Begin VB.CheckBox lblinvi_DEF_Uprev 
         Caption         =   "Управление:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   120
         TabIndex        =   12
         Top             =   3930
         Width           =   3000
      End
      Begin VB.TextBox txtinvi_DEF_Uprev 
         Height          =   300
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   11
         ToolTipText     =   "Управление"
         Top             =   4260
         Width           =   2550
      End
      Begin VB.CheckBox lblinvi_DEF_Otdel 
         Caption         =   "Отдел:"
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   120
         TabIndex        =   9
         Top             =   4635
         Width           =   3000
      End
      Begin VB.TextBox txtinvi_DEF_Otdel 
         Height          =   300
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   8
         ToolTipText     =   "Отдел"
         Top             =   4965
         Width           =   2550
      End
      Begin MTZ_PANEL.DropButton cmdinvi_DEF_Otdel 
         Height          =   300
         Left            =   2670
         TabIndex        =   7
         Tag             =   "refopen.ico"
         ToolTipText     =   "Отдел"
         Top             =   4965
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   529
         Caption         =   ""
      End
      Begin MTZ_PANEL.DropButton cmdinvi_DEF_Uprev 
         Height          =   300
         Left            =   2670
         TabIndex        =   10
         Tag             =   "refopen.ico"
         ToolTipText     =   "Управление"
         Top             =   4260
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   529
         Caption         =   ""
      End
      Begin MTZ_PANEL.DropButton cmdinvi_DEF_DIrection 
         Height          =   300
         Left            =   2670
         TabIndex        =   13
         Tag             =   "refopen.ico"
         ToolTipText     =   "Дирекция"
         Top             =   3555
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   529
         Caption         =   ""
      End
      Begin MTZ_PANEL.DropButton cmdinvi_DEF_TheOrg 
         Height          =   300
         Left            =   2670
         TabIndex        =   16
         Tag             =   "refopen.ico"
         ToolTipText     =   "Юр. лицо"
         Top             =   2850
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   529
         Caption         =   ""
      End
   End
   Begin VB.ListBox lstInv 
      Height          =   4110
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   2
      Top             =   120
      Width           =   4215
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7800
      TabIndex        =   1
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   6480
      TabIndex        =   0
      Top             =   6120
      Width           =   1215
   End
End
Attribute VB_Name = "frmMultiInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public col As Collection
Public OK As Boolean
Public ExportData  As Boolean

Private Sub CancelButton_Click()
OK = False
Me.Hide
End Sub

Private Sub cmdExport_Click()
 If lstInv.SelCount > 0 Then
    OK = True
    Me.Hide
    ExportData = True
  End If
End Sub

Private Sub cmdinvi_DEF_TheOrg_Click()
  On Error Resume Next
        Dim id As String, brief As String
        If Manager.GetReferenceDialogEx2("INVD_ORG", id, brief) Then
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
        If Manager.GetReferenceDialogEx2("INVD_DIR", id, brief) Then
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
        If Manager.GetReferenceDialogEx2("INVD_UPR", id, brief) Then
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
        If Manager.GetReferenceDialogEx2("INVD_OTDEL", id, brief) Then
          txtinvi_DEF_Otdel.Tag = Left(id, 38)
          txtinvi_DEF_Otdel = brief
        End If
End Sub

Private Sub Changing()

End Sub

Private Sub Form_Load()
  Dim rs As ADODB.Recordset
  Set rs = Session.GetData( _
    "select instanceid,invi_DEF_theorg,invi_Def_OrderNum,invi_def_startdate from v_autoinvi_def " & _
    " where statusName <> 'Оформляется' order by invi_DEF_theorg,invi_Def_OrderNum,invi_def_startdate" _
  )
  Set col = New Collection
  Dim db As DBuffer
  lstInv.Clear
  If Not rs Is Nothing Then
  While Not rs.EOF
     Set db = New DBuffer
     db.Name = rs!invi_DEF_theorg & " " & rs!invi_Def_OrderNum
     db.id = rs!InstanceID
     lstInv.AddItem db.Name
     col.Add db, db.id
     lstInv.ItemData(lstInv.NewIndex) = col.Count
     rs.MoveNext
  Wend
  End If
End Sub

Private Sub OKButton_Click()
  If lstInv.SelCount > 0 Then
    OK = True
    Me.Hide
    ExportData = False
  End If
End Sub
