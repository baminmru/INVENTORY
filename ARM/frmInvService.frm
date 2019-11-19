VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInvService 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Сервис инвентаризации"
   ClientHeight    =   7155
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Добавить потерянные проинвентаризированные объекты"
      Height          =   3135
      Left            =   120
      TabIndex        =   10
      Top             =   3360
      Width           =   5295
      Begin VB.CommandButton cmdAddLost 
         Caption         =   "Добавить потерянные проинвентаризированные объекты"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   2520
         Width           =   4695
      End
      Begin MSComCtl2.DTPicker dTo 
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   2040
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   661
         _Version        =   393216
         Format          =   100794369
         CurrentDate     =   40252
      End
      Begin MSComCtl2.DTPicker dFrom 
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   1320
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   661
         _Version        =   393216
         Format          =   100794369
         CurrentDate     =   40252
      End
      Begin VB.CommandButton cmdOrg 
         Caption         =   "..."
         Height          =   375
         Left            =   4200
         TabIndex        =   13
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtOrg 
         Height          =   405
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   600
         Width           =   3615
      End
      Begin VB.Label Label5 
         Caption         =   "По:"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "C:"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lblOrg 
         Caption         =   "Организация:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Исключение объектов"
      Height          =   3135
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5295
      Begin VB.CommandButton cmdOK 
         Caption         =   "Исключить объекты"
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   2520
         Width           =   4455
      End
      Begin VB.ComboBox cmbExclude 
         Height          =   315
         ItemData        =   "frmInvService.frx":0000
         Left            =   360
         List            =   "frmInvService.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   480
         Width           =   4455
      End
      Begin VB.TextBox txtINV 
         Height          =   405
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1200
         Width           =   3735
      End
      Begin VB.CommandButton cmdInvI 
         Caption         =   "..."
         Height          =   375
         Left            =   4320
         TabIndex        =   3
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox txtMask 
         Height          =   285
         Left            =   360
         TabIndex        =   2
         Top             =   2040
         Width           =   4455
      End
      Begin VB.Label Label1 
         Caption         =   "Что исключить:"
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label2 
         Caption         =   "Зарегистрированные в инвентаризации:"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   840
         Width           =   3135
      End
      Begin VB.Label Label3 
         Caption         =   "Маска названия:"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   1800
         Width           =   3735
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   6600
      Width           =   1215
   End
End
Attribute VB_Name = "frmInvService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public OK As Boolean
Public id As String

Private Sub cmdAddLost_Click()
  If txtOrg.Tag = "" Then
    Exit Sub
  End If
  If MsgBox("Добавить потерянные объекты в инвентаризацию ?", vbYesNo, "Подтвердите") = vbYes Then
    Dim f As String
    Dim inv As INV_INV.Application
    Dim cnt As Integer
    Dim cnt2 As Integer
    
    Set inv = Manager.GetInstanceObject(id)
    
    
    f = "select invos_info.invos_infoid,invos_inv.* from invos_info  join  invos_inv on invos_info.instanceid = invos_inv.instanceid " & _
     " where invos_inv.inventory not in (select invi_DEFid from invi_DEF)  and invos_info.theorg='" & txtOrg.Tag & "'"
    
    Dim rs As ADODB.Recordset
    Dim rsc As ADODB.Recordset
    f = f & " and invos_inv.invdate >=" & IIf(Session.IsMSSQL, MakeMSSQLDate(dFrom.Value), IIf(Session.IsORACLE, MakeORACLEDate(dFrom.Value), MakePGSQLDate(dFrom.Value)))
    f = f & " and invos_inv.invdate <=" & IIf(Session.IsMSSQL, MakeMSSQLDate(dTo.Value), IIf(Session.IsORACLE, MakeORACLEDate(dTo.Value), MakePGSQLDate(dTo.Value)))
    Set rs = Session.GetData(f)
    If Not rs Is Nothing Then
      If Not rs.EOF Then
        While Not rs.EOF
             ' проверяем  что такого объекта не было в инвентаризации
              Set rsc = Session.GetData("select * from  invi_obj where theos='" & rs!invos_infoid & "' and instanceid='" & id & "'")
              If rsc.EOF Then
              '  добавляем объект
                Session.GetData "insert into invi_obj ( instanceid, invi_objid, TheOS) values( '" & id & "','" & CreateGUID2() & "','" & rs!invos_infoid & "')"
                cnt = cnt + 1
              End If
              ' проверяем что такого объекта не было в прошедших
              Set rsc = Session.GetData("select * from  invi_done where theos='" & rs!invos_infoid & "' and instanceid='" & id & "'")
              If rsc.EOF Then
              
              ' добавляем в прошедшие ивентаризацию
              Session.GetData "insert into invi_done ( instanceid, invi_doneid, checkdate,TheOS,OSStatus) values( '" & id & "','" & CreateGUID2() & "'," & IIf(Session.IsMSSQL, MakeMSSQLDate(rs!InvDate), IIf(Session.IsORACLE, MakeORACLEDate(rs!InvDate), MakePGSQLDate(rs!InvDate))) & ",'" & rs!invos_infoid & "','" & rs!OSStatus & "')"
              cnt2 = cnt2 + 1
              End If
             ' прицепляем объект к  инвентаризации
             Session.GetData "update invos_inv set inventory='" & inv.invi_DEF.Item(1).id & "' where invos_invid='" & rs!invos_invid & "'"
          rs.MoveNext
        Wend
      End If
    End If
    MsgBox "Добавлено объектов: " & cnt & ". Добавлено в проверенные:" & cnt2
    
  End If
  
End Sub

Private Sub cmdCancel_Click()
OK = False
Me.Hide
End Sub

Private Sub cmdInvI_Click()
Dim id As String
Dim brief As String
If Manager.GetObjectListDialogEx3(id, brief, "", "INV_INV") Then
    txtINV.Text = brief
    txtINV.Tag = id
End If

End Sub

Private Sub cmdOK_Click()

If cmbExclude.ListIndex >= 0 Then
  If cmbExclude.ListIndex < 2 Then
    If txtINV.Tag <> "" Then
    
      OK = True
      Me.Hide
      Exit Sub
    End If
  Else
    If txtMask <> "" Then
       OK = True
       Me.Hide
       Exit Sub
    End If
  End If
End If
MsgBox "Необходимо задать параметры исключаемых данных", vbOKOnly + vbCritical, "Внимание"

End Sub


Private Sub cmdOrg_Click()
  Dim id As String
  Dim brief As String
  
  If Manager.GetReferenceDialogEx2("INVD_ORG", id, brief, "") Then
      txtOrg.Text = brief
      txtOrg.Tag = id
  End If
End Sub
