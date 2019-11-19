VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLoadPortal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Загрузка и актуализация данных о владельцах"
   ClientHeight    =   1080
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Загрузить справочник персонала"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4335
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "frmLoadPortal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public OK As Boolean
Private dic As INV_DIC.Application
Public invNum As INV_NUM.Application

Private Sub NextVal()
    pb.Value = (pb.Value + 1) Mod 100
End Sub



Private Sub CancelButton_Click()
  Me.Hide
End Sub

Private Sub Command1_Click()
  Dim txtSrv As String
  Dim txtDB As String
  Dim txtUsr As String
  Dim txtPass As String
  Dim txtTable As String
    
    
    txtSrv = GetSetting("ABOL", "PORTALDB", "PORTALSRV", "")
    txtDB = GetSetting("ABOL", "PORTALDB", "PORTALDB", "")
    txtUsr = GetSetting("ABOL", "PORTALDB", "PORTALUSR", "")
    txtPass = GetSetting("ABOL", "PORTALDB", "PORTALPASS", "")
    txtTable = GetSetting("ABOL", "PORTALDB", "PORTALTABLE", "S25")

  Dim conn As ADODB.Connection
  Set conn = New ADODB.Connection
  conn.Provider = "SQLoledb"
  conn.ConnectionString = "Server=" & txtSrv & ";DataBase=" & txtDB & ";UID=" & txtUsr & ";Pwd=" & txtPass & ";"
  conn.open
  If conn.State = ADODB.adStateOpen Then
   
  Else
    MsgBox "Ошибка параметров соединения"
    Set conn = Nothing
    Exit Sub
  End If
 

  Dim prs As ADODB.Recordset

    
    Dim rs As ADODB.Recordset
    Dim id As String
    Set rs = Manager.ListInstances("", "INV_DIC")
    If Not rs.EOF Then
      id = rs!InstanceID
    Else
      id = CreateGUID2
      Manager.NewInstance id, "INV_DIC", "Справочник"
    End If
    Set dic = Manager.GetInstanceObject(id)
    Manager.LockInstanceObject id
    Dim i As Integer
    
    
    On Error GoTo bye
    
     
    Set prs = conn.Execute("select * from " & txtTable)
    If prs Is Nothing Then
      MsgBox "Ошибка параметров соединения"
      Set conn = Nothing
      Exit Sub
    End If
    Dim s As String
    Dim K As String
    Dim P As String
    Dim OWNER As INVD_OWNER
    On Error Resume Next
    
    While Not prs.EOF
      s = prs("ФИО").Value
      K = prs("Кабинет").Value
      P = prs("р_место").Value
      Set OWNER = FindOwnerByName(s)
      If Not OWNER Is Nothing Then
       If Trim(K) <> "" And Trim(P) <> "" Then
        Session.GetData "update INVOS_PLACE set TheOwner ='" & OWNER.id & "' WHERE ComplNumber ='" & Trim(K) & "." & Trim(P) & "'"
       End If
      End If
      NextVal
      prs.MoveNext
    Wend
    Set prs = Nothing
    
    MsgBox "Данные актуализированы"
    Me.Hide
    Exit Sub
    
bye:
    MsgBox "Ошибка при запросе данных с портала.", vbCritical + vbOKOnly, "Ощибка"
End Sub

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
      Set rs = Session.GetData("select  * from INVD_owner where isnull(FamiliName,'') +' ' + isnull(Name,'') +' ' +isnull(SurName,'')+ ' ' ='" & Trim(Name) & " '")
    End If
    If Session.IsPOSTGRESQL Then
      Set rs = Session.GetData("select * from INVD_owner where  trim(COALESCE(cast(FamiliName as varchar),'') ||' ' ||  COALESCE(cast(Name as varchar),'') ||' ' || COALESCE(cast(SurName as varchar),'') || ' ')=trim('" & Trim(Name) & " ')")
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


