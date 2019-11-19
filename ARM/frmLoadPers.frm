VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmLoadPers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Загрузка справочника персонала"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPath 
      Caption         =   "..."
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   480
      Width           =   450
   End
   Begin VB.TextBox txtPath 
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   480
      Width           =   3735
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Загрузить справочник персонала"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   4335
   End
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   3960
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Файл с данными"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "frmLoadPers"
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


Private Sub cmdPath_Click()
 On Error Resume Next
  
  On Error GoTo bye
  Dim fn As String
  cdlg.CancelError = True
  cdlg.Filter = "Документ |*.xls"
  cdlg.DefaultExt = "XLS"
  cdlg.flags = cdlOFNPathMustExist + cdlOFNHideReadOnly + cdlOFNFileMustExist
  cdlg.ShowOpen
  txtPath.Tag = cdlg.Name
  txtPath = cdlg.fileName
bye:
End Sub

Private Sub Command1_Click()
    
    If txtPath.Text = "" Then Exit Sub
    
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
    
    
    Dim ex As Object 'excel.Application
    Dim wb As Object 'excel.Workbook
    Dim ws As Object 'excel.Worksheet
    Dim rng As Object 'excel.Range
    
    On Error GoTo bye
    Set wb = CreateObject(txtPath.Text)
    On Error Resume Next
    
    Set ws = wb.Worksheets.Item(1)
    
     
    
    Dim r As Long
    Dim c As Long
    For r = 16 To 65535
      Set rng = ws.Cells(r, 3)
      If rng.Value = "" Then Exit For
      FindOwnerByName rng.Value
      NextVal
    Next
    Me.Hide
    Exit Sub
bye:
    MsgBox "Ошибка открытия файла." & vbCrLf & "Проверьте формат файла. Ожидается Excel 2003 и выше.", vbCritical + vbOKOnly, "Ощибка"
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
      Set rs = Session.GetData("select  * from INVD_owner where  trim(COALESCE(cast(FamiliName as varchar),'') ||' ' ||  COALESCE(cast(Name as varchar),'') ||' ' || COALESCE(cast(SurName as varchar),'') || ' ')=trim('" & Trim(Name) & " ')")
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

