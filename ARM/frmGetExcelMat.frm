VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{1801C003-859D-471D-BF31-D4428050324B}#2.1#0"; "MTZ_PANEL.ocx"
Begin VB.Form frmGetExcelMat 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Загрузка данных по материалам"
   ClientHeight    =   6270
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtMOL 
      Height          =   300
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   24
      ToolTipText     =   "Владелец"
      Top             =   5250
      Width           =   2550
   End
   Begin VB.TextBox txtINVOS_PLACE_Uprav 
      Height          =   300
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   22
      ToolTipText     =   "Управление"
      Top             =   3810
      Width           =   2550
   End
   Begin VB.TextBox txtINVOS_PLACE_TheOwner 
      Height          =   300
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   19
      ToolTipText     =   "Владелец"
      Top             =   4530
      Width           =   2550
   End
   Begin VB.TextBox txtINVOS_PLACE_TheOrg 
      Height          =   300
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   14
      ToolTipText     =   "На учете в "
      Top             =   1800
      Width           =   2550
   End
   Begin VB.TextBox txtINVOS_PLACE_TheHouse 
      Height          =   300
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   12
      ToolTipText     =   "Здание"
      Top             =   2520
      Width           =   2550
   End
   Begin VB.TextBox txtINVOS_PLACE_DIrection 
      Height          =   300
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   10
      ToolTipText     =   "Дирекция"
      Top             =   3165
      Width           =   2550
   End
   Begin VB.TextBox txtINVOS_INFO_OSType 
      Height          =   300
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   8
      ToolTipText     =   "Группа ОС"
      Top             =   1170
      Width           =   2550
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   5760
      Visible         =   0   'False
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox txtPath 
      Height          =   375
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   2535
   End
   Begin VB.CommandButton cmdPath 
      Caption         =   "..."
      Height          =   375
      Left            =   2790
      TabIndex        =   2
      Top             =   360
      Width           =   450
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   4080
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MTZ_PANEL.DropButton cmdINVOS_INFO_OSType 
      Height          =   300
      Left            =   2790
      TabIndex        =   7
      Tag             =   "refopen.ico"
      ToolTipText     =   "Группа ОС"
      Top             =   1170
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   529
      Caption         =   "..."
      Caption         =   "..."
   End
   Begin MTZ_PANEL.DropButton cmdINVOS_PLACE_DIrection 
      Height          =   300
      Left            =   2790
      TabIndex        =   9
      Tag             =   "refopen.ico"
      ToolTipText     =   "Дирекция"
      Top             =   3165
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   529
      Caption         =   "..."
      Caption         =   "..."
   End
   Begin MTZ_PANEL.DropButton cmdINVOS_PLACE_TheHouse 
      Height          =   300
      Left            =   2790
      TabIndex        =   11
      Tag             =   "refopen.ico"
      ToolTipText     =   "Здание"
      Top             =   2520
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   529
      Caption         =   "..."
      Caption         =   "..."
   End
   Begin MTZ_PANEL.DropButton cmdINVOS_PLACE_TheOrg 
      Height          =   300
      Left            =   2790
      TabIndex        =   13
      Tag             =   "refopen.ico"
      ToolTipText     =   "На учете в "
      Top             =   1800
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   529
      Caption         =   "..."
      Caption         =   "..."
   End
   Begin MTZ_PANEL.DropButton cmdINVOS_PLACE_TheOwner 
      Height          =   300
      Left            =   2790
      TabIndex        =   18
      Tag             =   "refopen.ico"
      ToolTipText     =   "Владелец"
      Top             =   4560
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   529
      Caption         =   "..."
      Caption         =   "..."
   End
   Begin MTZ_PANEL.DropButton cmdINVOS_PLACE_Uprav 
      Height          =   300
      Left            =   2790
      TabIndex        =   21
      Tag             =   "refopen.ico"
      ToolTipText     =   "Управление"
      Top             =   3840
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   529
      Caption         =   "..."
      Caption         =   "..."
   End
   Begin MTZ_PANEL.DropButton cmdMOL 
      Height          =   300
      Left            =   2790
      TabIndex        =   25
      Tag             =   "refopen.ico"
      ToolTipText     =   "Владелец"
      Top             =   5280
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   529
      Caption         =   "..."
      Caption         =   "..."
   End
   Begin VB.Label Label8 
      Caption         =   "МОЛ:"
      Height          =   375
      Left            =   240
      TabIndex        =   26
      Top             =   4920
      Width           =   3135
   End
   Begin VB.Label Label7 
      Caption         =   "Управление:"
      Height          =   375
      Left            =   240
      TabIndex        =   23
      Top             =   3480
      Width           =   2655
   End
   Begin VB.Label Label6 
      Caption         =   "Владелец:"
      Height          =   375
      Left            =   240
      TabIndex        =   20
      Top             =   4200
      Width           =   3135
   End
   Begin VB.Label Label5 
      Caption         =   "Дирекция:"
      Height          =   375
      Left            =   240
      TabIndex        =   17
      Top             =   2880
      Width           =   2895
   End
   Begin VB.Label Label4 
      Caption         =   "Здание:"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   2160
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "На учете в :"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "Группа ОС:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Файл с данными"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmGetExcelMat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public OK As Boolean
Private dic As INV_DIC.Application
Public invNum As INV_NUM.Application
Private NewCount As Long
Public MatRows As Collection

Private Sub NextVal()
    pb.Value = (pb.Value + 1) Mod 100
End Sub
Private Sub CancelButton_Click()
    OK = False
    Me.Hide
End Sub

Private Sub cmdINVOS_INFO_OSType_Click()
  On Error Resume Next
        Dim id As String, brief As String
        If Manager.GetReferenceDialogEx2("INVD_OSTYPE", id, brief) Then
          txtINVOS_INFO_OSType.Tag = Left(id, 38)
          txtINVOS_INFO_OSType = brief
        End If
End Sub

Private Sub cmdINVOS_PLACE_DIrection_Click()
  On Error Resume Next
        Dim id As String, brief As String
        If Manager.GetReferenceDialogEx2("INVD_DIR", id, brief) Then
          txtINVOS_PLACE_DIrection.Tag = Left(id, 38)
          txtINVOS_PLACE_DIrection = brief
        End If
End Sub

Private Sub cmdINVOS_PLACE_TheHouse_Click()
On Error Resume Next
        Dim id As String, brief As String
        If Manager.GetReferenceDialogEx2("INVD_BLD", id, brief) Then
          txtINVOS_PLACE_TheHouse.Tag = Left(id, 38)
          txtINVOS_PLACE_TheHouse = brief
        End If
End Sub

Private Sub cmdINVOS_PLACE_TheOrg_Click()
 On Error Resume Next
        Dim id As String, brief As String
        If Manager.GetReferenceDialogEx2("INVD_ORG", id, brief) Then
          txtINVOS_PLACE_TheOrg.Tag = Left(id, 38)
          txtINVOS_PLACE_TheOrg = brief
        End If
End Sub

Private Sub cmdINVOS_PLACE_TheOwner_Click()
On Error Resume Next
        Dim id As String, brief As String
        If Manager.GetReferenceDialogEx2("INVD_OWNER", id, brief) Then
          txtINVOS_PLACE_TheOwner.Tag = Left(id, 38)
          txtINVOS_PLACE_TheOwner = brief
        End If
End Sub

Private Sub cmdINVOS_PLACE_Uprav_Click()
On Error Resume Next
        Dim id As String, brief As String
        If Manager.GetReferenceDialogEx2("INVD_UPR", id, brief) Then
          txtINVOS_PLACE_Uprav.Tag = Left(id, 38)
          txtINVOS_PLACE_Uprav = brief
        End If
End Sub



Private Sub cmdMOL_Click()
        On Error Resume Next
        Dim id As String, brief As String
        If Manager.GetReferenceDialogEx2("INVD_OWNER", id, brief) Then
          txtMOL.Tag = Left(id, 38)
          txtMOL.Text = brief
        End If
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

Private Sub OKButton_Click()
    If txtPath.Text <> "" Then
      If txtINVOS_PLACE_TheOrg.Tag = "" Then
        MsgBox "Необходимо задать организацию"
      Else
        Dim md5 As String
        Dim ccc As CMD5
        Set ccc = New CMD5
        md5 = ccc.FileMD5(txtPath.Text)
        Set ccc = Nothing
        If IsFileLoaded(txtPath.Text, md5, "МАТЕРИАЛЫ") Then
          MsgBox "Файл уже загружен в систему"
        Else
          If LoadXLS(txtPath.Text) Then
          '  файл меняется в процессе работы
             Set ccc = New CMD5
             md5 = ccc.FileMD5(txtPath.Text)
             Set ccc = Nothing
          
              RegisterFile txtPath.Text, md5, "МАТЕРИАЛЫ"
              OK = True
              MsgBox "Загрузка завершена" & vbCrLf & "Добавлено " & NewCount & " объектов."
              Me.Hide
          End If
        End If
       End If
    End If
End Sub




Private Function LoadXLS(ByVal path As String) As Boolean
    Dim res As Boolean
    res = True
    NewCount = 0
    Set MatRows = New Collection
    
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
    
     Set rs = Session.GetData("select * from invn_DEF where theORG='" & txtINVOS_PLACE_TheOrg.Tag & "'")
    If Not rs.EOF Then
      id = rs!InstanceID
      Set invNum = Manager.GetInstanceObject(id)

    Else
      id = CreateGUID2
      Manager.NewInstance id, "INV_NUM", "Нумерация"
      Set invNum = Manager.GetInstanceObject(id)
      With invNum.INVN_DEF.Add
        Set .theorg = FindOrg(txtINVOS_PLACE_TheOrg.Tag)
        .save
      End With
      
      
    End If
    
    Manager.LockInstanceObject id
    
    invNum.LockResource False
    If invNum.IsLocked <> LockSession Then
      MsgBox "Не удалось заблокировать нумератор"
      Exit Function
    End If
    
    Dim ex As Object 'excel.Application
    Dim wb As Object 'excel.Workbook
    Dim ws As Object 'excel.Worksheet
    Dim rng As Object 'excel.Range
    
    On Error GoTo bye
    Set wb = CreateObject(path)
    On Error Resume Next
    
    Set ws = wb.Worksheets.Item(1)
    
     Set rng = ws.Cells(2, 2)
     If Left(UCase(rng.Value), 8) = "МАТЕРИАЛ" Then
    
    Dim r As Long
    Dim c As Long
    Dim os As INV_OS.Application
    Dim inf As INVOS_INFO
    Dim Doc As INVOS_DOCS
   
    Dim inv As Long
    inv = 1000
    pb.Min = 0
    pb.Max = 100
    pb.Value = 0
    pb.Visible = True
    Dim q As Integer
    Dim cnum As String
    Dim Name As String
    Dim mIdx As Integer
    Dim theorg As INVD_ORG
    
    For r = 6 To 64000
        NextVal
  
        
        
        Me.Caption = r
        DoEvents
       
        Set rng = ws.Cells(r, 2)
        If rng.Value <> "конецфайла" Then
         
         Set rng = ws.Cells(r, 3)
         q = CInt(val(rng.Value))
         
         Set rng = ws.Cells(r, 13)
         
         If Trim(rng.Value) = "" Then
         
         
           Set rng = ws.Cells(r, 2)
           Name = rng.Value
           Set rng = ws.Cells(r, 4)
           cnum = rng.Value
           
          On Error Resume Next
          Dim mr As MatRow
          Set mr = New MatRow
          mr.Code = txtINVOS_PLACE_TheOrg.Tag & "|" & cnum & "|" & Name
          mr.Quantity = q
          If MatRows.Item(txtINVOS_PLACE_TheOrg.Tag & "|" & cnum & "|" & Name) Is Nothing Then
            MatRows.Add mr, mr.Code
            q = MatRows.Item(mr.Code).Quantity
          Else
            MatRows.Item(mr.Code).Quantity = MatRows.Item(mr.Code).Quantity + q
            q = MatRows.Item(mr.Code).Quantity
            Debug.Print mr.Code
            
          End If
          
         
           ' записываем поступление
           For mIdx = 1 To q
           
                
                Set rs = Session.GetData("select * from v_autoinvos_info where  INVOS_INFO_TheOrg_ID='" & txtINVOS_PLACE_TheOrg.Tag & "' and INVOS_INFO_CardNum='" & cnum & "' and INVOS_INFO_ShortName ='" & Name & "' and invos_info_InLineNum=" & mIdx)
                
                
                If rs.EOF Then
                   NewCount = NewCount + 1
                   id = CreateGUID2()
                   Set rng = ws.Cells(r, 2)
                   Manager.NewInstance id, "INV_OS", rng.Value
                   Set os = Manager.GetInstanceObject(id)
                   Set inf = os.INVOS_INFO.Add
                   
                   If txtINVOS_PLACE_TheOrg.Tag <> "" Then
                     Set theorg = FindOrg(txtINVOS_PLACE_TheOrg.Tag)
                     Set inf.theorg = theorg
                   End If
                   
                   inf.invNum = Right("00" & (val(theorg.NumPrefix) + 50), 2) & Right("00000000" & GetNextInvNum(), 8)
                   
                   Set Doc = os.INVOS_DOCS.Add
                   
                Else
                  id = rs!InstanceID
                  Set os = Manager.GetInstanceObject(id)
                  Set inf = os.INVOS_INFO.Item(1)
                  Set Doc = os.INVOS_DOCS.Item(1)
                End If
                  
                   
                   inf.InLineNum = mIdx
                   inf.IsMaterial = Boolean_Da
                   Dim compl As String
                  
                   For c = 2 To 16
                       Set rng = ws.Cells(r, c)
                       Select Case c
                       Case 2
                           inf.Name = rng.Value
                           inf.ShortName = rng.Value
                        
                           compl = GetCompl(inf.Name)
                           
                       Case 3
                           
                       Case 4
                           inf.CardNum = rng.Value
                       Case 5
                          Doc.InOrderNum = rng.Value
                       Case 6
                           Doc.NaklNum = rng.Value
                       Case 7
                           Set Doc.Contragent = FindAgentByName(rng.Value)
                       Case 8
                           Doc.DogNum = Left(rng.Value, 30)
                       Case 9
                           Doc.AccFNum = rng.Value
                       Case 10
                           Doc.AccNum = rng.Value
                        Case 11
                           inf.TheCost = val(Replace(rng.Value, ",", "."))
                       Case 16
                           If Trim(rng.Value & "") <> "" Then
                                inf.TheCost = val(Replace(rng.Value, ",", "."))
                           End If
                       Case 12
                       Case 13
                       Case Else
                       
                       End Select
                      
                   Next
                   
                   inf.SrokFI = 0
                   inf.SrokPI = 12
                   inf.SrokOI = 12
                  
                   ' save place data
                   If os.INVOS_PLACE.Count = 0 Then
                       os.INVOS_PLACE.Add
                   End If
                   
                   If txtINVOS_INFO_OSType.Tag <> "" Then
                       Set inf.ostype = FindOSType(txtINVOS_INFO_OSType.Tag)
                   End If
                   
                   Dim complArr() As String
                   With os.INVOS_PLACE.Item(1)
                       If compl <> "" Then
                          If .ComplNumber = "" Then
                           .ComplNumber = compl
                          End If
                          complArr = Split(compl, ".")
                          If UBound(complArr) >= 0 And .Flow = "" Then
                           .Flow = complArr(0)
                          End If
                          If UBound(complArr) >= 1 And .Room = "" Then
                           .Room = complArr(1)
                          End If
                          If UBound(complArr) >= 2 And .WorkPlaceNum = "" Then
                           .WorkPlaceNum = complArr(2)
                          End If
                       End If
                       If txtINVOS_PLACE_DIrection.Tag <> "" Then
                           Set .Direction = FindDir(txtINVOS_PLACE_DIrection.Tag)
                       End If
                       If txtINVOS_PLACE_TheHouse.Tag <> "" Then
                           Set .TheHouse = FindBuilding(txtINVOS_PLACE_TheHouse.Tag)
                       End If
                     
                       If txtINVOS_PLACE_TheOwner.Tag <> "" Then
                           Set .TheOwner = FindOwner(txtINVOS_PLACE_TheOwner.Tag)
                       End If
                       If txtINVOS_PLACE_Uprav.Tag <> "" Then
                           Set .Uprav = FindUPR(txtINVOS_PLACE_Uprav.Tag)
                       End If
                       
                        If txtMOL.Tag <> "" Then
                           Set .MatOtv = FindOwner(txtMOL.Tag)
                       End If
                       
                       .save
                   End With
                   
                   SaveHistory os.INVOS_PLACE.Item(1)
                 
                   
                    If os.INVOS_CODE.Count = 0 Then
                       os.INVOS_CODE.Add
                   End If
                   
                   With os.INVOS_CODE.Item(1)
                     .VisibleCode = inf.invNum
                     .ShCode = MTZUtil.Code128(.VisibleCode)
                     .save
                   End With
                   
                   inf.save
                   Doc.save
                   
                   If os.StatusID = "{179CB53A-CBE7-46B4-9905-22E35FAAE801}" Then
                       os.StatusID = "{8AD15E54-CF87-4FCF-8A1E-A85336E23C73}"
                   End If
                   
                  If os.INVOS_SROK.Count = 0 Then
                    os.INVOS_SROK.Add
                    With os.INVOS_SROK.Item(1)
                      .RecalcDate = DateAdd("m", 1, DateSerial(Year(Date), Month(Date), 1))
                      .save
                    End With
                      
                  End If
                  
               
            Next
 
          
          End If
        Else
            Exit For
        End If
        
        
      Manager.FreeAllInstanses
    Next
    
    
    ' списания
    For r = 6 To 64000
        NextVal
  
      
        
        Me.Caption = "Списание:" & r
        DoEvents
       
        Set rng = ws.Cells(r, 2)
        If rng.Value <> "конецфайла" Then
         
         Set rng = ws.Cells(r, 3)
         q = CInt(val(rng.Value))
         
         Set rng = ws.Cells(r, 13)
         
          If Trim(rng.Value) <> "" Then
      
           Set rng = ws.Cells(r, 2)
           Name = rng.Value
           Set rng = ws.Cells(r, 4)
           cnum = rng.Value
          
          
           Set rs = Session.GetData("select * from v_autoinvos_info where   INVOS_INFO_TheOrg_ID='" & txtINVOS_PLACE_TheOrg.Tag & "' and INVOS_INFO_CardNum='" & cnum & "' and INVOS_INFO_ShortName ='" & Name & "' and INTSANCEStatusID <>'{166D4978-0C4C-4575-8192-B251AC113781}'")
           While Not rs.EOF
                id = rs!InstanceID
                Set os = Manager.GetInstanceObject(id)
                Set inf = os.INVOS_INFO.Add()
                Set rng = ws.Cells(r, 13)
             
                If os.INVOS_OFFRULE.Count = 0 Then
                    os.INVOS_OFFRULE.Add
                End If
                With os.INVOS_OFFRULE.Item(1)
                    .Info = rng.Value
                    .save
                End With
                os.StatusID = "{166D4978-0C4C-4575-8192-B251AC113781}"
                q = q - 1
                If q = 0 Then
                 GoTo done
                End If
                rs.MoveNext
           Wend
done:
          
          
          End If
        Else
            Exit For
        End If
        
        
      Manager.FreeAllInstanses
    Next
    Else
      MsgBox "Неверный формат отчета"
    End If
    
    
    pb.Visible = False
    LoadXLS = res
    invNum.UnLockResource
    
    
    
  Exit Function
bye:
    invNum.UnLockResource
    MsgBox "Ошибка открытия файла." & vbCrLf & "Проверьте формат файла. Ожидается Excel 2003 и выше.", vbCritical + vbOKOnly, "Ощибка"
End Function


Private Function GetNextInvNum() As Integer
  If invNum.INVN_DEF.Count = 0 Then
    invNum.INVN_DEF.Add
  End If
  invNum.INVN_DEF.Item(1).TheNumber = invNum.INVN_DEF.Item(1).TheNumber + 1
  invNum.INVN_DEF.Item(1).save
  GetNextInvNum = invNum.INVN_DEF.Item(1).TheNumber

End Function


Private Function FindOSType(ByVal id As String) As INVD_OSTYPE
    Dim ost As INVD_OSTYPE
    Set ost = dic.INVD_OSTYPE.Item(id)
    Set FindOSType = ost
 End Function


Private Function FindBuilding(ByVal id As String) As INVD_BLD
   
    Dim bb As INV_DIC.INVD_BLD
   
    Set bb = dic.INVD_BLD.Item(id)
    
    Set FindBuilding = bb
End Function


Private Function FindUPR(ByVal id As String) As INVD_UPR
    Dim dd As INV_DIC.INVD_UPR
    Set dd = dic.INVD_UPR.Item(id)
    
    Set FindUPR = dd
End Function

Private Function FindDir(ByVal id As String) As INVD_DIR
    Dim dd As INV_DIC.INVD_DIR
    Set dd = dic.INVD_DIR.Item(id)
    
    Set FindDir = dd
End Function

Private Function FindOrg(ByVal id As String) As INVD_ORG
    Dim dd As INV_DIC.INVD_ORG
    Set dd = dic.INVD_ORG.Item(id)
    Set FindOrg = dd
End Function

Private Function FindOwner(ByVal id As String) As INVD_OWNER
    Dim dd As INV_DIC.INVD_OWNER
    Set dd = dic.INVD_OWNER.Item(id)
    Set FindOwner = dd
End Function

Private Function FindAgentByName(ByVal Name As String) As INVD_UR
    Dim ur As INVD_UR
    Dim rs As ADODB.Recordset
    Set rs = Session.GetData("select * from INVD_UR where sortname='" & Name & "' or fullname ='" & Name & "'")
    If rs.EOF Then
        Set ur = dic.INVD_UR.Add
        ur.SortName = Name
        ur.FullName = Name
        ur.save
    Else
        Set ur = dic.INVD_UR.Item(rs!invd_urid)
        ur.FullName = Name
        ur.save
    End If
    Set FindAgentByName = ur
    

End Function
