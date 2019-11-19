VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BEFORESTATUS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim theItem As INV_INV.Application
Dim Manager As MTZManager.Main
Dim session As MTZSession.session
Dim conn As adodb.Connection
  





Public Function Check(Item As Object, ByVal NewStatusid As String, ByVal CurrentUser As Object, ByVal Module As String) As Boolean
  Dim result As Boolean
  Dim id As String
  result = True
  
  On Error Resume Next
  Set theItem = Item
  Set Manager = theItem.Manager
  Set session = theItem.MTZSession
  Dim Obj As INV_OS.Application
  
 

  Dim i As Long
  
  Dim rs As adodb.Recordset
  Dim s As String
  
  
' состояния для типа:INV_OS Карточка основного средства
' "{2AA78799-2880-4541-99E0-3C8750AC33E6}" 'В аренде
' "{8E6E78D2-82AA-4913-B08C-1230A8C8B4A9}" 'В ремонте
' "{8AD15E54-CF87-4FCF-8A1E-A85336E23C73}" 'Введено в эксплуатацию
' "{179CB53A-CBE7-46B4-9905-22E35FAAE801}" 'Оформляется
' "{166D4978-0C4C-4575-8192-B251AC113781}" 'Списано
  
' состояния для типа:INV_INV Инвернтаризация
' "{FA929BE8-0966-46CD-99FC-FFF5E25EC4D5}" 'Идет инвентаризация
' "{03A3E27E-FF6E-4325-8174-462D88422A0E}" 'Инвентаризация завершена
' "{926A2E1C-FBF5-44A4-9536-E195AF47D32F}" 'Оформляется
' "{706FBA86-116E-4CF4-932E-32CF7DEBC573}" 'Утверждена

    If NewStatusid = "{706FBA86-116E-4CF4-932E-32CF7DEBC573}" Then
        ' строим запрос по параметрам инвентаризации
        ' очищаем коллекцию объектов инвентаризации
        ' копируем список объектов подлежащих инвентаризации в новую инвентарную ведомость
        
        s = " INTSANCEStatusID not in ('{166D4978-0C4C-4575-8192-B251AC113781}','{179CB53A-CBE7-46B4-9905-22E35FAAE801}') "

        With theItem.invi_DEF.Item(1)
        Dim f As frmProgress
            Set f = New frmProgress
            f.pb.Min = 0
            f.pb.Value = 0
            f.pb.Max = 100
            'f.Caption = ""
            f.Show
        
            If Not .TheOrg Is Nothing Then
              s = s & " and INVOS_INFO_TheOrg_ID='" & .TheOrg.id & "'"
            End If
            
            If Not .Building Is Nothing Then
              s = s & " and INVOS_Place_TheHouse_ID='" & .Building.id & "'"
            End If
        
        
        
            If Not .Direction Is Nothing Then
            s = s & " and INVOS_PLACE_DIrection_ID='" & .Direction.id & "'"

            End If
            If Not .Uprev Is Nothing Then
            s = s & " and INVOS_PLACE_Uprav_ID='" & .Uprev.id & "'"
            End If
            
            If Not .Otdel Is Nothing Then
            s = s & " and INVOS_PLACE_Otdel_ID='" & .Otdel.id & "'"
            End If
            
            If Not .TheOwner Is Nothing Then
            s = s & " and INVOS_PLACE_TheOwner_ID='" & .TheOwner.id & "'"
            End If
            
            If .TheFlow <> "" Then
              s = s & " and INVOS_PLACE_Flow='" & .TheFlow & "'"
            End If
       
            If .TheRoom <> "" Then
              s = s & " and INVOS_PLACE_Room='" & .TheRoom & "'"
            End If
            
             If .TheWorkPlace <> "" Then
              s = s & " and INVOS_PLACE_WorkPlaceNum='" & .TheWorkPlace & "'"
            End If
       
                
            While theItem.INVI_OBJ.Count > 0
                theItem.INVI_OBJ.Delete (1)
                Nextval f
            Wend
            
           While theItem.INVI_DONE.Count > 0
                theItem.INVI_DONE.Delete (1)
                Nextval f
           Wend
           Set rs = session.GetData("select * from v_autoinvos_info where " & s)
           While Not rs.EOF
            id = rs!InstanceID
            Set Obj = Manager.GetInstanceObject(id)
            With theItem.INVI_OBJ.Add
                Set .TheOS = Obj.INVOS_INFO.Item(1)
                .Save
            End With
            Nextval f
            rs.MoveNext
           Wend
           Unload f
           Set f = Nothing

        End With
        
        
    
    End If
    
    
  Check = result
End Function

Private Function Nextval(f As frmProgress)
  f.pb.Value = (f.pb.Value + 1) Mod 100
End Function

