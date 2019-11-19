VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TOOL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private dic As INV_DIC.Application
Private invNum As INV_NUM.Application


Public Function GetNumber(Item As Object, OldNum As String, TheOrgID As String, IsMaterial As String, CardNum As String) As String
    
    If Item Is Nothing Then
      GetNumber = OldNum
      Exit Function
    End If
    Dim rs As ADODB.Recordset
    Dim id As String
    Set rs = Item.Application.Manager.ListInstances("", "INV_DIC")
    If Not rs.EOF Then
      id = rs!InstanceID
    Else
      id = CreateGUID2
      Item.Application.Manager.NewInstance id, "INV_DIC", "Справочник"
    End If
    Set dic = Item.Application.Manager.GetInstanceObject(id)
    
    
    
    Set rs = Item.Application.MTZSession.GetData("select * from invn_DEF where theORG='" & TheOrgID & "'")
    If Not rs.EOF Then
      id = rs!InstanceID
      Set invNum = Item.Application.Manager.GetInstanceObject(id)

    Else
      id = CreateGUID2
      Item.Application.Manager.NewInstance id, "INV_NUM", "Нумерация"
      Set invNum = Item.Application.Manager.GetInstanceObject(id)
      With invNum.INVN_DEF.Add
        Set .TheOrg = FindOrg(TheOrgID)
        .Save
      End With
    End If
    
    invNum.LockResource False
    If invNum.IsLocked <> LockSession Then
      MsgBox "Не удалось заблокировать нумератор"
       GetNumber = OldNum
      Exit Function
    End If
    Dim gn As String
    If IsMaterial = "Да" Then
      gn = Right("00" & (Val(invNum.INVN_DEF.Item(1).TheOrg.NumPrefix) + 50), 2) & Right("00000000" & GetNextInvNum(), 8)
    Else
      If IsNumeric(CardNum) Then
        gn = Right("00" & (Val(invNum.INVN_DEF.Item(1).TheOrg.NumPrefix)), 2) & Right("00000000" & CardNum, 8)
      Else
        gn = Right("00" & (Val(invNum.INVN_DEF.Item(1).TheOrg.NumPrefix)), 2) & Right("00000000" & GetNextInvNumOS(), 8)
      End If
    End If
    Dim iapp As INV_OS.Application
    Set iapp = Item.Application
    If iapp.INVOS_CODE.Count = 0 Then
      iapp.INVOS_CODE.Add
    End If
    
    With iapp.INVOS_CODE.Item(1)
      .VisibleCode = gn
      .ShCode = Code128(gn)
      .MabualCreate = Boolean_Da
      .CodePrinted = Boolean_Net
      .Save
    End With
    
    If iapp.INVOS_SROK.Count = 0 Then
      iapp.INVOS_SROK.Add
      With iapp.INVOS_SROK.Item(1)
        .RecalcDate = DateAdd("m", 1, DateSerial(Year(Date), Month(Date), 1))
        .Save
      End With
        
    End If
    
    
    invNum.UnLockResource
    GetNumber = gn
End Function

Private Function GetNextInvNum() As Integer
 
  invNum.INVN_DEF.Item(1).TheNumber = invNum.INVN_DEF.Item(1).TheNumber + 1
  invNum.INVN_DEF.Item(1).Save
  GetNextInvNum = invNum.INVN_DEF.Item(1).TheNumber

End Function

Private Function GetNextInvNumOS() As Integer
 
  invNum.INVN_DEF.Item(1).TheOSNumber = invNum.INVN_DEF.Item(1).TheOSNumber + 1
  invNum.INVN_DEF.Item(1).Save
  GetNextInvNumOS = invNum.INVN_DEF.Item(1).TheOSNumber

End Function

Private Function FindOrg(ByVal id As String) As INVD_ORG
    Dim dd As INV_DIC.INVD_ORG
    Set dd = dic.INVD_ORG.Item(id)
    Set FindOrg = dd
End Function
