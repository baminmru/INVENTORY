Attribute VB_Name = "RoleSupport"
Option Explicit

Public MyRole As ROLES.Application  'Object

Public Enum RoleMenuStatus
  RoleMenuStatus_Unknown = 0
  RoleMenuStatus_Visible = 1
  RoleMenuStatus_Disabled = 2
  RoleMenuStatus_Hidden = 3
End Enum







Public Function BeforeChangeStatus(Item As Object, NewStatus As String) As Boolean
  Dim logic As Object
  Dim result As Boolean
  result = True
  On Error Resume Next
  Set logic = CreateObject(Item.TypeName & "BST.BEFORESTATUS")
  If Not logic Is Nothing Then
    result = logic.Check(Item, NewStatus, MyUser, Item.TypeName)
    Set logic = Nothing
  End If
  BeforeChangeStatus = result
End Function


Private Function Cast(g As String) As String
'    If Session.IsMSSQL Then
'        Cast = "'" & g & "'"
'    End If
'
'    If Session.IsORACLE Then
'        Cast = "'" & g & "'"
'    End If
'
'    If Session.IsPOSTGRESQL Then
'        Cast = "cast('" & g & "' as uuid)"
'    End If
Cast = g

End Function
Public Function ChooseRole() As Object
Dim rs As ADODB.Recordset
Dim Q1 As String, Q2 As String, Q3 As String, Q4 As String
Dim res1 As String, res2 As String, resroles As String, armroles As String

    ' ���� ������  � ������� ������ �����������
    Q1 = CreateGUID2
    
    Call Session.TheFinder.FIND_IDS(Cast(Q1), "GROUPUSER", "TheUser", OpEQ, Cast(MyUser.id))
    Q2 = CreateGUID2
    Call Session.TheFinder.RowsToParents("GROUPUSER", Q1, Q2)
    Q3 = CreateGUID2
    Call Session.TheFinder.FIND_IDS(Cast(Q3), "ROLES_MAP", "TheGroup", OpIN_RESULT, Cast(Q2))
    res1 = CreateGUID2
    Call Session.TheFinder.RowsToInstances("ROLES_MAP", Q3, res1)
    Session.TheFinder.DropResults Cast(Q1)
    Session.TheFinder.DropResults Cast(Q2)
    Session.TheFinder.DropResults Cast(Q3)
    
    Q1 = CreateGUID2
    Call Session.TheFinder.FIND_IDS(Cast(Q1), "ROLES_USER", "TheUser", OpEQ, Cast(MyUser.id))
    res2 = CreateGUID2
    Call Session.TheFinder.RowsToInstances("ROLES_USER", Q1, res2)
    Session.TheFinder.DropResults Cast(Q1)
    
    
    
    resroles = CreateGUID2
    ' �������� ����� ����� ������������
    Session.TheFinder.QR_OR_QR res1, res2, resroles
    Session.TheFinder.DropResults res1
    Session.TheFinder.DropResults res2
    
    
    
    ' ��������� ����� �� ���� ��������� ��� ���
    Q1 = CreateGUID2
    Call Session.TheFinder.FIND_IDS(Q1, "ROLES_WP", "WP", OpEQ, ARMID)
    res1 = CreateGUID2
    Call Session.TheFinder.RowsToInstances("ROLES_WP", Q1, res1)
    Session.TheFinder.DropResults Cast(Q1)
    armroles = CreateGUID2
    Session.TheFinder.QR_AND_QR resroles, res1, armroles
    
    Session.TheFinder.DropResults res1
    Session.TheFinder.DropResults resroles

    Set rs = Session.TheFinder.GetResults(armroles)
    If rs.EOF Then
        MsgBox "��� �� ��������� ������ � ���� �������", vbCritical + vbOKOnly, App.ProductName
        Set ChooseRole = Nothing
        Set rs = Nothing
        Session.TheFinder.DropResults armroles
        Exit Function
    End If
    Dim f As frmChooseRole
    Dim RoleObj As Object
    Set f = New frmChooseRole
    f.lstRole.Clear
    Dim i As Long
    Dim col As Collection
    Set col = New Collection
    i = 1
    While Not rs.EOF
        If Not IsNull(rs!result) Then
        Set RoleObj = Manager.GetInstanceObject(rs!result)
        col.Add RoleObj, RoleObj.id
         f.lstRole.AddItem RoleObj.Name
        f.lstRole.ItemData(f.lstRole.NewIndex) = i
        i = i + 1
        End If
        rs.MoveNext
    Wend
    Set rs = Nothing
    Session.TheFinder.DropResults armroles
    If col.Count = 1 Then
        Set ChooseRole = col.Item(f.lstRole.ItemData(0))
        Unload f
        Set f = Nothing
        Set col = Nothing
        Exit Function
    End If
    
    f.Show vbModal
    
    If f.OK Then
        Set ChooseRole = col.Item(f.lstRole.ItemData(f.lstRole.ListIndex))
        Unload f
        Set f = Nothing
        Set col = Nothing
        Exit Function
    Else
        Set ChooseRole = Nothing
        Unload f
        Set f = Nothing
        Set col = Nothing
        Exit Function
    End If
End Function

Public Function CheckMenu(ByVal menuName As String) As RoleMenuStatus
  Dim ms As RoleMenuStatus
  ms = RoleMenuStatus_Unknown
  If MyRole Is Nothing Then
    Exit Function
  End If
  Dim i As Long, j As Long
  Dim rwp As ROLES_WP
  Dim ract As ROLES_ACT
  
  For i = 1 To MyRole.ROLES_WP.Count
    If MyRole.ROLES_WP.Item(i).WP.id = ARMID Then
          Set rwp = MyRole.ROLES_WP.Item(i)
      Exit For
    End If
  Next
  
  Set ract = FindRoleAct(rwp.ROLES_ACT, menuName)
  If Not ract Is Nothing Then
    If ract.Accesible = YesNo_Da Then
      ms = RoleMenuStatus_Visible
    End If
    If ract.Accesible = YesNo_Net Then
      ms = RoleMenuStatus_Hidden
    End If
  End If
  CheckMenu = ms
End Function

Private Function FindRoleAct(ByVal col As ROLES_ACT_COL, ByVal Name As String) As ROLES_ACT
  Dim i As Long, j As Long
  Dim ract As ROLES_ACT
  
  Set ract = Nothing
  For i = 1 To col.Count
    If UCase(col.Item(i).EntryPoints.Name) = UCase(Name) Then
      Set ract = col.Item(i)
      Exit For
    End If
    If UCase(col.Item(i).EntryPoints.Caption) = UCase(Name) Then
      Set ract = col.Item(i)
      Exit For
    End If
    If ract Is Nothing Then
      Set ract = FindRoleAct(col.Item(i).ROLES_ACT, Name)
      If Not ract Is Nothing Then Exit For
    End If
  Next
  Set FindRoleAct = ract
End Function


Public Function GetDocumentMode(ByVal Obj As Object) As String
  Dim sid As String
  Dim tn As String
  Dim i As Long, j As Long
  Dim os As INV_OS.Application
  GetDocumentMode = ""
  If MyRole Is Nothing Then Exit Function
  tn = Obj.TypeName
  sid = Obj.StatusID
  For i = 1 To MyRole.ROLES_DOC.Count
    ' ����� ���
    If UCase(MyRole.ROLES_DOC.Item(i).The_Document.Name) = UCase(tn) Then
        ' ��� �������� � ������
        If MyRole.ROLES_DOC.Item(i).The_Denied = YesNo_Net Then
          For j = 1 To MyRole.ROLES_DOC.Item(i).ROLES_DOC_STATE.Count
            ' � ��������� �� ���������� ����������
            If sid = "" Then
              ' ���� ������ ��� ���������
              If MyRole.ROLES_DOC.Item(i).ROLES_DOC_STATE.Item(j).The_State Is Nothing Then
                ' �������� ����������
                GetDocumentMode = MyRole.ROLES_DOC.Item(i).ROLES_DOC_STATE.Item(j).The_Mode.Name
                On Error Resume Next
                Set os = Nothing
                Set os = Obj
                If Not os Is Nothing Then
                    If Not os.INVOS_INFO.Item(1).ostype Is Nothing Then
                        If os.INVOS_INFO.Item(1).ostype.ShowTech <> 0 Then
                        GetDocumentMode = "c" & GetDocumentMode
                        End If
                    End If
                End If
                
                
                Exit Function
              End If
            Else
              ' ���� ���������  -  ���������� ������ � ������������� ����������
              If Not MyRole.ROLES_DOC.Item(i).ROLES_DOC_STATE.Item(j).The_State Is Nothing Then
                ' �����
                If MyRole.ROLES_DOC.Item(i).ROLES_DOC_STATE.Item(j).The_State.id = sid Then
                  If MyRole.ROLES_DOC.Item(i).ROLES_DOC_STATE.Item(j).The_Mode Is Nothing Then
                     GetDocumentMode = ""
                  Else
                     ' �������� ����� ��������
                     GetDocumentMode = MyRole.ROLES_DOC.Item(i).ROLES_DOC_STATE.Item(j).The_Mode.Name
                     On Error Resume Next
                        Set os = Nothing
                        Set os = Obj
                        If Not os Is Nothing Then
                            If Not os.INVOS_INFO.Item(1).ostype Is Nothing Then
                                If os.INVOS_INFO.Item(1).ostype.ShowTech <> 0 Then
                                GetDocumentMode = "c" & GetDocumentMode
                                End If
                            End If
                        End If
                
                  End If
                  Exit Function
                End If
              End If

                  
            End If

          Next
          For j = 1 To MyRole.ROLES_DOC.Item(i).ROLES_DOC_STATE.Count
            
            ' ���� ������ ��� ���������
            If MyRole.ROLES_DOC.Item(i).ROLES_DOC_STATE.Item(j).The_State Is Nothing Then
                ' �������� ������ �����
                GetDocumentMode = MyRole.ROLES_DOC.Item(i).ROLES_DOC_STATE.Item(j).The_Mode.Name
                On Error Resume Next
                Set os = Nothing
                Set os = Obj
                If Not os Is Nothing Then
                    If Not os.INVOS_INFO.Item(1).ostype Is Nothing Then
                        If os.INVOS_INFO.Item(1).ostype.ShowTech <> 0 Then
                        GetDocumentMode = "c" & GetDocumentMode
                        End If
                    End If
                End If
                Exit Function
            End If
          Next
            
        End If
      Exit For
    End If
  Next
  
End Function


Public Function IsDocDenied(ByVal Obj As Object) As Boolean
  Dim sid As String
  Dim tn As String
  Dim mode As String
  Dim i As Long
  IsDocDenied = False
  If MyRole Is Nothing Then Exit Function
  tn = Obj.TypeName
  sid = Obj.StatusID
  For i = 1 To MyRole.ROLES_DOC.Count
    If UCase(MyRole.ROLES_DOC.Item(i).The_Document.Name) = UCase(tn) Then
      If MyRole.ROLES_DOC.Item(i).The_Denied = YesNo_Da Then
        IsDocDenied = True
        Exit Function
      End If
    End If
  Next
End Function


Public Function RoleDocAllowDelete(ByVal Obj As Object) As Boolean
  Dim sid As String
  Dim tn As String
  Dim mode As String
  Dim i As Long, j As Long
  If MyRole Is Nothing Then Exit Function
  tn = Obj.TypeName
  sid = Obj.StatusID
  RoleDocAllowDelete = True
  For i = 1 To MyRole.ROLES_DOC.Count
    If UCase(MyRole.ROLES_DOC.Item(i).The_Document.Name) = UCase(tn) Then
      If MyRole.ROLES_DOC.Item(i).AllowDeleteDoc = YesNo_Net Then
        RoleDocAllowDelete = False
        For j = 1 To MyRole.ROLES_DOC.Item(i).ROLES_DOC_STATE.Count
          If sid <> "" Then
            If Not MyRole.ROLES_DOC.Item(i).ROLES_DOC_STATE.Item(j).The_State Is Nothing Then
              If MyRole.ROLES_DOC.Item(i).ROLES_DOC_STATE.Item(j).The_State.id = sid Then
                If MyRole.ROLES_DOC.Item(i).ROLES_DOC_STATE.Item(j).AllowDelete = Boolean_Net Then
                  RoleDocAllowDelete = False
                Else
                  RoleDocAllowDelete = True
                End If
                Exit For
              End If
            End If
          End If
        Next
        Exit Function
      End If
    End If
  Next
End Function

Public Function RoleDocCanSwitchStatus(ByVal Obj As Object) As Boolean
  Dim sid As String
  Dim tn As String
  Dim mode As String
  Dim i As Long, j As Long
  If MyRole Is Nothing Then Exit Function
  tn = Obj.TypeName
  sid = Obj.StatusID
  RoleDocCanSwitchStatus = True
  For i = 1 To MyRole.ROLES_DOC.Count
    If UCase(MyRole.ROLES_DOC.Item(i).The_Document.Name) = UCase(tn) Then
        For j = 1 To MyRole.ROLES_DOC.Item(i).ROLES_DOC_STATE.Count
          If sid <> "" Then
            If Not MyRole.ROLES_DOC.Item(i).ROLES_DOC_STATE.Item(j).The_State Is Nothing Then
              If MyRole.ROLES_DOC.Item(i).ROLES_DOC_STATE.Item(j).The_State.id = sid Then
                If MyRole.ROLES_DOC.Item(i).ROLES_DOC_STATE.Item(j).StateChangeDisabled = Boolean_Da Then
                  RoleDocCanSwitchStatus = False
                Else
                  RoleDocCanSwitchStatus = True
                End If
                Exit For
              End If
            End If
          End If
        Next
        Exit Function
    End If
  Next
End Function


Public Function ARMID() As String
   ARMID = "{6DC47FFF-C1CD-4CA3-AA6F-3B15624A5A67}"
End Function



