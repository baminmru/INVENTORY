Attribute VB_Name = "LoadOS"
Option Explicit
Dim MatRows As Collection
Public Function LoadXLS_OS(ByVal path As String, ByVal ORG As String) As Boolean
    Dim res As Boolean
    res = True
    
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
    
    Dim ex As Object 'excel.Application
    Dim wb As Object 'excel.Workbook
    Dim ws As Object 'excel.Worksheet
    Dim rng As Object 'excel.Range
    Dim r As Long
    Dim c As Long
    Dim os As INV_OS.Application
    Dim inf As INVOS_INFO
    Dim Doc As INVOS_DOCS
    Dim theorg As INVD_ORG
    
    
    If ORG <> "" Then
      Set theorg = FindORGByName(ORG)
      
      If theorg Is Nothing Then
        LoadXLS_OS = False
        Exit Function
      End If
    End If
    On Error GoTo bye
    Set wb = CreateObject(path)
    On Error Resume Next
    
    Set ws = wb.Worksheets.Item(1)
    Set rng = ws.Cells(2, 2)
    If Left(UCase(rng.Value), 8) = "ОСНОВНЫЕ" Then
        
  
   
        If ORG = "" Then
          
          Set rng = ws.Cells(2, 6)
          ORG = rng.Value
          Set theorg = FindORGByName(ORG)
          
          If theorg Is Nothing Then
            LoadXLS_OS = False
            Set rng = Nothing
            Set ws = Nothing
            wb.Close
            Set wb = Nothing
            Exit Function
          End If
        End If
      
  
        Dim q As Integer
        Dim cnum As String
        Dim Name As String
        Dim mIdx As Integer
         Dim compl As String
                      

        
        For r = 6 To 64000
            Debug.Print "1:" & r
            DoEvents
           
            Set rng = ws.Cells(r, 2)
            If rng.Value <> "конецфайла" And Trim(rng.Value) <> "" Then
      
             
             Set rng = ws.Cells(r, 14)
             Dim s As String
             s = rng.Value
             
             
             If Trim(s) = "" Then
               Set rng = ws.Cells(r, 2)
               Name = rng.Value
               Set rng = ws.Cells(r, 3)
               cnum = rng.Value
               ' записываем поступление
               Set rs = Session.GetData("select * from v_autoinvos_info where INVOS_INFO_TheOrg_ID='" & theorg.id & "' and INVOS_INFO_CardNum='" & cnum & "'")
                    If rs.EOF Then
                       id = CreateGUID2()
                       Set rng = ws.Cells(r, 2)
                       Manager.NewInstance id, "INV_OS", rng.Value
                       Set os = Manager.GetInstanceObject(id)
                       Set inf = os.INVOS_INFO.Add
                       Set inf.theorg = theorg

                   
                       
                       inf.invNum = Right("00" & (Val(theorg.NumPrefix)), 2) & Right("00000000" & cnum, 8)
                       
                       Set Doc = os.INVOS_DOCS.Add
                    Else
                      id = rs!InstanceID
                      Set os = Manager.GetInstanceObject(id)
                      Set inf = os.INVOS_INFO.Item(1)
                      If os.INVOS_DOCS.Count = 0 Then
                      os.INVOS_DOCS.Add
                      End If
                      Set Doc = os.INVOS_DOCS.Item(1)
                   End If
                       
'                       If txtINVOS_INFO_OSType.Tag <> "" Then
'                           Set inf.OSType = FindOSType(txtINVOS_INFO_OSType.Tag)
'                       End If
                       
                       inf.InLineNum = 0
                       inf.IsMaterial = Boolean_Net
                       
                       For c = 2 To 13
                           Set rng = ws.Cells(r, c)
                           Select Case c
                           Case 2
                               inf.Name = rng.Value
                               inf.ShortName = rng.Value
                               compl = GetCompl(inf.Name)
                           Case 3
                               'inf.invNum = rng.Value
                               inf.CardNum = rng.Value
                           Case 4
                                Doc.InOrderNum = rng.Value
                           Case 5
                              Doc.NaklNum = rng.Value
                           Case 6
                               Set Doc.Contragent = FindAgentByName(rng.Value)
                           Case 7
                               Doc.DogNum = Left(rng.Value, 30)
                           Case 8
                               Doc.ActivateNum = Left(rng.Value, 30)
                           Case 9
                               inf.SrokPI = Val(rng.Value)
                           Case 10
                               inf.SrokFI = Val(rng.Value)
                           Case 11
                               inf.SrokOI = Val(rng.Value)
                           Case 12
                              inf.TheCost = Val(Replace(rng.Value, ",", "."))
                           Case 13
                              inf.Info = rng.Value
                           End Select
                          
                       Next
                      
                       ' save place data
                       If os.INVOS_PLACE.Count = 0 Then
                           os.INVOS_PLACE.Add
                       End If
                       
'                       If txtINVOS_PLACE_TheOrg.Tag <> "" Then
'                           Set inf.TheOrg = FindOrg(txtINVOS_PLACE_TheOrg.Tag)
'                       End If
                       
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
'                           If txtINVOS_PLACE_DIrection.Tag <> "" Then
'                               Set .Direction = FindDir(txtINVOS_PLACE_DIrection.Tag)
'                           End If
'                           If txtINVOS_PLACE_TheHouse.Tag <> "" Then
'                               Set .TheHouse = FindBuilding(txtINVOS_PLACE_TheHouse.Tag)
'                           End If
'
'                           If txtINVOS_PLACE_TheOwner.Tag <> "" Then
'                               Set .TheOwner = FindOwner(txtINVOS_PLACE_TheOwner.Tag)
'                           End If
'                           If txtINVOS_PLACE_Uprav.Tag <> "" Then
'                               Set .Uprav = FindUPR(txtINVOS_PLACE_Uprav.Tag)
'                           End If
                           
'                            If txtMOL.Tag <> "" Then
'                               Set .MatOtv = FindOwner(txtMOL.Tag)
'                           End If
                           
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
                  
                    
              End If
            Else
                Exit For
            End If
            
            Manager.FreeAllInstanses
        Next
        
        
        ' списания
        For r = 6 To 64000
            Debug.Print "2:" & r
      
          
            
           'me caption = "Списание:" & r
            DoEvents
           
            Set rng = ws.Cells(r, 2)
            If rng.Value <> "конецфайла" And Trim(rng.Value) <> "" Then
             
          
             Set rng = ws.Cells(r, 14)
             
              If Trim(rng.Value) <> "" Then
          
               Set rng = ws.Cells(r, 2)
               Name = rng.Value
               Set rng = ws.Cells(r, 3)
               cnum = rng.Value
              
              
               Set rs = Session.GetData("select * from v_autoinvos_info where INVOS_INFO_TheOrg_ID='" & theorg.id & "' and INVOS_INFO_CardNum='" & cnum & "'")
               If Not rs.EOF Then
                    id = rs!InstanceID
                    Set os = Manager.GetInstanceObject(id)
                    Set inf = os.INVOS_INFO.Add()
                    Set rng = ws.Cells(r, 14)
                 
                    If os.INVOS_OFFRULE.Count = 0 Then
                        os.INVOS_OFFRULE.Add
                    End If
                    With os.INVOS_OFFRULE.Item(1)
                        .Info = rng.Value
                        .save
                    End With
                    os.StatusID = "{166D4978-0C4C-4575-8192-B251AC113781}"
               End If
    
              
              
              End If
            Else
                Exit For
            End If
            
          Manager.FreeAllInstanses
        Next
        
    Else
        SetStatus "Неверный формат отчета"
    End If
    
    Set rng = Nothing
    Set ws = Nothing
    wb.Close
    Set wb = Nothing
    
  
    ' ' pb.Visible = False
    

    LoadXLS_OS = res

     Exit Function
bye:
    
    SetStatus "Ошибка открытия файла." & vbCrLf & "Проверьте формат файла. Ожидается Excel 2003 и выше."
    LoadXLS_OS = False
    
    
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
    Set rs = Session.GetData("select * from INVD_UR where sortname='" & Name & "'  or fullname ='" & Name & "'")
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

Private Function FindORGByName(ByVal Name As String) As INVD_ORG
    Dim ur As INVD_ORG
    Dim rs As ADODB.Recordset
    Set rs = Session.GetData("select * from INVD_ORG where sortname='" & Name & "'  or fullname ='" & Name & "'")
    If rs.EOF Then
        Set ur = dic.INVD_ORG.Add
        ur.SortName = Name
        ur.FullName = Name
        ur.save
    Else
        Set ur = dic.INVD_ORG.Item(rs!INVD_ORGid)
        ur.FullName = Name
        ur.save
    End If
    Set FindORGByName = ur
    

End Function


Public Function LoadXLS_MAT(ByVal path As String, ByVal ORG As String) As Boolean
    Dim res As Boolean
    Dim q As Integer
    Dim cnum As String
    Dim Name As String
    Dim mIdx As Integer
    Dim theorg As INVD_ORG
    
    res = True
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
    
    
    
    If ORG <> "" Then
      
      Set theorg = FindORGByName(ORG)
      If theorg Is Nothing Then
        LoadXLS_MAT = False
        Exit Function
      End If
      
    
      Manager.LockInstanceObject id
      
      Set rs = Session.GetData("select * from invn_DEF where theORG='" & theorg.id & "'")
      If Not rs.EOF Then
        id = rs!InstanceID
        Set invNum = Manager.GetInstanceObject(id)
  
      Else
        id = CreateGUID2
        Manager.NewInstance id, "INV_NUM", "Нумерация"
        Set invNum = Manager.GetInstanceObject(id)
        With invNum.INVN_DEF.Add
          Set .theorg = theorg
          .save
        End With
        
        
      End If
      
      
    
      
      
      Manager.LockInstanceObject id
      
      invNum.LockResource False
      If invNum.IsLocked <> LockSession Then
        SetStatus "Не удалось заблокировать нумератор"
        
        Exit Function
      End If
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
   
    
     If ORG = "" Then
     Set rng = ws.Cells(2, 6)
      ORG = rng.Value
      Set theorg = FindORGByName(ORG)
      If theorg Is Nothing Then
        LoadXLS_MAT = False
        Set rng = Nothing
        Set ws = Nothing
        wb.Close
        Set wb = Nothing
        Exit Function
      End If
      
    
      Manager.LockInstanceObject id
      
      Set rs = Session.GetData("select * from invn_DEF where theORG='" & theorg.id & "'")
      If Not rs.EOF Then
        id = rs!InstanceID
        Set invNum = Manager.GetInstanceObject(id)
  
      Else
        id = CreateGUID2
        Manager.NewInstance id, "INV_NUM", "Нумерация"
        Set invNum = Manager.GetInstanceObject(id)
        With invNum.INVN_DEF.Add
          Set .theorg = theorg
          .save
        End With
        
        
      End If
      
      
    
      
      
      Manager.LockInstanceObject id
      
      invNum.LockResource False
      If invNum.IsLocked <> LockSession Then
        SetStatus "Не удалось заблокировать нумератор"
        Exit Function
      End If
    End If
    
    
    
    For r = 6 To 64000
        NextVal
  
        
        
       'me caption = r
        DoEvents
        Debug.Print "1:" & r
       
        Set rng = ws.Cells(r, 2)
        If rng.Value <> "конецфайла" And Trim(rng.Value) <> "" Then
         
         Set rng = ws.Cells(r, 3)
         q = CInt(Val(rng.Value))
         
         Set rng = ws.Cells(r, 13)
         
         If Trim(rng.Value) = "" Then
         
         
           Set rng = ws.Cells(r, 2)
           Name = rng.Value
           Set rng = ws.Cells(r, 4)
           cnum = rng.Value
           
         
         
          On Error Resume Next
          Dim mr As MatRow
          Set mr = New MatRow
          mr.Code = theorg.id & "|" & cnum & "|" & Name
          mr.Quantity = q
          If MatRows.Item(theorg.id & "|" & cnum & "|" & Name) Is Nothing Then
            MatRows.Add mr, mr.Code
            q = MatRows.Item(mr.Code).Quantity
          Else
            MatRows.Item(mr.Code).Quantity = MatRows.Item(mr.Code).Quantity + q
            q = MatRows.Item(mr.Code).Quantity
            Debug.Print mr.Code
            
          End If
         
         
           ' записываем поступление
           For mIdx = 1 To q
           
                
                Set rs = Session.GetData("select * from v_autoinvos_info where  INVOS_INFO_TheOrg_ID='" & theorg.id & "' and INVOS_INFO_CardNum='" & cnum & "' and INVOS_INFO_ShortName ='" & Name & "' and invos_info_InLineNum=" & mIdx)
                
                
                If rs.EOF Then
                   id = CreateGUID2()
                   Set rng = ws.Cells(r, 2)
                   Manager.NewInstance id, "INV_OS", rng.Value
                   Set os = Manager.GetInstanceObject(id)
                   Set inf = os.INVOS_INFO.Add
                   
                 
                  Set inf.theorg = theorg
                 
                   
                   inf.invNum = Right("00" & (Val(theorg.NumPrefix) + 50), 2) & Right("00000000" & GetNextInvNum(), 8)
                   
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
                           inf.TheCost = Val(Replace(rng.Value, ",", "."))
                        Case 16
                           If Trim(rng.Value & "") <> "" Then
                                inf.TheCost = Val(Replace(rng.Value, ",", "."))
                           End If
                       Case 12
                       Case 13
                       End Select
                      
                   Next
                   
                   inf.SrokFI = 0
                   inf.SrokPI = 12
                   inf.SrokOI = 12
                  
                   ' save place data
                   If os.INVOS_PLACE.Count = 0 Then
                       os.INVOS_PLACE.Add
                   End If
                   
'                   If txtINVOS_INFO_OSType.Tag <> "" Then
'                       Set inf.OSType = FindOSType(txtINVOS_INFO_OSType.Tag)
'                   End If
                   
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
  
      
        
       'me caption = "Списание:" & r
        Debug.Print "2:" & r
        DoEvents
       
        Set rng = ws.Cells(r, 2)
        If rng.Value <> "конецфайла" And Trim(rng.Value) <> "" Then
         
         Set rng = ws.Cells(r, 3)
         q = CInt(Val(rng.Value))
         
         Set rng = ws.Cells(r, 13)
         
          If Trim(rng.Value) <> "" Then
      
           Set rng = ws.Cells(r, 2)
           Name = rng.Value
           Set rng = ws.Cells(r, 4)
           cnum = rng.Value
          
          
           Set rs = Session.GetData("select * from v_autoinvos_info where   INVOS_INFO_TheOrg_ID='" & theorg.id & "' and INVOS_INFO_CardNum='" & cnum & "' and INVOS_INFO_ShortName ='" & Name & "' and INTSANCEStatusID <>'{166D4978-0C4C-4575-8192-B251AC113781}'")
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
    invNum.UnLockResource
    
    Else
      'MsgBox
      SetStatus "Неверный формат отчета"
       res = False
    End If
    
    
    ' pb.Visible = False
    Set rng = Nothing
    Set ws = Nothing
    wb.Close
    Set wb = Nothing

    LoadXLS_MAT = res
     Exit Function
bye:
    invNum.UnLockResource
    SetStatus "Ошибка открытия файла." & vbCrLf & "Проверьте формат файла. Ожидается Excel 2003 и выше."
    LoadXLS_MAT = False
End Function



Public Function GetCompl(Name As String) As String
  Dim i As Integer
  Dim spcPos As Integer
  spcPos = -1
  For i = Len(Name) To 1 Step -1
    If Mid(Name, i, 1) = " " Then
      spcPos = i
      Exit For
    End If
  Next
  If spcPos < 0 Then
    GetCompl = ""
    Exit Function
  End If
  If Len(Name) - spcPos > 10 Then
    GetCompl = ""
    Exit Function
  End If
  
  Dim cnum As String
  cnum = Mid(Name, spcPos + 1, Len(Name) - spcPos)
  If IsNumeric(Replace(Replace(cnum, ".", ""), ",", "")) Then
    If InStr(1, cnum, ".") > 0 Or InStr(1, cnum, ",") > 0 Then
      GetCompl = Replace(cnum, ",", ".")
    End If
  End If

End Function


Public Sub SaveHistory(RowItem As INV_OS.INVOS_PLACE)
On Error GoTo bye
'''''''''''''''' change me

 Dim cRow As INV_OS.INVOS_PLACE
 Dim hist As INV_OS.INVOS_HIST
 Set cRow = RowItem
 Set hist = cRow.Application.INVOS_HIST.Add
 
 
 With hist
    .UntilDate = Now
    Set .ChangedBy = cRow.Application.FindRowObject("Users", cRow.Application.MTZSession.GetSessionUserID)
    'Set .TheOrg = cRow.TheOrg
    .ComplNumber = cRow.ComplNumber
    Set .Direction = cRow.Direction
    Set .Uprav = cRow.Uprav
    Set .Otdel = cRow.Otdel
    Set .TheHouse = cRow.TheHouse
    Set .MatOtv = cRow.MatOtv
    .Flow = cRow.Flow
    .Room = cRow.Room
    .WorkPlaceNum = cRow.WorkPlaceNum
    Set .TheOwner = cRow.TheOwner
    .save
 End With



 Exit Sub
bye:
  
End Sub

