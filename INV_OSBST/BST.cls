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

Dim theItem As INV_OS.Application
Dim Manager As MTZManager.Main
Dim session As MTZSession.session
Dim conn As ADODB.Connection
  





Public Function Check(Item As Object, ByVal NewStatusid As String, ByVal CurrentUser As Object, ByVal Module As String) As Boolean
  Dim result As Boolean
  Dim ID As String
  result = True
  
  On Error Resume Next
  Set theItem = Item
  Set Manager = theItem.Manager
  Set session = theItem.MTZSession
  Dim gui As Object
  
 
  Dim csstr As String
  Dim i As Long
  

' состояния для типа:INV_OS Карточка основного средства
' "{2AA78799-2880-4541-99E0-3C8750AC33E6}" 'В аренде
' "{72195EFD-2052-4539-AB55-1D7E6B3AA767}" 'В лизинге
' "{8E6E78D2-82AA-4913-B08C-1230A8C8B4A9}" 'В ремонте
' "{8AD15E54-CF87-4FCF-8A1E-A85336E23C73}" 'В эксплуатации
' "{DA1E3744-00B3-4D9E-AA07-BE499D2402E4}" 'На консервации
' "{55270A15-FA1D-4121-860B-A1B697B40A40}" 'На модернизации
' "{179CB53A-CBE7-46B4-9905-22E35FAAE801}" 'Оформляется
' "{166D4978-0C4C-4575-8192-B251AC113781}" 'Списано
    
    
    If NewStatusid = "{8AD15E54-CF87-4FCF-8A1E-A85336E23C73}" Then  ' в Эксплуатации
        If theItem.INVOS_CODE.Count = 0 Then
            With theItem.INVOS_CODE.Add
             Dim s As String
             s = Replace(.ID, "-", "")
             s = Replace(s, "{", "")
             s = Replace(s, "}", "")
            .ShCode = MTZUtil.Code128(s)
            .VisibleCode = s
            .Save
            End With
        End If
    
        If theItem.StatusID = "{179CB53A-CBE7-46B4-9905-22E35FAAE801}" Then
            ' проверить обязательные к заполнению поля
            result = True
        End If
        
        If theItem.StatusID = "{2AA78799-2880-4541-99E0-3C8750AC33E6}" Then  ' в аренде
            ' ищем строку где не задана дата завершения аренды
            ' и задаем там дату завершения аренды
            For i = 1 To theItem.INVOS_RENT.Count
                If theItem.INVOS_RENT.Item(i).EndDate = 0 Or theItem.INVOS_RENT.Item(i).EndDate < theItem.INVOS_RENT.Item(i).StartDate Then
                    theItem.INVOS_RENT.Item(i).EndDate = Date
                    theItem.INVOS_RENT.Item(i).Save
                    Exit For
                End If
            Next
        
        End If
        If theItem.StatusID = "{8E6E78D2-82AA-4913-B08C-1230A8C8B4A9}" Then   ' в ремонте
        
            ' ищем строку где не задана дата завершения ремонта
            ' и задаем там дату завершения ремонта
            For i = 1 To theItem.INVOS_REPAIR.Count
                If theItem.INVOS_REPAIR.Item(i).EndDate = 0 Or theItem.INVOS_REPAIR.Item(i).EndDate < theItem.INVOS_REPAIR.Item(i).StartDate Then
                    theItem.INVOS_REPAIR.Item(i).EndDate = Date
                    theItem.INVOS_REPAIR.Item(i).Save
                    Exit For
                End If
            Next
        End If
    End If
    
    If NewStatusid = "{2AA78799-2880-4541-99E0-3C8750AC33E6}" Then  ' в аренде
        
            Dim rent As INV_OS.INVOS_RENT
            Set rent = theItem.INVOS_RENT.Add
            With rent
                .StartDate = Date
                .Save
            End With
            Set gui = Manager.GetInstanceGUI(theItem.ID)
            If gui.ShowAddForm("", rent) Then
                rent.Save
            Else
                rent.Delete
                result = False
            End If
        
    End If
    
    If NewStatusid = "{8E6E78D2-82AA-4913-B08C-1230A8C8B4A9}" Then  ' в ремонте
        
            Dim rep As INV_OS.INVOS_REPAIR
            Set rep = theItem.INVOS_REPAIR.Add
            With rep
            
                .StartDate = Date
                .Save
            End With
                Set gui = Manager.GetInstanceGUI(theItem.ID)
            If gui.ShowAddForm("", rep) Then
                rep.Save
            Else
                rep.Delete
                result = False
            End If
        
    End If
    
    
    If NewStatusid = "{55270A15-FA1D-4121-860B-A1B697B40A40}" Then  ' На модернизации
        
            Dim moder As INV_OS.INVOS_MOD
            Set moder = theItem.INVOS_MOD.Add
            With moder
                .StartDate = Date
                .Save
            End With
                Set gui = Manager.GetInstanceGUI(theItem.ID)
            If gui.ShowAddForm("", moder) Then
                moder.Save
            Else
                moder.Delete
                result = False
            End If
        
    End If
    
    
    
    
     If NewStatusid = "{72195EFD-2052-4539-AB55-1D7E6B3AA767}" Then  'В лизинге
        
            Dim liz As INV_OS.INVOS_LIZING
            Set liz = theItem.INVOS_LIZING.Add
            With liz
                .Save
            End With
                Set gui = Manager.GetInstanceGUI(theItem.ID)
            If gui.ShowAddForm("", liz) Then
                liz.Save
            Else
                liz.Delete
                result = False
            End If
        
    End If
    
    

         If NewStatusid = "{DA1E3744-00B3-4D9E-AA07-BE499D2402E4}" Then  'На консервации
        
            Dim CNSRV As INV_OS.INVOS_CNSRV
            Set CNSRV = theItem.INVOS_CNSRV.Add
            With CNSRV
                .StartDate = Date
                .Save
            End With
                Set gui = Manager.GetInstanceGUI(theItem.ID)
            If gui.ShowAddForm("", CNSRV) Then
                CNSRV.Save
            Else
                CNSRV.Delete
                result = False
            End If
        
    End If
    
    
    If NewStatusid = "{166D4978-0C4C-4575-8192-B251AC113781}" Then  ' списано
    
            Dim roff As INV_OS.INVOS_OFFRULE
            If theItem.INVOS_OFFRULE.Count = 0 Then
                theItem.INVOS_OFFRULE.Add
            End If
            Set roff = theItem.INVOS_OFFRULE.Item(1)
            With roff
                .DocDate = Date
                .Save
            End With
            Set gui = Manager.GetInstanceGUI(theItem.ID)
            If gui.ShowAddForm("", roff) Then
                roff.Save
                result = True
            Else
                roff.Refresh
                result = False
            End If
    End If
    
    
  Check = result
End Function

