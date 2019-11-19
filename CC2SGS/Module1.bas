Attribute VB_Name = "Module1"
Dim ArgArray() As String
Private Type ACCDEF
  CODE As String
  FromPath As String
  used As Integer
End Type


Public Manager As MTZManager.Main
Public Session As MTZSession.Session
Public Status As String
Dim site As String
Dim login As String
Dim password As String
Dim txtFile As String
Dim ad() As ACCDEF, maxAD As Long, P() As String, maxP As Long
Public usr As MTZUsers.Application
Public MyUser As MTZUsers.Users
Public dic As INV_DIC.Application
Public invNum As INV_NUM.Application

Public Sub NextVal()
'    ' ' pb.Value = (' ' pb.Value + 1) Mod 100
End Sub


Public Sub Main()
Dim cnt As Integer, i As Integer
Dim arr() As String
  
  If Command = "" Then
   MsgBox "Вызов: СС2SGS paramfile APP:<site>  USR:<user name> PWD:<password> [LOG:<log file>]", vbOKOnly, "Загрузчик технической информации"
   End
  Else
    cnt = GetCommandLine
    For i = 2 To cnt
      arr() = Split(ArgArray(i), ":")
      If UBound(arr) >= 1 Then
      If UCase(arr(0)) = "APP" Then
        site = arr(1)
      End If
      If UCase(arr(0)) = "USR" Then
        login = arr(1)
      End If
      
      If UCase(arr(0)) = "PWD" Then
        password = arr(1)
      End If
      If UCase(arr(0)) = "LOG" Then
        LogPath = Mid(ArgArray(i), 5)
      End If
    End If
    Next
    
    
    
    Set Manager = New MTZManager.Main
    Set Session = Manager.GetSession(site)
    If Session Is Nothing Then
      End
    End If
    
    If Not Session.login(login, password) Then
      Set Session = Nothing
      End
    End If
    
    Dim rs As ADODB.Recordset
    Set rs = Manager.ListInstances(site, "MTZUsers")
    Set usr = Manager.GetInstanceObject(rs!InstanceID)
    Manager.LockInstanceObject usr.id
    
    
    Set rs = Nothing
    Set MyUser = usr.FindRowObject("Users", Session.GetSessionUserID())
    Set rs = Nothing
      
    Set W = New Writer
    If LogPath <> "" Then
      W.SetFilePath LogPath
    End If
    ProcessAll ArgArray(1)
    Session.Logout
    Set Session = Nothing
    Set Manager = Nothing
    End
  End If
End Sub


Function GetCommandLine(Optional MaxArgs) As Long
    On Error Resume Next
    'Declare variables.
    Dim c, CmdLine, CmdLnLen, InArg, i, NumArgs
    'See if MaxArgs was provided.
    If IsMissing(MaxArgs) Then MaxArgs = 20
    'Make array of the correct size.
    ReDim ArgArray(MaxArgs)
    NumArgs = 0: InArg = 0
    'Get command line arguments.
    CmdLine = Command()
    CmdLnLen = Len(CmdLine)
    'Go thru command line one character
    'at a time.
    For i = 1 To CmdLnLen
        c = Mid(CmdLine, i, 1)

        'Test for space or tab.
        If (InArg < 2 And c <> " " And c <> vbTab) Or _
            (InArg = 2 And c <> """") Or _
            (InArg = 3 And c <> "'") Then
        
            'Neither space nor tab.
            'Test if already in argument.
            If InArg = 0 Then
            'New argument begins.
            'Test for too many arguments.
                If NumArgs = MaxArgs Then Exit For
                NumArgs = NumArgs + 1
                If c = """" Then
                    InArg = 2
                    GoTo nnn
                ElseIf c = "'" Then
                    InArg = 3
                    GoTo nnn
                Else
                    InArg = 1
                End If
                ArgArray(NumArgs) = ""
            End If
            'Concatenate character to current argument.
            ArgArray(NumArgs) = ArgArray(NumArgs) & c
        Else
            
            InArg = 0
        End If
nnn:
    Next i
    'Resize array just enough to hold arguments.
    ReDim Preserve ArgArray(NumArgs)
    
    'Return Array in Function name.
    GetCommandLine = NumArgs
End Function





Public Sub ProcessAll(ByVal FilePath As String)
  Dim fd As Long, i, j, mypath, myname, str
  Dim t As Boolean
  If FilePath = "" Then Exit Sub
  Status = ""
  On Error GoTo noFile
  fd = FreeFile
  Open FilePath For Input Access Read As #fd
  On Error GoTo closefile
  maxAD = 1
  SetStatus "Читаем настроечный файл"
  Do While Not EOF(fd)
    ReDim Preserve ad(maxAD)
    Input #fd, ad(maxAD).CODE, ad(maxAD).FromPath
    ad(maxAD).used = 0
    SetStatus "Читаем настроечный файл строка № " & maxAD
    maxAD = maxAD + 1
  Loop
  
  Close fd
  On Error GoTo noFile

  Dim os As INV_OS.Application
  Dim rs As ADODB.Recordset
  
  
  
  'loop for all known path
  SetStatus "Начало загрузки"
  For i = 1 To maxAD - 1
      SetStatus "Загрузка файла " & ad(i).FromPath
      Set rs = Session.GetData("select instanceid from v_autoinvos_info where invos_code_visiblecode='" & ad(i).CODE & "'")
      If Not rs Is Nothing Then
        If Not rs.EOF Then
          Set os = Manager.GetInstanceObject(rs!instaceid)
          If Not os Is Nothing Then
            If Not IsFileLoaded(ad(i).FromPath, md5, "ТЕХНИЧЕСКАЯ ИНФОРМАЦИЯ") Then
              If Not LoadTech(os, ad(i).FromPath) Then
                RegisterFile ad(i).FromPath, md5, "ТЕХНИЧЕСКАЯ ИНФОРМАЦИЯ"
              End If
            End If
          End If
        End If
      End If
      

  Next
  SetStatus "Закончена загрузка"
  
  
   On Error GoTo noFile
  ' make list of source paths
  
closefile:
    Close fd
    Resume Next

noFile:
End Sub

Public Sub SetStatus(ByVal s As String)
  Status = s & vbCrLf & Status
  Debug.Print s
  DoEvents
End Sub



Private Function LoadTech(os As INV_OS.Application, ByVal ThePath As String) As Boolean
    Dim s As String
    Dim ff As Integer
    Dim arr() As String
    Dim spath As String
    LoadTech = False
    
    spath = ThePath
    If spath = "" Then
      spath = os.INVOS_INFO.Item(1).TechFilePath
    End If
    
  
    If spath <> "" Then
        ff = FreeFile
        On Error GoTo bye
        Open spath For Input As #ff
        os.INVOS_INFO.Item(1).TechFilePath = spath
        os.INVOS_INFO.Item(1).save
        s = input(LOF(ff), ff)
        Close #ff
        arr = Split(s, vbCrLf)
        Dim i As Integer
        For i = LBound(arr) To UBound(arr)
            If Left(arr(i), 1) = "[" Then
                If Not fProg Is Nothing Then
                  fProg.NextVal os.Brief & "->" & arr(i)
                End If
                If arr(i) = "[Info]" Then
                 
                    Call LoadOSInfo(arr, i, os)
                End If
                If arr(i) = "[Computer]" Then
                 Call LoadOSComputer(arr, i, os)
                End If
                If arr(i) = "[Current_Config]" Then
                 Call LoadOSCurCFG(arr, i, os)
                End If
                If arr(i) = "[Windows_Soft]" Then
                 Call LoadOSWinSoft(arr, i, os)
                End If
                If arr(i) = "[Windows_Devices]" Then
                 Call LoadOSWinDev(arr, i, os)
                End If
                
                If arr(i) = "[Config_changes]" Then
                 Call LoadCC(arr, i, os)
                End If
                 If arr(i) = "[Hardware]" Then
                 Call LoadHard(arr, i, os)
                End If
                
            End If
        Next
      LoadTech = True
    End If
      
bye:
    If Not fProg Is Nothing Then
            fProg.NextVal os.Brief & "->" & Err.Description
            Err.Clear
            'If MsgBox("Удалить данные о пути к файлу с Тех. информацией?", vbQuestion + vbYesNo, "Ошибка загрузки данных") = vbYes Then
            '  os.INVOS_INFO.Item(1).TechFilePath = ""
            '  os.INVOS_INFO.Item(1).save
            'End If
    End If
    LoadTech = False
End Function

Private Sub LoadOSInfo(arr() As String, StartIdx As Integer, os As INV_OS.Application)
    If os.INVOS_TECH.Count = 1 Then
        os.INVOS_TECH.Delete 1
    End If
    Dim bbb() As String
    Dim i As Integer
    With os.INVOS_TECH.Add
        For i = StartIdx + 1 To UBound(arr)
         If Left(arr(i), 1) = "[" Then
          .save
          Exit For
         End If
         bbb = Split(arr(i), "=")
         Select Case bbb(0)
         Case "MAC_Addr"
            .MAC_Addr = bbb(1)
         Case "Current_User_Name"
         .Current_User_Name = bbb(1)
         Case "Computer_Name"
         .Computer_Name = bbb(1)
         Case "IP_Addr"
         .IP_Addr = bbb(1)
         Case "System"
         .System = bbb(1)
         Case "Record_Date"
         .Record_Date = bbb(1)
         End Select
         
        Next
    End With
End Sub

Private Sub LoadOSComputer(arr() As String, StartIdx As Integer, os As INV_OS.Application)
 Dim bbb() As String
    Dim i As Integer
        For i = StartIdx + 1 To UBound(arr)
         If Left(arr(i), 1) = "[" Then
          Exit For
         End If
         bbb = Split(arr(i), "=")
            With os.INVOS_TECH.Item(1).INVOS_COMP.Add
            .Name = bbb(0)
            .ParamValue = bbb(1)
            .save
            End With
        Next
    
End Sub

Private Sub LoadOSCurCFG(arr() As String, StartIdx As Integer, os As INV_OS.Application)
Dim bbb() As String
    Dim i As Integer
        For i = StartIdx + 1 To UBound(arr)
         If Left(arr(i), 1) = "[" Then
          Exit For
         End If
         bbb = Split(arr(i), "=")
            With os.INVOS_TECH.Item(1).INVOS_CURCFG.Add
            .Name = bbb(0)
            .ParamValue = bbb(1)
            .save
            End With
        Next
End Sub

Private Sub LoadOSWinSoft(arr() As String, StartIdx As Integer, os As INV_OS.Application)
Dim bbb() As String
    Dim i As Integer
        For i = StartIdx + 1 To UBound(arr)
         If Left(arr(i), 1) = "[" Then
          Exit For
         End If
         bbb = Split(arr(i), "=")
            With os.INVOS_TECH.Item(1).INVOS_WINSOFT.Add
            .Name = bbb(0)
            .ParamValue = bbb(1)
            .save
            End With
        Next
End Sub


Private Sub LoadOSWinDev(arr() As String, StartIdx As Integer, os As INV_OS.Application)
Dim bbb() As String
    Dim i As Integer
        For i = StartIdx + 1 To UBound(arr)
         If Left(arr(i), 1) = "[" Then
          Exit For
         End If
         bbb = Split(arr(i), "=")
            With os.INVOS_TECH.Item(1).INVOS_DEVICES.Add
            .Name = bbb(0)
            .ParamValue = bbb(1)
            .save
            End With
        Next
End Sub

Private Sub LoadCC(arr() As String, StartIdx As Integer, os As INV_OS.Application)
On Error Resume Next
Dim bbb() As String
Dim ccc() As String
    Dim i As Integer
        For i = StartIdx + 1 To UBound(arr)
         If Left(arr(i), 1) = "[" Then
          Exit For
         End If
         bbb = Split(arr(i), "=")
         ccc = Split(bbb(0), " ")
            With os.INVOS_TECH.Item(1).INVOS_CFGCH.Add
            .ChDate = CDate(ccc(0) & " " & ccc(1))
            .ChNum = ccc(2)
            .TheValue = bbb(1)
            .save
            End With
        Next
End Sub

Private Sub LoadHard(arr() As String, StartIdx As Integer, os As INV_OS.Application)
Dim bbb() As String
    Dim i As Integer
        For i = StartIdx + 1 To UBound(arr)
         If Left(arr(i), 1) = "[" Then
          Exit For
         End If
         If arr(i) <> "" Then
         bbb = Split(arr(i), "=")
            With os.INVOS_TECH.Item(1).INVOS_HARD.Add
            .Name = bbb(0)
            .ParamValue = bbb(1)
            .save
            End With
        End If
        Next
End Sub


Public Sub RegisterFile(ByVal Filename As String, ByVal md5 As String, ByVal FileType As String)

  Dim invf As invf.Application
  Dim id As String
  id = CreateGUID2
  
  Manager.NewInstance id, "INVF", Filename
  Set invf = Manager.GetInstanceObject(id)
  With invf.INVF_DEF.Add
    .ThePath = Filename
    .TheHash = md5
    Set .TheUser = MyUser
    .TypeOfFile = FileType
    .Loaddate = Now
    .save
  End With
End Sub


Public Function IsFileLoaded(ByVal Filename As String, ByVal md5 As String, ByVal FileType As String) As Boolean
  Dim res As Boolean
  Dim rs As ADODB.Recordset
  Set rs = Session.GetData("select count(*) cnt from INVF_DEF where thePath ='" & Filename & "' and theHash='" & md5 & "' and TypeOfFile='" & FileType & "'")
  If rs Is Nothing Then
    res = False
  ElseIf rs.EOF Then
    res = False
  ElseIf rs!cnt = 0 Then
    res = False
  Else
    res = True
  End If
  rs.Close
  Set rs = Nothing
  IsFileLoaded = res
End Function
