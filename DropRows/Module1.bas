Attribute VB_Name = "Module1"


Dim ArgArray() As String
Private Type ACCDEF
  theorg As String
  FromPath As String
'  FileType As String
'  ToPath As String
'  renameTo As String
  used As Integer
End Type


Public Manager As MTZManager.Main
Public Session As MTZSession.Session
Public Status As String
Dim site As String
Dim login As String
Dim password As String
Dim txtFile As String
Dim logPath As String
Dim defFile As String
Dim ad() As ACCDEF, maxAD As Long, P() As String, maxP As Long
Public W As Writer
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
  
  logPath = "c:\droprows.log"
  
  
   Dim f As frm
   Set f = New frm
   f.Show vbModal
   If f.OK Then
   
    site = f.txtSite
    login = f.txtUser
    password = f.txtPWD
  
    If f.txtLog <> "" Then
      logPath = f.txtLog
    End If
   End If
   Unload f
   Set f = Nothing
    
    Set W = New Writer
    If logPath <> "" Then
      W.SetFilePath logPath
    End If
    
    Set Manager = New MTZManager.Main
    Set Session = Manager.GetSession(site)
    If Session Is Nothing Then
      SetStatus "Неверное название сайта"
      End
    End If
    
    If Not Session.login(login, password) Then
      Set Session = Nothing
      SetStatus "Неправильный пароль или имя пользователя"
      End
    End If
    
    Dim rs As ADODB.Recordset
    Set rs = Manager.ListInstances(site, "MTZUsers")
    Set usr = Manager.GetInstanceObject(rs!InstanceID)
    Manager.LockInstanceObject usr.id
    
    
    Set rs = Nothing
    Set MyUser = usr.FindRowObject("Users", Session.GetSessionUserID())
    Set rs = Nothing
      
    DropRows
  
    Session.Logout
    SetStatus "Завершение сессии"
    
    Set Session = Nothing
    Set Manager = Nothing
    Set W = Nothing
    End
  
End Sub


Private Sub DropRows()
On Error GoTo bye
Session.GetData "delete  from INVOS_RENT where instanceid in (select instanceid from invos_info where ismaterial=0 and (name ='' or name is null) and invnum='0100000000' )"
W.putBuf "delete INVOS_RENT"
Session.GetData "delete  from INVOS_HIST where instanceid in (select instanceid from invos_info where ismaterial=0 and (name ='' or name is null) and invnum='0100000000' )"
W.putBuf "delete INVOS_HIST"
Session.GetData "delete  from INVOS_OFFRULE where instanceid in (select instanceid from invos_info where ismaterial=0 and (name ='' or name is null) and invnum='0100000000' )"
W.putBuf "delete INVOS_OFFRULE"
Session.GetData "delete  from INVOS_SROK where instanceid in (select instanceid from invos_info where ismaterial=0 and (name ='' or name is null) and invnum='0100000000' )"
W.putBuf "delete INVOS_SROK"
Session.GetData "delete  from INVOS_DRAG where instanceid in (select instanceid from invos_info where ismaterial=0 and (name ='' or name is null) and invnum='0100000000' )"
W.putBuf "delete INVOS_DRAG"
Session.GetData "delete  from INVOS_DOCS where instanceid in (select instanceid from invos_info where ismaterial=0 and (name ='' or name is null) and invnum='0100000000' )"
W.putBuf "delete INVOS_DOCS"
Session.GetData "delete  from INVOS_INV where instanceid in (select instanceid from invos_info where ismaterial=0 and (name ='' or name is null) and invnum='0100000000' )"
W.putBuf "delete INVOS_INV"
Session.GetData "delete  from INVOS_LIZING where instanceid in (select instanceid from invos_info where ismaterial=0 and (name ='' or name is null) and invnum='0100000000' )"
W.putBuf "delete INVOS_LIZING"
Session.GetData "delete  from INVOS_CNSRV where instanceid in (select instanceid from invos_info where ismaterial=0 and (name ='' or name is null) and invnum='0100000000' )"
W.putBuf "delete INVOS_CNSRV"
Session.GetData "delete  from INVOS_MOD where instanceid in (select instanceid from invos_info where ismaterial=0 and (name ='' or name is null) and invnum='0100000000' )"
W.putBuf "delete INVOS_MOD"
Session.GetData "delete  from INVOS_CMNT where instanceid in (select instanceid from invos_info where ismaterial=0 and (name ='' or name is null) and invnum='0100000000' )"
W.putBuf "delete INVOS_CMNT"
Session.GetData "delete  from INVOS_CODE where instanceid in (select instanceid from invos_info where ismaterial=0 and (name ='' or name is null) and invnum='0100000000' )"
W.putBuf "delete INVOS_CODE"
Session.GetData "delete  from INVOS_REPAIR where instanceid in (select instanceid from invos_info where ismaterial=0 and (name ='' or name is null) and invnum='0100000000' )"
W.putBuf "delete INVOS_REPAIR"
Session.GetData "delete  from INVOS_PLACE where instanceid in (select instanceid from invos_info where ismaterial=0 and (name ='' or name is null) and invnum='0100000000' )"
W.putBuf "delete INVOS_PLACE"
Session.GetData "delete  from INVOS_INFO  where ismaterial=0 and (name ='' or name is null) and invnum='0100000000'"
W.putBuf "delete INVOS_INFO"
Session.GetData "delete  from instance where objtype='INV_OS' and instanceid not in (select instanceid from invos_info)"
W.putBuf "delete instance"
Exit Sub
bye:
W.putBuf Err.Description
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
    Input #fd, ad(maxAD).theorg, ad(maxAD).FromPath   ', ad(maxAD).renameTo
    ad(maxAD).used = 0
    SetStatus "Читаем настроечный файл строка № " & maxAD
    maxAD = maxAD + 1
  Loop
  
  Close fd
  On Error GoTo noFile

  
  'loop for all known path
  SetStatus "Сканируем директории"
  For i = 1 To maxAD - 1
    SetStatus "Сканируем директорию " & ad(i).FromPath
    mypath = ad(i).FromPath
    myname = Dir(mypath & "\*.xls", vbNormal)
    Do While myname <> ""
      If myname <> "." And myname <> ".." Then
       
        Dim spath As String
        spath = mypath & "\" & myname
        SetStatus "Файл " & spath
        
        Dim md5 As String
        Dim ccc As CMD5
        Set ccc = New CMD5
        md5 = ccc.FileMD5(spath)
        Set ccc = Nothing
        
      
        
        If Not IsFileLoaded(spath, md5, "МАТЕРИАЛЫ") Then
          If Not LoadXLS_MAT(spath, ad(i).theorg) Then
            If Not IsFileLoaded(spath, md5, "ОС") Then
              If LoadXLS_OS(spath, ad(i).theorg) Then
                Set ccc = New CMD5
                md5 = ccc.FileMD5(spath)
                Set ccc = Nothing
                RegisterFile spath, md5, "ОС"
                SetStatus "Регистрация файла"
              End If
            Else
              SetStatus "Файл уже загружался"
            End If
          Else
           Set ccc = New CMD5
           md5 = ccc.FileMD5(spath)
           Set ccc = Nothing
           RegisterFile spath, md5, "МАТЕРИАЛЫ"
           SetStatus "Регистрация файла"
          End If
        Else
           SetStatus "Файл уже загружался"
        End If
        
      End If
nxtName:
      myname = Dir  ' Get next entry.
    Loop
  Next
  SetStatus "Закончено сканирование"
  
  
   On Error GoTo noFile
  ' make list of source paths
  
closefile:
    Close fd
    Resume Next

noFile:
End Sub

Public Sub SetStatus(ByVal s As String)
  W.putBuf s
  Debug.Print s
  DoEvents
End Sub


Public Sub RegisterFile(ByVal Filename As String, ByVal md5 As String, ByVal FileType As String)

  Dim INVF As INVF.Application
  Dim id As String
  id = CreateGUID2
  
  Manager.NewInstance id, "INVF", Filename
  Set INVF = Manager.GetInstanceObject(id)
  With INVF.INVF_DEF.Add
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
