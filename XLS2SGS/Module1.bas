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
  
  logPath = "c:\xls2sgs.log"
  
  If Command = "" Then
   Dim f As frm
   Set f = New frm
   f.Show vbModal
   If f.OK Then
   
    site = f.txtSite
    login = f.txtUser
    password = f.txtPWD
    defFile = f.txtFile
    If f.txtLog <> "" Then
      logPath = f.txtLog
    End If
   Else
    MsgBox "�����: XLS2SGS paramfile APP:<site>  USR:<user name> PWD:<password> LOG:<log file>", vbOKOnly, "�������� ������ �� Excel"
    End
   End If
   Unload f
   Set f = Nothing
  Else
   
    
    cnt = GetCommandLine
    defFile = ArgArray(1)
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
        logPath = Mid(ArgArray(i), 5)
      End If
    End If
    Next
    
    End If
    
    Set W = New Writer
    If logPath <> "" Then
      W.SetFilePath logPath
    End If
    
    Set Manager = New MTZManager.Main
    Set Session = Manager.GetSession(site)
    If Session Is Nothing Then
      SetStatus "�������� �������� �����"
      End
    End If
    
    If Not Session.login(login, password) Then
      Set Session = Nothing
      SetStatus "������������ ������ ��� ��� ������������"
      End
    End If
    
    Dim rs As ADODB.Recordset
    Set rs = Manager.ListInstances(site, "MTZUsers")
    Set usr = Manager.GetInstanceObject(rs!InstanceID)
    Manager.LockInstanceObject usr.id
    
    
    Set rs = Nothing
    Set MyUser = usr.FindRowObject("Users", Session.GetSessionUserID())
    Set rs = Nothing
      
  
    ProcessAll defFile
    Session.Logout
    SetStatus "���������� ������"
    
    Set Session = Nothing
    Set Manager = Nothing
    Set W = Nothing
    End
  
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
  SetStatus "������ ����������� ����"
  Do While Not EOF(fd)
    ReDim Preserve ad(maxAD)
    Input #fd, ad(maxAD).theorg, ad(maxAD).FromPath   ', ad(maxAD).renameTo
    ad(maxAD).used = 0
    SetStatus "������ ����������� ���� ������ � " & maxAD
    maxAD = maxAD + 1
  Loop
  
  Close fd
  On Error GoTo noFile

  
  'loop for all known path
  SetStatus "��������� ����������"
  For i = 1 To maxAD - 1
    SetStatus "��������� ���������� " & ad(i).FromPath
    mypath = ad(i).FromPath
    myname = Dir(mypath & "\*.xls", vbNormal)
    Do While myname <> ""
      If myname <> "." And myname <> ".." Then
       
        Dim spath As String
        spath = mypath & "\" & myname
        SetStatus "���� " & spath
        
        Dim md5 As String
        Dim ccc As CMD5
        Set ccc = New CMD5
        md5 = ccc.FileMD5(spath)
        Set ccc = Nothing
        
      
        
        If Not IsFileLoaded(spath, md5, "���������") Then
          If Not LoadXLS_MAT(spath, ad(i).theorg) Then
            If Not IsFileLoaded(spath, md5, "��") Then
              If LoadXLS_OS(spath, ad(i).theorg) Then
                Set ccc = New CMD5
                md5 = ccc.FileMD5(spath)
                Set ccc = Nothing
                RegisterFile spath, md5, "��"
                SetStatus "����������� �����"
              End If
            Else
              SetStatus "���� ��� ����������"
            End If
          Else
           Set ccc = New CMD5
           md5 = ccc.FileMD5(spath)
           Set ccc = Nothing
           RegisterFile spath, md5, "���������"
           SetStatus "����������� �����"
          End If
        Else
           SetStatus "���� ��� ����������"
        End If
        
      End If
nxtName:
      myname = Dir  ' Get next entry.
    Loop
  Next
  SetStatus "��������� ������������"
  Exit Sub
  
   On Error GoTo noFile
  ' make list of source paths
  
closefile:
    Close fd
    Resume Next
    Exit Sub
noFile:
SetStatus Err.Description
 'MsgBox Err.Description
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
