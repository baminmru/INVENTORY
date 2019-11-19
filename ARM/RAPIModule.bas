Attribute VB_Name = "RAPIModule"
Option Explicit
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal _
hHandle As Long, ByVal dwMilliseconds As Long) As Long

Public Const ONE_SECOND = 1000
Public Const E_FAIL = &H80004005

Public Const FILE_ATTRIBUTE_NORMAL = &H80

Public Const INVALID_HANDLE_VALUE = -1

Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000

Public Const CREATE_NEW = 1
Public Const CREATE_ALWAYS = 2
Public Const OPEN_EXISTING = 3

Public Const ERROR_FILE_EXISTS = 80
Public Const ERROR_INVALID_PARAMETER = 87
Public Const ERROR_DISK_FULL = 112

Public Type CEOSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Public Type RAPIINIT
    cbSize As Long
    heRapiInit As Long
    hrRapiInit As Long
End Type

Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
End Type

Public Type MyType
    Value As Integer
End Type

Public Declare Function CeCloseHandle Lib "rapi.dll" ( _
    ByVal hObject As Long) As Boolean
        
Public Declare Function CeDeleteFile Lib "rapi.dll" ( _
    ByVal lpFileName As String) As Boolean
    
 Public Declare Function CeRemoveDirectory Lib "rapi.dll" ( _
    ByVal lpFileName As String) As Boolean
        
Public Declare Function CeCreateFile Lib "rapi.dll" ( _
    ByVal lpFileName As String, _
    ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, _
    lpSecurityAttributes As SECURITY_ATTRIBUTES, _
    ByVal dwCreationDistribution As Long, _
    ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As Long) As Long
        
Public Declare Function CeGetVersionEx Lib "rapi.dll" ( _
    lpVersionInformation As CEOSVERSIONINFO) As Boolean
        
Public Declare Function CeRapiInitEx Lib "rapi.dll" ( _
    pRapiInit As RAPIINIT) As Long
    
Public Declare Function CeRapiUninit Lib "rapi.dll" () As Long

Public Declare Function CeReadFile Lib "rapi.dll" ( _
    ByVal hFile As Long, _
    lpBuffer As Any, _
    ByVal nNumberOfBytesToRead As Long, _
    lpNumberOfBytesRead As Long, _
    ByVal lpOverlapped As Long) As Boolean
    
Public Declare Function CeWriteFile Lib "rapi.dll" ( _
    ByVal hFile As Long, _
    lpBuffer As Any, _
    ByVal nNumberOfBytesToWrite As Long, _
    lpNumberOfBytesWritten As Long, _
    ByVal lpOverlapped As Long) As Boolean
    
Public Declare Function CeGetLastError Lib "rapi.dll" () As Long

'Wrapper functions for above API calls

Private Function GetSub(Addr As Long) As Long
    'Used for the init call.
    GetSub = Addr
End Function

Public Sub ConnectedRapi()
    'Used for the init call. Do not remove.
End Sub

Public Function RapiConnect() As Boolean
    'Initiates a connection and returns true
    ' if it connected, false if it did not.
    
    'Modified to match suggestion of Microsoft KB article 831883
    'http://support.microsoft.com/default.aspx?scid=kb;en-us;831883

    Dim pRapiInit As RAPIINIT
    Dim dwWaitRet, dwTimeout As Long
    Dim hr As Long

    On Error GoTo RapiConnect_Err
    
    pRapiInit.cbSize = Len(pRapiInit)
    pRapiInit.heRapiInit = 0
    pRapiInit.hrRapiInit = 0

    hr = E_FAIL
    dwWaitRet = 0
    dwTimeout = 10 * ONE_SECOND 'However long you want to wait

    'Call CeRapiInitEx one time.
     hr = CeRapiInitEx(pRapiInit)
   
    If hr < 0 Then 'FAILED
      GoTo Failed
    End If
   
   'Wait for the RAPI event until timeout.

   'Use the WaitForSingleObject function for the worker thread
   'Use the WaitForMultipleObjects function if you are also waiting for other events.
   
   dwWaitRet = WaitForSingleObject(pRapiInit.heRapiInit, dwTimeout)
   
   If dwWaitRet = 0 Then 'WAIT_OBJECT_0
      'If the RAPI init is returned, check result
      
      If pRapiInit.hrRapiInit >= 0 Then 'SUCCEEDED
        GoTo Succeeded
      Else
        GoTo Failed
      End If
   Else
      'Timeout or failed.
      GoTo Failed
   
   End If
   
   'success
Succeeded:
     'Now you can make RAPI calls.
     RapiConnect = True
     Exit Function
Failed:
     'Uninitialize RAPI if you ever called CeRapiInitEx.
     If hr >= 0 Then 'SUCCEEDED
       Call CeRapiUninit
     End If
    
   RapiConnect = False
   Exit Function
    
RapiConnect_Err:
    RapiConnect = False
End Function

Public Sub RAPICopyCEFileToPC(ByVal CESourceFile As String, _
        ByVal PCDestFile As String)
        
    Dim lCeFileHandle   As Long
    Dim iFile           As Integer
    Dim BytePos         As Long
    Dim lBufferLen      As Long
    Dim lBytesRead      As Long
    Dim bytFile(2048)   As Byte
    Dim lResult         As Long
    Dim i               As Integer
    
    ' Open the CE file.
    lCeFileHandle = RapiOpenFile(CESourceFile, 1, False, _
            FILE_ATTRIBUTE_NORMAL)
            
    If lCeFileHandle <> INVALID_HANDLE_VALUE Then
        'Create a file on the PC and write
        ' the bytes from the CE file to it.
        iFile = FreeFile
        Open PCDestFile For Binary Access Write As iFile
        BytePos = 1
        lBufferLen = 2048
        Do
            lResult = CeReadFile(lCeFileHandle, bytFile(0), _
                    lBufferLen, lBytesRead, 0&)
                    
            If (lResult And (lBytesRead = 0)) Then
                lResult = CeCloseHandle(lCeFileHandle)
                Close iFile
                Exit Do
            Else
                For i = 0 To lBytesRead - 1
                    Put iFile, BytePos + i, bytFile(i)
                Next i
                BytePos = BytePos + lBytesRead
            End If
         
        Loop
    
                
    Else
        lResult = CeCloseHandle(lCeFileHandle)
        MsgBox "Device File Does Not Exist Or Is Empty (0 Bytes)!"
    End If
End Sub

Public Sub RAPICopyPCFileToCE(ByVal PCSourceFile As String, _
        ByVal CEDestFile As String)
        
    Dim iFile As Integer
    Dim bytFile() As MyType
    Dim lCeFileHandle As Long
    Dim BytePos As Long
    Dim lBufferLen As Long
    Dim TotalCopied As Long
    Dim lBytesWritten As Long
    Dim lResult As Long
    
    'Get bytes from PC file.
    iFile = FreeFile
    Open PCSourceFile For Binary Access Read As iFile
        ReDim bytFile(LOF(iFile))
        Get iFile, , bytFile
    Close iFile
    
    'Create a file on the CE Device and write
    ' the bytes from the PC file to it.
    lCeFileHandle = RapiOpenFile(CEDestFile, 2, True, FILE_ATTRIBUTE_NORMAL)
            
    If lCeFileHandle <> INVALID_HANDLE_VALUE Then
        BytePos = 0
        
        'Copy this many bytes at a time (MUST BE EVEN #).
        lBufferLen = 2048
        Do
            If UBound(bytFile) - TotalCopied > lBufferLen Then
                ' Copy the next set of bytes
                lResult = CeWriteFile(lCeFileHandle, bytFile(BytePos), _
                        lBufferLen, lBytesWritten, 0&)
                        
                TotalCopied = TotalCopied + lBytesWritten
                ' Unicode compensation.
                BytePos = BytePos + (lBufferLen \ 2)
'                Form1.Label6.Caption = "Bytes Copied: " & _
'                        TotalCopied & " Down."
'
'                Form1.Label6.Refresh
            Else
                ' Copy the remaining bytes if greater than 0
                lBufferLen = UBound(bytFile) - TotalCopied
                If lBufferLen > 0 Then
                    ' Copy remaining bytes at one time.
                    lResult = CeWriteFile(lCeFileHandle, _
                           bytFile(BytePos), lBufferLen, lBytesWritten, 0&)
 
                End If
                TotalCopied = TotalCopied + lBytesWritten
'                Form1.Label6.Caption = "Bytes Copied: " & _
'                        TotalCopied & " Down."
'
'                Form1.Label6.Refresh
                Exit Do
            End If
        Loop
    Else
        'CeCreateFile failed.  Why?
        Select Case CeGetLastError
            Case ERROR_FILE_EXISTS
                MsgBox "A file already exists with the specified name."
            Case ERROR_INVALID_PARAMETER
                MsgBox "A parameter was invalid."
            Case ERROR_DISK_FULL
                MsgBox "Disk if Full."
            Case Else
                MsgBox "An unknown error occurred."
        End Select
    End If
    'Form1.Label6.Caption = Form1.Label6.Caption & " Transfer Completed."
    lResult = CeCloseHandle(lCeFileHandle)
End Sub

Public Sub RapiDisconnect()
    Call CeRapiUninit
End Sub

Public Function RapiGetCEOSVersionString() As String
    ' Returns the Major, Minor, and Build number of the OS In a string.
    Dim ceosver As CEOSVERSIONINFO
    
    ceosver.dwOSVersionInfoSize = Len(ceosver)
    
    If CeGetVersionEx(ceosver) Then
        RapiGetCEOSVersionString = ceosver.dwMajorVersion & "." & _
            ceosver.dwMinorVersion & "." & _
            ceosver.dwBuildNumber & " " & _
            Left$(ceosver.szCSDVersion, _
            InStr(ceosver.szCSDVersion, Chr$(0)) - 1)
    Else
        RapiGetCEOSVersionString = ""
    End If
End Function

Public Function RapiIsConnected() As Boolean
    ' Returns whether there is a RAPI connection. If the Version
    'string is returned then we know we have a valid connection.
    RapiIsConnected = RapiGetCEOSVersionString <> ""
End Function

Public Function RapiOpenFile(ByVal fileName As String, _
        ByVal mode As Integer, _
        ByVal CreateNew As Boolean, _
        ByVal flags As Long) As Long
            
    Dim lReturn As Long
    Dim lFileMode As Long
    Dim Security As SECURITY_ATTRIBUTES
    Dim CreateDist As Long
    
    Select Case mode
        Case 1: lFileMode = GENERIC_READ
        Case 2: lFileMode = GENERIC_WRITE
        Case 3: lFileMode = GENERIC_READ Or GENERIC_WRITE
    End Select
    
    If CreateNew Then
        CreateDist = CREATE_ALWAYS 'CREATE_NEW
    Else
        CreateDist = OPEN_EXISTING
    End If
    
    lReturn = CeCreateFile(StrConv(fileName, vbUnicode), lFileMode, _
            0, Security, CreateDist, flags, 0&)
        
    RapiOpenFile = lReturn
End Function

Function FileExists(ByVal sFilename As String) As Boolean
'This function will check to make sure that a file exists. It will
'return True if the file was found and False if it was not found.
'Example: If Not FileExists("autoexec.bat") Then...

    Dim i As Integer
    
    On Error Resume Next
    
    i = Len(Dir$(sFilename))
    If Err Or i = 0 Or Trim(sFilename) = "" Then
        FileExists = False
    Else
        FileExists = True
    End If
End Function
        



