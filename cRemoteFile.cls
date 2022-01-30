VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRemoteFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'WinFTP, created by the KPD-Team 2000
'This file can be downloaded from http://www.allapi.net/
'For questions or comments, contact us at KPDTeam@Allapi.net

' You are free to use this code within your own applications,
' but you are expressly forbidden from selling or otherwise
' distributing this source code without prior written consent.
' This includes both posting free demo projects made from this
' code as well as reproducing the code in text or html format.

' Changes:
'    03/14/01, TPA: Increased Buffer size to 64k
'
'                   Commented Out DownloadFile() since it is not
'                   used and duplicates GetFile()
'
'                   Moved file writing from FOStart() to GetFile()
'                   to improve downloading speed, especially
'                   with large files on fast connections. (String
'                   concatination was getting out of hand and
'                   using too much memory.)
'
Public Enum tNewStatus
    nsConnected
    nsConnectionFailed
    nsOperationStarted
    nsOperationFailed
    nsOperationSuceeded
    nsDisconnected
    nsQuerySent
    nsQueryCompleted
    nsQueryFailed
End Enum
Public RemoteFile As String
Public FileData As String
Const BytesBuff As Long = 32# * 65# '1024# * 65#
Const FILE_BEGIN = 0
Const INTERNET_FLAG_RELOAD = &H80000000
Const HTTP_QUERY_CONTENT_LENGTH = 5
Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000     ' don't write this item to the cache
Const INTERNET_FLAG_DONT_CACHE = INTERNET_FLAG_NO_CACHE_WRITE
Const HTTP_VERSION = "HTTP/1.0"
Const INTERNET_DEFAULT_HTTP_PORT = 80
Const INTERNET_SERVICE_HTTP = 3
Const INTERNET_SERVICE_FTP = 1
Const INTERNET_INVALID_PORT_NUMBER = 0

Private Declare Function FtpDeleteFile Lib "wininet.dll" Alias "FtpDeleteFileA" (ByVal hFtpSession As Long, ByVal lpszFileName As String) As Boolean
Private Declare Function FtpOpenFile Lib "wininet.dll" Alias "FtpOpenFileA" (ByVal hFtpSession As Long, ByVal sBuff As String, ByVal Access As Long, ByVal Flags As Long, ByVal Context As Long) As Long
Private Declare Function FtpRenameFile Lib "wininet.dll" Alias "FtpRenameFileA" (ByVal hFtpSession As Long, ByVal lpszExisting As String, ByVal lpszNew As String) As Boolean
Private Declare Function InternetWriteFile Lib "wininet.dll" (ByVal hFile As Long, ByRef sBuffer As Byte, ByVal lNumBytesToWite As Long, dwNumberOfBytesWritten As Long) As Integer
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
Private Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Private Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal lpszUrl As String, ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUserName As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function InternetQueryDataAvailable Lib "wininet.dll" (ByVal hFile As Long, lpdwNumberOfBytesAvailable As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean
Private Declare Function InternetSetFilePointer Lib "wininet.dll" (ByVal hFile As Long, ByVal IDistanceToMove As Long, ByVal pReserved As Long, ByVal dwMoveMethod As Long, ByVal dwContext As Long) As Long
Private Declare Function HttpOpenRequest Lib "wininet.dll" Alias "HttpOpenRequestA" (ByVal hInternetSession As Long, ByVal lpszVerb As String, ByVal lpszObjectName As String, ByVal lpszVersion As String, ByVal lpszReferer As String, ByVal lpszAcceptTypes As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Private Declare Function HttpSendRequest Lib "wininet.dll" Alias "HttpSendRequestA" (ByVal hHttpRequest As Long, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal sOptional As String, ByVal lOptionalLength As Long) As Boolean
Private Declare Function HttpQueryInfo Lib "wininet.dll" Alias "HttpQueryInfoA" (ByVal hHttpRequest As Long, ByVal lInfoLevel As Long, ByVal sBuffer As String, ByRef lBufferLength As Long, ByRef lindex As Long) As Boolean
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)


Event ProgressChanged(nBytes As Long)
Event StatusChanged(NewStatus As tNewStatus, sOptionalInfo As String)
Public Sub ConvertBytes2String(ByRef BytesArray() As Byte, ByRef OutputString As String)
    On Error GoTo LOKAL_ERROR
    
    OutputString = Space(UBound(BytesArray) - LBound(BytesArray) + 1)
    CopyMemory ByVal StrPtr(OutputString), BytesArray(LBound(BytesArray)), Len(OutputString)
    OutputString = StrConv(OutputString, vbUnicode)
    
    Exit Sub
LOKAL_ERROR:
    MsgBox "cRemoteFile.ConvertBytes2String: " & err.Number & " / " & err.Description
End Sub
Public Function FormatMessage(nIndex As Long) As String
    On Error GoTo LOKAL_ERROR
    
    Select Case nIndex
        Case nsConnected
            FormatMessage = "Connected"
        Case nsConnectionFailed
            FormatMessage = "Connection Failed"
        Case nsOperationStarted
            FormatMessage = "Operation Started"
        Case nsOperationFailed
            FormatMessage = "Operation Failed"
        Case nsOperationSuceeded
            FormatMessage = "Operation Suceeded"
        Case nsDisconnected
            FormatMessage = "Disconnected"
        Case nsQuerySent
            FormatMessage = "Query Sent"
        Case nsQueryCompleted
            FormatMessage = "Query Completed"
        Case nsQueryFailed
            FormatMessage = "Query Failed"
    End Select
    
    Exit Function
LOKAL_ERROR:
    MsgBox "cRemoteFile.FormatMessage: " & err.Number & " / " & err.Description
End Function
Public Function GetFile(Connection As cConnection, LocalFileName As String) As Long
    On Error GoTo LOKAL_ERROR

    Dim FF          As Long
    Dim hFile       As Long
    Dim sBuffer     As String
    Dim Lret23        As Long
    Dim bFailed     As Boolean

    'änderung  jan 05
    If Right(LocalFileName, 9) = "not found" Then
        Exit Function
    End If


    If Connection.Connected = False Then
        RaiseEvent StatusChanged(nsOperationFailed, "Not connected")
        Exit Function
    End If
    RaiseEvent StatusChanged(nsOperationStarted, "")


    hFile = FtpOpenFile(Connection.hConnection, RemoteFile, &H80000000, INTERNET_FLAG_RELOAD, 0)
    If hFile = 0 Then  'Or 13369360
        RaiseEvent StatusChanged(nsOperationFailed, "Unable to open the specified URL.")
        Exit Function
    End If
    sBuffer = Space(BytesBuff)
    FileData = ""

    FF = FreeFile(0)
    Open LocalFileName For Binary Access Write Lock Read Write As #FF

    Do

        If InternetReadFile(hFile, sBuffer, BytesBuff, Lret23) = 0 Then
'            MsgBox Lret23
            bFailed = True
            RaiseEvent StatusChanged(nsOperationFailed, "Error while reading the remote file.")
            Exit Do
        End If


        If Lret23 <> BytesBuff Then
        
            sBuffer = Left$(sBuffer, Lret23)

        End If

        SentBytes = SentBytes + Lret23
        ActiveFileBytesSent = ActiveFileBytesSent + Lret23

        Put #FF, , sBuffer
        frmWKL38.UpdateProgress

        RaiseEvent ProgressChanged(SentBytes)

    Loop Until Lret23 <> BytesBuff



    InternetCloseHandle hFile
    Close FF
    If bFailed = False Then RaiseEvent StatusChanged(nsOperationSuceeded, "")
    sBuffer = "" ' Clear the buffer

    Exit Function
LOKAL_ERROR:
    MsgBox "cRemoteFile.GetFile: " & err.Number & " / " & err.Description
    Resume Next
End Function
Public Function GetFile1(Connection As cConnection, LocalFileName As String) As Long
    On Error GoTo LOKAL_ERROR

    Dim FF          As Long
    Dim hFile       As Long
    Dim sBuffer     As String
    Dim Lret23      As Long
    Dim bFailed     As Boolean
    Dim cInhalt     As String

    'änderung  jan 05
    If Right(LocalFileName, 9) = "not found" Then
        Exit Function
    End If


    If Connection.Connected = False Then
        RaiseEvent StatusChanged(nsOperationFailed, "Not connected")
        Exit Function
    End If
    RaiseEvent StatusChanged(nsOperationStarted, "")


    hFile = FtpOpenFile(Connection.hConnection, RemoteFile, &H80000000, INTERNET_FLAG_RELOAD, 0)
    If hFile = 0 Then  'Or 13369360
        RaiseEvent StatusChanged(nsOperationFailed, "Unable to open the specified URL.")
        Exit Function
    End If
    sBuffer = Space(BytesBuff)
    FileData = ""

    FF = FreeFile(0)
    Open LocalFileName For Binary Access Write Lock Read Write As #FF
'    cInhalt = Space$(LOF(FF))

    Do

        If InternetReadFile(hFile, sBuffer, BytesBuff, Lret23) = 0 Then

            bFailed = True
            RaiseEvent StatusChanged(nsOperationFailed, "Error while reading the remote file.")
            Exit Do
        End If

        If glFilesizeFTPGlobal = 0 Then
            bFailed = True
            RaiseEvent StatusChanged(nsOperationFailed, "Error while reading the remote file.")
            Exit Do
        End If
        
'        MsgBox Len(cInhalt)

        If Lret23 <> BytesBuff Then
        
            sBuffer = Left$(sBuffer, Lret23)

        End If
        
'        'neu
'        glFilesizeFTPGlobal = glFilesizeFTPGlobal - Lret23

        SentBytes = SentBytes + Lret23
        ActiveFileBytesSent = ActiveFileBytesSent + Lret23

        Put #FF, , sBuffer
        frmWKL38.UpdateProgress

        RaiseEvent ProgressChanged(SentBytes)

    Loop Until Lret23 <> BytesBuff



    InternetCloseHandle hFile
    Close FF
    If bFailed = False Then RaiseEvent StatusChanged(nsOperationSuceeded, "")
    sBuffer = "" ' Clear the buffer

    Exit Function
LOKAL_ERROR:
    MsgBox "cRemoteFile.GetFile1: " & err.Number & " / " & err.Description
    Resume Next
End Function
Function QueryFileSize(Connection As cConnection) As Long
    On Error GoTo LOKAL_ERROR
    
    Dim hHttpFile As Long
    Dim bQuery As Boolean
    Dim szSizeBuffer As String
    Dim dwLengthSizeBuffer As Long
    Dim hCon As Long
    
    If Connection.Connected = False Then
        QueryFileSize = -1
        RaiseEvent StatusChanged(nsOperationFailed, "Not connected.")
        Exit Function
    End If
    RaiseEvent StatusChanged(nsQuerySent, "")
    hHttpFile = InternetOpenUrl(Connection.hInternet, RemoteFile, vbNullString, 0, 0, 0)
    If hHttpFile = 0 Then
        QueryFileSize = -1
        RaiseEvent StatusChanged(nsQueryFailed, "Unable to open the specified URL.")
        Exit Function
    End If
    szSizeBuffer = Space(1024)
    dwLengthSizeBuffer = Len(szSizeBuffer)
    bQuery = HttpQueryInfo(hHttpFile, HTTP_QUERY_CONTENT_LENGTH, szSizeBuffer, dwLengthSizeBuffer, ByVal 0&)
    InternetCloseHandle hHttpFile
    If bQuery = False Then
        QueryFileSize = -1
        RaiseEvent StatusChanged(nsQueryFailed, "Unable to query information.")
    Else
        QueryFileSize = Val(szSizeBuffer)
        RaiseEvent StatusChanged(nsQueryCompleted, "")
    End If
    
    Exit Function
LOKAL_ERROR:
    MsgBox "cRemoteFile.QueryFileSize: " & err.Number & " / " & err.Description
End Function
Function UploadFile(Connection As cConnection, LocalFileName As String) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim cnt As Long, nFileLen As Long, sBuffer() As Byte, nRet As Long, nTotFileLen As Long
    Dim hFile As Long
    Dim nTotBytes As Long
    
    ReDim sBuffer(0 To BytesBuff - 1) As Byte
    If Connection.FTPService = False Then
        RaiseEvent StatusChanged(nsOperationFailed, "Not an FTP session.")
        Exit Function
    End If
    RaiseEvent StatusChanged(nsOperationStarted, "")
    'FtpDeleteFile Connection.hConnection, RemoteFile
    hFile = FtpOpenFile(Connection.hConnection, RemoteFile, &H40000000, UploadFlag, 0)
    If hFile = 0 Then
        RaiseEvent StatusChanged(nsOperationFailed, "Unable to create the remote file.")
        Exit Function
    End If
    Open LocalFileName For Binary Access Read As #1
        nTotFileLen = LOF(1)
        Do
            Get #1, , sBuffer
            If nFileLen < nTotFileLen - BytesBuff Then
                If InternetWriteFile(hFile, sBuffer(0), BytesBuff, nRet) = 0 Then
                    RaiseEvent StatusChanged(nsOperationFailed, "Error while writing the remote file.")
                    Exit Do
                End If
                SentBytes = SentBytes + BytesBuff
                ActiveFileBytesSent = ActiveFileBytesSent + BytesBuff
                nFileLen = nFileLen + BytesBuff
                nTotBytes = nTotBytes + BytesBuff
            Else
                If InternetWriteFile(hFile, sBuffer(0), nTotFileLen - nFileLen, nRet) = 0 Then
                    RaiseEvent StatusChanged(nsOperationFailed, "Error while writing the remote file.")
                    Exit Do
                End If
                SentBytes = SentBytes + (nTotFileLen - nFileLen)
                ActiveFileBytesSent = ActiveFileBytesSent + (nTotFileLen - nFileLen)
                nFileLen = nTotFileLen
                nTotBytes = nTotBytes + nTotFileLen - nFileLen
            End If
            frmWKL38.UpdateProgress
            RaiseEvent ProgressChanged(nFileLen)
        Loop Until nFileLen >= nTotFileLen
    Close
    InternetCloseHandle hFile
    If nFileLen = nTotFileLen Then
        UploadFile = True
        RaiseEvent StatusChanged(nsOperationSuceeded, "")
    End If
    
    Exit Function
LOKAL_ERROR:
    MsgBox "cRemoteFile.UploadFile: " & err.Number & " / " & err.Description
End Function
Public Function DeleteFile(Connection As cConnection)
    On Error GoTo LOKAL_ERROR
    
    If Connection.FTPService = False Or Connection.Connected = False Then
        RaiseEvent StatusChanged(nsOperationFailed, "Not an FTP session.")
        Exit Function
    End If
    RaiseEvent StatusChanged(nsOperationStarted, "")
    If FtpDeleteFile(Connection.hConnection, RemoteFile) = False Then
        RaiseEvent StatusChanged(nsOperationFailed, "Error while deleting the file.")
        Exit Function
    End If
    RaiseEvent StatusChanged(nsOperationSuceeded, "")
    
    Exit Function
LOKAL_ERROR:
    MsgBox "cRemoteFile.DeleteFile: " & err.Number & " / " & err.Description
End Function
Public Function RenameFile(Connection As cConnection, sNewFile As String)
    On Error GoTo LOKAL_ERROR
    
    If Connection.FTPService = False Or Connection.Connected = False Then
        RaiseEvent StatusChanged(nsOperationFailed, "Not an FTP session.")
        Exit Function
    End If
    RaiseEvent StatusChanged(nsOperationStarted, "")
    If FtpRenameFile(Connection.hConnection, RemoteFile, sNewFile) = False Then
        RaiseEvent StatusChanged(nsOperationFailed, "Error while renaming the file.")
        Exit Function
    End If
    RaiseEvent StatusChanged(nsOperationSuceeded, "")
    
    Exit Function
LOKAL_ERROR:
    MsgBox "cRemoteFile.RenameFile: " & err.Number & " / " & err.Description
End Function
