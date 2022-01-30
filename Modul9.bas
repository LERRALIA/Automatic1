Attribute VB_Name = "Modul9"
Option Explicit

Public Const FOP_UPLOAD = &H1
Public Const FOP_DOWNLOAD = &H2
Public Const FTP_TRANSFER_TYPE_ASCII = &H1
Public Const FTP_TRANSFER_TYPE_BINARY = &H2

Public Type tFO
    sName As String
    sPath As String
    bProcedure As Byte
    bCompleted As Boolean
    nFileSize As Long
End Type

Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function GetTickCount& Lib "kernel32" ()

Public bFOBusy As Boolean
Public foFiles() As tFO, foItems As Long, TotalFileSize As Long, SentBytes As Long, OldSpeed As Single
Public ActiveFileBytesSent As Long, ActiveFileBytesTotal As Long, UploadFlag As Long
Public ActiveFile As String, ActiveIndex As Long, ActiveProcedure As Byte, StartT As Long
Public Sub StartFO()
    On Error GoTo LOKAL_ERROR
    
    Dim ret As Long, FF As Integer
    bFOBusy = True
    If foItems <> 0 Then
        
        ret = GetNextFile
        StartT = GetTickCount
        While ret <> -1
            Pause (1)
            OldSpeed = 0
            ActiveIndex = ret
            ActiveFile = foFiles(ret).sName
            ActiveProcedure = foFiles(ret).bProcedure
            ActiveFileBytesTotal = foFiles(ret).nFileSize
            ActiveFileBytesSent = 0
            If foFiles(ret).bProcedure = FOP_UPLOAD Then
                frmWKL38.rfFile.RemoteFile = foFiles(ret).sName
                frmWKL38.rfFile.UploadFile frmWKL38.rfConnection, foFiles(ret).sPath + foFiles(ret).sName
                foFiles(ret).bCompleted = True
            ElseIf foFiles(ret).bProcedure = FOP_DOWNLOAD Then
                frmWKL38.rfFile.RemoteFile = foFiles(ret).sName
                frmWKL38.rfFile.GetFile frmWKL38.rfConnection, foFiles(ret).sPath + foFiles(ret).sName
                foFiles(ret).bCompleted = True
            End If
            GetStatus
            ret = GetNextFile
        Wend
        foItems = 0
        ReDim foFiles(1 To 1) As tFO
        TotalFileSize = 0
        SentBytes = 0
        frmWKL38.FillRemoteListView
        frmWKL38.FillLocalListView frmWKL38.sCurPath
        ActiveFile = ""
        ActiveFileBytesSent = 0
        ActiveFileBytesTotal = 1
        ActiveProcedure = 0
        TotalFileSize = 1
        SentBytes = 0
        frmWKL38.UpdateProgress
        NotifyWhenComplete
    End If
    bFOBusy = False
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul9"
    Fehler.gsFunktion = "StartFO"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Public Function GetNextFile() As Long
    On Error GoTo LOKAL_ERROR
    
    Dim cnt As Long
    GetNextFile = -1
    For cnt = 1 To foItems
        If foFiles(cnt).bCompleted = False Then
            GetNextFile = cnt
            Exit For
        End If
    Next cnt
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul9"
    Fehler.gsFunktion = "GetNextFile"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Public Sub StartFOEinzelDatfromLoToRe(sPfadUndDatname As String)
    On Error GoTo LOKAL_ERROR
    
    bFOBusy = True
    
    If foItems <> 0 Then
        
        
        StartT = GetTickCount
        OldSpeed = 0
        ActiveIndex = 1
        ActiveFile = foFiles(1).sName
        ActiveProcedure = foFiles(1).bProcedure
        ActiveFileBytesTotal = foFiles(1).nFileSize
        ActiveFileBytesSent = 0
        
        foFiles(1).bProcedure = FOP_UPLOAD
        frmWKL38.rfFile.RemoteFile = foFiles(1).sName
        frmWKL38.rfFile.UploadFile frmWKL38.rfConnection, sPfadUndDatname
        foFiles(1).bCompleted = True
            
        foItems = 0
        ReDim foFiles(1 To 1) As tFO
        TotalFileSize = 0
        SentBytes = 0
        
        ActiveFile = ""
        ActiveFileBytesSent = 0
        ActiveFileBytesTotal = 1
        ActiveProcedure = 0
        TotalFileSize = 1
        SentBytes = 0
        frmWKL38.FillRemoteListView
        frmWKL38.UpdateProgress
        NotifyWhenComplete
    End If
    bFOBusy = False
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul9"
    Fehler.gsFunktion = "StartFOEinzelDatfromLoToRe"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Public Function NotifyWhenComplete()
    On Error GoTo LOKAL_ERROR
    
    frmWKL38.WindowState = vbNormal
    SetForegroundWindow frmWKL38.hwnd
'    Beep
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul9"
    Fehler.gsFunktion = "NotifyWhenComplete"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Public Sub AddToCollection(bProcedure As Byte, sFile As String, sPath As String, nFileSize As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim cnt As Long, bOk As Boolean
    For cnt = 1 To foItems
        If foFiles(cnt).sName = sFile Then
            bOk = True
            Exit For
        End If
    Next cnt
    If bOk = False Then
        foItems = foItems + 1
        ReDim Preserve foFiles(1 To foItems) As tFO
        foFiles(foItems).bProcedure = bProcedure
        foFiles(foItems).nFileSize = nFileSize
        foFiles(foItems).sName = sFile
        foFiles(foItems).sPath = sPath
        TotalFileSize = TotalFileSize + nFileSize
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul9"
    Fehler.gsFunktion = "AddToCollection"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Public Sub GetStatus()
    On Error GoTo LOKAL_ERROR
    
    frmWKL38.txtStatus.Text = frmWKL38.txtStatus.Text + frmWKL38.rfConnection.GetLastResponseInfo
    frmWKL38.txtStatus.SelStart = Len(frmWKL38.txtStatus.Text)
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul9"
    Fehler.gsFunktion = "GetStatus"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
