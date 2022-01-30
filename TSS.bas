Attribute VB_Name = "TSS"
Public Const NotConnectedText As String = "TSE Fehler"
Public Const OfflineText As String = "TSE Offline"
Public Const OkText As String = "TSE Online"

Private Const mask As String = "0.00"

Public TRStart As String
Public TRFinish As String
Public TRNo As String
Public Serial As String

Private isStarted As Boolean

Public Enum ReceiptType
    Beleg = 0
    AVTraining = 1
    AVTransfer = 2
    AVBestellung = 3
    AVBelegabbruch = 4
    AVSachbezug = 5
    AVRechnung = 6
    AVSonstige = 7
    AVBelegstorno = 8
End Enum

Public Sub Connect(ByVal socket As Winsock, ByVal tim As Timer)
    If gbTSE_SCHREIBEN Then
        If socket.State <> sckClosed Then
            socket.Close
        End If
        socket.Connect
        tim.Enabled = True
    End If
End Sub

Public Sub Disconnect(ByVal socket As Winsock, ByVal tim As Timer)
    If gbTSE_SCHREIBEN Then
        socket.Close
        tim.Enabled = False
    End If
End Sub

Public Sub Start(ByVal socket As Winsock)
    If gbTSE_SCHREIBEN Then
        Send socket, SegmentTSSStart, ""
        isStarted = True
    End If
End Sub

Public Sub Cancel(ByVal socket As Winsock)
    If gbTSE_SCHREIBEN And isStarted Then
        Send socket, SegmentTSSCancel, ""
        isStarted = False
    End If
End Sub

Public Function Finish(ByVal socket As Winsock, ByVal rt As ReceiptType, ByVal regular As Double, ByVal reduced As Double, ByVal none As Double, ByVal cur As String, ByVal cash As Double, ByVal noncash As Double)
    If gbTSE_SCHREIBEN And isStarted Then
        Dim data As String
        data = CStr(rt) + ";" + Format(regular, mask) + ";" + Format(reduced, mask) + ";" + Format(none, mask) + ";" + cur + ";" + Format(cash, mask) + ";" + Format(noncash, mask)
        TRStart = "16.07.2019 16:10:00"
        TRFinish = "16.07.2019 16:10:20"
        TRNo = "215"
        Serial = "9ac89ff8-3cc4-46b6-ab7c-57a8c4ca5400"
        Send socket, SegmentTSSFinish, data
        isStarted = False
    End If
End Function

Public Sub TimerTick(ByVal socket As Winsock, ByVal tim As Timer)
    If gbTSE_SCHREIBEN And socket.State = sckConnected Then
        Send socket, SegmentHeartbeat, ""
    End If
End Sub

Public Sub Error(ByVal socket As Winsock, ByVal tim As Timer)
    If gbTSE_SCHREIBEN Then
        MsgBox "Verbindung zum TSE Server nicht möglich, bitte stellen Sie sicher, dass der Server ordnungsgemäß läuft."
        tim.Enabled = False
        socket.Close
    End If
End Sub

Public Sub Receive(ByVal socket As Winsock)
    If gbTSE_SCHREIBEN Then
        Dim data() As Byte
        socket.GetData data
        Dim msg As Message: Set msg = New Message
        Dim sgm As Segment: Set sgm = New Segment
        Set msg = NetworkDataConverter.DecodeMessage(socket, -1, data)
        If Not msg Is Nothing Then
            Set sgm = msg.Item(0)
        End If
    End If
End Sub

Private Sub Send(ByVal socket As Winsock, ByVal category As SegmentCategory, ByVal para As String)
    If gbTSE_SCHREIBEN And socket.State = sckConnected Then
        Dim msg As Message: Set msg = New Message
        Dim sgm As Segment: Set sgm = New Segment
        sgm.id = category
        sgm.sequence = 0
        sgm.Session = -1
        sgm.version = sgm.CurrentVersion
        sgm.Parameter = para
        msg.Item(0) = sgm
        Dim data() As Byte: data = NetworkDataConverter.EncodeMessage(msg)
        Call socket.senddata(data)
    End If
End Sub
