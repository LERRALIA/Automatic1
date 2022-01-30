Attribute VB_Name = "Modul13"
Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (ByRef Destination As Long, ByVal Source As Long, ByVal Length As Long)

Rem Beware: IFSF2.dll depends on libexpat.dll and SockKommu.dll
Rem         VB must be able to find them in the standard search
Rem         path.
Public Declare Sub IFSFINITIALIZELIBRARY Lib "IFSF2.dll" ()

Public Declare Sub IFSFCONSTRUCTREQUESTOBJECT Lib "IFSF2.dll" (ByRef pIfsf As Long, ByVal sServerHost As String, ByVal iServerPort As Long, ByVal sWorkstationID As String, ByVal sPopID As String, ByVal sApplicationID As String, ByVal iTimeout As Long)
Public Declare Sub IFSFDESTRUCTREQUESTOBJECT Lib "IFSF2.dll" (ByRef pIfsf As Long)
Public Declare Function IFSFPAYMENT Lib "IFSF2.dll" (ByVal pIfsf As Long, ByVal pAmout As String, ByVal pPaymentType As Integer, ByVal pCardType As String, ByVal iCardTypeLength As Long) As Integer

Public Declare Sub IFSFCONSTRUCTDEVICEOBJECT Lib "IFSF2.dll" (ByRef pDevice As Long, ByVal iPort As Long)
Public Declare Sub IFSFDESTRUCTDEVICEOBJECT Lib "IFSF2.dll" (ByRef pDevice As Long)

Public Declare Sub IFSFSTARTPRINTERLISTENER Lib "IFSF2.dll" (ByVal pDevice As Long, ByVal pCallback As Long)
Public Declare Sub IFSFSTARTDISPLAYLISTENER Lib "IFSF2.dll" (ByVal pDevice As Long, ByVal pCallback As Long)

Public Declare Sub IFSFSTOPPRINTERLISTENER Lib "IFSF2.dll" (ByVal pDevice As Long)
Public Declare Sub IFSFSTOPDISPLAYLISTENER Lib "IFSF2.dll" (ByVal pDevice As Long)

Public Declare Sub IFSFTICKETREPRINT Lib "IFSF2.dll" (ByVal pDevice As Long)

Private Declare Function ELMEGetPrint Lib "wk2elme.dll" (ByVal szDataBuffer As String, ByVal iBufferLen As Long) As Long

Dim pIfsf As Long
Dim pDevice As Long
Dim iCallbacks As Long

Public ifsfPrinter As TextBox
Public ifsfDisplay As Label
Public ifsfAmount As TextBox

Public Function charP2String(ByVal pLine As Long) As String
  Dim char As Long
  Dim line As String
  Call CopyMemory(char, pLine, 1)
  While char <> 0
    line = line & Chr(char)
    pLine = pLine + 1
    Call CopyMemory(char, pLine, 1)
  Wend
  charP2String = line
End Function

Public Function charPP2StringArray(ByVal pLines As Long, lines() As String) As Boolean
On Error GoTo LOKAL_ERROR

  Dim pLine As Long
  Dim i As Long
  charPP2StringArray = False
  If pLines <> 0 Then
    Call CopyMemory(pLine, pLines, 4)
    While pLine <> 0
      charPP2StringArray = True
      ReDim Preserve lines(i)
      lines(i) = charP2String(pLine)
      i = i + 1
      pLines = pLines + 4
      Call CopyMemory(pLine, pLines, 4)
    Wend
  End If
  
Exit Function
LOKAL_ERROR:
    MsgBox "charPP2StringArray" & " " & err.Number & " " & err.Description
End Function
Public Function printerCallback(ByVal pLines As Long) As Integer
On Error GoTo LOKAL_ERROR
    
    Dim i As Integer
    Dim lines() As String
    Read = charPP2StringArray(pLines, lines)
    If Read <> 0 Then
        For i = LBound(lines) To UBound(lines)
            ifsfPrinter.Text = ifsfPrinter.Text & vbCrLf & lines(i)
        Next i
    End If
  
  printerCallback = 1
  
Exit Function
LOKAL_ERROR:
    MsgBox "printerCallback" & " " & err.Number & " " & err.Description
End Function

Public Function displayCallback(ByVal pLine As Long) As Integer
  Dim line As String
  line = charP2String(pLine)
  
  ifsfDisplay.Caption = line
  
  displayCallback = 1
End Function
Public Function belegdrucken() As String

'    Call IFSFINITIALIZELIBRARY
'    Call IFSFCONSTRUCTREQUESTOBJECT(pIfsf, "192.168.1.11", 20002, "001", "61585526", "VB Reference", 1000000)
'    Call IFSFCONSTRUCTDEVICEOBJECT(pDevice, 20007, pIfsf)
    
    Call IFSFSTARTPRINTERLISTENER(pDevice, AddressOf printerCallback)
    Call IFSFTICKETREPRINT(pDevice)
    Call IFSFSTOPPRINTERLISTENER(pDevice)
    
'    Call IFSFDESTRUCTDEVICEOBJECT(pDevice)
'    Call IFSFDESTRUCTREQUESTOBJECT(pIfsf)
    

End Function
Public Sub initZahlung()


    Dim lret1 As Long
    
    
    
  Call IFSFINITIALIZELIBRARY
  Call IFSFCONSTRUCTREQUESTOBJECT(pIfsf, "192.168.1.11", 20002, "001", "61585526", "VB Reference", 1000000)
  Call IFSFCONSTRUCTDEVICEOBJECT(pDevice, 20007)
  
'  Call IFSFSTARTPRINTERLISTENER(pDevice, AddressOf printerCallback)
  Call IFSFSTARTDISPLAYLISTENER(pDevice, AddressOf displayCallback)
  
  Dim cardType As String * 255
  iRet = IFSFPAYMENT(pIfsf, ifsfAmount.Text, 1, cardType, Len(cardType))
  MsgBox iRet
  
  
  Dim sBLZ As String * 8000
  
  lret1 = ELMEGetPrint(sBLZ, 8000)
    If lret1 = 0 Then
        MsgBox sBLZ
'        SendeDaten2DruckerECCASH
    Else
        MsgBox "Fehler ELMEGetPrint: " & lret1, vbCritical, "Winkiss Fehler:"
'        gsAdtBeleg = ""
    End If
  
  
  

  Call IFSFSTOPDISPLAYLISTENER(pDevice)
'  Call IFSFSTOPPRINTERLISTENER(pDevice)

'    belegdrucken

  Call IFSFDESTRUCTDEVICEOBJECT(pDevice)
  Call IFSFDESTRUCTREQUESTOBJECT(pIfsf)
  
  
  
'  belegdrucken
End Sub
