Attribute VB_Name = "Modul12"
Option Explicit

Declare Function ELMEPay Lib "wk2elme.dll" (ByVal lAmount As Long, ByVal lUsePin As Long) As Long
Declare Function ELMEGetBankSortingCode Lib "wk2elme.dll" (ByVal szDataBuffer As String, ByVal iBufferLen As Long) As Long
Declare Function ELMEGetPrint Lib "wk2elme.dll" (ByVal szDataBuffer As String, ByVal iBufferLen As Long) As Long
Declare Function ELMEGetLastError Lib "wk2elme.dll" (ByRef piErrCode As Long, ByVal szErrBuffer As String, ByVal iBufferLen As Long) As Long

Declare Function ELMESettings Lib "wk2elme.dll" _
    (ByVal szWorkstationID As String, ByVal szPopID As String, _
    ByVal szApplID As String, ByVal szHostAddr As String, _
    ByVal iServerPort As Long, ByVal iTimeout As Long, ByVal iDevicePort As Long, _
    ByVal iColumns As Long, ByVal szLogPath As String) As Long

Declare Function ELMEReversal Lib "wk2elme.dll" (ByVal iSTAN As Long) As Long
Declare Function ELMEGetTerminalID Lib "wk2elme.dll" (ByVal szDataBuffer As String, ByVal iBufferLen As Long) As Long


Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, ByVal Source As Any, ByVal Length As Long)

Rem Beware: IFSF2.dll depends on libexpat.dll and SockKommu.dll
Rem         VB must be able to find them in the standard search
Rem         path.
Public Declare Sub IFSFINITIALIZELIBRARY Lib "c:/SEC/IFSF2.dll" ()

Public Declare Sub IFSFCONSTRUCTREQUESTOBJECT Lib "c:/SEC/IFSF2.dll" (pIfsf As Long, ByVal sServerHost As String, ByVal iServerPort As Long, ByVal sWorkstationID As String, ByVal sPopID As String, ByVal sApplicationID As String, ByVal iTimeout As Long)
Public Declare Sub IFSFDESTRUCTREQUESTOBJECT Lib "c:/SEC/IFSF2.dll" (pIfsf As Long)

Public Declare Function IFSFPAYMENT Lib "c:/SEC/IFSF2.dll" (ByVal pIfsf As Long, ByVal pAmout As String, ByVal pPaymentType As Integer, ByVal pCardType As String, ByVal iCardTypeLength As Long) As Integer
Public Declare Function IFSFREVERSAL Lib "c:/SEC/IFSF2.dll" (ByVal pIfsf As Long, ByVal pTraceNr As String, ByVal pCardType As String, ByVal iCardTypeLength As Long) As Integer


Public Declare Sub IFSFCONSTRUCTDEVICEOBJECT Lib "c:/SEC/IFSF2.dll" (pDevice As Long, ByVal iPort As Long, ByVal pIfsf As Long)
Public Declare Sub IFSFDESTRUCTDEVICEOBJECT Lib "c:/SEC/IFSF2.dll" (pDevice As Long)
Public Declare Sub IFSFSTARTPRINTERLISTENER Lib "c:/SEC/IFSF2.dll" (ByVal pDevice As Long, ByVal pCallback As Long)
Public Declare Sub IFSFSTARTDISPLAYLISTENER Lib "c:/SEC/IFSF2.dll" (ByVal pDevice As Long, ByVal pCallback As Long)
Public Declare Sub IFSFSTOPPRINTERLISTENER Lib "c:/SEC/IFSF2.dll" (ByVal pDevice As Long)
Public Declare Sub IFSFSTOPDISPLAYLISTENER Lib "c:/SEC/IFSF2.dll" (ByVal pDevice As Long)


Dim pIfsf As Long
Dim pDevice As Long
Dim iCallbacks As Long
Public ifsfPrinter As TextBox
Public ifsfDisplay As Label
Public ifsfAmount As TextBox

Dim sBon As String
Dim sDisplay As String
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
End Function
Public Function printerCallback(ByVal pLines As Long) As Integer
  Dim lines() As String
  Dim read As Boolean
  Dim i As Integer
  
'  Dim ifilenr         As Integer
'  Dim lPos As Long
'
'ifilenr = FreeFile
'Open "C:\SEC\Bon.txt" For Binary As #ifilenr
'
'lPos = LOF(ifilenr)

  read = charPP2StringArray(pLines, lines)
  If read <> 0 Then
    For i = LBound(lines) To UBound(lines)
    
'        lPos = lPos + 1
'        Put #ifilenr, lPos, sBon
        sBon = sBon & vbCrLf & lines(i)
        
        Next i
  End If
  
'  Close ifilenr
  
  printerCallback = 1
End Function

Public Function displayCallback(ByVal pLine As Long) As Integer
  Dim line As String
  line = charP2String(pLine)
  
'  ifsfDisplay.Caption = line
    sDisplay = line

  
  displayCallback = 1
End Function
Public Function initSECPOSII(sZahlbetrag As String, sIPAdress As String, sClient As String, sPort As String, sTermiID As String) As Integer


  Call IFSFINITIALIZELIBRARY
  Call IFSFCONSTRUCTREQUESTOBJECT(pIfsf, sIPAdress, sPort, sClient, sTermiID, "VB Reference", 100000)
  Call IFSFCONSTRUCTDEVICEOBJECT(pDevice, 20007, pIfsf)
  
  Call IFSFSTARTPRINTERLISTENER(pDevice, AddressOf printerCallback)
  Call IFSFSTARTDISPLAYLISTENER(pDevice, AddressOf displayCallback)
  
  Dim cardType As String * 255
  initSECPOSII = IFSFPAYMENT(pIfsf, sZahlbetrag, 1, cardType, Len(cardType))
  
  If initSECPOSII = 0 Then
'        MsgBox "zahlung erfolgt"
  Else
'        MsgBox "zahlung nicht erfolgt"
  End If

  Call IFSFSTOPDISPLAYLISTENER(pDevice)
  Call IFSFSTOPPRINTERLISTENER(pDevice)
  
  Call IFSFDESTRUCTDEVICEOBJECT(pDevice)
  Call IFSFDESTRUCTREQUESTOBJECT(pIfsf)
  
    
  gsAdtBeleg = sBon
  
End Function

Public Function initSECPOSII_Storno(strace As String, sIPAdress As String, sClient As String, sPort As String, sTermiID As String) As Integer


  Call IFSFINITIALIZELIBRARY
  Call IFSFCONSTRUCTREQUESTOBJECT(pIfsf, sIPAdress, sPort, sClient, sTermiID, "VB Reference", 100000)
  Call IFSFCONSTRUCTDEVICEOBJECT(pDevice, 20007, pIfsf)
  
  Call IFSFSTARTPRINTERLISTENER(pDevice, AddressOf printerCallback)
  Call IFSFSTARTDISPLAYLISTENER(pDevice, AddressOf displayCallback)
  
  Dim cardType As String * 255
  initSECPOSII_Storno = IFSFREVERSAL(pIfsf, strace, cardType, Len(cardType))
  
  If initSECPOSII_Storno = 0 Then
'        MsgBox "zahlung erfolgt"
  Else
'        MsgBox "zahlung nicht erfolgt"
  End If

  Call IFSFSTOPDISPLAYLISTENER(pDevice)
  Call IFSFSTOPPRINTERLISTENER(pDevice)
  
  Call IFSFDESTRUCTDEVICEOBJECT(pDevice)
  Call IFSFDESTRUCTREQUESTOBJECT(pIfsf)
  
    
  gsAdtBeleg = sBon
  
End Function

