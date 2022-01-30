Attribute VB_Name = "mdlEasyTSE"
Option Explicit

 'ist TSE aktiviert
 Global E_TSE_Aktiv As Boolean
 
 
 'durch dieses Objekt kommunizieren wir mit dem EPSON USB-Stick
 Global goTSE As Object
  
 
 'Zugangdaten zur Verbindung mit USB-Stick
 Global E_PUK As String            'max 6-stellig
 Global E_AdminPin As String       'max 5-stellig
 Global E_TimeAdminPin As String   'max 5-stellig
 Global E_SecretKey As String      'max 5-stellig
 Global E_IP As String
 Global E_Port As Integer
 Global E_DeviceID As String
 Global E_ClientID As String
 Global E_UserID As String
 
 'standardmäßig = 1 (single)
 Global E_SingleOrServerTSE As Integer
 
 'StorageInfo von USB-Stick (die werden zugewiesen,sobald Verbindung mit USB erfolgreich erstellt wird)
 Global USB_cdcId As String
 Global USB_cdcHash  As String
 Global USB_certificateExpirationDate As String
 Global USB_createdSignatures As String
 Global USB_hardwareVersion As String
 Global USB_hasPassedSelfTest As String
 Global USB_hasValidTime  As String
 Global USB_isExportEnabledIfCspTestFails As String
 Global USB_isTSEUnlocked As String
 Global USB_lastExportExecutedDate As String
 Global USB_maxRegisteredClients As String
 Global USB_maxSignatures As String
 Global USB_maxStartedTransactions As String
 Global USB_maxUpdateDelay As String
 Global USB_registeredClients As String
 Global USB_remainingSignatures As String
 Global USB_serialNumber As String
 Global USB_signatureAlgorithm As String
 Global USB_softwareVersion As String
 Global USB_startedTransactions As String
 Global USB_tarExportSize As String
 Global USB_timeUntilNextSelfTest As String
 Global USB_tseCapacity As String
 Global USB_tseCurrentSize As String
 Global USB_tseDescription As String
 Global USB_tseInitializationState As String
 Global USB_tsePublicKey As String
 Global USB_vendorType As String
 
 
 'TSE-Rückgaben Variabeln
 Global R_StartTime As String
 Global R_FinishTime As String
 Global R_TransactionNr As String
 Global R_QRCodeAlsText As String
 Global R_QRCodeAlsImgPath As String
 Global R_StartSignatur As String
 Global R_FinishSignatur As String
 Global R_FINISH_SIG_Zaehler As Double
 Global R_START_SIG_Zaehler As Double
 
 'wenn TSE vernünftig initialisiert ist, ist TSE_OK = True , TSE_Err = ""
 Global TSE_OK As Boolean
 Global TSE_Err As String
 
 'TSE ExportPfad
 Global gbTSEExportPfad As String
 
 'Soll QR-Code gedruckt werden
 Global MitQrCode As Boolean
 
 'true:   wenn Bondrucker alt ist
 'Falsch: wenn Bondrucker modern ist
 Global altDruckModus As Boolean
 
 Global TSE_ID As Integer
 
 Global FalschUSB_serialNumber As String
 
  
 'True  = (bei Stack_UpdateTime wird die Zeit vom Internet abgefragt)
 'False = (bei Stack_UpdateTime wird die Zeit vom Lokal_PC eingelesen)
 Global TSE_InternetZeitAbfragen As Boolean
 
 
 
 
 
 
 
 

Public Sub TSE_Initialisieren()
 On Error GoTo LOKAL_ERROR
      
     If IstTseInstalliert Then
     
        TSE_OK = True
        TSE_Err = ""
        
          
        TseEinstellungen.lblZeig.ForeColor = vbBlack
        TseEinstellungen.lblZeig.Caption = "ein Instanz von EasyTSE.EpsonTSE wird erstellt . . ."
        TseEinstellungen.lblZeig.Refresh
        
        frmWKL00.lbl_TSE.ForeColor = vbBlack
        frmWKL00.lbl_TSE.Caption = "ein Instanz von EasyTSE.EpsonTSE wird erstellt . . ."
        frmWKL00.lbl_TSE.Refresh
        
        TseEinstellungen.btnNeueClient.Enabled = False
        'TseEinstellungen.btnTSEInfo.Enabled = False
        TseEinstellungen.btnSpeicher.Enabled = False
        frmWKL00.lbl_TSE.Visible = True
        
     
        Set goTSE = CreateObject("EasyTSE.EpsonTSE")
        
        If Not goTSE Is Nothing Then
            LeseUSBStickZugangdaten
         Else
            TseEinstellungen.lblZeig.ForeColor = vbRed
            TseEinstellungen.lblZeig.Caption = "kein Instanz erstellt . . ."
            TseEinstellungen.lblZeig.Refresh
            
            frmWKL00.lbl_TSE.ForeColor = vbRed
            frmWKL00.lbl_TSE.Caption = "kein Instanz erstellt . . ."
            frmWKL00.lbl_TSE.Refresh
            
        End If
    
    Else
    
     TSE_OK = False
     TSE_Err = "EasyTSE ist auf diesem Rechner nicht installiert !!!"
     
     TseEinstellungen.lblZeig.ForeColor = vbRed
     TseEinstellungen.lblZeig.Caption = TSE_Err
     TseEinstellungen.lblZeig.Refresh
     
     frmWKL00.lbl_TSE.ForeColor = vbRed
     frmWKL00.lbl_TSE.Caption = TSE_Err
     frmWKL00.lbl_TSE.Refresh
     
     Exit Sub
    End If
    
 Exit Sub
    
LOKAL_ERROR:

    Screen.MousePointer = 0

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdlEasyTSE"
    Fehler.gsFunktion = "TSE_Initialisieren"
    Fehler.gsFehlertext = "Im Programmteil mdlEasyTSE ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub

 

Sub LeseUSBStickZugangdaten()
    On Error GoTo LOKAL_ERROR
     
     TSE_OK = True
     TSE_Err = ""
     
     TseEinstellungen.lblZeig.Caption = "TSESettings werden abgefragt . . ."
     TseEinstellungen.lblZeig.Refresh
     
     frmWKL00.lbl_TSE.Caption = "TSESettings werden abgefragt . . ."
     frmWKL00.lbl_TSE.Refresh
      
      
     Dim rsrs As Recordset
     
        Set rsrs = gdApp.OpenRecordset("select * from TSESettings")
        If Not rsrs.EOF Then

            If Not IsNull(rsrs!TSE_PUK) Then
                E_PUK = rsrs!TSE_PUK
            End If

            If Not IsNull(rsrs!TSE_AdminPin) Then
                E_AdminPin = rsrs!TSE_AdminPin
                TseEinstellungen.txtTSE_Adminpin.Text = E_AdminPin
            End If
 
            If Not IsNull(rsrs!TSE_TimeAdminPin) Then
                E_TimeAdminPin = rsrs!TSE_TimeAdminPin
                TseEinstellungen.txtTSE_TimeAdminpin.Text = E_TimeAdminPin
            End If
 
 
            If Not IsNull(rsrs!TSE_ClientID) Then
                E_ClientID = rsrs!TSE_ClientID
                TseEinstellungen.comClients.Text = E_ClientID
            End If
 
            If Not IsNull(rsrs!TSE_IP) Then
                E_IP = rsrs!TSE_IP
                TseEinstellungen.txtTSE_IP.Text = E_IP
            End If
 
            If Not IsNull(rsrs!TSE_Port) Then
                E_Port = rsrs!TSE_Port
                TseEinstellungen.txtTSE_Port.Text = E_Port
            End If
            
            If Not IsNull(rsrs!TSE_DeviceID) Then
                E_DeviceID = rsrs!TSE_DeviceID
            End If
 

             'statische Variabeln
             E_SingleOrServerTSE = 1
             E_SecretKey = goTSE.TSE_SecretKey
             E_UserID = goTSE.TSE_UserID
              
        End If
        rsrs.Close: Set rsrs = Nothing
 
       TSE_Verbinden
       
    Exit Sub
LOKAL_ERROR:

    Screen.MousePointer = 0
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdlEasyTSE"
    Fehler.gsFunktion = "LeseUSBStickZugangdaten"
    Fehler.gsFehlertext = "Im Programmteil mdlEasyTSE ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub


Sub TSE_Verbinden()
 On Error GoTo LOKAL_ERROR
     
     TSE_OK = True
     TSE_Err = ""
     
     TseEinstellungen.lblZeig.Caption = "Verbindung mit TSE wird hergestellt . . ."
     TseEinstellungen.lblZeig.Refresh
     
     frmWKL00.lbl_TSE.Caption = "Verbindung mit TSE wird hergestellt . . ."
     frmWKL00.lbl_TSE.Refresh
     
     goTSE.TSE_PUK = E_PUK
     goTSE.TSE_AdminPin = E_AdminPin
     goTSE.TSE_TimeAdminPin = E_TimeAdminPin
 
     goTSE.TSE_IP = E_IP
     goTSE.TSE_Port = E_Port

     goTSE.TSE_SecretKey = E_SecretKey
     goTSE.TSE_DeviceID = E_DeviceID
     goTSE.TSE_UserID = E_UserID
     
     'goTSE.nSingleOrServerTSE = E_SingleOrServerTSE
     
     'nAutoSetup muss auf 1 gesetzt werden
     goTSE.nAutoSetup = 1
     
     If goTSE.TSEConnectOpenSend("GetStorageInfo") = 1 Then
      
        
      
        USB_cdcId = goTSE.oGetStorageInfo.Output.TSEInformation.cdcId
        USB_cdcHash = goTSE.oGetStorageInfo.Output.TSEInformation.cdcHash
        USB_certificateExpirationDate = goTSE.oGetStorageInfo.Output.TSEInformation.certificateExpirationDate
        USB_createdSignatures = goTSE.oGetStorageInfo.Output.TSEInformation.createdSignatures
        USB_hardwareVersion = goTSE.oGetStorageInfo.Output.TSEInformation.hardwareVersion
        USB_hasPassedSelfTest = goTSE.oGetStorageInfo.Output.TSEInformation.hasPassedSelfTest
        USB_hasValidTime = goTSE.oGetStorageInfo.Output.TSEInformation.hasValidTime
        
        USB_isExportEnabledIfCspTestFails = goTSE.oGetStorageInfo.Output.TSEInformation.isExportEnabledIfCspTestFails
        USB_isTSEUnlocked = goTSE.oGetStorageInfo.Output.TSEInformation.isTSEUnlocked
        USB_lastExportExecutedDate = goTSE.oGetStorageInfo.Output.TSEInformation.lastExportExecutedDate
        USB_maxRegisteredClients = goTSE.oGetStorageInfo.Output.TSEInformation.maxRegisteredClients
        USB_maxSignatures = goTSE.oGetStorageInfo.Output.TSEInformation.maxSignatures
        USB_maxStartedTransactions = goTSE.oGetStorageInfo.Output.TSEInformation.maxStartedTransactions
        USB_maxUpdateDelay = goTSE.oGetStorageInfo.Output.TSEInformation.maxUpdateDelay
        
        USB_registeredClients = goTSE.oGetStorageInfo.Output.TSEInformation.registeredClients
        USB_remainingSignatures = goTSE.oGetStorageInfo.Output.TSEInformation.remainingSignatures
        
        'Serial Nummer lesen <<<<<<<<< START
        
            'das ist die falsche SerialNummer, und muss in den Tabellen [ KASSBON , GetStorageInfo ] gegen die richtige Nummer getauscht werden
            FalschUSB_serialNumber = goTSE.oGetStorageInfo.Output.TSEInformation.SerialNumber
            
            'get die richtige Serial Nummmer
            goTSE.BuildSNfromPublicKey
            USB_serialNumber = goTSE.cSNfromPublicKey
            
            'das Korrigieren (Tauschen) der Serial-Nummer findest du hier in diesem Modul unter dem Titel 'Serial Nummer erstmal korrigieren'
            
            
        'Serial Nummer lesen <<<<<<<<< ENDE
        
        
        USB_signatureAlgorithm = goTSE.oGetStorageInfo.Output.TSEInformation.signatureAlgorithm
        USB_softwareVersion = goTSE.oGetStorageInfo.Output.TSEInformation.softwareVersion
        USB_startedTransactions = goTSE.oGetStorageInfo.Output.TSEInformation.startedTransactions
        USB_tarExportSize = goTSE.oGetStorageInfo.Output.TSEInformation.tarExportSize
        
        USB_timeUntilNextSelfTest = goTSE.oGetStorageInfo.Output.TSEInformation.timeUntilNextSelfTest
        USB_tseCapacity = goTSE.oGetStorageInfo.Output.TSEInformation.tseCapacity
        USB_tseCurrentSize = goTSE.oGetStorageInfo.Output.TSEInformation.tseCurrentSize
        USB_tseDescription = goTSE.oGetStorageInfo.Output.TSEInformation.tseDescription
        USB_tseInitializationState = goTSE.oGetStorageInfo.Output.TSEInformation.tseInitializationState
        USB_tsePublicKey = goTSE.oGetStorageInfo.Output.TSEInformation.tsePublicKey
        USB_vendorType = goTSE.oGetStorageInfo.Output.TSEInformation.vendorType
      
        TseEinstellungen.txtTSE_SN.Text = USB_serialNumber
        
        SetupForPrinter
        
    Else
         TSE_OK = False
         TSE_Err = goTSE.cErrorList
         
         TseEinstellungen.lblZeig.ForeColor = vbRed
         TseEinstellungen.lblZeig.Caption = TSE_Err
         TseEinstellungen.lblZeig.Refresh
         
         frmWKL00.lbl_TSE.ForeColor = vbRed
         frmWKL00.lbl_TSE.Caption = TSE_Err
         frmWKL00.lbl_TSE.Refresh
         
     End If
     
     'immer prüfen, ob es einen TSE-Fehler gab  <<<<< START
      If Not Trim(goTSE.cErrorList) = "" Then
             
             MsgBox ("TSE Fehler:" & vbNewLine & goTSE.cErrorList)
             
      End If
     'immer prüfen, ob es einen TSE-Fehler gab  <<<<< ENDE
     
     
 Exit Sub
    
LOKAL_ERROR:
    Screen.MousePointer = 0
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdlEasyTSE"
    Fehler.gsFunktion = "TSE_Verbinden"
    Fehler.gsFehlertext = "Im Programmteil mdlEasyTSE ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

 Sub SetupForPrinter()
 On Error GoTo LOKAL_ERROR
 
   TSE_OK = True
   TSE_Err = ""
   
   TseEinstellungen.lblZeig.Caption = "TSE: Setup For Printer . . ."
   TseEinstellungen.lblZeig.Refresh
   
   frmWKL00.lbl_TSE.Caption = "TSE: Setup For Printer . . ."
   frmWKL00.lbl_TSE.Refresh
   
   
 
     If goTSE.Stack_SetupForPrinter() = 1 Then
          USBSelfTest
     Else
          TSE_OK = False
          TSE_Err = "TSE: " & goTSE.cErrorList
          
          TseEinstellungen.lblZeig.ForeColor = vbRed
          TseEinstellungen.lblZeig.Caption = TSE_Err
          TseEinstellungen.lblZeig.Refresh
          
          frmWKL00.lbl_TSE.ForeColor = vbRed
          frmWKL00.lbl_TSE.Caption = TSE_Err
          frmWKL00.lbl_TSE.Refresh
          
          
     End If
     
     'immer prüfen, ob es einen TSE-Fehler gab  <<<<< START
      If Not Trim(goTSE.cErrorList) = "" Then
             
             MsgBox ("TSE Fehler:" & vbNewLine & goTSE.cErrorList)
             
      End If
     'immer prüfen, ob es einen TSE-Fehler gab  <<<<< ENDE
     
 Exit Sub
LOKAL_ERROR:
    Screen.MousePointer = 0
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdlEasyTSE"
    Fehler.gsFunktion = "SetupForPrinter"
    Fehler.gsFehlertext = "Im Programmteil mdlEasyTSE ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
 
 
Sub USBSelfTest()
 On Error GoTo LOKAL_ERROR
 
   TSE_OK = True
   TSE_Err = ""
   
   TseEinstellungen.lblZeig.Caption = "TSE: USB Self-Test . . ."
   TseEinstellungen.lblZeig.Refresh
   
   frmWKL00.lbl_TSE.Caption = "TSE: USB Self-Test . . ."
   frmWKL00.lbl_TSE.Refresh
   
   
 
     If goTSE.Stack_RunTSESelfTest() = 1 Then
          UpdateTime
        Else
          TSE_OK = False
          TSE_Err = "TSE: " & goTSE.cErrorList
          
          TseEinstellungen.lblZeig.ForeColor = vbRed
          TseEinstellungen.lblZeig.Caption = TSE_Err
          TseEinstellungen.lblZeig.Refresh
          
          frmWKL00.lbl_TSE.ForeColor = vbRed
          frmWKL00.lbl_TSE.Caption = TSE_Err
          frmWKL00.lbl_TSE.Refresh
          
          
     End If
     
     'immer prüfen, ob es einen TSE-Fehler gab  <<<<< START
      If Not Trim(goTSE.cErrorList) = "" Then
             
            MsgBox ("TSE Fehler:" & vbNewLine & goTSE.cErrorList)
             
      End If
     'immer prüfen, ob es einen TSE-Fehler gab  <<<<< ENDE
     
 Exit Sub
LOKAL_ERROR:
    Screen.MousePointer = 0
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdlEasyTSE"
    Fehler.gsFunktion = "USBSelfTest"
    Fehler.gsFehlertext = "Im Programmteil mdlEasyTSE ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub


 
Sub UpdateTime()
 On Error GoTo LOKAL_ERROR
   
   TSE_OK = True
   TSE_Err = ""
   
   TseEinstellungen.lblZeig.Caption = "TSE: UpdateTime . . ."
   TseEinstellungen.lblZeig.Refresh
   
   frmWKL00.lbl_TSE.Caption = "TSE: UpdateTime . . ."
   frmWKL00.lbl_TSE.Refresh
   
     If Not TSE_InternetZeitAbfragen Then
     
        'bedeutet keine Zeit vom Internet abfragen
        goTSE.cTimeServers = ""
     
     End If
       
     If goTSE.Stack_UpdateTime() = 1 Then
     
         TseEinstellungen.lblZeig.Caption = "TSE ist erfolgreich initialisiert"
         TseEinstellungen.lblZeig.Refresh
         
         frmWKL00.lbl_TSE.Caption = "TSE ist erfolgreich initialisiert"
         frmWKL00.lbl_TSE.Refresh
          
         'wenn der USB-Stick neu ist(der alte USB-Stick ist schon voll und hat keinen Speicherplatz mehr), dann screibe StorageInfo-Daten dieses neues USB-Stick in der Tabelle TSEStorageInfo
          NeuTseUsbStickRegestrieren
         'alle regestrierte Clients in USB abfragen und in Combobox zeigen
          getAllClients
          
         TseEinstellungen.btnNeueClient.Enabled = True
'         TseEinstellungen.btnTSEInfo.Enabled = True
         TseEinstellungen.btnSpeicher.Enabled = True
         TseEinstellungen.comClients.Enabled = True
         frmWKL00.lbl_TSE.Visible = False
         
         Screen.MousePointer = 0
         
         
        Else
        
          TSE_OK = False
          TSE_Err = "TSE: " & goTSE.cErrorList
          
          TseEinstellungen.lblZeig.ForeColor = vbRed
          TseEinstellungen.lblZeig.Caption = TSE_Err
          TseEinstellungen.lblZeig.Refresh
          
          frmWKL00.lbl_TSE.ForeColor = vbRed
          frmWKL00.lbl_TSE.Caption = TSE_Err
          frmWKL00.lbl_TSE.Refresh
          
     End If
     
      
     
     'immer prüfen, ob es einen TSE-Fehler gab  <<<<< START
      If Not Trim(goTSE.cErrorList) = "" Then
             
            MsgBox ("TSE Fehler:" & vbNewLine & goTSE.cErrorList)
             
      End If
     'immer prüfen, ob es einen TSE-Fehler gab  <<<<< ENDE
     
 Exit Sub
LOKAL_ERROR:
    Screen.MousePointer = 0
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdlEasyTSE"
    Fehler.gsFunktion = "UpdateTime"
    Fehler.gsFehlertext = "Im Programmteil mdlEasyTSE ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

 
Sub NeuTseUsbStickRegestrieren()
 On Error GoTo LOKAL_ERROR
    
     
    
    'Serial Nummer erstmal korrigieren  <<<< START
     
     'Hinweis:
     'diese korrektur mache ich, weil ich vorher die FalscheSerialNummer in den Tabellen[ KASSBON,TSEStorageInfo ]gespeichert
     'und jetzt muss ich es hier korrigieren
     
       Dim cmdd As String
       'korrigiere in der Tabelle [ KASSBON ]
       cmdd = "UPDATE KASSBON SET TSESERIAL ='" & USB_serialNumber & "' WHERE TSESERIAL='" & FalschUSB_serialNumber & "'"
       gdBase.Execute cmdd, dbFailOnError
      
       'korrigiere in der Tabelle [ TSEStorageInfo ]
       cmdd = "UPDATE TSEStorageInfo SET SerialNum ='" & USB_serialNumber & "' WHERE SerialNum='" & FalschUSB_serialNumber & "'"
       gdBase.Execute cmdd, dbFailOnError
    
  
    'Serial Nummer erstmal korrigieren  <<<< ENDE
  
  
   'ich habe in einer vorherigen Version vergissen, TSEID in der Tabelle [ KASSBON ] mit den gesamten TSE-Info zu speichern
   'deswegen lese ich TSEID aus der Tabelle [ TSEStorageInfo ] und speichere in [ KASSBON ]
   'durch INNER JOIN
   'Hinweis: das wird nur einmal durchgeführt, deswegen erstelle ich Sperrtabelle nach der Ausführung
    
    If Not NewTableSuchenDB("TSEID_Ist_geschrieben", gdBase) Then
    
      gdBase.Execute "UPDATE KASSBON KB INNER JOIN TSEStorageInfo SI ON KB.TSESERIAL=SI.SerialNum SET KB.TSEID=SI.TSEID", dbFailOnError
      
      'TSEID_Ist_geschrieben ist einfach Sperrtabelle
      gdBase.Execute "Create Table TSEID_Ist_geschrieben(sperrTablle varchar(3))", dbFailOnError
     
    End If
   
  
  'jetzt mach weiter
  Dim schonExistiert As Boolean
  Dim rsrs As Recordset
         
  Set rsrs = gdBase.OpenRecordset("select count(*)as rFlag FROM TSEStorageInfo WHERE SerialNum='" & USB_serialNumber & "'")
   
    If Not rsrs.EOF Then
      
         If Not IsNull(rsrs!rFlag) Then
           
             If rsrs!rFlag > 0 Then
                 'USB_Stick ist schon regestriert
                  schonExistiert = True
             Else
                 'einen neuen USB_Stick regestrieren
                 schonExistiert = False
             End If
           
         End If
     
    End If
   
   
    If schonExistiert Then
   
        'lese TSEID des regestrierten USB_Sticks
        Set rsrs = gdBase.OpenRecordset("select TSEID FROM TSEStorageInfo WHERE SerialNum='" & USB_serialNumber & "'")
        If Not rsrs.EOF Then
           TSE_ID = rsrs!TSEID
        End If
    
    Else
        'einen neuen USB_Stick regestrieren
        Dim sSQL As String
        sSQL = "INSERT INTO TSEStorageInfo ( VendorTyp,SoftVersion,CDCID,CDCHash,SerialNum,SignaturAlg,GueltigBis,PublicKey)VALUES(" & "'" & USB_vendorType & "','" & USB_softwareVersion & "','" & USB_cdcId & "','" & USB_cdcHash & "','" & USB_serialNumber & "','" & USB_signatureAlgorithm & "','" & USB_certificateExpirationDate & "','" & USB_tsePublicKey & "')"
        gdBase.Execute sSQL, dbFailOnError
        'jetzt TSEID lesen
        Set rsrs = gdBase.OpenRecordset("select TSEID FROM TSEStorageInfo WHERE SerialNum='" & USB_serialNumber & "'")
        If Not rsrs.EOF Then
           TSE_ID = rsrs!TSEID
        End If
        
    End If
    
 Exit Sub
LOKAL_ERROR:
    Screen.MousePointer = 0
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdlEasyTSE"
    Fehler.gsFunktion = "NeuTseUsbStickRegestrieren"
    Fehler.gsFehlertext = "Im Programmteil mdlEasyTSE ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub


Sub getAllClients()
 On Error GoTo LOKAL_ERROR
    
    TseEinstellungen.comClients.Clear
    'Clients abfrage
    goTSE.Stack_GetRegisteredClientList
    
    Dim ClientsString As String
    ClientsString = goTSE.cRegisteredClientList ' Abfrage Resultat
    
    
    Dim arrClients() As String
    arrClients = Split(ClientsString, ";")
    
    Dim clientZahl As Integer
    clientZahl = UBound(arrClients)
    
    Dim i As Integer
    
    For i = 0 To clientZahl
     TseEinstellungen.comClients.AddItem (arrClients(i))
    Next
     
    TseEinstellungen.comClients.Text = E_ClientID
    
 Exit Sub
LOKAL_ERROR:
    Screen.MousePointer = 0
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdlEasyTSE"
    Fehler.gsFunktion = "getAllClients"
    Fehler.gsFehlertext = "Im Programmteil mdlEasyTSE ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

 
 
 
Sub Check_TSE_Einstellugen()
On Error GoTo LOCAL_ERROR

    Dim sSQL As String
        
    If Not NewTableSuchenDB("TSESettings", gdApp) Then
     
           'Tabelle TSESettings anlegen
            sSQL = "Create Table TSESettings "
            sSQL = sSQL & "("
            sSQL = sSQL & " SinOrSer varchar(6)"
            sSQL = sSQL & ",TSE_IP varchar(15)"
            sSQL = sSQL & ",TSE_PORT int"
            sSQL = sSQL & ",TSE_AdminPin varchar(5)"
            sSQL = sSQL & ",TSE_TimeAdminPin varchar(5)"
            sSQL = sSQL & ",TSE_PUK varchar(6)"
            sSQL = sSQL & ",TSE_ClientID varchar(30)"
            sSQL = sSQL & ",TSE_DeviceID varchar(30)"
            sSQL = sSQL & ",TSE_SN varchar(80)"
            sSQL = sSQL & ",TSE_ExportPfad varchar(255)"
            sSQL = sSQL & ",TSE_Aktiv bit"
            sSQL = sSQL & ",mitQR bit"
            sSQL = sSQL & ",AltDModus bit"
            sSQL = sSQL & " )"
            gdApp.Execute sSQL, dbFailOnError
          
          
          'INSERT standarde Werte in der Tabelle TSESettings
          Dim DefaultClientID As String
          sSQL = "INSERT INTO TSESettings (TSE_IP,TSE_PORT,TSE_AdminPin,TSE_TimeAdminPin,TSE_PUK,TSE_ClientID,SinOrSer,TSE_DeviceID,TSE_ExportPfad,TSE_Aktiv,mitQR,AltDModus)VALUES('127.0.0.1','8009','12345','54321','123456','','Single','local_TSE','','0','0','0') "
          gdApp.Execute sSQL, dbFailOnError
    
          
          TseEinstellungen.chkTSE_Status.value = vbUnchecked
          HideShowAlleControls False
          TseEinstellungen.chkQrCode.value = vbUnchecked
          E_TSE_Aktiv = False
          TseEinstellungen.txtTSE_IP.Text = "127.0.0.1"
          TseEinstellungen.txtTSE_Port = "8009"
          TseEinstellungen.txtTSE_Adminpin = "12345"
          TseEinstellungen.txtTSE_TimeAdminpin = "54321"
          TseEinstellungen.txtTSE_DeviceID = "local_TSE"
          
          gbTSEExportPfad = ""
          
    Else
    
         If SpalteInTabellegefundenNEW("TSESettings", "TSE_ExportPfad", gdApp) = False Then
                sSQL = "Alter table TSESettings add column TSE_ExportPfad varchar(255)"
                gdApp.Execute sSQL, dbFailOnError
                gbTSEExportPfad = ""
         End If
         
         If SpalteInTabellegefundenNEW("TSESettings", "mitQR", gdApp) = False Then
                sSQL = "Alter table TSESettings add column mitQR bit"
                gdApp.Execute sSQL, dbFailOnError
                MitQrCode = False
         End If
         
         If SpalteInTabellegefundenNEW("TSESettings", "AltDModus", gdApp) = False Then
                sSQL = "Alter table TSESettings add column AltDModus bit"
                gdApp.Execute sSQL, dbFailOnError
                MitQrCode = False
         End If
         
         
         If SpalteInTabellegefundenNEW("TSESettings", "ZeitVomInternet", gdApp) = False Then
                sSQL = "Alter table TSESettings add column ZeitVomInternet bit"
                gdApp.Execute sSQL, dbFailOnError
                TSE_InternetZeitAbfragen = False
         End If
         
         
          Dim rsrs As Recordset
         
            Set rsrs = gdApp.OpenRecordset("select * from TSESettings")
            If Not rsrs.EOF Then
     
                If Not IsNull(rsrs!TSE_AdminPin) Then
                    TseEinstellungen.txtTSE_Adminpin.Text = rsrs!TSE_AdminPin
                End If
     
                If Not IsNull(rsrs!TSE_TimeAdminPin) Then
                     TseEinstellungen.txtTSE_TimeAdminpin.Text = rsrs!TSE_TimeAdminPin
                End If
     
                If Not IsNull(rsrs!TSE_IP) Then
                     TseEinstellungen.txtTSE_IP.Text = rsrs!TSE_IP
                End If
     
                If Not IsNull(rsrs!TSE_Port) Then
                     TseEinstellungen.txtTSE_Port.Text = rsrs!TSE_Port
                End If
                
                If Not IsNull(rsrs!TSE_DeviceID) Then
                     TseEinstellungen.txtTSE_DeviceID.Text = rsrs!TSE_DeviceID
                End If
                
                If Not IsNull(rsrs!TSE_SN) Then
                     TseEinstellungen.txtTSE_SN.Text = rsrs!TSE_SN
                End If
     
                If Not IsNull(rsrs!TSE_ClientID) Then
                    TseEinstellungen.comClients.Text = rsrs!TSE_ClientID
                End If
                
                If Not IsNull(rsrs!TSE_ExportPfad) Then
                    gbTSEExportPfad = rsrs!TSE_ExportPfad
                End If
                
                 If Not IsNull(rsrs!mitQR) Then
                    If rsrs!mitQR = True Then
                        TseEinstellungen.chkQrCode.value = vbChecked
                        MitQrCode = True
                    Else
                        TseEinstellungen.chkQrCode.value = vbUnchecked
                        MitQrCode = False
                    End If
                End If
                
                If Not IsNull(rsrs!TSE_Aktiv) Then

                    If rsrs!TSE_Aktiv = True Then
                    
                      TseEinstellungen.chkTSE_Status.value = vbChecked
                      HideShowAlleControls True
                      
                    Else
                    
                      TseEinstellungen.chkTSE_Status.value = vbUnchecked
                      HideShowAlleControls False
                      
                    End If

                End If
                
                
                If Not IsNull(rsrs!AltDModus) Then

                    If rsrs!AltDModus = True Then
                    
                      TseEinstellungen.chkAltDru.value = vbChecked
                      altDruckModus = True
                    Else
                    
                      TseEinstellungen.chkAltDru.value = vbUnchecked
                      altDruckModus = False
                    End If

                End If
                
               If Not IsNull(rsrs!ZeitVomInternet) Then
               
                    If rsrs!ZeitVomInternet = True Then
                        TseEinstellungen.chkInterntZeit.value = vbChecked
                        TSE_InternetZeitAbfragen = True
                    Else
                        TseEinstellungen.chkInterntZeit.value = vbUnchecked
                        TSE_InternetZeitAbfragen = False
                    End If
                    
              End If
                
    
            End If
            rsrs.Close: Set rsrs = Nothing
             
    End If
    
    
    ''''''''''''''' Check TSEStorageInfo '''''''''''''''''''''''''
    
        If Not NewTableSuchenDB("TSEStorageInfo", gdBase) Then
         
               'Tabelle TSEStorageInfo anlegen
                sSQL = "Create table TSEStorageInfo ( VendorTyp varchar(5),SoftVersion varchar(5),CDCID  varchar(50),CDCHash varchar(150),SerialNum varchar(200),SignaturAlg varchar(25),GueltigBis date,PublicKey varchar(250))"
                gdBase.Execute sSQL, dbFailOnError
        End If
        
        If SpalteInTabellegefundenNEW("TSEStorageInfo", "TSEID", gdBase) = False Then
                    sSQL = "ALTER TABLE TSEStorageInfo ADD COLUMN TSEID AUTOINCREMENT"
                    gdBase.Execute sSQL, dbFailOnError
        End If
    
    ''''''''''''''' Check TSEStorageInfo '''''''''''''''''''''''''
    
    'KASSBON Tabelle prüfen   ( START )
    'mal prüfen, ob die Tabelle KASSBON schon erweitert wurde(erweitert = paar Spalten wurden hinzugefügt)
        
        If SpalteInTabellegefundenNEW("KASSBON", "QRCODE", gdBase) = False Then
                sSQL = "Alter table KASSBON add column QRCODE TEXT"
                gdBase.Execute sSQL, dbFailOnError
        End If
        
        If SpalteInTabellegefundenNEW("KASSBON", "TSESTART", gdBase) = False Then
                sSQL = "Alter table KASSBON add column TSESTART DATETIME"
                gdBase.Execute sSQL, dbFailOnError
        End If
        
        If SpalteInTabellegefundenNEW("KASSBON", "TSEEND", gdBase) = False Then
                sSQL = "Alter table KASSBON add column TSEEND DATETIME"
                gdBase.Execute sSQL, dbFailOnError
        End If
        
        If SpalteInTabellegefundenNEW("KASSBON", "TSESERIAL", gdBase) = False Then
                sSQL = "Alter table KASSBON add column TSESERIAL varchar(80)"
                gdBase.Execute sSQL, dbFailOnError
        End If
        
        If SpalteInTabellegefundenNEW("KASSBON", "TSETRANSACTION", gdBase) = False Then
                sSQL = "Alter table KASSBON add column TSETRANSACTION NUMBER"
                gdBase.Execute sSQL, dbFailOnError
        End If
        
        If SpalteInTabellegefundenNEW("KASSBON", "TSEFEHLER", gdBase) = False Then
                sSQL = "Alter table KASSBON add column TSEFEHLER varchar(200)"
                gdBase.Execute sSQL, dbFailOnError
        End If
        
        If SpalteInTabellegefundenNEW("KASSBON", "TSEClientID", gdBase) = False Then
                sSQL = "Alter table KASSBON add column TSEClientID varchar(30)"
                gdBase.Execute sSQL, dbFailOnError
        End If
        
       If SpalteInTabellegefundenNEW("KASSBON", "TSESTARTSIG", gdBase) = False Then
                sSQL = "Alter table KASSBON add column TSESTARTSIG varchar(128)"
                gdBase.Execute sSQL, dbFailOnError
       End If
        
       If SpalteInTabellegefundenNEW("KASSBON", "TSEFINISHSIG", gdBase) = False Then
                sSQL = "Alter table KASSBON add column TSEFINISHSIG varchar(128)"
                gdBase.Execute sSQL, dbFailOnError
       End If
       
       If SpalteInTabellegefundenNEW("KASSBON", "TSEID", gdBase) = False Then
                sSQL = "Alter table KASSBON add column TSEID NUMBER"
                gdBase.Execute sSQL, dbFailOnError
       End If
       
       If SpalteInTabellegefundenNEW("KASSBON", "STARTSIGZAHLER", gdBase) = False Then
                sSQL = "Alter table KASSBON add column STARTSIGZAHLER NUMBER"
                gdBase.Execute sSQL, dbFailOnError
       End If
       
       If SpalteInTabellegefundenNEW("KASSBON", "FINISHSIGZAHLER", gdBase) = False Then
                sSQL = "Alter table KASSBON add column FINISHSIGZAHLER NUMBER"
                gdBase.Execute sSQL, dbFailOnError
       End If
       
    'KASSBON Tabelle prüfen   ( ENDE )
    
 Exit Sub
    
LOCAL_ERROR:
    
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "mdlEasyTSE"
        Fehler.gsFunktion = "Check_TSE_Einstellugen"
        Fehler.gsFehlertext = "Im Programmteil TseEinstellungen ist ein Fehler aufgetreten."
        
        Fehlermeldung1
        
 End Sub
 
 
 
Sub TSE_SilentInstall()
On Error GoTo LOCAL_ERROR
     
     Dim TSE_SETUP_File1 As String
     Dim TSE_SETUP_File2 As String
     Dim PruefFlag As Boolean
     
     'TSE Silent Installieren, nur wenn die beide TSE_Setup-Dateien vorhanden sind
     TSE_SETUP_File1 = App.Path & "\" & "easytsesetup.exe"
     TSE_SETUP_File2 = App.Path & "\" & "EpsonTSEDriverSetup_1.0.7-4587.exe"
     
     Dim cmdObj As Object
     Set cmdObj = CreateObject("Shell.Application")
      
        If (FileExists(TSE_SETUP_File1)) Then
        
           PruefFlag = True
           
         Else
           PruefFlag = False
           MsgBox ("easytsesetup.exe existiert nicht ! ! !")
           Exit Sub
        End If
        
        
        
        If (FileExists(TSE_SETUP_File2)) Then
        
           PruefFlag = True
           
         Else
           PruefFlag = False
           MsgBox ("EpsonTSEDriverSetup_1.0.7-4587.exe existiert nicht ! ! !")
           Exit Sub
        End If
         
         
      
      
      If PruefFlag Then
       
        'Run (TSE_SETUP_File1) und (TSE_SETUP_File2) als Administrator
        cmdObj.ShellExecute TSE_SETUP_File1, "", , "runas", 1
        cmdObj.ShellExecute TSE_SETUP_File2, "-q EPOSPORT=8009 SHPORT=23500", , "runas", 1
          
      End If
       
 Exit Sub
    
LOCAL_ERROR:
    
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "mdlEasyTSE"
        Fehler.gsFunktion = "TSE_SilentInstall"
        Fehler.gsFehlertext = "Im Programmteil TseEinstellungen ist ein Fehler aufgetreten."
        
        Fehlermeldung1
        
 End Sub
 
 
 
 
'*********************************************************************************************************
'                                        TSE Methoden
'*********************************************************************************************************
 
 
 
 

'*********************************************************************************************************
'neue Client ( Kasse ) Regestrieren
'*********************************************************************************************************
Public Sub NeuClientRegestrieren(ClientName As String)
On Error GoTo LOKAL_ERROR

    
    goTSE.TSE_NewClientID = ClientName
    
    If goTSE.Stack_RegisterNewClient() = 1 Then
        MsgBox ("Client : " & ClientName & " wurde erfolgreich regestriert")
    Else
        MsgBox (goTSE.cErrorList)
    End If

Exit Sub
LOKAL_ERROR:
   
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdlEasyTSE"
    Fehler.gsFunktion = "NeuClientRegestrieren"
    Fehler.gsFehlertext = "Im Programmteil mdlEasyTSE ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub




'*********************************************************************************************************
'Client ( Kasse ) DeRegestrieren
'*********************************************************************************************************
Public Sub ClientDeRegestrieren(ClientName As String)
On Error GoTo LOKAL_ERROR

    
    goTSE.TSE_NewClientID = ClientName
    
    If goTSE.Stack_DeRegisterClient() = 1 Then
        MsgBox ("Client : " & ClientName & " wurde erfolgreich deregestriert")
    Else
        MsgBox (goTSE.cErrorList)
    End If

Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdlEasyTSE"
    Fehler.gsFunktion = "ClientDeRegestrieren"
    Fehler.gsFehlertext = "Im Programmteil mdlEasyTSE ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub






'*********************************************************************************************************
'Transaction (KASSENBON Schreiben)
'*********************************************************************************************************
Public Sub TransactionSchreiben(DataString As String, ProcessTyp As Integer, VorgangTyp As Integer, MWST1 As Double, MWST2 As Double, MWST3 As Double, MWST4 As Double, MWST0 As Double, BetragBar As Double, BetragUnBar As Double)
On Error GoTo LOKAL_ERROR

    R_StartTime = ""
    R_FinishTime = ""
    R_TransactionNr = ""
    R_QRCodeAlsText = ""
    R_QRCodeAlsImgPath = ""
    R_FinishSignatur = ""
    R_StartSignatur = ""
    R_FINISH_SIG_Zaehler = 0
    R_START_SIG_Zaehler = 0
    
    
    'ClientID,TSE_UserID müssen vor dem TransactionsSchreiben gesetzt
 
    goTSE.TSE_ClientID = E_ClientID
    goTSE.TSE_UserID = goTSE.TSE_ClientID
 

    If goTSE.Stack_StartFinishTransaction(DataString, ProcessTyp, VorgangTyp, MWST1, MWST2, MWST3, MWST4, MWST0, BetragBar, BetragUnBar) = 1 Then
         
         R_StartTime = goTSE.oTransActionObjectStart.LogTime
         R_FinishTime = goTSE.oTransActionObjectFinish.LogTime
         R_TransactionNr = goTSE.oTransActionObjectStart.TransactionNumber
         R_QRCodeAlsText = goTSE.cQRCode
         R_QRCodeAlsImgPath = goTSE.cQRCodePath
         R_FinishSignatur = goTSE.oTransActionObjectFinish.Signature
         R_StartSignatur = goTSE.oTransActionObjectStart.Signature
         R_START_SIG_Zaehler = goTSE.oTransActionObjectStart.SignatureCounter
         R_FINISH_SIG_Zaehler = goTSE.oTransActionObjectFinish.SignatureCounter
         
         TSE_OK = True
         TSE_Err = goTSE.cErrorList
         
    Else
    
         'TSE_OK = False
         TSE_Err = goTSE.cErrorList
       
    End If
    
Exit Sub
LOKAL_ERROR:
 
 
 
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdlEasyTSE"
    Fehler.gsFunktion = "TransactionSchreiben"
    Fehler.gsFehlertext = "Im Programmteil mdlEasyTSE ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
 
  
  
End Sub

Sub TSE_Einstellungen_Aktualisieren(neuIP As String, neuPort As String, neuAdminPin As String, neuTimeAdminPin As String, neuClientID As String, neuDeviceID As String, neuSN As String)
On Error GoTo LOKAL_ERROR

  If PruefVerbindungMitNeuerDaten(neuIP, neuPort, neuAdminPin, neuTimeAdminPin) Then
     Dim sSQL As String
     sSQL = "UPDATE TSESettings SET TSE_IP='" & neuIP & "',TSE_PORT='" & neuPort & "',TSE_AdminPin='" & neuAdminPin & "',TSE_TimeAdminPin='" & neuTimeAdminPin & "',TSE_ClientID='" & neuClientID & "',TSE_DeviceID='" & neuDeviceID & "',TSE_SN='" & neuSN & "'"
     gdApp.Execute sSQL, dbFailOnError
     
     E_IP = neuIP
     E_Port = neuPort
     E_AdminPin = neuAdminPin
     E_TimeAdminPin = neuTimeAdminPin
     E_ClientID = neuClientID
     
     MsgBox ("TSE-Einstellungen wurden erfolgreich aktualisiert")
  
  Else
      TseEinstellungen.lblZeig.ForeColor = vbRed
      TseEinstellungen.lblZeig.Caption = goTSE.cErrorList
      TseEinstellungen.lblZeig.Refresh
     'MsgBox ("Keine Verbindung wurde hergestellt !!! " & vbNewLine & "bitte die eingegebene Zugangdaten prüfen")
  End If


 Exit Sub
 
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdlEasyTSE"
    Fehler.gsFunktion = "TSE_Einstellungen_Aktualisieren"
    Fehler.gsFehlertext = "Im Programmteil mdlEasyTSE ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub


Function IstTseInstalliert() As Boolean
On Error GoTo Error:

IstTseInstalliert = False
Dim testActivX As Object
Set testActivX = CreateObject("EasyTSE.EpsonTSE")
IstTseInstalliert = True

Exit Function

Error:
IstTseInstalliert = False

End Function


Function PruefVerbindungMitNeuerDaten(Tip As String, Tport As String, TApin As String, TTApin) As Boolean
 On Error GoTo LOKAL_ERROR
      
     PruefVerbindungMitNeuerDaten = False
      
     goTSE.TSE_IP = Tip
     goTSE.TSE_Port = Tport
     goTSE.TSE_AdminPin = TApin
     goTSE.TSE_TimeAdminPin = TTApin
     
     'nAutoSetup muss auf 1 gesetzt werden
     goTSE.nAutoSetup = 1
     
     If goTSE.TSEConnectOpenSend("GetStorageInfo") = 1 Then
           PruefVerbindungMitNeuerDaten = True
      Else
           PruefVerbindungMitNeuerDaten = False
     End If
     
     Exit Function
LOKAL_ERROR:
     
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdlEasyTSE"
    Fehler.gsFunktion = "PruefVerbindungMitNeuerDaten"
    Fehler.gsFehlertext = "Im Programmteil mdlEasyTSE ist ein Fehler aufgetreten."
    
    Fehlermeldung1
     
     
End Function



Sub HideShowAlleControls(X As Boolean)
 
 If X Then
 
 
 TseEinstellungen.Label1.Visible = True
 TseEinstellungen.Label2.Visible = True
 TseEinstellungen.Label4.Visible = True
 TseEinstellungen.Label5.Visible = True
 TseEinstellungen.Label6.Visible = True
 TseEinstellungen.Label7.Visible = True
 TseEinstellungen.Label8.Visible = True
 
  
 TseEinstellungen.txtTSE_IP.Visible = True
 TseEinstellungen.txtTSE_Port.Visible = True
 TseEinstellungen.txtTSE_Adminpin.Visible = True
 TseEinstellungen.txtTSE_TimeAdminpin.Visible = True
 TseEinstellungen.txtTSE_DeviceID.Visible = True
 TseEinstellungen.comClients.Visible = True
 TseEinstellungen.txtTSE_SN.Visible = True
 
 TseEinstellungen.btnVerbinden.Visible = True
 TseEinstellungen.btnNeueClient.Visible = True
 TseEinstellungen.btnTSEInfo.Visible = True
 TseEinstellungen.btnSpeicher.Visible = True
 TseEinstellungen.btnDatenExport.Visible = True
 TseEinstellungen.lblZeig.Visible = True
 
 TseEinstellungen.chkQrCode.Visible = True
 TseEinstellungen.chkInterntZeit.Visible = True
 
 
 Else
 
 
 TseEinstellungen.Label1.Visible = False
 TseEinstellungen.Label2.Visible = False
 TseEinstellungen.Label4.Visible = False
 TseEinstellungen.Label5.Visible = False
 TseEinstellungen.Label6.Visible = False
 TseEinstellungen.Label7.Visible = False
 TseEinstellungen.Label8.Visible = False
 
 TseEinstellungen.txtTSE_IP.Visible = False
 TseEinstellungen.txtTSE_Port.Visible = False
 TseEinstellungen.txtTSE_Adminpin.Visible = False
 TseEinstellungen.txtTSE_TimeAdminpin.Visible = False
 TseEinstellungen.txtTSE_DeviceID.Visible = False
 TseEinstellungen.comClients.Visible = False
 TseEinstellungen.txtTSE_SN.Visible = False
 
 TseEinstellungen.btnVerbinden.Visible = False
 TseEinstellungen.btnNeueClient.Visible = False
 TseEinstellungen.btnTSEInfo.Visible = False
 TseEinstellungen.btnSpeicher.Visible = False
 TseEinstellungen.btnDatenExport.Visible = False
 TseEinstellungen.lblZeig.Visible = False
 
 TseEinstellungen.chkQrCode.Visible = False
 TseEinstellungen.chkInterntZeit.Visible = False
 
 frmWKL00.lbl_TSE.Visible = False
 
 End If
 
 
 
End Sub


'QRadapter.exe ist eine VB.Net Programm, das die von USB-Stick lieferte QR-Code druckt
Public Sub QRcodeDrucken()
On Error GoTo LOKAL_ERROR

  If Trim(R_QRCodeAlsImgPath) <> "" Then

      If (FileExists(App.Path & "\" & "QRadapter.exe")) Then
        Shell App.Path & "\" & "QRadapter.exe " & R_QRCodeAlsImgPath & "?" & gcBonDrucker & "?" & "", vbMinimizedFocus
      Else
         MsgBox ("QRadapter.exe existiert leider nicht ! ! !")
      End If
      
  End If


Exit Sub

LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdlEasyTSE"
    Fehler.gsFunktion = "QRcodeDrucken"
    Fehler.gsFehlertext = "Im Programmteil mdlEasyTSE ist ein Fehler aufgetreten."
    
    Fehlermeldung1
     
     
End Sub





'QRadapter.exe schneidet das Papier auch
Public Sub CutPapier()
On Error GoTo LOKAL_ERROR
 
      If (FileExists(App.Path & "\" & "QRadapter.exe")) Then
        Shell App.Path & "\" & "QRadapter.exe " & "" & "?" & gcBonDrucker & "?" & "SchneideKassenbonBitte", vbMinimizedFocus
      Else
        MsgBox ("QRadapter.exe existiert leider nicht ! ! !")
      End If
 
Exit Sub

LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdlEasyTSE"
    Fehler.gsFunktion = "CutPapier"
    Fehler.gsFehlertext = "Im Programmteil mdlEasyTSE ist ein Fehler aufgetreten."
    
    Fehlermeldung1
     
     
End Sub



'QRadapter.exe druckt paar leere zeilen
Public Sub PaarLeereZeilenDrucken()
On Error GoTo LOKAL_ERROR
 
      If (FileExists(App.Path & "\" & "QRadapter.exe")) Then
        Shell App.Path & "\" & "QRadapter.exe " & "" & "?" & gcBonDrucker & "?" & "PaarLeereZeilen", vbMinimizedFocus
      Else
         MsgBox ("QRadapter.exe existiert leider nicht ! ! !")
      End If
 
Exit Sub

LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdlEasyTSE"
    Fehler.gsFunktion = "PaarLeereZeilenDrucken"
    Fehler.gsFehlertext = "Im Programmteil mdlEasyTSE ist ein Fehler aufgetreten."
    
    Fehlermeldung1
     
     
End Sub


Function SplitStringNachCharZahl(spaceVonLinks As Integer, Gstr As String, Gzahl) As String
On Error GoTo LOKAL_ERROR

'Hier wird [Gstr] in [Gzahl]-char gruppe zusammengestellt mit spaceVonLinks
'Beispiel : wenn Gstr = "HalloAlle" ,Gzahl=3 und spaceVonLinks= 2
'dann Resultat:
'  Hal
'  loA
'  lle

'Beispie2 : wenn Gstr = "HalloAlle" ,Gzahl=3 und spaceVonLinks= 0
'dann Resultat:
'Hal
'loA
'lle

'Beispie3 : wenn Gstr = "HalloAlle" ,Gzahl=4 und spaceVonLinks= 0
'dann Resultat:
'Hall
'oAll
'e

'Beispie4 : wenn Gstr = "HalloAlle" ,Gzahl=6 und spaceVonLinks= 4
'dann Resultat:
'    HalloA
'    lle


 
'...............usw



    SplitStringNachCharZahl = ""

    Dim Tmp As String
    Dim i As Integer
    Dim cout As Integer
     
    Dim lAnzZeile As Integer
    Dim arr() As String
    
    Tmp = ""
    cout = 1
     
    
    For i = 1 To Len(Gstr)
      
      If cout = Gzahl Then
      
         Tmp = Tmp & Mid(Gstr, i, 1)
         Tmp = Space(spaceVonLinks) & Tmp
         lAnzZeile = lAnzZeile + 1
         ReDim Preserve arr(1 To lAnzZeile) As String
         arr(lAnzZeile) = Tmp
         Tmp = ""
         cout = 1
         
       Else
       
         Tmp = Tmp & Mid(Gstr, i, 1)
         cout = cout + 1
         
      End If
       
    Next i
    
    If Trim(Tmp) <> "" Then
    
     Tmp = Space(spaceVonLinks) & Tmp
     lAnzZeile = lAnzZeile + 1
     ReDim Preserve arr(1 To lAnzZeile) As String
     arr(lAnzZeile) = Tmp
    
    End If
    
   
    Tmp = ""
   
    For i = 1 To UBound(arr)
     Tmp = Tmp & arr(i) & vbNewLine
    Next i
    
    SplitStringNachCharZahl = Tmp
    Exit Function
 
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdlEasyTSE"
    Fehler.gsFunktion = "SplitStringNachCharZahl"
    Fehler.gsFehlertext = "Im Programmteil mdlEasyTSE ist ein Fehler aufgetreten."
    
    Fehlermeldung1
 
End Function

 

'********************************************************************************************************'
'Daten-Exportieren (Alles)                                                                               '
'********************************************************************************************************'
Public Sub DatenExportieren()
On Error GoTo LOKAL_ERROR
                                                                                                         
   Screen.MousePointer = 11
                                                                                                         
   goTSE.cExportDir = gbTSEExportPfad
   goTSE.TSE_ClientID = E_ClientID
                                                                                                         
   TSEDataExport.lblExportStatus.Caption = "Daten werden aus USB-Stick exportiert . . ."
   TSEDataExport.lblExportStatus.Refresh
                                                                                                         
    If goTSE.Stack_ExportArchiveData() = 1 Then
      TSEDataExport.lblExportStatus.Caption = "Fertig"
      TSEDataExport.lblExportStatus.Refresh
                                                                                                         
      TSEDataExport.btnExport.Enabled = True
                                                                                                         
    Else
      TSEDataExport.lblExportStatus.Caption = goTSE.cErrorList
      TSEDataExport.lblExportStatus.Refresh
                                                                                                         
       TSEDataExport.btnExport.Enabled = True
    End If
                                                                                                          
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:

    Screen.MousePointer = 0
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdlEasyTSE"
    Fehler.gsFunktion = "DatenExportieren"
    Fehler.gsFehlertext = "Im Programmteil mdlEasyTSE ist ein Fehler aufgetreten."
                                                                                                         
    Fehlermeldung1
                                                                                                         
End Sub



'********************************************************************************************************
'Daten-Exportieren(zwischen TransNr Interval)
'********************************************************************************************************
Public Sub DatenExportierenNachTransNr(StartNr As Long, EndNr As Long)
On Error GoTo LOKAL_ERROR
                                                                                                         
   Screen.MousePointer = 11
   goTSE.cExportDir = gbTSEExportPfad
   goTSE.TSE_ClientID = E_ClientID
                                                                                                         
   goTSE.TSE_StartTransactionNumber = StartNr
   goTSE.TSE_EndTransactionNumber = EndNr
                                                                                                         
   TSEDataExport.lblExportStatus.Caption = "Daten werden aus USB-Stick exportiert . . ."
   TSEDataExport.lblExportStatus.Refresh
                                                                                                         
    If goTSE.Stack_ExportFilteredByTransactionNumberInterval() = 1 Then
      TSEDataExport.lblExportStatus.Caption = "Fertig"
      TSEDataExport.lblExportStatus.Refresh
      TSEDataExport.btnExport.Enabled = True
                                                                                                         
    Else
      TSEDataExport.lblExportStatus.Caption = goTSE.cErrorList
      TSEDataExport.lblExportStatus.Refresh
      TSEDataExport.btnExport.Enabled = True
                                                                                                         
    End If
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:

    Screen.MousePointer = 0
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdlEasyTSE"
    Fehler.gsFunktion = "DatenExportierenNachTransNr"
    Fehler.gsFehlertext = "Im Programmteil mdlEasyTSE ist ein Fehler aufgetreten."
                                                                                                         
    Fehlermeldung1
                                                                                                         
End Sub


'*********************************************************************************************************
'Daten-Exportieren(für einen beliebigen Zeitraum)
'*********************************************************************************************************
Public Sub DatenExportierenNachZeitraum(StartDatum As String, EndDatum As String)
On Error GoTo LOKAL_ERROR
   
   Screen.MousePointer = 11
   
   goTSE.cExportDir = gbTSEExportPfad
   goTSE.TSE_ClientID = E_ClientID
   
   goTSE.TSE_StartDate = StartDatum
   goTSE.TSE_EndDate = EndDatum
   
   TSEDataExport.lblExportStatus.Caption = "Daten werden aus USB-Stick exportiert . . ."
   TSEDataExport.lblExportStatus.Refresh
   
    If goTSE.Stack_ExportFilteredByPeriodOfTime() = 1 Then
      TSEDataExport.lblExportStatus.Caption = "Fertig"
      TSEDataExport.lblExportStatus.Refresh
      TSEDataExport.btnExport.Enabled = True
      
    Else
      TSEDataExport.lblExportStatus.Caption = goTSE.cErrorList
      TSEDataExport.lblExportStatus.Refresh
      TSEDataExport.btnExport.Enabled = True
      
    End If
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:

    Screen.MousePointer = 0
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdlEasyTSE"
    Fehler.gsFunktion = "DatenExportierenNachZeitraum"
    Fehler.gsFehlertext = "Im Programmteil mdlEasyTSE ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
 
 
