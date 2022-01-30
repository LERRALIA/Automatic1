VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form TSEStorageInfo 
   BackColor       =   &H00C0C000&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "TSE_StorageInfo"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   9885
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command1 
      Caption         =   "Drucken"
      Height          =   255
      Left            =   7920
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid dgvInfo 
      Height          =   3885
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   6853
      _Version        =   393216
      Rows            =   1
      Cols            =   3
      RowHeightMin    =   300
      AllowUserResizing=   3
   End
End
Attribute VB_Name = "TSEStorageInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 
 reportbildschirm "", "rptTSEinfo"
  
End Sub

Private Sub Form_Load()

    dgvInfo.TextMatrix(0, 1) = "Eigenschaft"
    dgvInfo.TextMatrix(0, 2) = "Wert"
    
    dgvInfo.ColWidth(1) = 2400
    dgvInfo.ColWidth(2) = 4000
     
    dgvInfo.ColAlignment(1) = flexAlignLeftCenter
    dgvInfo.ColAlignment(2) = flexAlignLeftCenter
    
    ZeigAlleStorageInfo
End Sub

Sub ZeigAlleStorageInfo()
    
    dgvInfo.Rows = dgvInfo.Rows + 1
    Dim newRow As Integer
    newRow = dgvInfo.Rows - 1
    
    dgvInfo.TextMatrix(newRow, 0) = newRow
    dgvInfo.TextMatrix(newRow, 1) = "cdcId"
    dgvInfo.TextMatrix(newRow, 2) = USB_cdcId
    
    dgvInfo.Rows = dgvInfo.Rows + 1
    newRow = dgvInfo.Rows - 1
    
    dgvInfo.TextMatrix(newRow, 0) = newRow
    dgvInfo.TextMatrix(newRow, 1) = "cdcHash"
    dgvInfo.TextMatrix(newRow, 2) = USB_cdcHash
    
    dgvInfo.Rows = dgvInfo.Rows + 1
    newRow = dgvInfo.Rows - 1
    
    dgvInfo.TextMatrix(newRow, 0) = newRow
    dgvInfo.TextMatrix(newRow, 1) = "certificateExpirationDate"
    dgvInfo.TextMatrix(newRow, 2) = USB_certificateExpirationDate
    
    dgvInfo.Rows = dgvInfo.Rows + 1
    newRow = dgvInfo.Rows - 1
    
    dgvInfo.TextMatrix(newRow, 0) = newRow
    dgvInfo.TextMatrix(newRow, 1) = "createdSignatures"
    dgvInfo.TextMatrix(newRow, 2) = USB_createdSignatures
    dgvInfo.Rows = dgvInfo.Rows + 1
    newRow = dgvInfo.Rows - 1
    
    dgvInfo.TextMatrix(newRow, 0) = newRow
    dgvInfo.TextMatrix(newRow, 1) = "hardwareVersion"
    dgvInfo.TextMatrix(newRow, 2) = USB_hardwareVersion

    dgvInfo.Rows = dgvInfo.Rows + 1
    newRow = dgvInfo.Rows - 1
    
    dgvInfo.TextMatrix(newRow, 0) = newRow
    dgvInfo.TextMatrix(newRow, 1) = "hasPassedSelfTest"
    dgvInfo.TextMatrix(newRow, 2) = USB_hasPassedSelfTest
    
    dgvInfo.Rows = dgvInfo.Rows + 1
    newRow = dgvInfo.Rows - 1
    
    dgvInfo.TextMatrix(newRow, 0) = newRow
    dgvInfo.TextMatrix(newRow, 1) = "hasValidTime"
    dgvInfo.TextMatrix(newRow, 2) = USB_hasValidTime
    
    dgvInfo.Rows = dgvInfo.Rows + 1
    newRow = dgvInfo.Rows - 1
    
    dgvInfo.TextMatrix(newRow, 0) = newRow
    dgvInfo.TextMatrix(newRow, 1) = "isExportEnabledIfCspTestFails"
    dgvInfo.TextMatrix(newRow, 2) = USB_isExportEnabledIfCspTestFails

    dgvInfo.Rows = dgvInfo.Rows + 1
    newRow = dgvInfo.Rows - 1
    
    dgvInfo.TextMatrix(newRow, 0) = newRow
    dgvInfo.TextMatrix(newRow, 1) = "isTSEUnlocked"
    dgvInfo.TextMatrix(newRow, 2) = USB_isTSEUnlocked
    
    dgvInfo.Rows = dgvInfo.Rows + 1
    newRow = dgvInfo.Rows - 1
    
    dgvInfo.TextMatrix(newRow, 0) = newRow
    dgvInfo.TextMatrix(newRow, 1) = "lastExportExecutedDate"
    dgvInfo.TextMatrix(newRow, 2) = USB_lastExportExecutedDate
 
    dgvInfo.Rows = dgvInfo.Rows + 1
    newRow = dgvInfo.Rows - 1
    
    dgvInfo.TextMatrix(newRow, 0) = newRow
    dgvInfo.TextMatrix(newRow, 1) = "maxRegisteredClients"
    dgvInfo.TextMatrix(newRow, 2) = USB_maxRegisteredClients
    
    dgvInfo.Rows = dgvInfo.Rows + 1
    newRow = dgvInfo.Rows - 1
    
    dgvInfo.TextMatrix(newRow, 0) = newRow
    dgvInfo.TextMatrix(newRow, 1) = "maxSignatures"
    dgvInfo.TextMatrix(newRow, 2) = USB_maxSignatures
    
    dgvInfo.Rows = dgvInfo.Rows + 1
    newRow = dgvInfo.Rows - 1
    
    dgvInfo.TextMatrix(newRow, 0) = newRow
    dgvInfo.TextMatrix(newRow, 1) = "maxStartedTransactions"
    dgvInfo.TextMatrix(newRow, 2) = USB_maxStartedTransactions
    
    dgvInfo.Rows = dgvInfo.Rows + 1
    newRow = dgvInfo.Rows - 1
    
    dgvInfo.TextMatrix(newRow, 0) = newRow
    dgvInfo.TextMatrix(newRow, 1) = "maxUpdateDelay"
    dgvInfo.TextMatrix(newRow, 2) = USB_maxUpdateDelay
   
    dgvInfo.Rows = dgvInfo.Rows + 1
    newRow = dgvInfo.Rows - 1
    
    dgvInfo.TextMatrix(newRow, 0) = newRow
    dgvInfo.TextMatrix(newRow, 1) = "registeredClients"
    dgvInfo.TextMatrix(newRow, 2) = USB_registeredClients
    
    dgvInfo.Rows = dgvInfo.Rows + 1
    newRow = dgvInfo.Rows - 1
    
    dgvInfo.TextMatrix(newRow, 0) = newRow
    dgvInfo.TextMatrix(newRow, 1) = "remainingSignatures"
    dgvInfo.TextMatrix(newRow, 2) = USB_remainingSignatures
    
    dgvInfo.Rows = dgvInfo.Rows + 1
    newRow = dgvInfo.Rows - 1
    
    dgvInfo.TextMatrix(newRow, 0) = newRow
    dgvInfo.TextMatrix(newRow, 1) = "SerialNumber"
    dgvInfo.TextMatrix(newRow, 2) = USB_serialNumber
 
    dgvInfo.Rows = dgvInfo.Rows + 1
    newRow = dgvInfo.Rows - 1
    
    dgvInfo.TextMatrix(newRow, 0) = newRow
    dgvInfo.TextMatrix(newRow, 1) = "signatureAlgorithm"
    dgvInfo.TextMatrix(newRow, 2) = USB_signatureAlgorithm
    
    dgvInfo.Rows = dgvInfo.Rows + 1
    newRow = dgvInfo.Rows - 1
    
    dgvInfo.TextMatrix(newRow, 0) = newRow
    dgvInfo.TextMatrix(newRow, 1) = "softwareVersion"
    dgvInfo.TextMatrix(newRow, 2) = USB_softwareVersion
    
    dgvInfo.Rows = dgvInfo.Rows + 1
    newRow = dgvInfo.Rows - 1
    
    dgvInfo.TextMatrix(newRow, 0) = newRow
    dgvInfo.TextMatrix(newRow, 1) = "startedTransactions"
    dgvInfo.TextMatrix(newRow, 2) = USB_startedTransactions
     
    dgvInfo.Rows = dgvInfo.Rows + 1
    newRow = dgvInfo.Rows - 1
    
    dgvInfo.TextMatrix(newRow, 0) = newRow
    dgvInfo.TextMatrix(newRow, 1) = "tarExportSize"
    dgvInfo.TextMatrix(newRow, 2) = USB_tarExportSize
    
    dgvInfo.Rows = dgvInfo.Rows + 1
    newRow = dgvInfo.Rows - 1
    
    dgvInfo.TextMatrix(newRow, 0) = newRow
    dgvInfo.TextMatrix(newRow, 1) = "timeUntilNextSelfTest"
    dgvInfo.TextMatrix(newRow, 2) = USB_timeUntilNextSelfTest
     
    dgvInfo.Rows = dgvInfo.Rows + 1
    newRow = dgvInfo.Rows - 1
    
    dgvInfo.TextMatrix(newRow, 0) = newRow
    dgvInfo.TextMatrix(newRow, 1) = "tseCapacity"
    dgvInfo.TextMatrix(newRow, 2) = USB_tseCapacity

    dgvInfo.Rows = dgvInfo.Rows + 1
    newRow = dgvInfo.Rows - 1
    
    dgvInfo.TextMatrix(newRow, 0) = newRow
    dgvInfo.TextMatrix(newRow, 1) = "tseCurrentSize"
    dgvInfo.TextMatrix(newRow, 2) = USB_tseCurrentSize
    
    dgvInfo.Rows = dgvInfo.Rows + 1
    newRow = dgvInfo.Rows - 1
    
    dgvInfo.TextMatrix(newRow, 0) = newRow
    dgvInfo.TextMatrix(newRow, 1) = "tseDescription"
    dgvInfo.TextMatrix(newRow, 2) = USB_tseDescription
 
    dgvInfo.Rows = dgvInfo.Rows + 1
    newRow = dgvInfo.Rows - 1
    
    dgvInfo.TextMatrix(newRow, 0) = newRow
    dgvInfo.TextMatrix(newRow, 1) = "tseInitializationState"
    dgvInfo.TextMatrix(newRow, 2) = USB_tseInitializationState
    
    dgvInfo.Rows = dgvInfo.Rows + 1
    newRow = dgvInfo.Rows - 1
    
    dgvInfo.TextMatrix(newRow, 0) = newRow
    dgvInfo.TextMatrix(newRow, 1) = "tsePublicKey"
    dgvInfo.TextMatrix(newRow, 2) = USB_tsePublicKey
    
    dgvInfo.Rows = dgvInfo.Rows + 1
    newRow = dgvInfo.Rows - 1
    
    dgvInfo.TextMatrix(newRow, 0) = newRow
    dgvInfo.TextMatrix(newRow, 1) = "vendorType"
    dgvInfo.TextMatrix(newRow, 2) = USB_vendorType
  

End Sub
