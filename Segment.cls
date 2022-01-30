VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Segment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
    Public Enum SegmentCategory
        SegmentHeartbeat = &H0
        SegmentResponse = &H80
        SegmentTSSStatus = &HE0
        SegmentTSSStart = &HE1
        SegmentTSSCancel = &HE2
        SegmentTSSFinish = &HE3
        SegmentTSSTransaction = &HE4
    End Enum
    
    Public Enum ResponseCode
        CodeInternalCatalogueMissmatch = &H2000
        CodeInternalSegmentIdMissmatch = &H2001
        CodeCatalogueUnknown = &H2010
        CodeCatalogueUnsupported = &H2011
        CodeSegmentUnknown = &H2020
        CodeSegmentUnsupported = &H2021
        CodeCrcCheckForwardFailed = &H4000
        CodeCrcCheckBackwardFailed = &H4001
        CodeMessageSizeTooSmall = &H4010
        CodeMessageSizeMissmatch = &H4011
        CodeMessageVersionUnsupported = &H4012
        CodeSegmentSizeTooSmall = &H4020
        CodeSegmentSizeMissmatch = &H4021
        CodeSegmentVersionUnsupported = &H4022
        CodeUnknownError = &H4F00
        CodeOK = &H8000
    End Enum
    
    Public id As SegmentCategory
    Public Session As Long
    Public version As Integer
    Public sequence As Long
    Public Parameter As String
    Public Code As ResponseCode
    
    Property Get CurrentVersion() As Integer
        CurrentVersion = 1
    End Property
    
    Private Sub Class_Initialize()
        Session = 0
        sequence = 0
    End Sub
    
    Public Function ToString() As String
        ToString = "[0x00] " & Hex$(id) & " V" & version
    End Function
