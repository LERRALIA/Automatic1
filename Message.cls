VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Message"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
    Public Enum MessageConstants
        messageDelimiter = &H7E
        MessageEscape = &HAC
    End Enum
        
    Private Const hashId As Long = 0
    Private Const startIndexCrc As Integer = 8
    
    Public version As Integer
    Private Segments() As Segment
    
    Private Sub Class_Initialize()
        version = 1
        ReDim Preserve Segments(0)
    End Sub
    
    Public Property Get CurrentVersion() As Integer
        CurrentVersion = 1
    End Property
    
    Public Property Get Size() As Long
        Size = UBound(Segments) - LBound(Segments) + 1
    End Property

    Public Property Get Item(ByVal index As Integer) As Segment
        Set Item = Segments(index)
    End Property
    
    Public Property Let Item(ByVal index As Integer, ByVal value As Segment)
        If index > Size Then
            ReDim Preserve Segments(index + 1)
        End If
        Set Segments(index) = value
    End Property
