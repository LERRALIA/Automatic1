VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
    Private data() As Byte
    Public position As Long

    Property Get Size() As Long
        If Not (Not data) Then
            Size = UBound(data) - LBound(data) + 1
        Else
            Size = 0
        End If
    End Property

    Private Sub Class_Initialize()
        position = 0
    End Sub
    
    Public Sub Clear()
        position = 0
        Erase data
    End Sub
    
    Public Function GetData() As Byte()
        GetData = data
    End Function
    
    Public Sub SetData(ByRef sourceData() As Byte)
        Erase data
        length = UBound(sourceData) - LBound(sourceData) + 1
        For n = 0 To length - 1
            WriteByte sourceData(n)
        Next n
        position = 0
    End Sub
    
    Public Function ReadByte() As Byte
        Dim result As Byte
        result = data(position)
        position = position + 1
        ReadByte = result
    End Function

    Public Function ReadInteger() As Integer
        Dim result As Integer
        If data(position + 1) > 127 Then
            result = (((data(position + 1) And &H7F) * &H100) Or &H80000) Or data(position)
        Else
            result = (data(position + 1) * &H100) Or data(position)
        End If
        position = position + 2
        ReadInteger = result
    End Function
    
    Public Function ReadLong() As Long
        Dim result As Long
        If data(position + 3) > 127 Then
            result = (((data(position + 3) And &H7F) * &H1000000) Or &H80000000) Or _
                   (data(position + 2) * &H10000) Or _
                   (data(position + 1) * CLng(&H100)) Or data(position)
        Else
            result = (data(position + 3) * &H1000000) Or _
                   (data(position + 2) * &H10000) Or _
                   (data(position + 1) * CLng(&H100)) Or data(position)
        End If
        position = position + 4
        ReadLong = result
    End Function
    
    Public Sub WriteByte(ByVal value As Byte)
        ReDim Preserve data(Size)
        data(position) = value
        position = position + 1
    End Sub
    
    Public Sub WriteInteger(ByVal value As Integer)
        ReDim Preserve data(Size + 1)
        data(position) = CByte((value And &HFF) And &HFF)
        data(position + 1) = CByte(((value And &HFF00) \ &H100) And &HFF)
        position = position + 2
    End Sub
    
    Public Sub WriteLong(ByVal value As Long)
        ReDim Preserve data(Size + 3)
        data(position) = CByte((value And &HFF) And &HFF)
        data(position + 1) = CByte(((value And &HFF00) \ &H100) And &HFF)
        data(position + 2) = CByte(((value And &HFF0000) \ &H10000) And &HFF)
        data(position + 3) = CByte(((value And &HFF000000) \ &H1000000) And &HFF)
        position = position + 4
    End Sub
