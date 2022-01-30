Attribute VB_Name = "NetworkDataConverter"
    Private Const messageBaseSize As Long = 16
    Private Const segmentBaseSize As Long = 12
    Private Const startIndexCrc As Long = 8
    
    Private isCrc32Initialized As Boolean
    Private crc32LookUp() As Long
    
    Public Function DecodeMessage(ByVal socket As Winsock, ByVal sessionId As Long, ByRef data() As Byte) As Message
        Dim result As New Message
        data = Deescape(data)
        Dim length As Long: length = UBound(data) - LBound(data) + 1
        
        If length < messageBaseSize Then
            SendResponse socket, CreateResponse(sessionId, CodeMessageSizeTooSmall)
            Exit Function
        End If
        
        Dim handler As New DataHandler
        handler.SetData data
        Dim crc32f As Long: crc32f = handler.ReadLong()
        Dim crc32b As Long: crc32b = handler.ReadLong()
        Dim computedCrc32f As Long: computedCrc32f = ComputeHash(data, startIndexCrc, False)
        Dim computedCrc32b As Long: computedCrc32b = ComputeHash(data, startIndexCrc, True)
        
        If computedCrc32f <> crc32f Then
            SendResponse socket, CreateResponse(sessionId, CodeCrcCheckForwardFailed)
            Exit Function
        ElseIf computedCrc32b <> crc32b Then
            SendResponse socket, CreateResponse(sessionId, CodeCrcCheckBackwardFailed)
            Exit Function
        End If
        
        Dim Size As Long: Size = handler.ReadLong()
        If Size <> length Then
            SendResponse socket, CreateResponse(sessionId, CodeMessageSizeMissmatch)
            Exit Function
        End If
        
        result.version = handler.ReadInteger()
        If result.version > result.CurrentVersion Then
            SendResponse socket, CreateResponse(sessionId, CodeMessageVersionUnsupported)
            Exit Function
        End If
        
        Dim segmentCount As Integer: segmentCount = handler.ReadInteger()
        For N = 0 To segmentCount - 1
            Dim segmentData() As Byte
            Dim position As Long: position = handler.position
            Dim segmentSize As Long: segmentSize = handler.ReadLong()
            ReDim segmentData(segmentSize - 1)
            handler.position = position
            For i = 0 To segmentSize - 1
                segmentData(i) = handler.ReadByte()
            Next i
            Dim innerSegment As Segment: Set innerSegment = DecodeSegment(sessionId, segmentData)
            result.Item(N) = innerSegment
        Next N
        
        Set DecodeMessage = result
    End Function
    
    Public Function DecodeSegment(ByVal sessionId As Long, ByRef data() As Byte) As Segment
        Dim length As Long: length = UBound(data) - LBound(data) + 1
        
        If length < segmentBaseSize Then
            Set DecodeSegment = CreateResponse(sessionId, CodeSegmentSizeTooSmall)
            Exit Function
        End If
        
        Dim handler As New DataHandler
        handler.SetData data
        
        Dim result As New Segment
        Dim Size As Long: Size = handler.ReadLong()
        Dim catalogue As Byte: catalogue = handler.ReadByte()
        Dim id As Byte: id = handler.ReadByte()
        
        If Not IsValidId(id) Then
            Set DecodeSegment = CreateResponse(sessionId, CodeSegmentUnsupported)
            Exit Function
        End If
        
        result.id = id
        result.Session = sessionId
        result.version = handler.ReadInteger()
        result.sequence = handler.ReadLong()
        
        If Size <> length Then
            Set DecodeSegment = CreateResponse(sessionId, CodeSegmentSizeMissmatch)
            Exit Function
        ElseIf catalogue <> 0 Then
            Set DecodeSegment = CreateResponse(sessionId, CodeCatalogueUnsupported)
            Exit Function
        ElseIf result.version > result.CurrentVersion Then
            Set DecodeSegment = CreateResponse(sessionId, CodeSegmentVersionUnsupported)
            Exit Function
        End If
        
        Set DecodeSegment = result
    End Function
    
    Public Function EncodeMessage(ByVal msg As Message) As Byte()
        Dim Size As Long: Size = messageBaseSize
        Dim segmentCount As Long: segmentCount = msg.Size
        Dim segmentData() As Byte
        Dim position As Long
        Dim tempData() As Byte
        Dim tempLength As Long
        
        For N = 0 To segmentCount - 1
            Dim tempSegment As Segment
            Set tempSegment = msg.Item(N)
            tempData = EncodeSegment(tempSegment)
            tempLength = UBound(tempData) - LBound(tempData) + 1
            Size = Size + tempLength
            ReDim Preserve segmentData(position + tempLength - 1)
            For X = 0 To tempLength - 1
                segmentData(position) = tempData(X)
                position = position + 1
            Next X
        Next N
        
        Dim handler As New DataHandler
        handler.WriteLong Size
        handler.WriteInteger msg.version
        handler.WriteInteger segmentCount
        tempLength = UBound(segmentData) - LBound(segmentData) + 1
        For N = 0 To tempLength - 1
            handler.WriteByte segmentData(N)
        Next N
        
        Dim messageData() As Byte: messageData = handler.GetData()
        handler.Clear
        
        handler.WriteLong ComputeHash(messageData, 0, False)
        handler.WriteLong ComputeHash(messageData, 0, True)
        tempLength = UBound(messageData) - LBound(messageData) + 1
        For N = 0 To tempLength - 1
            handler.WriteByte messageData(N)
        Next N
        messageData = Escape(handler.GetData())
        tempLength = UBound(messageData) - LBound(messageData) + 1
        ReDim Preserve messageData(tempLength)
        messageData(tempLength) = MessageDelimiter

        handler.Clear
        
        handler.WriteByte MessageDelimiter
        tempLength = UBound(messageData) - LBound(messageData) + 1
        For N = 0 To tempLength - 1
            handler.WriteByte messageData(N)
        Next N
        
        EncodeMessage = handler.GetData()
    End Function
    
    Public Function EncodeSegment(ByVal Source As Segment) As Byte()
        Dim handler As New DataHandler
        Dim Size As Long: Size = segmentBaseSize
        Dim data() As Byte: data = EncodeSegmentParameters(Source)
        Dim length As Long
        
        If Not (Not data) Then
            length = UBound(data) - LBound(data) + 1
        Else
            length = 0
        End If
        
        If length > 0 Then
            Size = Size + length
        End If
        
        handler.WriteLong Size
        handler.WriteByte 0
        handler.WriteByte Source.id
        handler.WriteInteger Source.version
        handler.WriteLong Source.sequence

        For N = 0 To length - 1
            handler.WriteByte data(N)
        Next N

        EncodeSegment = handler.GetData()
    End Function
    
    Private Function EncodeSegmentParameters(ByVal Source As Segment) As Byte()
        Dim result() As Byte
        If Source.id = SegmentResponse Then
            Dim handler As New DataHandler
            handler.WriteInteger Source.Code
            result = handler.GetData()
        ElseIf Source.id = SegmentTSSFinish Then
            Dim length As Long: length = Len(Source.Parameter)
            If length > 0 Then
                ReDim result(length - 1)
                For N = 1 To length
                    result(N - 1) = Asc(Mid(Source.Parameter, N, 1))
                Next N
            End If
        End If
        
        EncodeSegmentParameters = result
    End Function
    
    Private Function Escape(ByRef data() As Byte) As Byte()
        Dim insertAt() As Long
        Dim index As Long: index = -1
        Dim length As Long: length = UBound(data) - LBound(data) + 1
        
        For N = 0 To length - 1
            If data(N) = MessageDelimiter Or data(N) = MessageEscape Then
                ReDim Preserve insertAt(index + 1)
                index = index + 1
                insertAt(index) = N
                isSpecialCharFound = False
            End If
        Next N
        
        For i = 0 To index
            length = UBound(data) - LBound(data) + 1
            For N = length To insertAt(i) + 1 Step -1
                data(N) = data(N - 1)
            Next N
            data(insertAt(i)) = MessageEscape
        Next i
        
        Escape = data
    End Function
    
    Private Function Deescape(ByRef data() As Byte) As Byte()
        Dim removeAt() As Long
        Dim index As Long: index = -1
        Dim length As Long: length = UBound(data) - LBound(data) + 1
    
        For N = length - 1 To 0 Step -1
            If (data(N) = MessageDelimiter Or data(N) = MessageEscape) And N = 0 Then
                ReDim Preserve removeAt(index + 1)
                index = index + 1
                removeAt(index) = N
            ElseIf (data(N) = MessageDelimiter Or data(N) = MessageEscape) And data(N - 1) <> MessageEscape Then
                ReDim Preserve removeAt(index + 1)
                index = index + 1
                removeAt(index) = N
            End If
        Next N
        
        For i = 0 To index
            length = UBound(data) - LBound(data)
            For N = removeAt(i) To length - 1
                data(N) = data(N + 1)
            Next N
            ReDim Preserve data(length - 1)
        Next i
    
        Deescape = data
    End Function
    
    Private Sub SendResponse(ByVal socket As Winsock, ByVal response As Segment)
        Dim msg As New Message
        msg.version = 1
        msg.Item(0) = response
        socket.senddata EncodeMessage(msg)
    End Sub

    Private Function CreateResponse(ByVal sessionId As Long, ByVal Code As ResponseCode) As Segment
        Dim response As New Segment
        response.id = SegmentResponse
        response.Session = sessionId
        response.version = response.CurrentVersion
        response.sequence = 0
        response.Code = Code
        Set CreateResponse = response
    End Function
    
    Private Function IsValidId(ByVal id As Byte) As Boolean
        If id = SegmentHeartbeat Or id = SegmentSession Or id = SegmentPing Or id = SegmentVersion Or _
           id = SegmentResponse Or id = SegmentBypass Or id = SegmentClosing Then
            IsValidId = True
        End If
        IsValidId = False
    End Function
    
    Public Function ComputeHash(ByRef Bytes() As Byte, ByVal startIndex As Long, ByVal isReverse As Boolean) As Long
        Dim i As Long
        Dim index As Long
        Dim Size As Long
        Dim result As Long: result = &HFFFFFFFF
        Dim isNegative As Boolean: isNegative = False
        
        If Not isCrc32Initialized Then Crc32Init
        
        Size = UBound(Bytes)
        If isReverse Then
            For i = Size To startIndex Step -1
                index = (result And &HFF) Xor Bytes(i)
                If result < 0 Then
                    isNegative = True
                    result = result And &H7FFFFFFF
                End If
                result = (result \ &H100)
                If isNegative = True Then
                    result = result Or &H800000
                    isNegative = False
                End If
                result = result Xor crc32LookUp(index)
            Next i
        Else
            For i = startIndex To Size
                index = (result And &HFF) Xor Bytes(i)
                If result < 0 Then
                    isNegative = True
                    result = result And &H7FFFFFFF
                End If
                result = (result \ &H100)
                If isNegative = True Then
                    result = result Or &H800000
                    isNegative = False
                End If
                result = result Xor crc32LookUp(index)
            Next i
        End If
        
        ComputeHash = result Xor &HFFFFFFFF
    End Function
    
        Public Sub Crc32Init()
        If isCrc32Initialized Then Exit Sub
        ReDim crc32LookUp(255)
     
        crc32LookUp(0) = &H0
        crc32LookUp(1) = &H77073096
        crc32LookUp(2) = &HEE0E612C
        crc32LookUp(3) = &H990951BA
        crc32LookUp(4) = &H76DC419
        crc32LookUp(5) = &H706AF48F
        crc32LookUp(6) = &HE963A535
        crc32LookUp(7) = &H9E6495A3
        crc32LookUp(8) = &HEDB8832
        crc32LookUp(9) = &H79DCB8A4
        crc32LookUp(10) = &HE0D5E91E
        crc32LookUp(11) = &H97D2D988
        crc32LookUp(12) = &H9B64C2B
        crc32LookUp(13) = &H7EB17CBD
        crc32LookUp(14) = &HE7B82D07
        crc32LookUp(15) = &H90BF1D91
        crc32LookUp(16) = &H1DB71064
        crc32LookUp(17) = &H6AB020F2
        crc32LookUp(18) = &HF3B97148
        crc32LookUp(19) = &H84BE41DE
        crc32LookUp(20) = &H1ADAD47D
        crc32LookUp(21) = &H6DDDE4EB
        crc32LookUp(22) = &HF4D4B551
        crc32LookUp(23) = &H83D385C7
        crc32LookUp(24) = &H136C9856
        crc32LookUp(25) = &H646BA8C0
        crc32LookUp(26) = &HFD62F97A
        crc32LookUp(27) = &H8A65C9EC
        crc32LookUp(28) = &H14015C4F
        crc32LookUp(29) = &H63066CD9
        crc32LookUp(30) = &HFA0F3D63
        crc32LookUp(31) = &H8D080DF5
        crc32LookUp(32) = &H3B6E20C8
        crc32LookUp(33) = &H4C69105E
        crc32LookUp(34) = &HD56041E4
        crc32LookUp(35) = &HA2677172
        crc32LookUp(36) = &H3C03E4D1
        crc32LookUp(37) = &H4B04D447
        crc32LookUp(38) = &HD20D85FD
        crc32LookUp(39) = &HA50AB56B
        crc32LookUp(40) = &H35B5A8FA
        crc32LookUp(41) = &H42B2986C
        crc32LookUp(42) = &HDBBBC9D6
        crc32LookUp(43) = &HACBCF940
        crc32LookUp(44) = &H32D86CE3
        crc32LookUp(45) = &H45DF5C75
        crc32LookUp(46) = &HDCD60DCF
        crc32LookUp(47) = &HABD13D59
        crc32LookUp(48) = &H26D930AC
        crc32LookUp(49) = &H51DE003A
        crc32LookUp(50) = &HC8D75180
        crc32LookUp(51) = &HBFD06116
        crc32LookUp(52) = &H21B4F4B5
        crc32LookUp(53) = &H56B3C423
        crc32LookUp(54) = &HCFBA9599
        crc32LookUp(55) = &HB8BDA50F
        crc32LookUp(56) = &H2802B89E
        crc32LookUp(57) = &H5F058808
        crc32LookUp(58) = &HC60CD9B2
        crc32LookUp(59) = &HB10BE924
        crc32LookUp(60) = &H2F6F7C87
        crc32LookUp(61) = &H58684C11
        crc32LookUp(62) = &HC1611DAB
        crc32LookUp(63) = &HB6662D3D
        crc32LookUp(64) = &H76DC4190
        crc32LookUp(65) = &H1DB7106
        crc32LookUp(66) = &H98D220BC
        crc32LookUp(67) = &HEFD5102A
        crc32LookUp(68) = &H71B18589
        crc32LookUp(69) = &H6B6B51F
        crc32LookUp(70) = &H9FBFE4A5
        crc32LookUp(71) = &HE8B8D433
        crc32LookUp(72) = &H7807C9A2
        crc32LookUp(73) = &HF00F934
        crc32LookUp(74) = &H9609A88E
        crc32LookUp(75) = &HE10E9818
        crc32LookUp(76) = &H7F6A0DBB
        crc32LookUp(77) = &H86D3D2D
        crc32LookUp(78) = &H91646C97
        crc32LookUp(79) = &HE6635C01
        crc32LookUp(80) = &H6B6B51F4
        crc32LookUp(81) = &H1C6C6162
        crc32LookUp(82) = &H856530D8
        crc32LookUp(83) = &HF262004E
        crc32LookUp(84) = &H6C0695ED
        crc32LookUp(85) = &H1B01A57B
        crc32LookUp(86) = &H8208F4C1
        crc32LookUp(87) = &HF50FC457
        crc32LookUp(88) = &H65B0D9C6
        crc32LookUp(89) = &H12B7E950
        crc32LookUp(90) = &H8BBEB8EA
        crc32LookUp(91) = &HFCB9887C
        crc32LookUp(92) = &H62DD1DDF
        crc32LookUp(93) = &H15DA2D49
        crc32LookUp(94) = &H8CD37CF3
        crc32LookUp(95) = &HFBD44C65
        crc32LookUp(96) = &H4DB26158
        crc32LookUp(97) = &H3AB551CE
        crc32LookUp(98) = &HA3BC0074
        crc32LookUp(99) = &HD4BB30E2
        crc32LookUp(100) = &H4ADFA541
        crc32LookUp(101) = &H3DD895D7
        crc32LookUp(102) = &HA4D1C46D
        crc32LookUp(103) = &HD3D6F4FB
        crc32LookUp(104) = &H4369E96A
        crc32LookUp(105) = &H346ED9FC
        crc32LookUp(106) = &HAD678846
        crc32LookUp(107) = &HDA60B8D0
        crc32LookUp(108) = &H44042D73
        crc32LookUp(109) = &H33031DE5
        crc32LookUp(110) = &HAA0A4C5F
        crc32LookUp(111) = &HDD0D7CC9
        crc32LookUp(112) = &H5005713C
        crc32LookUp(113) = &H270241AA
        crc32LookUp(114) = &HBE0B1010
        crc32LookUp(115) = &HC90C2086
        crc32LookUp(116) = &H5768B525
        crc32LookUp(117) = &H206F85B3
        crc32LookUp(118) = &HB966D409
        crc32LookUp(119) = &HCE61E49F
        crc32LookUp(120) = &H5EDEF90E
        crc32LookUp(121) = &H29D9C998
        crc32LookUp(122) = &HB0D09822
        crc32LookUp(123) = &HC7D7A8B4
        crc32LookUp(124) = &H59B33D17
        crc32LookUp(125) = &H2EB40D81
        crc32LookUp(126) = &HB7BD5C3B
        crc32LookUp(127) = &HC0BA6CAD
        crc32LookUp(128) = &HEDB88320
        crc32LookUp(129) = &H9ABFB3B6
        crc32LookUp(130) = &H3B6E20C
        crc32LookUp(131) = &H74B1D29A
        crc32LookUp(132) = &HEAD54739
        crc32LookUp(133) = &H9DD277AF
        crc32LookUp(134) = &H4DB2615
        crc32LookUp(135) = &H73DC1683
        crc32LookUp(136) = &HE3630B12
        crc32LookUp(137) = &H94643B84
        crc32LookUp(138) = &HD6D6A3E
        crc32LookUp(139) = &H7A6A5AA8
        crc32LookUp(140) = &HE40ECF0B
        crc32LookUp(141) = &H9309FF9D
        crc32LookUp(142) = &HA00AE27
        crc32LookUp(143) = &H7D079EB1
        crc32LookUp(144) = &HF00F9344
        crc32LookUp(145) = &H8708A3D2
        crc32LookUp(146) = &H1E01F268
        crc32LookUp(147) = &H6906C2FE
        crc32LookUp(148) = &HF762575D
        crc32LookUp(149) = &H806567CB
        crc32LookUp(150) = &H196C3671
        crc32LookUp(151) = &H6E6B06E7
        crc32LookUp(152) = &HFED41B76
        crc32LookUp(153) = &H89D32BE0
        crc32LookUp(154) = &H10DA7A5A
        crc32LookUp(155) = &H67DD4ACC
        crc32LookUp(156) = &HF9B9DF6F
        crc32LookUp(157) = &H8EBEEFF9
        crc32LookUp(158) = &H17B7BE43
        crc32LookUp(159) = &H60B08ED5
        crc32LookUp(160) = &HD6D6A3E8
        crc32LookUp(161) = &HA1D1937E
        crc32LookUp(162) = &H38D8C2C4
        crc32LookUp(163) = &H4FDFF252
        crc32LookUp(164) = &HD1BB67F1
        crc32LookUp(165) = &HA6BC5767
        crc32LookUp(166) = &H3FB506DD
        crc32LookUp(167) = &H48B2364B
        crc32LookUp(168) = &HD80D2BDA
        crc32LookUp(169) = &HAF0A1B4C
        crc32LookUp(170) = &H36034AF6
        crc32LookUp(171) = &H41047A60
        crc32LookUp(172) = &HDF60EFC3
        crc32LookUp(173) = &HA867DF55
        crc32LookUp(174) = &H316E8EEF
        crc32LookUp(175) = &H4669BE79
        crc32LookUp(176) = &HCB61B38C
        crc32LookUp(177) = &HBC66831A
        crc32LookUp(178) = &H256FD2A0
        crc32LookUp(179) = &H5268E236
        crc32LookUp(180) = &HCC0C7795
        crc32LookUp(181) = &HBB0B4703
        crc32LookUp(182) = &H220216B9
        crc32LookUp(183) = &H5505262F
        crc32LookUp(184) = &HC5BA3BBE
        crc32LookUp(185) = &HB2BD0B28
        crc32LookUp(186) = &H2BB45A92
        crc32LookUp(187) = &H5CB36A04
        crc32LookUp(188) = &HC2D7FFA7
        crc32LookUp(189) = &HB5D0CF31
        crc32LookUp(190) = &H2CD99E8B
        crc32LookUp(191) = &H5BDEAE1D
        crc32LookUp(192) = &H9B64C2B0
        crc32LookUp(193) = &HEC63F226
        crc32LookUp(194) = &H756AA39C
        crc32LookUp(195) = &H26D930A
        crc32LookUp(196) = &H9C0906A9
        crc32LookUp(197) = &HEB0E363F
        crc32LookUp(198) = &H72076785
        crc32LookUp(199) = &H5005713
        crc32LookUp(200) = &H95BF4A82
        crc32LookUp(201) = &HE2B87A14
        crc32LookUp(202) = &H7BB12BAE
        crc32LookUp(203) = &HCB61B38
        crc32LookUp(204) = &H92D28E9B
        crc32LookUp(205) = &HE5D5BE0D
        crc32LookUp(206) = &H7CDCEFB7
        crc32LookUp(207) = &HBDBDF21
        crc32LookUp(208) = &H86D3D2D4
        crc32LookUp(209) = &HF1D4E242
        crc32LookUp(210) = &H68DDB3F8
        crc32LookUp(211) = &H1FDA836E
        crc32LookUp(212) = &H81BE16CD
        crc32LookUp(213) = &HF6B9265B
        crc32LookUp(214) = &H6FB077E1
        crc32LookUp(215) = &H18B74777
        crc32LookUp(216) = &H88085AE6
        crc32LookUp(217) = &HFF0F6A70
        crc32LookUp(218) = &H66063BCA
        crc32LookUp(219) = &H11010B5C
        crc32LookUp(220) = &H8F659EFF
        crc32LookUp(221) = &HF862AE69
        crc32LookUp(222) = &H616BFFD3
        crc32LookUp(223) = &H166CCF45
        crc32LookUp(224) = &HA00AE278
        crc32LookUp(225) = &HD70DD2EE
        crc32LookUp(226) = &H4E048354
        crc32LookUp(227) = &H3903B3C2
        crc32LookUp(228) = &HA7672661
        crc32LookUp(229) = &HD06016F7
        crc32LookUp(230) = &H4969474D
        crc32LookUp(231) = &H3E6E77DB
        crc32LookUp(232) = &HAED16A4A
        crc32LookUp(233) = &HD9D65ADC
        crc32LookUp(234) = &H40DF0B66
        crc32LookUp(235) = &H37D83BF0
        crc32LookUp(236) = &HA9BCAE53
        crc32LookUp(237) = &HDEBB9EC5
        crc32LookUp(238) = &H47B2CF7F
        crc32LookUp(239) = &H30B5FFE9
        crc32LookUp(240) = &HBDBDF21C
        crc32LookUp(241) = &HCABAC28A
        crc32LookUp(242) = &H53B39330
        crc32LookUp(243) = &H24B4A3A6
        crc32LookUp(244) = &HBAD03605
        crc32LookUp(245) = &HCDD70693
        crc32LookUp(246) = &H54DE5729
        crc32LookUp(247) = &H23D967BF
        crc32LookUp(248) = &HB3667A2E
        crc32LookUp(249) = &HC4614AB8
        crc32LookUp(250) = &H5D681B02
        crc32LookUp(251) = &H2A6F2B94
        crc32LookUp(252) = &HB40BBE37
        crc32LookUp(253) = &HC30C8EA1
        crc32LookUp(254) = &H5A05DF1B
        crc32LookUp(255) = &H2D02EF8D

        isCrc32Initialized = True
    End Sub

