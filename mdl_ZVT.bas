Attribute VB_Name = "mdl_ZVT"
Option Explicit
Public Sub lese_ZVT_opt()
    On Error GoTo LOKAL_ERROR
    
    Dim rsrs As Recordset
    
    gZVTclientId = ""
    gZVTioPfad = ""
    giZVT_Fehler = 0
    gZVTDruckVar = "1"
    gZVTPName = ""
    gZVTPTitel = ""
    gZVTTimeOut = "120"
    
    If NewTableSuchenDBKombi("ZVTOPT", gdApp) Then

        Set rsrs = gdApp.OpenRecordset("select * from ZVTOPT")
        If Not rsrs.EOF Then
            
            If Not IsNull(rsrs!clientID) Then
                gZVTclientId = rsrs!clientID
            End If
            
            If Not IsNull(rsrs!ioPfad) Then
                gZVTioPfad = rsrs!ioPfad
            End If
            
            If Not IsNull(rsrs!DRUCKVAR) Then
                gZVTDruckVar = rsrs!DRUCKVAR
            End If
            
            If Not IsNull(rsrs!pname) Then
                gZVTPName = rsrs!pname
            End If
            
            If Not IsNull(rsrs!ptitel) Then
                gZVTPTitel = rsrs!ptitel
            End If
            
            If Not IsNull(rsrs!TimeOut) Then
                gZVTTimeOut = rsrs!TimeOut
            End If
        End If
        rsrs.Close: Set rsrs = Nothing
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdl_ZVT"
    Fehler.gsFunktion = "lese_ZVT_opt"
    Fehler.gsFehlertext = "Im Programmteil ZVT ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Function Lebst_Du_noch_ZVT() As Boolean
    On Error GoTo LOKAL_ERROR
    
    
    Lebst_Du_noch_ZVT = False
    
    
    Dim iFileNr             As Integer
    Dim lPos                As Long
    Dim cSatz               As String
    
    cSatz = "Lebst du noch?"
    
    'vorher Outfile löschen
    Kill gZVTioPfad & "\ping.ck"
    Kill gZVTioPfad & "\pong.ck"
    
    iFileNr = FreeFile
    Open gZVTioPfad & "\ping.ck" For Binary As #iFileNr
    
    lPos = LOF(iFileNr)
    lPos = lPos + 1
    Put #iFileNr, lPos, cSatz
        
    Close iFileNr
    
    
    PauseSi 2

    If FileExists(gZVTioPfad & "\pong.ck") Then
        Lebst_Du_noch_ZVT = True
    End If
    
    
Exit Function
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "mdl_ZVT"
        Fehler.gsFunktion = "Lebst_Du_noch_ZVT"
        Fehler.gsFehlertext = "Im Programmteil ZVT ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Function
Public Function Zahlung_ZVT(sBetrag As String) As String
    On Error GoTo LOKAL_ERROR

    Zahlung_ZVT = ""
    
    Create_Infile_ZVT gZVTioPfad, sBetrag, gZVTclientId
    
    Outfile_suchen_ZVT gZVTioPfad, gZVTclientId, gZVTDruckVar, gZVTTimeOut
        
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdl_ZVT"
    Fehler.gsFunktion = "Zahlung_ZVT"
    Fehler.gsFehlertext = "Im Programmteil ZVT ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Sub Create_Infile_ZVT(sPfad As String, sCentBetrag As String, sClient As String)
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr             As Integer
    Dim sPfad_und_Datei     As String
    Dim lPos                As Long
    Dim cSatz               As String
    
'''    sPfad = "C:\Test"
    
    cSatz = sCentBetrag & vbCrLf
    cSatz = cSatz & "00" & vbCrLf
    cSatz = cSatz & "0" & vbCrLf '0 = Auto
    
    'vorher Outfile löschen
    Kill sPfad & "\outfile." & sClient
    
    sPfad_und_Datei = sPfad & "\infile." & sClient
    
    iFileNr = FreeFile
    Open sPfad_und_Datei For Binary As #iFileNr
    
    lPos = LOF(iFileNr)
    lPos = lPos + 1
    Put #iFileNr, lPos, cSatz
        
    Close iFileNr

Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "mdl_ZVT"
        Fehler.gsFunktion = "Create_Infile_ZVT"
        Fehler.gsFehlertext = "Im Programmteil ZVT ist ein Fehler aufgetreten."
    End If
    
    Fehlermeldung1
End Sub
Public Sub Create_Infile_ZVT_Kassenschnitt(sPfad As String, sClient As String)
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr             As Integer
    Dim sPfad_und_Datei     As String
    Dim lPos                As Long
    Dim cSatz               As String
    
    cSatz = "0" & vbCrLf
    cSatz = cSatz & "02" & vbCrLf
    cSatz = cSatz & "0" & vbCrLf
    
    'vorher Outfile löschen
    Kill sPfad & "\outfile." & sClient
    
    sPfad_und_Datei = sPfad & "\infile." & sClient
    
    iFileNr = FreeFile
    Open sPfad_und_Datei For Binary As #iFileNr
    
    lPos = LOF(iFileNr)
    lPos = lPos + 1
    Put #iFileNr, lPos, cSatz
        
    Close iFileNr

Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "mdl_ZVT"
        Fehler.gsFunktion = "Create_Infile_ZVT_Kassenschnitt"
        Fehler.gsFehlertext = "Im Programmteil ZVT ist ein Fehler aufgetreten."
    End If
    
    Fehlermeldung1
End Sub
Public Sub Create_Infile_ZVT_BelegWiederholung(sPfad As String, sClient As String)
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr             As Integer
    Dim sPfad_und_Datei     As String
    Dim lPos                As Long
    Dim cSatz               As String
    
    cSatz = "0" & vbCrLf
    cSatz = cSatz & "02" & vbCrLf
    cSatz = cSatz & "0" & vbCrLf
    
    'vorher Outfile löschen
    Kill sPfad & "\outfile." & sClient
    
    sPfad_und_Datei = sPfad & "\infile." & sClient
    
    iFileNr = FreeFile
    Open sPfad_und_Datei For Binary As #iFileNr
    
    lPos = LOF(iFileNr)
    lPos = lPos + 1
    Put #iFileNr, lPos, cSatz
        
    Close iFileNr

Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "mdl_ZVT"
        Fehler.gsFunktion = "Create_Infile_ZVT_BelegWiederholung"
        Fehler.gsFehlertext = "Im Programmteil ZVT ist ein Fehler aufgetreten."
    End If
    
    Fehlermeldung1
End Sub
Public Sub Outfile_suchen_ZVT(sZVT_Pfad As String, sClient As String, sDruckVar As String, lTimeOut As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim iErr_Zaehler As Integer
    Dim iTimeout As Integer
    Dim iRet As Integer
    Dim bTimeout_erreicht As Boolean
    Dim ctmp As String
    
    bTimeout_erreicht = False
    
    iTimeout = lTimeOut * 10
    
    iErr_Zaehler = 0
    Do While FileExists(sZVT_Pfad & "\outfile." & sClient) = False
        PauseSi 0.1
        iErr_Zaehler = iErr_Zaehler + 1
        If iErr_Zaehler > iTimeout Then
            bTimeout_erreicht = True
            Exit Do
        End If
    Loop
    
    If bTimeout_erreicht = True Then
        ctmp = "Ist die Kartenzahlung erfolgt?" & vbCrLf & vbCrLf
        ctmp = ctmp & "Bitte auf das Kartenterminal schauen!" & vbCrLf
        
        iRet = MsgBox(ctmp, vbQuestion + vbYesNo, "Kartenterminal?")
        If iRet = vbNo Then
            giZVT_Fehler = 1
            Exit Sub
        End If
    End If
    
    PauseSi 0.1
    If FileExists(sZVT_Pfad & "\outfile." & sClient) Then
        Outfile_nachFehler_checken_ZVT sZVT_Pfad, sClient
        
        Select Case sDruckVar
        
            Case "1"
                Outfile_drucken_ZVT_V1 sZVT_Pfad, sClient
            Case "2"
                Outfile_drucken_ZVT_V2 sZVT_Pfad, sClient
        End Select
    Else
        giZVT_Fehler = 1
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdl_ZVT"
    Fehler.gsFunktion = "Outfile_suchen_ZVT"
    Fehler.gsFehlertext = "Im Programmteil ZVT ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub Outfile_nachFehler_checken_ZVT(sZVT_Pfad As String, sClient As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cSatz                   As String
    Dim iFileNr                 As Integer
    Dim lPos                    As Long
    Dim cEinzelsatz             As String
    
    giZVT_Fehler = 0
    lPos = 1
    
    iFileNr = FreeFile
    Open sZVT_Pfad & "\outfile." & sClient For Binary As #iFileNr
    If LOF(iFileNr) > 0 Then
    
        cSatz = Space$(LOF(iFileNr))
        Get #iFileNr, 1, cSatz
        
        If InStr(UCase(cSatz), "LIZENZ FEHLERHAFT") Then
            giZVT_Fehler = 99
            Exit Sub
        End If
        
        lPos = InStr(lPos, cSatz, "@@@@@")
        
        If lPos > 1 Then
        
            
            'auslesen
            cEinzelsatz = Mid(cSatz, lPos + 7, 4)
        End If
        
        If Val(cEinzelsatz) > 0 Then
           giZVT_Fehler = Val(cEinzelsatz)
        Else
            giZVT_Fehler = 0
        End If
        

    End If
    
    Close iFileNr
        
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdl_ZVT"
    Fehler.gsFunktion = "Outfile_nachFehler_checken_ZVT"
    Fehler.gsFehlertext = "Im Programmteil ZVT ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub Outfile_drucken_ZVT_V1(sZVT_Pfad As String, sClient As String)
    On Error GoTo LOKAL_ERROR

    ReDim cDruckZeile(1 To 1) As String
    ReDim cDruckFormat(1 To 1) As String
    Dim cSatz                   As String
    Dim iFileNr                 As Integer
    Dim lLenfil                 As Long
    Dim lPos                    As Long
    Dim lPosEnde                As Long
    Dim cEinzelsatz             As String
    Dim cWert                   As String
    Dim cWert_Format            As String
    Dim iDruckzeilen_count      As Integer
    Dim lAnzBuchstaben          As Long
    Dim iüAnzahl                As Integer

    lPos = 1
    lPosEnde = 1
    lAnzBuchstaben = 1
    
    iüAnzahl = 0

    iDruckzeilen_count = 0
    ReDim cDruckZeile(1 To 1) As String
    ReDim cDruckFormat(1 To 1) As String

    iFileNr = FreeFile
    Open sZVT_Pfad & "\outfile." & sClient For Binary As #iFileNr
    If LOF(iFileNr) > 0 Then

        cSatz = Space$(LOF(iFileNr))
        Get #iFileNr, 1, cSatz

        lLenfil = Len(cSatz)

        Do
            lPosEnde = InStr(lPos, cSatz, vbCrLf)

            lAnzBuchstaben = lPosEnde - lPos
            cEinzelsatz = Mid(cSatz, lPos, lAnzBuchstaben)

            cEinzelsatz = SwapStr(cEinzelsatz, "Ã¤", "ä")
            cEinzelsatz = SwapStr(cEinzelsatz, "Ã¼", "ü")
            cEinzelsatz = SwapStr(cEinzelsatz, "Ã", "Ä")
            cEinzelsatz = SwapStr(cEinzelsatz, Chr(132), "")
            
            
            
            cEinzelsatz = Trim(cEinzelsatz)

            If Left(cEinzelsatz, 1) = "ü" Then
            
                
            
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String


                cWert_Format = ""
                cWert_Format = "S"


                cWert = ""
                cWert = cEinzelsatz

                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
                
                iüAnzahl = iüAnzahl + 1
                
                If iüAnzahl = 2 Then
                    GoTo ENDE
                End If
                
            ElseIf Left(cEinzelsatz, 7) = "Kassen-" Then
            
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String

                cWert_Format = ""
                cWert_Format = "Z"

                cWert = ""
                cWert = cEinzelsatz

                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
                
                iüAnzahl = iüAnzahl + 1 'hier das 2.ü vorgauckeln
                
            ElseIf Left(cEinzelsatz, 3) = "EUR" Then
            
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String

                cWert_Format = ""
                cWert_Format = "F"

                cWert = ""
                cWert = cEinzelsatz

                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
                
            ElseIf Left(cEinzelsatz, 12) = "Unterschrift" Then
            
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String

                cWert_Format = ""
                cWert_Format = "F"

                cWert = ""
                cWert = cEinzelsatz

                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
                
            ElseIf Left(cEinzelsatz, 8) = "umseitig" Then
            
                cEinzelsatz = "erforderlich"
            
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String

                cWert_Format = ""
                cWert_Format = "F"

                cWert = ""
                cWert = cEinzelsatz

                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
                
                'Jetzt hängen wir die Datenschutzerklärung an:
                
                cWert_Format = "K"
                
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String

                cWert = ""

                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
                
            
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String

                cWert = "Ermächtigung Lastschrifteinzug:"

                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
                
                
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String

                cWert = ""

                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
                
                
                
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String

                cWert = "Ich ermächtige hiermit das"

                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
            
               
    
                
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String

                cWert = "oben genannte Unternehmen,"

                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
                
    
                
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String
                cWert = "den als Endsumme ausgewiesenen"
                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
                
                
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String
                cWert = "Betrag von meinem durch BIC"
                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
    
    
                
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String
                cWert = "und IBAN bezeichneten Konto"
                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
    
                
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String
                cWert = "durch Lastschrift einzuziehen."
                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
                
                
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String
                cWert = ""
                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
    
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String
                cWert = "Ermächtigung Adressweitergabe:"
                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format

               
    
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String
                cWert = ""
                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
                
                
                
                
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String
                cWert = "Ich weise mein Kreditinstitut,"
                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
    
        
                
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String
                cWert = "das durch die BIC "
                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
    
                
    
    
    
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String
                cWert = "bezeichnet ist, unwiderruflich"
                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
                
                
                
                
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String
                cWert = "an, bei Nichteinlösung der Last-"
                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
                
                
                
                
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String
                cWert = "schrift oder bei Widerspruch"
                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
                
                
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String
                cWert = "gegen die Lastschrift dem"
                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
                
                
                
                
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String
                cWert = "Unternehmen oder einem von"
                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
                
                
                
                
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String
                cWert = "ihm beauftragten Dritten auf "
                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
                
                
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String
                cWert = "Anforderung meinen Namen und"
                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
                
                
                
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String
                cWert = "meine Adresse mitzuteilen, damit"
                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
    
                
                
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String
                cWert = "das Unternehmen seinen Anspruch"
                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
                
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String
                cWert = "gegen mich geltend machen kann."
                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
                
                
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String
                cWert = ""
                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
           
                
           
            
                
                
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String
                cWert = ""
                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
                
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String
                cWert = ""
                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
                
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String
                cWert = ""
                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
                
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String
                cWert = ""
                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
                
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String
                cWert = "_____________________"
                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
                
                
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String
                cWert = "Unterschrift:"
                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
                
                
                
                
                'Ende Jetzt hängen wir in die Datenschutzerklärung an:
                
                
                
                
            Else
            
                
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String


                cWert_Format = ""
                cWert_Format = "Z"


                cWert = ""
                cWert = cEinzelsatz

                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
                
                
                

            End If

            lPos = lPos + lAnzBuchstaben + 2

        Loop While lLenfil >= lPos

    End If
    
    
    If giZVT_Fehler > 0 Then
    
        iDruckzeilen_count = iDruckzeilen_count + 1
        ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
        ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String


        cWert_Format = ""
        cWert_Format = "S"


        cWert = ""
        cWert = cEinzelsatz

        cDruckZeile(iDruckzeilen_count) = cWert
        cDruckFormat(iDruckzeilen_count) = cWert_Format
    
    End If
    
ENDE:
    
    Close iFileNr
    
    DruckeEndlosBeleg_ZVT cDruckZeile(), cDruckFormat(), iDruckzeilen_count, True, 18


Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdl_ZVT"
    Fehler.gsFunktion = "Outfile_drucken_ZVT_V1"
    Fehler.gsFehlertext = "Im Programmteil ZVT ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Public Sub Outfile_drucken_ZVT_V2(sZVT_Pfad As String, sClient As String)
    On Error GoTo LOKAL_ERROR

    ReDim cDruckZeile(1 To 1) As String
    ReDim cDruckFormat(1 To 1) As String
    Dim cSatz                   As String
    Dim iFileNr                 As Integer
    Dim lLenfil                 As Long
    Dim lPos                    As Long
    Dim lPosEnde                As Long
    Dim cEinzelsatz             As String
    Dim cWert                   As String
    Dim cWert_Format            As String
    Dim iDruckzeilen_count      As Integer
    Dim lAnzBuchstaben          As Long
    Dim iüAnzahl                As Integer

    lPos = 1
    lPosEnde = 1
    lAnzBuchstaben = 1
    
    iüAnzahl = 0

    iDruckzeilen_count = 0
    ReDim cDruckZeile(1 To 1) As String
    ReDim cDruckFormat(1 To 1) As String

    iFileNr = FreeFile
    Open sZVT_Pfad & "\outfile." & sClient For Binary As #iFileNr
    If LOF(iFileNr) > 0 Then

        cSatz = Space$(LOF(iFileNr))
        Get #iFileNr, 1, cSatz

        lLenfil = Len(cSatz)

        Do
            lPosEnde = InStr(lPos, cSatz, vbCrLf)

            lAnzBuchstaben = lPosEnde - lPos
            cEinzelsatz = Mid(cSatz, lPos, lAnzBuchstaben)

            cEinzelsatz = SwapStr(cEinzelsatz, "Ã¤", "ä")
            cEinzelsatz = SwapStr(cEinzelsatz, "Ã¼", "ü")
            cEinzelsatz = SwapStr(cEinzelsatz, "Ã", "Ä")
            cEinzelsatz = SwapStr(cEinzelsatz, Chr(132), "")
            cEinzelsatz = SwapStr(cEinzelsatz, "PBetrag ", "Betrag ")
            cEinzelsatz = SwapStr(cEinzelsatz, "PZahlung ", "Zahlung ")
            cEinzelsatz = SwapStr(cEinzelsatz, "PAutorisierung ", "Autorisierung ")
            
            'besonderheit bei Unger alle Leerzeilen enthalten dieses Zeichen
            cEinzelsatz = SwapStr(cEinzelsatz, "Â", "")
            
            'besonderheit bei Unger einige Zeilen haben ein @ vorangstellt
            'ich nehme @ weg, so ist die Auswertung der Fehlermeldung @@@@@ passé
            cEinzelsatz = SwapStr(cEinzelsatz, "@", "")
            
            cEinzelsatz = Trim(cEinzelsatz)

            If Left(cEinzelsatz, 1) = "ü" Then
            
                
            
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String


                cWert_Format = ""
                cWert_Format = "S"


                cWert = ""
                cWert = cEinzelsatz

                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
                
                iüAnzahl = iüAnzahl + 1
                
                If iüAnzahl = 2 Then
                    GoTo ENDE
                End If
                
            ElseIf Left(cEinzelsatz, 7) = "Kassen-" Then
            
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String

                cWert_Format = ""
                cWert_Format = "Z"

                cWert = ""
                cWert = cEinzelsatz

                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
                
                iüAnzahl = iüAnzahl + 1 'hier das 2.ü vorgauckeln
                
            ElseIf InStr(UCase(cEinzelsatz), "STORNO EUR") > 0 Then
            
                Dim lPosEur As Long
                lPosEur = InStr(UCase(cEinzelsatz), "EUR")
                cEinzelsatz = "Storno EUR " & Trim(Right(cEinzelsatz, Len(cEinzelsatz) - lPosEur - 3))
            
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String

                cWert_Format = ""
                cWert_Format = "F"

                cWert = ""
                cWert = cEinzelsatz

                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
                
            ElseIf InStr(UCase(cEinzelsatz), "BETRAG EUR") > 0 Then
            
'                Dim lPosEur As Long
                lPosEur = InStr(UCase(cEinzelsatz), "EUR")
                cEinzelsatz = "Betrag EUR " & Trim(Right(cEinzelsatz, Len(cEinzelsatz) - lPosEur - 3))
            
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String

                cWert_Format = ""
                cWert_Format = "F"

                cWert = ""
                cWert = cEinzelsatz

                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
                
            ElseIf Left(cEinzelsatz, 12) = "Unterschrift" Then
            
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String

                cWert_Format = ""
                cWert_Format = "F"

                cWert = ""
                cWert = cEinzelsatz

                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
                
            ElseIf Left(cEinzelsatz, 8) = "umseitig" Then
            
                cEinzelsatz = "erforderlich"
            
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String

                cWert_Format = ""
                cWert_Format = "F"

                cWert = ""
                cWert = cEinzelsatz

                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
                
                'Jetzt hängen wir die Datenschutzerklärung an:
                
                cWert_Format = "K"
                
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String

                cWert = ""

                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
                
            
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String

                cWert = "Ermächtigung Lastschrifteinzug:"

                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
                
                
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String

                cWert = ""

                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
                
                
                
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String

                cWert = "Ich ermächtige hiermit das"

                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
            
               
    
                
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String

                cWert = "oben genannte Unternehmen,"

                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
                
    
                
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String
                cWert = "den als Endsumme ausgewiesenen"
                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
                
                
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String
                cWert = "Betrag von meinem durch BIC"
                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
    
    
                
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String
                cWert = "und IBAN bezeichneten Konto"
                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
    
                
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String
                cWert = "durch Lastschrift einzuziehen."
                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
                
                
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String
                cWert = ""
                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
    
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String
                cWert = "Ermächtigung Adressweitergabe:"
                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format

               
    
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String
                cWert = ""
                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
                
                
                
                
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String
                cWert = "Ich weise mein Kreditinstitut,"
                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
    
        
                
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String
                cWert = "das durch die BIC "
                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
    
                
    
    
    
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String
                cWert = "bezeichnet ist, unwiderruflich"
                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
                
                
                
                
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String
                cWert = "an, bei Nichteinlösung der Last-"
                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
                
                
                
                
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String
                cWert = "schrift oder bei Widerspruch"
                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
                
                
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String
                cWert = "gegen die Lastschrift dem"
                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
                
                
                
                
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String
                cWert = "Unternehmen oder einem von"
                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
                
                
                
                
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String
                cWert = "ihm beauftragten Dritten auf "
                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
                
                
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String
                cWert = "Anforderung meinen Namen und"
                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
                
                
                
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String
                cWert = "meine Adresse mitzuteilen, damit"
                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
    
                
                
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String
                cWert = "das Unternehmen seinen Anspruch"
                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
                
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String
                cWert = "gegen mich geltend machen kann."
                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
                
                
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String
                cWert = ""
                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
           
                
           
            
                
                
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String
                cWert = ""
                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
                
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String
                cWert = ""
                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
                
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String
                cWert = ""
                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
                
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String
                cWert = ""
                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
                
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String
                cWert = "_____________________"
                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
                
                
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String
                cWert = "Unterschrift:"
                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
                
                
                
                
                'Ende Jetzt hängen wir in die Datenschutzerklärung an:
                
                
                
                
            Else
            
                
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String


                cWert_Format = ""
                cWert_Format = "Z"


                cWert = ""
                cWert = cEinzelsatz

                cDruckZeile(iDruckzeilen_count) = cWert
                cDruckFormat(iDruckzeilen_count) = cWert_Format
                
                
                

            End If

            lPos = lPos + lAnzBuchstaben + 2

        Loop While lLenfil >= lPos

    End If
    
    
    If giZVT_Fehler > 0 Then
    
        iDruckzeilen_count = iDruckzeilen_count + 1
        ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
        ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String


        cWert_Format = ""
        cWert_Format = "S"


        cWert = ""
        cWert = cEinzelsatz

        cDruckZeile(iDruckzeilen_count) = cWert
        cDruckFormat(iDruckzeilen_count) = cWert_Format
    
    End If
    
ENDE:
    
    Close iFileNr
    
    DruckeEndlosBeleg_ZVT cDruckZeile(), cDruckFormat(), iDruckzeilen_count, True, 12


Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdl_ZVT"
    Fehler.gsFunktion = "Outfile_drucken_ZVT_V2"
    Fehler.gsFehlertext = "Im Programmteil ZVT ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Public Sub DruckeEndlosBeleg_ZVT(cZeilen() As String, cFormat() As String, iMax As Integer, bschnitt As Boolean, iFettgr As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim aDeviceName As String
    Dim cEscapeSequenz As String
    ReDim cDruckZeile(1 To 1) As String
    ReDim cDruckFormat(1 To 1) As String
    Dim lAnzZeile As Long
    Dim lcount As Long
    Dim i As Integer
    
    'zum Test
    setzedrucker gcBonDrucker
    
    'Drucker an, Display aus, Init Drucker
    
    
    aDeviceName = Printer.DeviceName
'''    cEscapeSequenz = gcInit
'''    OpenDrawer aDeviceName, cEscapeSequenz
'''
'''    '***********************************************
'''    'Drucker ein- und Kundendisplay ausschalten
'''    '***********************************************
'''
'''    cEscapeSequenz = gcInit
    lAnzZeile = lAnzZeile + 1
'''    ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
'''    cDruckZeile(lAnzZeile) = cEscapeSequenz
    
    ReDim Preserve cDruckFormat(1 To lAnzZeile) As String
    cDruckFormat(lAnzZeile) = ""

    For i = 1 To iMax
        lAnzZeile = lAnzZeile + 1
        ReDim Preserve cDruckZeile(1 To lAnzZeile) As String
        ReDim Preserve cDruckFormat(1 To lAnzZeile) As String
        
        If InStr(1, cFormat(i), "Z") > 0 Then
            cZeilen(i) = Space$((26 - Len(cZeilen(i))) / 2) & cZeilen(i)
        End If
        
        
'        KonvertAnsiAscii cZeilen(i)
        cDruckZeile(lAnzZeile) = cZeilen(i) & vbCrLf
        cDruckFormat(lAnzZeile) = cFormat(i)
    Next i


    
    OpenDrawer4_ZVT aDeviceName, cDruckZeile(), cDruckFormat(), lAnzZeile, bschnitt, iFettgr



Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdl_ZVT"
    Fehler.gsFunktion = "DruckeEndlosBeleg_ZVT"
    Fehler.gsFehlertext = "Beim Drucken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
End Sub
Public Sub OpenDrawer4_ZVT(aDeviceName As String, cDruckZeile() As String, cDruckFormat() As String, lAnzZeile As Long, bschnitt As Boolean, iFettgr As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    Dim ctmp As String
    
    Dim cEscapeSequenz As String
    
    If gbNOBONDRUCKER = True Then
        Exit Sub
    End If
    
    'Einstellungen für Drucker vornehmen
    
    Printer.FontName = "Lucida Console"
    Printer.FontSize = 9


    For lcount = 1 To lAnzZeile
        ctmp = Space(4) & cDruckZeile(lcount)
''        KonvertAsciiAnsi ctmp
        If Right(ctmp, 2) = vbCrLf Then
            ctmp = Left(ctmp, Len(ctmp) - 2)
        End If
        
        If InStr(1, cDruckFormat(lcount), "F") > 0 Then
            Printer.FontSize = iFettgr
            Printer.FontBold = True
            Printer.Print ctmp
            
        ElseIf InStr(1, cDruckFormat(lcount), "K") > 0 Then
            Printer.FontSize = 6
            Printer.FontBold = True
            Printer.Print ctmp
        ElseIf InStr(1, cDruckFormat(lcount), "S") > 0 Then
            Printer.EndDoc
            
            'Nur Krakau
            If gbAPI = True And bschnitt = True Then
                aDeviceName = Printer.DeviceName
                cEscapeSequenz = gcSchneiden
                OpenDrawer aDeviceName, cEscapeSequenz
            End If
            
        ElseIf InStr(1, cDruckFormat(lcount), "N") > 0 Then
            
        Else
            Printer.FontBold = False
            Printer.FontSize = 9
            Printer.Print ctmp
        End If
        
    Next lcount
    

    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdl_ZVT"
    Fehler.gsFunktion = "OpenDrawer4_ZVT"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub OpenDrawer5_V3(aDeviceName As String, cDruckZeile() As String, lAnzZeile As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    Dim ctmp As String
    
    Dim cEscapeSequenz As String
    
    If gbNOBONDRUCKER = True Then
        Exit Sub
    End If
    
    'Einstellungen für Drucker vornehmen
    
'    Printer.FontName = "Lucida Console"
'    Printer.FontSize = 9
    
'    Printer.FontName = "Courier New"
'        Printer.FontSize = 10
'        Printer.FontBold = True
        
        
        Printer.FontName = "15 CPI"
    Printer.FontSize = 9.5
    Printer.FontBold = False


    For lcount = 1 To lAnzZeile
        ctmp = Space(4) & cDruckZeile(lcount)

        If Right(ctmp, 2) = vbCrLf Then
            ctmp = Left(ctmp, Len(ctmp) - 2)
        End If
        
        If lcount = 1 Then
        
'            Printer.FontSize = 12
            Printer.FontBold = True
            Printer.Print ctmp
            
        Else
            Printer.FontBold = False
            Printer.FontSize = 10
            Printer.Print ctmp
        End If
        
    Next lcount
    
    Printer.Print
    Printer.Print
    Printer.Print
    
    Printer.EndDoc
    

    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdl_ZVT"
    Fehler.gsFunktion = "OpenDrawer5_V3"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Function BelegWiederholung_ZVT() As String
    On Error GoTo LOKAL_ERROR

    BelegWiederholung_ZVT = ""
    
    Create_Infile_ZVT_BelegWiederholung gZVTioPfad, gZVTclientId
    
    Outfile_suchen_ZVT gZVTioPfad, gZVTclientId, gZVTDruckVar, gZVTTimeOut
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdl_ZVT"
    Fehler.gsFunktion = "BelegWiederholung_ZVT"
    Fehler.gsFehlertext = "Im Programmteil ZVT ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function Kassenschnitt_ZVT() As String
    On Error GoTo LOKAL_ERROR

    Kassenschnitt_ZVT = ""
    
    Create_Infile_ZVT_Kassenschnitt gZVTioPfad, gZVTclientId
    
    Outfile_suchen_ZVT gZVTioPfad, gZVTclientId, gZVTDruckVar, gZVTTimeOut
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdl_ZVT"
    Fehler.gsFunktion = "Kassenschnitt_ZVT"
    Fehler.gsFehlertext = "Im Programmteil ZVT ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function Storno_ZVT(sBeNR As String) As String
    On Error GoTo LOKAL_ERROR

    Storno_ZVT = ""
    
    Create_Infile_ZVT_Storno gZVTioPfad, gZVTclientId, sBeNR
    
    Outfile_suchen_ZVT gZVTioPfad, gZVTclientId, gZVTDruckVar, gZVTTimeOut
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdl_ZVT"
    Fehler.gsFunktion = "Storno_ZVT"
    Fehler.gsFehlertext = "Im Programmteil ZVT ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Sub Create_Infile_ZVT_Storno(sPfad As String, sClient As String, sBelegNummer As String)
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr             As Integer
    Dim sPfad_und_Datei     As String
    Dim lPos                As Long
    Dim cSatz               As String
    
    cSatz = sBelegNummer & vbCrLf
    cSatz = cSatz & "01" & vbCrLf
    
    
    'vorher Outfile löschen
    Kill sPfad & "\outfile." & sClient
    
    sPfad_und_Datei = sPfad & "\infile." & sClient
    
    iFileNr = FreeFile
    Open sPfad_und_Datei For Binary As #iFileNr
    
    lPos = LOF(iFileNr)
    lPos = lPos + 1
    Put #iFileNr, lPos, cSatz
        
    Close iFileNr

Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "mdl_ZVT"
        Fehler.gsFunktion = "Create_Infile_ZVT_Storno"
        Fehler.gsFehlertext = "Im Programmteil ZVT ist ein Fehler aufgetreten."
    End If
    
    Fehlermeldung1
End Sub







