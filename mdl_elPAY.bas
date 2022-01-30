Attribute VB_Name = "mdl_elPAY"
Option Explicit
Public Function Zahlung_elPAY(sBetrag As String) As String
    On Error GoTo LOKAL_ERROR

    Zahlung_elPAY = ""
    
    Create_Infile_ELPAY gELPioPfad, sBetrag, gELPclientId
    
    Outfile_suchen_ELPAY gELPioPfad, gELPclientId
        
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdl_elPAY"
    Fehler.gsFunktion = "Zahlung_elPAY"
    Fehler.gsFehlertext = "Im Programmteil elPAY ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function Zahlung_elPAY_Manuell(sBetrag As String, sPan As String, sVerfall As String) As String
    On Error GoTo LOKAL_ERROR

    Zahlung_elPAY_Manuell = ""
    
    Create_Infile_ELPAY_Manuell gELPioPfad, sBetrag, gELPclientId, sPan, sVerfall
    
    Outfile_suchen_ELPAY gELPioPfad, gELPclientId
        
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdl_elPAY"
    Fehler.gsFunktion = "Zahlung_elPAY_Manuell"
    Fehler.gsFehlertext = "Im Programmteil elPAY ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function Storno_elPAY_Manuell(sTANR As String, sBetrag As String, sPan As String, sVerfall As String) As String
    On Error GoTo LOKAL_ERROR

    Storno_elPAY_Manuell = ""
    
    Create_Infile_ELPAY_Manuell_Storno gELPioPfad, sBetrag, gELPclientId, sPan, sVerfall, sTANR
    
    Outfile_suchen_ELPAY gELPioPfad, gELPclientId
        
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdl_elPAY"
    Fehler.gsFunktion = "Storno_elPAY_Manuell"
    Fehler.gsFehlertext = "Im Programmteil elPAY ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function Storno_elPAY(sTANR As String, sBetrag As String) As String
    On Error GoTo LOKAL_ERROR

    Storno_elPAY = ""
    
    Create_Infile_ELPAY_Storno gELPioPfad, gELPclientId, sTANR, sBetrag
    
    Outfile_suchen_ELPAY gELPioPfad, gELPclientId
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdl_elPAY"
    Fehler.gsFunktion = "Storno_elPAY"
    Fehler.gsFehlertext = "Im Programmteil elPAY ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function BelegWiederholung_elPAY() As String
    On Error GoTo LOKAL_ERROR

    BelegWiederholung_elPAY = ""
    
    Create_Infile_ELPAY_BelegWiederholung gELPioPfad, gELPclientId
    
    Outfile_suchen_ELPAY gELPioPfad, gELPclientId
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdl_elPAY"
    Fehler.gsFunktion = "BelegWiederholung_elPAY"
    Fehler.gsFehlertext = "Im Programmteil elPAY ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function Kassenschnitt_elPAY() As String
    On Error GoTo LOKAL_ERROR

    Kassenschnitt_elPAY = ""
    
    Create_Infile_ELPAY_Kassenschnitt gELPioPfad, gELPclientId
    
    Outfile_suchen_ELPAY gELPioPfad, gELPclientId
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdl_elPAY"
    Fehler.gsFunktion = "Kassenschnitt_elPAY"
    Fehler.gsFehlertext = "Im Programmteil elPAY ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function Diagnose_elPAY() As String
    On Error GoTo LOKAL_ERROR

    Diagnose_elPAY = ""
    
    Create_Infile_ELPAY_Diagnose gELPioPfad, gELPclientId
    
    Outfile_suchen_ELPAY gELPioPfad, gELPclientId
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdl_elPAY"
    Fehler.gsFunktion = "Diagnose_elPAY"
    Fehler.gsFehlertext = "Im Programmteil elPAY ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function Lebst_Du_noch_ELPAY() As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim iErr_Zaehler As Integer
    iErr_Zaehler = 0
    
    
    
    Lebst_Du_noch_ELPAY = False
    
    Kill gELPioPfad & "\Aktiv.$$$"
    
    PauseSi 1.5

    Kill gELPioPfad & "\Aktiv.$$$"
    
    If iErr_Zaehler = 0 Then
    
        Lebst_Du_noch_ELPAY = True
    Else
        Lebst_Du_noch_ELPAY = False
    End If
    
    
    
    
    
    
    
    
    
'    Do While FileExists(gELPioPfad & "\Aktiv.$$$") = False
'
'        PauseSi 0.1
'        iErr_Zaehler = iErr_Zaehler + 1
'        If iErr_Zaehler > iTimeout Then
'            Exit Do
'        End If
'    Loop
'
'    PauseSi 0.1
'
'    If FileExists(gELPioPfad & "\Aktiv.$$$") Then
'        Lebst_Du_noch_ELPAY = True
'    End If

Exit Function
LOKAL_ERROR:
    If err.Number = 53 Then
        iErr_Zaehler = iErr_Zaehler + 1
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "mdl_elPAY"
        Fehler.gsFunktion = "Lebst_Du_noch_ELPAY"
        Fehler.gsFehlertext = "Im Programmteil elPAY ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Function
Public Sub Outfile_suchen_ELPAY(sELPAY_Pfad As String, sClient As String)
    On Error GoTo LOKAL_ERROR
    
    Dim iErr_Zaehler As Integer
    Dim iTimeout As Integer
    iTimeout = 600
    
    iErr_Zaehler = 0
    Do While FileExists(sELPAY_Pfad & "\outfile." & sClient) = False
    
        PauseSi 0.1
        iErr_Zaehler = iErr_Zaehler + 1
        If iErr_Zaehler > iTimeout Then
            Exit Do
        End If
    Loop
    
    PauseSi 0.1
    If FileExists(sELPAY_Pfad & "\outfile." & sClient) Then
        Outfile_nachFehler_checken_ELPAY sELPAY_Pfad, sClient
        Outfile_drucken_ELPAY sELPAY_Pfad, sClient
    Else
        giELPAY_Fehler = 1
'        MsgBox "Timeout ist schneller"
    End If
    
   

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdl_elPAY"
    Fehler.gsFunktion = "Outfile_suchen_ELPAY"
    Fehler.gsFehlertext = "Im Programmteil elPAY ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub Outfile_nachFehler_checken_ELPAY(sELPAY_Pfad As String, sClient As String)
    On Error GoTo LOKAL_ERROR
    
    
    Dim cSatz                   As String
    Dim iFileNr                 As Integer
    Dim lPos                    As Long
    Dim cEinzelsatz             As String
    
    giELPAY_Fehler = 0
    lPos = 1
    
    iFileNr = FreeFile
    Open sELPAY_Pfad & "\outfile." & sClient For Binary As #iFileNr
    If LOF(iFileNr) > 0 Then
    
        cSatz = Space$(LOF(iFileNr))
        Get #iFileNr, 1, cSatz
        
        lPos = InStr(lPos, cSatz, "Fehlercode:")
        
        If lPos > 1 Then
            'auslesen
            cEinzelsatz = Mid(cSatz, lPos + 11, 4)
        End If
        
        If Val(cEinzelsatz) > 0 Then
           giELPAY_Fehler = Val(cEinzelsatz)
        Else
            giELPAY_Fehler = 0
        End If
        

    End If
    
    Close iFileNr
        
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdl_elPAY"
    Fehler.gsFunktion = "Outfile_nachFehler_checken_ELPAY"
    Fehler.gsFehlertext = "Im Programmteil elPAY ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub Outfile_drucken_ELPAY(sELPAY_Pfad As String, sClient As String)
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
    Dim lposDoppelpunkt         As Long
    Dim lposSemikolon           As Long
    
    lPos = 1
    lPosEnde = 1
    lAnzBuchstaben = 1
    
    iDruckzeilen_count = 0
    ReDim cDruckZeile(1 To 1) As String
    ReDim cDruckFormat(1 To 1) As String
    
    iFileNr = FreeFile
    Open sELPAY_Pfad & "\outfile." & sClient For Binary As #iFileNr
    If LOF(iFileNr) > 0 Then
    
        cSatz = Space$(LOF(iFileNr))
        Get #iFileNr, 1, cSatz
        
        lLenfil = Len(cSatz)
        
        Do
            
            lPosEnde = InStr(lPos, cSatz, vbCrLf)
            
            lAnzBuchstaben = lPosEnde - lPos
            cEinzelsatz = Mid(cSatz, lPos, lAnzBuchstaben)
            
            If Left(cEinzelsatz, 10) = "Druckzeile" Then
                iDruckzeilen_count = iDruckzeilen_count + 1
                ReDim Preserve cDruckZeile(1 To iDruckzeilen_count) As String
                ReDim Preserve cDruckFormat(1 To iDruckzeilen_count) As String
                
                cWert_Format = ""
                cWert = ""
                lposDoppelpunkt = 0
                lposDoppelpunkt = InStr(1, cEinzelsatz, ":")
                
                lposSemikolon = 0
                lposSemikolon = InStr(1, cEinzelsatz, ";")
                
                If lposSemikolon > 0 Then
                    cWert_Format = Mid(cEinzelsatz, lposDoppelpunkt + 1, lposSemikolon - lposDoppelpunkt - 1)
                    cWert = Mid(cEinzelsatz, lposSemikolon + 1, Len(cEinzelsatz) - lposSemikolon + 1)
                Else
                    cWert = Mid(cEinzelsatz, lposDoppelpunkt + 1, Len(cEinzelsatz) - lposDoppelpunkt + 1)
                End If
                
                cDruckFormat(iDruckzeilen_count) = cWert_Format
                cDruckZeile(iDruckzeilen_count) = cWert
            End If
            
            lPos = lPos + lAnzBuchstaben + 2
            
        Loop While lLenfil >= lPos
        
    End If
    
    Close iFileNr
    
    DruckeEndlosBeleg_ELPAY cDruckZeile(), cDruckFormat(), iDruckzeilen_count, True
        
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdl_elPAY"
    Fehler.gsFunktion = "Outfile_drucken_ELPAY"
    Fehler.gsFehlertext = "Im Programmteil elPAY ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub DruckeEndlosBeleg_ELPAY(cZeilen() As String, cFormat() As String, iMax As Integer, bschnitt As Boolean)
    On Error GoTo LOKAL_ERROR
    
    Dim aDeviceName As String
    Dim cEscapeSequenz As String
    ReDim cDruckZeile(1 To 1) As String
    ReDim cDruckFormat(1 To 1) As String
    Dim lAnzZeile As Long
    Dim lcount As Long
    Dim i As Integer
    
    'zum Test
'    setzedrucker gcBonDrucker
    
    
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
            cZeilen(i) = Space$((32 - Len(cZeilen(i))) / 2) & cZeilen(i)
        End If
        
        cDruckZeile(lAnzZeile) = cZeilen(i) & vbCrLf
        cDruckFormat(lAnzZeile) = cFormat(i)
    Next i

    
    OpenDrawer4_ELPAY aDeviceName, cDruckZeile(), cDruckFormat(), lAnzZeile, bschnitt

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdl_elPAY"
    Fehler.gsFunktion = "DruckeEndlosBeleg_ELPAY"
    Fehler.gsFehlertext = "Beim Drucken ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Public Sub OpenDrawer4_ELPAY(aDeviceName As String, cDruckZeile() As String, cDruckFormat() As String, lAnzZeile As Long, bschnitt As Boolean)
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
        ctmp = Space$(3) & cDruckZeile(lcount)
'        KonvertAsciiAnsi ctmp
        If Right(ctmp, 2) = vbCrLf Then
            ctmp = Left(ctmp, Len(ctmp) - 2)
        End If
        
        If InStr(1, cDruckFormat(lcount), "F") > 0 Then
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
            
        Else
            Printer.FontBold = False
            Printer.Print ctmp
        End If
        
    Next lcount
    

    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdl_elPAY"
    Fehler.gsFunktion = "OpenDrawer4_ELPAY"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub Create_Infile_ELPAY(sPfad As String, sCentBetrag As String, sClient As String)
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr             As Integer
    Dim sPfad_und_Datei     As String
    Dim lPos                As Long
    Dim cSatz               As String
    
    cSatz = "Funktion:00" & vbCrLf
    cSatz = cSatz & "Betrag:" & sCentBetrag & vbCrLf
    cSatz = cSatz & "Druckbreite:32" & vbCrLf
    cSatz = cSatz & "Softwarename:Winkiss" & vbCrLf
    cSatz = cSatz & "Softwareversion:" & glpVers & vbCrLf
    
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
        Fehler.gsFormular = "mdl_elPAY"
        Fehler.gsFunktion = "Create_Infile_ELPAY"
        Fehler.gsFehlertext = "Im Programmteil elPAY ist ein Fehler aufgetreten."
    End If
    
    Fehlermeldung1
End Sub
Public Sub Create_Infile_ELPAY_Manuell(sPfad As String, sCentBetrag As String, sClient As String, sPan As String, sVerfall As String)
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr             As Integer
    Dim sPfad_und_Datei     As String
    Dim lPos                As Long
    Dim cSatz               As String
    
    cSatz = "Funktion:10" & vbCrLf
    cSatz = cSatz & "Betrag:" & sCentBetrag & vbCrLf
    cSatz = cSatz & "PAN:" & sPan & vbCrLf
    cSatz = cSatz & "Verfalldatum:" & sVerfall & vbCrLf
    cSatz = cSatz & "Druckbreite:32" & vbCrLf
    cSatz = cSatz & "Softwarename:Winkiss" & vbCrLf
    cSatz = cSatz & "Softwareversion:" & glpVers & vbCrLf
    
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
        Fehler.gsFormular = "mdl_elPAY"
        Fehler.gsFunktion = "Create_Infile_ELPAY_Manuell"
        Fehler.gsFehlertext = "Im Programmteil elPAY ist ein Fehler aufgetreten."
    End If
    
    Fehlermeldung1
End Sub
Public Sub Create_Infile_ELPAY_Manuell_Storno(sPfad As String, sCentBetrag As String, sClient As String, sPan As String, sVerfall As String, sTANR As String)
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr             As Integer
    Dim sPfad_und_Datei     As String
    Dim lPos                As Long
    Dim cSatz               As String
    
    cSatz = "Funktion:11" & vbCrLf
    cSatz = cSatz & "Beleg:" & sTANR & vbCrLf
    cSatz = cSatz & "Betrag:" & sCentBetrag & vbCrLf
    cSatz = cSatz & "PAN:" & sPan & vbCrLf
    cSatz = cSatz & "Verfalldatum:" & sVerfall & vbCrLf
    cSatz = cSatz & "Druckbreite:32" & vbCrLf
    cSatz = cSatz & "Softwarename:Winkiss" & vbCrLf
    cSatz = cSatz & "Softwareversion:" & glpVers & vbCrLf
    
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
        Fehler.gsFormular = "mdl_elPAY"
        Fehler.gsFunktion = "Create_Infile_ELPAY_Manuell_Storno"
        Fehler.gsFehlertext = "Im Programmteil elPAY ist ein Fehler aufgetreten."
    End If
    
    Fehlermeldung1
End Sub
Public Sub Create_Infile_ELPAY_BelegWiederholung(sPfad As String, sClient As String)
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr             As Integer
    Dim sPfad_und_Datei     As String
    Dim lPos                As Long
    Dim cSatz               As String
    
    cSatz = "Funktion:60" & vbCrLf
    cSatz = cSatz & "Druckbreite:32" & vbCrLf
    cSatz = cSatz & "Softwarename:Winkiss" & vbCrLf
    cSatz = cSatz & "Softwareversion:" & glpVers & vbCrLf
    
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
        Fehler.gsFormular = "mdl_elPAY"
        Fehler.gsFunktion = "Create_Infile_ELPAY_BelegWiederholung"
        Fehler.gsFehlertext = "Im Programmteil elPAY ist ein Fehler aufgetreten."
    End If
    
    Fehlermeldung1
End Sub
Public Sub Create_Infile_ELPAY_Kassenschnitt(sPfad As String, sClient As String)
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr             As Integer
    Dim sPfad_und_Datei     As String
    Dim lPos                As Long
    Dim cSatz               As String
    
    cSatz = "Funktion:99" & vbCrLf
    cSatz = cSatz & "Druckbreite:32" & vbCrLf
    cSatz = cSatz & "Softwarename:Winkiss" & vbCrLf
    cSatz = cSatz & "Softwareversion:" & glpVers & vbCrLf
    
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
        Fehler.gsFormular = "mdl_elPAY"
        Fehler.gsFunktion = "Create_Infile_ELPAY_Kassenschnitt"
        Fehler.gsFehlertext = "Im Programmteil elPAY ist ein Fehler aufgetreten."
    End If
    
    Fehlermeldung1
End Sub
Public Sub Create_Infile_ELPAY_Diagnose(sPfad As String, sClient As String)
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr             As Integer
    Dim sPfad_und_Datei     As String
    Dim lPos                As Long
    Dim cSatz               As String
    
    cSatz = "Funktion:97" & vbCrLf
    cSatz = cSatz & "Druckbreite:32" & vbCrLf
    cSatz = cSatz & "Softwarename:Winkiss" & vbCrLf
    cSatz = cSatz & "Softwareversion:" & glpVers & vbCrLf
    
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
        Fehler.gsFormular = "mdl_elPAY"
        Fehler.gsFunktion = "Create_Infile_ELPAY_Diagnose"
        Fehler.gsFehlertext = "Im Programmteil elPAY ist ein Fehler aufgetreten."
    End If
    
    Fehlermeldung1
End Sub
Public Sub Create_Infile_ELPAY_Storno(sPfad As String, sClient As String, sTANR As String, sCentBetrag As String)
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr             As Integer
    Dim sPfad_und_Datei     As String
    Dim lPos                As Long
    Dim cSatz               As String
    
    cSatz = "Funktion:01" & vbCrLf
    cSatz = cSatz & "Betrag:" & sCentBetrag & vbCrLf
    cSatz = cSatz & "Beleg:" & sTANR & vbCrLf
    cSatz = cSatz & "Druckbreite:32" & vbCrLf
    cSatz = cSatz & "Softwarename:Winkiss" & vbCrLf
    cSatz = cSatz & "Softwareversion:" & glpVers & vbCrLf
    
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
        Fehler.gsFormular = "mdl_elPAY"
        Fehler.gsFunktion = "Create_Infile_ELPAY_Storno"
        Fehler.gsFehlertext = "Im Programmteil elPAY ist ein Fehler aufgetreten."
    End If
    
    Fehlermeldung1
End Sub
Public Sub lese_ELPAY_opt()
    On Error GoTo LOKAL_ERROR
    
    Dim rsrs As Recordset
    
    gELPclientId = ""
    gELPioPfad = ""
    giELPAY_Fehler = 0
    
    If NewTableSuchenDBKombi("ELPOPT", gdApp) Then

        Set rsrs = gdApp.OpenRecordset("select * from ELPOPT")
        If Not rsrs.EOF Then
            
            If Not IsNull(rsrs!clientID) Then
                gELPclientId = rsrs!clientID
            End If
            
            If Not IsNull(rsrs!ioPfad) Then
                gELPioPfad = rsrs!ioPfad
            End If
        End If
        rsrs.Close: Set rsrs = Nothing
        
        
        'Ist elpay an?
        
        
'        If Lebst_Du_noch_ELPAY() Then
'
'        Else
'            MsgBox "ELPAY5 muss noch gestartet werden", vbCritical, "Winkiss Hinweis:"
'        End If
        
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "mdl_elPAY"
    Fehler.gsFunktion = "lese_ELPAY_opt"
    Fehler.gsFehlertext = "Im Programmteil elPAY ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
