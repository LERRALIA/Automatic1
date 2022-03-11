Attribute VB_Name = "Modul3"
Option Explicit
Public Sub zeige_Best_Hist_GDPdU(cDatum As String, labelx As Label, Optional bAlsMail As Boolean = False, Optional sAGN As String = "")
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim GDPdU_DB    As Database
    Dim cPfad       As String
    Dim cPfad2      As String
    
    Screen.MousePointer = 11
    
    anzeige "normal", "Exportdatei wird erstellt...", labelx
    
    cPfad2 = gcDBPfad
    If Right$(cPfad2, 1) <> "\" Then
        cPfad2 = cPfad2 & "\"
    End If
    
    cPfad = gcDBPfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "GDPdU\GDPdU.MDB"
    
    Set GDPdU_DB = OpenDatabase(cPfad, False, False, "MS Access;PWD=" & gsGDPdU_Passwort)
    
    loeschNEW "Lieflw", GDPdU_DB
    CreateTable "LIEFLW", GDPdU_DB
    
    loeschNEW "ArtTemp", GDPdU_DB

    sSQL = "select artnr,bestand,linr,ekpr,kvkpr1,bezeich into arttemp from GLAGER_GDPdU "
    sSQL = sSQL & " where Bestand > 0 "
    sSQL = sSQL & " and datum = " & CLng(DateValue(cDatum)) & ""
    GDPdU_DB.Execute sSQL, dbFailOnError
    
    sSQL = "Update arttemp AS A inner join "
    sSQL = sSQL & "[;DATABASE=" & cPfad2 & "kissdata.mdb;pwd=" & gsPasswort & "].ARTIKEL AS B ON A.artnr = b.artnr"
    sSQL = sSQL & " set a.ekpr = b.lekpr where a.ekpr = 0 "
    GDPdU_DB.Execute sSQL, dbFailOnError

    sSQL = "INSERT into LIEFLW Select LINR, Sum(arttemp.BESTAND) as BESTAND "
    sSQL = sSQL & ", Sum(KVKPR1* arttemp.BESTAND) as LagerVK"
    sSQL = sSQL & ", Sum(EKPR* arttemp.BESTAND) as LagerEK"
    sSQL = sSQL & " from arttemp "
    sSQL = sSQL & " Where arttemp.Bestand > 0  "

    sSQL = sSQL & " group BY arttemp.LINR "
    GDPdU_DB.Execute sSQL, dbFailOnError

    loeschNEW "ArtTemp", GDPdU_DB
    
    sSQL = "Update LIEFLW set BGrund = 'Schnitteinkaufswert' "
    GDPdU_DB.Execute sSQL, dbFailOnError
    
    
    sSQL = "Update LiefLW inner join "
    sSQL = sSQL & "[;DATABASE=" & cPfad2 & "kissdata.mdb;pwd=" & gsPasswort & "].lisrt on lieflw.linr = lisrt.linr set lieflw.LIEFBEZ = lisrt.liefbez"
    GDPdU_DB.Execute sSQL, dbFailOnError
    
    sSQL = "Update LiefLW set auswahl = '" & cDatum & "' "
    GDPdU_DB.Execute sSQL, dbFailOnError
    
    loeschNEW "LiefLW", gdBase
    TransferTab GDPdU_DB, gcDBPfad & "\kissdata.mdb", "LiefLW"
    
    If bAlsMail = True Then
    
        Dim cPfad1 As String
        cPfad1 = gcDBPfad
        If Right$(cPfad1, 1) <> "\" Then
            cPfad1 = cPfad1 & "\"
        End If
        
        Dim ctmp As String
        Dim cName As String
        Dim lWert As Long
        Dim sTime As String
        
        sTime = TimeValue(Now)
        sTime = Right(sTime, 8)
        sTime = Left(sTime, 5)

        lWert = DateValue(Now)
        ctmp = Format$(lWert, "MM.DD")
       
        ctmp = ctmp & sTime
        ctmp = SwapStr(ctmp, ".", "")
        ctmp = SwapStr(ctmp, ":", "")

        cName = ctmp
            
        Kill cPfad1 & "Export\*.txt"
    
        reportbildschirmtoText "awkl46j", cPfad1 & "Export\" & cName & "_" & gcKasNum & ".txt"
                
        gcBestellEmail.Attachment1 = cPfad1 & "Export\" & cName & "_" & gcKasNum & ".txt"
        Screen.MousePointer = 0
        frmWKL129.Show 1
        
    Else
        reportbildschirm "", "awkl46j"
    End If
    
    GDPdU_DB.Close
    
Exit Sub
LOKAL_ERROR:

    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul3"
        Fehler.gsFunktion = "zeige_Best_Hist_GDPdU"
        Fehler.gsFehlertext = "Im Programmteil GDPdU/DATEV ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Public Sub zeige_Best_Hist_Einzel_GDPdU(cDatum As String, labelx As Label, Optional bAlsMail As Boolean = False, Optional sAGN As String = "")
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim GDPdU_DB    As Database
    Dim cPfad       As String
    Dim cPfad2      As String
    
    Screen.MousePointer = 11
    
    anzeige "normal", "Exportdatei wird erstellt...", labelx
    
    cPfad2 = gcDBPfad
    If Right$(cPfad2, 1) <> "\" Then
        cPfad2 = cPfad2 & "\"
    End If
    
    cPfad = gcDBPfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "GDPdU\GDPdU.MDB"
    
    Set GDPdU_DB = OpenDatabase(cPfad, False, False, "MS Access;PWD=" & gsGDPdU_Passwort)
    
    loeschNEW "ARTHISTE", GDPdU_DB
    CreateTableT2 "ARTHISTE", GDPdU_DB
    
    loeschNEW "ArtTemp", GDPdU_DB

    sSQL = "select artnr,bestand,linr,ekpr,kvkpr1,bezeich into arttemp from GLAGER_GDPdU "
    sSQL = sSQL & " where Bestand > 0 "
    sSQL = sSQL & " and datum = " & CLng(DateValue(cDatum)) & ""
    GDPdU_DB.Execute sSQL, dbFailOnError
    
    sSQL = "Update arttemp AS A inner join "
    sSQL = sSQL & "[;DATABASE=" & cPfad2 & "kissdata.mdb;pwd=" & gsPasswort & "].ARTIKEL AS B ON A.artnr = b.artnr"
    sSQL = sSQL & " set a.ekpr = b.lekpr where a.ekpr = 0 "
    GDPdU_DB.Execute sSQL, dbFailOnError
    
    sSQL = "INSERT into ARTHISTE Select artnr,bezeich,linr,ekpr,kvkpr1,'' as liefbez"
    sSQL = sSQL & " ,'Schnitteinkaufswert' as BGRUND "
    sSQL = sSQL & " ,'" & cDatum & "'  as AUSWAHL "
    sSQL = sSQL & " , Sum(arttemp.BESTAND) as BESTAND "
    sSQL = sSQL & " from arttemp "
    sSQL = sSQL & " Where arttemp.Bestand > 0  "

    sSQL = sSQL & " group BY artnr "
    sSQL = sSQL & ",bezeich,linr,ekpr,kvkpr1"
    GDPdU_DB.Execute sSQL, dbFailOnError

    loeschNEW "ArtTemp", GDPdU_DB
    
    sSQL = "Update ARTHISTE inner join "
    sSQL = sSQL & "[;DATABASE=" & cPfad2 & "kissdata.mdb;pwd=" & gsPasswort & "].lisrt on ARTHISTE.linr = lisrt.linr set ARTHISTE.LIEFBEZ = lisrt.liefbez"
    GDPdU_DB.Execute sSQL, dbFailOnError
    
    loeschNEW "ARTHISTE", gdBase
    TransferTab GDPdU_DB, gcDBPfad & "\kissdata.mdb", "ARTHISTE"
    
    reportbildschirm "", "awkl46k"
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "zeige_Best_Hist_Einzel_GDPdU"
    Fehler.gsFehlertext = "Im Programmteil Inventur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Public Function WhatIsXfromXtab(cSchluesselwert As String, cSchluesselspalte As String, cTab As String, sSpalte As String) As String
    On Error GoTo LOKAL_ERROR
    Dim sSQL As String
    Dim rs As Recordset
    
    WhatIsXfromXtab = ""
    
    sSQL = "select " & sSpalte & " as maxi from " & cTab & " where " & cSchluesselspalte & " = " & cSchluesselwert
    Set rs = gdBase.OpenRecordset(sSQL)
    If Not rs.EOF Then
        If Not IsNull(rs!maxi) Then
            WhatIsXfromXtab = rs!maxi
        End If
    End If
    rs.Close
    
    Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "WhatIsXfromXtab"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Sub Etikettenbestellung_Per_Mail(lKissArtnr As Long, sEtikettenbezeichnung As String)
On Error GoTo LOKAL_ERROR


    'Mail schicken

    gcBestellEmail.Attachment1 = ""
    gcBestellEmail.Attachment2 = ""
    gcBestellEmail.Attachment3 = ""
    gcBestellEmail.Attachment4 = ""
    gcBestellEmail.Attachment5 = ""

    Dim sTemp As String
    
    gcBestellEmail.Subject = "Etikettenbestellung (" & gFirma.FirmaName & ")"
    
    sTemp = "Liebes Kiss Team," & vbCrLf & vbCrLf
    sTemp = sTemp & "hiermit möchte ich eine Etikettenbestellung in Auftrag geben." & vbCrLf
    sTemp = sTemp & "Kiss Artikelnummer: " & lKissArtnr & vbCrLf
    sTemp = sTemp & "Bestellmenge: 1" & vbCrLf
    sTemp = sTemp & "Bezeichnung: " & sEtikettenbezeichnung & vbCrLf
    
    sTemp = sTemp & vbCrLf
    sTemp = sTemp & "Lieferanschrift:" & vbCrLf
    
    sTemp = sTemp & gFirma.FirmaName & vbCrLf
    sTemp = sTemp & gFirma.Plz & " " & gFirma.Ort & vbCrLf
    sTemp = sTemp & gFirma.strasse & vbCrLf
    
    sTemp = sTemp & vbCrLf
    sTemp = sTemp & "Ansprechpartner:                                         " & vbCrLf
    sTemp = sTemp & "Tel: " & gFirma.Tel & vbCrLf
    
    sTemp = sTemp & vbCrLf
    sTemp = sTemp & "meine Anmerkung:"
    sTemp = sTemp & vbCrLf & vbCrLf
    
    
    
    
    
    sTemp = sTemp & vbCrLf & vbCrLf
    sTemp = sTemp & "Diese Bestellung ist erst verbindlich nach telefonischer Bestätigung durch die Firma KISS Hannover."
    
    
    gcBestellEmail.Message = sTemp
    'gcBestellEmail.Recipient = "s.fazlija@kisswws.de"
    gcBestellEmail.Recipient = "vertrieb@kisswws.de"
            
    frmWKL129.Show 1
            
    gcBestellEmail.Attachment1 = ""
    gcBestellEmail.Subject = ""
    gcBestellEmail.Message = ""
    gcBestellEmail.Recipient = ""


    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "Etikettenbestellung_Per_Mail"
    Fehler.gsFehlertext = "Beim Öffnen eines Programmteils ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Public Sub Termine_versenden_Frage()
On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsrs As DAO.Recordset
    
    Dim bGehrein As Boolean
    bGehrein = False
    

    If NewTableSuchenDBKombi("SMS_UEBERSICHT", gdBase) = False Then
        frmWKL217.Show 1
    Else
    
        cSQL = "Select max(DATUMERINNERUNG) as maxdat from SMS_UEBERSICHT"
        Set rsrs = gdBase.OpenRecordset(cSQL)
        If Not rsrs.EOF Then
            If Not IsNull(rsrs!Maxdat) Then
                If CLng(DateValue(Now)) <> CLng(DateValue(rsrs!Maxdat)) Then
                    bGehrein = True
'                    frmWKL217.Show 1
                End If
            End If
        End If
        rsrs.Close: Set rsrs = Nothing
        
        If bGehrein = True Then
            frmWKL217.Show 1
        End If
    
    End If


        
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "Termine_versenden_Frage"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub LeseOpeningsWKL82()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim cWoTag As String
    Dim iLfdNr As Integer
    Dim cVon As String
    Dim cBis As String
    Dim cZeitblock As String
    Dim iWert As Integer
    
    Dim dUhrzeit As Double
    Dim dStartzeit As Double
    Dim dZeit As Double
    Dim lcount As Long
    
    iWert = 0
    cSQL = "Select * from OPENINGS order by WOTAG, LFDNR"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            iWert = iWert + 1
            If Not IsNull(rsrs!WoTag) Then
                cWoTag = rsrs!WoTag
            Else
                cWoTag = "0"
            End If
            If Not IsNull(rsrs!LFDNR) Then
                iLfdNr = rsrs!LFDNR
            Else
                iLfdNr = 0
            End If
            If Not IsNull(rsrs!Von) Then
                cVon = rsrs!Von
            Else
                cVon = ""
            End If
            If Not IsNull(rsrs!Bis) Then
                cBis = rsrs!Bis
            Else
                cBis = ""
            End If
            If Not IsNull(rsrs!Zeitblock) Then
                cZeitblock = rsrs!Zeitblock
            Else
                cZeitblock = ""
            End If
            
            If cZeitblock <> "" Then
                gcZeitBlock = cZeitblock
            End If
            
            gZeiten(iWert).WoTag = Val(cWoTag)
            gZeiten(iWert).LFDNR = iLfdNr
            gZeiten(iWert).Von = cVon
            gZeiten(iWert).Bis = cBis
            gZeiten(iWert).Zeitblock = Val(cZeitblock)
            rsrs.MoveNext
        Loop
    Else
        For iWert = 1 To 21
            gZeiten(iWert).WoTag = 0
            gZeiten(iWert).LFDNR = 0
            gZeiten(iWert).Von = ""
            gZeiten(iWert).Bis = ""
            gZeiten(iWert).Zeitblock = 0
            If cZeitblock <> "" Then
                gcZeitBlock = "15"
            End If
        Next iWert
    End If
    rsrs.Close: Set rsrs = Nothing
    
    dUhrzeit = Val(gcZeitBlock) / 1440
    gcZeitBlock = Format$(dUhrzeit, "HH:MM")
    
    dStartzeit = 1
    For lcount = 1 To 21        '(eine Woche mit max. 3 Öffnungszeiten pro Tag)
        If gZeiten(lcount).Von <> "" Then
            dZeit = TimeValue(gZeiten(lcount).Von)
            If dZeit < dStartzeit Then
                dStartzeit = dZeit
            End If
        End If
    Next lcount
    
    dStartzeit = dStartzeit - dUhrzeit
    gcStartZeit = Format$(dStartzeit, "HH:MM")
    
    dStartzeit = 0
    For lcount = 1 To 21
        If gZeiten(lcount).Bis <> "" Then
            dZeit = TimeValue(gZeiten(lcount).Bis)
            If dZeit > dStartzeit Then
                dStartzeit = dZeit
            End If
        End If
    Next lcount
    
    gcEndeZeit = Format$(dStartzeit, "HH:MM")
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "LeseOpeningsWKL82"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub TrageSMSBenachrichtigungEin(DateBEH As Date)
On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    
    If NewTableSuchenDBKombi("SMS_UEBERSICHT", gdBase) = False Then
        CreateTableT3 "SMS_UEBERSICHT", gdBase
    End If
    
    cSQL = "Delete from SMS_UEBERSICHT where DATUMBEHANDLUNG = " & CLng(DateBEH) & " "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into SMS_UEBERSICHT ( "
    cSQL = cSQL & " DATUMBEHANDLUNG,DATUMERINNERUNG "
    cSQL = cSQL & " ) values ( "
    cSQL = cSQL & " " & CLng(DateBEH) & " "
    cSQL = cSQL & ", " & CLng(DateValue(Now)) & " "
    cSQL = cSQL & " ) "
    gdBase.Execute cSQL, dbFailOnError


Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "TrageSMSBenachrichtigungEin"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Public Sub VersendeTermineSMS(DateHeut As Date)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim rsrs        As Recordset
    Dim cVon        As Date
    Dim cFeld       As String
    Dim cDauer      As String
    Dim lBehDat     As Long
    Dim lWeekday    As Long
    
    Screen.MousePointer = 11
            
    cVon = DateHeut
    
    loeschNEW "TERMPRINT_EP", gdBase
    CreateTableT2 "TERMPRINT_EP", gdBase
    
    cSQL = "Insert into TERMPRINT_EP select "
    cSQL = cSQL & " BEDNAME "
    cSQL = cSQL & ", BEDNU "
    cSQL = cSQL & ", BEHANDLUNG "
    cSQL = cSQL & ", BUCHUNGSNR "
    cSQL = cSQL & ", DATUM "
    cSQL = cSQL & ", KABINE "
    cSQL = cSQL & ", KUERZEL "
    cSQL = cSQL & ", KUNDNR "
    cSQL = cSQL & ", UHRZEIT "
    cSQL = cSQL & ", BEDEINTRAG "
    cSQL = cSQL & " from termine where datum = " & CLng(cVon) & " "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update TERMPRINT_EP set von = " & CLng(cVon)
    gdBase.Execute cSQL, dbFailOnError

    cSQL = "Update TERMPRINT_EP set adate = datum "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update TERMPRINT_EP inner join kunden on TERMPRINT_EP.KUNDNR = Kunden.Kundnr "
    cSQL = cSQL & " set TERMPRINT_EP.Name = Kunden.Name "
    cSQL = cSQL & " , TERMPRINT_EP.TEL = Kunden.TEL "
    cSQL = cSQL & " , TERMPRINT_EP.MOBILTEL = Kunden.MOBILTEL "
    cSQL = cSQL & " , TERMPRINT_EP.VORNAME = Kunden.VORNAME "
    cSQL = cSQL & " , TERMPRINT_EP.EMAIL = Kunden.EMAIL "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Select * from TERMPRINT_EP "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!Behandlung) Then
                cFeld = rsrs!Behandlung
            Else
                cFeld = ""
            End If
            
            cFeld = SwapStr(cFeld, Chr(13), " ")
            cFeld = SwapStr(cFeld, Chr(10), " ")
            
            rsrs.Edit
            rsrs!Behandlung = Trim(cFeld)

            rsrs.Update
    
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Dim cStartzeit As String
    Dim cEndZeit As String
    Dim dStart As Double
    Dim dEnde As Double
    Dim dDauer As Double
    Dim lBuchnr As Long
    
    loeschNEW "TERMPRINT_MEP", gdBase
    CreateTableT2 "TERMPRINT_MEP", gdBase
    
    cSQL = "Select BUCHUNGSNR, max(Uhrzeit) as maxizeit, min(Uhrzeit) as minizeit from TERMPRINT_EP group by Buchungsnr "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        
        Do While Not rsrs.EOF
        
            If Not IsNull(rsrs!BUCHUNGSNR) Then
                lBuchnr = rsrs!BUCHUNGSNR
            End If
            
            If Not IsNull(rsrs!maxizeit) Then
                cEndZeit = rsrs!maxizeit
            Else
                cEndZeit = ""
            End If
            
            If Not IsNull(rsrs!minizeit) Then
                cStartzeit = rsrs!minizeit
            Else
                cStartzeit = ""
            End If
            
            dStart = TimeValue(cStartzeit)
            dEnde = TimeValue(cEndZeit)
            dEnde = dEnde + TimeValue(gcZeitBlock)
    
            dDauer = dEnde - dStart
            cDauer = Format$(dDauer, "HH:MM")
            
            
            cSQL = "Insert into TERMPRINT_MEP (buchnr,Dauer,Uhrzeit_ende,UHRZEIT) values ("
            cSQL = cSQL & " " & lBuchnr & ",'" & cDauer & "',  '" & Format$(dEnde, "HH:MM") & "',  '" & Format$(dStart, "HH:MM") & "')"
            gdBase.Execute cSQL, dbFailOnError
        
    
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    cSQL = "Update TERMPRINT_MEP M inner join TERMPRINT_EP E on M.buchnr = E.BUCHUNGSNR "
    cSQL = cSQL & " set M.Name = E.Name "
    cSQL = cSQL & ",M.BEDNAME = E.BEDNAME "
    cSQL = cSQL & ",M.BEDNU = E.BEDNU  "
    cSQL = cSQL & ",M.BEHANDLUNG = E.BEHANDLUNG "
    cSQL = cSQL & ",M.DATUM = E.DATUM  "
    cSQL = cSQL & ",M.KABINE = E.KABINE  "
    cSQL = cSQL & ",M.KUERZEL = E.KUERZEL  "
    cSQL = cSQL & ",M.KUNDNR = E.KUNDNR  "
    
    cSQL = cSQL & ",M.TEL = E.TEL  "
    cSQL = cSQL & ",M.MOBILTEL = E.MOBILTEL  "
    cSQL = cSQL & ",M.VORNAME = E.VORNAME  "

    
    cSQL = cSQL & ",M.EMAIL = E.EMAIL  "
    cSQL = cSQL & ",M.adate = E.adate  "
    cSQL = cSQL & ",M.von = E.von  "
    cSQL = cSQL & ",M.bis = E.bis  "
    cSQL = cSQL & ",M.BEDEINTRAG = E.BEDEINTRAG   "
    gdBase.Execute cSQL, dbFailOnError
    
    If Not SpalteInTabellegefundenNEW("TERMPRINT_MEP", "Anrede", gdBase) Then
        SpalteAnfuegenNEW "TERMPRINT_MEP", "Anrede", "Text(35)", gdBase
        SpalteAnfuegenNEW "TERMPRINT_MEP", "Geschlecht", "Text(1)", gdBase
        SpalteAnfuegenNEW "TERMPRINT_MEP", "AnzTermine", "Text(1)", gdBase
    End If
    
    cSQL = "Update TERMPRINT_MEP set AnzTermine = '1' "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update TERMPRINT_MEP inner join kunden on TERMPRINT_MEP.KUNDNR = Kunden.Kundnr "
    cSQL = cSQL & " set TERMPRINT_MEP.Anrede = Kunden.Anrede "
    cSQL = cSQL & " , TERMPRINT_MEP.Geschlecht = Kunden.Geschlecht "
    gdBase.Execute cSQL, dbFailOnError
    

    Dim cAbsenderEmail As String
    Dim cAnEmailadresse As String
    Dim cBetreff As String
    Dim cMessagetext As String
    Dim sAttachment As String
    
    Dim sAnrede As String
    Dim sName As String
    Dim sDatum As String
    Dim sUhrzeit As String
    Dim sTel As String
    
    Dim lcount As Long
    lcount = 0
        
    sAttachment = ""
    
    
            
    cAbsenderEmail = ermFirmenMail
    If cAbsenderEmail = "" Then
        MsgBox "Bitte auch in den Unternehmensdaten eine Emaildadresse als Absendermailadresse hinterlegen (Service/Einstellungen/Unternehmens-Daten)", vbInformation, "Winkiss Hinweis:"
        Exit Sub
    End If
    
    If NewTableSuchenDBKombi("SMSTEXT", gdBase) = False Then
        MsgBox "Bitte hinterlegen Sie erst den SMS-Text (Termine/Vorgaben/SMStext)", vbInformation, "Winkiss Hinweis:"
        Exit Sub
    End If
    
    
    
    Dim sZeile1 As String
    Dim sZeile2 As String
    Dim sZeile3 As String
    Dim sZeile4 As String
    Dim sZeile5 As String
    Dim sZeile6 As String
    Dim sZeile7 As String
    
    Dim i As Integer
    Dim sZeichen(18) As String
    
    For i = 0 To 17
        sZeichen(i) = ""
    Next i
    
    cSQL = "Select * from SMSTEXT"
    
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!Zeile1) Then
            sZeile1 = rsrs!Zeile1
        End If
        
        If Not IsNull(rsrs!Zeile2) Then
            sZeile2 = rsrs!Zeile2
        End If
        
        If Not IsNull(rsrs!Zeile3) Then
            sZeile3 = rsrs!Zeile3
        End If
        
        If Not IsNull(rsrs!Zeile4) Then
            sZeile4 = rsrs!Zeile4
        End If
        
        If Not IsNull(rsrs!Zeile5) Then
            sZeile5 = rsrs!Zeile5
        End If
        
        If Not IsNull(rsrs!Zeile6) Then
            sZeile6 = rsrs!Zeile6
        End If
        
        If Not IsNull(rsrs!Zeile7) Then
            sZeile7 = rsrs!Zeile7
        End If
        
        
        
        
        If Not IsNull(rsrs!bo0) Then
            If rsrs!bo0 = -1 Then
                sZeichen(0) = vbCrLf
            End If
        End If
        
        If Not IsNull(rsrs!bo1) Then
            If rsrs!bo1 = -1 Then
                sZeichen(1) = vbCrLf
            End If
        End If
        
        If Not IsNull(rsrs!bo2) Then
            If rsrs!bo2 = -1 Then
                sZeichen(2) = vbCrLf
            End If
        End If
        
        If Not IsNull(rsrs!bo3) Then
            If rsrs!bo3 = -1 Then
                sZeichen(3) = vbCrLf
            End If
        End If
        
        If Not IsNull(rsrs!bo4) Then
            If rsrs!bo4 = -1 Then
                sZeichen(4) = vbCrLf
            End If
        End If
        
        If Not IsNull(rsrs!bo5) Then
            If rsrs!bo5 = -1 Then
                sZeichen(5) = vbCrLf
            End If
        End If
        
        If Not IsNull(rsrs!bo6) Then
            If rsrs!bo6 = -1 Then
                sZeichen(6) = vbCrLf
            End If
        End If
        
        If Not IsNull(rsrs!bo7) Then
            If rsrs!bo7 = -1 Then
                sZeichen(7) = vbCrLf
            End If
        End If
        
        If Not IsNull(rsrs!bo8) Then
            If rsrs!bo8 = -1 Then
                sZeichen(8) = vbCrLf
            End If
        End If
        
        If Not IsNull(rsrs!bo9) Then
            If rsrs!bo9 = -1 Then
                sZeichen(9) = vbCrLf
            End If
        End If
        
        If Not IsNull(rsrs!bo10) Then
            If rsrs!bo10 = -1 Then
                sZeichen(10) = vbCrLf
            End If
        End If
        
        If Not IsNull(rsrs!bo11) Then
            If rsrs!bo11 = -1 Then
                sZeichen(11) = vbCrLf
            End If
        End If
        
        If Not IsNull(rsrs!bo12) Then
            If rsrs!bo12 = -1 Then
                sZeichen(12) = vbCrLf
            End If
        End If
        
        If Not IsNull(rsrs!bo13) Then
            If rsrs!bo13 = -1 Then
                sZeichen(13) = vbCrLf
            End If
        End If
        
        If Not IsNull(rsrs!bo14) Then
            If rsrs!bo14 = -1 Then
                sZeichen(14) = vbCrLf
            End If
        End If
        
        If Not IsNull(rsrs!bo15) Then
            If rsrs!bo15 = -1 Then
                sZeichen(15) = vbCrLf
            End If
        End If
        
        If Not IsNull(rsrs!bo16) Then
            If rsrs!bo16 = -1 Then
                sZeichen(16) = vbCrLf
            End If
        End If
        
        If Not IsNull(rsrs!bo17) Then
            If rsrs!bo17 = -1 Then
                sZeichen(17) = vbCrLf
            End If
        End If
        
    End If
    rsrs.Close: Set rsrs = Nothing
    
    'Ende der SMS INFOS
    
    Dim rsBonPause As DAO.Recordset
    Dim sUebersichtMess    As String
    sUebersichtMess = ""
    
    Dim lAnz As Long
    
    Dim sKUNDNR As String
    cSQL = "Select distinct(kundnr) from TERMPRINT_MEP  "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
        
            sKUNDNR = ""
            
            If Not IsNull(rsrs!Kundnr) Then
                sKUNDNR = rsrs!Kundnr
            End If
            
            
            cSQL = " Select count(kundnr) as anz from TERMPRINT_MEP where KUNDNR = " & sKUNDNR & ""
            Set rsBonPause = gdBase.OpenRecordset(cSQL)
            
            If Not rsBonPause.EOF Then
            
                If Not IsNull(rsBonPause!anz) Then
                    lAnz = Val(rsBonPause!anz)
                End If
            
                
                
                If lAnz > 1 Then
                
                
                    Dim rsRSminUhrz As DAO.Recordset
                    
                    Dim sMinUhtz As String
                    
                    
                    cSQL = "Select Uhrzeit from TERMPRINT_MEP where KUNDNR = " & sKUNDNR & " order by Uhrzeit asc"
                    Set rsRSminUhrz = gdBase.OpenRecordset(cSQL)
                    
                    If Not rsRSminUhrz.EOF Then
                    
                        If Not IsNull(rsRSminUhrz!Uhrzeit) Then
                            sMinUhtz = rsRSminUhrz!Uhrzeit
                        End If
                    
            
                        If lAnz > 1 Then
                    
                            cSQL = "Delete * from TERMPRINT_MEP where KUNDNR = " & sKUNDNR & " and Uhrzeit <> '" & sMinUhtz & "'"
                            gdBase.Execute cSQL, dbFailOnError
                            
                            cSQL = "Update TERMPRINT_MEP set anztermine = " & lAnz & " where KUNDNR = " & sKUNDNR & " and Uhrzeit = '" & sMinUhtz & "'"
                            gdBase.Execute cSQL, dbFailOnError
                        
                        End If
                    
                    End If
                    
                    rsRSminUhrz.Close
            

                
                End If
            
            End If
            
            rsBonPause.Close
    
    
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
    Dim sAnzTermine As String
            
    cSQL = "Select * from TERMPRINT_MEP  "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        
        Do While Not rsrs.EOF
        
            sAnzTermine = ""
        
            If Not IsNull(rsrs!AnzTermine) Then
                sAnzTermine = rsrs!AnzTermine
            End If
            
                    
            sAnrede = ""
        
            If Not IsNull(rsrs!anrede) Then
                sAnrede = rsrs!anrede
            End If
            
            If UCase(sAnrede) = "FRAU" Or UCase(sAnrede) = "HERR" Then
            
            Else
            
                If Not IsNull(rsrs!geschlecht) Then
                    sAnrede = rsrs!geschlecht
                End If
                
                If UCase(sAnrede) = "M" Or UCase(sAnrede) = "W" Then
                    If UCase(sAnrede) = "M" Then
                        sAnrede = "Herr"
                    ElseIf UCase(sAnrede) = "W" Then
                        sAnrede = "Frau"
                    End If
                Else
                    sAnrede = "Frau"
                End If
            End If
            
            sName = ""
            If Not IsNull(rsrs!name) Then
                sName = rsrs!name
            End If
            
            
            
            sDatum = ""
            If Not IsNull(rsrs!Datum) Then
            
                sDatum = Format(rsrs!Datum, "DD.MM.")
            
                lBehDat = CLng(rsrs!Datum)
                lWeekday = Weekday(lBehDat, vbMonday)
            
                Select Case lWeekday
                    Case Is = 1 '"MO"
                        sDatum = "Mo " & sDatum
                    Case Is = 2 '"DI"
                        sDatum = "Di " & sDatum
                    Case Is = 3 '"MI"
                        sDatum = "Mi " & sDatum
                    Case Is = 4 '"DO"
                        sDatum = "Do " & sDatum
                    Case Is = 5 '"FR"
                        sDatum = "Fr " & sDatum
                    Case Is = 6 '"SA"
                        sDatum = "Sa " & sDatum
                    Case Is = 7 '"SO"
                        sDatum = "So " & sDatum
                End Select
            End If
            
            sUhrzeit = ""
            If Not IsNull(rsrs!Uhrzeit) Then
                sUhrzeit = rsrs!Uhrzeit
            End If
            
            
            sTel = ""
            If Not IsNull(rsrs!Mobiltel) Then
                sTel = Trim(rsrs!Mobiltel)
            End If
            
''            If sTel = "" Then
''                If Not IsNull(rsrs!Tel) Then
''                    sTel = Trim(rsrs!Tel)
''                End If
''            End If
            
            sTel = SwapStr(sTel, "  ", "")
            sTel = SwapStr(sTel, " ", "")
            sTel = SwapStr(sTel, "/", "")
            sTel = SwapStr(sTel, "\", "")
            sTel = SwapStr(sTel, "-", "")

            If sTel <> "" Then
                If IsNumeric(sTel) Then
            
                    cAnEmailadresse = sTel & "@echoemail.net"
    
                    'schicke Mail an die hinterlegte Adresse
                    
                    cBetreff = ""
                    
                    If UCase(sAnrede) = "FRAU" Then
                        cMessagetext = "Liebe Frau " & sName & "," & sZeichen(0) & sZeichen(1)
                    ElseIf UCase(sAnrede) = "HERR" Then
                        cMessagetext = "Lieber Herr " & sName & "," & sZeichen(0) & sZeichen(1)
                    End If
                    
                    cMessagetext = cMessagetext & sZeile1 & sZeichen(2) & sZeichen(3)
                    cMessagetext = cMessagetext & sDatum & sZeichen(4)
                    cMessagetext = cMessagetext & sZeile2 & sZeichen(5) & sZeichen(6)
                    cMessagetext = cMessagetext & sUhrzeit & sZeichen(7)
                    cMessagetext = cMessagetext & sZeile3 & sZeichen(8) & sZeichen(9)
                    cMessagetext = cMessagetext & sZeile4 & sZeichen(10) & sZeichen(11)
                    cMessagetext = cMessagetext & sZeile5 & sZeichen(12) & sZeichen(13)
                    cMessagetext = cMessagetext & sZeile6 & sZeichen(14) & sZeichen(15)
                    cMessagetext = cMessagetext & sZeile7 & sZeichen(16) & sZeichen(17)
                    
                    If Val(sAnzTermine) > 1 Then
                    
                        Dim lAnzAnzeige As Long
                        lAnzAnzeige = Val(sAnzTermine) - 1
                        
                        If lAnzAnzeige = 1 Then
                            cMessagetext = cMessagetext & "Achtung: Sie haben " & lAnzAnzeige & " weiteren Termin."
                        Else
                            cMessagetext = cMessagetext & "Achtung: Sie haben " & lAnzAnzeige & " weitere Termine."
                        End If
                    
                        
                    End If
                    
''                    MsgBox cMessagetext
                    

                    schickeMailimHintergrundSSL ermFirmenBez, cAbsenderEmail, "", cAnEmailadresse _
                    , cAbsenderEmail, gcSMTP_SERVER, gcSMTP_PORT, gcSMTP_USER, gcSMTP_PW, cBetreff, cMessagetext, sAttachment
                    
                    sUebersichtMess = sUebersichtMess & sAnrede & " " & sName & " " & sTel & " " & sDatum & " " & sUhrzeit & vbCrLf & vbCrLf

                    lcount = lcount + 1
                End If
                
            End If
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Dim sUebersichtBetreff As String
    
    Dim sUebersichtAttach  As String
    sUebersichtAttach = ""
    
    sUebersichtBetreff = lcount & " SMS wurde/n erfolgreich versendet."
    
    
    Dim sAnMail As String
    sAnMail = ermFirmenMail
    
    schickeMailimHintergrundSSL ermFirmenBez, cAbsenderEmail, cAbsenderEmail, sAnMail _
    , "bestsend@kisswws.de", gcSMTP_SERVER, gcSMTP_PORT, gcSMTP_USER, gcSMTP_PW, sUebersichtBetreff, sUebersichtMess, sUebersichtAttach
    
    
    
    
    
    
    If gsTerminReminderstart = "" Then
    
        MsgBox lcount & " SMS wurde/n erfolgreich versendet.", vbInformation + vbOKOnly, "Winkiss Hinweis:"
        
    End If
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "VersendeTermineSMS"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
End Sub
Public Function checkLinrForKISS(lblx As Label) As Long
    On Error GoTo LOKAL_ERROR
    
    If checkLinrForKISS = 0 Then
        Screen.MousePointer = 0
        gF2Prompt.cFeld = ""
        gF2Prompt.cWert = ""
        gF2Prompt.cWert2 = ""
        gF2Prompt.cWahl = ""
        gF2Prompt.bMultiple = False
        
        gF2Prompt.cFeld = "LINR"
        If gF2Prompt.cFeld <> "" Then
            gsAnzeige00a = "Bitte einen Lieferant auswählen!"
            frmWK00a.Show 1
        End If
        gsAnzeige00a = ""
        
        anzeige "normal", "Der Lieferant: " & gF2Prompt.cWahl & " wurde zugeordnet.", lblx
        
        If gF2Prompt.cWahl <> "" Then
             checkLinrForKISS = CDbl(gF2Prompt.cWahl)
        End If
    End If
   
Exit Function
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "checkLinrForKISS"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Public Sub delInterart(cArtNr As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If cArtNr = "" Then
        Exit Sub
    End If
    
    sSQL = "Delete from INTERART where ARTNR = " & cArtNr
    gdBase.Execute sSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "delInterart"
    Fehler.gsFehlertext = "Im Programmteil Shop Informationen ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Public Function NextfreieArtnr(lartV As Long, lartB As Long) As Long
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    
    NextfreieArtnr = 0

    Do While NextfreieArtnr = 0
        If lartV >= lartB Then
            NextfreieArtnr = 0
            Exit Function
            
        Else
        
            If lartV = 666665 Then
                lartV = lartV + 2
            Else
                lartV = lartV + 1
            End If
            
            sSQL = "Select * from artikel where artnr = " & lartV
            Set rsrs = gdBase.OpenRecordset(sSQL)
            If rsrs.RecordCount = 0 Then
                NextfreieArtnr = lartV
            Else
                NextfreieArtnr = 0
            End If
        End If
    Loop
    
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "NextfreieArtnr"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Public Function HoleFreieArtikelNrab(lartab As Long, lartbis As Long) As Long
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    
    HoleFreieArtikelNrab = 0
    
    Do While HoleFreieArtikelNrab = 0
        If lartab >= lartbis Then
            HoleFreieArtikelNrab = 0
            Exit Function
        Else
            
            If lartab = 666665 Then
                lartab = lartab + 2
            Else
                lartab = lartab + 1
            End If
            
            sSQL = "Select * from artikel where artnr = " & lartab
            Set rsrs = gdBase.OpenRecordset(sSQL)
            If rsrs.RecordCount = 0 Then
                HoleFreieArtikelNrab = lartab
            Else
                HoleFreieArtikelNrab = 0
            End If
        End If
    Loop
    
    rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "HoleFreieArtikelNrab"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Public Function UMS_LINRaktuell() As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset

    UMS_LINRaktuell = False
    
    sSQL = "select Lastdate from UMS_LINR   "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!LASTDATE) Then
            If DateValue(Now) = rsrs!LASTDATE Then
                UMS_LINRaktuell = True
            End If
        End If
    End If
    rsrs.Close
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "UMS_LINRaktuell"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function

Public Function UMS_ARTNRaktuell() As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset

    UMS_ARTNRaktuell = False
    
    sSQL = "select Lastdate from UMS_ARTNR   "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!LASTDATE) Then
            If DateValue(Now) = rsrs!LASTDATE Then
                UMS_ARTNRaktuell = True
            End If
        End If
    End If
    rsrs.Close
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "UMS_ARTNRaktuell"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function UMS_LPZaktuell() As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset

    UMS_LPZaktuell = False
    
    sSQL = "select Lastdate from UMS_LPZ   "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!LASTDATE) Then
            If DateValue(Now) = rsrs!LASTDATE Then
                UMS_LPZaktuell = True
            End If
        End If
    End If
    rsrs.Close
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "UMS_LPZaktuell"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Sub ErzeugeLpzUmsatz()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    Screen.MousePointer = 11

    loeschNEW "UMS_LPZ", gdBase
    CreateTableT2 "UMS_LPZ", gdBase
    
    cSQL = "Create Index PRIMKEY on UMS_LPZ(LINR,LPZ,JAHR, MONAT)"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into UMS_LPZ "
    cSQL = cSQL & "Select "
    cSQL = cSQL & " YEAR(ADATE) as JAHR"
    cSQL = cSQL & ", MONTH(ADATE) as MONAT"
    cSQL = cSQL & ", LINR"
    cSQL = cSQL & ", LPZ"
    cSQL = cSQL & ", '" & DateValue(Now) & "' as LASTDATE "
    cSQL = cSQL & ", SUM(PREIS) as UMSATZ"
    cSQL = cSQL & ", SUM(Menge * EKPR) as UMSATZSEK"
    cSQL = cSQL & ", SUM(Menge) as ABSATZ"
    cSQL = cSQL & " from KASSJOUR"
    cSQL = cSQL & " where UMS_OK = 'J' "
    cSQL = cSQL & " group by  YEAR(ADATE), MONTH(ADATE), LINR ,LPZ"
    gdBase.Execute cSQL, dbFailOnError
    
    loeschNEW "AVGLAG_LPZ", gdBase
    CreateTableT2 "AVGLAG_LPZ", gdBase
    
    cSQL = "Create Index PRIMKEY on AVGLAG_LPZ(LINR,LPZ,JAHR, MONAT)"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into AVGLAG_LPZ "
    cSQL = cSQL & "Select "
    cSQL = cSQL & " YEAR(DATUM) as JAHR"
    cSQL = cSQL & ", MONTH(DATUM) as MONAT"
    cSQL = cSQL & ", LINR"
    cSQL = cSQL & ", LPZ"
    cSQL = cSQL & ", AVG(SEK) as AVGSEK"
    cSQL = cSQL & " from LAGERLLW"
    cSQL = cSQL & " group by  YEAR(DATUM), MONTH(DATUM), LINR ,LPZ"
    gdBase.Execute cSQL, dbFailOnError
    
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    If err.Number = 3376 Or err.Number = 3372 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul3"
        Fehler.gsFunktion = "ErzeugeLpzUmsatz"
        Fehler.gsFehlertext = "Beim Erzeugen der Tabelle UMS_LINR ist ein Fehler aufgetreten."
    
        Fehlermeldung1
    End If
End Sub
Public Sub LagerwerteschreibenLPZJetzt(lblanzeige As Label, lLinr As Long)
On Error GoTo LOKAL_ERROR

    Dim sSQL        As String
    Dim rsrs        As Recordset

    Screen.MousePointer = 11
    
    loeschNEW "LAGERDLJETZT", gdBase
    loeschNEW "LAGERLLWJETZT", gdBase
    
    CreateTableT2 "LAGERLLWJETZT", gdBase
    CreateTableT2 "LAGERDLJETZT", gdBase
    
    sSQL = "insert into LAGERDLJETZT Select ARTNR,LINR,0 as LPZ,0 as EKPR, BESTAND from Artikel "
    
    If lLinr = 0 Then
    
    Else
        sSQL = sSQL & " where Linr = " & lLinr
    End If
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from LAGERDLJETZT where bestand <= 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update LAGERDLJETZT inner join Artikel on LAGERDLJETZT.ARTNR = ARTIKEL.ARTNR  "
    sSQL = sSQL & " Set LAGERDLJETZT.LPZ = ARTIKEL.LPZ "
    sSQL = sSQL & " , LAGERDLJETZT.EKPR = ARTIKEL.EKPR "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into LagerlLwJETZT Select Datevalue(now) as datum , LINR,LPZ,sum(EKPR * bestand) as SEK , sum(BESTAND) as BEST from LAGERDLJETZT "
    sSQL = sSQL & " group by LINR,LPZ "
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "LAGERDJETZT", gdBase
    
    anzeige "normal", "", lblanzeige
    
    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul8"
    Fehler.gsFunktion = "LagerwerteschreibenLINRJetzt"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Function MittelwertLugaufLPZ(lLinr As Long, lLpz As Long) As Double
    On Error GoTo LOKAL_ERROR

    Dim rec As Recordset
    Dim sSQL As String
    
    MittelwertLugaufLPZ = 0
    
    sSQL = "Select avg(lug) as mittel from ALLARTLU where Linr = " & lLinr
    sSQL = sSQL & " and lpz = " & lLpz
    Set rec = gdBase.OpenRecordset(sSQL)
    If Not rec.EOF Then
        If Not IsNull(rec!Mittel) Then
            MittelwertLugaufLPZ = rec!Mittel
        End If
    End If
    rec.Close

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "MittelwertLugaufLPZ"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function EinkaufsumsatzermittlungLPZ(cLinr As String, db As Database, iJahr As Integer, lLpz As Long) As String
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL    As String
    Dim rsrs    As Recordset
    Dim ctmp    As String
    
    If Trim$(cLinr) = "" Then
        Exit Function
    End If
    
    Screen.MousePointer = 11

    loeschNEW "EUMS" & srechnertab, db
    loeschNEW "EUMS4" & srechnertab, db
    
    cSQL = "Select artnr,LINR,0 as LPZ, BEWEGUNG , EKPR  into EUMS4" & srechnertab
    cSQL = cSQL & " from ZUGANG "
    cSQL = cSQL & " where YEAR(ADATE) = " & iJahr
    cSQL = cSQL & " and LINR = " & cLinr & " "
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Update EUMS4" & srechnertab & " inner join Artikel on"
    cSQL = cSQL & " EUMS4" & srechnertab & ".artnr =  Artikel.Artnr "
    cSQL = cSQL & " Set EUMS4" & srechnertab & ".LPZ = Artikel.lpz "
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from  EUMS4" & srechnertab
    cSQL = cSQL & " where LPZ <> " & lLpz
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Select LINR,lpz, SUM(BEWEGUNG * EKPR) as EAKJahr into EUMS" & srechnertab
    cSQL = cSQL & " from EUMS4" & srechnertab
    cSQL = cSQL & " group by LINR,lpz"
    db.Execute cSQL, dbFailOnError
    
    ctmp = "0,00"
    Set rsrs = db.OpenRecordset("EUMS" & srechnertab, dbOpenTable)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!EAKJahr) Then
            ctmp = rsrs!EAKJahr
        End If
    End If
    rsrs.Close
    
    loeschNEW "EUMS" & srechnertab, db
    loeschNEW "EUMS4" & srechnertab, db
    
    EinkaufsumsatzermittlungLPZ = ctmp
    Screen.MousePointer = 0
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "EinkaufsumsatzermittlungLPZ"
    Fehler.gsFehlertext = "Beim Ermitteln des Einkaufsumsatzes ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function EinkaufsStückermittlungLPZ(cLinr As String, db As Database, iJahr As Integer, imon As Byte, lLpz As Long) As Long
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL    As String
    Dim rsrs    As Recordset
    
    EinkaufsStückermittlungLPZ = 0
    
    If Trim$(cLinr) = "" Then
        Exit Function
    End If
    
    Screen.MousePointer = 11
    
    cSQL = "Select sum(BEWEGUNG) as maxi from ZUGANG z, Artikel a  "
    cSQL = cSQL & " where YEAR(z.ADATE) = " & iJahr
    cSQL = cSQL & " and z.artnr = a.artnr "
    cSQL = cSQL & " and z.LINR = " & cLinr & " "
    
    If lLpz > -1 Then
        cSQL = cSQL & " and a.LPZ = " & lLpz & " "
    End If
    
    If imon > 0 Then
        cSQL = cSQL & " and Month(z.adate) = " & imon
    End If
    
    
    Set rsrs = db.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!maxi) Then
            EinkaufsStückermittlungLPZ = rsrs!maxi
        End If
    End If
    rsrs.Close
    
    Screen.MousePointer = 0
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "EinkaufsStückermittlungLPZ"
    Fehler.gsFehlertext = "Beim Ermitteln des Einkaufsumsatzes ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermgesAbsatzLPZ(bymonat As Byte, iJahr As Integer, lLinr As Long, lLpz As Long) As Double
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    
    ermgesAbsatzLPZ = 0
    
    sSQL = "Select sum(ABSATZ) as maxi"
    sSQL = sSQL & " from UMS_LPZ "
    sSQL = sSQL & " where Jahr = " & iJahr
    If bymonat > 0 Then
        sSQL = sSQL & " and Monat = " & bymonat
    End If
    
    If lLinr = 0 Then
    
    Else
        sSQL = sSQL & " and LINR = " & lLinr
    End If
    
    If lLpz = 0 Then
    
    Else
        sSQL = sSQL & " and lpz = " & lLpz
    End If
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermgesAbsatzLPZ = rsrs!maxi
        End If
    End If
    rsrs.Close
                
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "ermgesAbsatzLPZ"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Function
Public Function ermgesUmsatzLpz(bymonat As Byte, iJahr As Integer, lLinr As Long, lLpz As Long) As Double
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    
    ermgesUmsatzLpz = 0
    
    sSQL = "Select sum(umsatz) as maxi"
    sSQL = sSQL & " from UMS_LPZ "
    sSQL = sSQL & " where Jahr = " & iJahr
    If bymonat > 0 Then
        sSQL = sSQL & " and Monat = " & bymonat
    End If
    
    If lLinr = 0 Then
    
    Else
        sSQL = sSQL & " and LINR = " & lLinr
    End If
    
    If lLpz = -1 Then
    
    Else
        sSQL = sSQL & " and Lpz = " & lLpz
    End If
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermgesUmsatzLpz = rsrs!maxi
        End If
    End If
    rsrs.Close
                
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "ermgesUmsatzLpz"
    Fehler.gsFehlertext = "In der Preislagenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Function
Public Function ermgesEKUmsatzLPZ(bymonat As Byte, iJahr As Integer, lLinr As Long, lLpz As Long) As Double
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    
    ermgesEKUmsatzLPZ = 0
    
    sSQL = "Select sum(umsatzsek) as maxi"
    sSQL = sSQL & " from UMS_LPZ "
    sSQL = sSQL & " where Jahr = " & iJahr
    If bymonat > 0 Then
        sSQL = sSQL & " and Monat = " & bymonat
    End If
    
    If lLinr = 0 Then
    
    Else
        sSQL = sSQL & " and LINR = " & lLinr
    End If
    
    If lLpz = -1 Then
    
    Else
        sSQL = sSQL & " and LPZ = " & lLpz
    End If
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermgesEKUmsatzLPZ = rsrs!maxi
        End If
    End If
    rsrs.Close
                
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "ermgesEKUmsatzLPZ"
    Fehler.gsFehlertext = "In der Preislagenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Function
Public Function LAGEREKermittlungJetztLPZ(lLinr As Long, lLpz As Long) As Double
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL    As String
    Dim rsrs    As Recordset
    
    If lLinr = 0 Then
        Exit Function
    End If
    
    LAGEREKermittlungJetztLPZ = 0
    
    cSQL = "Select SEK from LagerllwJETZT"
    cSQL = cSQL & " where linr = " & lLinr
    cSQL = cSQL & " and lpz = " & lLpz
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!sEK) Then
            LAGEREKermittlungJetztLPZ = rsrs!sEK
        End If
    End If
    rsrs.Close
   
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "LAGEREKermittlungJetztLPZ"
    Fehler.gsFehlertext = "Bei der Ermittlung des Lagereinkaufswertes ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
        
End Function
Public Function LAGERStückErmittlungJetztLPZ(lLinr As Long, lLpz As Long) As Long
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL    As String
    Dim rsrs    As Recordset
    
    If lLinr = 0 Then
        Exit Function
    End If
    
    LAGERStückErmittlungJetztLPZ = 0
    
    cSQL = "Select BEST from LagerllwJETZT"
    cSQL = cSQL & " where linr = " & lLinr
    cSQL = cSQL & " and lpz = " & lLpz
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!best) Then
            LAGERStückErmittlungJetztLPZ = rsrs!best
        End If
    End If
    rsrs.Close
   
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "LAGERStückErmittlungJetztLPZ"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
        
End Function
Public Function PennerEKermittlungJetztLPZ(lLinr As Long, lLpz As Long) As Double
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL    As String
    Dim rsrs    As Recordset
    
    If lLinr = 0 Then
        Exit Function
    End If
    
    PennerEKermittlungJetztLPZ = 0
    
    cSQL = "Select Top 1 datum,SEK from PENLAGERLLW "
    cSQL = cSQL & " where linr = " & lLinr
    cSQL = cSQL & " and lpz = " & lLpz
    cSQL = cSQL & " order by datum desc "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!sEK) Then
            PennerEKermittlungJetztLPZ = rsrs!sEK
        End If
    End If
    rsrs.Close
   
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "PennerEKermittlungJetztLPZ"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
        
End Function
Public Function PENNERStückErmittlungJetztLPZ(lLinr As Long, lLpz As Long) As Long
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL    As String
    Dim rsrs    As Recordset
    
    If lLinr = 0 Then
        Exit Function
    End If
    
    PENNERStückErmittlungJetztLPZ = 0
    
    cSQL = "Select TOP 1 DATUM,BEST from PENLAGERLLW "
    cSQL = cSQL & " where linr = " & lLinr
    cSQL = cSQL & " and lpz = " & lLpz
    cSQL = cSQL & " order by Datum desc "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!best) Then
            PENNERStückErmittlungJetztLPZ = rsrs!best
        End If
    End If
    rsrs.Close
   
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "PENNERStückErmittlungJetztLPZ"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
        
End Function
Public Function ermLINBEZ1(lLpz As Long, lLinr As Long) As String
On Error GoTo LOKAL_ERROR
Dim sSQL As String
Dim rsrs As Recordset

ermLINBEZ1 = ""

sSQL = "Select LINBEZEICH from LINBEZ where "
sSQL = sSQL & " LPZ  = " & lLpz
sSQL = sSQL & " and LINR  = " & lLinr

Set rsrs = gdBase.OpenRecordset(sSQL)
If Not rsrs.EOF Then
    rsrs.MoveFirst
    If Not IsNull(rsrs!LINBEZEICH) Then
        ermLINBEZ1 = rsrs!LINBEZEICH
    End If
End If
rsrs.Close

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "ermLINBEZ1"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Public Sub KBmLINR(sArtikelstatus As String, sKundenstatus As String, cLinr As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String

    
    loeschNEW "KUOKB", gdBase
    CreateTable "KUOKB", gdBase
    
    sSQL = "Insert into KUOKB Select "
    sSQL = sSQL & " KUNDBEST.ARTNR "
    sSQL = sSQL & ", KUNDBEST.BEZEICH"
    sSQL = sSQL & ", KUNDBEST.BEDNU  "
    sSQL = sSQL & ", KUNDBEST.EKPR "
    sSQL = sSQL & ", KUNDBEST.VKPR "
    sSQL = sSQL & ", KUNDBEST.MWST"
    sSQL = sSQL & ", KUNDBEST.FARBE "
    sSQL = sSQL & ", KUNDBEST.FARBTEXT "
    sSQL = sSQL & ", KUNDBEST.Filiale "
    sSQL = sSQL & ", KUNDBEST.SENDOK "
    sSQL = sSQL & ", KUNDBEST.STATUSARTIKEL "
    sSQL = sSQL & ", KUNDBEST.STATUSKUNDE "
    sSQL = sSQL & ", KUNDBEST.BESTELLTAM  "
    sSQL = sSQL & ", KUNDBEST.BESTELLTUM  "
    sSQL = sSQL & ", KUNDBEST.BESTELLTPREIS  "
    sSQL = sSQL & ", KUNDBEST.BESTELLTMENGE  "
    sSQL = sSQL & ", KUNDBEST.KUNDNR "
    sSQL = sSQL & "  from KUNDBEST , artlief where artlief.artnr = kundbest.artnr "
    sSQL = sSQL & " and Artlief.LINR = " & cLinr & " and STATUSARTIKEL = '" & sArtikelstatus & "' "
    
    If sKundenstatus <> "" Then
        sSQL = sSQL & "  and  STATUSKunde = '" & sKundenstatus & "' "
    End If
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update KUOKB inner join KUNDEN on KUOKB.KUNDNR = KUNDEN.KUNDNR"
    sSQL = sSQL & " SET KUOKB.TEL = KUNDEN.TEL "
    sSQL = sSQL & ", KUOKB.FAXNR = KUNDEN.FAXNR "
    sSQL = sSQL & ", KUOKB.EMAIL = KUNDEN.EMAIL "
    sSQL = sSQL & ", KUOKB.MOBILTEL = KUNDEN.MOBILTEL "
    sSQL = sSQL & ", KUOKB.VORNAME = KUNDEN.VORNAME "
    
    sSQL = sSQL & ", KUOKB.NAME = KUNDEN.NAME "
    sSQL = sSQL & ", KUOKB.STRASSE = KUNDEN.STRASSE "
    sSQL = sSQL & ", KUOKB.PLZ = KUNDEN.PLZ "
    sSQL = sSQL & ", KUOKB.ORT = KUNDEN.STADT "
    sSQL = sSQL & ", KUOKB.TITEL = KUNDEN.TITEL "
    sSQL = sSQL & ", KUOKB.FIRMA = KUNDEN.FIRMA "
    gdBase.Execute sSQL, dbFailOnError
      
    Select Case sArtikelstatus
    
        Case "INBESTELLUNG"
            reportbildschirm "", "aWKL77a"
        Case "BESTELLT"
            reportbildschirm "", "aWKL77b"
        Case "GELIEFERT"
            reportbildschirm "", "aWKL77c"
        Case "NICHTGELIEFERT"
            reportbildschirm "", "aWKL77d"
    End Select
    
    loeschNEW "KUOKB", gdBase
    

Exit Sub
LOKAL_ERROR:
   
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "KBmLINR"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
  
End Sub
Public Function ermoffenKUB(cLinr As String) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    
    ermoffenKUB = False
    
    cSQL = "Select KUNDBEST.ARTNR from KUNDBEST inner join ARTLIEF on KUNDBEST.artnr = ARTLIEF.ARTNR"
    cSQL = cSQL & " where ARTLIEF.LINR = " & Val(cLinr) & " and KUNDBEST.StatusARTIKEL = 'INBESTELLUNG' "
    cSQL = cSQL & "  and KUNDBEST.Sendok = false "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        ermoffenKUB = True
    End If
    rsrs.Close
    
    Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "ermoffenKUB"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
    
End Function

Public Function MittelwertLugaufLINR(lLinr As Long) As Double
    On Error GoTo LOKAL_ERROR

    Dim rec As Recordset
    Dim sSQL As String
    
    MittelwertLugaufLINR = 0
    
    sSQL = "Select avg(lug) as mittel from ALLARTLU where Linr = " & lLinr
    Set rec = gdBase.OpenRecordset(sSQL)
    If Not rec.EOF Then
        If Not IsNull(rec!Mittel) Then
            MittelwertLugaufLINR = rec!Mittel
        End If
    End If
    rec.Close

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "mittelwertLugaufLINR"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function Einkaufsumsatzermittlung(cLinr As String, db As Database, iJahr As Integer) As String
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL    As String
    Dim rsrs    As Recordset
    Dim ctmp    As String
    
    If Trim$(cLinr) = "" Then
        Exit Function
    End If
    
    Screen.MousePointer = 11

    loeschNEW "EUMS" & srechnertab, db
    loeschNEW "EUMS4" & srechnertab, db
    
    cSQL = "Select LINR, BEWEGUNG , REK  into EUMS4" & srechnertab
'    cSQL = "Select LINR, BEWEGUNG , EKPR  into EUMS4" & srechnertab
    cSQL = cSQL & " from ZUGANG "
    cSQL = cSQL & " where YEAR(ADATE) = " & iJahr
    cSQL = cSQL & " and LINR = " & cLinr & " "
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Select LINR, SUM(BEWEGUNG * REK) as EAKJahr into EUMS" & srechnertab
'    cSQL = "Select LINR, SUM(BEWEGUNG * EKPR) as EAKJahr into EUMS" & srechnertab
    cSQL = cSQL & " from EUMS4" & srechnertab
    cSQL = cSQL & " group by LINR"
    db.Execute cSQL, dbFailOnError
    
    ctmp = "0,00"
    Set rsrs = db.OpenRecordset("EUMS" & srechnertab, dbOpenTable)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!EAKJahr) Then
            ctmp = rsrs!EAKJahr
        End If
    End If
    rsrs.Close
    
    loeschNEW "EUMS" & srechnertab, db
    loeschNEW "EUMS4" & srechnertab, db
    
    Einkaufsumsatzermittlung = ctmp
    Screen.MousePointer = 0
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "Einkaufsumsatzermittlung"
    Fehler.gsFehlertext = "Beim Ermitteln des Einkaufsumsatzes ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function EinkaufsumsatzermittlungArtikel(cArtNr As String, db As Database, iJahr As Integer) As String
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL    As String
    Dim rsrs    As Recordset
    Dim ctmp    As String
    
    If Trim$(cArtNr) = "" Then
        Exit Function
    End If
    
    Screen.MousePointer = 11

    EinkaufsumsatzermittlungArtikel = "0,00"
    
    cSQL = "Select SUM(BEWEGUNG * EKPR) as EAKJahr "
    cSQL = cSQL & " from ZUGANG "
    cSQL = cSQL & " where YEAR(ADATE) = " & iJahr
    cSQL = cSQL & " and ARTNR = " & cArtNr & " "
    
    Set rsrs = db.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!EAKJahr) Then
            EinkaufsumsatzermittlungArtikel = rsrs!EAKJahr
        End If
    End If
    rsrs.Close
    
    Screen.MousePointer = 0
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "EinkaufsumsatzermittlungArtikel"
    Fehler.gsFehlertext = "Beim Ermitteln des Einkaufsumsatzes ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermgesUmsatzLinr(bymonat As Byte, iJahr As Integer, lLinr As Long) As Double
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    
    ermgesUmsatzLinr = 0
    
    sSQL = "Select sum(umsatz) as maxi"
    sSQL = sSQL & " from UMS_LINR "
    sSQL = sSQL & " where Jahr = " & iJahr
    If bymonat > 0 Then
        sSQL = sSQL & " and Monat = " & bymonat
    End If
    
    If lLinr = 0 Then
    
    Else
        sSQL = sSQL & " and LINR = " & lLinr
    End If
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermgesUmsatzLinr = rsrs!maxi
        End If
    End If
    rsrs.Close
                
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "ermgesUmsatzLinr"
    Fehler.gsFehlertext = "In der Preislagenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Function
Public Function ermgesUmsatzARTnr(bymonat As Byte, iJahr As Integer, cArtNr As String) As Double
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    
    ermgesUmsatzARTnr = 0
    
    sSQL = "Select sum(umsatz) as maxi"
    sSQL = sSQL & " from UMS_ARTNR" & srechnertab
    sSQL = sSQL & " where Jahr = " & iJahr
    If bymonat > 0 Then
        sSQL = sSQL & " and Monat = " & bymonat
    End If
    
    If cArtNr = "" Then
    
    Else
        sSQL = sSQL & " and artnr = " & cArtNr
    End If
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermgesUmsatzARTnr = rsrs!maxi
        End If
    End If
    rsrs.Close
                
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "ermgesUmsatzARTnr"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Function
Public Function ermgesVKMengeARTnr(bymonat As Byte, iJahr As Integer, cArtNr As String) As Long
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    
    ermgesVKMengeARTnr = 0
    
    sSQL = "Select sum(vkmenge) as maxi"
    sSQL = sSQL & " from UMS_ARTNR" & srechnertab
    sSQL = sSQL & " where Jahr = " & iJahr
    If bymonat > 0 Then
        sSQL = sSQL & " and Monat = " & bymonat
    End If
    
    If cArtNr = "" Then
    
    Else
        sSQL = sSQL & " and artnr = " & cArtNr
    End If
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermgesVKMengeARTnr = rsrs!maxi
        End If
    End If
    rsrs.Close
                
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "ermgesVKMengeARTnr"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Function
Public Function ermgesEKUmsatzLinr(bymonat As Byte, iJahr As Integer, lLinr As Long) As Double
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    
    ermgesEKUmsatzLinr = 0
    
    sSQL = "Select sum(umsatzsek) as maxi"
    sSQL = sSQL & " from UMS_LINR "
    sSQL = sSQL & " where Jahr = " & iJahr
    If bymonat > 0 Then
        sSQL = sSQL & " and Monat = " & bymonat
    End If
    
    If lLinr = 0 Then
    
    Else
        sSQL = sSQL & " and LINR = " & lLinr
    End If
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermgesEKUmsatzLinr = rsrs!maxi
        End If
    End If
    rsrs.Close
                
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "ermgesEKUmsatzLinr"
    Fehler.gsFehlertext = "In der Preislagenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Function
Public Function ermgesEKUmsatzARTNR(bymonat As Byte, iJahr As Integer, cArtNr As String) As Double
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    
    ermgesEKUmsatzARTNR = 0
    
    sSQL = "Select sum(umsatzsek) as maxi"
    sSQL = sSQL & " from UMS_ARTNR" & srechnertab
    sSQL = sSQL & " where Jahr = " & iJahr
    If bymonat > 0 Then
        sSQL = sSQL & " and Monat = " & bymonat
    End If
    
    If cArtNr = "" Then
    
    Else
        sSQL = sSQL & " and Artnr = " & cArtNr
    End If
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermgesEKUmsatzARTNR = rsrs!maxi
        End If
    End If
    rsrs.Close
                
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "ermgesEKUmsatzARTNR"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Function
Public Function LAGEREKermittlungJetzt(lLinr As Long) As Double
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL    As String
    Dim rsrs    As Recordset
    
    If lLinr = 0 Then
        Exit Function
    End If
    
    LAGEREKermittlungJetzt = 0
    
    cSQL = "Select SEK from LagerlwJETZT"
    cSQL = cSQL & " where linr = " & lLinr
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!sEK) Then
            LAGEREKermittlungJetzt = rsrs!sEK
        End If
    End If
    rsrs.Close
   
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "LAGEREKermittlungJetzt"
    Fehler.gsFehlertext = "Bei der Ermittlung des Lagereinkaufswertes ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
        
End Function
Public Function LAGEREKermittlungJetztARTNR(cArtNr As String) As Double
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL    As String
    Dim rsrs    As Recordset
    
    If cArtNr = "" Then
        Exit Function
    End If
    
    LAGEREKermittlungJetztARTNR = 0
    
    cSQL = "Select Bestand * EKPR as SEK from ARTIKEL"
    cSQL = cSQL & " where artnr = " & cArtNr
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!sEK) Then
            LAGEREKermittlungJetztARTNR = rsrs!sEK
        End If
    End If
    rsrs.Close
   
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "LAGEREKermittlungJetztARTNR"
    Fehler.gsFehlertext = "Bei der Ermittlung des Lagereinkaufswertes ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
        
End Function
Public Function LAGERStückErmittlungJetzt(lLinr As Long) As Long
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL    As String
    Dim rsrs    As Recordset
    
    If lLinr = 0 Then
        Exit Function
    End If
    
    LAGERStückErmittlungJetzt = 0
    
    cSQL = "Select BEST from LagerlwJETZT"
    cSQL = cSQL & " where linr = " & lLinr
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!best) Then
            LAGERStückErmittlungJetzt = rsrs!best
        End If
    End If
    rsrs.Close
   
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "LAGERStückErmittlungJetzt"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
        
End Function

Public Function PennerEKermittlungJetzt(lLinr As Long) As Double
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL    As String
    Dim rsrs    As Recordset
    
    If lLinr = 0 Then
        Exit Function
    End If
    
    PennerEKermittlungJetzt = 0
    
    cSQL = "Select top 1 datum, SEK from PenLagerlw"
    cSQL = cSQL & " where linr = " & lLinr
    cSQL = cSQL & " order by datum desc"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!sEK) Then
            PennerEKermittlungJetzt = rsrs!sEK
        End If
    End If
    rsrs.Close
   
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "PennerEKermittlungJetzt"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
        
End Function
Public Function PennerStückErmittlungJetzt(lLinr As Long) As Long
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL    As String
    Dim rsrs    As Recordset
    
    If lLinr = 0 Then
        Exit Function
    End If
    
    PennerStückErmittlungJetzt = 0
    
    cSQL = "Select top 1 datum, BEST from PenLagerlw"
    cSQL = cSQL & " where linr = " & lLinr
    cSQL = cSQL & " order by datum desc"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!best) Then
            PennerStückErmittlungJetzt = rsrs!best
        End If
    End If
    rsrs.Close
   
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "PennerStückErmittlungJetzt"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
        
End Function
Public Function ermdatSAP() As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset

    ermdatSAP = False
    
    If NewTableSuchenDBKombi("SAP", gdBase) Then
        sSQL = "select Datum from SAP"
        Set rsrs = gdBase.OpenRecordset(sSQL)
        If Not rsrs.EOF Then
            If Not IsNull(rsrs!Datum) Then
                If DateValue(Now) > rsrs!Datum Then
                    ermdatSAP = True
                End If
            End If
        End If
        rsrs.Close
    End If
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "ermdatSAP"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function

Public Sub ErzeugeLinrUmsatz()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    Screen.MousePointer = 11

    loeschNEW "UMS_LINR", gdBase
    CreateTableT2 "UMS_LINR", gdBase
    
    cSQL = "Create Index PRIMKEY on UMS_LINR(LINR,JAHR, MONAT)"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into UMS_LINR "
    cSQL = cSQL & "Select "
    cSQL = cSQL & " YEAR(ADATE) as JAHR"
    cSQL = cSQL & ", MONTH(ADATE) as MONAT"
    cSQL = cSQL & ", LINR"
    cSQL = cSQL & ", '" & DateValue(Now) & "' as LASTDATE "
    cSQL = cSQL & ", SUM(PREIS) as UMSATZ"
    cSQL = cSQL & ", SUM(Menge * EKPR) as UMSATZSEK"
    cSQL = cSQL & ", SUM(Menge) as ABSATZ"
    cSQL = cSQL & " from KASSJOUR"
    cSQL = cSQL & " where UMS_OK = 'J' "
    cSQL = cSQL & " group by  YEAR(ADATE), MONTH(ADATE), LINR "
    gdBase.Execute cSQL, dbFailOnError
    
    loeschNEW "AVGLAG_LINR", gdBase
    CreateTableT2 "AVGLAG_LINR", gdBase
    
    cSQL = "Create Index PRIMKEY on AVGLAG_LINR(LINR,JAHR, MONAT)"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into AVGLAG_LINR "
    cSQL = cSQL & "Select "
    cSQL = cSQL & " YEAR(DATUM) as JAHR"
    cSQL = cSQL & ", MONTH(DATUM) as MONAT"
    cSQL = cSQL & ", LINR"
    cSQL = cSQL & ", AVG(SEK) as AVGSEK"
    cSQL = cSQL & " from LAGERLW"
    cSQL = cSQL & " group by  YEAR(DATUM), MONTH(DATUM), LINR "
    gdBase.Execute cSQL, dbFailOnError
    
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    If err.Number = 3376 Or err.Number = 3372 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul3"
        Fehler.gsFunktion = "ErzeugeLinrUmsatz"
        Fehler.gsFehlertext = "Beim Erzeugen der Tabelle UMS_LINR ist ein Fehler aufgetreten."
    
        Fehlermeldung1
    End If
End Sub
Public Sub ErzeugeArtnrUmsatz()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    Screen.MousePointer = 11

    loeschNEW "UMS_ARTNR" & srechnertab, gdBase
    
    cSQL = "Create Table UMS_ARTNR" & srechnertab
    cSQL = cSQL & "( "
    cSQL = cSQL & " JAHR integer"
    cSQL = cSQL & ", MONAT integer"
    cSQL = cSQL & ", ARTNR long"
    cSQL = cSQL & ", LASTDATE DATETIME"
    cSQL = cSQL & ", VKMENGE long"
    cSQL = cSQL & ", UMSATZ double"
    cSQL = cSQL & ", UMSATZSEK double"
    cSQL = cSQL & ") "
    gdBase.Execute cSQL, dbFailOnError
    
'    CreateTableT2 "UMS_ARTNR", gdBase
    
    cSQL = "Create Index PRIMKEY on UMS_ARTNR" & srechnertab & "(ARTNR,JAHR, MONAT)"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into UMS_ARTNR" & srechnertab
    cSQL = cSQL & " Select "
    cSQL = cSQL & " YEAR(ADATE) as JAHR"
    cSQL = cSQL & ", MONTH(ADATE) as MONAT"
    cSQL = cSQL & ", KASSJOUR.ARTNR"
    cSQL = cSQL & ", '" & DateValue(Now) & "' as LASTDATE "
    cSQL = cSQL & ", SUM(MENGE) as VKMENGE"
    cSQL = cSQL & ", SUM(PREIS) as UMSATZ"
    cSQL = cSQL & ", SUM(Menge * KASSJOUR.EKPR) as UMSATZSEK"
    cSQL = cSQL & " from KASSJOUR inner join Top" & srechnertab & " on KASSJOUR.ARTNR = Top" & srechnertab & ".Artnr  "
    cSQL = cSQL & " where UMS_OK = 'J' "
    
'    cSQL = cSQL & " and Artnr in (Select Artnr from Top " & srechnertab & " )"
    
    
    
    
    cSQL = cSQL & " group by  YEAR(ADATE), MONTH(ADATE), KASSJOUR.ARTNR "
    gdBase.Execute cSQL, dbFailOnError
    
    
    loeschNEW "Last_VK" & srechnertab, gdBase
    
    cSQL = "Select Kassjour.Artnr, Max(adate) as LASTVK into Last_VK" & srechnertab
    cSQL = cSQL & " from KASSJOUR inner join Top" & srechnertab & " on KASSJOUR.ARTNR = Top" & srechnertab & ".Artnr  "
    cSQL = cSQL & " group by  KASSJOUR.ARTNR "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update Top" & srechnertab & " inner join Last_VK" & srechnertab & " on Top" & srechnertab & ".Artnr = Last_VK" & srechnertab & ".Artnr "
    cSQL = cSQL & " Set Top" & srechnertab & ".LVK = Last_VK" & srechnertab & ".LASTVK  "
    gdBase.Execute cSQL, dbFailOnError
    
    loeschNEW "Last_VK" & srechnertab, gdBase
    
    loeschNEW "LAGER_EK" & srechnertab, gdBase
    
    cSQL = "Select Artikel.Artnr, Artikel.Bestand * Artikel.EKPR as SEK into LAGER_EK" & srechnertab
    cSQL = cSQL & " from Artikel inner join Top" & srechnertab & " on Artikel.ARTNR = Top" & srechnertab & ".Artnr  "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update Top" & srechnertab & " inner join LAGER_EK" & srechnertab & " on Top" & srechnertab & ".Artnr = LAGER_EK" & srechnertab & ".Artnr "
    cSQL = cSQL & " Set Top" & srechnertab & ".LAGERWSEK = LAGER_EK" & srechnertab & ".SEK  "
    gdBase.Execute cSQL, dbFailOnError
    
    loeschNEW "LAGER_EK" & srechnertab, gdBase
    
    
    
    
    
    loeschNEW "Last_ZU" & srechnertab, gdBase
    
    cSQL = "Select Zugang.Artnr, Max(adate) as LASTZU into Last_ZU" & srechnertab
    cSQL = cSQL & " from Zugang inner join Top" & srechnertab & " on Zugang.ARTNR = Top" & srechnertab & ".Artnr  "
    cSQL = cSQL & " group by  Zugang.ARTNR "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update Top" & srechnertab & " inner join Last_ZU" & srechnertab & " on Top" & srechnertab & ".Artnr = Last_ZU" & srechnertab & ".Artnr "
    cSQL = cSQL & " Set Top" & srechnertab & ".LWE = Last_ZU" & srechnertab & ".LASTZU  "
    gdBase.Execute cSQL, dbFailOnError
    
    loeschNEW "Last_ZU" & srechnertab, gdBase
    
    
    
    
    
    
    

    
    
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    If err.Number = 3376 Or err.Number = 3372 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul3"
        Fehler.gsFunktion = "ErzeugeArtnrUmsatz"
        Fehler.gsFehlertext = "Beim Erzeugen der Tabelle UMS_ARTNR ist ein Fehler aufgetreten."
    
        Fehlermeldung1
    End If
End Sub
Public Sub LagerwerteschreibenLINRJetzt(lblanzeige As Label)
On Error GoTo LOKAL_ERROR

    Dim sSQL        As String
    Dim rsrs        As Recordset

    Screen.MousePointer = 11
    
    loeschNEW "LAGERDJETZT", gdBase
    loeschNEW "LagerlwJETZT", gdBase
    
    CreateTableT2 "LAGERLWJETZT", gdBase
    CreateTableT2 "LAGERDJETZT", gdBase
    
    sSQL = "insert into LAGERDJETZT Select ARTNR,LINR,0 as EKPR,  BESTAND from Artikel "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from LAGERDJETZT where bestand <= 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update LAGERDJETZT inner join Artikel on LAGERDJETZT.ARTNR = ARTIKEL.ARTNR  "
    sSQL = sSQL & " set LAGERDJETZT.EKPR = ARTIKEL.EKPR "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into LagerlwJETZT Select Datevalue(now) as datum , LINR,sum(EKPR * bestand) as SEK , sum(BESTAND) as BEST from LAGERDJETZT "
    sSQL = sSQL & " group by LINR "
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "LAGERDJETZT", gdBase
    
    anzeige "normal", "", lblanzeige
    
    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "LagerwerteschreibenLINRJetzt"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Function fnBildeSQLZENartFMb(sTab As String, sEAN As String, sBez As String, sLinr As String, _
sLPZ As String, slibesnr As String, sAGN As String, iOrder As Byte, sawm As String, _
Listx As ListBox, mitLU As Boolean, sPGN As String, sMARKE As String, brk As Boolean) As String
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim bAnd        As Boolean
    Dim ctmp        As String
    Dim cArtNr      As String
    Dim i           As Integer
    Dim sSQL        As String
    Dim lGBest      As Long
    Dim lBestO      As Long
    Dim lgMinBest   As Long
    Dim rec         As Recordset
    Dim rsartt      As Recordset
    Dim cPfad       As String
    Dim cPfad2      As String
    Dim siEkpr      As Double
    
    fnBildeSQLZENartFMb = ""
    
       
    
    bAnd = False
    
    
    loeschNEW sTab, gdApp
    loeschNEW sTab, gdBase
    
    
    cSQL = "Create Table " & sTab
    cSQL = cSQL & "(ARTNR Long "
    cSQL = cSQL & ", BEZEICH Text(35) "
    cSQL = cSQL & ", LINR long "
    cSQL = cSQL & ", LPZ double "
    cSQL = cSQL & ", LIBESNR Text(13) "
    cSQL = cSQL & ", LEKPR double "
    cSQL = cSQL & ", VKPR double "
    cSQL = cSQL & ", KVKPR1 double "
    cSQL = cSQL & ", EKPR double "
    cSQL = cSQL & ", MINBEST double "
    cSQL = cSQL & ", GEFUEHRT Text(1) "
    cSQL = cSQL & ", RABATT_OK Text(1) "
    cSQL = cSQL & ", PREISSCHU Text(1) "
    cSQL = cSQL & ", NOTIZEN Text(40) "
    cSQL = cSQL & ", AGN double "
    cSQL = cSQL & ", RKZ Text(1) "
    cSQL = cSQL & ", EAN Text(13) "
    cSQL = cSQL & ", EAN2 Text(13) "
    cSQL = cSQL & ", EAN3 Text(13) "
    cSQL = cSQL & ", MINMEN double "
    cSQL = cSQL & ", MWST Text(1) "
    cSQL = cSQL & ", ETIMERK Text(1) "
    cSQL = cSQL & ", AWM Text(2)"
    cSQL = cSQL & ", UMS_OK Text(1) "
    cSQL = cSQL & ", FARBNR double "
    cSQL = cSQL & ", MOPREIS double "
    cSQL = cSQL & ", BESTAND double "
    cSQL = cSQL & ", BESTO double "
    cSQL = cSQL & ", VKMENGE double "
    cSQL = cSQL & ", VKDATUM DATETIME "
    cSQL = cSQL & ", INHALT double "
    cSQL = cSQL & ", INHALTBEZ TEXT(3) "
    cSQL = cSQL & ", GRUNDPREIS TEXT(1) "
    cSQL = cSQL & ", BONUS_OK TEXT(1) "
    cSQL = cSQL & ", LASTDATE DATETIME "
    cSQL = cSQL & ", LASTTIME Text(10) "
    cSQL = cSQL & ", AUFDAT DATETIME "
    cSQL = cSQL & ", EXDAT DATETIME "
    cSQL = cSQL & ", MARKE Text(20) "
    cSQL = cSQL & ", GROESSE text(10) "
    cSQL = cSQL & ", SPANNE double "
    cSQL = cSQL & ", AUFSCHLAG double "
    cSQL = cSQL & ", SYNSTATUS Text(1) "
    cSQL = cSQL & ", LFNR Autoincrement "
    cSQL = cSQL & ", MINBESTN double "
    cSQL = cSQL & ", LAGERU double "
    cSQL = cSQL & ", PGN BYTE "
    
    cSQL = cSQL & ", LAGERP Long "
    cSQL = cSQL & ", ETIKETT  TEXT(1) "
    
    cSQL = cSQL & " ) "
    gdBase.Execute cSQL, dbFailOnError
    

    cSQL = "Insert into " & sTab & " Select distinct Artikel.* from ARTIKEL  inner join Artlief  on ARTIKEL.artnr = ARTLIEF.ARTNR where "
'    cSQL = cSQL & " Artikel.ARTNR = B.ARTNR "
    'ArtNr oder EAN
    ctmp = Trim$(sEAN)
    
    If ctmp <> "" Then
        If Len(ctmp) <= 6 Then
            'KISS-ArtNr
            cSQL = cSQL & "Artikel.ARTNR = " & ctmp & " "
            bAnd = True
            
        ElseIf Len(ctmp) = 8 Then
            'KISS-ArtNr als Barcode oder echter EAN-8
            If Left$(ctmp, 1) = "2" Or Left$(ctmp, 1) = "0" Then
                ctmp = Mid$(ctmp, 2, 6)
                cSQL = cSQL & "Artikel.ARTNR = " & ctmp & " "
                bAnd = True
            Else
                cSQL = cSQL & "( Artikel.EAN = '" & ctmp & "' "
                cSQL = cSQL & "or Artikel.EAN2 = '" & ctmp & "' "
                cSQL = cSQL & "or Artikel.EAN3 = '" & ctmp & "' ) "
                bAnd = True
            End If
        Else
            'Irgendwas anderes für die EAN-Felder
            cSQL = cSQL & "( Artikel.EAN = '" & ctmp & "' "
            cSQL = cSQL & "or Artikel.EAN2 = '" & ctmp & "' "
            cSQL = cSQL & "or Artikel.EAN3 = '" & ctmp & "' ) "
            bAnd = True
        End If
    End If
    
    'Artikelbezeichnung
    If InStr(sBez, "#") > 0 Then
        sBez = "*[#]"
    End If
    
    ctmp = Trim$(sBez)
    If ctmp <> "" Then
        If bAnd Then
            cSQL = cSQL & "and "
        End If
        cSQL = cSQL & "Artikel.BEZEICH like '" & ctmp & "*' "
        bAnd = True
    End If
    
    'LiefNr
    ctmp = Trim$(sLinr)
    If ctmp <> "" Then
        If bAnd Then
            cSQL = cSQL & "and "
        End If
        cSQL = cSQL & "Artlief.LINR = " & ctmp & " "
        bAnd = True
    End If
    
    'Linie
    
    If Listx.Visible = True And Listx.ListCount > 0 Then
        If bAnd Then
            cSQL = cSQL & "and "
        End If
    
        cSQL = cSQL & "( Artikel.lpz=" & Mid$(Listx.list(0), 1, InStr(1, Listx.list(0), " "))
        For i = 1 To Listx.ListCount - 1
            cSQL = cSQL & " or Artikel.lpz=" & Mid$(Listx.list(i), 1, InStr(1, Listx.list(i), " "))
        Next i
        cSQL = cSQL & " ) "
        bAnd = True
        
    Else
        
        'Linie
        ctmp = Trim$(sLPZ)
        If ctmp <> "" Then
            If bAnd Then
                cSQL = cSQL & "and "
            End If
            cSQL = cSQL & "Artikel.LPZ = " & ctmp & " "
            bAnd = True
        End If
        
    End If
    
    'LiefBestNr
    ctmp = Trim$(slibesnr)
    If ctmp <> "" Then
        If bAnd Then
            cSQL = cSQL & "and "
        End If
        cSQL = cSQL & "Artikel.LIBESNR like '" & ctmp & "*' "
        bAnd = True
    End If
    
    'AGN
    ctmp = Trim$(sAGN)
    If ctmp <> "" Then
        If bAnd Then
            cSQL = cSQL & "and "
        End If
        cSQL = cSQL & "Artikel.AGN = " & ctmp & " "
        bAnd = True
    End If
    
    'Marke
    ctmp = Trim$(sMARKE)
    If ctmp <> "" Then
        If LoeseMarkenInArtnr(ctmp) Then
            If bAnd Then
                cSQL = cSQL & "and "
            End If
            cSQL = cSQL & " Artikel.artnr in(Select artnr from MY" & srechnertab & ")"
            bAnd = True
        End If
    End If
    
    If mitLU Then
    
        If bAnd Then
            cSQL = cSQL & "and "
        End If
        cSQL = cSQL & " Artikel.artnr in(Select artnr from MBORDER )"
        bAnd = True
    End If

    
    'AGN
    ctmp = Trim$(sPGN)
    If ctmp <> "" Then
        If bAnd Then
            cSQL = cSQL & "and "
        End If
        cSQL = cSQL & "Artikel.PGN = " & ctmp & " "
        bAnd = True
    End If
    
    ctmp = Trim$(sawm)
    If ctmp <> "" Then
        If bAnd Then
            cSQL = cSQL & "and "
        End If
        cSQL = cSQL & "Artikel.AWM = '" & ctmp & "' "
        bAnd = True
    End If
    
'    MussRKZ
    If brk Then
        If bAnd Then
            cSQL = cSQL & "and "
        End If
        cSQL = cSQL & "Artikel.RKZ = 'J' "
        bAnd = True
    End If
    
    If bAnd Then
        cSQL = cSQL & "and "
    End If
    
    cSQL = cSQL & " ( Artikel.SYNSTATUS is null or Artikel.SYNSTATUS = 'E' or Artikel.SYNSTATUS = 'A' )"
    
    bAnd = True
    
'    If bAnd Then
'        cSQL = cSQL & "and "
'    End If

'    cSQL = cSQL & " ARTLIEF.SYNSTATUS <> 'D' "
    
    bAnd = True
    
    Dim corder As String
    
    If iOrder = 1 Then
        corder = " order by Artikel.LINR, Artikel.LPZ, Artikel.BEZEICH, Artikel.ARTNR "
    ElseIf iOrder = 2 Then
        corder = " order by Artikel.BEZEICH, Artikel.ARTNR "
    ElseIf iOrder = 3 Then
        corder = " order by Artikel.AWM desc , Artikel.BEZEICH, Artikel.ARTNR "
    ElseIf iOrder = 4 Then
        corder = " order by Artikel.AWM desc , Artikel.LINR, Artikel.LPZ, Artikel.BEZEICH, Artikel.ARTNR "
    End If
    cSQL = cSQL & corder
'    MsgBox cSQL
    gdBase.Execute cSQL, dbFailOnError
    

    cPfad2 = gcDBPfad
    If Right$(cPfad2, 1) <> "\" Then
        cPfad2 = cPfad2 & "\"
    End If
    
    cSQL = "select * into " & sTab & " from " & sTab & " in '" & cPfad2 & "kissdata.MDB'"
    gdApp.Execute cSQL, dbFailOnError
    
    sSQL = " Create index  ARTNR on " & sTab & "(ARTNR) "
    gdApp.Execute sSQL, dbFailOnError
   
    Dim lMax As Long
    Dim iRet As Integer
    
    Set rsartt = gdApp.OpenRecordset(sTab, dbOpenTable)
    If Not rsartt.EOF Then
        lMax = rsartt.RecordCount
        If lMax > 1000 Then
            iRet = MsgBox("Uppss..." & vbCrLf & "Es wurden mehr als 1000 Datensätze gefunden.(" & lMax & ")" & vbCrLf & "Wirklich anzeigen?", vbQuestion + vbYesNo, "DATENVOLUMEN")
            If iRet = vbNo Then
                fnBildeSQLZENartFMb = ""
                Exit Function
            End If
        End If
        rsartt.MoveFirst
        Do While Not rsartt.EOF
            If Not IsNull(rsartt!artnr) Then
                cArtNr = rsartt!artnr
            End If
            
            
            If Not IsNull(rsartt!ekpr) Then
                siEkpr = rsartt!ekpr
            Else
                siEkpr = 0
            End If
             
            
            
'
'            If Not IsNull(rsartt!linr) Then
'                If Not IsNull(rsartt!LPZ) Then
'                    rsartt!MARKE = ermittleMarke(rsartt!linr, rsartt!LPZ)
'                End If
'            End If
            
        rsartt.MoveNext
        Loop
    End If
    rsartt.Close
      
    fnBildeSQLZENartFMb = "Select * from " & sTab & " order by lfnr"
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "fnBildeSQLZENartFMb"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Public Sub FaerbenFlexHKunde(ctmp As String, mshflex As MSHFlexGrid, iSpalte As Integer, lrow As Long)
    On Error GoTo LOKAL_ERROR
    
    With mshflex

        If ctmp <> "0" Then
            .Row = lrow
            .Col = iSpalte
            If ctmp = "1" Then
                .CellBackColor = glfarbe(1)
                .CellForeColor = vbBlack
            ElseIf ctmp = "2" Then
                .CellBackColor = glfarbe(2)
                .CellForeColor = vbBlack
            ElseIf ctmp = "3" Then
                .CellBackColor = glfarbe(3)
                .CellForeColor = vbBlack
            ElseIf ctmp = "4" Then
                .CellBackColor = glfarbe(4)
                .CellForeColor = vbBlack
            ElseIf ctmp = "5" Then
                .CellBackColor = glfarbe(5)
                .CellForeColor = vbBlack
            ElseIf ctmp = "6" Then
                .CellBackColor = glfarbe(6)
                .CellForeColor = vbBlack
            ElseIf ctmp = "7" Then
                .CellBackColor = glfarbe(7)
                .CellForeColor = vbBlack
            ElseIf ctmp = "8" Then
                .CellBackColor = glfarbe(8)
                .CellForeColor = vbBlack
            ElseIf ctmp = "9" Then
                .CellBackColor = glfarbe(9)
                .CellForeColor = vbBlack
            ElseIf ctmp = "99" Then         'eben angefügte Artikel
                .CellBackColor = vbWhite
                .CellForeColor = vbBlue
            ElseIf ctmp = "98" Then         'neue Artikel
                .CellBackColor = vbWhite
                .CellForeColor = vbRed
            ElseIf ctmp = "97" Then
                .CellBackColor = vbYellow   'automatisch kalkulierte
                .CellForeColor = vbBlue
            ElseIf ctmp = "96" Then
                .CellBackColor = vbWhite   'doppelte in Anzeige
                .CellForeColor = vbGreen
            ElseIf ctmp = "95" Then         'nicht geliefert
                .CellBackColor = vbBlue
                .CellForeColor = vbBlack
            ElseIf ctmp = "94" Then         'Preisaktion in Vorbereitung
                .CellBackColor = glfarbe(0)
                .CellForeColor = vbBlue
            ElseIf ctmp = "93" Then         'Preisaktion aktiv
                .CellBackColor = vbWhite
                .CellForeColor = vbGreen
            ElseIf ctmp = "92" Then         'seit 2 Jahren oder noch nie verkauft
                .CellBackColor = &H80000012 'Black
                .CellForeColor = vbWhite
            Else
                .CellBackColor = glfarbe(0)
                .CellForeColor = vbBlack
            End If
        Else
            .Col = iSpalte
            .CellBackColor = glfarbe(0)
            .CellForeColor = vbBlack
        End If
        
    End With
    
    Exit Sub
LOKAL_ERROR:

    

        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul3"
        Fehler.gsFunktion = "FaerbenFlexHKunde"
        Fehler.gsFehlertext = "Beim Faerben der Tabelle ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    

End Sub
Public Sub schreibeDAOtxt()
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr     As Integer
    Dim lPos        As Long
    Dim lWert       As Long
    Dim sTime       As String
    Dim ctmp        As String
    Dim cZeil       As String
    Dim cZeile2     As String
    Dim cPfad       As String
    Dim cdatei      As String
    Dim cSatz       As String
    
    sTime = TimeValue(Now)
    sTime = Right(sTime, 8)
    
    cPfad = App.Path
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If

    lWert = DateValue(Now)
    ctmp = Format$(lWert, "DD.MM.YY")
    
    cZeil = ""
    cZeil = ctmp & Space(1) & sTime & Space(2) & vbCrLf
    cSatz = cPfad & "DAO.TXT"
    
    iFileNr = FreeFile
    Open cSatz For Binary As #iFileNr
    
        lPos = LOF(iFileNr)
        lPos = lPos + 1
        Put #iFileNr, lPos, cZeil
    Close iFileNr
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul3"
        Fehler.gsFunktion = "schreibeDAOtxt"
        Fehler.gsFehlertext = "Beim Erstellen des allgemeinen Protokolls ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Public Sub ZeigArtHistInList(sArt As String, Listx As ListBox, sarti As String, sOrder As String)
On Error GoTo LOKAL_ERROR

Dim sSQL        As String
Dim sSatz       As String
Dim rsrs        As Recordset
Dim cBez As String
Dim cLiefBez As String



Screen.MousePointer = 11
Listx.Clear
Listx.Visible = False

If UCase$(sArt) = "EINKAUF" Then
    If Trim(sarti) = "" Then
        Screen.MousePointer = 0
        Exit Sub
    End If

    sSQL = "select * from Zugang where artnr = " & sarti
    sSQL = sSQL & " order by  " & sOrder
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
        If Not IsNull(rsrs!ADATE) Then
            sSatz = Format(rsrs!ADATE, "DD.MM.YY") & Space(10 - Len(Format(rsrs!ADATE, "DD.MM.YY")))
        End If
        
        If Not IsNull(rsrs!Uhrzeit) Then
            sSatz = sSatz & Format(rsrs!Uhrzeit, "HH:MM") & Space(8 - Len(Format(rsrs!Uhrzeit, "HH:MM")))
        Else
            sSatz = sSatz & Space(8)
        End If
        
        If Not IsNull(rsrs!BEWEGUNG) Then
            sSatz = sSatz & rsrs!BEWEGUNG & Space(7 - Len(rsrs!BEWEGUNG))
        Else
            sSatz = sSatz & "0" & Space(6)
        End If
        
        If Not IsNull(rsrs!bestandalt) Then
            sSatz = sSatz & rsrs!bestandalt & Space(6 - Len(rsrs!bestandalt))
        Else
            sSatz = sSatz & Space(6)
        End If
        
        If Not IsNull(rsrs!BESTANDneu) Then
            sSatz = sSatz & rsrs!BESTANDneu & Space(7 - Len(rsrs!BESTANDneu))
        Else
            sSatz = sSatz & Space(7)
        End If
        
        If Not IsNull(rsrs!linr) Then
            cLiefBez = ermLiefBez(CLng(rsrs!linr))
            sSatz = sSatz & rsrs!linr & Space(7 - Len(rsrs!linr)) & " " & cLiefBez & Space(36 - Len(cLiefBez))
        Else
            sSatz = sSatz & Space(44)
        End If
        
    
        
        If Not IsNull(rsrs!rek) Then
            sSatz = sSatz & Format$(rsrs!rek, "####0.00") & Space(14 - Len(Format$(rsrs!rek, "####0.00")))
        Else
            sSatz = sSatz & Space(14)
        End If
        
        Listx.AddItem sSatz
        
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
ElseIf UCase$(sArt) = "BON" Then

    

    If NewTableSuchenDBKombi("KAT" & srechnertab, gdBase) = False Then
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    sSQL = "select * from KAT" & srechnertab
    sSQL = sSQL & " order by azeit "
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
        If Not IsNull(rsrs!artnr) Then
            sSatz = rsrs!artnr & Space(7 - Len(rsrs!artnr))
        End If
        
        If Not IsNull(rsrs!BEZEICH) Then
            sSatz = sSatz & rsrs!BEZEICH & Space(36 - Len(rsrs!BEZEICH))
        Else
            sSatz = sSatz & Space(36)
        End If
        
        If Not IsNull(rsrs!AZEIT) Then
            sSatz = sSatz & Format(rsrs!AZEIT, "HH:MM") & Space(12 - Len(Format(rsrs!AZEIT, "HH:MM")))
        Else
            sSatz = sSatz & Space(12)
        End If
        
        If Not IsNull(rsrs!Menge) Then
            sSatz = sSatz & rsrs!Menge & Space(4 - Len(rsrs!Menge))
        Else
            sSatz = sSatz & "0" & Space(3)
        End If
        
        If Not IsNull(rsrs!Preis) Then
            sSatz = sSatz & Format$(rsrs!Preis, "####0.00") & Space(9 - Len(Format$(rsrs!Preis, "####0.00")))
        Else
            sSatz = sSatz & Space(9)
        End If
        
        If Not IsNull(rsrs!Kundnr) Then
            sSatz = sSatz & rsrs!Kundnr & Space(8 - Len(rsrs!Kundnr))
        Else
            sSatz = sSatz & Space(8)
        End If
        
        If Not IsNull(rsrs!Kundnr) Then
            sSatz = sSatz & WhatIsXfromKu(rsrs!Kundnr, "Name")
        End If
        
        
        Listx.AddItem sSatz
        
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
ElseIf UCase$(sArt) = "BESTPDRU" Then
    If Trim(sarti) = "" Then
        Screen.MousePointer = 0
        Exit Sub
    End If
   
    sSQL = "select * from BESTPDRU" & srechnertab & " where artnr = " & sarti
    sSQL = sSQL & "  order by " & sOrder
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
        If Not IsNull(rsrs!LASTDATE) Then
            sSatz = Format(rsrs!LASTDATE, "DD.MM.YY") & Space(10 - Len(Format(rsrs!LASTDATE, "DD.MM.YY")))
        End If
        
        If Not IsNull(rsrs!LASTTIME) Then
            sSatz = sSatz & Format(rsrs!LASTTIME, "HH:MM") & Space(12 - Len(Format(rsrs!LASTTIME, "HH:MM")))
        Else
            sSatz = sSatz & Space(12)
        End If
        
        If Not IsNull(rsrs!FILIALE) Then
            sSatz = sSatz & rsrs!FILIALE & Space(3 - Len(rsrs!FILIALE))
        Else
            sSatz = sSatz & Space(3)
        End If
        
        
        
        If Not IsNull(rsrs!OLDBEST) Then
            sSatz = sSatz & Space(6 - Len(rsrs!OLDBEST)) & rsrs!OLDBEST
        Else
            sSatz = sSatz & Space(5) & "0"
        End If
        
        If Not IsNull(rsrs!UMENGE) Then
            sSatz = sSatz & Space(6 - Len(rsrs!UMENGE)) & rsrs!UMENGE
        Else
            sSatz = sSatz & Space(5) & "0"
        End If
        
        
        If Not IsNull(rsrs!NEWBEST) Then
            sSatz = sSatz & Space(6 - Len(rsrs!NEWBEST)) & rsrs!NEWBEST
        Else
            sSatz = sSatz & Space(5) & "0"
        End If
        
        sSatz = sSatz & Space(2)
        
        If Not IsNull(rsrs!AENART) Then
            sSatz = sSatz & rsrs!AENART & Space(21 - Len(rsrs!AENART))
        Else
            sSatz = sSatz & Space(21)
        End If
        
        If Not IsNull(rsrs!AENGRUND) Then
            sSatz = sSatz & rsrs!AENGRUND & Space(21 - Len(rsrs!AENGRUND))
        Else
            sSatz = sSatz & Space(21)
        End If
               
        If Not IsNull(rsrs!BEDIENER) Then
            sSatz = sSatz & rsrs!BEDIENER
        Else
            sSatz = sSatz & Space(7)
        End If
        Listx.AddItem sSatz
        
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
ElseIf UCase$(sArt) = "KUBE" Then

    If Trim(sarti) = "" Then
        Screen.MousePointer = 0
        Exit Sub
    End If
    loeschNEW "KUBE", gdBase
    
    sSQL = "select KUNDBEST.artnr "
    sSQL = sSQL & " , KUNDBEST.bestelltam "
    sSQL = sSQL & " , KUNDBEST.bestelltum "
    sSQL = sSQL & " , KUNDBEST.bestelltmenge "
    sSQL = sSQL & " , KUNDBEST.Filiale "
    sSQL = sSQL & " , KUNDBEST.BEZEICH "
    sSQL = sSQL & " , KUNDBEST.BESTELLTPREIS "
    sSQL = sSQL & " , KUNDBEST.BEDNU "
    sSQL = sSQL & " , KUNDBEST.KUNDNR "
    sSQL = sSQL & " , KUNDBEST.StatusARTIKEL "
    sSQL = sSQL & " into KUBE from KUNDBEST "
    sSQL = sSQL & " where KUNDNR = " & sarti
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update KUBE SET StatusARTIKEL = 'A'  "
    sSQL = sSQL & " where StatusARTIKEL = 'INBESTELLUNG' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update KUBE SET StatusARTIKEL = 'B'  "
    sSQL = sSQL & " where StatusARTIKEL = 'BESTELLT' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update KUBE SET StatusARTIKEL = 'C'  "
    sSQL = sSQL & " where StatusARTIKEL = 'GELIEFERT' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update KUBE SET StatusARTIKEL = 'D'  "
    sSQL = sSQL & " where StatusARTIKEL = 'NICHTGELIEFERT' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "select * from KUBE "
    sSQL = sSQL & " order by " & sOrder
    

    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
        If Not IsNull(rsrs!Bestelltam) Then
            sSatz = Format(rsrs!Bestelltam, "DD.MM.YY") & Space(9 - Len(Format(rsrs!Bestelltam, "DD.MM.YY")))
        End If
        
        If Not IsNull(rsrs!Bestelltum) Then
            sSatz = sSatz & Format(rsrs!Bestelltum, "HH:MM:SS") & Space(9 - Len(Format(rsrs!Bestelltum, "HH:MM")))
        Else
            sSatz = sSatz & Space(9)
        End If
        
        If Not IsNull(rsrs!BestelltMenge) Then
            sSatz = sSatz & rsrs!BestelltMenge & Space(6 - Len(rsrs!BestelltMenge))
        Else
            sSatz = sSatz & "0" & Space(5)
        End If
        
        If Not IsNull(rsrs!artnr) Then
            sSatz = sSatz & rsrs!artnr & Space(7 - Len(rsrs!artnr))
        Else
            sSatz = sSatz & "0" & Space(6)
        End If
        
        If Not IsNull(rsrs!BEZEICH) Then
            sSatz = sSatz & rsrs!BEZEICH & Space(36 - Len(rsrs!BEZEICH))
        Else
            sSatz = sSatz & Space(36)
        End If
        
        If Not IsNull(rsrs!FILIALE) Then
            sSatz = sSatz & rsrs!FILIALE & Space(3 - Len(rsrs!FILIALE))
        Else
            sSatz = sSatz & Space(3)
        End If
        
        If Not IsNull(rsrs!BestelltPreis) Then
            sSatz = sSatz & Format$(rsrs!BestelltPreis, "####0.00") & Space(9 - Len(Format$(rsrs!BestelltPreis, "####0.00")))
        Else
            sSatz = sSatz & Space(9)
        End If
        
'        If Not IsNull(rsrs!KUNDNR) Then
'            sSatz = sSatz & rsrs!KUNDNR & Space(7 - Len(rsrs!KUNDNR))
'        Else
'            sSatz = sSatz & Space(7)
'        End If
        
        If Not IsNull(rsrs!BEDNU) Then
            sSatz = sSatz & rsrs!BEDNU & Space(3 - Len(rsrs!BEDNU))
        Else
            sSatz = sSatz & Space(3)
        End If
        
        If Not IsNull(rsrs!StatusARTIKEL) Then
            Select Case rsrs!StatusARTIKEL
                Case "A"
                    sSatz = sSatz & "noch nicht bestellt"
                Case "B"
                    sSatz = sSatz & "ist bestellt"
                Case "C"
                    sSatz = sSatz & "geliefert"
                Case "D"
                    sSatz = sSatz & "nicht geliefert"
                
                
                
            End Select
        Else
            sSatz = sSatz & Space(20)
        End If
        
        
        
        
        Listx.AddItem sSatz
        
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
ElseIf UCase$(sArt) = "EANPDRU" Then
    If Trim(sarti) = "" Then
        Screen.MousePointer = 0
        Exit Sub
    End If
   
    sSQL = "select * from EANPDRU where artnr = " & sarti
    sSQL = sSQL & "  order by " & sOrder
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
        If Not IsNull(rsrs!LASTDATE) Then
            sSatz = Format(rsrs!LASTDATE, "DD.MM.YY") & Space(10 - Len(Format(rsrs!LASTDATE, "DD.MM.YY")))
        End If
        
        If Not IsNull(rsrs!LASTTIME) Then
            sSatz = sSatz & Format(rsrs!LASTTIME, "HH:MM") & Space(12 - Len(Format(rsrs!LASTTIME, "HH:MM")))
        Else
            sSatz = sSatz & Space(12)
        End If
        
        If Not IsNull(rsrs!FILIALE) Then
            sSatz = sSatz & rsrs!FILIALE & Space(3 - Len(rsrs!FILIALE))
        Else
            sSatz = sSatz & Space(3)
        End If
        
        sSatz = sSatz & Space(2)
        
        If Not IsNull(rsrs!EAN) Then
            sSatz = sSatz & rsrs!EAN & Space(16 - Len(rsrs!EAN))
        Else
            sSatz = sSatz & Space(15) & " "
        End If
        
        
        If Not IsNull(rsrs!AENART) Then
            sSatz = sSatz & rsrs!AENART & Space(22 - Len(rsrs!AENART))
        Else
            sSatz = sSatz & Space(21)
        End If
               
        If Not IsNull(rsrs!BEDIENER) Then
            sSatz = sSatz & rsrs!BEDIENER
        Else
            sSatz = sSatz & Space(7)
        End If
        Listx.AddItem sSatz
        
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
ElseIf UCase$(sArt) = "EANPDRUEAN" Then
    If Trim(sarti) = "" Then
        Screen.MousePointer = 0
        Exit Sub
    End If
   
    sSQL = "select * from EANPDRU where EAN = '" & sarti & "'"
    sSQL = sSQL & "  order by " & sOrder
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
        If Not IsNull(rsrs!LASTDATE) Then
            sSatz = Format(rsrs!LASTDATE, "DD.MM.YY") & Space(10 - Len(Format(rsrs!LASTDATE, "DD.MM.YY")))
        End If
        
        If Not IsNull(rsrs!LASTTIME) Then
            sSatz = sSatz & Format(rsrs!LASTTIME, "HH:MM") & Space(8 - Len(Format(rsrs!LASTTIME, "HH:MM")))
        Else
            sSatz = sSatz & Space(8)
        End If
        
        If Not IsNull(rsrs!FILIALE) Then
            sSatz = sSatz & rsrs!FILIALE & Space(3 - Len(rsrs!FILIALE))
        Else
            sSatz = sSatz & Space(3)
        End If
        
        sSatz = sSatz & Space(1)
        
        If Not IsNull(rsrs!artnr) Then
            sSatz = sSatz & rsrs!artnr & Space(7 - Len(rsrs!artnr))
        Else
            sSatz = sSatz & Space(6) & " "
        End If
        
        cBez = ermBezeichausWGN(rsrs!artnr)
        
        sSatz = sSatz & cBez & Space(36 - Len(cBez))
        
        If Not IsNull(rsrs!AENART) Then
            sSatz = sSatz & rsrs!AENART & Space(22 - Len(rsrs!AENART))
        Else
            sSatz = sSatz & Space(21)
        End If
               
        If Not IsNull(rsrs!BEDIENER) Then
            sSatz = sSatz & rsrs!BEDIENER
        Else
            sSatz = sSatz & Space(7)
        End If
        Listx.AddItem sSatz
        
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing

ElseIf UCase$(sArt) = "KVKPR1PDRU" Then
    If Trim(sarti) = "" Then
        Screen.MousePointer = 0
        Exit Sub
    End If
   
    sSQL = "select * from KVKPR1PDRU where artnr = " & sarti
    sSQL = sSQL & "  order by " & sOrder
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
        If Not IsNull(rsrs!LASTDATE) Then
            sSatz = Format(rsrs!LASTDATE, "DD.MM.YY") & Space(10 - Len(Format(rsrs!LASTDATE, "DD.MM.YY")))
        End If
        
        If Not IsNull(rsrs!LASTTIME) Then
            sSatz = sSatz & Format(rsrs!LASTTIME, "HH:MM") & Space(12 - Len(Format(rsrs!LASTTIME, "HH:MM")))
        Else
            sSatz = sSatz & Space(12)
        End If
        
        If Not IsNull(rsrs!FILIALE) Then
            sSatz = sSatz & rsrs!FILIALE & Space(3 - Len(rsrs!FILIALE))
        Else
            sSatz = sSatz & Space(3)
        End If
        
        sSatz = sSatz & Space(2)
        
        If Not IsNull(rsrs!KVKPR1) Then
            sSatz = sSatz & Format$(rsrs!KVKPR1, "####0.00") & Space(14 - Len(Format$(rsrs!KVKPR1, "####0.00")))
        Else
            sSatz = sSatz & Space(14)
        End If
        
        
        If Not IsNull(rsrs!AENART) Then
            sSatz = sSatz & rsrs!AENART & Space(21 - Len(rsrs!AENART))
        Else
            sSatz = sSatz & Space(21)
        End If
               
        If Not IsNull(rsrs!BEDIENER) Then
            sSatz = sSatz & rsrs!BEDIENER
        Else
            sSatz = sSatz & Space(7)
        End If
        Listx.AddItem sSatz
        
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing


    
ElseIf UCase$(sArt) = "ZUGANG" Then
    If Trim(sarti) = "" Then
        Screen.MousePointer = 0
        Exit Sub
    End If
    sSQL = "select * from Zugang where artnr = " & sarti
    sSQL = sSQL & " and not rek is null "
    sSQL = sSQL & " order by  " & sOrder
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
        If Not IsNull(rsrs!ADATE) Then
            sSatz = Format(rsrs!ADATE, "DD.MM.YY") & Space(10 - Len(Format(rsrs!ADATE, "DD.MM.YY")))
        End If
        
        If Not IsNull(rsrs!Uhrzeit) Then
            sSatz = sSatz & Format(rsrs!Uhrzeit, "HH:MM") & Space(12 - Len(Format(rsrs!Uhrzeit, "HH:MM")))
        Else
            sSatz = sSatz & Space(12)
        End If
        
        If Not IsNull(rsrs!BEWEGUNG) Then
            sSatz = sSatz & rsrs!BEWEGUNG & Space(14 - Len(rsrs!BEWEGUNG))
        Else
            sSatz = sSatz & "0" & Space(13)
        End If
        
        If Not IsNull(rsrs!rek) Then
            sSatz = sSatz & Format$(rsrs!rek, "####0.00") & Space(14 - Len(Format$(rsrs!rek, "####0.00")))
        Else
            sSatz = sSatz & Space(14)
        End If
        
        
        Listx.AddItem sSatz
        
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing

ElseIf UCase$(sArt) = "VERKAUF" Then

    If Trim(sarti) = "" Then
        Screen.MousePointer = 0
        Exit Sub
    End If
    sSQL = "select * from Kassjour where artnr = " & sarti
    sSQL = sSQL & " order by " & sOrder
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
        If Not IsNull(rsrs!ADATE) Then
            sSatz = Format(rsrs!ADATE, "DD.MM.YY") & Space(10 - Len(Format(rsrs!ADATE, "DD.MM.YY")))
        End If
        
        If Not IsNull(rsrs!AZEIT) Then
            sSatz = sSatz & Format(rsrs!AZEIT, "HH:MM") & Space(12 - Len(Format(rsrs!AZEIT, "HH:MM")))
        Else
            sSatz = sSatz & Space(12)
        End If
        
        If Not IsNull(rsrs!Menge) Then
            sSatz = sSatz & rsrs!Menge & Space(4 - Len(rsrs!Menge))
        Else
            sSatz = sSatz & "0" & Space(3)
        End If
        
        If Not IsNull(rsrs!FILIALE) Then
            sSatz = sSatz & rsrs!FILIALE & Space(3 - Len(rsrs!FILIALE))
        Else
            sSatz = sSatz & Space(3)
        End If
        
        If Not IsNull(rsrs!Preis) Then
            sSatz = sSatz & Format$(rsrs!Preis, "####0.00") & Space(9 - Len(Format$(rsrs!Preis, "####0.00")))
        Else
            sSatz = sSatz & Space(9)
        End If
        
        If Not IsNull(rsrs!Kundnr) Then
            sSatz = sSatz & rsrs!Kundnr & Space(7 - Len(rsrs!Kundnr))
        Else
            sSatz = sSatz & Space(7)
        End If
        
        If Not IsNull(rsrs!BEDIENER) Then
            sSatz = sSatz & rsrs!BEDIENER
        Else
            sSatz = sSatz & Space(7)
        End If
        Listx.AddItem sSatz
        
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
ElseIf UCase$(sArt) = "VERKAUFKU" Then


    If Trim(sarti) = "" Then
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    
    
    'Neu Rabattkennzeichen
    
    If SpalteInTabellegefundenNEW("KUNDAZE", "RABATTF", gdBase) = False Then
        SpalteAnfuegenNEW "KUNDAZE", "RABATTF", "TEXT(1)", gdBase
        
        sSQL = "Update KUNDAZE k inner join Artikel a on k.artnr = a.artnr  Set k.rabattf = '*' "
        sSQL = sSQL & " where a.rabatt_OK = 'N'"
        gdBase.Execute sSQL, dbFailOnError
    
    End If
    
    'Neu Ende

    sSQL = "select * from KUNDAZE "
    sSQL = sSQL & " order by " & sOrder
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
        If Not IsNull(rsrs!ADATE) Then
            sSatz = Format(rsrs!ADATE, "DD.MM.YY") & Space(10 - Len(Format(rsrs!ADATE, "DD.MM.YY")))
        End If
        
        If Not IsNull(rsrs!Menge) Then
            sSatz = sSatz & rsrs!Menge & Space(7 - Len(rsrs!Menge))
        Else
            sSatz = sSatz & "0" & Space(6)
        End If
        
        If Not IsNull(rsrs!rabattf) Then
            sSatz = sSatz & rsrs!rabattf
        Else
            sSatz = sSatz & " "
        End If
        
        If Not IsNull(rsrs!artnr) Then
            sSatz = sSatz & rsrs!artnr & Space(7 - Len(rsrs!artnr))
        Else
            sSatz = sSatz & "0" & Space(6)
        End If
        
        If Not IsNull(rsrs!BEZEICH) Then
            sSatz = sSatz & rsrs!BEZEICH & Space(36 - Len(rsrs!BEZEICH))
        Else
            sSatz = sSatz & "0" & Space(35)
        End If
        
        If Not IsNull(rsrs!FILIALE) Then
            sSatz = sSatz & rsrs!FILIALE & Space(4 - Len(rsrs!FILIALE))
        Else
            sSatz = sSatz & Space(4)
        End If
        
        If Not IsNull(rsrs!Preis) Then
            sSatz = sSatz & Format$(rsrs!Preis, "####0.00") & Space(9 - Len(Format$(rsrs!Preis, "####0.00")))
        Else
            sSatz = sSatz & Space(9)
        End If
        
        If Not IsNull(rsrs!BEDNR) Then
            sSatz = sSatz & rsrs!BEDNR
        Else
            sSatz = sSatz & Space(7)
        End If
        Listx.AddItem sSatz
        
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    If SpalteInTabellegefundenNEW("KUNDAZE", "RABATTF", gdBase) = False Then
        sSQL = "Alter Table KUNDAZE drop rabattf  "
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    
ElseIf UCase$(sArt) = "KUB1" Then

    If sarti = "" Then
        Screen.MousePointer = 0
        Exit Sub
    End If
    sSQL = "select KUNDBEST.artnr "
    sSQL = sSQL & " , KUNDBEST.bestelltam "
    sSQL = sSQL & " , KUNDBEST.bestelltum "
    sSQL = sSQL & " , KUNDBEST.KUNDNR "
    sSQL = sSQL & " , KUNDBEST.bestelltmenge "
    sSQL = sSQL & " , KUNDBEST.Filiale "
    sSQL = sSQL & " , KUNDBEST.BEZEICH "
    sSQL = sSQL & " , KUNDBEST.BESTELLTPREIS "
    sSQL = sSQL & " , KUNDBEST.BEDNU "
    sSQL = sSQL & " from KUNDBEST , artlief where artlief.artnr = kundbest.artnr "
    sSQL = sSQL & " and Artlief.LINR = " & sarti
    sSQL = sSQL & " and KUNDBEST.StatusARTIKEL = 'INBESTELLUNG' "
    sSQL = sSQL & " and KUNDBEST.sendok = false "
    
    sSQL = sSQL & " order by " & sOrder
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
        If Not IsNull(rsrs!Bestelltam) Then
            sSatz = Format(rsrs!Bestelltam, "DD.MM.YY") & Space(10 - Len(Format(rsrs!Bestelltam, "DD.MM.YY")))
        End If
        
        If Not IsNull(rsrs!Bestelltum) Then
            sSatz = sSatz & Format(rsrs!Bestelltum, "HH:MM:SS") & Space(12 - Len(Format(rsrs!Bestelltum, "HH:MM:SS")))
        Else
            sSatz = sSatz & Space(12)
        End If
        
        If Not IsNull(rsrs!BestelltMenge) Then
            sSatz = sSatz & rsrs!BestelltMenge & Space(4 - Len(rsrs!BestelltMenge))
        Else
            sSatz = sSatz & "0" & Space(3)
        End If
        
        If Not IsNull(rsrs!artnr) Then
            sSatz = sSatz & rsrs!artnr & Space(7 - Len(rsrs!artnr))
        Else
            sSatz = sSatz & "0" & Space(6)
        End If
        
        If Not IsNull(rsrs!BEZEICH) Then
            sSatz = sSatz & rsrs!BEZEICH & Space(36 - Len(rsrs!BEZEICH))
        Else
            sSatz = sSatz & "0" & Space(35)
        End If
        
        If Not IsNull(rsrs!FILIALE) Then
            sSatz = sSatz & rsrs!FILIALE & Space(3 - Len(rsrs!FILIALE))
        Else
            sSatz = sSatz & Space(3)
        End If
        
        If Not IsNull(rsrs!BestelltPreis) Then
            sSatz = sSatz & Format$(rsrs!BestelltPreis, "####0.00") & Space(9 - Len(Format$(rsrs!BestelltPreis, "####0.00")))
        Else
            sSatz = sSatz & Space(9)
        End If
        
        If Not IsNull(rsrs!Kundnr) Then
            sSatz = sSatz & rsrs!Kundnr & Space(8 - Len(rsrs!Kundnr))
        Else
            sSatz = sSatz & Space(8)
        End If
        
        If Not IsNull(rsrs!BEDNU) Then
            sSatz = sSatz & rsrs!BEDNU
        Else
            sSatz = sSatz & Space(7)
        End If
        Listx.AddItem sSatz
        
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close
ElseIf UCase$(sArt) = "KUBART" Then
    If Trim(sarti) = "" Then
        Screen.MousePointer = 0
        Exit Sub
    End If

    sSQL = "select KUNDBEST.artnr "
    sSQL = sSQL & " , KUNDBEST.bestelltam "
    sSQL = sSQL & " , KUNDBEST.bestelltum "
    sSQL = sSQL & " , KUNDBEST.KUNDNR "
    sSQL = sSQL & " , KUNDBEST.bestelltmenge "
    sSQL = sSQL & " , KUNDBEST.Filiale "
    sSQL = sSQL & " , KUNDBEST.BEZEICH "
    sSQL = sSQL & " , KUNDBEST.BESTELLTPREIS "
    sSQL = sSQL & " , KUNDBEST.BEDNU "
    sSQL = sSQL & " from KUNDBEST , Artikel where Artikel.artnr = kundbest.artnr "
    sSQL = sSQL & " and Artikel.artnr = " & sarti
    sSQL = sSQL & " and KUNDBEST.StatusARTIKEL = 'INBESTELLUNG' "
    
    sSQL = sSQL & " order by " & sOrder
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
        If Not IsNull(rsrs!Bestelltam) Then
            sSatz = Format(rsrs!Bestelltam, "DD.MM.YY") & Space(10 - Len(Format(rsrs!Bestelltam, "DD.MM.YY")))
        End If
        
        If Not IsNull(rsrs!Bestelltum) Then
            sSatz = sSatz & Format(rsrs!Bestelltum, "HH:MM") & Space(12 - Len(Format(rsrs!Bestelltum, "HH:MM")))
        Else
            sSatz = sSatz & Space(12)
        End If
        
        If Not IsNull(rsrs!BestelltMenge) Then
            sSatz = sSatz & rsrs!BestelltMenge & Space(4 - Len(rsrs!BestelltMenge))
        Else
            sSatz = sSatz & "0" & Space(3)
        End If
        
        If Not IsNull(rsrs!artnr) Then
            sSatz = sSatz & rsrs!artnr & Space(7 - Len(rsrs!artnr))
        Else
            sSatz = sSatz & "0" & Space(6)
        End If
        
        If Not IsNull(rsrs!BEZEICH) Then
            sSatz = sSatz & rsrs!BEZEICH & Space(36 - Len(rsrs!BEZEICH))
        Else
            sSatz = sSatz & "0" & Space(35)
        End If
        
        If Not IsNull(rsrs!FILIALE) Then
            sSatz = sSatz & rsrs!FILIALE & Space(3 - Len(rsrs!FILIALE))
        Else
            sSatz = sSatz & Space(3)
        End If
        
        If Not IsNull(rsrs!BestelltPreis) Then
            sSatz = sSatz & Format$(rsrs!BestelltPreis, "####0.00") & Space(9 - Len(Format$(rsrs!BestelltPreis, "####0.00")))
        Else
            sSatz = sSatz & Space(9)
        End If
        
        If Not IsNull(rsrs!Kundnr) Then
            sSatz = sSatz & rsrs!Kundnr & Space(7 - Len(rsrs!Kundnr))
        Else
            sSatz = sSatz & Space(7)
        End If
        
        If Not IsNull(rsrs!BEDNU) Then
            sSatz = sSatz & rsrs!BEDNU
        Else
            sSatz = sSatz & Space(7)
        End If
        Listx.AddItem sSatz
        
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
ElseIf UCase$(sArt) = "KUB" Then
    If Trim(sarti) = "" Then
        Screen.MousePointer = 0
        Exit Sub
    End If

    sSQL = "select KUNDBEST.artnr "
    sSQL = sSQL & " , KUNDBEST.bestelltam "
    sSQL = sSQL & " , KUNDBEST.bestelltum "
    sSQL = sSQL & " , KUNDBEST.KUNDNR "
    sSQL = sSQL & " , KUNDBEST.bestelltmenge "
    sSQL = sSQL & " , KUNDBEST.Filiale "
    sSQL = sSQL & " , KUNDBEST.BEZEICH "
    sSQL = sSQL & " , KUNDBEST.BESTELLTPREIS "
    sSQL = sSQL & " , KUNDBEST.BEDNU "
    sSQL = sSQL & " from KUNDBEST , artlief where artlief.artnr = kundbest.artnr "
    sSQL = sSQL & " and Artlief.LINR = " & sarti
    sSQL = sSQL & " and KUNDBEST.StatusARTIKEL = 'INBESTELLUNG' "
    
    sSQL = sSQL & " order by " & sOrder
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
        If Not IsNull(rsrs!Bestelltam) Then
            sSatz = Format(rsrs!Bestelltam, "DD.MM.YY") & Space(10 - Len(Format(rsrs!Bestelltam, "DD.MM.YY")))
        End If
        
        If Not IsNull(rsrs!Bestelltum) Then
            sSatz = sSatz & Format(rsrs!Bestelltum, "HH:MM") & Space(12 - Len(Format(rsrs!Bestelltum, "HH:MM")))
        Else
            sSatz = sSatz & Space(12)
        End If
        
        If Not IsNull(rsrs!BestelltMenge) Then
            sSatz = sSatz & rsrs!BestelltMenge & Space(4 - Len(rsrs!BestelltMenge))
        Else
            sSatz = sSatz & "0" & Space(3)
        End If
        
        If Not IsNull(rsrs!artnr) Then
            sSatz = sSatz & rsrs!artnr & Space(7 - Len(rsrs!artnr))
        Else
            sSatz = sSatz & "0" & Space(6)
        End If
        
        If Not IsNull(rsrs!BEZEICH) Then
            sSatz = sSatz & rsrs!BEZEICH & Space(36 - Len(rsrs!BEZEICH))
        Else
            sSatz = sSatz & "0" & Space(35)
        End If
        
        If Not IsNull(rsrs!FILIALE) Then
            sSatz = sSatz & rsrs!FILIALE & Space(3 - Len(rsrs!FILIALE))
        Else
            sSatz = sSatz & Space(3)
        End If
        
        If Not IsNull(rsrs!BestelltPreis) Then
            sSatz = sSatz & Format$(rsrs!BestelltPreis, "####0.00") & Space(9 - Len(Format$(rsrs!BestelltPreis, "####0.00")))
        Else
            sSatz = sSatz & Space(9)
        End If
        
        If Not IsNull(rsrs!Kundnr) Then
            sSatz = sSatz & rsrs!Kundnr & Space(7 - Len(rsrs!Kundnr))
        Else
            sSatz = sSatz & Space(7)
        End If
        
        If Not IsNull(rsrs!BEDNU) Then
            sSatz = sSatz & rsrs!BEDNU
        Else
            sSatz = sSatz & Space(7)
        End If
        Listx.AddItem sSatz
        
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
ElseIf UCase$(sArt) = "KOND" Then

    If Trim(sarti) = "" Then
        Screen.MousePointer = 0
        Exit Sub
    End If
    sSQL = "select KONDITIONEN.artnr "
    sSQL = sSQL & " , KONDITIONEN.KONDI "
    sSQL = sSQL & " , KONDITIONEN.FAKTOR "
    sSQL = sSQL & " , ARTIKEL.BEZEICH "
    sSQL = sSQL & " from KONDITIONEN , artlief, artikel where artlief.artnr = KONDITIONEN.artnr "
    sSQL = sSQL & " and Artikel.ARTNR = KONDITIONEN.artnr "
    sSQL = sSQL & " and Artlief.LINR = " & sarti
    
    
    sSQL = sSQL & " order by " & sOrder
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
        If Not IsNull(rsrs!artnr) Then
            sSatz = rsrs!artnr & Space(7 - Len(rsrs!artnr))
        End If
        
        If Not IsNull(rsrs!BEZEICH) Then
            sSatz = sSatz & rsrs!BEZEICH & Space(36 - Len(rsrs!BEZEICH))
        Else
            sSatz = sSatz & "0" & Space(35)
        End If
        
        If Not IsNull(rsrs!kondi) Then
            sSatz = sSatz & rsrs!kondi & Space(3 - Len(rsrs!kondi))
        Else
            sSatz = sSatz & Space(3)
        End If
        
        sSatz = sSatz & " + "
        
        If Not IsNull(rsrs!Faktor) Then
            sSatz = sSatz & rsrs!Faktor & Space(3 - Len(rsrs!Faktor))
        Else
            sSatz = sSatz & Space(3)
        End If
        Listx.AddItem sSatz
        
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
ElseIf UCase$(sArt) = "UMSCHLAG" Then

    sSQL = "select * from TB" & srechnertab
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
        If Not IsNull(rsrs!Monat) Then
            sSatz = rsrs!Monat & Space(8 - Len(rsrs!Monat))
        End If
        
        If Not IsNull(rsrs!jahr) Then
            sSatz = sSatz & rsrs!jahr & Space(13 - Len(rsrs!jahr))
        End If
        
        If Not IsNull(rsrs!BESTAND) Then
            sSatz = sSatz & rsrs!BESTAND & Space(11 - Len(rsrs!BESTAND))
        Else
            sSatz = sSatz & "0" & Space(10)
        End If
        
        If Not IsNull(rsrs!Verkauf) Then
            sSatz = sSatz & rsrs!Verkauf & Space(8 - Len(rsrs!Verkauf))
        Else
            sSatz = sSatz & "0" & Space(7)
        End If
        Listx.AddItem sSatz
        
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
End If

Listx.Visible = True
Listx.Refresh

Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "ZeigArtHistInList"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
End Sub
Public Function WhatIsXfromKu(cKunde As String, sSpalte As String) As String
    On Error GoTo LOKAL_ERROR
    Dim sSQL As String
    Dim rs As Recordset
    
    WhatIsXfromKu = ""
    
    sSQL = "select " & sSpalte & " as maxi from kunden where kundnr = " & cKunde
    Set rs = gdBase.OpenRecordset(sSQL)
    If Not rs.EOF Then
        If Not IsNull(rs!maxi) Then
            WhatIsXfromKu = rs!maxi
        End If
    End If
    rs.Close: Set rs = Nothing
    
    Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "WhatIsXfromKu"
    Fehler.gsFehlertext = "Beim Ermitteln von Kundendaten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function HoleLagerumschlag1(cART As String) As Double
    On Error GoTo LOKAL_ERROR

    Dim sSQL    As String
    
    HoleLagerumschlag1 = 0
    
    VorBereitLagerumschlag

    sSQL = "Update TB" & srechnertab
    sSQL = sSQL & " set Bestand = 0 and Verkauf = 0"
    gdBase.Execute sSQL, dbFailOnError
    
    HoleLagerumschlag1 = LagerumschlagRestberechnung(cART, "TB" & srechnertab)
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "HoleLagerumschlag1"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function HoleLagerumschlag2(cART As String) As Double
    On Error GoTo LOKAL_ERROR

    Dim sSQL    As String
    
    HoleLagerumschlag2 = 0
    
'    VorBereitLagerumschlag2

    sSQL = "Update TB" & srechnertab
    sSQL = sSQL & " set Bestand = 0 and Verkauf = 0"
    gdBase.Execute sSQL, dbFailOnError
    
    HoleLagerumschlag2 = LagerumschlagRestberechnung(cART, "TB" & srechnertab)
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "HoleLagerumschlag2"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function HoleLagerumschlag1VJ(cART As String, siEkpr As Double) As Double
    On Error GoTo LOKAL_ERROR

    Dim sSQL        As String
    
    HoleLagerumschlag1VJ = 0

    sSQL = "Update TBVJ" & srechnertab
    sSQL = sSQL & " set Bestand = 0 and Verkauf = 0"
    gdBase.Execute sSQL, dbFailOnError
    
    HoleLagerumschlag1VJ = LagerumschlagRestberechnung(cART, "TBVJ" & srechnertab)
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "HoleLagerumschlag1VJ"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function HoleLagerumschlag1VVJ(cART As String, siEkpr As Double) As Double
    On Error GoTo LOKAL_ERROR

    Dim sSQL        As String

    HoleLagerumschlag1VVJ = 0

    sSQL = "Update TBVVJ" & srechnertab
    sSQL = sSQL & " set Bestand = 0 and Verkauf = 0"
    gdBase.Execute sSQL, dbFailOnError
    
    HoleLagerumschlag1VVJ = LagerumschlagRestberechnung(cART, "TBVVJ" & srechnertab)
   
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "HoleLagerumschlag1VVJ"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function LagerumschlagRestberechnung(cART As String, sTab As String) As Double
    On Error GoTo LOKAL_ERROR

    Dim lGBest      As Long
    Dim sSQL        As String
    Dim rsrs        As Recordset

    Dim i           As Integer
    Dim lMoni       As Byte
    Dim lJahri      As Integer
    Dim lsumBest    As Long
    Dim lsumVerk    As Long

    Dim schnittBest As Single
    Dim schnittVerk As Single
    
    Dim lTeiler     As Long
    
    lTeiler = 0
    
    cART = SwapStr(cART, ",", "")
    If cART = "" Then
        Exit Function
    End If

    LagerumschlagRestberechnung = 0

    sSQL = "Update " & sTab & " set Bestand = 0 , Verkauf = 0"
    gdBase.Execute sSQL, dbFailOnError

    sSQL = "Update " & sTab & " inner join DVKART on " & sTab & ".monat = DVKART.Monat "
    sSQL = sSQL & " and " & sTab & ".jahr = DVKART.Jahr "
    sSQL = sSQL & " set " & sTab & ".Verkauf =  DVKART.Dvk "
    sSQL = sSQL & " where DVKART.artnr=" & cART
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update " & sTab & " inner join GDLAGER on " & sTab & ".monat = GDLAGER.Monat "
    sSQL = sSQL & " and " & sTab & ".jahr = GDLAGER.Jahr "
    sSQL = sSQL & " set " & sTab & ".BESTAND =  GDLAGER.SBESTAND "
    sSQL = sSQL & " where GDLAGER.artnr=" & cART
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update " & sTab
    sSQL = sSQL & " set BESTAND =  0  "
    sSQL = sSQL & " where BESTAND < 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "select bestand from " & sTab & " where bestand > 0 "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lTeiler = rsrs.RecordCount
    End If
    rsrs.Close: Set rsrs = Nothing

    sSQL = "select sum(bestand) as maxbestand from " & sTab
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxbestand) Then
            lsumBest = rsrs!maxbestand
        Else
            lsumBest = 0
        End If
    End If
    rsrs.Close: Set rsrs = Nothing

    sSQL = "select sum(Verkauf) as maxverkauf from " & sTab
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxverkauf) Then
            lsumVerk = rsrs!maxverkauf
        Else
            lsumVerk = 0
        End If
    End If
    rsrs.Close: Set rsrs = Nothing

    If lTeiler <> 0 Then
        schnittBest = lsumBest / lTeiler ' 12
    End If
    schnittVerk = lsumVerk '/ 12

'    schnittBest = schnittBest * siEkpr
'    schnittVerk = schnittVerk * siEkpr

    If schnittBest = 0 Then
        If schnittVerk = 0 Then
            LagerumschlagRestberechnung = 0
        Else
            LagerumschlagRestberechnung = 0
        End If
    Else
        LagerumschlagRestberechnung = schnittVerk / schnittBest
    End If

    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "LagerumschlagRestberechnung"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Public Function LUGSAktuell() As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset

    LUGSAktuell = False
    
    sSQL = "select top 1 adate from allartlukopf "
    sSQL = sSQL & " order by adate desc "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!ADATE) Then
            If Month(DateValue(Now)) = Month(rsrs!ADATE) Then
                If Year(DateValue(Now)) = Year(rsrs!ADATE) Then
                    LUGSAktuell = True
                End If
            End If
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "LUGSAktuell"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function KUDDAktuell() As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset

    KUDDAktuell = False
    
    sSQL = "select top 1 adate from KUDD "
    sSQL = sSQL & " order by adate desc "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!ADATE) Then
            If Month(DateValue(Now)) = Month(rsrs!ADATE) Then
                If Year(DateValue(Now)) = Year(rsrs!ADATE) Then
                    KUDDAktuell = True
                End If
            End If
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "KUDDAktuell"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermSchBest() As Double
    On Error GoTo LOKAL_ERROR

    Dim sSQL        As String
    Dim rsrs        As Recordset
    Dim lTeiler     As Long
    
    Dim lsumBest    As Long

    ermSchBest = 0
    lTeiler = 0
    
    sSQL = "select bestand from TB" & srechnertab & " where bestand > 0 "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lTeiler = rsrs.RecordCount
    End If
    rsrs.Close: Set rsrs = Nothing
    
    sSQL = "select sum(bestand) as maxbestand from TB" & srechnertab
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxbestand) Then
            lsumBest = rsrs!maxbestand
        Else
            lsumBest = 0
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    If lTeiler <> 0 Then
        ermSchBest = lsumBest / lTeiler '12
    End If
        
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "ermSchBest"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Function
Public Function ermSchVerk() As Double
    On Error GoTo LOKAL_ERROR

    Dim sSQL        As String
    Dim rsrs        As Recordset
    Dim lsumVerk    As Long
    
    ermSchVerk = 0

    sSQL = "select sum(Verkauf) as maxverkauf from TB" & srechnertab
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
    
        If Not IsNull(rsrs!maxverkauf) Then
            lsumVerk = rsrs!maxverkauf
        Else
            lsumVerk = 0
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
   
    ermSchVerk = lsumVerk '/ 12
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "ermSchVerk"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Function
Public Sub schreibeWKEtidru(sArt As String, lnewBest As Long, lFil As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL    As String
    Dim rsrs    As Recordset
    Dim rsArtikel   As Recordset
    
    If lnewBest = 0 Then
        Exit Sub
    End If
    
    sSQL = "Select * from ETIDRU where artnr = " & sArt
    sSQL = sSQL & " and filnr = " & lFil
    Set rsrs = gdBase.OpenRecordset(sSQL)
    
    sSQL = "Select * from Artikel where artnr = " & sArt
    Set rsArtikel = gdBase.OpenRecordset(sSQL)
    If Not rsArtikel.EOF Then
    
        If Not rsrs.EOF Then
            rsrs.Edit
            
            'mal gucken, ob dass funktioniert
            '21.08.14
            
            '29.01.15 es funktionierte nicht
            'rückgängig für Walther
            
'            rsrs!BESTAND = lnewBest
'            rsrs!ANZAHL = lnewBest
            
            rsrs!BESTAND = rsrs!BESTAND + lnewBest
            rsrs!ANZAHL = rsrs!ANZAHL + lnewBest
        Else
            rsrs.AddNew
            rsrs!BESTAND = lnewBest
            rsrs!ANZAHL = lnewBest
        End If
        
        rsrs!artnr = sArt
        rsrs!BEZEICH = rsArtikel!BEZEICH
        rsrs!vkpr = rsArtikel!KVKPR1
        rsrs!LIBESNR = rsArtikel!LIBESNR
        rsrs!EAN = rsArtikel!EAN
        rsrs!LPZ = rsArtikel!LPZ
        rsrs!linr = rsArtikel!linr
        rsrs!filnr = lFil
        rsrs!Pcname = srechnertab
        rsrs.Update
        
    End If
    
    rsrs.Close: Set rsrs = Nothing
    rsArtikel.Close: Set rsArtikel = Nothing
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "schreibeWKEtidru"
    Fehler.gsFehlertext = "Beim Etiketten erzeugen für Winkiss ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Public Sub IsinArtlief(cArtNr As String, cLinr As String, cLEKPR As String, cLiBesNr As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL    As String
    Dim rsArtlief    As Recordset
    
    cSQL = "Select * from ARTLIEF where ARTNR = " & cArtNr & " and LINR = " & cLinr & " "
    Set rsArtlief = gdBase.OpenRecordset(cSQL)
    
    If rsArtlief.EOF Then
        rsArtlief.AddNew
        rsArtlief!SYNStatus = "A"
        rsArtlief!artnr = cArtNr
        rsArtlief!linr = cLinr
        rsArtlief!lekpr = cLEKPR
        rsArtlief!LIBESNR = cLiBesNr
        rsArtlief!MINMEN = 0
    Else
        rsArtlief.Edit
        rsArtlief!SYNStatus = "E"
        rsArtlief!artnr = cArtNr
        rsArtlief!linr = cLinr
        rsArtlief!lekpr = cLEKPR
    End If
    rsArtlief.Update
    rsArtlief.Close: Set rsArtlief = Nothing
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "IsinArtlief"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Public Function ermLangString(langnr As Byte, srtnumber As Long) As String
    On Error GoTo LOKAL_ERROR
    
    Dim rs As Recordset
    Dim sSQL As String

    ermLangString = ""
    
    sSQL = "Select lang" & langnr & " as strErg from LANG where STRINGNR = " & srtnumber
    Set rs = gdApp.OpenRecordset(sSQL)
    If Not rs.EOF Then
        If Not IsNull(rs!strErg) Then
            ermLangString = rs!strErg
        End If
    End If
    rs.Close: Set rs = Nothing

    Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "ermLangString"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Sub FaerbenFlexH(ctmp As String, mshflex As MSHFlexGrid, iSpalte As Integer, lrow As Integer)
    On Error GoTo LOKAL_ERROR
    
    ''''''''''''''''''''''''''''''''''''''''' ODAYY  <<<<< START
    Dim j       As Integer
    ''''''''''''''''''''''''''''''''''''''''' ODAYY  <<<<< ENDE
    
    With mshflex

        If ctmp <> "0" Then
            .Row = lrow
            .Col = iSpalte
            If ctmp = "1" Then
                .CellBackColor = glfarbe(1)
                .CellForeColor = vbBlack
            ElseIf ctmp = "2" Then
                .CellBackColor = glfarbe(2)
                .CellForeColor = vbBlack
            ElseIf ctmp = "3" Then
            
                   ''''''''''''''''''''''''''''''''''''''''' ODAYY  <<<<< START
                    If IchBinBeiKundeSuche Then
                       
                       For j = 0 To .Cols - 1
                         .Col = j
                         .CellBackColor = glfarbe(3)
                         .CellForeColor = vbBlack
                       Next j
                       
                    Else
                   ''''''''''''''''''''''''''''''''''''''''' ODAYY  <<<<< ENDE
                   
                      .CellBackColor = glfarbe(3)
                      .CellForeColor = vbBlack
                      
                   ''''''''''''''''''''''''''''''''''''''''' ODAYY  <<<<< START
                     End If
                   ''''''''''''''''''''''''''''''''''''''''' ODAYY  <<<<< ENDE
                  
            ElseIf ctmp = "4" Then
                .CellBackColor = glfarbe(4)
                .CellForeColor = vbBlack
            ElseIf ctmp = "5" Then
                .CellBackColor = glfarbe(5)
                .CellForeColor = vbBlack
            ElseIf ctmp = "6" Then
                .CellBackColor = glfarbe(6)
                .CellForeColor = vbBlack
            ElseIf ctmp = "7" Then
                .CellBackColor = glfarbe(7)
                .CellForeColor = vbBlack
            ElseIf ctmp = "8" Then
                .CellBackColor = glfarbe(8)
                .CellForeColor = vbBlack
            ElseIf ctmp = "9" Then
                .CellBackColor = glfarbe(9)
                .CellForeColor = vbBlack
            ElseIf ctmp = "11" Then
                .CellBackColor = glfarbe2(1)
                .CellForeColor = vbBlack
            ElseIf ctmp = "12" Then
                .CellBackColor = glfarbe2(2)
                .CellForeColor = vbBlack
            ElseIf ctmp = "13" Then
                .CellBackColor = glfarbe2(3)
                .CellForeColor = vbBlack
            ElseIf ctmp = "14" Then
                .CellBackColor = glfarbe2(4)
                .CellForeColor = vbBlack
            ElseIf ctmp = "15" Then
                .CellBackColor = glfarbe2(5)
                .CellForeColor = vbBlack
            ElseIf ctmp = "16" Then
                .CellBackColor = glfarbe2(6)
                .CellForeColor = vbBlack
            ElseIf ctmp = "17" Then
                .CellBackColor = glfarbe2(7)
                .CellForeColor = vbBlack
            ElseIf ctmp = "18" Then
                .CellBackColor = glfarbe2(8)
                .CellForeColor = vbBlack
            ElseIf ctmp = "19" Then
                .CellBackColor = glfarbe2(9)
                .CellForeColor = vbBlack
            ElseIf ctmp = "99" Then         'eben angefügte Artikel
                .CellBackColor = vbWhite
                .CellForeColor = vbBlue
            ElseIf ctmp = "98" Then         'neue Artikel
                .CellBackColor = vbWhite
                .CellForeColor = vbRed
            ElseIf ctmp = "97" Then
                .CellBackColor = vbYellow   'automatisch kalkulierte
                .CellForeColor = vbBlue
                
            ElseIf ctmp = "95" Then
                .CellBackColor = vbBlue   'nicht geliefert
                .CellForeColor = vbBlack
                
            ElseIf ctmp = "94" Then
                .CellBackColor = glfarbe(0)   'Preisaktion in Vorbereitung
                .CellForeColor = vbBlue
                
            ElseIf ctmp = "93" Then
                .CellBackColor = vbWhite  'Preisaktion aktiv
                .CellForeColor = vbGreen
                
            ElseIf ctmp = "92" Then
                .CellBackColor = &H80000012  'seit 2 Jahren oder noch nie verkauft
                .CellForeColor = vbWhite
                
            Else
                .CellBackColor = glfarbe(0)
                .CellForeColor = vbBlack
            End If
            
        Else
            .Col = iSpalte
            .CellBackColor = glfarbe(0)
            .CellForeColor = vbBlack
        End If
        
    End With
    
    Exit Sub
LOKAL_ERROR:
    If err.Number = 13 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul3"
        Fehler.gsFunktion = "FaerbenFlexH"
        Fehler.gsFehlertext = "Beim Faerben der Tabelle ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Public Sub FaerbenFlex(ctmp As String, mshflex As MSFlexGrid, iSpalte As Integer, lrow As Integer)
    On Error GoTo LOKAL_ERROR
     
    ''''''''''''''''''''''''''''''''''''''''' ODAYY  <<<<< START
    Dim j       As Integer
    ''''''''''''''''''''''''''''''''''''''''' ODAYY  <<<<< ENDE
    
    With mshflex

        If ctmp <> "0" Then
            .Row = lrow
            .Col = iSpalte
            If ctmp = "1" Then
                .CellBackColor = glfarbe(1)
                .CellForeColor = vbBlack
            ElseIf ctmp = "2" Then
                .CellBackColor = glfarbe(2)
                .CellForeColor = vbBlack
            ElseIf ctmp = "3" Then
            
                   ''''''''''''''''''''''''''''''''''''''''' ODAYY  <<<<< START
                    If IchBinBeiKundeSuche Then
                       
                       For j = 0 To .Cols - 1
                         .Col = j
                         .CellBackColor = glfarbe(3)
                         .CellForeColor = vbBlack
                       Next j
                       
                    Else
                   ''''''''''''''''''''''''''''''''''''''''' ODAYY  <<<<< ENDE
                   
                      .CellBackColor = glfarbe(3)
                      .CellForeColor = vbBlack
                      
                   ''''''''''''''''''''''''''''''''''''''''' ODAYY  <<<<< START
                     End If
                   ''''''''''''''''''''''''''''''''''''''''' ODAYY  <<<<< ENDE
                   
                    
            ElseIf ctmp = "4" Then
                .CellBackColor = glfarbe(4)
                .CellForeColor = vbBlack
            ElseIf ctmp = "5" Then
                .CellBackColor = glfarbe(5)
                .CellForeColor = vbBlack
            ElseIf ctmp = "6" Then
                .CellBackColor = glfarbe(6)
                .CellForeColor = vbBlack
            ElseIf ctmp = "7" Then
                .CellBackColor = glfarbe(7)
                .CellForeColor = vbBlack
            ElseIf ctmp = "8" Then
                .CellBackColor = glfarbe(8)
                .CellForeColor = vbBlack
            ElseIf ctmp = "9" Then
                .CellBackColor = glfarbe(9)
                .CellForeColor = vbBlack
            ElseIf ctmp = "11" Then
                .CellBackColor = glfarbe2(1)
                .CellForeColor = vbBlack
            ElseIf ctmp = "12" Then
                .CellBackColor = glfarbe2(2)
                .CellForeColor = vbBlack
            ElseIf ctmp = "13" Then
                .CellBackColor = glfarbe2(3)
                .CellForeColor = vbBlack
            ElseIf ctmp = "14" Then
                .CellBackColor = glfarbe2(4)
                .CellForeColor = vbBlack
            ElseIf ctmp = "15" Then
                .CellBackColor = glfarbe2(5)
                .CellForeColor = vbBlack
            ElseIf ctmp = "16" Then
                .CellBackColor = glfarbe2(6)
                .CellForeColor = vbBlack
            ElseIf ctmp = "17" Then
                .CellBackColor = glfarbe2(7)
                .CellForeColor = vbBlack
            ElseIf ctmp = "18" Then
                .CellBackColor = glfarbe2(8)
                .CellForeColor = vbBlack
            ElseIf ctmp = "19" Then
                .CellBackColor = glfarbe2(9)
                .CellForeColor = vbBlack
            ElseIf ctmp = "99" Then         'eben angefügte Artikel
                .CellBackColor = vbWhite
                .CellForeColor = vbBlue
            ElseIf ctmp = "98" Then         'neue Artikel
                .CellBackColor = vbWhite
                .CellForeColor = vbRed
            ElseIf ctmp = "97" Then
                .CellBackColor = vbYellow   'automatisch kalkulierte
                .CellForeColor = vbBlue
            
            ElseIf ctmp = "95" Then
                .CellBackColor = vbBlue   'nicht geliefert
                .CellForeColor = vbBlack
                
            ElseIf ctmp = "94" Then
                .CellBackColor = glfarbe(0)   'Preisaktion in Vorbereitung
                .CellForeColor = vbBlue
                
            ElseIf ctmp = "93" Then
                .CellBackColor = vbWhite  'Preisaktion aktiv
                .CellForeColor = vbGreen
                
            ElseIf ctmp = "92" Then
                .CellBackColor = &H80000012  'seit 2 Jahren oder noch nie verkauft
                .CellForeColor = vbWhite
                
            Else
                .CellBackColor = glfarbe(0)
                .CellForeColor = vbBlack
            End If
            
        Else
            .Row = lrow
            .Col = iSpalte
            .CellBackColor = glfarbe(0)
            .CellForeColor = vbBlack
        End If
        
    End With
    
    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "FaerbenFlex"
    Fehler.gsFehlertext = "Beim Faerben der Tabelle ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Public Function WhatIsAwm(cartikel As String) As String
    On Error GoTo LOKAL_ERROR
    Dim sSQL As String
    Dim rs As Recordset
    
    WhatIsAwm = 0
    
    sSQL = "select Awm from artikel where artnr = " & cartikel
    Set rs = gdBase.OpenRecordset(sSQL)
    If Not rs.EOF Then
        If Not IsNull(rs!AWM) Then
            WhatIsAwm = rs!AWM
        End If
    End If
    rs.Close: Set rs = Nothing
    
    
    
    Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "WhatIsAwm"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function WhatIsAwmKU(cKunde As String) As String
    On Error GoTo LOKAL_ERROR
    Dim sSQL As String
    Dim rs As Recordset
    
    WhatIsAwmKU = 0
    
    sSQL = "select Awm from kunden where Kundnr = " & cKunde
    Set rs = gdBase.OpenRecordset(sSQL)
    If Not rs.EOF Then
        If Not IsNull(rs!AWM) Then
            WhatIsAwmKU = rs!AWM
        End If
    End If
    rs.Close: Set rs = Nothing
    
    
    
    Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "WhatIsAwmKU"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Sub OpenDrawer(aDeviceName, cEscapeSequenz)
    On Error GoTo LOKAL_ERROR
    
    Dim lhPrinter As Long
    Dim lReturn As Long
    Dim lpcWritten As Long
    Dim lDoc As Long
    Dim MyDocInfo As DOCINFO
    Dim lSize As Long
    Dim aDat As String
    Dim ctmp As String
    Dim cESC As String
    Dim llen As Long
    
    If gbNOBONDRUCKER = True Then
        Exit Sub
    End If
    
    If gbDebug Then
        MsgBox "Drucker = " & aDeviceName
        llen = Len(cEscapeSequenz)
        cESC = ""
        If llen > 2 Then
            ctmp = Trim$(Str$(Asc(Mid(cEscapeSequenz, 1, 1))))
            cESC = cESC & "Chr(" & ctmp & ") + "
            ctmp = Trim$(Str$(Asc(Mid(cEscapeSequenz, 2, 1))))
            cESC = cESC & "Chr(" & ctmp & ") + "
            ctmp = Trim$(Str$(Asc(Mid(cEscapeSequenz, 3, 1))))
            cESC = cESC & "Chr(" & ctmp & ") + "
            ctmp = Trim$(Str$(Asc(Mid(cEscapeSequenz, 4, 1))))
            cESC = cESC & "Chr(" & ctmp & ") + "
            ctmp = Trim$(Str$(Asc(Mid(cEscapeSequenz, 5, 1))))
            cESC = cESC & "Chr(" & ctmp & ") "
            MsgBox "Öffne Kasse mit = " & cESC
        Else
            ctmp = Trim$(Str$(Asc(Mid(cEscapeSequenz, 1, 1))))
            cESC = cESC & "Chr(" & ctmp & ") + "
            ctmp = Trim$(Str$(Asc(Mid(cEscapeSequenz, 2, 1))))
            cESC = cESC & "Chr(" & ctmp & ") "
            MsgBox "Schneide Kassenbon mit = " & cESC
        End If
    End If
    lReturn = 0
    lReturn = OpenPrinter(aDeviceName, lhPrinter, 0)
    If lReturn = 0 And gbDebug Then
        MsgBox "Drucker " & aDeviceName & " nicht gefunden!", vbCritical, "STOP!"
        Exit Sub
    End If
    
    MyDocInfo.pDocName = "Open Drawer"
    MyDocInfo.pOutputFile = vbNullString
    MyDocInfo.pDatatype = vbNullString
    
    lDoc = 0
    lDoc = StartDocPrinter(lhPrinter, 1, MyDocInfo)
    If lDoc = 0 And gbDebug Then
        MsgBox "Fehler bei StartDocPrinter (" & Trim$(Str$(lhPrinter)) & ") "
        Exit Sub
    End If
    
    lReturn = 0
    lReturn = StartPagePrinter(lhPrinter)
    If lReturn = 0 And gbDebug Then
        MsgBox "Fehler bei StartPagePrinter (" & Trim$(Str$(lhPrinter)) & ") "
        Exit Sub
    End If
    
    aDat = cEscapeSequenz
    lSize = Len(aDat)
    lReturn = 0
    lReturn = WritePrinter(lhPrinter, ByVal aDat, lSize, lpcWritten)
    If lReturn = 0 And gbDebug Then
        MsgBox "Fehler bei WritePrinter (" & Trim$(Str$(lhPrinter)) & " / " & aDat & ") "
        Exit Sub
    End If
    
    
    If lReturn = 0 And gbDebug Then
        MsgBox "WritePrinter fehlerhaft"
        Exit Sub
    End If
    lReturn = 0
    lReturn = EndPagePrinter(lhPrinter)
    If lReturn = 0 And gbDebug Then
        MsgBox "Fehler bei EndPagePrinter (" & Trim$(Str$(lhPrinter)) & ") "
        Exit Sub
    End If
    
    If lReturn = 0 And gbDebug Then
        MsgBox "EndPagePrinter fehlerhaft"
        Exit Sub
    End If
        
    lReturn = 0
    lReturn = EndDocPrinter(lhPrinter)
    If lReturn = 0 And gbDebug Then
        MsgBox "Fehler bei EndDocPrinter (" & Trim$(Str$(lhPrinter)) & ") "
        Exit Sub
    End If
    
    If lReturn = 0 And gbDebug Then
        MsgBox "EndDocPrinter fehlerhaft"
        Exit Sub
    End If
    
    lReturn = 0
    lReturn = ClosePrinter(lhPrinter)
    If lReturn = 0 And gbDebug Then
        MsgBox "Fehler bei ClosePrinter (" & Trim$(Str$(lhPrinter)) & ") "
        Exit Sub
    End If
    
    If lReturn = 0 And gbDebug Then
        MsgBox "ClosePrinter fehlerhaft"
        Exit Sub
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "OpenDrawer"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Public Sub OpenDrawer3(aDeviceName As String, cDruckZeile() As String, lAnzZeile As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim lhPrinter As Long
    Dim lReturn As Long
    Dim lpcWritten As Long
    Dim lDoc As Long
    Dim MyDocInfo As DOCINFO
    Dim lSize As Long
    Dim aDat As String
    Dim ctmp As String
    Dim cESC As String
    Dim llen As Long
    Dim lcount As Long
    Dim lStart As Long
    Dim lAktuell As Long
    
    If gbNOBONDRUCKER = True Then
        Exit Sub
    End If
    
    
    lReturn = 0
    lReturn = OpenPrinter(aDeviceName, lhPrinter, 0)
    If lReturn = 0 Then
        MsgBox "Drucker " & aDeviceName & " nicht gefunden!", vbCritical, "STOP!"
        Exit Sub
    End If
    DoEvents
    
'''    cEscapeSequenz = Chr(29) & Chr(104) & Chr(40)
'''    OpenDrawer aDeviceName, cEscapeSequenz
    
    
    If gsBONFONTNAME = "Standard" Then
        
        MyDocInfo.pDocName = "Open Drawer"
        MyDocInfo.pOutputFile = vbNullString
        MyDocInfo.pDatatype = vbNullString
        
        
        lDoc = 0
        lDoc = StartDocPrinter(lhPrinter, 1, MyDocInfo)
        If lDoc = 0 Then
            MsgBox "Fehler bei StartDocPrinter (" & Trim$(Str$(lhPrinter)) & ") "
            Exit Sub
        End If
        DoEvents
        
        lReturn = 0
        lReturn = StartPagePrinter(lhPrinter)
        If lReturn = 0 Then
            MsgBox "Fehler bei StartPagePrinter (" & Trim$(Str$(lhPrinter)) & ") "
            Exit Sub
        End If
        DoEvents
        
        
        If InStr(UCase$(aDeviceName), "ZEBRA") > 0 Then
            lStart = Timer
            Do
                lAktuell = Timer
            Loop While lAktuell < lStart + 10
        End If
        
        For lcount = 1 To lAnzZeile
        
            If glZeichenAnzahlBon = 32 Then
                aDat = Space$(5) & cDruckZeile(lcount)
            Else
                aDat = cDruckZeile(lcount)
            End If
        

            lSize = Len(aDat)
            lReturn = 0
            lReturn = WritePrinter(lhPrinter, ByVal aDat, lSize, lpcWritten)
            
            If InStr(UCase$(aDeviceName), "ZEBRA") > 0 Then
                lStart = Timer
                Do
                    lAktuell = Timer
                Loop While lAktuell < lStart + 2
            End If
            If lReturn = 0 Then
                MsgBox "Fehler bei WritePrinter (" & Trim$(Str$(lhPrinter)) & " / " & aDat & ") "
                Exit Sub
            End If
            DoEvents
            
            If lReturn = 0 And gbDebug Then
                MsgBox "WritePrinter fehlerhaft"
                Exit Sub
            End If
        
        Next lcount
           
        If InStr(UCase$(aDeviceName), "ZEBRA") > 0 Then
            lStart = Timer
            Do
                lAktuell = Timer
            Loop While lAktuell < lStart + 10
        End If
        
        lReturn = 0
        lReturn = EndPagePrinter(lhPrinter)
        If lReturn = 0 Then
            MsgBox "Fehler bei EndPagePrinter (" & Trim$(Str$(lhPrinter)) & ") "
            Exit Sub
        End If
        DoEvents
        
        If lReturn = 0 And gbDebug Then
            MsgBox "EndPagePrinter fehlerhaft"
            Exit Sub
        End If
            
        lReturn = 0
        lReturn = EndDocPrinter(lhPrinter)
        If lReturn = 0 Then
            MsgBox "Fehler bei EndDocPrinter (" & Trim$(Str$(lhPrinter)) & ") "
            Exit Sub
        End If
        DoEvents
        
        If lReturn = 0 And gbDebug Then
            MsgBox "EndDocPrinter fehlerhaft"
            Exit Sub
        End If
        
        lReturn = 0
        lReturn = ClosePrinter(lhPrinter)
        If lReturn = 0 Then
            MsgBox "Fehler bei ClosePrinter (" & Trim$(Str$(lhPrinter)) & ") "
            Exit Sub
        End If
        DoEvents
        
        If lReturn = 0 And gbDebug Then
            MsgBox "ClosePrinter fehlerhaft"
            Exit Sub
        End If
        
        
    Else
    
       
        
    
        Dim cEscapeSequenz As String

        Printer.FontName = gsBONFONTNAME
        Printer.FontSize = gdBONFONTSIZE

        For lcount = 1 To lAnzZeile
            ctmp = cDruckZeile(lcount)



            If ctmp <> gcInit Then

                KonvertAsciiAnsi ctmp
                If Right(ctmp, 2) = vbCrLf Then
                    ctmp = Left(ctmp, Len(ctmp) - 2)
                End If
                Printer.Print ctmp

            End If

        Next lcount

        Printer.Print
        Printer.Print
        Printer.Print
        
        

        Printer.EndDoc
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "OpenDrawer3"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub


Public Sub OpenDrawer4(aDeviceName As String, cDruckZeile() As String, lAnzZeile As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    Dim cEscapeSequenz As String
    Dim ctmp As String
    
    Dim bError As Boolean
    Dim iStufe As Integer
    
    If gbNOBONDRUCKER = True Then
        Exit Sub
    End If
    
    
    bError = False
    
    'Einstellungen für Drucker vornehmen
    iStufe = 1
    Printer.FontName = "15 CPI"
    Printer.FontSize = 9.5
    Printer.FontBold = False

    If bError = True Then
        iStufe = 2

        Printer.FontName = "Courier New"
        Printer.FontSize = 10
        Printer.FontBold = True
    End If
    
    For lcount = 1 To lAnzZeile
        ctmp = Space$(3) & cDruckZeile(lcount)
        KonvertAsciiAnsi ctmp
        If Right(ctmp, 2) = vbCrLf Then
            ctmp = Left(ctmp, Len(ctmp) - 2)
        End If
        Printer.Print ctmp
    Next lcount
    
    Printer.Print
    Printer.Print
    Printer.Print
    
    Printer.EndDoc
    


Exit Sub
LOKAL_ERROR:
    If err.Number = 380 Then
        If iStufe = 1 Then
            bError = True
            Resume Next
        Else
            Fehler.gsDescr = err.Description
            Fehler.gsNumber = err.Number
            Fehler.gsFormular = "Modul3"
            Fehler.gsFunktion = "OpenDrawer4"
            Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
            
            Fehlermeldung1
        End If
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul3"
        Fehler.gsFunktion = "OpenDrawer4"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub

Public Sub OpenDrawer4Groß(aDeviceName As String, cDruckZeile() As String, lAnzZeile As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim lcount As Long
    Dim cEscapeSequenz As String
    Dim ctmp As String
    Dim sSQL As String
    
    If gbNOBONDRUCKER = True Then
        Exit Sub
    End If
    
    
    'Einstellungen für Drucker vornehmen

    Printer.FontName = "Courier New"
    Printer.FontSize = 48
    Printer.FontBold = True
    
    
    For lcount = 1 To lAnzZeile
        ctmp = Space(1) & cDruckZeile(lcount)
        KonvertAsciiAnsi ctmp

        Printer.Print ctmp
    Next lcount
    
    Printer.Print
    Printer.Print
    Printer.Print
    
    Printer.EndDoc

Exit Sub
LOKAL_ERROR:

    If err.Number = 482 Then
    
        sSQL = "Update KASSEIN Set nografik = True"
        gdBase.Execute sSQL, dbFailOnError
        gbNoGrafik = True
    
    Else
        
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul3"
        Fehler.gsFunktion = "OpenDrawer4Groß"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
    
End Sub

Public Function SetDefaultPrinter(PrName As String)
    On Error GoTo LOKAL_ERROR
 
 ' Parameter: Druckername
 ' Rückgabewert: Erfolg der Aktion

    Dim Buffer As String
    Dim RW As Integer
    Dim Tmp As String

 
 
 
   Buffer = String(255, 0)
   RW = GetProfileString(ByVal "devices", ByVal PrName, ByVal "", Buffer, Len(Buffer))
   If RW <= 0 Then
     SetDefaultPrinter = False
     Exit Function
   Else
     Tmp = PrName & "," & Mid(Buffer, 1, RW)
   End If
   
   
   
 ' Standarddrucker setzen
   RW = WriteProfileString(ByVal "Windows", ByVal "Device", ByVal Tmp)
   If RW <> 1 Then
     SetDefaultPrinter = False
     Exit Function
   End If
   
   
   
   
 ' und mitteilen, daß sich die WIN.INI geändert hat
'   RW = SendMessage(HWND_BROADCAST, WM_WININICHANGE, 0&, 0&)
   SetDefaultPrinter = True
   
Ex:
   Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "SetDefaultPrinter"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    SetDefaultPrinter = False
    Resume Ex
End Function
Public Sub insert_Display(sZeile1 As String, sZeile2 As String, Optional sSatz As String, Optional sZSUM As String, Optional ilfnr As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    Dim sBezeich        As String
    Dim iAnzahl         As Integer
    Dim sEPR            As String
    Dim sGPR            As String
    Dim sZwiSUM         As String
    Dim rsrs            As DAO.Recordset
    Dim iLokallfnr      As Integer
    Dim Task$
    
    iLokallfnr = 0
    iAnzahl = 0
    sBezeich = ""
    sEPR = "0"
    sGPR = "0"
    sZwiSUM = "0"
    
    sZeile1 = SwapStr(sZeile1, "'", "")
    sZeile2 = SwapStr(sZeile2, "'", "")
    
    If sSatz <> "" Then
        'zerhacken
        iAnzahl = Val(Trim(Mid(sSatz, 1, 5)))
        sBezeich = Trim(Mid(sSatz, 14, 35))
        sEPR = Trim(Mid(sSatz, 50, 9))
        sGPR = Trim(Mid(sSatz, 60, 9))
    End If
    
    If sZSUM <> "" Then
        sZwiSUM = Trim(sZSUM)
    End If
    
    sSQL = "select max(lfnr)as maxlfnr from DISPLAY_" & srechnertab
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxlfnr) Then
             iLokallfnr = rsrs!maxlfnr
        End If
    End If
    rsrs.Close
    iLokallfnr = iLokallfnr + 1
    
    
    
    
    
    If ilfnr = -1 Then
        iLokallfnr = 0
        
        sSQL = "Delete from DISPLAY_" & srechnertab
        gdBase.Execute sSQL, dbFailOnError
        
'        Task = Shell(App.Path & "\Display.exe", 1) 'Display öffnen
        
    End If
    
    
    
    
'    sZeile1 = SwapStr(sZeile1, "'", "")
'    sZeile2 = SwapStr(sZeile2, "'", "")
    sBezeich = SwapStr(sBezeich, "'", "")
    
    sSQL = "Insert into DISPLAY_" & srechnertab & "("
    sSQL = sSQL & "  Zeile1 "
    sSQL = sSQL & ", Zeile2 "
    sSQL = sSQL & ", LFNR "
    sSQL = sSQL & ", Anzahl "
    sSQL = sSQL & ", BEZEICH "
    sSQL = sSQL & ", EPR "
    sSQL = sSQL & ", GPR "
    sSQL = sSQL & ", ZSUM "
    sSQL = sSQL & ") values "
    sSQL = sSQL & "('" & sZeile1 & "'"
    sSQL = sSQL & ",'" & sZeile2 & "'"
    sSQL = sSQL & "," & iLokallfnr & ""
    sSQL = sSQL & "," & iAnzahl & ""
    sSQL = sSQL & ",'" & sBezeich & "'"
    sSQL = sSQL & ",'" & sEPR & "'"
    sSQL = sSQL & ",'" & sGPR & "'"
    sSQL = sSQL & ",'" & sZwiSUM & "'"
    sSQL = sSQL & ") "
    gdBase.Execute sSQL, dbFailOnError
    
    
    schreibeDatei iLokallfnr, sZeile1, sZeile2, sBezeich, iAnzahl, sEPR, sGPR, sZwiSUM
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "insert_Display"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Public Sub schreibeDatei(ilfnr As Integer, sZeile1 As String, sZeile2 As String, sBezeich As String, iAnzahl As Integer, sEPR As String, sGPR As String, sZwiSUM As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cPfad As String
    Dim cSatz As String
    Dim iFileNr As Integer
    Dim lPos As Long
    
    schreibeProtokollMonitorTXT "lfnr: " & CStr(ilfnr) & " " & sZeile1 & " " & sZeile2
    
    
    
    
    cPfad = gcPfad  'Anwendungspfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    cPfad = cPfad & "\DISPLAY\"
    
    
    Kill cPfad & "D" & ilfnr & ".txt"

    iFileNr = FreeFile
    Open cPfad & "D" & ilfnr & ".txt" For Binary As #iFileNr
    
    cSatz = ""
    cSatz = cSatz & sZeile1 & ";"
    cSatz = cSatz & sZeile2 & ";"
    cSatz = cSatz & ilfnr & ";"
    cSatz = cSatz & iAnzahl & ";"
    cSatz = cSatz & sBezeich & ";"
    cSatz = cSatz & sEPR & ";"
    cSatz = cSatz & sGPR & ";"
    cSatz = cSatz & sZwiSUM
    
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
        Fehler.gsFormular = "Modul3"
        Fehler.gsFunktion = "schreibeDatei"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
        Fehlermeldung1
    End If
End Sub
Public Sub ZeigeKundenDisplay(cZeile1 As String, cZeile2 As String, Optional sSatz As String, Optional sZwSUM As String, Optional ilfnr As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim aDeviceName As String
    Dim cEscapeSequenz As String
    
    If gbSPY Then
        frmWKL20!Winsock22.senddata cZeile1 & vbCrLf
        frmWKL20!Winsock22.senddata cZeile2 & vbCrLf
    End If
    
    If gbZweitMoni Then
    
        If gbZweitMoniMinimieren = True Then
    
            Dim hwnd&
            Dim Y As String
            Dim result&
            Dim Title$
    
            Y = "Ihre Kundeninformationen"
    
            hwnd = GetWindow(frmWKL00.hwnd, GW_HWNDFIRST)
    
            Do
                result = GetWindowTextLength(hwnd) + 1
                Title = Space(result)
                result = GetWindowText(hwnd, Title, result)
                Title = Left$(Title, Len(Title) - 1)
    
                If InStr(1, Title, Y) Then
                
                    If ilfnr = -1 Then
                        ShowWindow hwnd, vbMinimizedNoFocus
                    Else
    
                        ShowWindow hwnd, vbMaximizedFocus
                    End If
    
    
                End If
    
                hwnd = GetWindow(hwnd, GW_HWNDNEXT)
            Loop Until hwnd = 0
            
        End If
    
        insert_Display cZeile1, cZeile2, sSatz, sZwSUM, ilfnr
    End If
    
    If Not gbDisplay Then
        Exit Sub
    End If
           
    KonvertAnsiAscii cZeile1
    KonvertAnsiAscii cZeile2
    If gbDisplaySeriell Then
        Select Case gcDisplay
            Case Is = "Epson"
                frmWKL20!MSComm3.CommPort = giDisplaySeriellComPort
                frmWKL20!MSComm3.Settings = "9600,N,8,1"
                frmWKL20!MSComm3.InputLen = 0
                frmWKL20!MSComm3.PortOpen = True
            
                frmWKL20!MSComm3.Output = Chr$(12) & Chr$(11) & cZeile1 & Chr$(31) & Chr$(36) & Chr$(1) & Chr$(2) & cZeile2
                frmWKL20!MSComm3.PortOpen = False
            
            Case Is = "JarlTech"
                frmWKL20!MSComm3.CommPort = giDisplaySeriellComPort
                frmWKL20!MSComm3.Settings = "9600,N,8,1"
                frmWKL20!MSComm3.InputLen = 0
                frmWKL20!MSComm3.PortOpen = True
                frmWKL20!MSComm3.Output = Chr$(26) & Chr$(27) & Chr$(96) & Chr$(1) & cZeile1 & Chr$(10) & Chr$(13) & cZeile2
                frmWKL20!MSComm3.PortOpen = False
                
            Case Is = "Peacock"
                frmWKL20!MSComm3.CommPort = giDisplaySeriellComPort
                frmWKL20!MSComm3.Settings = "9600,N,8,1"
                frmWKL20!MSComm3.InputLen = 0
                frmWKL20!MSComm3.PortOpen = True
                frmWKL20!MSComm3.Output = Chr$(12) & Chr$(11) & cZeile1 & Chr$(31) & Chr$(36) & Chr$(1) & Chr$(2) & cZeile2
                frmWKL20!MSComm3.PortOpen = False
                
            Case Is = "Aures neu"
                frmWKL20!MSComm3.Output = Chr$(12) & Chr$(11) & cZeile1 & Chr$(31) & Chr$(36) & Chr$(1) & Chr$(2) & cZeile2
                
            Case Is = "Sango"

                cZeile1 = cZeile1 & Space(20 - Len(cZeile1))
                cZeile2 = cZeile2 & Space(20 - Len(cZeile2))
                frmWKL20!MSComm3.Output = Chr$(12) & Chr$(11) & cZeile1 & cZeile2
            Case Is = "Peacock (alt)"
            
                frmWKL20!MSComm3.CommPort = giDisplaySeriellComPort
                frmWKL20!MSComm3.Settings = "9600,E,8,1"
                frmWKL20!MSComm3.InputLen = 0
                frmWKL20!MSComm3.PortOpen = True
                frmWKL20!MSComm3.Output = Chr$(13)
                frmWKL20!MSComm3.Output = Chr$(10)
                cZeile1 = cZeile1 & Space(20 - Len(cZeile1))
                cZeile2 = cZeile2 & Space(20 - Len(cZeile2))
                frmWKL20!MSComm3.Output = cZeile1 & cZeile2
                frmWKL20!MSComm3.PortOpen = False
           
        End Select
        
        Exit Sub
    Else
        aDeviceName = gcBonDrucker
        'Init Drucker
        cEscapeSequenz = Chr$(27) & Chr$(64)
        'Anzeigecursor aus
        cEscapeSequenz = cEscapeSequenz & Chr$(31) & Chr$(67) & "0"
        'Drucker aus, Display an
        cEscapeSequenz = cEscapeSequenz & Chr$(27) & Chr$(61) & Chr$(2)
        'Clear Display, Move Home
        cEscapeSequenz = cEscapeSequenz & Chr$(27) & Chr$(12) & Chr$(27) & Chr$(11)
        'Move Cursor to Pos 1, Line 1
        cEscapeSequenz = cEscapeSequenz & Chr$(31) & Chr$(36) & Chr$(1) & Chr$(1)
        'Show Message Zeile 1
        cEscapeSequenz = cEscapeSequenz & cZeile1
        'Move Cursor to Pos 1, Line 2
        cEscapeSequenz = cEscapeSequenz & Chr$(31) & Chr$(36) & Chr$(1) & Chr$(2)
        'Show Message Zeile 2
        cEscapeSequenz = cEscapeSequenz & cZeile2
        
        'Message über Drucker an Display senden
        OpenDrawer aDeviceName, cEscapeSequenz
    End If
    
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 8005 Then
        Resume Next
    ElseIf err.Number = 8002 Then
        MsgBox "Der COM - Port " & giDisplaySeriellComPort & " steht nicht zur Verfügung.", vbInformation, "Winkiss Hinweis:"
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul3"
        Fehler.gsFunktion = "ZeigeKundenDisplay"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Public Sub ZeigeKundenDisplay_forTest(cZeile1 As String, cZeile2 As String, Optional sSatz As String, Optional sZwSUM As String, Optional ilfnr As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim aDeviceName As String
    Dim cEscapeSequenz As String
    
    If Not gbDisplay Then
        Exit Sub
    End If
           
    KonvertAnsiAscii cZeile1
    KonvertAnsiAscii cZeile2
    
    If gbDisplaySeriell Then
        Select Case gcDisplay
            Case Is = "Epson"
                frmWKL50!MSComm3.CommPort = giDisplaySeriellComPort
                frmWKL50!MSComm3.Settings = "9600,N,8,1"
                frmWKL50!MSComm3.InputLen = 0
                
                frmWKL50!MSComm3.PortOpen = True
                frmWKL50!MSComm3.Output = Chr$(12) & Chr$(11) & cZeile1 & Chr$(31) & Chr$(36) & Chr$(1) & Chr$(2) & cZeile2
                frmWKL50!MSComm3.PortOpen = False
            
            Case Is = "JarlTech"
                frmWKL50!MSComm3.CommPort = giDisplaySeriellComPort
                frmWKL50!MSComm3.Settings = "9600,N,8,1"
                frmWKL50!MSComm3.InputLen = 0
                frmWKL50!MSComm3.PortOpen = True
                frmWKL50!MSComm3.Output = Chr$(26) & Chr$(27) & Chr$(96) & Chr$(1) & cZeile1 & Chr$(10) & Chr$(13) & cZeile2
                frmWKL50!MSComm3.PortOpen = False
                
            Case Is = "Peacock"
                frmWKL50!MSComm3.CommPort = giDisplaySeriellComPort
                frmWKL50!MSComm3.Settings = "9600,N,8,1"
                frmWKL50!MSComm3.InputLen = 0
                frmWKL50!MSComm3.PortOpen = True
                frmWKL50!MSComm3.Output = Chr$(12) & Chr$(11) & cZeile1 & Chr$(31) & Chr$(36) & Chr$(1) & Chr$(2) & cZeile2
                frmWKL50!MSComm3.PortOpen = False
            Case Is = "Aures neu"
            
                frmWKL50!MSComm3.CommPort = giDisplaySeriellComPort
                frmWKL50!MSComm3.Settings = "9600,N,8,1"
                frmWKL50!MSComm3.InputLen = 0
                frmWKL50!MSComm3.PortOpen = True
                frmWKL50!MSComm3.Output = Chr$(12) & Chr$(11) & cZeile1 & Chr$(31) & Chr$(36) & Chr$(1) & Chr$(2) & cZeile2
                frmWKL50!MSComm3.PortOpen = False
            Case Is = "Sango"

                cZeile1 = cZeile1 & Space(20 - Len(cZeile1))
                cZeile2 = cZeile2 & Space(20 - Len(cZeile2))
                
                frmWKL50!MSComm3.CommPort = giDisplaySeriellComPort
                frmWKL50!MSComm3.Settings = "9600,N,8,1"
                frmWKL50!MSComm3.InputLen = 0
                frmWKL50!MSComm3.PortOpen = True
                frmWKL50!MSComm3.Output = Chr$(12) & Chr$(11) & cZeile1 & cZeile2
                frmWKL50!MSComm3.PortOpen = False
                
            Case Is = "Peacock (alt)"
            
                frmWKL50!MSComm3.CommPort = giDisplaySeriellComPort
                frmWKL50!MSComm3.Settings = "9600,E,8,1"
                frmWKL50!MSComm3.InputLen = 0
                frmWKL50!MSComm3.PortOpen = True
                
                PauseSi 0.5
                frmWKL50!MSComm3.Output = Chr$(13)
                frmWKL50!MSComm3.Output = Chr$(10)
                
                cZeile1 = cZeile1 & Space(20 - Len(cZeile1))
                cZeile2 = cZeile2 & Space(20 - Len(cZeile2))
                
                frmWKL50!MSComm3.Output = cZeile1 & cZeile2
                frmWKL50!MSComm3.PortOpen = False
                
        End Select
        
        Exit Sub
    Else
        aDeviceName = gcBonDrucker
        'Init Drucker
        cEscapeSequenz = Chr$(27) & Chr$(64)
        'Anzeigecursor aus
        cEscapeSequenz = cEscapeSequenz & Chr$(31) & Chr$(67) & "0"
        'Drucker aus, Display an
        cEscapeSequenz = cEscapeSequenz & Chr$(27) & Chr$(61) & Chr$(2)
        'Clear Display, Move Home
        cEscapeSequenz = cEscapeSequenz & Chr$(27) & Chr$(12) & Chr$(27) & Chr$(11)
        'Move Cursor to Pos 1, Line 1
        cEscapeSequenz = cEscapeSequenz & Chr$(31) & Chr$(36) & Chr$(1) & Chr$(1)
        'Show Message Zeile 1
        cEscapeSequenz = cEscapeSequenz & cZeile1
        'Move Cursor to Pos 1, Line 2
        cEscapeSequenz = cEscapeSequenz & Chr$(31) & Chr$(36) & Chr$(1) & Chr$(2)
        'Show Message Zeile 2
        cEscapeSequenz = cEscapeSequenz & cZeile2
        
        'Message über Drucker an Display senden
        OpenDrawer aDeviceName, cEscapeSequenz
    End If
    
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 8005 Then
        Resume Next
    ElseIf err.Number = 8002 Then
        MsgBox "Der COM - Port " & giDisplaySeriellComPort & " steht nicht zur Verfügung.", vbInformation, "Winkiss Hinweis:"
    ElseIf err.Number = 8018 Then
        frmWKL50!MSComm3.PortOpen = False
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul3"
        Fehler.gsFunktion = "ZeigeKundenDisplay_forTest"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Public Sub in_Kiss_Lieferavis_wandeln(sFormat As String, sLinr As String)
    On Error GoTo LOKAL_ERROR
    
    Select Case sFormat
        Case "BUDNI"
            Budni_Lieferavis_wandeln sLinr

    End Select
    
Exit Sub

LOKAL_ERROR:
     Fehler.gsDescr = err.Description
     Fehler.gsNumber = err.Number
     Fehler.gsFormular = "Modul3"
     Fehler.gsFunktion = "in_Kiss_Lieferavis_wandeln"
     Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
     
     Fehlermeldung1
End Sub
Private Sub Budni_Lieferavis_wandeln(sLinr As String)
    On Error GoTo LOKAL_ERROR
    
    Dim i           As Integer
    Dim cPfad       As String
    Dim sSQL        As String
    Dim sName       As String
    Dim cDatum      As String
    Dim sAuftragsnr As String
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    cPfad = cPfad & "In"
    
    frmWKL00.File2.Path = cPfad
    frmWKL00.File2.Pattern = "BUDNIDESADV*.csv"
    frmWKL00.File2.Refresh
    
    cPfad = cPfad & "\"
    
    If frmWKL00.File2.ListCount = 0 Then Exit Sub
    
    For i = 0 To frmWKL00.File2.ListCount - 1
    
        sName = frmWKL00.File2.list(i)
        
        sAuftragsnr = Mid(sName, 30, 6)
        
        import_in_Desadv cPfad & sName, sLinr, sAuftragsnr

        Kill cPfad & frmWKL00.File2.list(i)
    Next i

Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "Budni_Lieferavis_wandeln"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub import_in_Desadv(sPfadundDatei As String, sLinr As String, sAuftragsnummer As String)
On Error GoTo LOKAL_ERROR

    Dim lPosEnde        As Long
    Dim cEinzelsatz     As String
    Dim lLenfil         As Long
    Dim lposSemi        As Long
    Dim lposSemiEnde    As Long
    Dim cWert           As String
    Dim lfnr1           As Long
    Dim cPreis          As String
    Dim lPos            As String
    Dim rsrs            As Recordset
    Dim iFileNr         As Integer
    Dim cSatz1          As String
    Dim dWert           As Double
    Dim lcount          As Long
    Dim sSQL            As String
    
    If sAuftragsnummer = "Erstbe" Then sAuftragsnummer = 100000
    If sAuftragsnummer = "Tester" Then sAuftragsnummer = 100000
    If sAuftragsnummer = "WPR Zu" Then sAuftragsnummer = 100000
    If IsNumeric(sAuftragsnummer) = False Then sAuftragsnummer = 100000

            
    lPos = 1
    lPosEnde = 1
    lposSemiEnde = 1
    
    sSQL = "Delete * from DESADV where AUFTRAGSNR = " & sAuftragsnummer
    gdBase.Execute sSQL, dbFailOnError
    
    Set rsrs = gdBase.OpenRecordset("DESADV")

    iFileNr = FreeFile
    Open sPfadundDatei For Binary As #iFileNr
    If LOF(iFileNr) > 0 Then
    
        cSatz1 = Space$(LOF(iFileNr))
        Get #iFileNr, 1, cSatz1
    
        lLenfil = Len(cSatz1)
        lPosEnde = InStr(lPos, cSatz1, vbCrLf)
        lPos = lPos + lPosEnde - lPos + 2 'Kopfzeile überspringen
        
        lcount = 0
        
        Do
            lcount = lcount + 1
'            anzeige "normal", CStr(lcount), lblAnzeige
            
            lPosEnde = InStr(lPos, cSatz1, vbCrLf)
            cEinzelsatz = Mid(cSatz1, lPos, lPosEnde)
            lPos = lPos + lPosEnde - lPos + 2
            lposSemi = 1
            
            rsrs.AddNew

            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
            '1. überspringen

            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
            '2. überspringen
            
            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbTab): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
            rsrs!LIBESNR = CStr(Val(cWert))
            lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbCrLf): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
            rsrs!Menge = Val(cWert)
            
            rsrs!linr = sLinr
            rsrs!AUFTRAGSNR = sAuftragsnummer
            rsrs.Update
        Loop While lLenfil >= lPos
        
    End If
    
    Close iFileNr
    rsrs.Close: Set rsrs = Nothing
    
    sSQL = "Update DESADV d inner join Artlief a on d.libesnr = a.libesnr "
    sSQL = sSQL & " set d.artnr = a.artnr "
    sSQL = sSQL & " where a.linr = " & sLinr
    sSQL = sSQL & " and a.RKZ = 'N' "
    sSQL = sSQL & " and d.AUFTRAGSNR = " & sAuftragsnummer
    gdBase.Execute sSQL, dbFailOnError

Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul3"
    Fehler.gsFunktion = "import_in_Desadv"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
'    Resume Next
End Sub


