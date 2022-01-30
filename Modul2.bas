Attribute VB_Name = "Modul2"

Option Explicit
Dim gllfnr       As Long
Dim newPreislage() As Preislage
Dim byteanzPreisl   As Byte

Public Function FRAGENeueNachrichten() As Boolean
    On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsrs As DAO.Recordset
    
    FRAGENeueNachrichten = False
    
    If gbKL_LIVENACHRICHTEN = True Then
    
        If fTestLogin_SQLDABA_Error = 0 Then 'ist alles OK? Datenbank erreichbar?
            'alles okay
        Else
        
            schreibeProtokollVPNTXT "Unterbrechung"
        
            Dim sTemp As String
            sTemp = "Bitte starten Sie diesen Rechner neu" & vbCrLf
            sTemp = sTemp & "oder schließen Sie das Schloss und starten Sie WinKiss neu."
        
            MsgBox sTemp, vbCritical + vbOKOnly, "Datenbank nicht erreichbar"
            Exit Function
        End If
        
        Dim stConnect As String
        
        If gsKL_DSN <> "" Then
            stConnect = "ODBC;DSN=" & gsKL_DSN & ";UID=" & gsKL_BENUTZER & ";PWD=" & gsKL_PASSWORT & ""
        Else
            stConnect = "ODBC;DRIVER=SQL Server;SERVER=" & gsKL_ADRESSE & ";DATABASE=" & gsKL_DATENBANKNAME & ";UID=" & gsKL_BENUTZER & ";PWD=" & gsKL_PASSWORT & ""
        End If
        
        Dim dbEAN As DAO.Database
        Set dbEAN = OpenDatabase(gsKL_DATENBANKNAME, dbDriverNoPrompt, False, stConnect)
        
        Dim lMaxNaNr As Long
        lMaxNaNr = ermMaxNachrichtenNummer
        
        cSQL = "Select *  from NACHRICHTEN where lfnr > " & lMaxNaNr
        Set rsrs = dbEAN.OpenRecordset(cSQL)
        If Not rsrs.EOF Then
            
            FRAGENeueNachrichten = True
            
        End If
        rsrs.Close: Set rsrs = Nothing
        
        dbEAN.Close
        
    End If
        
Exit Function
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "FrageNeueNachrichten"
    Fehler.gsFehlertext = "Es trat ein Fehler auf. "
    
    Fehlermeldung1
End Function
Public Sub ImportiereNeueNachrichten()
    On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsrs As DAO.Recordset
    
    Dim lLFNR As Long
    Dim sADATE As String
    Dim sAzeit As String
    Dim lAN_FILIALE As Long
    Dim sMESSAGETEXT As String
    Dim sBetreff As String
    Dim sABSENDER As String
    Dim bVERSENDET As Boolean
    Dim bANGEKOMMEN As Boolean
    Dim bGelesen As Boolean
    Dim sGELESEN_VONWEM As String
    Dim daGELESEN_ADATE As Date
    Dim sGELESEN_AZEIT As String
    Dim bFELD1 As Boolean
    Dim bFELD2 As Boolean
    Dim bFELD3 As Boolean
    Dim bFELD4 As Boolean
    Dim bFELD5 As Boolean
    Dim bFELD6 As Boolean
    Dim bFELD7 As Boolean
    Dim bFELD8 As Boolean
    Dim bFELD9 As Boolean
    Dim bFELD10 As Boolean
    Dim sINFO1 As String
    Dim sINFO2 As String
    Dim sINFO3 As String
    
    lLFNR = -1
    sADATE = CStr(DateValue(Now))
    sAzeit = ""
    lAN_FILIALE = 0
    sMESSAGETEXT = ""
    sBetreff = ""
    sABSENDER = ""
    bVERSENDET = False
    bANGEKOMMEN = False
    bGelesen = False
    sGELESEN_VONWEM = ""
    daGELESEN_ADATE = DateValue(Now)
    sGELESEN_AZEIT = ""
    bFELD1 = False
    bFELD2 = False
    bFELD3 = False
    bFELD4 = False
    bFELD5 = False
    bFELD6 = False
    bFELD7 = False
    bFELD8 = False
    bFELD9 = False
    bFELD10 = False
    sINFO1 = ""
    sINFO2 = ""
    sINFO3 = ""
    
''    cSQL = "delete from Nachrichten"
''    gdBase.Execute cSQL, dbFailOnError
''
    
    
    
    
    If gbKL_LIVENACHRICHTEN = True Then
    
        If fTestLogin_SQLDABA_Error = 0 Then 'ist alles OK? Datenbank erreichbar?
            'alles okay
        Else
        
            schreibeProtokollVPNTXT "Unterbrechung"
        
            Dim sTemp As String
            sTemp = "Bitte starten Sie diesen Rechner neu" & vbCrLf
            sTemp = sTemp & "oder schließen Sie das Schloss und starten Sie WinKiss neu."
        
            MsgBox sTemp, vbCritical + vbOKOnly, "Datenbank nicht erreichbar"
            Exit Sub
        End If
        
        Dim stConnect As String
        
        If gsKL_DSN <> "" Then
            stConnect = "ODBC;DSN=" & gsKL_DSN & ";UID=" & gsKL_BENUTZER & ";PWD=" & gsKL_PASSWORT & ""
        Else
            stConnect = "ODBC;DRIVER=SQL Server;SERVER=" & gsKL_ADRESSE & ";DATABASE=" & gsKL_DATENBANKNAME & ";UID=" & gsKL_BENUTZER & ";PWD=" & gsKL_PASSWORT & ""
        End If
        
        Dim dbEAN As DAO.Database
        Set dbEAN = OpenDatabase(gsKL_DATENBANKNAME, dbDriverNoPrompt, False, stConnect)
        
        Dim lMaxNaNr As Long
        lMaxNaNr = ermMaxNachrichtenNummer
        
        cSQL = "Select *  from NACHRICHTEN where lfnr > " & lMaxNaNr
        Set rsrs = dbEAN.OpenRecordset(cSQL)
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            Do While Not rsrs.EOF
            
            If Not IsNull(rsrs!lfnr) Then
                lLFNR = rsrs!lfnr
            End If
            
            If Not IsNull(rsrs!ADATE) Then
                sADATE = rsrs!ADATE
            End If
            
            If Not IsNull(rsrs!AZEIT) Then
                sAzeit = rsrs!AZEIT
            End If
            
            If Not IsNull(rsrs!AN_FILIALE) Then
                lAN_FILIALE = rsrs!AN_FILIALE
            End If
            
            If Not IsNull(rsrs!MESSAGETEXT) Then
                sMESSAGETEXT = rsrs!MESSAGETEXT
            End If
            
            If Not IsNull(rsrs!BETREFF) Then
                sBetreff = rsrs!BETREFF
            End If
            
            If Not IsNull(rsrs!ABSENDER) Then
                sABSENDER = rsrs!ABSENDER
            End If
            
            If Not IsNull(rsrs!VERSENDET) Then
                bVERSENDET = rsrs!VERSENDET
            End If
            
            If Not IsNull(rsrs!ANGEKOMMEN) Then
                bANGEKOMMEN = rsrs!ANGEKOMMEN
            End If
            
            If Not IsNull(rsrs!gelesen) Then
                bGelesen = rsrs!gelesen
            End If
            
            If Not IsNull(rsrs!GELESEN_VONWEM) Then
                sGELESEN_VONWEM = rsrs!GELESEN_VONWEM
            End If
            
            If Not IsNull(rsrs!GELESEN_ADATE) Then
                daGELESEN_ADATE = rsrs!GELESEN_ADATE
            End If
            
            If Not IsNull(rsrs!GELESEN_AZEIT) Then
                sGELESEN_AZEIT = rsrs!GELESEN_AZEIT
            End If
        
            If Not IsNull(rsrs!FELD1) Then
                bFELD1 = rsrs!FELD1
            End If
            
            If Not IsNull(rsrs!FELD2) Then
                bFELD2 = rsrs!FELD2
            End If
            
            If Not IsNull(rsrs!FELD3) Then
                bFELD3 = rsrs!FELD3
            End If
            
            If Not IsNull(rsrs!feld4) Then
                bFELD4 = rsrs!feld4
            End If
            
            If Not IsNull(rsrs!FELD5) Then
                bFELD5 = rsrs!FELD5
            End If
            
            If Not IsNull(rsrs!FELD6) Then
                bFELD6 = rsrs!FELD6
            End If
            
            If Not IsNull(rsrs!FELD7) Then
                bFELD7 = rsrs!FELD7
            End If
            
            If Not IsNull(rsrs!FELD8) Then
                bFELD8 = rsrs!FELD8
            End If
            
            If Not IsNull(rsrs!FELD9) Then
                bFELD9 = rsrs!FELD9
            End If
            
            If Not IsNull(rsrs!FELD10) Then
                bFELD10 = rsrs!FELD10
            End If
            
            If Not IsNull(rsrs!INFO1) Then
                sINFO1 = rsrs!INFO1
            End If
            
            If Not IsNull(rsrs!INFO2) Then
                sINFO2 = rsrs!INFO2
            End If
            
            If Not IsNull(rsrs!INFO3) Then
                sINFO3 = rsrs!INFO3
            End If
            
            
            
            cSQL = "Insert into Nachrichten"
            cSQL = cSQL & " ( "
            cSQL = cSQL & " lfnr  "
            cSQL = cSQL & ", ADATE "
            cSQL = cSQL & ", AZEIT  "
            cSQL = cSQL & ", AN_FILIALE "
            cSQL = cSQL & ", MESSAGETEXT  "
            cSQL = cSQL & ", BETREFF  "
            cSQL = cSQL & ", ABSENDER  "
            cSQL = cSQL & ", VERSENDET "
            cSQL = cSQL & ", ANGEKOMMEN "
            cSQL = cSQL & ", GELESEN "
            cSQL = cSQL & ", GELESEN_VONWEM  "
            cSQL = cSQL & ", GELESEN_ADATE "
            cSQL = cSQL & ", GELESEN_AZEIT  "
            cSQL = cSQL & ", FELD1 "
            cSQL = cSQL & ", FELD2 "
            cSQL = cSQL & ", FELD3 "
            cSQL = cSQL & ", FELD4 "
            cSQL = cSQL & ", FELD5 "
            cSQL = cSQL & ", FELD6 "
            cSQL = cSQL & ", FELD7 "
            cSQL = cSQL & ", FELD8 "
            cSQL = cSQL & ", FELD9 "
            cSQL = cSQL & ", FELD10 "
            cSQL = cSQL & ", INFO1 "
            cSQL = cSQL & ", INFO2 "
            cSQL = cSQL & ", INFO3 "
            cSQL = cSQL & " ) values ( "
            cSQL = cSQL & " " & lLFNR & "  "
            cSQL = cSQL & ", '" & sADATE & "' "
            cSQL = cSQL & ", '" & sAzeit & "' "
            cSQL = cSQL & ", " & lAN_FILIALE & " "
            cSQL = cSQL & ", '" & sMESSAGETEXT & "' "
            cSQL = cSQL & ", '" & sBetreff & "' "
            cSQL = cSQL & ", '" & sABSENDER & "' "
            
            If bVERSENDET = True Then
                cSQL = cSQL & ", True "
            Else
                cSQL = cSQL & ", False "
            End If
            
            If bANGEKOMMEN = True Then
                cSQL = cSQL & ", True "
            Else
                cSQL = cSQL & ", False "
            End If
            
            If bGelesen = True Then
                cSQL = cSQL & ", True "
            Else
                cSQL = cSQL & ", False "
            End If
            
            cSQL = cSQL & ", '" & sGELESEN_VONWEM & "'   "
            cSQL = cSQL & ", '" & DateValue(Now) & "' "
            cSQL = cSQL & ", '" & sGELESEN_AZEIT & "'   "
            
            If bFELD1 = True Then
                cSQL = cSQL & ", True "
            Else
                cSQL = cSQL & ", False "
            End If
            
            If bFELD2 = True Then
                cSQL = cSQL & ", True "
            Else
                cSQL = cSQL & ", False "
            End If
            
            If bFELD3 = True Then
                cSQL = cSQL & ", True "
            Else
                cSQL = cSQL & ", False "
            End If
            
            If bFELD4 = True Then
                cSQL = cSQL & ", True "
            Else
                cSQL = cSQL & ", False "
            End If
            
            If bFELD5 = True Then
                cSQL = cSQL & ", True "
            Else
                cSQL = cSQL & ", False "
            End If
            
            
            If bFELD6 = True Then
                cSQL = cSQL & ", True "
            Else
                cSQL = cSQL & ", False "
            End If
            
            If bFELD7 = True Then
                cSQL = cSQL & ", True "
            Else
                cSQL = cSQL & ", False "
            End If
            
            If bFELD8 = True Then
                cSQL = cSQL & ", True "
            Else
                cSQL = cSQL & ", False "
            End If
            
            If bFELD9 = True Then
                cSQL = cSQL & ", True "
            Else
                cSQL = cSQL & ", False "
            End If
            
            If bFELD10 = True Then
                cSQL = cSQL & ", True "
            Else
                cSQL = cSQL & ", False "
            End If
            
            
            
            
            
            
            cSQL = cSQL & ", '" & sINFO1 & "' "
            cSQL = cSQL & ", '" & sINFO2 & "' "
            cSQL = cSQL & ", '" & sINFO3 & "' "
            cSQL = cSQL & " )"
'            MsgBox cSQL
            gdBase.Execute cSQL, dbFailOnError
            
                
            rsrs.MoveNext
            Loop
            
        End If
        rsrs.Close: Set rsrs = Nothing
        
        dbEAN.Close
        
    End If
        
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ImportiereNeueNachrichten"
    Fehler.gsFehlertext = "Es trat ein Fehler auf. "
    
    Fehlermeldung1
End Sub

Public Function DabaFileSize() As String
    On Error GoTo LOKAL_ERROR
    
    Dim dFilesize As Double
    Dim cPfad As String
    
    DabaFileSize = ""
    
    cPfad = gcDBPfad      'Dabapfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    dFilesize = FileLen(cPfad & "Kissdata.mdb")     'in BYTE
    dFilesize = dFilesize / 1024                    'in KBYTE
    dFilesize = dFilesize / 1024                    'in MBYTE
    
    DabaFileSize = Format$(dFilesize, "####0.00")
    
    Exit Function
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "DabaFileSize"
    Fehler.gsFehlertext = "Es trat ein Fehler auf. "
    
    Fehlermeldung1
End Function
Public Function fnArtMBORDERSuchenMB(sArtnr As String) As String
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rs As Recordset
    
    fnArtMBORDERSuchenMB = ""
    
    If sArtnr = "" Then
        Exit Function
    End If
    
    fnArtMBORDERSuchenMB = ""
    sSQL = "Select MB, Lastdate from MBORDER where Artnr = " & sArtnr
    Set rs = gdBase.OpenRecordset(sSQL)
    If Not rs.EOF Then
        If Not IsNull(rs!MB) Then
            fnArtMBORDERSuchenMB = rs!MB & " am: " & rs!LASTDATE & " festgesetzt"
        End If
    End If
    rs.Close: Set rs = Nothing
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "fnArtMBORDERSuchenMB"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function fnArtBezSuchen(sArtnr As String) As String
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rs As Recordset
    
    fnArtBezSuchen = ""
    
    If sArtnr = "" Then
        Exit Function
    End If
    
    fnArtBezSuchen = ""
    sSQL = "Select BEZEICH from ARTIKEL where Artnr = " & sArtnr
    Set rs = gdBase.OpenRecordset(sSQL)
    If Not rs.EOF Then
        If Not IsNull(rs!BEZEICH) Then
            fnArtBezSuchen = rs!BEZEICH
        End If
    End If
    rs.Close: Set rs = Nothing
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "fnArtBezSuchen"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function fnArtEanSuchen(sArtnr As String) As String
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rs As Recordset
    
    fnArtEanSuchen = ""
    
    If sArtnr = "" Then
        Exit Function
    End If
    
    fnArtEanSuchen = ""
    sSQL = "Select EAN from ARTIKEL where Artnr = " & sArtnr
    Set rs = gdBase.OpenRecordset(sSQL)
    If Not rs.EOF Then
        If Not IsNull(rs!EAN) Then
            fnArtEanSuchen = rs!EAN
        End If
    End If
    rs.Close: Set rs = Nothing
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "fnArtEanSuchen"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermBonusfähigkeitArtikel(sArtnr As String) As String
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rs As Recordset
    
    ermBonusfähigkeitArtikel = "J"
    
    If sArtnr = "" Then
        Exit Function
    End If
    
    ermBonusfähigkeitArtikel = "J"
    sSQL = "Select Bonus_ok from ARTIKEL where Artnr = " & sArtnr
    Set rs = gdBase.OpenRecordset(sSQL)
    If Not rs.EOF Then
        If Not IsNull(rs!BONUS_OK) Then
            ermBonusfähigkeitArtikel = rs!BONUS_OK
        End If
    End If
    rs.Close: Set rs = Nothing
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ermBonusfähigkeitArtikel"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function Artnr_Over_EAN(sEAN As String) As String
On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rs As Recordset
    
    Artnr_Over_EAN = ""
    
    If sEAN = "" Then
        Exit Function
    End If
    
    sSQL = "Select ARTNR from ARTIKEL "
    sSQL = sSQL & " where (EAN = '" & sEAN & "' "
    sSQL = sSQL & " or EAN2 = '" & sEAN & "' "
    sSQL = sSQL & " or EAN3 = '" & sEAN & "' )"
    
    
    Set rs = gdBase.OpenRecordset(sSQL)
    If Not rs.EOF Then
        If Not IsNull(rs!artnr) Then
            Artnr_Over_EAN = rs!artnr
        End If
    End If
    rs.Close: Set rs = Nothing
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "Artnr_Over_EAN"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermgesbestand() As Long
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    
    ermgesbestand = 0
    
    sSQL = "select sum(Bestand) as Maxi from artikel where bestand > 0 "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermgesbestand = rsrs!maxi
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ermgesbestand"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Function
Public Function ermgesSEKwert() As Double
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    
    ermgesSEKwert = 0
    
    sSQL = "select sum(Bestand * EKPR) as Maxi from artikel where bestand > 0 "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermgesSEKwert = rsrs!maxi
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ermgesSEKwert"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Function
Public Sub ExcelExport(sTab As String, db As Database)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    Dim cdatei              As String
    Dim cPfad1              As String
    Dim cPfad               As String
    Dim iFileNr             As Integer
    Dim sPfad               As String
    Dim iRet                As Integer
    Dim sAusgabedatname     As String
    
    sAusgabedatname = ""
'    sAusgabedatname = "Inventur" & ".xls"
    
    cdatei = cPfad1 & "BOX\" & sAusgabedatname
    cPfad = cPfad1 & "BOX"

    With frmWKL00.cdlopen
        
        .CancelError = True
        On Error GoTo err
        .DialogTitle = "Speichern der Tabelle"
        .Filter = "Excel - Dateien (*.xls)|*.xls"
'        .FileName = ""
'        .FileName = cPfad & "\" & sAusgabedatname
        .ShowSave
    End With

    sPfad = frmWKL00.cdlopen.FileName
    
    If FileExists(sPfad) Then
        iRet = MsgBox("Eine gleichnamige Datei ist schon vorhanden, möchten Sie diese überschreiben?", vbQuestion + vbYesNo, "Winkiss Frage:")
        If iRet = vbNo Then
            Exit Sub
        Else
            Kill sPfad
        End If
    Else
    
    End If
    
    

    sSQL = "Select * into " & sTab & " IN '" & sPfad & "' 'Excel 8.0;' from " & sTab
    gdBase.Execute sSQL, dbFailOnError

    MsgBox "Diese Datei ist unter (" & sPfad & ") abgespeichert", vbInformation, "Winkiss Information:"

err:
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ExcelExport"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Function ZeigeSchwerpunktLinr(cKundnr As String, Listx As ListBox) As Long
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    Dim lLinr As Long
    Dim dSumme As Single
    Dim siAnteil As Single
    Dim siSumTeil As Single
    Dim sSatz As String
    
    ZeigeSchwerpunktLinr = 0
    
    Listx.Clear
    
    loeschNEW "SCHWPUNKT" & srechnertab, gdBase
    sSQL = "select distinct(Linr) as LIN,sum(preis) as sumPreis into SCHWPUNKT" & srechnertab & " from KUNDAZE where KUNDNR = " & cKundnr
    sSQL = sSQL & " and preis > 0 "
    sSQL = sSQL & " group by Linr "
    gdBase.Execute sSQL, dbFailOnError
    
    dSumme = 0
    sSQL = "select sum(sumPreis) as maxi from SCHWPUNKT" & srechnertab & " "
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!maxi) Then
            dSumme = rsrs!maxi
        End If
    End If
    rsrs.Close
    
    sSQL = "select * from SCHWPUNKT" & srechnertab & " order by sumPreis desc"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
        sSatz = ""
        If Not IsNull(rsrs!LIN) Then
            lLinr = rsrs!LIN
        Else
            lLinr = 0
        End If
        
        sSatz = Left(ermLiefBez(lLinr), 20) & ".."
        sSatz = sSatz & Space(23 - Len(sSatz))
        If Not IsNull(rsrs!sumpreis) Then
            sSatz = sSatz & Space(9 - Len(Format$(rsrs!sumpreis, "####0.00"))) & Format$(rsrs!sumpreis, "####0.00")
            siSumTeil = rsrs!sumpreis
            
        Else
            sSatz = sSatz & Space(9)
        End If
        
        siAnteil = 0
        
        If dSumme <> 0 Then
            siAnteil = siSumTeil * 100 / dSumme
            If siAnteil > 80 Then
                ZeigeSchwerpunktLinr = lLinr
            End If
        End If
        
        sSatz = sSatz & Space(8 - Len(Format$(siAnteil, "####0.00"))) & Format$(siAnteil, "####0.00") & " %"
        
        Listx.AddItem sSatz
        
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close

    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ZeigeSchwerpunktLinr"
    Fehler.gsFehlertext = "Im Programmteil Kunden Verkauf ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Sub ZeigeSchwerpunktMarke(cKundnr As String, Listx As ListBox)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    Dim cMarke As String
    Dim dSumme As Single
    Dim siAnteil As Single
    Dim siSumTeil As Single
    Dim sSatz As String
    
    Listx.Clear
    
    loeschNEW "SCHWPUNKT", gdBase
    sSQL = "select distinct(Marke) as LIN,sum(preis) as sumPreis into SCHWPUNKT from KUNDAZE where KUNDNR = " & cKundnr
    sSQL = sSQL & " and preis > 0 "
    sSQL = sSQL & " group by Marke "
    gdBase.Execute sSQL, dbFailOnError
    
    dSumme = 0
    sSQL = "select sum(sumPreis) as maxi from SCHWPUNKT "
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!maxi) Then
            dSumme = rsrs!maxi
        End If
    End If
    rsrs.Close
    
    sSQL = "select * from SCHWPUNKT order by sumPreis desc"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
        sSatz = ""
        If Not IsNull(rsrs!LIN) Then
            cMarke = rsrs!LIN
        Else
            cMarke = ""
        End If
        
        sSatz = Left(cMarke, 20) & ".."
        sSatz = sSatz & Space(23 - Len(sSatz))
        If Not IsNull(rsrs!sumpreis) Then
            sSatz = sSatz & Space(9 - Len(Format$(rsrs!sumpreis, "####0.00"))) & Format$(rsrs!sumpreis, "####0.00")
            siSumTeil = rsrs!sumpreis
        Else
            sSatz = sSatz & Space(9)
        End If
        
        siAnteil = 0
        
        If dSumme <> 0 Then
            siAnteil = siSumTeil * 100 / dSumme
        End If
        
        sSatz = sSatz & Space(8 - Len(Format$(siAnteil, "####0.00"))) & Format$(siAnteil, "####0.00") & " %"
        
        Listx.AddItem sSatz
        
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ZeigeSchwerpunktMarke"
    Fehler.gsFehlertext = "Im Programmteil Kunden Verkauf ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub ZeigeSchwerpunktAGN(cKundnr As String, Listx As ListBox)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    Dim lagn As Long
    Dim dSumme As Single
    Dim siAnteil As Single
    Dim siSumTeil As Single
    Dim sSatz As String
    
    Listx.Clear
    
    loeschNEW "SCHWPUNKT", gdBase
    sSQL = "select distinct(AGN) as LIN,sum(preis) as sumPreis into SCHWPUNKT from KUNDAZE where KUNDNR = " & cKundnr
    sSQL = sSQL & " and preis > 0 "
    sSQL = sSQL & " group by AGN "
    gdBase.Execute sSQL, dbFailOnError
    
    dSumme = 0
    sSQL = "select sum(sumPreis) as maxi from SCHWPUNKT "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!maxi) Then
            dSumme = rsrs!maxi
        End If
    End If
    rsrs.Close
    
    sSQL = "select * from SCHWPUNKT order by sumPreis desc"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
        sSatz = ""
        If Not IsNull(rsrs!LIN) Then
            lagn = rsrs!LIN
        Else
            lagn = 0
        End If
        
        sSatz = Left(ermAGNbez(CStr(lagn), gdBase), 20) & ".."
        sSatz = sSatz & Space(23 - Len(sSatz))
        If Not IsNull(rsrs!sumpreis) Then
            sSatz = sSatz & Space(9 - Len(Format$(rsrs!sumpreis, "####0.00"))) & Format$(rsrs!sumpreis, "####0.00")
            siSumTeil = rsrs!sumpreis
        Else
            sSatz = sSatz & Space(9)
        End If
        
        siAnteil = 0
        
        If dSumme <> 0 Then
            siAnteil = siSumTeil * 100 / dSumme
        End If
        
        sSatz = sSatz & Space(8 - Len(Format$(siAnteil, "####0.00"))) & Format$(siAnteil, "####0.00") & " %"
        
        Listx.AddItem sSatz
        
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ZeigeSchwerpunktAGN"
    Fehler.gsFehlertext = "Im Programmteil Kunden Verkauf ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub ZeigeSchwerpunktPGN(cKundnr As String, Listx As ListBox)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    Dim lPGN As Long
    Dim dSumme As Single
    Dim siAnteil As Single
    Dim siSumTeil As Single
    Dim sSatz As String
    
    Listx.Clear
    
    loeschNEW "SCHWPUNKT", gdBase
    sSQL = "select distinct(PGN) as LIN,sum(preis) as sumPreis into SCHWPUNKT from KUNDAZE where KUNDNR = " & cKundnr
    sSQL = sSQL & " and preis > 0 "
    sSQL = sSQL & " group by PGN "
    gdBase.Execute sSQL, dbFailOnError
    
    dSumme = 0
    sSQL = "select sum(sumPreis) as maxi from SCHWPUNKT "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!maxi) Then
            dSumme = rsrs!maxi
        End If
    End If
    rsrs.Close
    
    sSQL = "select * from SCHWPUNKT order by sumPreis desc"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
        sSatz = ""
        If Not IsNull(rsrs!LIN) Then
            lPGN = rsrs!LIN
        Else
            lPGN = 0
        End If
        
        sSatz = Left(Ermittlepgntext(CStr(lPGN)), 20) & ".."
        sSatz = sSatz & Space(23 - Len(sSatz))
        If Not IsNull(rsrs!sumpreis) Then
            sSatz = sSatz & Space(9 - Len(Format$(rsrs!sumpreis, "####0.00"))) & Format$(rsrs!sumpreis, "####0.00")
            siSumTeil = rsrs!sumpreis
        Else
            sSatz = sSatz & Space(9)
        End If
        
        siAnteil = 0
        
        If dSumme <> 0 Then
            siAnteil = siSumTeil * 100 / dSumme
        End If
        
        sSatz = sSatz & Space(8 - Len(Format$(siAnteil, "####0.00"))) & Format$(siAnteil, "####0.00") & " %"
        
        Listx.AddItem sSatz
        
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ZeigeSchwerpunktPGN"
    Fehler.gsFehlertext = "Im Programmteil Kunden Verkauf ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub GFKstat(cKundnr As String, bAutomatik As Boolean, sEmail As String)
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    
    loeschNEW "GFKSTAT", gdBase
    
    If Trim(cKundnr) = "" Then
        
        Exit Sub
    ElseIf Trim(cKundnr) = "0" Then
    
        Exit Sub
    End If
    
    If bAutomatik Then
        If Trim(sEmail) = "" Then
            
            Exit Sub
        End If
    End If
    
    CreateTableT2 "GFKSTAT", gdBase
    
    If bAutomatik = True Then
        sSQL = "Insert into GFKSTAT (KUNDNR,AUTOMATIK,EMAIL) values (" & cKundnr & ",TRUE,'" & sEmail & "')"
    Else
        sSQL = "Insert into GFKSTAT (KUNDNR,AUTOMATIK,EMAIL) values (" & cKundnr & ",False,'" & sEmail & "')"
    End If
    gdBase.Execute sSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "GFKstat"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Function leseGFKstat(sSpalte As String) As String
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rsrs As Recordset
    
    leseGFKstat = ""
    
    If NewTableSuchenDBKombi("GFKSTAT", gdBase) Then
    
        Select Case sSpalte
        
            Case "email"
   
                sSQL = "select email from GFKSTAT "
                Set rsrs = gdBase.OpenRecordset(sSQL)
                If Not rsrs.EOF Then
                    If Not IsNull(rsrs!Email) Then
                        leseGFKstat = rsrs!Email
                    End If
                End If
                rsrs.Close
                
            Case "lastdate"
   
                sSQL = "select lastdate from GFKSTAT "
                Set rsrs = gdBase.OpenRecordset(sSQL)
                If Not rsrs.EOF Then
                    If Not IsNull(rsrs!LASTDATE) Then
                        leseGFKstat = rsrs!LASTDATE
                    End If
                End If
                rsrs.Close
                
            Case "kundnr"
   
                sSQL = "select kundnr from GFKSTAT "
                Set rsrs = gdBase.OpenRecordset(sSQL)
                If Not rsrs.EOF Then
                    If Not IsNull(rsrs!Kundnr) Then
                        leseGFKstat = rsrs!Kundnr
                    End If
                End If
                rsrs.Close
                
            Case "automatik"
                leseGFKstat = "0"
   
                sSQL = "select automatik from GFKSTAT "
                Set rsrs = gdBase.OpenRecordset(sSQL)
                If Not rsrs.EOF Then
                    If Not IsNull(rsrs!automatik) Then
                        leseGFKstat = rsrs!automatik
                    End If
                End If
                rsrs.Close
                
        End Select
        
    End If
    
Exit Function
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "leseGFKstat"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Function
Public Function IPstat(cMarktnr As String, blive As Boolean) As Boolean
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    IPstat = False
    
    
    If Trim(cMarktnr) = "" Then
       
        Exit Function
    ElseIf Trim(cMarktnr) = "0" Then
       
        Exit Function
    End If
    
    gbIPSTAT = blive
    gsIPMarktNr = cMarktnr
    
   
    loeschNEW "IPSTAT", gdBase
    CreateTableT2 "IPSTAT", gdBase
    
    If blive = True Then
        sSQL = "Insert into IPSTAT (MARKTNR,LIVE) values (" & cMarktnr & ",TRUE)"
        IPstat = True

    Else
        sSQL = "Insert into IPSTAT (MARKTNR,LIVE) values (" & cMarktnr & ",False)"
    End If
    gdBase.Execute sSQL, dbFailOnError
    
Exit Function
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "GFKstat"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function VEDESstat(cMarktnr As String, blive As Boolean) As Boolean
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    VEDESstat = False
    
    
    If Trim(cMarktnr) = "" Then
       
        Exit Function
    ElseIf Trim(cMarktnr) = "0" Then
       
        Exit Function
    End If
    
    gbVEDESSTAT = blive
    gsVEDESMarktNr = cMarktnr
    
   
    loeschNEW "VEDESSTAT", gdBase
    CreateTableT2 "VEDESSTAT", gdBase
    
    If blive = True Then
        sSQL = "Insert into VEDESSTAT (MARKTNR,LIVE) values (" & cMarktnr & ",TRUE)"
        VEDESstat = True

    Else
        sSQL = "Insert into VEDESSTAT (MARKTNR,LIVE) values (" & cMarktnr & ",False)"
    End If
    gdBase.Execute sSQL, dbFailOnError
    
Exit Function
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "VEDESstat"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function leseCouponStat(sSpalte As String) As String
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rsrs As Recordset
    
    leseCouponStat = ""
    
    If NewTableSuchenDBKombi("COUPONSTAT", gdBase) Then
    
        Select Case sSpalte
        
            Case "lastdate"
                sSQL = "select lastdate from COUPONSTAT "
                Set rsrs = gdBase.OpenRecordset(sSQL)
                If Not rsrs.EOF Then
                    If Not IsNull(rsrs!LASTDATE) Then
                        leseCouponStat = rsrs!LASTDATE
                    End If
                End If
                rsrs.Close
                   
        End Select
        
    End If
    
Exit Function
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "leseCouponStat"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Function
Public Function leseIPstat(sSpalte As String) As String
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rsrs As Recordset
    
    leseIPstat = ""
    
    If NewTableSuchenDBKombi("IPSTAT", gdBase) Then
    
        Select Case sSpalte
        
            
                
            Case "lastdate"
   
                sSQL = "select lastdate from IPSTAT "
                Set rsrs = gdBase.OpenRecordset(sSQL)
                If Not rsrs.EOF Then
                    If Not IsNull(rsrs!LASTDATE) Then
                        leseIPstat = rsrs!LASTDATE
                    End If
                End If
                rsrs.Close
                
            Case "marktnr"
   
                sSQL = "select marktnr from IPSTAT "
                Set rsrs = gdBase.OpenRecordset(sSQL)
                If Not rsrs.EOF Then
                    If Not IsNull(rsrs!marktnr) Then
                        leseIPstat = rsrs!marktnr
                    End If
                End If
                rsrs.Close
                
            Case "live"
                leseIPstat = "0"
   
                sSQL = "select live from IPSTAT "
                Set rsrs = gdBase.OpenRecordset(sSQL)
                If Not rsrs.EOF Then
                    If Not IsNull(rsrs!live) Then
                        leseIPstat = rsrs!live
                    End If
                End If
                rsrs.Close
                
        End Select
        
    End If
    
Exit Function
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "leseIPstat"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Function
Public Function leseVEDESstat(sSpalte As String) As String
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rsrs As Recordset
    
    leseVEDESstat = ""
    
    If NewTableSuchenDBKombi("VEDESSTAT", gdBase) Then
    
        Select Case sSpalte
        
            
                
            Case "lastdate"
   
                sSQL = "select lastdate from VEDESSTAT "
                Set rsrs = gdBase.OpenRecordset(sSQL)
                If Not rsrs.EOF Then
                    If Not IsNull(rsrs!LASTDATE) Then
                        leseVEDESstat = rsrs!LASTDATE
                    End If
                End If
                rsrs.Close
                
            Case "marktnr"
   
                sSQL = "select marktnr from VEDESSTAT "
                Set rsrs = gdBase.OpenRecordset(sSQL)
                If Not rsrs.EOF Then
                    If Not IsNull(rsrs!marktnr) Then
                        leseVEDESstat = rsrs!marktnr
                    End If
                End If
                rsrs.Close
                
            Case "live"
                leseVEDESstat = "0"
   
                sSQL = "select live from VEDESSTAT "
                Set rsrs = gdBase.OpenRecordset(sSQL)
                If Not rsrs.EOF Then
                    If Not IsNull(rsrs!live) Then
                        leseVEDESstat = rsrs!live
                    End If
                End If
                rsrs.Close
                
        End Select
        
    End If
    
Exit Function
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "leseVEDESstat"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Function
Public Sub Couponeinloesung(cBudniKundNr As String, lAuswerttag As Long)
On Error GoTo LOKAL_ERROR

    Dim sSQL As String

    loeschNEW "COUPONPRINT", gdBase
    CreateTableT2 "COUPONPRINT", gdBase
    
'    sSQL = "Insert into COUPONPRINT Select  "
'    sSQL = sSQL & " ARTNR  "
'    sSQL = sSQL & ", BEZEICH "
'    sSQL = sSQL & ", 'DRONOVA' as KETTE "
'    sSQL = sSQL & ", 'DRONOVA' as GRUPPE "
'    sSQL = sSQL & ", '" & cBudniKundNr & "' as KUNDNR "
'    sSQL = sSQL & ", ADATE  "
'    sSQL = sSQL & ", EAN "
'    sSQL = sSQL & ", MENGE  "
'    sSQL = sSQL & ", PREIS from Kassjour"
'    sSQL = sSQL & " where "
'    sSQL = sSQL & " adate >= " & cVon & " and adate <= " & cBis & " "
'    sSQL = sSQL & " and artnr in(select artnr from artlief where Linr = " & lCouponLinr & ")"
'    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into COUPONPRINT Select  "
    sSQL = sSQL & " ARTNR  "
    sSQL = sSQL & ", BEZEICH "
    sSQL = sSQL & ", 'DRONOVA' as KETTE "
    sSQL = sSQL & ", 'DRONOVA' as GRUPPE "
    sSQL = sSQL & ", '" & cBudniKundNr & "' as KUNDNR "
    sSQL = sSQL & ", ADATE  "
    sSQL = sSQL & ", EAN "
    sSQL = sSQL & ", MENGE  "
    sSQL = sSQL & ", PREIS from Kassjour"
    sSQL = sSQL & " where "
    sSQL = sSQL & " adate = " & lAuswerttag & " "
    sSQL = sSQL & " and ean like '98232*'"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update COUPONPRINT Set "
    sSQL = sSQL & " PREIS  =PREIS * (-1)  "
    gdBase.Execute sSQL, dbFailOnError
    
    If Datendrin("COUPONPRINT", gdBase) Then
        schreibe_CouponCSV cBudniKundNr, lAuswerttag
    End If
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "Couponeinloesung"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub IP_AUSW_erstellen(lAuswerttag As Long)
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rsrs As Recordset
    
    Dim cSatz           As String
    Dim cArtNr          As String
    Dim cUhrZeit        As String
    Dim cBezeich        As String
    Dim cEAN            As String
    Dim cKasnum         As String
    Dim cBELEGNR        As String
    Dim cMenge          As String
    Dim cAgn            As String
    Dim cVKEinzel       As String
    Dim cPreis          As String
    Dim cPfad           As String
    Dim cDatum          As String
    Dim iFileNr         As Integer
    Dim lPos            As Long
    Dim iRet            As Integer
    
    Screen.MousePointer = 11

    loeschNEW "IP_TAG", gdBase
    CreateTableT2 "IP_TAG", gdBase
    
    sSQL = "Insert into IP_TAG Select "
    sSQL = sSQL & " artnr "
    sSQL = sSQL & ",bezeich "
    sSQL = sSQL & ",ean "
    sSQL = sSQL & ",agn "
    sSQL = sSQL & ",menge "
    sSQL = sSQL & ",preis "
    sSQL = sSQL & ",adate "
    sSQL = sSQL & ",azeit "
    sSQL = sSQL & ",BELEGNR "
    sSQL = sSQL & ",KASNUM "
    sSQL = sSQL & ",FILIALE "
    sSQL = sSQL & " from kassjour "
    sSQL = sSQL & " WHERE adate= " & lAuswerttag
    gdBase.Execute sSQL, dbFailOnError
    
    cDatum = Format(lAuswerttag, "YYYY.MM.DD")
    cDatum = SwapStr(cDatum, ".", "")
    
    Dim cMarktnr As String
    
    cMarktnr = gsIPMarktNr
    
    While Len(Trim(cMarktnr)) < 8
        cMarktnr = "0" & cMarktnr
    Wend
    
    cPfad = gcDBPfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "IP\"
    cPfad = UCase$(cPfad)

    Kill cPfad & "IP_" & CStr(lAuswerttag) & "_" & gsIPMarktNr & ".TXT"
    
    iFileNr = FreeFile
    Open cPfad & "IP_" & CStr(lAuswerttag) & "_" & gsIPMarktNr & ".TXT" For Binary As #iFileNr
    
    cSatz = "Marktnr(1-8) Datum(YYYYMMDD 9-16) Uhrzeit(HHMMSS 17-22) EAN(23-35) Artikelbezeichnung(36-70) ArtikelNr(71-76) VKMenge(77-81)"
    cSatz = cSatz & "VKPreis(82-91) AGN(92-97) Belegnr(98-101) KassenNr(102-103)"
    cSatz = cSatz & Chr$(13) & Chr$(10)
    lPos = LOF(iFileNr)
    lPos = lPos + 1
    Put #iFileNr, lPos, cSatz
    
    sSQL = "Select * from IP_TAG order by kasnum, belegnr"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            If Not IsNull(rsrs!EAN) Then
                cEAN = rsrs!EAN
            Else
                cEAN = ""
            End If
            
            While Len(Trim(cEAN)) < 13
                cEAN = "0" & cEAN
            Wend
            
            If Not IsNull(rsrs!BEZEICH) Then
                cBezeich = Left(rsrs!BEZEICH, 35)
            Else
                cBezeich = ""
            End If
            
            cBezeich = SwapStr(cBezeich, Chr(10), " ")  'umbruch
            cBezeich = SwapStr(cBezeich, Chr(13), " ")  'umbruch
            
            
            While Len(cBezeich) < 35
                cBezeich = cBezeich & " "
            Wend
            
            If Not IsNull(rsrs!artnr) Then
                cArtNr = rsrs!artnr
            Else
                cArtNr = ""
            End If
            
            While Len(Trim(cArtNr)) < 6
                cArtNr = "0" & cArtNr
            Wend
            
            If Not IsNull(rsrs!MENGE) Then
                cMenge = rsrs!MENGE
            Else
                cMenge = ""
            End If
            
            While Len(cMenge) < 5
                cMenge = " " & cMenge
            Wend
            
            If Not IsNull(rsrs!PREIS) Then
                cPreis = Format(rsrs!PREIS, "#####0.00")
            Else
                cPreis = ""
            End If
            
            cPreis = SwapStr(cPreis, ",", ".")
            
            While Len(cPreis) < 10
                cPreis = " " & cPreis
            Wend
            
            If Not IsNull(rsrs!AGN) Then
                cAgn = rsrs!AGN
            Else
                cAgn = ""
            End If
            
            While Len(Trim(cAgn)) < 6
                cAgn = "0" & cAgn
            Wend
            
            If Not IsNull(rsrs!AZEIT) Then
                cUhrZeit = rsrs!AZEIT
            Else
                cUhrZeit = ""
            End If
            
            cUhrZeit = SwapStr(cUhrZeit, ":", "")
            
            While Len(Trim(cUhrZeit)) < 6
                cUhrZeit = "0" & cUhrZeit
            Wend
            
            If Not IsNull(rsrs!BELEGNR) Then
                cBELEGNR = rsrs!BELEGNR
            Else
                cBELEGNR = ""
            End If
            
            While Len(Trim(cBELEGNR)) < 4
                cBELEGNR = "0" & cBELEGNR
            Wend
            
            If Not IsNull(rsrs!KASNUM) Then
                cKasnum = rsrs!KASNUM
            Else
                cKasnum = ""
            End If
            
            While Len(Trim(cKasnum)) < 2
                cKasnum = "0" & cKasnum
            Wend
            
            cSatz = cMarktnr & cDatum & cUhrZeit & cEAN & cBezeich & cArtNr & cMenge & cPreis & cAgn & cBELEGNR & cKasnum
            cSatz = cSatz & Chr$(13) & Chr$(10)
            lPos = LOF(iFileNr)
            lPos = lPos + 1
            Put #iFileNr, lPos, cSatz
    
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    Close iFileNr
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "IP_AUSW_erstellen"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Public Sub IP_AUSW_uebertragen()
On Error GoTo LOKAL_ERROR

    'IP Auswertung auf den FTPSERVER
    
    Dim bmerke As Boolean
    bmerke = gbFTPautomatic
    gbFTPautomatic = True
        
    giKissFtpMode = 30
    frmWKL38.Show 1
    
    gbFTPautomatic = bmerke
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "IP_AUSW_uebertragen"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Public Sub COUPON_AUSW_uebertragen()
On Error GoTo LOKAL_ERROR

    'IP Auswertung auf den FTPSERVER
    
    Dim bmerke As Boolean
    bmerke = gbFTPautomatic
    gbFTPautomatic = True
        
    giKissFtpMode = 43
    frmWKL38.Show 1
    
    gbFTPautomatic = bmerke
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "COUPON_AUSW_uebertragen"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub VEDES_AUSW_erstellen(lAuswerttag As Long)
On Error GoTo LOKAL_ERROR

    Dim sSQL            As String
    Dim rsrs            As Recordset
    Dim cSatz           As String
    Dim cUhrZeit        As String
    Dim lMenge          As Long
    Dim cEAN            As String
    Dim cKasnum         As String
    Dim cBELEGNR        As String
    Dim cAgn            As String
    Dim cVKEinzel       As String
    Dim cPreis          As String
    Dim cPfad           As String
    Dim cDatum          As String
    Dim iFileNr         As Integer
    Dim lPos            As Long
    Dim iRet            As Integer
    Dim lVorgang        As Long
    
    Dim cEkPreis        As String
    Dim cMwStWert       As String
    Dim cRabwert        As String
    Dim cMwst           As String
    
    Dim cSortiment      As String
    Dim cLieferant      As String
    Dim cLiefArt        As String
    
    cSortiment = ""
    cLieferant = ""
    cLiefArt = ""
    
    Dim cMarktnr        As String
    cMarktnr = gsVEDESMarktNr
    
    Dim sTime           As String
    Dim sDate           As String
    
    Screen.MousePointer = 11
    
    sTime = Format$(TimeValue(Now), "HHMM")
    sDate = Format$(lAuswerttag, "DDMMYYYY")
    
    Dim cFil As String
    cFil = "0"
    
    cPfad = gcDBPfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "VEDES\"
    cPfad = UCase$(cPfad)

    Kill cPfad & "POSDMEX_" & sDate & sTime & ".txt"
    
    iFileNr = FreeFile
    Open cPfad & "POSDMEX_" & sDate & sTime & ".txt" For Binary As #iFileNr
    
    sSQL = "Select * from kassjour "
    sSQL = sSQL & " WHERE adate= " & lAuswerttag
    sSQL = sSQL & " order by kasnum, belegnr"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            If Not IsNull(rsrs!EAN) Then
                cEAN = rsrs!EAN
            Else
                cEAN = ""
            End If
            
            lMenge = 0
            If Not IsNull(rsrs!MENGE) Then
                lMenge = rsrs!MENGE
            End If
            
            lVorgang = 1
            If lMenge < 0 Then
                lVorgang = 2
            End If
            
            If Not IsNull(rsrs!PREIS) Then
                cPreis = Format(rsrs!PREIS, "#####0.00")
            Else
                cPreis = "0"
            End If
            
            If Not IsNull(rsrs!ekpr) Then
                cEkPreis = Format(rsrs!ekpr, "#####0.00")
            Else
                cEkPreis = "0"
            End If
            
        
            If Not IsNull(rsrs!AGN) Then
                cAgn = rsrs!AGN
            Else
                cAgn = ""
            End If
            
            If Len(Trim(cAgn)) > 4 Then
                cAgn = Left(cAgn, 4)
            End If
            
            If Not IsNull(rsrs!AZEIT) Then
                cUhrZeit = rsrs!AZEIT
            Else
                cUhrZeit = ""
            End If
            
            cUhrZeit = Format$(cUhrZeit, "HHMM")
            cDatum = Format$(lAuswerttag, "YYYYMMDD")
            
            If Not IsNull(rsrs!BELEGNR) Then
                cBELEGNR = rsrs!BELEGNR
            Else
                cBELEGNR = ""
            End If
            
            If Not IsNull(rsrs!KASNUM) Then
                cKasnum = rsrs!KASNUM
            Else
                cKasnum = ""
            End If
            
            If Not IsNull(rsrs!MWST) Then
                cMwst = rsrs!MWST
            Else
                cMwst = ""
            End If
            
            
            
            Select Case cMwst
                Case Is = "V"
                    cMwStWert = CStr(CDbl(cPreis) * gdMWStV / (100 + gdMWStV))
                Case Is = "E"
                    cMwStWert = CStr(CDbl(cPreis) * gdMWStE / (100 + gdMWStE))
                Case Is = "O"
                    cMwStWert = cPreis
                
            End Select
            
            cMwStWert = Format(cMwStWert, "#####0.00")
            
            If Not IsNull(rsrs!vkpr) Then
                cVKEinzel = rsrs!vkpr
            Else
                cVKEinzel = ""
            End If
            
             
            
            
            cRabwert = CDbl(cVKEinzel) - CDbl(cPreis) / lMenge
            cRabwert = Format(cRabwert, "#####0.00")
            
            cSatz = cMarktnr & vbTab & cFil & vbTab & cKasnum & vbTab & cDatum & vbTab & cUhrZeit & vbTab & cBELEGNR & vbTab & cEAN & vbTab & cAgn & vbTab
            cSatz = cSatz & cSortiment & vbTab & cLieferant & vbTab & cLiefArt & vbTab
            cSatz = cSatz & lVorgang & vbTab & lMenge & vbTab & cEkPreis & vbTab & cPreis & vbTab & cMwStWert & vbTab & cRabwert
            cSatz = cSatz & Chr$(13) & Chr$(10)
            lPos = LOF(iFileNr)
            lPos = lPos + 1
            Put #iFileNr, lPos, cSatz
    
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    Close iFileNr
    
'    Kill cPfad & "VEDES_" & CStr(lAuswerttag) & "_" & gsVEDESMarktNr & ".TXT"
    
'    MsgBox "Die Datei befindet sich hier: " & cPfad & "POSDMEX_" & sDate & sTime & ".TXT", vbInformation, "Winkiss Hinweis:"
    
    Screen.MousePointer = 0
    

    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "VEDES_AUSW_erstellen"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1
'        Resume Next
    End If
End Sub
Public Sub VEDES_AUSW_uebertragen()
On Error GoTo LOKAL_ERROR

    'IP Auswertung auf den FTPSERVER
    
    Dim bmerke As Boolean
    bmerke = gbFTPautomatic
    gbFTPautomatic = True
        
    giKissFtpMode = 41
    frmWKL38.Show 1
    
    gbFTPautomatic = bmerke
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "VEDES_AUSW_uebertragen"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Public Sub GFKerstellen(cKW As String, iJahr As Integer, cKundnr As String, bAutomatik As Boolean, sEmail As String)
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rsrs As Recordset
    
    Dim cSatz           As String
    Dim cArtNr          As String
    Dim cBezeich        As String
    Dim cEAN            As String
    Dim cInhaltBez      As String
    Dim cBestand        As String
    Dim cVKMENGE        As String
    Dim cAgn            As String
    Dim cVKEinzel       As String
    Dim cVKPreis        As String
    Dim cPfad           As String
    Dim cDatum          As String
    Dim iFileNr         As Integer
    Dim lPos            As Long
    Dim iRet            As Integer
    
    
    Screen.MousePointer = 11
    
    loeschNEW "GFKWOCHE", gdBase
    CreateTableT2 "GFKWOCHE", gdBase
    
    sSQL = "Insert into GFKWOCHE Select "
    sSQL = sSQL & " artnr "
    sSQL = sSQL & ",bezeich "
    sSQL = sSQL & ",ean "
    sSQL = sSQL & ",agn "
    sSQL = sSQL & ",sum(menge)as vkmenge "
    sSQL = sSQL & ",sum(preis) as vkpreis "
    
    sSQL = sSQL & " from kassjour "
    
    sSQL = sSQL & " WHERE year(adate)= " & iJahr
    sSQL = sSQL & " and ((DatePart('ww',kassjour.adate)= " & cKW & "))"
    
    
    sSQL = sSQL & " and menge > 0 "
    sSQL = sSQL & " group by artnr "
    sSQL = sSQL & ",bezeich "
    sSQL = sSQL & ",ean "
    sSQL = sSQL & ",agn "
    
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update GFKWOCHE inner join artikel on gfkwoche.artnr = artikel.artnr "
    sSQL = sSQL & " set gfkwoche.inhaltbez = artikel.inhaltbez"
    sSQL = sSQL & " , gfkwoche.bestand = artikel.bestand"
    gdBase.Execute sSQL, dbFailOnError
    
    While Len(Trim(cKundnr)) < 4
        cKundnr = "0" & cKundnr
    Wend
    
    cDatum = Format(DateValue(Now), "DD.MM.YY")
    cDatum = SwapStr(cDatum, ".", "")
    
    cPfad = gcDBPfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "GFK\"
    cPfad = UCase$(cPfad)

    Kill cPfad & "GFK_" & cKW & "_" & cKundnr & ".TXT"
    
    iFileNr = FreeFile
    Open cPfad & "GFK_" & cKW & "_" & cKundnr & ".TXT" For Binary As #iFileNr
    
    sSQL = "Select * from GFKWOCHE order by artnr"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            If Not IsNull(rsrs!EAN) Then
                cEAN = rsrs!EAN
            Else
                cEAN = ""
            End If
            
            While Len(Trim(cEAN)) < 13
                cEAN = "0" & cEAN
            Wend
            
            If Not IsNull(rsrs!BEZEICH) Then
                cBezeich = Left(rsrs!BEZEICH, 30)
            Else
                cBezeich = ""
            End If
            
            While Len(cBezeich) < 30
                cBezeich = " " & cBezeich
            Wend
            
            If Not IsNull(rsrs!INHALTBEZ) Then
                cInhaltBez = Left(rsrs!INHALTBEZ, 2)
            Else
                cInhaltBez = ""
            End If
            
            While Len(cInhaltBez) < 2
                cInhaltBez = " " & cInhaltBez
            Wend
            
            If Not IsNull(rsrs!artnr) Then
                cArtNr = rsrs!artnr
            Else
                cArtNr = ""
            End If
            
            While Len(Trim(cArtNr)) < 6
                cArtNr = "0" & cArtNr
            Wend
            
            If Not IsNull(rsrs!BESTAND) Then
                cBestand = rsrs!BESTAND
            Else
                cBestand = ""
            End If
            
            While Len(cBestand) < 6
                cBestand = " " & cBestand
            Wend
            
            If Not IsNull(rsrs!VKMENGE) Then
                cVKMENGE = rsrs!VKMENGE
            Else
                cVKMENGE = ""
            End If
            
            While Len(cVKMENGE) < 5
                cVKMENGE = " " & cVKMENGE
            Wend
            
            If Not IsNull(rsrs!VKPREIS) Then
                cVKPreis = Format(rsrs!VKPREIS, "#####0.00")
            Else
                cVKPreis = ""
            End If
            
            cVKEinzel = Format(CDbl(cVKPreis) / CLng(cVKMENGE), "#####0.00")
            cVKEinzel = SwapStr(cVKEinzel, ",", ".")
            
            cVKPreis = SwapStr(cVKPreis, ",", ".")
            
            While Len(cVKPreis) < 10
                cVKPreis = " " & cVKPreis
            Wend
            
            While Len(cVKEinzel) < 7
                cVKEinzel = " " & cVKEinzel
            Wend

            If Not IsNull(rsrs!AGN) Then
                cAgn = rsrs!AGN
            Else
                cAgn = ""
            End If
            
            While Len(Trim(cAgn)) < 6
                cAgn = "0" & cAgn
            Wend
            
''            If Not IsNull(rsrs!Adate) Then
''                cKaufdat = rsrs!Adate
''            Else
''                cKaufdat = ""
''            End If
''
''            cKaufdat = Format(cKaufdat, "DD.MM.YY")
''            cKaufdat = SwapStr(cKaufdat, ".", "")
''
''            cWeek = DatePart("ww", DateValue(rsrs!Adate))
            
            While Len(Trim(cKW)) < 2
                cKW = "0" & cKW
            Wend
            
            cSatz = cKundnr & cDatum & cEAN & cBezeich & cInhaltBez & cArtNr & cBestand & cVKMENGE & cVKEinzel & cVKPreis & cAgn & "01"
            cSatz = cSatz & Chr$(13) & Chr$(10)
            lPos = LOF(iFileNr)
            lPos = lPos + 1
            Put #iFileNr, lPos, cSatz
    
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    Close iFileNr
    
    sSQL = "Update GFKSTAT Set LASTDATE = '" & DateValue(Now) & "'"
    gdBase.Execute sSQL, dbFailOnError
    
    Screen.MousePointer = 0
    
    

    If bAutomatik Then
        'GfK auf den FTPSERVER
    
        Dim bmerke As Boolean
        bmerke = gbFTPautomatic
        gbFTPautomatic = True
            
        giKissFtpMode = 29
        frmWKL38.Show 1
        
        gbFTPautomatic = bmerke
    Else
        MsgBox "Datei 'GFK_" & cKW & "_" & cKundnr & ".TXT' in " & cPfad & " erstellt!", vbInformation, "Winkiss Hinweis:"
    End If
    
    
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "GFKerstellen"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Public Sub GFKerstellenJahr(iJahr As Integer, cKundnr As String, bAutomatik As Boolean, sEmail As String)
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rsrs As Recordset
    
    Dim cSatz           As String
    Dim cArtNr          As String
    Dim cBezeich        As String
    Dim cEAN            As String
    Dim cInhaltBez      As String
    Dim cBestand        As String
    Dim cVKMENGE        As String
    Dim cAgn            As String
    Dim cVKEinzel       As String
    Dim cVKPreis        As String
    Dim cPfad           As String
    Dim cDatum          As String
    Dim iFileNr         As Integer
    Dim lPos            As Long
    Dim iRet            As Integer
    Dim cKaufdat        As String
    Dim cWeek           As String
    
    Screen.MousePointer = 11
    
    loeschNEW "GFKWOCHE", gdBase
    CreateTableT2 "GFKWOCHE", gdBase
    
'    sSQL = "Insert into GFKWOCHE Select "
'    sSQL = sSQL & " artnr "
'    sSQL = sSQL & ",bezeich "
'    sSQL = sSQL & ",ean "
'    sSQL = sSQL & ",agn "
'    sSQL = sSQL & ",sum(menge)as vkmenge "
'    sSQL = sSQL & ",sum(preis) as vkpreis "
'    sSQL = sSQL & " from kassjour "
'    sSQL = sSQL & " WHERE year(adate)= " & iJahr
'    sSQL = sSQL & " and menge > 0 "
'    sSQL = sSQL & " group by artnr "
'    sSQL = sSQL & ",bezeich "
'    sSQL = sSQL & ",ean "
'    sSQL = sSQL & ",agn "
'    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into GFKWOCHE Select "
    sSQL = sSQL & " artnr "
    sSQL = sSQL & ",bezeich "
    sSQL = sSQL & ",ean "
    sSQL = sSQL & ",agn "
    sSQL = sSQL & ",menge as vkmenge "
    sSQL = sSQL & ",preis as vkpreis "
    sSQL = sSQL & ",adate "
    sSQL = sSQL & " from kassjour "
    sSQL = sSQL & " WHERE year(adate)= " & iJahr
    sSQL = sSQL & " and menge > 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update GFKWOCHE inner join artikel on gfkwoche.artnr = artikel.artnr "
    sSQL = sSQL & " set gfkwoche.inhaltbez = artikel.inhaltbez"
    sSQL = sSQL & " , gfkwoche.bestand = artikel.bestand"
    gdBase.Execute sSQL, dbFailOnError
    
    While Len(Trim(cKundnr)) < 4
        cKundnr = "0" & cKundnr
    Wend
    
    cDatum = Format(DateValue(Now), "DD.MM.YY")
    cDatum = SwapStr(cDatum, ".", "")
    
    cPfad = gcDBPfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "GFK\"
    cPfad = UCase$(cPfad)

    Kill cPfad & "GFK_" & CStr(iJahr) & "_" & cKundnr & ".TXT"
    
    iFileNr = FreeFile
    Open cPfad & "GFK_" & CStr(iJahr) & "_" & cKundnr & ".TXT" For Binary As #iFileNr
    
    sSQL = "Select * from GFKWOCHE order by adate"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            If Not IsNull(rsrs!EAN) Then
                cEAN = rsrs!EAN
            Else
                cEAN = ""
            End If
            
            While Len(Trim(cEAN)) < 13
                cEAN = "0" & cEAN
            Wend
            
            If Not IsNull(rsrs!BEZEICH) Then
                cBezeich = Left(rsrs!BEZEICH, 30)
            Else
                cBezeich = ""
            End If
            
            While Len(cBezeich) < 30
                cBezeich = " " & cBezeich
            Wend
            
            If Not IsNull(rsrs!INHALTBEZ) Then
                cInhaltBez = Left(rsrs!INHALTBEZ, 2)
            Else
                cInhaltBez = ""
            End If
            
            While Len(cInhaltBez) < 2
                cInhaltBez = " " & cInhaltBez
            Wend
            
            If Not IsNull(rsrs!artnr) Then
                cArtNr = rsrs!artnr
            Else
                cArtNr = ""
            End If
            
            While Len(Trim(cArtNr)) < 6
                cArtNr = "0" & cArtNr
            Wend
            
            If Not IsNull(rsrs!BESTAND) Then
                cBestand = rsrs!BESTAND
            Else
                cBestand = ""
            End If
            
            While Len(cBestand) < 6
                cBestand = " " & cBestand
            Wend
            
            If Not IsNull(rsrs!VKMENGE) Then
                cVKMENGE = rsrs!VKMENGE
            Else
                cVKMENGE = ""
            End If
            
            While Len(cVKMENGE) < 5
                cVKMENGE = " " & cVKMENGE
            Wend
            
            If Not IsNull(rsrs!VKPREIS) Then
                cVKPreis = Format(rsrs!VKPREIS, "#####0.00")
            Else
                cVKPreis = ""
            End If
            
            cVKEinzel = Format(CDbl(cVKPreis) / CLng(cVKMENGE), "#####0.00")
            cVKEinzel = SwapStr(cVKEinzel, ",", ".")
            
            cVKPreis = SwapStr(cVKPreis, ",", ".")
            
            While Len(cVKPreis) < 10
                cVKPreis = " " & cVKPreis
            Wend
            
            While Len(cVKEinzel) < 7
                cVKEinzel = " " & cVKEinzel
            Wend

            If Not IsNull(rsrs!AGN) Then
                cAgn = rsrs!AGN
            Else
                cAgn = ""
            End If
            
            While Len(Trim(cAgn)) < 6
                cAgn = "0" & cAgn
            Wend
            
            If Not IsNull(rsrs!ADATE) Then
                cKaufdat = rsrs!ADATE
            Else
                cKaufdat = ""
            End If
            
            cKaufdat = Format(cKaufdat, "DD.MM.YY")
            cKaufdat = SwapStr(cKaufdat, ".", "")
            
            cWeek = DatePart("ww", DateValue(rsrs!ADATE))
            
'            While Len(Trim(cKW)) < 2
'                cKW = "0" & cKW
'            Wend
            
            cSatz = cKundnr & cDatum & cEAN & cBezeich & cInhaltBez & cArtNr & cBestand & cVKMENGE & cVKEinzel & cVKPreis & cAgn & "01" & cKaufdat & cWeek
            cSatz = cSatz & Chr$(13) & Chr$(10)
            lPos = LOF(iFileNr)
            lPos = lPos + 1
            Put #iFileNr, lPos, cSatz
    
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close
    
    Close iFileNr
    
'    sSQL = "Update GFKSTAT Set LASTDATE = '" & DateValue(Now) & "'"
'    gdBase.Execute sSQL, dbFailOnError

    'GfK auf den FTPSERVER
    
    Dim bmerke As Boolean
    bmerke = gbFTPautomatic
    gbFTPautomatic = True
        
    giKissFtpMode = 29
    frmWKL38.Show 1
    
    gbFTPautomatic = bmerke
    

    
    
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "GFKerstellenJahr"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Public Function checkwarengru() As Boolean
On Error GoTo LOKAL_ERROR

    checkwarengru = False
    
    Dim cSQL As String
    
    If NewTableSuchenDBKombi("WARENGRU", gdBase) Then
        checkwarengru = True
    Else
        checkwarengru = False
    End If
            
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "checkwarengru"
    Fehler.gsFehlertext = "Beim Ermitteln der Tabelle Warengru ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Sub delARTTOINV()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    loeschNEW "ARTTOINV", gdBase
    cSQL = "Select * into ARTTOINV from Artikel where artnr = -1"
    gdBase.Execute cSQL, dbFailOnError

    cSQL = "Delete from ARTTOINV "
    gdBase.Execute cSQL, dbFailOnError
    
    SpalteAnfuegenNEW "ARTTOINV", "lfnr", "autoincrement", gdBase
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "delARTTOINV"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Function ermgesAbsatzLinr(bymonat As Byte, iJahr As Integer, lLinr As Long) As Double
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    
    ermgesAbsatzLinr = 0
    
    sSQL = "Select sum(ABSATZ) as maxi"
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
            ermgesAbsatzLinr = rsrs!maxi
        End If
    End If
    rsrs.Close
                
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "ermgesAbsatzLinr"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Function
'Public Function ermgesUmsatzLinr(byMonat As Byte, iJahr As Integer, llinr As Long) As Double
'On Error GoTo LOKAL_ERROR
'
'    Dim sSQL        As String
'    Dim rsrs        As Recordset
'
'    ermgesUmsatzLinr = 0
'
'    sSQL = "Select sum(umsatz) as maxi"
'    sSQL = sSQL & " from UMS_LINR "
'    sSQL = sSQL & " where Jahr = " & iJahr
'    If byMonat > 0 Then
'        sSQL = sSQL & " and Monat = " & byMonat
'    End If
'
'    If llinr = 0 Then
'
'    Else
'        sSQL = sSQL & " and LINR = " & llinr
'    End If
'
'    Set rsrs = gdBase.OpenRecordset(sSQL)
'    If Not rsrs.EOF Then
'        If Not IsNull(rsrs!maxi) Then
'            ermgesUmsatzLinr = rsrs!maxi
'        End If
'    End If
'    rsrs.Close
'
'Exit Function
'LOKAL_ERROR:
'
'    Fehler.gsDescr = err.Description
'    Fehler.gsNumber = err.Number
'    Fehler.gsFormular = "Modul2"
'    Fehler.gsFunktion = "ermgesUmsatzLinr"
'    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
'
'    Fehlermeldung1
'
'End Function
Public Function EinkaufsStückermittlung(cLinr As String, db As Database, iJahr As Integer, imon As Byte) As Long
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL    As String
    Dim rsrs    As Recordset
    
    EinkaufsStückermittlung = 0
    
    If Trim$(cLinr) = "" Then
        Exit Function
    End If
    
    Screen.MousePointer = 11
    
    cSQL = "Select sum(BEWEGUNG) as maxi from ZUGANG  "
    cSQL = cSQL & " where YEAR(ADATE) = " & iJahr
    cSQL = cSQL & " and LINR = " & cLinr & " "
    
    If imon > 0 Then
        cSQL = cSQL & " and Month(adate) = " & imon
    End If
    
    
    Set rsrs = db.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!maxi) Then
            EinkaufsStückermittlung = rsrs!maxi
        End If
    End If
    rsrs.Close
    
    Screen.MousePointer = 0
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "EinkaufsStückermittlung"
    Fehler.gsFehlertext = "Beim Ermitteln des Einkaufsumsatzes ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermAGNbez1(lagn As Long) As String
    On Error GoTo LOKAL_ERROR


    Dim sSQL As String
    Dim rs As Recordset
    
    ermAGNbez1 = ""
    
    sSQL = "Select AGTEXT From AGNDBF"
    sSQL = sSQL & "  where agn = " & lagn

    Set rs = gdBase.OpenRecordset(sSQL)
    
    If Not rs.EOF Then
    rs.MoveFirst
        If Not IsNull(rs!AGTEXT) Then
            ermAGNbez1 = rs!AGTEXT
        End If
    End If
    
    rs.Close
    
Exit Function
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ermAGNbez1"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
    
End Function

Public Function ermAGNbez(sArt As String, db As Database) As String
On Error GoTo LOKAL_ERROR
Dim sSQL As String
Dim rsrs As Recordset

ermAGNbez = ""

sSQL = "Select A.AGTEXT from AGNDBF A where "
sSQL = sSQL & "  a.AGN  = " & sArt

Set rsrs = db.OpenRecordset(sSQL)
If Not rsrs.EOF Then
    rsrs.MoveFirst
    If Not IsNull(rsrs!AGTEXT) Then
        ermAGNbez = rsrs!AGTEXT
    End If
End If
rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ermAGNBEZ"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Public Function ermLINBEZ(sLPZ As String, lLinr As Long, db As Database) As String
On Error GoTo LOKAL_ERROR
Dim sSQL As String
Dim rsrs As Recordset

ermLINBEZ = ""

sSQL = "Select A.LINBEZEICH from LINBEZ A where "
sSQL = sSQL & "  a.LPZ  = " & sLPZ
sSQL = sSQL & " and a.LINR  = " & lLinr

Set rsrs = db.OpenRecordset(sSQL)
If Not rsrs.EOF Then
    rsrs.MoveFirst
    If Not IsNull(rsrs!LINBEZEICH) Then
        ermLINBEZ = rsrs!LINBEZEICH
    End If
End If
rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ermLINBEZ"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Public Sub insertLINBEZ(lLpz As Long, lLinr As Long, cLiefBez As String, cMarke As String, cMarkekürzel As String)
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    
    sSQL = "Delete from LINBEZ  "
    sSQL = sSQL & "  where LPZ  = " & lLpz
    sSQL = sSQL & " and LINR  = " & lLinr
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into LINBEZ (LPZ,LINR,LINBEZEICH,MARKE,KUERZEL)"
    sSQL = sSQL & "  values "
    sSQL = sSQL & " (" & lLpz & " "
    sSQL = sSQL & " ," & lLinr & " "
    sSQL = sSQL & " ,'" & cLiefBez & "' "
    sSQL = sSQL & " ,'" & cMarke & "' "
    sSQL = sSQL & " ,'" & cMarkekürzel & "' "
    sSQL = sSQL & "  )"
    gdBase.Execute sSQL, dbFailOnError

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "insertLINBEZ"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Public Sub zeigePreislage(Listx As ListBox)
    On Error GoTo LOKAL_ERROR
    
    Dim rsrs As Recordset
    Dim sSQL As String
    Dim cSatz As String
    Dim cFeld As String
    
    Listx.Clear
    sSQL = "Select * from Preisl order by lfnr "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            If Not IsNull(rsrs!vonP) Then
            
                cFeld = Format(rsrs!vonP, "##0.00 EUR ")
                cSatz = cSatz & Space(13 - Len(cFeld)) & cFeld
                
                If Not IsNull(rsrs!bisP) Then
                
                    cFeld = Format(rsrs!bisP, "##0.00 EUR ")
                    cSatz = cSatz & "-" & Space(13 - Len(cFeld)) & cFeld
                
                    If Not IsNull(rsrs!lfnr) Then
                        cSatz = cSatz & Space(50) & rsrs!lfnr
                        Listx.AddItem cSatz
                    End If
                End If
            End If
            cFeld = ""
            cSatz = ""
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "zeigePreislage"
    Fehler.gsFehlertext = "In der Preislagenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    Resume Next
   
End Sub
Public Sub speicherPreislage(textvon As TextBox, textbis As TextBox, lblx As Label)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim cVon As String
    Dim cBis As String
    
    textvon.Text = Trim(textvon.Text)
    textbis.Text = Trim(textbis.Text)
    
    If IsNumeric(textvon.Text) Then
        If CDbl(textvon.Text) > 9999.99 Then
            anzeigeNew "rot", "Geben Sie hier einen kleineren Währungbetrag ein!", lblx
            textvon.SetFocus
            Exit Sub
        ElseIf CDbl(textvon.Text) < -9999.99 Then
            anzeigeNew "rot", "Geben Sie hier einen größeren Währungbetrag ein!", lblx
            textvon.SetFocus
            Exit Sub
        Else
            cVon = textvon.Text
        End If
    Else
        anzeigeNew "rot", "Geben Sie hier einen Währungbetrag ein!", lblx
        textvon.SetFocus
        Exit Sub
    End If
    
    If IsNumeric(textbis.Text) Then
        If CDbl(textbis.Text) > 9999.99 Then
            anzeigeNew "rot", "Geben Sie hier einen kleineren Währungbetrag ein!", lblx
            textbis.SetFocus
            Exit Sub
        ElseIf CDbl(textbis.Text) < -9999.99 Then
            anzeigeNew "rot", "Geben Sie hier einen größeren Währungbetrag ein!", lblx
            textbis.SetFocus
            Exit Sub
        Else
            cBis = textbis.Text
        End If
    Else
        anzeigeNew "rot", "Geben Sie hier einen Währungbetrag ein!", lblx
        textbis.SetFocus
        Exit Sub
    End If
        
    sSQL = "Insert into Preisl (vonP,BisP,lfnr)"
    sSQL = sSQL & " values ( '" & cVon & "', '" & cBis & "',  10 )"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "speicherPreislage"
    Fehler.gsFehlertext = "In der Preislagenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub sortPreislage()
    On Error GoTo LOKAL_ERROR
    
    Dim rsrs As Recordset
    Dim sSQL As String
    Dim lcount As Byte
   
    lcount = 0
    sSQL = "Select * from Preisl order by vonp "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            lcount = lcount + 1
            rsrs.Edit
            rsrs!lfnr = lcount
            rsrs.Update
            
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "sortPreislage"
    Fehler.gsFehlertext = "In der Preislagenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Public Sub fülleCboEtiketten(cbox As ComboBox)
    On Error GoTo LOKAL_ERROR
    
    cbox.Clear
    cbox.AddItem "69 x 14 (Var 1)"
    cbox.AddItem "69 x 14 (Var 2)"
    cbox.AddItem "69 x 14 (Var 3)"
    
    cbox.AddItem "40 x 18 (Var 1)"
    cbox.AddItem "40 x 18 (Var 2)"
    cbox.AddItem "40 x 18 (Var 3)"
    cbox.AddItem "40 x 18 (Var 4)"
    cbox.AddItem "40 x 18 (Var 5)"
    cbox.AddItem "40 x 18 (Var 6)"
    
    cbox.AddItem "45 x 23 (Var 1)"
    cbox.AddItem "45 x 23 (Var 2)"
    cbox.AddItem "45 x 23 (Var 3)"
    cbox.AddItem "45 x 23 (Var 4)"
    
    cbox.AddItem "38 x 23 (Var 1)"
    cbox.AddItem "38 x 23 (Var 2)"
    cbox.AddItem "38 x 23 (Var 3)"
    
    cbox.AddItem "51 x 19 (Var 1)"
    cbox.AddItem "51 x 19 (Var 2)"
    cbox.AddItem "51 x 19 (Var 3)"
    
    cbox.AddItem "49 x 19 (Var 1)"
    
    cbox.AddItem "44 x 21 (Var 1)"
    
    cbox.AddItem "30 x 15 (Var 1)"
    cbox.AddItem "30 x 15 (Var 2)"
    cbox.AddItem "30 x 15 (Var 3)"
    
    cbox.AddItem "48 x 18 (Var 1)"
    
    cbox.AddItem "35 x 15 (Var 1)"
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "fülleCboEtiketten"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub fülleCboEtikettenSpezTermin(cbox As ComboBox)
    On Error GoTo LOKAL_ERROR
    
    cbox.Clear
    cbox.AddItem "bitte wählen"
    cbox.AddItem "50 x 40"
    cbox.Text = "bitte wählen"
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "fülleCboEtikettenSpezTermin"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub füllePreislage(cbox As ComboBox)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    Dim cSatz As String
    Dim cFeld As String
    
    If Not NewTableSuchenDBKombi("PREISL", gdBase) Then
        CreateTable "PREISL", gdBase
        
        sSQL = "Insert into Preisl (vonP,BisP,lfnr)"
        sSQL = sSQL & " values ( '4,99', '9,99' , 1 )"
        schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
        
        sSQL = "Insert into Preisl (vonP,BisP,lfnr)"
        sSQL = sSQL & " values ( '10,00', '14,99' , 2 )"
        schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
        
        sSQL = "Insert into Preisl (vonP,BisP,lfnr)"
        sSQL = sSQL & " values ( '15,00', '19,99' , 3 )"
        schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
        
        sSQL = "Insert into Preisl (vonP,BisP,lfnr)"
        sSQL = sSQL & " values ( '20,00', '29,99' , 4 )"
        schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
        
        sSQL = "Insert into Preisl (vonP,BisP,lfnr)"
        sSQL = sSQL & " values ( '30,00', '49,99' , 5 )"
        schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
        
        sSQL = "Insert into Preisl (vonP,BisP,lfnr)"
        sSQL = sSQL & " values ( '50,00', '99,99' , 6 )"
        schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
        
    End If
    
    cbox.Clear
    cbox.AddItem "alle Preislagen"
    cbox.Text = "alle Preislagen"
    
    sSQL = "Select * from preisl order by lfnr"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!vonP) Then
            
                cFeld = Format(rsrs!vonP, "##0.00 EUR ")
                cSatz = cSatz & Space(13 - Len(cFeld)) & cFeld
                
                If Not IsNull(rsrs!bisP) Then
                
                    cFeld = Format(rsrs!bisP, "##0.00 EUR ")
                    cSatz = cSatz & "-" & Space(13 - Len(cFeld)) & cFeld
                
                    If Not IsNull(rsrs!lfnr) Then
                        cSatz = cSatz & Space(50) & rsrs!lfnr
                        cbox.AddItem cSatz
                    End If
                End If
            End If
            cSatz = ""
            cFeld = ""
        
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "füllePreislage"
    Fehler.gsFehlertext = "In der Preislagenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub füllePreislagestandard(cbox As ComboBox)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    Dim cSatz As String
    Dim cFeld As String
    
    loeschNEW "PREISL", gdBase
    CreateTable "PREISL", gdBase
    
    sSQL = "Insert into Preisl (vonP,BisP,lfnr)"
    sSQL = sSQL & " values ( '4,99', '9,99' , 1 )"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into Preisl (vonP,BisP,lfnr)"
    sSQL = sSQL & " values ( '10,00', '14,99' , 2 )"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into Preisl (vonP,BisP,lfnr)"
    sSQL = sSQL & " values ( '15,00', '19,99' , 3 )"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into Preisl (vonP,BisP,lfnr)"
    sSQL = sSQL & " values ( '20,00', '29,99' , 4 )"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into Preisl (vonP,BisP,lfnr)"
    sSQL = sSQL & " values ( '30,00', '49,99' , 5 )"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into Preisl (vonP,BisP,lfnr)"
    sSQL = sSQL & " values ( '50,00', '99,99' , 6 )"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    
    cbox.Clear
    cbox.AddItem "alle Preislagen"
    cbox.Text = "alle Preislagen"
    
    sSQL = "Select * from preisl order by lfnr"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!vonP) Then
            
                cFeld = Format(rsrs!vonP, "##0.00 EUR ")
                cSatz = cSatz & Space(13 - Len(cFeld)) & cFeld
                
                If Not IsNull(rsrs!bisP) Then
                
                    cFeld = Format(rsrs!bisP, "##0.00 EUR ")
                    cSatz = cSatz & "-" & Space(13 - Len(cFeld)) & cFeld
                
                    If Not IsNull(rsrs!lfnr) Then
                        cSatz = cSatz & Space(50) & rsrs!lfnr
                        cbox.AddItem cSatz
                    End If
                End If
            End If
            cSatz = ""
            cFeld = ""
        
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "füllePreislagestandard"
    Fehler.gsFehlertext = "In der Preislagenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub delPreislage(lf As Byte)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    sSQL = "Delete from preisl where lfnr = " & lf
    gdBase.Execute sSQL, dbFailOnError

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "delPreislage"
    Fehler.gsFehlertext = "In der Preislagenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub wählePreislage(textvon As TextBox, textbis As TextBox, Listx As ListBox)
    On Error GoTo LOKAL_ERROR
    
    Dim cLBSatz As String
    cLBSatz = Listx.list(Listx.ListIndex)
   
    textvon.Text = Trim(Mid(cLBSatz, 1, 8))
    textbis.Text = Trim(Mid(cLBSatz, 16, 8))
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "wählePreislage"
    Fehler.gsFehlertext = "In der Preislagenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub

Public Sub leseZeitSteu()
On Error GoTo LOKAL_ERROR

Dim cSQL As String
Dim rsrs As Recordset

If NewTableSuchenDBKombi("ZEITSTEU", gdApp) = False Then
    CreateTable "ZEITSTEU", gdApp
End If

cSQL = "select * from ZEITSTEU "
Set rsrs = gdApp.OpenRecordset(cSQL)
If Not rsrs.EOF Then
    If Not IsNull(rsrs!adresse) Then
        frmWKL08.Text1(1).Text = rsrs!adresse
        gsVMPadresse = rsrs!adresse
    Else
        gsVMPadresse = ""
    End If
    
    If Not IsNull(rsrs!BETREFF) Then
        frmWKL08.Text1(3).Text = rsrs!BETREFF
        gsVMPbetreff = rsrs!BETREFF
    Else
        gsVMPbetreff = ""
    End If
    
    If Not IsNull(rsrs!KdNr) Then
        frmWKL08.Text1(2).Text = rsrs!KdNr
        gsVMPKdNr = rsrs!KdNr
    Else
        gsVMPKdNr = ""
    End If
    
    If Not IsNull(rsrs!zLinr) Then
        frmWKL08.Text1(4).Text = rsrs!zLinr
        gsVMPzLinr = rsrs!zLinr
    Else
        gsVMPzLinr = ""
    End If
    
    If Not IsNull(rsrs!Endung) Then
        If Val(rsrs!Endung) = 1 Then
            frmWKL08.Option1(0).value = True
        ElseIf Val(rsrs!Endung) = 2 Then
            frmWKL08.Option1(1).value = True
        End If
        gsVMPEndung = rsrs!Endung
    Else
        gsVMPEndung = ""
    End If
    
    If Not IsNull(rsrs!art) Then
        If Val(rsrs!art) = 1 Then
            frmWKL08.Option2(0).value = True
        ElseIf Val(rsrs!art) = 2 Then
            frmWKL08.Option2(1).value = True
        End If
        gsVMPArt = rsrs!art
    Else
        gsVMPArt = ""
    End If

End If
rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
  
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "leseZeitSteu"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub lese_Ex_Steu()
On Error GoTo LOKAL_ERROR

Dim cSQL As String
Dim rsrs As Recordset

If NewTableSuchenDBKombi("EXSTEU", gdApp) = False Then
    CreateTableT3 "EXSTEU", gdApp
Else
    If SpalteInTabellegefundenNEW("EXSTEU", "PLUSEAN", gdApp) = False Then
        SpalteAnfuegenNEW "EXSTEU", "EXNOR", "BIT", gdApp
        SpalteAnfuegenNEW "EXSTEU", "BL", "BIT", gdApp
        SpalteAnfuegenNEW "EXSTEU", "BLKENNUNG", "Text(20)", gdApp
        SpalteAnfuegenNEW "EXSTEU", "PLUSEAN", "BIT", gdApp
        
        cSQL = "Update Exsteu set exnor = true"
        gdApp.Execute cSQL, dbFailOnError
        
        cSQL = "Update Exsteu set PLUSEAN = false"
        gdApp.Execute cSQL, dbFailOnError
    End If
    
    If SpalteInTabellegefundenNEW("EXSTEU", "PLUSBEZEICH", gdApp) = False Then
        SpalteAnfuegenNEW "EXSTEU", "PLUSBEZEICH", "BIT", gdApp
        
        cSQL = "Update Exsteu set PLUSBEZEICH = false"
        gdApp.Execute cSQL, dbFailOnError
    End If
    
    If SpalteInTabellegefundenNEW("EXSTEU", "DATEIENDUNG", gdApp) = False Then
        SpalteAnfuegenNEW "EXSTEU", "DATEIENDUNG", "Text(3)", gdApp
        
        cSQL = "Update Exsteu set DATEIENDUNG = 'txt'"
        gdApp.Execute cSQL, dbFailOnError
    End If
    
    If SpalteInTabellegefundenNEW("EXSTEU", "FELDTRENNER", gdApp) = False Then
        SpalteAnfuegenNEW "EXSTEU", "FELDTRENNER", "Text(20)", gdApp
        
        cSQL = "Update Exsteu set FELDTRENNER = 'Tab'"
        gdApp.Execute cSQL, dbFailOnError
    End If
    
    
    
    If SpalteInTabellegefundenNEW("EXSTEU", "SHOPARTIKEL", gdApp) = False Then
        SpalteAnfuegenNEW "EXSTEU", "SHOPARTIKEL", "BIT", gdApp
        
        cSQL = "Update Exsteu set SHOPARTIKEL = false"
        gdApp.Execute cSQL, dbFailOnError
    End If
    
    If SpalteInTabellegefundenNEW("EXSTEU", "ShopPreis", gdApp) = False Then
        SpalteAnfuegenNEW "EXSTEU", "ShopPreis", "BIT", gdApp
        
        cSQL = "Update Exsteu set ShopPreis = false"
        gdApp.Execute cSQL, dbFailOnError
    End If
    
    If SpalteInTabellegefundenNEW("EXSTEU", "MITUEBERSCHRIFT", gdApp) = False Then
        SpalteAnfuegenNEW "EXSTEU", "MITUEBERSCHRIFT", "BIT", gdApp
        
        cSQL = "Update Exsteu set MITUEBERSCHRIFT = false"
        gdApp.Execute cSQL, dbFailOnError
    End If
    
    
    
End If

cSQL = "select * from EXSTEU "
Set rsrs = gdApp.OpenRecordset(cSQL)
If Not rsrs.EOF Then
    If Not IsNull(rsrs!nurmitBestand) Then
        gbNMB = rsrs!nurmitBestand
    Else
        gbNMB = False
    End If
    
    If Not IsNull(rsrs!EXNOR) Then
        gbEXNOR = rsrs!EXNOR
    Else
        gbEXNOR = False
    End If
    
    If Not IsNull(rsrs!SHOPARTIKEL) Then
        gbSHOPARTIKEL = rsrs!SHOPARTIKEL
    Else
        gbSHOPARTIKEL = False
    End If
    
    If Not IsNull(rsrs!BL) Then
        gbBL = rsrs!BL
    Else
        gbBL = False
    End If
    
    If Not IsNull(rsrs!BLKENNUNG) Then
        gsBLKENNUNG = rsrs!BLKENNUNG
    Else
        gsBLKENNUNG = ""
    End If
    
    If Not IsNull(rsrs!PLUSEAN) Then
        gbPlusEAN = rsrs!PLUSEAN
    Else
        gbPlusEAN = False
    End If
    
    If Not IsNull(rsrs!PLUSBEZEICH) Then
        gbPlusBezeich = rsrs!PLUSBEZEICH
    Else
        gbPlusBezeich = False
    End If
    
    If Not IsNull(rsrs!DATEIENDUNG) Then
        gsDATEIENDUNG = rsrs!DATEIENDUNG
    Else
        gsDATEIENDUNG = "txt"
    End If
    
    If Not IsNull(rsrs!FELDTRENNER) Then
        gsFELDTRENNER = rsrs!FELDTRENNER
    Else
        gsFELDTRENNER = "Tab"
    End If
    
    If Not IsNull(rsrs!ShopPreis) Then
        gbPlusShopPreis = rsrs!ShopPreis
    Else
        gbPlusShopPreis = False
    End If
    
    If Not IsNull(rsrs!MITUEBERSCHRIFT) Then
        gbMITUEBERSCHRIFT = rsrs!MITUEBERSCHRIFT
    Else
        gbMITUEBERSCHRIFT = False
    End If
    
End If
rsrs.Close: Set rsrs = Nothing
    
Exit Sub
LOKAL_ERROR:
  
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "lese_Ex_Steu"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Function VMPZeitPunktwegschicken() As String
On Error GoTo LOKAL_ERROR

    Dim lDatum      As Long
    Dim sFText      As String
    Dim sZeitung    As String
    Dim ldatum1Jan  As Long
    Dim lex         As Long
    Dim cex         As String
    Dim cPfad       As String
    
    cPfad = gcDBPfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    cPfad = cPfad & "ZEITUNG"

    lDatum = DateValue(Now)
    ldatum1Jan = DateValue("01.01." & Year(DateValue(Now)))
    lex = lDatum - ldatum1Jan + 1
    cex = CStr(lex)
    
    If Len(cex) = 1 Then
        cex = "00" & cex
    ElseIf Len(cex) = 2 Then
        cex = "0" & cex
    End If
    
    If gsVMPEndung = "1" Then
    
    ElseIf gsVMPEndung = "2" Then
        cex = "txt"
    End If
    
    Kill cPfad & "\" & Val(gsVMPKdNr) & ".*"
    sZeitung = cPfad & "\" & Val(gsVMPKdNr) & "." & cex

    VMPZeitPunktwegschicken = ""

    sFText = AUSwertungZP(0, sZeitung, gsVMPKdNr, Val(gsVMPzLinr))

    If sFText <> "" Then

    Else
        Dim sAbsenderadresse As String
        


        schickeMailimHintergrundSSL ermFirmenBez, ermFirmenMail, ermFirmenMail, gsVMPadresse _
        , ermFirmenMail, gcSMTP_SERVER, gcSMTP_PORT, gcSMTP_USER, gcSMTP_PW, gsVMPbetreff, "Zeitungsdaten", sZeitung
        

        
        gcBestellEmail.Attachment1 = ""
        gcBestellEmail.Subject = ""
        gcBestellEmail.Message = ""
        gcBestellEmail.Recipient = ""
    End If

Exit Function
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "VMPZeitPunktwegschicken"
        Fehler.gsFehlertext = "Es trat ein Fehler auf. "
        Fehlermeldung1
    End If
End Function
Public Function Export_Artikelbestände() As Boolean
On Error GoTo LOKAL_ERROR

    Dim cPfad                   As String
    Dim iFileNr                 As Integer
    Dim lPos                    As Long
    Dim cSatz                   As String
    Dim rsrs                    As Recordset
    Dim cSQL                    As String
    Dim cFeldtrennZeichen       As String
    
    Dim cLEK                    As String
    Dim cKVK                    As String
    Dim cSHOPKVK                As String
    Dim cBezeich                As String
    
    Dim cUeber                  As String
    
    Select Case gsFELDTRENNER
        Case "Tab"
            cFeldtrennZeichen = vbTab
        Case "Komma"
            cFeldtrennZeichen = ","
        Case "Semikolon"
            cFeldtrennZeichen = ";"
    End Select
        
    
    Export_Artikelbestände = False
    
    cPfad = gcDBPfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    cPfad = cPfad & "STAT\"
    
    lese_Ex_Steu
    
    Kill cPfad & "Best_Fil_" & gcFilNr & "." & gsDATEIENDUNG
    
    iFileNr = FreeFile
    Open cPfad & "Best_Fil_" & gcFilNr & "." & gsDATEIENDUNG For Binary As #iFileNr
    
    
    
    loeschNEW "tart", gdBase
    cSQL = "Select ARTIKEL.artnr,ARTIKEL.KVKPR1,ARTIKEL.bestand,'0' as LEKPR,ARTIKEL.EAN,ARTIKEL.EAN2,ARTIKEL.EAN3,ARTIKEL.BEZEICH,'0' as SHOPKVK into tArt from ARTIKEL "
    
    If gbSHOPARTIKEL = True Then
        cSQL = cSQL & " inner join INTERART on Artikel.artnr = INTERART.artnr "
    End If
    
    If gbNMB Then
        cSQL = cSQL & " where bestand > 0 "
    End If
    gdBase.Execute cSQL, dbFailOnError
    
    loeschNEW "miartl", gdBase
    cSQL = "Select min(LEKPR) as lek , artnr  into miartl from Artlief group by artnr "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "update tart inner join miartl on tart.artnr = miartl.artnr set tart.lekpr = miartl.lek  "
    gdBase.Execute cSQL, dbFailOnError
    
    
    If gbSHOPARTIKEL Then
    
        cSQL = "update tart inner join INTERART on tart.artnr = INTERART.artnr set tart.SHOPKVK = INTERART.SHOPKVK  "
        gdBase.Execute cSQL, dbFailOnError
    
    End If
    
    
    cSQL = " select * from tart order by Bestand desc "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        
        If gbMITUEBERSCHRIFT Then
                
            cUeber = "Artnr" & cFeldtrennZeichen & "LEK" & cFeldtrennZeichen & "KVK" & cFeldtrennZeichen & "Bestand"
            
            If gbPlusEAN Then
                cUeber = cUeber & cFeldtrennZeichen
                cUeber = cUeber & "EAN" & cFeldtrennZeichen
                cUeber = cUeber & "EAN2" & cFeldtrennZeichen
                cUeber = cUeber & "EAN3"
            End If
            
            If gbPlusBezeich Then
                cUeber = cUeber & cFeldtrennZeichen
                cUeber = cUeber & "Bezeich"
            End If
            
            If gbPlusShopPreis Then
                cUeber = cUeber & cFeldtrennZeichen
                cUeber = cUeber & "SHOPKVK"
            End If
        
        
            cUeber = cUeber & vbCrLf
        
            lPos = LOF(iFileNr)
            lPos = lPos + 1
            Put #iFileNr, lPos, cUeber
        
        End If
        
        
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!artnr) Then
            
                
            
                cLEK = "0"
                If Not IsNull(rsrs!lekpr) Then
                    cLEK = rsrs!lekpr
                End If
                
                cKVK = "0"
                If Not IsNull(rsrs!KVKPR1) Then
                    cKVK = rsrs!KVKPR1
                End If
                
                cSHOPKVK = "0"
                If Not IsNull(rsrs!SHOPKVK) Then
                    cSHOPKVK = rsrs!SHOPKVK
                End If
                
                
                
                cBezeich = ""
                If Not IsNull(rsrs!BEZEICH) Then
                    cBezeich = rsrs!BEZEICH
                End If
                
                If gsFELDTRENNER = "Komma" Then
                    cLEK = SwapStr(cLEK, ",", ".")
                    cKVK = SwapStr(cKVK, ",", ".")
                    cSHOPKVK = SwapStr(cSHOPKVK, ",", ".")
                    cBezeich = SwapStr(cBezeich, ",", ".")
                ElseIf gsFELDTRENNER = "Semikolon" Then
                    cBezeich = SwapStr(cBezeich, ";", ",")
                End If
            
                cSatz = ""
                cSatz = cSatz & rsrs!artnr & cFeldtrennZeichen
                cSatz = cSatz & cLEK & cFeldtrennZeichen
                cSatz = cSatz & cKVK & cFeldtrennZeichen
                cSatz = cSatz & rsrs!BESTAND
                
                If gbPlusEAN Then
                    cSatz = cSatz & cFeldtrennZeichen
                    cSatz = cSatz & rsrs!EAN & cFeldtrennZeichen
                    cSatz = cSatz & rsrs!EAN2 & cFeldtrennZeichen
                    cSatz = cSatz & rsrs!EAN3
                End If
                
                If gbPlusBezeich Then
                    cSatz = cSatz & cFeldtrennZeichen
                    cSatz = cSatz & cBezeich
                End If
                
                If gbPlusShopPreis Then
                    cSatz = cSatz & cFeldtrennZeichen
                    If cSHOPKVK = cKVK Then
                        cSHOPKVK = "0"
                    End If
                    cSatz = cSatz & cSHOPKVK
                   
                End If
                
                
                
                
                
                cSatz = cSatz & vbCrLf
                
                lPos = LOF(iFileNr)
                lPos = lPos + 1
                Put #iFileNr, lPos, cSatz
                
            End If
            rsrs.MoveNext
        Loop
    End If

    rsrs.Close
    Close iFileNr
    
    loeschNEW "tart", gdBase
    loeschNEW "miartl", gdBase
    
    Export_Artikelbestände = True

Exit Function
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "Export_Artikelbestände"
        Fehler.gsFehlertext = "Es trat ein Fehler auf. "
        Fehlermeldung1
'        Resume Next
    End If
End Function
Public Function Export_Artikelbestände_Komplett_Vedes() As Boolean
On Error GoTo LOKAL_ERROR

    Dim cPfad               As String
    Dim iFileNr             As Integer
    Dim lPos                As Long
    Dim cSatz               As String
    Dim rsrs                As Recordset
    Dim cSQL                As String
    Dim sdateiname          As String
    
    Dim cFeldtrennZeichen   As String
    cFeldtrennZeichen = ";"
   
    Dim slibesnr            As String
    Dim sEAN                As String
    Dim sLinr               As String
    Dim cKVK                As String
    
    
    
    sdateiname = "retailer_" & Format(DateValue(Now), "YYYYMMDD") & Format(TimeValue(Now), "HHMMSS") & "_daily.csv"

    Export_Artikelbestände_Komplett_Vedes = False
    
    cPfad = gcDBPfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    cPfad = cPfad & "VEDESDSL\"
    
    lese_Ex_Steu
    
    
    Kill cPfad & sdateiname
    iFileNr = FreeFile
    Open cPfad & sdateiname For Binary As #iFileNr
    
    
    cSatz = ""
    cSatz = cSatz & "Kundennummer" & cFeldtrennZeichen
    cSatz = cSatz & "VEDES-Sortimentsnummer" & cFeldtrennZeichen
    cSatz = cSatz & "EAN" & cFeldtrennZeichen
    cSatz = cSatz & "Lieferantennummer" & cFeldtrennZeichen
    cSatz = cSatz & "Bestand" & cFeldtrennZeichen
    cSatz = cSatz & "VK-Preis" & cFeldtrennZeichen
    cSatz = cSatz & "Kennzeichen ohne GH" & cFeldtrennZeichen
    
    cSatz = cSatz & "Kennzeichen ohne Eigenversand" & cFeldtrennZeichen
    cSatz = cSatz & "Lieferzeit in Tagen" & cFeldtrennZeichen
    cSatz = cSatz & "VK-Preis Liste" & cFeldtrennZeichen
    cSatz = cSatz & "VK-Preis POS" & cFeldtrennZeichen
    cSatz = cSatz & "VK-Preis VEDES Marktplätze" & cFeldtrennZeichen
    cSatz = cSatz & "Versandkosten-Art" & cFeldtrennZeichen
    cSatz = cSatz & "Bestand Demo" & cFeldtrennZeichen
    cSatz = cSatz & "Bestand Drittlager" & cFeldtrennZeichen
    cSatz = cSatz & "" & vbCrLf

    lPos = LOF(iFileNr)
    lPos = lPos + 1
    Put #iFileNr, lPos, cSatz
    

    loeschNEW "tart", gdBase
    cSQL = "Select ARTIKEL.artnr,ARTIKEL.KVKPR1,ARTIKEL.bestand,'0' as LEKPR,0 as LINR, '' as Libesnr,ARTIKEL.EAN,ARTIKEL.EAN2,ARTIKEL.EAN3,ARTIKEL.BEZEICH into tArt from ARTIKEL "
    cSQL = cSQL & " inner join INTERART on Artikel.artnr = INTERART.artnr "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update tart inner join artlief on "
    cSQL = cSQL & " tart.artnr = artlief.artnr "
    cSQL = cSQL & " set tart.linr = artlief.linr "
    cSQL = cSQL & " where artlief.RKZ = 'N' "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from tart "
    cSQL = cSQL & " where linr = 0 "
    gdBase.Execute cSQL, dbFailOnError
    
    
    
    cSQL = "update tart set bestand = 0 where bestand is null "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = " select * from tart order by Bestand desc "
    
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
        
            sEAN = ""
            If Not IsNull(rsrs!EAN) Then
                sEAN = rsrs!EAN
            End If
            
            If sEAN = "" Then
                If Not IsNull(rsrs!EAN2) Then
                    sEAN = rsrs!EAN2
                End If
            End If
            
            If sEAN = "" Then
                If Not IsNull(rsrs!EAN3) Then
                    sEAN = rsrs!EAN3
                End If
            End If
            
            slibesnr = ""
            If Not IsNull(rsrs!LIBESNR) Then
                slibesnr = rsrs!LIBESNR
            End If
            
            sLinr = ""
            If Not IsNull(rsrs!linr) Then
                sLinr = rsrs!linr
            End If
            
'
'            sLinr = "0008"
            
            cKVK = "0"
            If Not IsNull(rsrs!KVKPR1) Then
                cKVK = rsrs!KVKPR1
            End If
            
            
        
            If sEAN <> "" Or slibesnr <> "" Then
                cSatz = ""
                cSatz = cSatz & gsBLKENNUNG & cFeldtrennZeichen
                cSatz = cSatz & slibesnr & cFeldtrennZeichen
                cSatz = cSatz & sEAN & cFeldtrennZeichen
                cSatz = cSatz & sLinr & cFeldtrennZeichen
                cSatz = cSatz & rsrs!BESTAND & cFeldtrennZeichen
                cSatz = cSatz & cKVK & cFeldtrennZeichen
                cSatz = cSatz & "" & cFeldtrennZeichen
                
                
                cSatz = cSatz & cFeldtrennZeichen & cFeldtrennZeichen & cFeldtrennZeichen & cFeldtrennZeichen & cFeldtrennZeichen & cFeldtrennZeichen & cFeldtrennZeichen
                
                
                
                
                cSatz = cSatz & "" & vbCrLf
                
                lPos = LOF(iFileNr)
                lPos = lPos + 1
                Put #iFileNr, lPos, cSatz
            End If
            
            rsrs.MoveNext
        Loop
    End If

    rsrs.Close
    Close iFileNr
    
    Export_Artikelbestände_Komplett_Vedes = True

Exit Function
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "Export_Artikelbestände_Komplett_Vedes"
        Fehler.gsFehlertext = "Es trat ein Fehler auf. "
        Fehlermeldung1
    End If
End Function
Private Sub Emailverschicken()
'    On Error GoTo LOKAL_ERROR
'
'    ' =============================
'    ' Hier beginnt der Sendevorgang
'    ' =============================
'    Screen.MousePointer = 11
'    mailocxcheck
'
'
'
'
'    With frmWKL00.sevSMTP1
'        .SenderName = gcBestellEmail.SenderName
'        .ReplyTo = gcBestellEmail.ReplyTo
'        .SenderEMail = gcBestellEmail.SenderEMail
'
'        gbCCfromBestlief = True
'        If gbCCfromBestlief = True Then
'            .CC = gcBestellEmail.CC
'        End If
'
'        If gcBestellEmail.BCC <> "" Then
'            .BCC = gcBestellEmail.BCC
'        End If
'
'        .Recipient = gcBestellEmail.Recipient
'        .SMTPAUTH = gcBestellEmail.SMTPAUTH
'        .ServerName = gcBestellEmail.ServerName
'        .ServerPort = gcBestellEmail.ServerPort
'        .Username = gcBestellEmail.Username
'        .Password = gcBestellEmail.Password
'        .Subject = gcBestellEmail.Subject
'        .Message = gcBestellEmail.Message
'        .AutoZIP = gcBestellEmail.AutoZIP
'
'
'        .AttachmentClear
'        If gcBestellEmail.Attachment1 <> "" Then
'            .AttachmentAdd gcBestellEmail.Attachment1
'        End If
'
'        If gcBestellEmail.Attachment2 <> "" Then
'            .AttachmentAdd gcBestellEmail.Attachment2
'        End If
'
'        If gcBestellEmail.Attachment3 <> "" Then
'            .AttachmentAdd gcBestellEmail.Attachment3
'        End If
'
'        If gcBestellEmail.Attachment4 <> "" Then
'            .AttachmentAdd gcBestellEmail.Attachment4
'        End If
'
'        If gcBestellEmail.Attachment5 <> "" Then
'            .AttachmentAdd gcBestellEmail.Attachment5
'        End If
'    End With
'
'    ' Anmeldung erfolgreich?
'    If frmWKL00.sevSMTP1.Connect() = True Then
'        Screen.MousePointer = 11
'        Dim lngBytesSent As Long
'        lngBytesSent = frmWKL00.sevSMTP1.SendMail()
'
'        frmWKL00.sevSMTP1.Disconnect
'    End If
'
'    Screen.MousePointer = 0
'
'Exit Sub
'LOKAL_ERROR:
'    Fehler.gsDescr = err.Description
'    Fehler.gsNumber = err.Number
'    Fehler.gsFormular = "Modul2"
'    Fehler.gsFunktion = "Emailverschicken"
'    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
'
'    Fehlermeldung1
End Sub
Private Sub EmailverschickenSSL()
    On Error GoTo LOKAL_ERROR

    ' =============================
    ' Hier beginnt der Sendevorgang
    ' =============================
    Screen.MousePointer = 11
    
    mailDLLcheck
    
    
    Dim From As String
    Dim to_addr As String
    Dim cc_addr As String
    Dim ServerAddr As String
    
    
    From = gcBestellEmail.SenderName '(textFrom.Text)
    to_addr = gcBestellEmail.Recipient
    cc_addr = gcBestellEmail.CC
    ServerAddr = gcBestellEmail.ServerName
   
    
    
    
    'Declare and create easendmail mail object instance
    Dim oSmtp As EASendMailObjLib.Mail
    Set oSmtp = New EASendMailObjLib.Mail
    'The license code for EASendMail ActiveX Object,
    'for evaluation usage, please use "TryIt" as the license code.
    
    
    'ES-D1508812687-00538-A396A7A3E3AC9A9B-U1EB72UA1E5C16FC
                         
    oSmtp.LicenseCode = "ES-D1508812687-00538-A396A7A3E3AC9A9B-U1EB72UA1E5C16FC"
    'oSmtp.LogFileName = App.Path & "\smtp.txt" 'enable smtp log
    oSmtp.ServerAddr = gcBestellEmail.ServerName
    oSmtp.ServerPort = gcBestellEmail.ServerPort
    
    
    
    oSmtp.Protocol = 0 'lstProtocol.ListIndex
    
    If gcBestellEmail.ServerName <> "" Then
        
        oSmtp.ServerPort = CLng(gcBestellEmail.ServerPort)
        
        
        oSmtp.Username = Trim(gcBestellEmail.Username)
        oSmtp.Password = Trim(gcBestellEmail.Password)
        
        If gcBestellEmail.SSL = True Then
       
            oSmtp.SSL_init
            'If SSL port is 465 or other port rather than 25 or 587 port, please use
            'oSmtp.ServerPort = 465
            'oSmtp.SSL_starttls = 0
        End If
    End If
    
    
    
    oSmtp.Charset = "utf-8" 'm_arCharset(lstCharset.ListIndex, 1)
'    Dim name, addr As String
'    fnParseAddr From, name, addr
    
    'Using this email to be replied to another address
    'oSmtp.ReplyTo = ReplyAddress
    
    oSmtp.From = gcBestellEmail.SenderName ' name
    oSmtp.FromAddr = gcBestellEmail.SenderEMail 'addr
    
    'add digital signature
    oSmtp.SignerCert.Unload
'    If chkSign.Value = Checked Then
'        If Not oSmtp.SignerCert.FindSubject(addr, CERT_SYSTEM_STORE_CURRENT_USER, "my") Then
'            MsgBox oSmtp.SignerCert.GetLastError() & ":" & addr
'            btnSend.Enabled = True
'        Exit Sub
'        End If
'        If Not oSmtp.SignerCert.HasPrivateKey Then
'            MsgBox "Signer certificate has not private key, this certificate can not be used to sign email!"
'            btnSend.Enabled = True
'            Exit Sub
'        End If
'    End If
    
    oSmtp.AddRecipientEx to_addr, 0  ' 0, Normal recipient, 1, cc, 2, bcc
    oSmtp.AddRecipientEx cc_addr, 0
    
    Dim recipients As String
    recipients = to_addr & "," & cc_addr
    fnTrim recipients, ","
    
    Dim i, Count As Integer
    'encrypt email by recipients certificate
    oSmtp.RecipientsCerts.Clear
'    If chkEncrypt.Value = Checked Then
'        Dim arAddr
'        arAddr = SplitEx(recipients, ",")   'split the multiple address to an array
'        Count = UBound(arAddr)
'        For i = LBound(arAddr) To Count
'            addr = arAddr(i)
'            fnTrim addr, " ,;"
'            If addr <> "" Then
'                'find the encrypting certificate for every recipients
'                Dim oEncryptCert As New EASendMailObjLib.Certificate
'                If Not oEncryptCert.FindSubject(addr, CERT_SYSTEM_STORE_CURRENT_USER, "AddressBook") Then
'                    If Not oEncryptCert.FindSubject(addr, CERT_SYSTEM_STORE_CURRENT_USER, "my") Then
'                        MsgBox oEncryptCert.GetLastError() & ":" & addr
'                        btnSend.Enabled = True
'                        Exit Sub
'                    End If
'                End If
'                oSmtp.RecipientsCerts.Add oEncryptCert
'            End If
'        Next
'    End If
    
    
    Dim m_arAttachment() As String
    Dim iCount As Integer
    iCount = 0
    
    
    ReDim m_arAttachment(iCount)
    
    If gcBestellEmail.Attachment1 <> "" Then
        m_arAttachment(iCount) = gcBestellEmail.Attachment1
    End If
    
    iCount = iCount + 1: ReDim Preserve m_arAttachment(iCount)
    
    
    
    
    If gcBestellEmail.Attachment2 <> "" Then
        m_arAttachment(iCount) = gcBestellEmail.Attachment2
    End If
    
    iCount = iCount + 1: ReDim Preserve m_arAttachment(iCount)
    
    
    If gcBestellEmail.Attachment3 <> "" Then
        m_arAttachment(iCount) = gcBestellEmail.Attachment3
    End If
    
    iCount = iCount + 1: ReDim Preserve m_arAttachment(iCount)
    
    If gcBestellEmail.Attachment4 <> "" Then
        m_arAttachment(iCount) = gcBestellEmail.Attachment4
    End If
    
    iCount = iCount + 1: ReDim Preserve m_arAttachment(iCount)
    
    
    If gcBestellEmail.Attachment5 <> "" Then
        m_arAttachment(iCount) = gcBestellEmail.Attachment5
    End If
    
    iCount = iCount + 1: ReDim Preserve m_arAttachment(iCount)
    
    If gcBestellEmail.Attachment6 <> "" Then
        m_arAttachment(iCount) = gcBestellEmail.Attachment6
    End If
    
    iCount = iCount + 1: ReDim Preserve m_arAttachment(iCount)
    
    
    
    iCount = UBound(m_arAttachment)
    For i = 0 To iCount - 1
        If oSmtp.AddAttachment(m_arAttachment(i)) <> 0 Then
'            MsgBox oSmtp.GetLastErrDescription() & ":" & m_arAttachment(i)
'            btnSend.Enabled = True
'            Exit Sub
        End If
    Next
    
    
    
    
    
    
    
    
    
    Dim Subject As String
    Dim Bodytext As String
    
    Subject = gcBestellEmail.Subject
    Bodytext = gcBestellEmail.Message
    
'    Bodytext = Replace(Bodytext, "[$from]", From)
'    Bodytext = Replace(Bodytext, "[$to]", recipients)
'    Bodytext = Replace(Bodytext, "[$subject]", Subject)
'
    oSmtp.Subject = Subject
    oSmtp.Bodytext = Bodytext
    
    'oSmtp.BodyFormat = 1    ' Using HTML FORMAT to send mail
    
'''    If InStr(1, recipients, ",", 1) > 1 And ServerAddr = "" Then
'''        'To send email without specified smtp server, we have to send the emails one by one
'''        ' to multiple recipients. That is because every recipient has different smtp server.
'''        DirectSend oSmtp, recipients
'''''        btnSend.Enabled = True
'''''        textStatus.Caption = ""
'''        Exit Sub
'''    End If
'''
    
    
    
    
    If oSmtp.SendMail() = 0 Then
'        Label16.Caption = "Nachricht erfolgreich versendet"
    Else
'        Label16.Caption = oSmtp.GetLastErrDescription()  'Get last error description
    End If
'    Label16.Refresh
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
   

    
'    With frmWKL00.sevSMTP1
'        .SenderName = gcBestellEmail.SenderName
'        .ReplyTo = gcBestellEmail.ReplyTo
'        .SenderEMail = gcBestellEmail.SenderEMail
'
'        gbCCfromBestlief = True
'        If gbCCfromBestlief = True Then
'            .CC = gcBestellEmail.CC
'        End If
'
'        If gcBestellEmail.BCC <> "" Then
'            .BCC = gcBestellEmail.BCC
'        End If
'
'        .Recipient = gcBestellEmail.Recipient
'        .SMTPAUTH = gcBestellEmail.SMTPAUTH
'        .ServerName = gcBestellEmail.ServerName
'        .ServerPort = gcBestellEmail.ServerPort
'        .Username = gcBestellEmail.Username
'        .Password = gcBestellEmail.Password
'        .Subject = gcBestellEmail.Subject
'        .Message = gcBestellEmail.Message
'        .AutoZIP = gcBestellEmail.AutoZIP
'
'
'        .AttachmentClear
'        If gcBestellEmail.Attachment1 <> "" Then
'            .AttachmentAdd gcBestellEmail.Attachment1
'        End If
'
'        If gcBestellEmail.Attachment2 <> "" Then
'            .AttachmentAdd gcBestellEmail.Attachment2
'        End If
'
'        If gcBestellEmail.Attachment3 <> "" Then
'            .AttachmentAdd gcBestellEmail.Attachment3
'        End If
'
'        If gcBestellEmail.Attachment4 <> "" Then
'            .AttachmentAdd gcBestellEmail.Attachment4
'        End If
'
'        If gcBestellEmail.Attachment5 <> "" Then
'            .AttachmentAdd gcBestellEmail.Attachment5
'        End If
'    End With
'
'    ' Anmeldung erfolgreich?
'    If frmWKL00.sevSMTP1.Connect() = True Then
'        Screen.MousePointer = 11
'        Dim lngBytesSent As Long
'        lngBytesSent = frmWKL00.sevSMTP1.SendMail()
'
'        frmWKL00.sevSMTP1.Disconnect
'    End If

    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "EmailverschickenSSL"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
'========================================================
' fnParseAddr
'========================================================
Public Function fnParseAddr(src, ByRef name, ByRef addr)
    Dim nIndex
    nIndex = InStr(1, src, "<")
    If nIndex > 0 Then
        name = Mid(src, 1, nIndex - 1)
        addr = Mid(src, nIndex)
    Else
        name = ""
        addr = src
    End If
    
    Call fnTrim(name, " ,;<>""'")
    Call fnTrim(addr, " ,;<>""'")
End Function
'========================================================
' fnTrim
'========================================================
Public Function fnTrim(ByRef src, trimer)
    Dim i, nCount, ch
    nCount = Len(src)
    For i = 1 To nCount
        ch = Mid(src, i, 1)
        If InStr(1, trimer, ch) < 1 Then
            Exit For
        End If
    Next
    
    src = Mid(src, i)
    nCount = Len(src)
    For i = nCount To 1 Step -1
        ch = Mid(src, i, 1)
        If InStr(1, trimer, ch) < 1 Then
            Exit For
        End If
    Next
    src = Mid(src, 1, i)
End Function
Public Function schickeMailimHintergrund(sSendername As String, sReplyTo As String _
, sCC As String, sAdresse As String, sABSENDER As String, sServerName As String, sServerport As String _
, sUser As String, sPass As String, sBetreff As String, sMsg As String, sAttachment1 As String) As Boolean
'    On Error GoTo LOKAL_ERROR
'
'    gcBestellEmail.SenderName = sSendername 'ermFirmenBez
'    gcBestellEmail.ReplyTo = sReplyTo ' ermFirmenMail
'    gcBestellEmail.SenderEMail = sABSENDER '"bestsend@kisswws.de"
'
'    gbCCfromBestlief = True
'    If gbCCfromBestlief = True Then
'        gcBestellEmail.CC = sCC 'ermFirmenMail
'    End If
'
'    gcBestellEmail.Recipient = sAdresse ' Combo1.Text
'    gcBestellEmail.SMTPAUTH = True
'    gcBestellEmail.ServerName = sServerName '"smtp.strato.de"
'    gcBestellEmail.ServerPort = sServerport '25
'    gcBestellEmail.Username = sUser '"bestsend@kisswws.de"
'    gcBestellEmail.Password = sPass '"geheim"
'
'    'Betreff
'    gcBestellEmail.Subject = sBetreff 'Text1(0).Text
'    gcBestellEmail.Message = sMsg 'Text1(1).Text
'
'    gcBestellEmail.AutoZIP = False
'    gcBestellEmail.Attachment1 = sAttachment1 '""
'    gcBestellEmail.Attachment2 = ""
'    gcBestellEmail.Attachment3 = ""
'    gcBestellEmail.Attachment4 = ""
'    gcBestellEmail.Attachment5 = ""
'
'    Emailverschicken
'
'Exit Function
'LOKAL_ERROR:
'    Fehler.gsDescr = err.Description
'    Fehler.gsNumber = err.Number
'    Fehler.gsFormular = "Modul2"
'    Fehler.gsFunktion = "schickeMailimHintergrund"
'    Fehler.gsFehlertext = "Es trat ein Fehler auf. "
'    Fehlermeldung1
End Function
Public Function schickeMailimHintergrundSSL(sSendername As String, sReplyTo As String _
, sCC As String, sAdresse As String, sABSENDER As String, sServerName As String, sServerport As String _
, sUser As String, sPass As String, sBetreff As String, sMsg As String, sAttachment1 As String) As Boolean
    On Error GoTo LOKAL_ERROR
    
    gcBestellEmail.SenderName = sSendername
    gcBestellEmail.ReplyTo = sReplyTo
    gcBestellEmail.SenderEMail = sABSENDER
    
    gbCCfromBestlief = True
    If gbCCfromBestlief = True Then
        gcBestellEmail.CC = sCC
    End If
    
    gcBestellEmail.Recipient = sAdresse
    gcBestellEmail.SMTPAUTH = True
    gcBestellEmail.ServerName = sServerName
    gcBestellEmail.ServerPort = sServerport
    gcBestellEmail.Username = sUser
    gcBestellEmail.Password = sPass
    gcBestellEmail.SSL = True
    
    'Betreff
    gcBestellEmail.Subject = sBetreff
    gcBestellEmail.Message = sMsg

    gcBestellEmail.AutoZIP = False
    gcBestellEmail.Attachment1 = sAttachment1
    gcBestellEmail.Attachment2 = ""
    gcBestellEmail.Attachment3 = ""
    gcBestellEmail.Attachment4 = ""
    gcBestellEmail.Attachment5 = ""
    gcBestellEmail.Attachment6 = ""
    
    EmailverschickenSSL

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "schickeMailimHintergrundSSL"
    Fehler.gsFehlertext = "Es trat ein Fehler auf. "
    Fehlermeldung1
End Function
Public Function LeseBankleitzahl(cKundnr As String) As String
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    
    LeseBankleitzahl = ""

    sSQL = "select * from  BANKKU where kundnr = " & cKundnr
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!BLZ) Then
            LeseBankleitzahl = rsrs!BLZ
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "LeseBankleitzahl"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Kontodaten auf. "
    Fehlermeldung1
End Function
Public Function ermKundenumsatz(bKundenbindung As Boolean, bymonat As Byte, lJahr As Long, iFil As Integer) As Double
    On Error GoTo LOKAL_ERROR

    Dim sSQL  As String
    Dim rsrs As Recordset
   
    ermKundenumsatz = 0
    
    sSQL = "Select sum(preis)as maxi from "
    If bKundenbindung = True Then
        sSQL = sSQL & " KUNZTmitK  "
    Else
        sSQL = sSQL & " KUNZTohneK "
    End If
    sSQL = sSQL & " where month(adate) = " & bymonat
    sSQL = sSQL & " and year(adate) = " & lJahr
    sSQL = sSQL & " and Filiale = " & iFil
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermKundenumsatz = rsrs!maxi
        End If
    End If
    rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ermKundenumsatz"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Public Function ermKundenMenge(bKundenbindung As Boolean, bymonat As Byte, lJahr As Long) As Long
    On Error GoTo LOKAL_ERROR

    Dim sSQL  As String
    Dim rsrs As Recordset
   
    ermKundenMenge = 0
    
    sSQL = "Select sum(menge)as maxi from "
    If bKundenbindung = True Then
        sSQL = sSQL & " KUNZTmitK  "
    Else
        sSQL = sSQL & " KUNZTohneK "
    End If
    sSQL = sSQL & " where month(adate) = " & bymonat
    sSQL = sSQL & " and year(adate) = " & lJahr
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermKundenMenge = rsrs!maxi
        End If
    End If
    rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "ermKundenMenge"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Public Function ermKundenumsatzproKauf(bKundenbindung As Boolean, bymonat As Byte, lJahr As Long, iFil As Integer) As Long
    On Error GoTo LOKAL_ERROR

    Dim sSQL  As String
    Dim rsrs As Recordset
   
    ermKundenumsatzproKauf = 0
    
    loeschNEW "KUNZTEMP", gdBase

    sSQL = "Select * into KUNZTEMP from    "
    If bKundenbindung = True Then
        sSQL = sSQL & " KUNZTmitK  "
    Else
        sSQL = sSQL & " KUNZTohneK "
    End If
    sSQL = sSQL & " where month(adate) = " & bymonat
    sSQL = sSQL & " and year(adate) = " & lJahr
    sSQL = sSQL & " and Filiale = " & iFil
    gdBase.Execute sSQL, dbFailOnError
    
    CheckIndex "KUNZTEMP", "Adate", "", gdBase
    CheckIndex "KUNZTEMP", "belegnr", "", gdBase
    
    
    sSQL = "Select distinct(adate) from KUNZTEMP"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!ADATE) Then
            
                
            
                'addieren und Schleife
                ermKundenumsatzproKauf = ermKundenumsatzproKauf + ermkunzTab(CLng(rsrs!ADATE), gdBase, "KUNZTEMP")
            End If
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "ermKundenumsatzproKauf"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Public Function KundenZahl_Bediener_Mon(imon As Integer, iJahr As Integer, lbed As Long) As Long
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL                As String
    Dim rsrs                As Recordset
    
    loeschNEW "KUANT", gdBase
    
    cSQL = "select  distinct adate, BELEGNR  "
    cSQL = cSQL & " into KUANT from KUANTE "
    
'    cSQL = cSQL & " Where Month(ADATE) = " & imon
'    cSQL = cSQL & " and year(ADATE) = " & iJahr
    
    cSQL = cSQL & " Where Bediener = " & lbed
    cSQL = cSQL & " group by adate,BELEGNR "
    gdBase.Execute cSQL, dbFailOnError
    
    KundenZahl_Bediener_Mon = 0
    cSQL = "select count(*) as maxi from KUANT"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            KundenZahl_Bediener_Mon = rsrs!maxi
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "KundenZahl_Bediener_Mon"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Public Function KundenZahl_Bediener(lDatVon As Long, lDatBis As Long, lbed As Long) As Long
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL                As String
    Dim rsrs                As Recordset
    
    loeschNEW "KUANT", gdBase
    
    cSQL = "select  distinct adate, BELEGNR  "
    cSQL = cSQL & " into KUANT from Kassjour where ADATE >= " & Trim$(Str$(lDatVon))
    cSQL = cSQL & " and ADATE <= " & Trim$(Str$(lDatBis))
    cSQL = cSQL & " and Bediener = " & lbed
    cSQL = cSQL & " group by adate,BELEGNR "
    gdBase.Execute cSQL, dbFailOnError
    
    KundenZahl_Bediener = 0
    cSQL = "select count(*) as maxi from KUANT"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            KundenZahl_Bediener = rsrs!maxi
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "KundenZahl_Bediener"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Public Function ermKundenumsatzproKaufalle() As Long
    On Error GoTo LOKAL_ERROR

    Dim sSQL  As String
    Dim rsrs As Recordset
   
    ermKundenumsatzproKaufalle = 0
    
    loeschNEW "KUNZTEMP", gdBase

    sSQL = "Select * into KUNZTEMP from kassjour "
    sSQL = sSQL & " where adate > " & CLng(DateValue(Now) - 180)
    sSQL = sSQL & " and ums_ok = 'J'"
    sSQL = sSQL & " and kundnr > 0"
    gdBase.Execute sSQL, dbFailOnError
    
    CheckIndex "KUNZTEMP", "Adate", "", gdBase
    CheckIndex "KUNZTEMP", "belegnr", "", gdBase
    
    sSQL = "Select distinct(adate) from KUNZTEMP"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!ADATE) Then
                'addieren und Schleife
                ermKundenumsatzproKaufalle = ermKundenumsatzproKaufalle + ermkunzTab(CLng(rsrs!ADATE), gdBase, "KUNZTEMP")
            End If
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "ermKundenumsatzproKaufalle"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermAktiveKundenzahlAusgesamtkassjour() As Double
    On Error GoTo LOKAL_ERROR

    Dim sSQL  As String
    Dim rsrs As Recordset
   
    ermAktiveKundenzahlAusgesamtkassjour = 0

    loeschNEW "KUNDSCHNITT", gdBase
    
    sSQL = "Select sum(preis) as mittel,kundnr into kundschnitt from kassjour "
    sSQL = sSQL & " where ums_ok = 'J'"
    sSQL = sSQL & " and KUNDNR > 0 "
    sSQL = sSQL & " group by KUNDNR "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Select avg(mittel) as durchsch from kundschnitt "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!durchsch) Then
                ermAktiveKundenzahlAusgesamtkassjour = rsrs!durchsch
            End If
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    loeschNEW "KUNDSCHNITT", gdBase

Exit Function
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "ermAktiveKundenzahlAusgesamtkassjour"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Public Function ermumsatzproZR() As Double
    On Error GoTo LOKAL_ERROR

    Dim sSQL  As String
    Dim rsrs As Recordset
   
    ermumsatzproZR = 0

    sSQL = "Select sum(preis) as maxi from kassjour"
    sSQL = sSQL & " where adate > " & CLng(DateValue(Now) - 180)
    sSQL = sSQL & " and ums_ok = 'J' "
    sSQL = sSQL & " and kundnr > 0 "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermumsatzproZR = rsrs!maxi
        End If
    End If
    rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "ermumsatzproZR"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Public Function ermkunzTab(lDat As Long, db As Database, sTab As String) As Long
On Error GoTo LOKAL_ERROR

    ermkunzTab = 0
    Dim rsrs        As Recordset
    Dim sSQL        As String

    loeschNEW "zeitzone", db

    sSQL = "select adate, belegnr into zeitzone  "

    sSQL = sSQL & " from " & sTab & " "
    sSQL = sSQL & "  where ADATE = " & lDat & " "
    db.Execute sSQL, dbFailOnError
    
    loeschNEW "zbonanz", db
    
    sSQL = "select distinct(belegnr) into zbonanz from zeitzone "
    sSQL = sSQL & " group by belegnr "
    db.Execute sSQL, dbFailOnError
    
    sSQL = "select * from zbonanz "
    Set rsrs = db.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        
        rsrs.MoveLast
        ermkunzTab = rsrs.RecordCount
        
    End If
    rsrs.Close: Set rsrs = Nothing
    
    loeschNEW "zbonanz", db
    loeschNEW "zeitzone", db
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ermkunzTab"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function

Public Function LeseKontonummer(cKundnr As String) As String
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    
    LeseKontonummer = ""

    sSQL = "select * from  BANKKU where kundnr = " & cKundnr
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!KTNR) Then
            LeseKontonummer = rsrs!KTNR
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "LeseKontonummer"
    Fehler.gsFehlertext = "Es trat ein Fehler im Programmteil Kontodaten auf. "
    Fehlermeldung1
End Function
Public Function GTBONHEUTENOCHNICHT() As Boolean
On Error GoTo LOKAL_ERROR

    Dim rsrs As Recordset

    GTBONHEUTENOCHNICHT = True
    
    If Not NewTableSuchenDBKombi("LGTBON", gdApp) Then
        CreateTable "LGTBON", gdApp
        Exit Function
    End If
    
    Set rsrs = gdApp.OpenRecordset("LGTBON")
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!lDate) Then
            If CLng(rsrs!lDate) = CLng(DateValue(Now)) Then
                GTBONHEUTENOCHNICHT = False
            End If
        End If
    
    End If
    rsrs.Close: Set rsrs = Nothing

   
    
    Screen.MousePointer = 0
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "GTBONHEUTENOCHNICHT"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function TagesabschlussGesternErfolgreich() As Double
On Error GoTo LOKAL_ERROR

    Dim rsrs As Recordset
    Dim cSQL As String

    TagesabschlussGesternErfolgreich = 0
    
    cSQL = "Select SUM(APREIS) as UMSATZ from AFCBUCH where kasnum  = " & gcKasNum
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!UMSATZ) Then
            TagesabschlussGesternErfolgreich = rsrs!UMSATZ
        End If
    End If
    rsrs.Close: Set rsrs = Nothing

    Screen.MousePointer = 0
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "TagesabschlussGesternErfolgreich"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Sub DruckenGutenTagBon()
    On Error GoTo LOKAL_ERROR
    Screen.MousePointer = 11
    
    Dim cDatum      As String
    Dim czeit       As String
    Dim cHH         As String
    Dim cZeile1     As String
    ReDim cZeilen(0 To 12) As String
    Dim sSQL        As String
    
    
    cDatum = DateValue(Now)
    czeit = TimeValue(Now)
    
    cHH = Left(czeit, 2)
    If Val(cHH) < 11 Then
        cZeile1 = "Guten Morgen!"
    ElseIf Val(cHH) < 18 Then
        cZeile1 = "Guten Tag!"
    Else
        cZeile1 = "Guten Abend!"
    End If
    
    'Drucke den Beleg
    
    
    cZeilen(0) = cZeile1
    cZeilen(1) = String$(glZeichenAnzahlBon, "-")
    
'    cZeilen(1) = "--------------------------------------"
    cZeilen(2) = "Ihre Filiale: " & gcFilNr & " startet heute: "
    cZeilen(3) = cDatum & " um: " & czeit & " Uhr"
    cZeilen(4) = "die Kasse."
    cZeilen(5) = ""
    If TagesabschlussGesternErfolgreich = 0 Then
        cZeilen(6) = "Der letzte Tagesabschluss."
        cZeilen(7) = "war erfolgreich."
    Else
        cZeilen(6) = "Es sind noch Umsätze"
        cZeilen(7) = "im Tagesabschluss."
    End If
    cZeilen(8) = ""
    cZeilen(9) = "Einen erfolgreichen Tag wünscht Ihnen"
    cZeilen(10) = "Ihr KISS Team aus Hannover"
    cZeilen(11) = "0511/955910"
    cZeilen(12) = "www.kisslive.de"
    
    DruckeArbeitszeitBelegWK20d cZeilen(), 12
    
    sSQL = "Delete from LGTBON "
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into LGTBON (LDATE) values ( '" & DateValue(Now) & "')"
    gdApp.Execute sSQL, dbFailOnError
    
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "DruckenGutenTagBon"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub HOLHELP(sPteil As String, txt As TextBox)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    
    txt.Text = ""
    sSQL = "Select HILFE from ZENTHELP where Pteil =  '" & sPteil & "'"
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!HILFE) Then
            If Not IsNull(rsrs!HILFE) Then
                txt.Text = rsrs!HILFE
            End If
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "HOLHELP"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Public Sub SAVEHELP(sPteil As String, txt As TextBox)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    txt.Text = SwapStr(txt.Text, "'", " ")
    
    sSQL = "Delete from ZENTHELP where Pteil =  '" & sPteil & "'"
    gdBase.Execute sSQL, dbFailOnError
    
    
    sSQL = "Insert into  ZENTHELP (Pteil,HILFE) "
    sSQL = sSQL & " values ('" & sPteil & "','" & txt.Text & "')"
    gdBase.Execute sSQL, dbFailOnError
    
    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "SAVEHELP"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Public Sub PrintHELP(sPteil As String, txt As TextBox)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    loeschNEW "PRINTHELP", gdBase
    CreateTable "PRINTHELP", gdBase

    sSQL = "Insert into  PRINTHELP select * from zenthelp where Pteil =  '" & sPteil & "'"
    gdBase.Execute sSQL, dbFailOnError
    
    reportbildschirm "", "aZENHE"
    
    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "PrintHELP"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Public Sub HOLARTIKEL(sArt As String, txt As TextBox)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    
    If sArt = "" Then
        Exit Sub
    End If
    
    txt.Text = ""
    sSQL = "Select HILFE from NOTART where ARTNR =  " & sArt
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!HILFE) Then
            If Not IsNull(rsrs!HILFE) Then
                txt.Text = rsrs!HILFE
            End If
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "HOLARTIKEL"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Public Sub SAVEARTIKEL(sArt As String, txt As TextBox, cBez As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    txt.Text = SwapStr(txt.Text, "'", " ")
    
    sSQL = "Delete from NOTART where ARTNR =  " & sArt
    gdBase.Execute sSQL, dbFailOnError
    
    
    sSQL = "Insert into  NOTART (ARTNR,HILFE,BEZEICH) "
    sSQL = sSQL & " values (" & sArt & ",'" & txt.Text & "','" & cBez & "')"
    gdBase.Execute sSQL, dbFailOnError
    
    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "SAVEARTIKEL"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Public Sub PrintARTIKEL(sArt As String, txt As TextBox)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    loeschNEW "PRINTART", gdBase
    CreateTable "PRINTART", gdBase

    sSQL = "Insert into  PRINTART select * from NOTART where artnr =  " & sArt
    gdBase.Execute sSQL, dbFailOnError
    
    reportbildschirm "", "aWKL111"
    
    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "PrintARTIKEL"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Public Function checkthisean(sEAN As String, sartn As String) As Boolean
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rsrs As Recordset
    
    If sEAN = "" Then
        checkthisean = True
        Exit Function
    End If
    
    If sartn = "" Then
        checkthisean = True
        Exit Function
    End If
    
    If Val(sEAN) = 0 Then
        checkthisean = True
        Exit Function
    End If
    
    If Len(sEAN) < "7" Then
        checkthisean = True
        Exit Function
    End If
    
    checkthisean = False
    
    sSQL = "Select * from artikel where ean = '" & sEAN & "'"
    sSQL = sSQL & " and artnr <> " & sartn
    sSQL = sSQL & " and (SYNSTATUS <> 'D' or SYNSTATUS is null )"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.Close: Set rsrs = Nothing
        Exit Function
    End If
    rsrs.Close: Set rsrs = Nothing
    
    sSQL = "Select * from artikel where ean2 = '" & sEAN & "'"
    sSQL = sSQL & " and artnr <> " & sartn
    sSQL = sSQL & " and (SYNSTATUS <> 'D' or SYNSTATUS is null )"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.Close: Set rsrs = Nothing
        Exit Function
    End If
    rsrs.Close: Set rsrs = Nothing
    
    sSQL = "Select * from artikel where ean3 = '" & sEAN & "'"
    sSQL = sSQL & " and artnr <> " & sartn
    sSQL = sSQL & " and (SYNSTATUS <> 'D' or SYNSTATUS is null )"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.Close: Set rsrs = Nothing
        Exit Function
    End If
    rsrs.Close: Set rsrs = Nothing
    
    sSQL = "Select * from ARTEAN_K where ean = '" & sEAN & "'"
    sSQL = sSQL & " and artnr <> " & sartn

    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.Close: Set rsrs = Nothing
        Exit Function
    End If
    rsrs.Close: Set rsrs = Nothing
    
    checkthisean = True

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "checkthisean"
    Fehler.gsFehlertext = "Beim Checken der EAN Nummer ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Sub Altinformation_speichern(sEAN As String, sartn As String)
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rsrs As Recordset
    
    If sEAN = "" Then
        Exit Sub
    End If
    
    If sartn = "" Then
        Exit Sub
    End If
    
    If Val(sEAN) = 0 Then
        Exit Sub
    End If
    
    If Len(sEAN) < "7" Then
        Exit Sub
    End If
    
    Dim lBestand As Long
    Dim dKVK As Double
    Dim lAltArtnr As Long
    
    lAltArtnr = 0
    
    If lAltArtnr = 0 Then
        sSQL = "Select * from artikel where ean = '" & sEAN & "'"
        sSQL = sSQL & " and artnr <> " & sartn
        sSQL = sSQL & " and (SYNSTATUS <> 'D' or SYNSTATUS is null )"
        Set rsrs = gdBase.OpenRecordset(sSQL)
        If Not rsrs.EOF Then
            If Not IsNull(rsrs!artnr) Then
                lAltArtnr = rsrs!artnr
            End If
        End If
        rsrs.Close: Set rsrs = Nothing
    End If
    
    
    
    If lAltArtnr = 0 Then
        sSQL = "Select * from artikel where ean2 = '" & sEAN & "'"
        sSQL = sSQL & " and artnr <> " & sartn
        sSQL = sSQL & " and (SYNSTATUS <> 'D' or SYNSTATUS is null )"
        Set rsrs = gdBase.OpenRecordset(sSQL)
        If Not rsrs.EOF Then
            If Not IsNull(rsrs!artnr) Then
                lAltArtnr = rsrs!artnr
            End If
        End If
        rsrs.Close: Set rsrs = Nothing
    End If
    
    If lAltArtnr = 0 Then
        sSQL = "Select * from artikel where ean3 = '" & sEAN & "'"
        sSQL = sSQL & " and artnr <> " & sartn
        sSQL = sSQL & " and (SYNSTATUS <> 'D' or SYNSTATUS is null )"
        Set rsrs = gdBase.OpenRecordset(sSQL)
        If Not rsrs.EOF Then
            If Not IsNull(rsrs!artnr) Then
                lAltArtnr = rsrs!artnr
            End If
        End If
        rsrs.Close: Set rsrs = Nothing
    End If
    
    
    If lAltArtnr = 0 Then
        sSQL = "Select * from artean_K where ean = '" & sEAN & "'"
        sSQL = sSQL & " and artnr <> " & sartn
'        sSQL = sSQL & " and (SYNSTATUS <> 'D' or SYNSTATUS is null )"
        Set rsrs = gdBase.OpenRecordset(sSQL)
        If Not rsrs.EOF Then
            If Not IsNull(rsrs!artnr) Then
                lAltArtnr = rsrs!artnr
            End If
        End If
        rsrs.Close: Set rsrs = Nothing
    End If
    
    If lAltArtnr > 0 Then
        sSQL = "Select Bestand,KVKPR1 from artikel where  artnr = " & lAltArtnr
        Set rsrs = gdBase.OpenRecordset(sSQL)
        If Not rsrs.EOF Then
            If Not IsNull(rsrs!BESTAND) Then
                lBestand = rsrs!BESTAND
            End If
            
            If Not IsNull(rsrs!KVKPR1) Then
                dKVK = rsrs!KVKPR1
            End If
        End If
        rsrs.Close: Set rsrs = Nothing
    End If
    
    sSQL = "Insert into Altinfo" & srechnertab & " (altartnr,EAN,neuArtnr,Bestand,KVKPR1) values ( "
    sSQL = sSQL & lAltArtnr
    sSQL = sSQL & ", '" & sEAN & "'"
    sSQL = sSQL & ", " & sartn
    sSQL = sSQL & ", " & lBestand
    sSQL = sSQL & ", '" & dKVK & "'"
    sSQL = sSQL & " ) "
    gdBase.Execute sSQL, dbFailOnError
    

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "Altinformation_speichern"
    Fehler.gsFehlertext = "Beim Checken der EAN Nummer ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
End Sub
Public Sub EAN_Updaten(lAltArtnr As Long)
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    
    sSQL = "Update artikel set ean = '',ean2='',ean3='' where artnr = " & lAltArtnr & ""
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete * from artean_K where artnr = " & lAltArtnr & ""
    gdBase.Execute sSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "EAN_Updaten"
    Fehler.gsFehlertext = "Beim Checken der EAN Nummer ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Function checkthiseankoml(sartn As String) As Boolean
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rsrs As Recordset
    
    checkthiseankoml = False
    
    sSQL = "Select * from eankoml33 where artnr = " & sartn
    sSQL = sSQL & " and farbe = '3'"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.Close: Set rsrs = Nothing
        Exit Function
    End If
    rsrs.Close: Set rsrs = Nothing
    
    checkthiseankoml = True

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "checkthiseankoml"
    Fehler.gsFehlertext = "Beim Checken der EAN Nummer ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Sub FaerbenGrid(grid As MSFlexGrid, iawmSpalte As Integer, Izufarbspalte As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim i       As Integer
    Dim j       As Integer
    Dim cAWM    As String
    
    With grid
        .Redraw = False
    
        For i = 0 To .Rows - 1
            .Row = i
            For j = 0 To .Cols - 1
            .Col = j
                If .Col = iawmSpalte Then
                    cAWM = .TextMatrix(i, j)
                    If cAWM = "" Then cAWM = "0"
                    FaerbenFlex cAWM, grid, Izufarbspalte, i
                End If
                
            Next j
        Next i
        .Redraw = True
    End With
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "FaerbenGrid"
    Fehler.gsFehlertext = "Beim Faerben eines Grids ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Public Sub FaerbenGrid_INBEST(grid As MSFlexGrid, iINBESTSpalte As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim i       As Integer
    Dim j       As Integer
    Dim cINBEST    As String
    
    With grid
        .Redraw = False
    
        For i = 1 To .Rows - 1
            .Row = i
            For j = 0 To .Cols - 1
            .Col = j
                If .Col = iINBESTSpalte Then
                    cINBEST = .TextMatrix(i, j)
                    If cINBEST = "" Then cINBEST = "0"
                    
                    If CInt(cINBEST) > 0 Then
                        .CellBackColor = &HFF00FF
                    Else
                        .CellBackColor = vbWhite
                    End If
                    
                    
                    
                    
                End If
                
            Next j
        Next i
        .Redraw = True
    End With
    
    
   
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "FaerbenGrid_INBEST"
    Fehler.gsFehlertext = "Beim Faerben eines Grids ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Public Sub FaerbenHGrid(grid As MSHFlexGrid, iawmSpalte As Integer, Izufarbspalte As Integer)
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer
    Dim j As Integer
    
    Dim cAWM                As String
    
    With grid
        .Redraw = False
    
        For i = 0 To .Rows - 1
            .Row = i
            For j = 0 To .Cols - 1
            .Col = j
                If .Col = iawmSpalte Then
                    cAWM = .TextMatrix(i, j)
                    If cAWM = "" Then cAWM = "0"
                    FaerbenFlexH cAWM, grid, Izufarbspalte, i
                End If
                
            Next j
        Next i
        .Redraw = True
    End With
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "FaerbenHGrid"
    Fehler.gsFehlertext = "Beim Faerben eines Grids ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Function ermdopio(cART As String) As Integer
On Error GoTo LOKAL_ERROR

Dim rsrs As Recordset
Dim sSQL As String

If Not IsNumeric(cART) Then
    Exit Function
End If

sSQL = "Select * from artlief where artnr = " & cART
sSQL = sSQL & " and  ( SYNSTATUS is null or SYNSTATUS = 'E' or SYNSTATUS = 'A' ) "

Set rsrs = gdBase.OpenRecordset(sSQL)
If Not rsrs.EOF Then
    rsrs.MoveLast
    ermdopio = rsrs.RecordCount
Else
    ermdopio = 0
End If
rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ermdopio"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function FaerbenGridaufGrundmehrLINR(grid As MSFlexGrid, iLINRSpalte As Integer, Iartspalte As Integer) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer
    Dim j As Integer
    FaerbenGridaufGrundmehrLINR = False
    
    Dim cArtNr                As String
    
    With grid
        .Redraw = False
    
        For i = 0 To .Rows - 1
            .Row = i
            .Col = Iartspalte
            cArtNr = .Text
            
            If ermdopio(cArtNr) > 1 Then
                FaerbenGridaufGrundmehrLINR = True
                .Col = iLINRSpalte
                
                .CellBackColor = vbRed
                .CellForeColor = vbBlack
                
            End If
                
               
            
            
        Next i
        .Redraw = True
    End With
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "FaerbenGridaufGrundmehrLINR"
    Fehler.gsFehlertext = "Beim Faerben eines Grids ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Sub FnOpenrecordset(glrs As Recordset, rsSQL As String, iArt As Integer, db As Database)
On Error GoTo LOKAL_ERROR

Set glrs = Nothing

If db.name = gdBase.name Then

    Select Case iArt
    Case 1 'select
        Set glrs = db.OpenRecordset(rsSQL)
    Case 2 'Addnew,Edit = update
        schreibeProtokollDabaAblauf rsSQL
        Set glrs = db.OpenRecordset(rsSQL, dbOpenDynaset, dbDenyWrite + dbSeeChanges, dbPessimistic)
    Case 3 'delete
        schreibeProtokollDabaAblauf rsSQL
        Set glrs = db.OpenRecordset(rsSQL, dbOpenDynaset, dbDenyWrite + dbSeeChanges, dbPessimistic)
    
    End Select
End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "FnOpenrecordset"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1

End Sub
Public Sub AfcstatIstNull(sTab As String)
On Error GoTo LOKAL_ERROR

Dim rsrs As Recordset
Dim cSQL As String

    cSQL = "Select * from " & sTab & " where KASNUM = " & gcKasNum & " "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF

            rsrs.Edit
            
            If Not IsNull(rsrs!UMS_BAR) Then
                
            Else
                rsrs!UMS_BAR = 0
            End If
            
            If Not IsNull(rsrs!UMS_Kred) Then
                
            Else
                rsrs!UMS_Kred = 0
            End If
            
            If Not IsNull(rsrs!UMS_SCHECK) Then
                
            Else
                rsrs!UMS_SCHECK = 0
            End If
            
            If Not IsNull(rsrs!UMS_KARTE) Then
                
            Else
                rsrs!UMS_KARTE = 0
            End If
            
            
            
            If Not IsNull(rsrs!UMS_LAST) Then
                
            Else
                rsrs!UMS_LAST = 0
            End If
    
            If Not IsNull(rsrs!SPREIS_ANZ) Then
               
            Else
                rsrs!SPREIS_ANZ = 0
            End If
            
            If Not IsNull(rsrs!SPREIS_GES) Then
               
            Else
                rsrs!SPREIS_GES = 0
            End If
            
            If Not IsNull(rsrs!ANZSCHECK) Then
               
            Else
                rsrs!ANZSCHECK = 0
            End If
            
            If Not IsNull(rsrs!Kundenzahl) Then
                
            Else
                rsrs!Kundenzahl = 0
            End If
            
            If Not IsNull(rsrs!GELDFACH) Then
               
            Else
                rsrs!GELDFACH = 0
            End If
                
            If Not IsNull(rsrs!ARTRAB_ANZ) Then
               
            Else
                rsrs!ARTRAB_ANZ = 0
            End If
            
            If Not IsNull(rsrs!ARTRAB_GES) Then
              
            Else
                rsrs!ARTRAB_GES = 0
            End If
            
            If Not IsNull(rsrs!GESRAB_ANZ) Then
               
            Else
                rsrs!GESRAB_ANZ = 0
            End If
            
            If Not IsNull(rsrs!GESRAB_GES) Then
                
            Else
                rsrs!GESRAB_GES = 0
            End If
            
            If Not IsNull(rsrs!STORNO_ANZ) Then
                
            Else
                rsrs!STORNO_ANZ = 0
            End If
    
            If Not IsNull(rsrs!STORNO_GES) Then
                
            Else
                rsrs!STORNO_GES = 0
            End If
            
            If Not IsNull(rsrs!EINZAHLUNG) Then
                
            Else
                rsrs!EINZAHLUNG = 0
            End If
            
            If Not IsNull(rsrs!AUSZAHLUNG) Then
                
            Else
                rsrs!AUSZAHLUNG = 0
            End If
            
            If Not IsNull(rsrs!GUTSCHEIN) Then
              
            Else
                rsrs!GUTSCHEIN = 0
            End If
            
            If Not IsNull(rsrs!ZHLGGUTSCH) Then
              
            Else
                rsrs!ZHLGGUTSCH = 0
            End If
            
            If Not IsNull(rsrs!BELEGNR) Then
                
            Else
                rsrs!BELEGNR = 0
            End If
            
            If Not IsNull(rsrs!GUTSCHBAR) Then
               
            Else
                rsrs!GUTSCHBAR = 0
            End If
            
            If Not IsNull(rsrs!GUTSCHSCH) Then
               
            Else
                rsrs!GUTSCHSCH = 0
            End If
            
            If Not IsNull(rsrs!GUTSCHKRE) Then
               
            Else
                rsrs!GUTSCHKRE = 0
            End If
            
            If Not IsNull(rsrs!GUTSCHKAR) Then
                
            Else
                rsrs!GUTSCHKAR = 0
            End If
            
            If Not IsNull(rsrs!GUTSCHLAST) Then
               
            Else
                rsrs!GUTSCHLAST = 0
            End If
            
    
            If Not IsNull(rsrs!BARVERKAUF) Then
                
            Else
                rsrs!BARVERKAUF = 0
            End If
            
            If Not IsNull(rsrs!SCHVERKAUF) Then
                
            Else
                rsrs!SCHVERKAUF = 0
            End If
            
            If Not IsNull(rsrs!TILGBAR) Then
               
            Else
                rsrs!TILGBAR = 0
            End If
            
            If Not IsNull(rsrs!TILGSCH) Then
                
            Else
                rsrs!TILGSCH = 0
            End If
            
            If Not IsNull(rsrs!TILGGUT) Then
                
            Else
                rsrs!TILGGUT = 0
            End If
            
            If Not IsNull(rsrs!TILGKAR) Then
               
            Else
                rsrs!TILGKAR = 0
            End If

            If Not IsNull(rsrs!EINRGUTSCH) Then
                
            Else
                rsrs!EINRGUTSCH = 0
            End If
            
            If Not IsNull(rsrs!RESTGUTSCH) Then
               
            Else
                rsrs!RESTGUTSCH = 0
            End If
            
            If Not IsNull(rsrs!AUSZGUTSCH) Then
                
            Else
                rsrs!AUSZGUTSCH = 0
            End If
            
            If Not IsNull(rsrs!Wechsel) Then

            Else
                rsrs!Wechsel = 0
            End If
            
            If Not IsNull(rsrs!NUMSKARTE) Then

            Else
                rsrs!NUMSKARTE = 0
            End If
            
            If Not IsNull(rsrs!DUKA) Then

            Else
                rsrs!DUKA = 0
            End If
            

            rsrs.Update
            
            rsrs.MoveNext
            
            
        Loop
    End If

    rsrs.Close: Set rsrs = Nothing


Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "AfcstatIstNull"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Public Sub LogtoStart(Frm As Form)
    On Error GoTo LOKAL_ERROR
    Dim formularUeberschrift  As String
    Dim formname As String
    
    formname = Frm.name
    
    formularUeberschrift = SwapStr(Frm.Caption, "'", " ")
    
'    Formularstart wegschreiben

    schreibeProtokollProgrammablauf " betritt  " & formname & " " & formularUeberschrift

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "LogtoStart"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Public Sub LogtoEnd(Frm As Form)
    On Error GoTo LOKAL_ERROR
    Dim formularUeberschrift  As String
    Dim formname As String
    
    formname = Frm.name
    formularUeberschrift = SwapStr(Frm.Caption, "'", " ")
    
'    Formularende wegschreiben

    schreibeProtokollProgrammablauf " verlässt " & formname & " " & formularUeberschrift

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "LogtoEnd"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Public Sub BestBedProvision(txtStatus As TextBox, picprogress As PictureBox)
    On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11

    Dim sSQL As String
    
    txtStatus.Text = 0
    picprogress.Visible = True
    
    txtStatus.Text = 14
    
    loeschNEW "BAT", gdBase
    CreateTable "BAT", gdBase
    
    txtStatus.Text = 55
    
    sSQL = "Insert into BAT select artnr ,bezeich, bediener, Preis,Menge,adate,azeit,mwst "
    sSQL = sSQL & " from Provision "
    
    If Month(DateValue(Now)) - 1 = 0 Then
        sSQL = sSQL & " where month(adate) = 12 "
    Else
        sSQL = sSQL & " where month(adate) = " & Month(DateValue(Now)) - 1
    End If
    
    If Month(DateValue(Now)) - 1 = 0 Then
        sSQL = sSQL & " and  year(adate) = " & Year(DateValue(Now)) - 1
    Else
        sSQL = sSQL & " and  year(adate) = " & Year(DateValue(Now))
    End If

    gdBase.Execute sSQL, dbFailOnError
    txtStatus.Text = 64
    
    sSQL = "Update BAT inner join artikel on bat.artnr = artikel.artnr set BAT.farbnr = val(artikel.awm)"
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 70

    sSQL = "Update BAT inner join Bedname on BAT.bediener = Bedname.BEDNU "
    sSQL = sSQL & " SET BAT.BEDNAME = BEDNAME.bedname "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update BAT  SET Preis = Preis/100"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update BAT set NPreis = (PREIS * 100)/(100 + " & gdMWStV & ") "
    sSQL = sSQL & " where mwst = 'V' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update BAT set NPreis = (PREIS * 100)/(100 + " & gdMWStE & ")  "
    sSQL = sSQL & " where mwst = 'E' "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update BAT set NPreis = (PREIS * 100)/(100 + " & gdMWStO & ")  "
    sSQL = sSQL & " where mwst = 'O' "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 86
    
    If Month(DateValue(Now)) - 1 = 0 Then
        sSQL = "Update BAT  SET mont = 'Dezember'"
        gdBase.Execute sSQL, dbFailOnError
    Else
        sSQL = "Update BAT  SET mont = '" & MonthName(Month(DateValue(Now)) - 1) & "'"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    txtStatus.Text = 89
    
    BringFarbeInsSpiel "BAT", gdBase
    
    reportbildschirm "", "aZEN0V3"
    
    txtStatus.Text = 0
    picprogress.Visible = False

    Screen.MousePointer = 0
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "BestBedProvision"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Public Sub BestBedProvisionRab(txtStatus As TextBox, picprogress As PictureBox)
    On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11

    Dim sSQL As String
    
    txtStatus.Text = 0
    picprogress.Visible = True
    
    
    txtStatus.Text = 14
    
    loeschNEW "AAT", gdBase
    
    txtStatus.Text = 55
    
    sSQL = "Select artnr ,bezeich, bediener,'' as bedname,'' as rabkenn,'' as mont, Preis,Menge,adate,azeit into AAT "
    sSQL = sSQL & " from Kassjour "
    
    If Month(DateValue(Now)) - 1 = 0 Then
        sSQL = sSQL & " where month(adate) = 12 "
    Else
        sSQL = sSQL & " where month(adate) = " & Month(DateValue(Now)) - 1
    End If
    
    If Month(DateValue(Now)) - 1 = 0 Then
        sSQL = sSQL & " and  year(adate) = " & Year(DateValue(Now)) - 1
    Else
        sSQL = sSQL & " and  year(adate) = " & Year(DateValue(Now))
    End If

    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError

    txtStatus.Text = 70

    sSQL = "Update AAT inner join Bedname on AAT.bediener = Bedname.BEDNU "
    sSQL = sSQL & " SET AAT.BEDNAME = BEDNAME.bedname "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update AAT inner join Artikel on AAT.artnr = ARTIKEL.artnr "
    sSQL = sSQL & " SET AAT.rabkenn = Artikel.rabatt_ok "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from AAT where rabkenn = '' "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from AAT where rabkenn is null "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from AAT where rabkenn = 'N' "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 86
    
    If Month(DateValue(Now)) - 1 = 0 Then
        sSQL = "Update AAT  SET mont = 'Dezember'"
        schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    Else
        sSQL = "Update AAT  SET mont = '" & MonthName(Month(DateValue(Now)) - 1) & "'"
        schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    End If
    
    txtStatus.Text = 89
    
    reportbildschirm "", "aZEN0V4"
    
    txtStatus.Text = 0
    picprogress.Visible = False

    Screen.MousePointer = 0
    

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "BestBedProvisionRab"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
  
End Sub
Public Function erminBestell(cArtNr As String) As Long
On Error GoTo LOKAL_ERROR

    Dim cSQL    As String
    Dim rsrs1   As Recordset

    erminBestell = 0
    
    cSQL = "Select SUM(BESTVOR) as INBEST from BESTREST where ARTNR = " & cArtNr & " "
    Set rsrs1 = gdBase.OpenRecordset(cSQL)
    If Not rsrs1.EOF Then
        rsrs1.MoveFirst
        
        If Not IsNull(rsrs1!INBEST) Then
            erminBestell = rsrs1!INBEST
        End If
        
    End If
    
    rsrs1.Close: Set rsrs1 = Nothing

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "erminBestell"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function

Public Sub BestBedKuCut1(txtStatus As TextBox, picprogress As PictureBox)
    On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11

    Dim sSQL As String
    
    txtStatus.Text = 0
    picprogress.Visible = True

    txtStatus.Text = 10
    
    loeschNEW "TOPBED", gdBase
    loeschNEW "TOPBED1", gdBase
    txtStatus.Text = 12
    loeschNEW "TOPBEDP2", gdBase
    txtStatus.Text = 13
    CreateTable "TOPBEDP2", gdBase
    
    
    txtStatus.Text = 14
    
    loeschNEW "AAT", gdBase
    
    txtStatus.Text = 15
    
    sSQL = "Select distinct adate, BELEGNR as ANZKUNDEN , bediener, sum(Preis) as APreis into AAT "
    sSQL = sSQL & " from Kassjour "
    
    If Month(DateValue(Now)) - 1 = 0 Then
        sSQL = sSQL & " where month(adate) = 12 "
    Else
        sSQL = sSQL & " where month(adate) = " & Month(DateValue(Now)) - 1
    End If
    
    If Month(DateValue(Now)) - 1 = 0 Then
        sSQL = sSQL & " and  year(adate) = " & Year(DateValue(Now)) - 1
    Else
        sSQL = sSQL & " and  year(adate) = " & Year(DateValue(Now))
    End If
    
    sSQL = sSQL & " and UMS_OK = 'J' "
    sSQL = sSQL & " group by adate,bediener,BELEGNR"

    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 30
    
    sSQL = "Select sum(APreis) as SPreis ,BEDIENER,count(ANZKUNDEN) as belegnr into TopBED1 "
    sSQL = sSQL & " from AAT "
    sSQL = sSQL & " group by bediener"
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 42
    
    sSQL = "Insert into TOPBEDP2 Select SPreis ,BEDIENER, belegnr as bonanz from TopBED1 "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 54
    
    sSQL = "Update TopBEDP2 set KUCUT = sPreis/bonanz where bonanz <> 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 67

    loeschNEW "TOPP4", gdBase
    
    txtStatus.Text = 68
    CreateTable "TOPP4", gdBase
    
    txtStatus.Text = 70
    
    sSQL = "Insert into TOPP4 SELECT  Bediener, KUCUT ,sPreis as tPreis, bonanz as tanzku "
    sSQL = sSQL & " from TopBEDP2 order by KUCUT desc"
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 78

    sSQL = "Update TOPP4 inner join Bedname on TOPP4.bediener = Bedname.BEDNU "
    sSQL = sSQL & " SET TOPP4.BEDNAME = BEDNAME.bedname "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 86
    
    If Month(DateValue(Now)) - 1 = 0 Then
        sSQL = "Update TOPP4  SET mont = 'Dezember'"
        gdBase.Execute sSQL, dbFailOnError
    Else
        sSQL = "Update TOPP4  SET mont = '" & MonthName(Month(DateValue(Now)) - 1) & "'"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    reportbildschirm "", "aZEN0V1"
    
    txtStatus.Text = 0
    picprogress.Visible = False

    Screen.MousePointer = 0
    
    
    
    loeschNEW "TOPP4", gdBase
    loeschNEW "TopBEDP2", gdBase
    loeschNEW "AAT", gdBase
    loeschNEW "TopBED1", gdBase

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "BestBedKuCut1"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
   
End Sub
Public Sub BestBedKuCut(txtStatus As TextBox, picprogress As PictureBox)
    On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11

    Dim sSQL As String
    
    txtStatus.Text = 0
    picprogress.Visible = True

    txtStatus.Text = 10
    
    
    loeschNEW "TOPBED", gdBase
    txtStatus.Text = 12
    loeschNEW "TOPBED1", gdBase
    loeschNEW "TOPBED2", gdBase
    txtStatus.Text = 13
    CreateTable "TOPBED2", gdBase
    
    
    txtStatus.Text = 14
    
    loeschNEW "AAT", gdBase
    
    txtStatus.Text = 15
    
    sSQL = "Select distinct adate, BELEGNR as ANZKUNDEN , bediener, sum(Menge) as AMenge into AAT "
    sSQL = sSQL & " from Kassjour "
    
    If Month(DateValue(Now)) - 1 = 0 Then
        sSQL = sSQL & " where month(adate) = 12 "
    Else
        sSQL = sSQL & " where month(adate) = " & Month(DateValue(Now)) - 1
    End If
    
    If Month(DateValue(Now)) - 1 = 0 Then
        sSQL = sSQL & " and  year(adate) = " & Year(DateValue(Now)) - 1
    Else
        sSQL = sSQL & " and  year(adate) = " & Year(DateValue(Now))
    End If
    
    sSQL = sSQL & " and UMS_OK = 'J' "
    sSQL = sSQL & " group by adate,bediener,BELEGNR"

    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 30
    
    sSQL = "Select sum(AMenge) as SMENGE ,BEDIENER,count(ANZKUNDEN) as belegnr into TopBED1 "
    sSQL = sSQL & " from AAT "
    sSQL = sSQL & " group by bediener"
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 42
    
    sSQL = "Insert into TOPBED2 Select SMENGE ,BEDIENER, belegnr as bonanz from TopBED1 "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 54
    
    sSQL = "Update TopBED2 set KUCUT = sMenge/bonanz where bonanz <> 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 67

    loeschNEW "TOP4", gdBase
    
    txtStatus.Text = 68
    CreateTable "TOP4", gdBase
    
    txtStatus.Text = 70
    
    sSQL = "Insert into TOP4 SELECT  Bediener, KUCUT ,smenge as tMenge, bonanz as tanzku "
    sSQL = sSQL & " from TopBED2 order by KUCUT desc"
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 78

    sSQL = "Update TOP4 inner join Bedname on TOP4.bediener = Bedname.BEDNU "
    sSQL = sSQL & " SET TOP4.BEDNAME = BEDNAME.bedname "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 86
    
    If Month(DateValue(Now)) - 1 = 0 Then
        sSQL = "Update TOP4  SET mont = 'Dezember'"
        gdBase.Execute sSQL, dbFailOnError
    Else
        sSQL = "Update TOP4  SET mont = '" & MonthName(Month(DateValue(Now)) - 1) & "'"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    

    reportbildschirm "", "aZEN00V"
    
    txtStatus.Text = 0
    picprogress.Visible = False

    Screen.MousePointer = 0
    
    
    
    loeschNEW "TOP4", gdBase
    loeschNEW "TopBED2", gdBase
    loeschNEW "AAT", gdBase
    loeschNEW "TopBED1", gdBase

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "BestBedKuCut"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
   
End Sub
Public Sub BestBedKuCut2(txtStatus As TextBox, picprogress As PictureBox)
    On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11

    Dim sSQL As String
    
    txtStatus.Text = 0
    picprogress.Visible = True

    txtStatus.Text = 10
    
    
    loeschNEW "TOPBED", gdBase
    loeschNEW "TOPBED1", gdBase
    txtStatus.Text = 12
    loeschNEW "TOPBEDP3", gdBase
    txtStatus.Text = 13
    CreateTable "TOPBEDP3", gdBase
    
    
    txtStatus.Text = 14
    
    loeschNEW "AAT1", gdBase
    
    
    
    sSQL = "Select  adate, BELEGNR  , bediener, 0.0 as Ertrag, mwst , EKPR as LEKPR, Menge as VKMENGE , Preis as VKPREIS into AAT1 "
    sSQL = sSQL & " from Kassjour "
    
    If Month(DateValue(Now)) - 1 = 0 Then
        sSQL = sSQL & " where month(adate) = 12 "
    Else
        sSQL = sSQL & " where month(adate) = " & Month(DateValue(Now)) - 1
    End If
    
    If Month(DateValue(Now)) - 1 = 0 Then
        sSQL = sSQL & " and  year(adate) = " & Year(DateValue(Now)) - 1
    Else
        sSQL = sSQL & " and  year(adate) = " & Year(DateValue(Now))
    End If
    
    sSQL = sSQL & " and UMS_OK = 'J' "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    
    
    txtStatus.Text = 16
    
    
    sSQL = "Update AAT1 set ertrag = ((VKPREIS * 100)/(100 + " & gdMWStV & ")) - (LEKPR * VKMENGE) "
    sSQL = sSQL & " where mwst = 'V' "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 19
    
    sSQL = "Update AAT1 set ertrag = ((VKPREIS * 100)/(100 + " & gdMWStE & ")) - (LEKPR * VKMENGE) "
    sSQL = sSQL & " where mwst = 'E' "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 22
    
    sSQL = "Update AAT1 set ertrag = ((VKPREIS * 100)/(100 + " & gdMWStO & " )) - (LEKPR * VKMENGE) "
    sSQL = sSQL & " where mwst = 'O' "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    
    
    
    
    txtStatus.Text = 25
    
    loeschNEW "AAT", gdBase
    
    sSQL = "Select distinct adate, BELEGNR as ANZKUNDEN , bediener, sum(Ertrag) as AErtrag into AAT "
    sSQL = sSQL & " from aat1 "
    sSQL = sSQL & " group by adate,bediener,BELEGNR"
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    
    txtStatus.Text = 30
    
    sSQL = "Select sum(AErtrag) as SErtrag ,BEDIENER,count(ANZKUNDEN) as belegnr into TopBED1 "
    sSQL = sSQL & " from AAT "
    sSQL = sSQL & " group by bediener"
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 42
    
    sSQL = "Insert into TOPBEDP3 Select SErtrag ,BEDIENER, belegnr as bonanz from TopBED1 "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 54
    
    sSQL = "Update TopBEDP3 set KUCUT = sErtrag/bonanz where bonanz <> 0"
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 67

    loeschNEW "TOPP6", gdBase
    
    txtStatus.Text = 68
    CreateTable "TOPP6", gdBase
    
    txtStatus.Text = 70
    
    sSQL = "Insert into TOPP6 SELECT  Bediener, KUCUT ,sErtrag as tErtrag, bonanz as tanzku "
    sSQL = sSQL & " from TopBEDP3 order by KUCUT desc"
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 78

    sSQL = "Update TOPP6 inner join Bedname on TOPP6.bediener = Bedname.BEDNU "
    sSQL = sSQL & " SET TOPP6.BEDNAME = BEDNAME.bedname "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 86
    
    If Month(DateValue(Now)) - 1 = 0 Then
        sSQL = "Update TOPP6  SET mont = 'Dezember'"
        gdBase.Execute sSQL, dbFailOnError
    Else
        sSQL = "Update TOPP6  SET mont = '" & MonthName(Month(DateValue(Now)) - 1) & "'"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    

    reportbildschirm "", "aZEN0V2"
    
    txtStatus.Text = 0
    picprogress.Visible = False

    Screen.MousePointer = 0
    
    
    
    loeschNEW "TOPP6", gdBase
    loeschNEW "TopBEDP3", gdBase
    loeschNEW "AAT", gdBase
    loeschNEW "AAT1", gdBase
    loeschNEW "TopBED1", gdBase

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "BestBedKuCut2"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
   
End Sub
Public Sub BestBedKuCutDEVELo(txtStatus As TextBox, picprogress As PictureBox)
    On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11

    Dim sSQL    As String
    Dim i       As Integer
    
    txtStatus.Text = 0
    picprogress.Visible = True

    txtStatus.Text = 10



    For i = 1 To 6
    
        

        loeschNEW "AAT" & i, gdBase
        sSQL = "Select distinct adate, BELEGNR as ANZKUNDEN , bediener , sum(Menge) as AMenge into AAT" & i
        sSQL = sSQL & " from Kassjour "
        
        If Month(DateValue(Now)) - i = 0 Then
            sSQL = sSQL & " where month(adate) = 12 "
            
        ElseIf Month(DateValue(Now)) - i = -1 Then
            sSQL = sSQL & " where month(adate) = 11 "
            
        ElseIf Month(DateValue(Now)) - i = -2 Then
            sSQL = sSQL & " where month(adate) = 10 "
            
        ElseIf Month(DateValue(Now)) - i = -3 Then
            sSQL = sSQL & " where month(adate) = 9 "
        ElseIf Month(DateValue(Now)) - i = -4 Then
            sSQL = sSQL & " where month(adate) = 8 "
        ElseIf Month(DateValue(Now)) - i = -5 Then
            sSQL = sSQL & " where month(adate) = 7 "
        ElseIf Month(DateValue(Now)) - i = -6 Then
            sSQL = sSQL & " where month(adate) = 6 "
        Else
            sSQL = sSQL & " where month(adate) = " & Month(DateValue(Now)) - i
        End If
        
        If Month(DateValue(Now)) - i = 0 Then
            sSQL = sSQL & " and year(adate) = " & Year(DateValue(Now)) - 1
        ElseIf Month(DateValue(Now)) - i = -1 Then
            sSQL = sSQL & " and year(adate) = " & Year(DateValue(Now)) - 1
        ElseIf Month(DateValue(Now)) - i = -2 Then
            sSQL = sSQL & " and year(adate) = " & Year(DateValue(Now)) - 1
        ElseIf Month(DateValue(Now)) - i = -3 Then
            sSQL = sSQL & " and year(adate) = " & Year(DateValue(Now)) - 1
        ElseIf Month(DateValue(Now)) - i = -4 Then
            sSQL = sSQL & " and year(adate) = " & Year(DateValue(Now)) - 1
        ElseIf Month(DateValue(Now)) - i = -5 Then
            sSQL = sSQL & " and year(adate) = " & Year(DateValue(Now)) - 1
        ElseIf Month(DateValue(Now)) - i = -6 Then
            sSQL = sSQL & " and year(adate) = " & Year(DateValue(Now)) - 1
        Else
            sSQL = sSQL & " and year(adate) = " & Year(DateValue(Now))
        End If
        
        sSQL = sSQL & " and UMS_OK = 'J' "
        sSQL = sSQL & " group by adate,bediener,BELEGNR"
        
        gdBase.Execute sSQL, dbFailOnError

        txtStatus.Text = 15 + i

    Next i

    For i = 1 To 6

        loeschNEW "TOPBEDT" & i, gdBase

        sSQL = "Select sum(AMenge) as SMENGE ,BEDIENER,count(ANZKUNDEN) as belegnr into TOPBEDT" & i
        sSQL = sSQL & " from AAT" & i
        sSQL = sSQL & " group by bediener"
        gdBase.Execute sSQL, dbFailOnError

        txtStatus.Text = 22 + i
    Next i


    For i = 1 To 6
        loeschNEW "TOPBED" & i, gdBase
        
        sSQL = "Create Table TOPBED" & i
        sSQL = sSQL & "( BEDIENER INTEGER"
        sSQL = sSQL & ", SMENGE LONG"
        sSQL = sSQL & ", BONANZ LONG"
        sSQL = sSQL & ", KUCUT single"
        sSQL = sSQL & ") "
        gdBase.Execute sSQL, dbFailOnError

        sSQL = "Insert into TOPBED" & i & " Select SMENGE ,BEDIENER, belegnr as bonanz  from TOPBEDT" & i
        gdBase.Execute sSQL, dbFailOnError

        txtStatus.Text = 29 + i

        sSQL = "Update TopBED" & i & " set KUCUT = sMenge/bonanz where bonanz <> 0"
        gdBase.Execute sSQL, dbFailOnError
    Next i

    loeschNEW "TOP5", gdBase

    txtStatus.Text = 68
    CreateTable "TOP5", gdBase

    txtStatus.Text = 69

    sSQL = "Insert into TOP5 SELECT  Bediener, KUCUT as KUCUT1 ,smenge as tMenge1, bonanz as tanzku1 "
    sSQL = sSQL & " from TopBED1 order by KUCUT desc"
    gdBase.Execute sSQL, dbFailOnError
    
    For i = 2 To 6
    
        txtStatus.Text = 69 + i
    
        sSQL = "Update TOP5 inner join TopBED" & i & " on TOP5.bediener = TopBED" & i & ".bediener "
        sSQL = sSQL & " SET TOP5.KUCUT" & i & " = TopBED" & i & ".KUCUT "
        sSQL = sSQL & " , TOP5.tMenge" & i & " = TopBED" & i & ".sMenge "
        sSQL = sSQL & " , TOP5.tanzku" & i & " = TopBED" & i & ".bonanz "
        gdBase.Execute sSQL, dbFailOnError
    
    Next i
    
    txtStatus.Text = 76

    sSQL = "Update TOP5 inner join Bedname on TOP5.bediener = Bedname.BEDNU "
    sSQL = sSQL & " SET TOP5.BEDNAME = BEDNAME.bedname "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 79
    
    sSQL = "Update TOP5  SET KUCUTS = (KUCUT1 + KUCUT2 + KUCUT3 + KUCUT4 + KUCUT5 + KUCUT6)/6 "
    gdBase.Execute sSQL, dbFailOnError
    
    Dim cMontName
    
    For i = 1 To 6
        txtStatus.Text = 79 + i
        
        If Month(DateValue(Now)) - i = 0 Then
            cMontName = "Dezember"
        ElseIf Month(DateValue(Now)) - i = -1 Then
            cMontName = "November"
        ElseIf Month(DateValue(Now)) - i = -2 Then
            cMontName = "Oktober"
        ElseIf Month(DateValue(Now)) - i = -3 Then
            cMontName = "September"
        ElseIf Month(DateValue(Now)) - i = -4 Then
            cMontName = "August"
        ElseIf Month(DateValue(Now)) - i = -5 Then
            cMontName = "Juli"
        ElseIf Month(DateValue(Now)) - i = -6 Then
            cMontName = "Juni"
        Else
            cMontName = MonthName(Month(DateValue(Now)) - i)
        End If
        
        
        sSQL = "Update TOP5  SET mont" & i & " = '" & cMontName & "'"
        gdBase.Execute sSQL, dbFailOnError
    
    Next i
    
    loeschNEW "TOP6", gdBase
    sSQL = "select * into top6 from TOP5 order by KUCUTS desc"
    gdBase.Execute sSQL, dbFailOnError

    reportbildschirm "", "aZEN00B"
    
    txtStatus.Text = 0
    picprogress.Visible = False

    Screen.MousePointer = 0
    
    
    
    loeschNEW "TOP5", gdBase

    For i = 1 To 6
        loeschNEW "ATT" & i, gdBase
        loeschNEW "TOPBED" & i, gdBase
        loeschNEW "TOPBEDT" & i, gdBase
    Next i

    
    

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "BestBedKuCutDEVELo"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
    
'    Resume Next
   
End Sub
Public Sub schreibeProtokoll_Artikel_VKPREISE(sZeile As String)
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr     As Integer
    Dim lPos        As Long
    Dim lWert       As Long
    Dim cSatz       As String
    Dim cdatei      As String
    Dim cPfad       As String
    Dim sTime       As String
    Dim ctmp        As String
    Dim cZeil       As String
    Dim cZeile2     As String
    Dim sRechner    As String

    sRechner = rechnername
    sRechner = SwapStr(sRechner, ".", "")
    
    
    sTime = TimeValue(Now)
    sTime = Right(sTime, 8)
    
    cPfad = gcDBPfad    'Datenbankpfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "LPROTOK\"

    lWert = DateValue(Now)
    ctmp = Format$(lWert, "DD.MM.YY")
    
    cdatei = "Artikel_VKPREISE"
    
    cZeil = ""
    cZeil = ctmp & Space(1) & sTime & Space(2) & sRechner & sZeile & vbCrLf
    
    
    cSatz = cPfad & cdatei & ".TXT"
    
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
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "schreibeProtokoll_Artikel_VKPREISE"
        Fehler.gsFehlertext = "Beim Erstellen des Fehler Protokolls ist ein Fehler aufgetreten."

        Fehlermeldung1
       
    End If
End Sub
Public Sub schreibeProtokoll_Artikel_EX(sZeile As String)
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr     As Integer
    Dim lPos        As Long
    Dim lWert       As Long
    Dim cSatz       As String
    Dim cdatei      As String
    Dim cPfad       As String
    Dim sTime       As String
    Dim ctmp        As String
    Dim cZeil       As String
    Dim cZeile2     As String
    Dim sRechner    As String

    sRechner = rechnername
    sRechner = SwapStr(sRechner, ".", "")
    
    sTime = TimeValue(Now)
    sTime = Right(sTime, 8)
    
    cPfad = gcDBPfad    'Datenbankpfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "LPROTOK\"

    lWert = DateValue(Now)
    ctmp = Format$(lWert, "DD.MM.YY")
    
    cdatei = "Artikel_EX"
    
    cZeil = ""
    cZeil = ctmp & Space(1) & sTime & Space(2) & sRechner & sZeile & vbCrLf
    
    
    cSatz = cPfad & cdatei & ".TXT"
    
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
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "schreibeProtokoll_Artikel_EX"
        Fehler.gsFehlertext = "Beim Erstellen des Fehler Protokolls ist ein Fehler aufgetreten."

        Fehlermeldung1
       
    End If
End Sub
Public Sub schreibeProtokollProgrammablauf(sZeile As String)
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr     As Integer
    Dim lPos        As Long
    Dim lWert       As Long
    Dim cSatz       As String
    Dim cdatei      As String
    Dim cPfad       As String
    Dim sTime       As String
    Dim ctmp        As String
    Dim cZeil       As String
    Dim cZeile2     As String
    Dim sRechner    As String

    sRechner = rechnername
    sRechner = SwapStr(sRechner, ".", "")
    
    
    sTime = TimeValue(Now)
    sTime = Right(sTime, 8)
    
    cPfad = gcDBPfad    'Datenbankpfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "LPROTOK\"

    lWert = DateValue(Now)
    ctmp = Format$(lWert, "DD.MM.YY")
    
    cdatei = "PROABL"
    
    cZeil = ""
    cZeil = ctmp & Space(1) & sTime & Space(2) & sRechner & sZeile & vbCrLf
    
    
    cSatz = cPfad & cdatei & ".TXT"
    
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
    
'        Fehler.gsDescr = err.Description
'        Fehler.gsNumber = err.Number
'        Fehler.gsFormular = "Modul2"
'        Fehler.gsFunktion = "schreibeProtokollProgrammablauf"
'        Fehler.gsFehlertext = "Beim Erstellen des Fehler Protokolls ist ein Fehler aufgetreten."
'
'        Fehlermeldung1
        Resume Next
    End If
End Sub
Public Sub schreibeProtokollBESTablauf(sZeile As String)
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr     As Integer
    Dim lPos        As Long
    Dim lWert       As Long
    Dim cSatz       As String
    Dim cdatei      As String
    Dim cPfad       As String
    Dim sTime       As String
    Dim ctmp        As String
    Dim cZeil       As String
    Dim cZeile2     As String
    Dim sRechner    As String

    sRechner = rechnername
    sRechner = SwapStr(sRechner, ".", "")
    
    
    sTime = TimeValue(Now)
    sTime = Right(sTime, 8)
    
    cPfad = gcDBPfad    'Datenbankpfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "LPROTOK\"

    lWert = DateValue(Now)
    ctmp = Format$(lWert, "DD.MM.YY")
    
    cdatei = "BEST"
    
    cZeil = ""
    cZeil = ctmp & Space(1) & sTime & Space(2) & sRechner & sZeile & vbCrLf
    
    
    cSatz = cPfad & cdatei & ".TXT"
    
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
    
'        Fehler.gsDescr = err.Description
'        Fehler.gsNumber = err.Number
'        Fehler.gsFormular = "Modul2"
'        Fehler.gsFunktion = "schreibeProtokollProgrammablauf"
'        Fehler.gsFehlertext = "Beim Erstellen des Fehler Protokolls ist ein Fehler aufgetreten."
'
'        Fehlermeldung1
        Resume Next
    End If
End Sub
Public Sub schreibeProtokollKVKPR1ablauf(sZeile As String)
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr     As Integer
    Dim lPos        As Long
    Dim lWert       As Long
    Dim cSatz       As String
    Dim cdatei      As String
    Dim cPfad       As String
    Dim sTime       As String
    Dim ctmp        As String
    Dim cZeil       As String
    Dim cZeile2     As String
    Dim sRechner    As String

    sRechner = rechnername
    sRechner = SwapStr(sRechner, ".", "")
    
    
    sTime = TimeValue(Now)
    sTime = Right(sTime, 8)
    
    cPfad = gcDBPfad    'Datenbankpfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "LPROTOK\"

    lWert = DateValue(Now)
    ctmp = Format$(lWert, "DD.MM.YY")
    
    cdatei = "KVKPR1"
    
    cZeil = ""
    cZeil = ctmp & Space(1) & sTime & Space(2) & sRechner & sZeile & vbCrLf
    
    
    cSatz = cPfad & cdatei & ".TXT"
    
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
    
'        Fehler.gsDescr = err.Description
'        Fehler.gsNumber = err.Number
'        Fehler.gsFormular = "Modul2"
'        Fehler.gsFunktion = "schreibeProtokollProgrammablauf"
'        Fehler.gsFehlertext = "Beim Erstellen des Fehler Protokolls ist ein Fehler aufgetreten."
'
'        Fehlermeldung1
        Resume Next
    End If
End Sub
Public Sub schreibeProtokollAWMablauf(sZeile As String)
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr     As Integer
    Dim lPos        As Long
    Dim lWert       As Long
    Dim cSatz       As String
    Dim cdatei      As String
    Dim cPfad       As String
    Dim sTime       As String
    Dim ctmp        As String
    Dim cZeil       As String
    Dim cZeile2     As String
    Dim sRechner    As String

    sRechner = rechnername
    sRechner = SwapStr(sRechner, ".", "")
    
    
    sTime = TimeValue(Now)
    sTime = Right(sTime, 8)
    
    cPfad = gcDBPfad    'Datenbankpfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "LPROTOK\"

    lWert = DateValue(Now)
    ctmp = Format$(lWert, "DD.MM.YY")
    
    cdatei = "Farbe"
    
    cZeil = ""
    cZeil = ctmp & Space(1) & sTime & Space(2) & sRechner & sZeile & vbCrLf
    
    
    cSatz = cPfad & cdatei & ".TXT"
    
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
    
'        Fehler.gsDescr = err.Description
'        Fehler.gsNumber = err.Number
'        Fehler.gsFormular = "Modul2"
'        Fehler.gsFunktion = "schreibeProtokollProgrammablauf"
'        Fehler.gsFehlertext = "Beim Erstellen des Fehler Protokolls ist ein Fehler aufgetreten."
'
'        Fehlermeldung1
        Resume Next
    End If
End Sub
Public Sub schreibeProtokollEANablauf(sZeile As String)
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr     As Integer
    Dim lPos        As Long
    Dim lWert       As Long
    Dim cSatz       As String
    Dim cdatei      As String
    Dim cPfad       As String
    Dim sTime       As String
    Dim ctmp        As String
    Dim cZeil       As String
    Dim cZeile2     As String
    Dim sRechner    As String

    sRechner = rechnername
    sRechner = SwapStr(sRechner, ".", "")
    
    
    sTime = TimeValue(Now)
    sTime = Right(sTime, 8)
    
    cPfad = gcDBPfad    'Datenbankpfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "LPROTOK\"

    lWert = DateValue(Now)
    ctmp = Format$(lWert, "DD.MM.YY")
    
    cdatei = "EAN"
    
    cZeil = ""
    cZeil = ctmp & Space(1) & sTime & Space(2) & sRechner & sZeile & vbCrLf
    
    
    cSatz = cPfad & cdatei & ".TXT"
    
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
    
'        Fehler.gsDescr = err.Description
'        Fehler.gsNumber = err.Number
'        Fehler.gsFormular = "Modul2"
'        Fehler.gsFunktion = "schreibeProtokollProgrammablauf"
'        Fehler.gsFehlertext = "Beim Erstellen des Fehler Protokolls ist ein Fehler aufgetreten."
'
'        Fehlermeldung1
        Resume Next
    End If
End Sub
Public Sub schreibeProtokollArtikelMengenLoeschen(sZeile As String)
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr     As Integer
    Dim lPos        As Long
    Dim lWert       As Long
    Dim cSatz       As String
    Dim cdatei      As String
    Dim cPfad       As String
    Dim sTime       As String
    Dim ctmp        As String
    Dim cZeil       As String
    Dim cZeile2     As String
    Dim sRechner    As String

    sRechner = rechnername
    sRechner = SwapStr(sRechner, ".", "")
    
    
    sTime = TimeValue(Now)
    sTime = Right(sTime, 8)
    
    cPfad = gcDBPfad    'Datenbankpfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "LPROTOK\"

    lWert = DateValue(Now)
    ctmp = Format$(lWert, "DD.MM.YY")
    
    cdatei = "ArtMenDEL"
    
    cZeil = ""
    cZeil = ctmp & Space(1) & sTime & Space(2) & sRechner & sZeile & vbCrLf
    
    
    cSatz = cPfad & cdatei & ".TXT"
    
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
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "schreibeProtokollArtikelMengenLoeschen"
        Fehler.gsFehlertext = "Beim Erstellen des Fehler Protokolls ist ein Fehler aufgetreten."

        Fehlermeldung1
       
    End If
End Sub
Public Sub schreibeProtokollInventurImport(sZeile As String)
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr     As Integer
    Dim lPos        As Long
    Dim lWert       As Long
    Dim cSatz       As String
    Dim cdatei      As String
    Dim cPfad       As String
    Dim sTime       As String
    Dim ctmp        As String
    Dim cZeil       As String
    Dim cZeile2     As String
    Dim sRechner    As String

    sRechner = rechnername
    sRechner = SwapStr(sRechner, ".", "")
    
    
    sTime = TimeValue(Now)
    sTime = Right(sTime, 8)
    
    cPfad = gcDBPfad    'Datenbankpfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "LPROTOK\"

    lWert = DateValue(Now)
    ctmp = Format$(lWert, "DD.MM.YY")
    
    cdatei = "Inventur_Import"
    
    cZeil = ""
    cZeil = ctmp & Space(1) & sTime & Space(2) & sRechner & sZeile & vbCrLf
    
    
    cSatz = cPfad & cdatei & ".TXT"
    
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
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "schreibeProtokollInventurImport"
        Fehler.gsFehlertext = "Beim Erstellen des Fehler Protokolls ist ein Fehler aufgetreten."

        Fehlermeldung1
       
    End If
End Sub
Public Sub schreibeProtokollgKUN(sZeile As String)
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr     As Integer
    Dim lPos        As Long
    Dim lWert       As Long
    Dim cSatz       As String
    Dim cdatei      As String
    Dim cPfad       As String
    Dim sTime       As String
    Dim ctmp        As String
    Dim cZeil       As String
    
    
    
    sTime = TimeValue(Now)
    sTime = Right$(sTime, 8)
    
    cPfad = gcDBPfad    'Datenbankpfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "LPROTOK\"

    lWert = DateValue(Now)
    ctmp = Format$(lWert, "DD.MM.YY")
    
    cdatei = "geloeschteKunden"
    
    
    cZeil = ""
    cSatz = cPfad & cdatei & ".RTF"
    iFileNr = FreeFile
    Open cSatz For Binary As #iFileNr
    cZeil = ctmp & Space(1) & sTime & Space(1) & sZeile
    lPos = LOF(iFileNr)
    lPos = lPos + 1
    Put #iFileNr, lPos, cZeil & vbCrLf
    
    
    Close iFileNr
    
  Exit Sub
LOKAL_ERROR:
    
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "schreibeProtokollgKUN"
        Fehler.gsFehlertext = "Beim Erstellen des allgemeinen Protokolls ist ein Fehler aufgetreten."
        
        Fehlermeldung1
End Sub

Public Sub schreibeProtokollBENUTZERablauf(sZeile As String)
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr     As Integer
    Dim lPos        As Long
    Dim lWert       As Long
    Dim cSatz       As String
    Dim cdatei      As String
    Dim cPfad       As String
    Dim sTime       As String
    Dim ctmp        As String
    Dim cZeil       As String
    Dim cZeile2     As String
    Dim sRechner    As String

    sRechner = rechnername
    sRechner = SwapStr(sRechner, ".", "")
    
    
    sTime = TimeValue(Now)
    sTime = Right(sTime, 8)
    
    cPfad = gcDBPfad    'Datenbankpfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "LPROTOK\"

    lWert = DateValue(Now)
    ctmp = Format$(lWert, "DD.MM.YY")
    
    cdatei = "BENUTZER"
    
    cZeil = ""
    cZeil = ctmp & Space(1) & sTime & Space(2)
    cZeil = cZeil & Space(30 - Len(cZeil)) & sRechner
    cZeil = cZeil & Space(50 - Len(cZeil)) & sZeile & vbCrLf
    
    cSatz = cPfad & cdatei & ".TXT"
    
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
    
'        Fehler.gsDescr = err.Description
'        Fehler.gsNumber = err.Number
'        Fehler.gsFormular = "Modul2"
'        Fehler.gsFunktion = "schreibeProtokollProgrammablauf"
'        Fehler.gsFehlertext = "Beim Erstellen des Fehler Protokolls ist ein Fehler aufgetreten."
'
'        Fehlermeldung1
        Resume Next
    End If
End Sub
Public Sub schreibelokalFehlerproto(sZeile As String)
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr     As Integer
    Dim lPos        As Long
    Dim lWert       As Long
    Dim cSatz       As String
    Dim cdatei      As String
    Dim cPfad       As String
    Dim sTime       As String
    Dim ctmp        As String
    Dim cZeil       As String
    Dim cZeile2     As String
    Dim sRechner    As String

    sRechner = rechnername
    sRechner = SwapStr(sRechner, ".", "")
    
    
    sTime = TimeValue(Now)
    sTime = Right(sTime, 8)
    
    cPfad = gcPfad    'apppfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "BIGERR\"

    lWert = DateValue(Now)
    ctmp = Format$(lWert, "DD.MM.YY")
    
    cdatei = "RED60"
    
    cZeil = ""
    cZeil = ctmp & Space(1) & sTime & Space(2) & sRechner & Space(1) & sZeile & vbCrLf
    
    
    cSatz = cPfad & cdatei & ".TXT"
    
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
    
'        Fehler.gsDescr = err.Description
'        Fehler.gsNumber = err.Number
'        Fehler.gsFormular = "Modul2"
'        Fehler.gsFunktion = "schreibeProtokollProgrammablauf"
'        Fehler.gsFehlertext = "Beim Erstellen des Fehler Protokolls ist ein Fehler aufgetreten."
'
'        Fehlermeldung1
        Resume Next
    End If
End Sub
Public Sub schreibeProtokollKundenkorrektur(sZeile As String)
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr     As Integer
    Dim lPos        As Long
    Dim lWert       As Long
    Dim cSatz       As String
    Dim cdatei      As String
    Dim cPfad       As String
    Dim sTime       As String
    Dim ctmp        As String
    Dim cZeile1     As String
    Dim cZeile2     As String
    Dim sRechner    As String

    sRechner = rechnername

    sTime = TimeValue(Now)
    sTime = Right(sTime, 8)
    
    cPfad = gcDBPfad    'Datenbankpfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "LPROTOK\"

    lWert = DateValue(Now)
    ctmp = Format$(lWert, "DD.MM.YY")
    
    cdatei = "Kundenkontokorrektur"
    
    cZeile1 = ""
    cZeile1 = ctmp & Space(1) & sTime & Space(2) & sRechner & Space(1) & sZeile & vbCrLf
    
    
    cSatz = cPfad & cdatei & ".TXT"
    
 
    
    iFileNr = FreeFile
    Open cSatz For Binary As #iFileNr

        lPos = LOF(iFileNr)
        lPos = lPos + 1
        Put #iFileNr, lPos, cZeile1

    Close iFileNr
    
  Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
    
        Resume Next
    Else
    
'        Fehler.gsDescr = err.Description
'        Fehler.gsNumber = err.Number
'        Fehler.gsFormular = "Modul2"
'        Fehler.gsFunktion = "schreibeProtokollNachtAblauf"
'        Fehler.gsFehlertext = "Beim Erstellen des Fehler Protokolls ist ein Fehler aufgetreten."
'
'        Fehlermeldung1
        Resume Next
    End If
End Sub
Public Sub schreibeProtokollStamda(sZeile As String)
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr     As Integer
    Dim lPos        As Long
    Dim lWert       As Long
    Dim cSatz       As String
    Dim cdatei      As String
    Dim cPfad       As String
    Dim sTime       As String
    Dim ctmp        As String
    Dim cZeil       As String
    Dim cZeile2     As String
    
    
    sTime = TimeValue(Now)
    sTime = Right$(sTime, 8)
    
    cPfad = gcDBPfad    'Datenbankpfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "LPROTOK\"

    lWert = DateValue(Now)
    ctmp = Format$(lWert, "DD.MM.YY")
    
    cdatei = "Stamda"
    
    cZeil = ""
    cZeil = ctmp & Space(1) & sTime & Space(2) & sZeile & vbCrLf
    cSatz = cPfad & cdatei & ".TXT"
    
    iFileNr = FreeFile
    Open cSatz For Binary As #iFileNr
    If LOF(iFileNr) > 0 Then
        cZeile2 = Space$(LOF(iFileNr))
        Get #iFileNr, 1, cZeile2
        Close iFileNr
    Else
        Close iFileNr
        Kill cSatz
        
    End If
    
    Kill cSatz
    
    iFileNr = FreeFile
    Open cSatz For Binary As #iFileNr
    
        lPos = LOF(iFileNr)
        lPos = lPos + 1
        Put #iFileNr, lPos, cZeil
    
        lPos = LOF(iFileNr)
        lPos = lPos + 1
        
        Put #iFileNr, lPos, cZeile2
        
    Close iFileNr
    
    
    
  Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
    
        Resume Next
    Else
    
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "schreibeProtokollStamda"
        Fehler.gsFehlertext = "Beim Erstellen des allgemeinen Protokolls ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Public Sub schreibeProtokoll_Bargeld_Handling(sZeile As String)
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr     As Integer
    Dim lPos        As Long
    Dim lWert       As Long
    Dim cSatz       As String
    Dim cdatei      As String
    Dim cPfad       As String
    Dim sTime       As String
    Dim ctmp        As String
    Dim cZeile1     As String
    
    sTime = TimeValue(Now)
    sTime = Right$(sTime, 8)
    
    cPfad = gcDBPfad    'Datenbankpfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "LPROTOK\"

    lWert = DateValue(Now)
    ctmp = Format$(lWert, "DD.MM.YY")
    
    cdatei = "Bargeld_Handling"
    
    cZeile1 = ""
    cZeile1 = ctmp & Space(1) & sTime & Space(2) & sZeile & vbCrLf
    cSatz = cPfad & cdatei & ".TXT"
    
    iFileNr = FreeFile
    Open cSatz For Binary As #iFileNr

        lPos = LOF(iFileNr)
        lPos = lPos + 1
        Put #iFileNr, lPos, cZeile1

    Close iFileNr
    
    
    
  Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "schreibeProtokoll_Bargeld_Handling"
        Fehler.gsFehlertext = "Beim Erstellen des allgemeinen Protokolls ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Public Sub schreibeProtokollDabaAblauf(sZeile As String)
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr     As Integer
    Dim lPos        As Long
    Dim lWert       As Long
    Dim cSatz       As String
    Dim cdatei      As String
    Dim cPfad       As String
    Dim sTime       As String
    Dim ctmp        As String
    Dim cZeile1     As String
    Dim cZeile2     As String
    Dim sRechner    As String

    sRechner = rechnername

    sTime = TimeValue(Now)
    sTime = Right(sTime, 8)
    
    cPfad = gcDBPfad    'Datenbankpfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "LPROTOK\"

    lWert = DateValue(Now)
    ctmp = Format$(lWert, "DD.MM.YY")
    
    cdatei = "DABAABL"
    
    cZeile1 = ""
    cZeile1 = ctmp & Space(1) & sTime & Space(2) & sRechner & Space(1) & sZeile & vbCrLf
    
    
    cSatz = cPfad & cdatei & ".TXT"
    
 
    
    iFileNr = FreeFile
    Open cSatz For Binary As #iFileNr

        lPos = LOF(iFileNr)
        lPos = lPos + 1
        Put #iFileNr, lPos, cZeile1

    Close iFileNr
    
  Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
    
        Resume Next
    Else
    
'        Fehler.gsDescr = err.Description
'        Fehler.gsNumber = err.Number
'        Fehler.gsFormular = "Modul2"
'        Fehler.gsFunktion = "schreibeProtokollProgrammablauf"
'        Fehler.gsFehlertext = "Beim Erstellen des Fehler Protokolls ist ein Fehler aufgetreten."
'
'        Fehlermeldung1
        Resume Next
    End If
End Sub
Public Sub schreibeTSE(sZeile As String)
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr     As Integer
    Dim lPos        As Long
    Dim cSatz       As String
    Dim cdatei      As String
    Dim cPfad       As String
    
    cPfad = gcPfad    'app Pfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "TSE\"
    
    cdatei = CStr(GetUnixTimestamp)
    
    cSatz = cPfad & cdatei & ".TXT"
    
    iFileNr = FreeFile
    Open cSatz For Binary As #iFileNr

        lPos = LOF(iFileNr)
        lPos = lPos + 1
        Put #iFileNr, lPos, sZeile

    Close iFileNr
    
  Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
    
        Resume Next
    Else
    
'        Fehler.gsDescr = err.Description
'        Fehler.gsNumber = err.Number
'        Fehler.gsFormular = "Modul2"
'        Fehler.gsFunktion = "schreibeTSE"
'        Fehler.gsFehlertext = "Beim Erstellen des Fehler Protokolls ist ein Fehler aufgetreten."
'
'        Fehlermeldung1
        Resume Next
    End If
End Sub

Private Function GetUnixTimestamp() As Long
  ' Zeit-Differenz zum 01.01.1970 00:00:00 berechnen
  Dim nSek As Long
  nSek = DateDiff("s", CDate("01.01.1970 00:00:00"), Now)
 
  ' jetzt noch Sommern/Winterzeit berücksichtigen
  Dim nDiff As Long
  Dim st As SYSTEMTIME
 
  ' Systemzeit ermitteln
  GetSystemTime st
 
  ' Zeit-Differenz zur GMT-Zeit in Sekunden
  nDiff = DateDiff("s", DateSerial(st.wYear, st.wMonth, st.wDay) + _
    TimeSerial(st.wHour, st.wMinute, st.wSecond), Now)
  nSek = nSek - nDiff
 
  GetUnixTimestamp = nSek
End Function
Public Sub schreibeProtokollGZwarn(sZeile As String)
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr     As Integer
    Dim lPos        As Long
    Dim lWert       As Long
    Dim cSatz       As String
    Dim cdatei      As String
    Dim cPfad       As String
    Dim sTime       As String
    Dim ctmp        As String
    Dim cZeile1     As String
    Dim cZeile2     As String
    Dim sRechner    As String

    sRechner = rechnername

    sTime = TimeValue(Now)
    sTime = Right(sTime, 8)
    
    cPfad = gcDBPfad    'Datenbankpfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "LPROTOK\"

    lWert = DateValue(Now)
    ctmp = Format$(lWert, "DD.MM.YY")
    
    cdatei = "GZWARN"
    
    cZeile1 = ""
    cZeile1 = ctmp & Space(1) & sTime & Space(2) & sRechner & Space(1) & sZeile & vbCrLf
    
    
    cSatz = cPfad & cdatei & ".TXT"
    
 
    
    iFileNr = FreeFile
    Open cSatz For Binary As #iFileNr

        lPos = LOF(iFileNr)
        lPos = lPos + 1
        Put #iFileNr, lPos, cZeile1

    Close iFileNr
    
  Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
    
        Resume Next
    Else
    
'        Fehler.gsDescr = err.Description
'        Fehler.gsNumber = err.Number
'        Fehler.gsFormular = "Modul2"
'        Fehler.gsFunktion = "schreibeProtokollGZwarn"
'        Fehler.gsFehlertext = "Beim Erstellen des Fehler Protokolls ist ein Fehler aufgetreten."
'
'        Fehlermeldung1
        Resume Next
    End If
End Sub
Public Sub ZentraleWillsWissen(sZeile As String)
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr     As Integer
    Dim lPos        As Long
    Dim lWert       As Long
    Dim cSatz       As String
    Dim cdatei      As String
    Dim cPfad       As String
    Dim sTime       As String
    Dim ctmp        As String
    Dim cZeil       As String
    Dim cZeile1     As String
    Dim cZeile2     As String
    Dim sRechner    As String
    Dim cName       As String
    
    lWert = DateValue(Now)
    ctmp = Format$(lWert, "DD.MM")
    ctmp = SwapStr(ctmp, ".", "")
    cName = gcFilNr & "F" & ctmp
    
    If gcFilNr <= 1 Then
        Exit Sub
    End If

    sTime = TimeValue(Now)
    sTime = Right(sTime, 8)
    
    cPfad = gcDBPfad    'Datenbankpfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "ZPROTOK\"

    lWert = DateValue(Now)
    ctmp = Format$(lWert, "DD.MM.YY")
    
    cdatei = cName
    
    cZeil = ""
    cZeil = ctmp & Space(1) & sTime & Space(2) & sZeile & vbCrLf
    
    
    cSatz = cPfad & cdatei & ".TXT"
    
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
    
'        Fehler.gsDescr = err.Description
'        Fehler.gsNumber = err.Number
'        Fehler.gsFormular = "Modul2"
'        Fehler.gsFunktion = "ZentraleWillsWissen"
'        Fehler.gsFehlertext = "Beim Erstellen des Fehler Protokolls ist ein Fehler aufgetreten."
'
'        Fehlermeldung1
'
    End If
End Sub
Public Sub schreibeProtokollIndex(sZeile As String)
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr     As Integer
    Dim lPos        As Long
    Dim lWert       As Long
    Dim cSatz       As String
    Dim cdatei      As String
    Dim cPfad       As String
    Dim sTime       As String
    Dim ctmp        As String
    Dim cZeile1     As String
    Dim cZeile2     As String
    Dim sRechner    As String

    sRechner = rechnername

    sTime = TimeValue(Now)
    sTime = Right(sTime, 8)
    
    cPfad = gcDBPfad    'Datenbankpfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "LPROTOK\"

    lWert = DateValue(Now)
    ctmp = Format$(lWert, "DD.MM.YY")
    
    cdatei = "INDEX"
    
    cZeile1 = ""
    cZeile1 = ctmp & Space(1) & sTime & Space(2) & sRechner & Space(1) & sZeile & vbCrLf
    
    
    cSatz = cPfad & cdatei & ".TXT"
    
 
    
    iFileNr = FreeFile
    Open cSatz For Binary As #iFileNr

        lPos = LOF(iFileNr)
        lPos = lPos + 1
        Put #iFileNr, lPos, cZeile1

    Close iFileNr
    
  Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
    
        Resume Next
    Else
    
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "schreibeProtokollIndex"
        Fehler.gsFehlertext = "Beim Erstellen des Fehler Protokolls ist ein Fehler aufgetreten."

        Fehlermeldung1
        
    End If
End Sub
Public Sub schreibeProtokollKassenFehler(sZeile As String)
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr     As Integer
    Dim lPos        As Long
    Dim lWert       As Long
    Dim cSatz       As String
    Dim cdatei      As String
    Dim cPfad       As String
    Dim sTime       As String
    Dim ctmp        As String
    Dim cZeile1     As String
    Dim cZeile2     As String
    Dim sRechner    As String

    sRechner = rechnername

    sTime = TimeValue(Now)
    sTime = Right(sTime, 8)
    
    DabaPfadNew84
    
    cPfad = gcDBPfad    'Datenbankpfad
    
    DabaPfadNew83
    
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "PROTOK\"

    lWert = DateValue(Now)
    ctmp = Format$(lWert, "DD.MM.YY")
    
    cdatei = "FKASSE"
    
    cZeile1 = ""
    cZeile1 = ctmp & Space(1) & sTime & Space(2) & sRechner & Space(1) & sZeile & vbCrLf
    
    
    cSatz = cPfad & cdatei & ".TXT"
    
 
    
    iFileNr = FreeFile
    Open cSatz For Binary As #iFileNr

        lPos = LOF(iFileNr)
        lPos = lPos + 1
        Put #iFileNr, lPos, cZeile1

    Close iFileNr
    
  Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
    
        Resume Next
    Else
    
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "schreibeProtokollKassenFehler"
        Fehler.gsFehlertext = "Beim Erstellen des Fehler Protokolls ist ein Fehler aufgetreten."

        Fehlermeldung1
        
    End If
End Sub
Public Function ermERTRAG(sArt As String, cVon As String, cBis As String, iFil As Integer, sMWST As String) As Double
On Error GoTo LOKAL_ERROR
Dim sSQL As String
Dim rsrs As Recordset
Dim gdMws As Double

ermERTRAG = 0#

Select Case sMWST
    Case "V"
        gdMws = gdMWStV
    
    Case "E"
    
        gdMws = gdMWStE
        
    Case "O"
        gdMws = gdMWStO
End Select

sSQL = "Select sum(((Preis * 100)/(100 + " & gdMws & ")) - (EKPR * menge)) as maxi from Kassjour where "
sSQL = sSQL & " adate between  " & cVon & " And " & cBis
If iFil = 0 Then
Else
   sSQL = sSQL & " and filiale = " & iFil
End If
sSQL = sSQL & " and artnr = " & sArt
sSQL = sSQL & " and MWST = '" & sMWST & "'"


Set rsrs = gdBase.OpenRecordset(sSQL)
If Not rsrs.EOF Then
    rsrs.MoveFirst
    If Not IsNull(rsrs!maxi) Then
        ermERTRAG = CDbl(rsrs!maxi)
    End If
End If
rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ermERTRAG"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Public Function ermMWST(cWg As String) As String
On Error GoTo LOKAL_ERROR

Dim sSQL As String
Dim rsrs As Recordset

ermMWST = "V"

If IsNumeric(cWg) Then

    sSQL = " Select MWST from ARTIKEL where artnr = " & cWg
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then

        rsrs.MoveFirst
        If Not IsNull(rsrs!MWST) Then
            ermMWST = rsrs!MWST
        End If

    End If
    rsrs.Close: Set rsrs = Nothing

End If

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ermMWST"
    Fehler.gsFehlertext = "Im Programmteil Kasse ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Public Function ermNS(sArt As String, cVon As String, cBis As String, iFil As Integer, sMWST As String) As Double
On Error GoTo LOKAL_ERROR
Dim sSQL As String
Dim rsrs As Recordset
Dim gdMws As Double

ermNS = 0#

Select Case sMWST
    Case "V"
        gdMws = gdMWStV
    
    Case "E"
    
        gdMws = gdMWStE
        
    Case "O"
        gdMws = gdMWStO
End Select

sSQL = "Select (((Preis/(100 + " & gdMws & "))* 100) - (EKPR * menge)* 100 /((Preis/100 + " & gdMws & " )) * 100 ) as maxi from Kassjour where "
sSQL = sSQL & " adate between  " & cVon & " And " & cBis
If iFil = 0 Then
Else
   sSQL = sSQL & " and filiale = " & iFil
End If
sSQL = sSQL & " and artnr = " & sArt
sSQL = sSQL & " and MWST = '" & sMWST & "'"

Set rsrs = gdBase.OpenRecordset(sSQL)
If Not rsrs.EOF Then
    rsrs.MoveFirst
    If Not IsNull(rsrs!maxi) Then
        ermNS = CDbl(rsrs!maxi)
    End If
End If
rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ermNS"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Public Function ermMENGE(sArt As String, cVon As String, cBis As String, iFil As Integer) As Double
On Error GoTo LOKAL_ERROR
Dim sSQL As String
Dim rsrs As Recordset

If sArt = "" Then
    Exit Function
End If

ermMENGE = 0#

sSQL = "Select sum(Menge) as maxi from Kassjour where "
sSQL = sSQL & "  artnr = " & sArt

If cVon <> "" And cBis <> "" Then
    sSQL = sSQL & " and adate between  " & cVon & " And " & cBis
End If

If iFil = 0 Then
Else
   sSQL = sSQL & " and filiale = " & iFil
End If



Set rsrs = gdBase.OpenRecordset(sSQL)
If Not rsrs.EOF Then
    rsrs.MoveFirst
    If Not IsNull(rsrs!maxi) Then
        ermMENGE = CDbl(rsrs!maxi)
    End If
End If
rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ermMENGE"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Public Function ermavgEK(sArt As String, cVon As String, cBis As String, iFil As Integer) As Double
On Error GoTo LOKAL_ERROR
Dim sSQL As String
Dim rsrs As Recordset

ermavgEK = 0#

sSQL = "Select avg(EKPr) as maxi from Kassjour where "
sSQL = sSQL & " adate between  " & cVon & " And " & cBis
If iFil = 0 Then
Else
   sSQL = sSQL & " and filiale = " & iFil
End If
sSQL = sSQL & " and artnr = " & sArt


Set rsrs = gdBase.OpenRecordset(sSQL)
If Not rsrs.EOF Then
    rsrs.MoveFirst
    If Not IsNull(rsrs!maxi) Then
        ermavgEK = CDbl(rsrs!maxi)
    End If
End If
rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ermavgEK"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Public Function ermUMSATZ(sArt As String, cVon As String, cBis As String, iFil As Integer) As Double
On Error GoTo LOKAL_ERROR
Dim sSQL As String
Dim rsrs As Recordset

ermUMSATZ = 0#

sSQL = "Select sum(Preis) as maxi from Kassjour where "
sSQL = sSQL & " adate between  " & cVon & " And " & cBis
If iFil = 0 Then
Else
   sSQL = sSQL & " and filiale = " & iFil
End If
sSQL = sSQL & " and artnr = " & sArt


Set rsrs = gdBase.OpenRecordset(sSQL)
If Not rsrs.EOF Then
    rsrs.MoveFirst
    If Not IsNull(rsrs!maxi) Then
        ermUMSATZ = CDbl(rsrs!maxi)
    End If
End If
rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ermUMSATZ"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function

Public Sub schreibeProtokollNachtAblauf(sZeile As String)
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr     As Integer
    Dim lPos        As Long
    Dim lWert       As Long
    Dim cSatz       As String
    Dim cdatei      As String
    Dim cPfad       As String
    Dim sTime       As String
    Dim ctmp        As String
    Dim cZeile1     As String
    Dim cZeile2     As String
    Dim sRechner    As String

    sRechner = rechnername

    sTime = TimeValue(Now)
    sTime = Right(sTime, 8)
    
    cPfad = gcDBPfad    'Datenbankpfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "LPROTOK\"

    lWert = DateValue(Now)
    ctmp = Format$(lWert, "DD.MM.YY")
    
    cdatei = "NACHTV"
    
    cZeile1 = ""
    cZeile1 = ctmp & Space(1) & sTime & Space(2) & sRechner & Space(1) & sZeile & vbCrLf
    
    
    cSatz = cPfad & cdatei & ".TXT"
    
 
    
    iFileNr = FreeFile
    Open cSatz For Binary As #iFileNr

        lPos = LOF(iFileNr)
        lPos = lPos + 1
        Put #iFileNr, lPos, cZeile1

    Close iFileNr
    
  Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
    
        Resume Next
    Else
    
'        Fehler.gsDescr = err.Description
'        Fehler.gsNumber = err.Number
'        Fehler.gsFormular = "Modul2"
'        Fehler.gsFunktion = "schreibeProtokollNachtAblauf"
'        Fehler.gsFehlertext = "Beim Erstellen des Fehler Protokolls ist ein Fehler aufgetreten."
'
'        Fehlermeldung1
        Resume Next
    End If
End Sub


Public Sub schreibeProtokollDaba(sZeile As String)
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr     As Integer
    Dim lPos        As Long
    Dim lWert       As Long
    Dim cSatz       As String
    Dim cdatei      As String
    Dim cPfad       As String
    Dim sTime       As String
    Dim ctmp        As String
    Dim cZeil       As String
    Dim cZeile2     As String
    
    
    sTime = TimeValue(Now)
    sTime = Right(sTime, 8)
    
    cPfad = gcDBPfad    'Datenbankpfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "LPROTOK\"

    lWert = DateValue(Now)
    ctmp = Format$(lWert, "DD.MM.YY")
    
    cdatei = "Datenbank"
    
    cZeil = ""
    cZeil = ctmp & Space(1) & sTime & Space(2) & sZeile & vbCrLf
    cSatz = cPfad & cdatei & ".TXT"
    
    iFileNr = FreeFile
    Open cSatz For Binary As #iFileNr
    If LOF(iFileNr) > 0 Then
        cZeile2 = Space$(LOF(iFileNr))
        Get #iFileNr, 1, cZeile2
        Close iFileNr
    Else
        Close iFileNr
        Kill cSatz
        
    End If
    
    Kill cSatz
    
    iFileNr = FreeFile
    Open cSatz For Binary As #iFileNr
    
        lPos = LOF(iFileNr)
        lPos = lPos + 1
        Put #iFileNr, lPos, cZeil
    
        lPos = LOF(iFileNr)
        lPos = lPos + 1
        
        Put #iFileNr, lPos, cZeile2
        
    Close iFileNr
    
    
    
  Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
    
        Resume Next
    Else
    
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "schreibeProtokollDaba"
        Fehler.gsFehlertext = "Beim Erstellen des allgemeinen Protokolls ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Public Sub schreibeProtokollBEZ(sZeile As String)
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr     As Integer
    Dim lPos        As Long
    Dim lWert       As Long
    Dim cSatz       As String
    Dim cdatei      As String
    Dim cPfad       As String
    Dim sTime       As String
    Dim ctmp        As String
    Dim cZeil       As String
    Dim cZeile2     As String
    
    
    sTime = TimeValue(Now)
    sTime = Right(sTime, 8)
    
    cPfad = gcDBPfad    'Datenbankpfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "LPROTOK\"

    lWert = DateValue(Now)
    ctmp = Format$(lWert, "DD.MM.YY")
    
    cdatei = "BEZEICH"
    
    cZeil = ""
    cZeil = ctmp & Space(1) & sTime & Space(2) & sZeile & vbCrLf
    cSatz = cPfad & cdatei & ".TXT"
    
    iFileNr = FreeFile
    Open cSatz For Binary As #iFileNr
    If LOF(iFileNr) > 0 Then
        cZeile2 = Space$(LOF(iFileNr))
        Get #iFileNr, 1, cZeile2
        Close iFileNr
    Else
        Close iFileNr
        Kill cSatz
        
    End If
    
    Kill cSatz
    
    iFileNr = FreeFile
    Open cSatz For Binary As #iFileNr
    
        lPos = LOF(iFileNr)
        lPos = lPos + 1
        Put #iFileNr, lPos, cZeil
    
        lPos = LOF(iFileNr)
        lPos = lPos + 1
        
        Put #iFileNr, lPos, cZeile2
        
    Close iFileNr
    
    
    
  Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
    
        Resume Next
    Else
    
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "schreibeProtokollBEZ"
        Fehler.gsFehlertext = "Beim Erstellen des allgemeinen Protokolls ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Public Sub schreibeEinzelFehlermeldungExtra(Mitteilung As Errormessage, firma As FIRMA_, nameFile As String)
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr     As Integer
    Dim lPos        As Long
    Dim lWert       As Long
    Dim cSatz       As String
    Dim cdatei      As String
    Dim cPfad       As String
    Dim sTime       As String
    Dim ctmp        As String
    Dim cZeil       As String
    Dim cZeile2     As String
    Dim sRechner    As String

    sRechner = rechnername
    sRechner = SwapStr(sRechner, ".", "")
    
    sTime = TimeValue(Now)
    sTime = Right$(sTime, 8)
    
    cPfad = App.Path    'anwendungspfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "LERR\"

    lWert = DateValue(Now)
    ctmp = Format$(lWert, "DD.MM.YY")
    
    cdatei = nameFile
    
    cZeil = ""
    cZeil = cZeil & " Leider hat das Winkiss - Programm an dieser Stelle Schwierigkeiten." & vbCrLf
    cZeil = cZeil & " Sie können uns diese Meldung, ggf. um Ihre Kommentare ergänzt, an die" & vbCrLf
    cZeil = cZeil & " 0511 - 9559144 faxen." & vbCrLf
    
    cZeil = cZeil & " ____________________________________________________" & vbCrLf & vbCrLf
    

    cZeil = cZeil & " Datum:            " & ctmp & Space(1) & sTime & Space(2) & vbCrLf
    
    cZeil = cZeil & " Firma:            " & gFirma.FirmaName & vbCrLf
    cZeil = cZeil & " Telefon:          " & gFirma.Tel & vbCrLf & vbCrLf
    
    cZeil = cZeil & " Formular:         " & Fehler.gsFormular & vbCrLf
    cZeil = cZeil & " Funktion:         " & Fehler.gsFunktion & vbCrLf
    cZeil = cZeil & " Nummer:           " & Fehler.gsNumber & vbCrLf
    cZeil = cZeil & " Beschreibung:     " & Fehler.gsDescr & vbCrLf
    cZeil = cZeil & " F Beschreibung:   " & Fehler.gsFehlertext & vbCrLf
    cZeil = cZeil & " Rechnername:      " & sRechner & vbCrLf
    cZeil = cZeil & " Winkiss Version: " & WKVersion & vbCrLf
    
    cZeil = cZeil & " ____________________________________________________" & vbCrLf & vbCrLf
    cZeil = cZeil & " Möchten Sie von uns zurückgerufen werden? ja / nein (bitte unterstreichen)" & vbCrLf
    cZeil = cZeil & " Kommentar:"
    cZeil = cZeil & vbCrLf
    cZeil = cZeil & vbCrLf
    cZeil = cZeil & vbCrLf
    cZeil = cZeil & vbCrLf
    cZeil = cZeil & vbCrLf
    cZeil = cZeil & vbCrLf
    cZeil = cZeil & vbCrLf
    cZeil = cZeil & vbCrLf
    
    cZeil = cZeil & " ____________________________________________________" & vbCrLf
    cZeil = cZeil & " Vielen Dank! Wir bemühen uns, diesen Fehler " & vbCrLf
    cZeil = cZeil & " mit dem nächsten Programmupdate zu beheben." & vbCrLf & vbCrLf
    cZeil = cZeil & " Ihr K.I.S.S. Team" & vbCrLf & vbCrLf & vbCrLf
    
    cZeil = cZeil & " K.I.S.S. Warenwirtschaftssysteme GmbH" & vbCrLf
    cZeil = cZeil & " Brabeckstr. 167, 30539 Hannover" & vbCrLf
    cZeil = cZeil & " Telefon +49 511 955910" & vbCrLf
    cZeil = cZeil & " Telefax +49 511 95591-44" & vbCrLf


    cSatz = cPfad & cdatei & ".TXT"
    
    Kill cPfad & cdatei & ".TXT"
    
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
        MsgBox "Formular: Modul2" & vbCrLf _
     & "Funktion: Fehlermeldung1 " & vbCrLf _
     & "Fehlernummer: " & err.Number & vbCrLf _
     & "Fehlerbeschreibung: " & err.Description & vbCrLf _
     & "Programmversion: " & WKVersion, vbCritical + vbOKOnly, "Winkiss Fehlermeldung:"
     
    End If
End Sub
Public Sub schreibeEinzelFehlermeldung(Mitteilung As Errormessage, firma As FIRMA_, nameFile As String)
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr     As Integer
    Dim lPos        As Long
    Dim lWert       As Long
    Dim cSatz       As String
    Dim cdatei      As String
    Dim cPfad       As String
    Dim sTime       As String
    Dim ctmp        As String
    Dim cZeil       As String
    Dim cZeile2     As String
    Dim sRechner    As String

    sRechner = rechnername
    sRechner = SwapStr(sRechner, ".", "")
    
    sTime = TimeValue(Now)
    sTime = Right(sTime, 8)
    
    cPfad = App.Path    'anwendungspfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "LERR\"

    lWert = DateValue(Now)
    ctmp = Format$(lWert, "DD.MM.YY")
    
    cdatei = nameFile
    
    cZeil = ""
    cZeil = cZeil & " Bitte faxen Sie diese Fehlermeldung faxen an: " & vbCrLf
    cZeil = cZeil & " 0511/9559144" & vbCrLf
    cZeil = cZeil & "_______________________________________________________" & vbCrLf & vbCrLf
    
    
    
    
    cZeil = cZeil & " Datum:            " & ctmp & Space(1) & sTime & Space(2) & vbCrLf
    
    cZeil = cZeil & " Firma:            " & gFirma.FirmaName & vbCrLf
    cZeil = cZeil & " Telefon:          " & gFirma.Tel & vbCrLf & vbCrLf
    
    cZeil = cZeil & " Formular:         " & Fehler.gsFormular & vbCrLf
    cZeil = cZeil & " Funktion:         " & Fehler.gsFunktion & vbCrLf
    cZeil = cZeil & " Nummer:           " & Fehler.gsNumber & vbCrLf
    cZeil = cZeil & " Beschreibung:     " & Fehler.gsDescr & vbCrLf
    cZeil = cZeil & " F Beschreibung:   " & Fehler.gsFehlertext & vbCrLf
    cZeil = cZeil & " Rechnername:      " & sRechner & vbCrLf
    cZeil = cZeil & " Winkiss Version:  " & WKVersion & vbCrLf
    cZeil = cZeil & "_______________________________________________________" & vbCrLf
    
    
    cSatz = cPfad & cdatei & ".TXT"
    
    Kill cPfad & cdatei & ".TXT"
    
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
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "schreibeEinzelFehlermeldung"
        Fehler.gsFehlertext = "Beim Erstellen des Fehler Protokolls ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub

Public Sub schreibeProtokollFehlermeldung(Mitteilung As Errormessage)
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr     As Integer
    Dim lPos        As Long
    Dim lWert       As Long
    Dim cSatz       As String
    Dim cdatei      As String
    Dim cPfad       As String
    Dim sTime       As String
    Dim ctmp        As String
    Dim cZeil       As String
    Dim cZeile2     As String
    Dim sRechner    As String

    sRechner = rechnername
    sRechner = SwapStr(sRechner, ".", "")
    
    
    sTime = TimeValue(Now)
    sTime = Right(sTime, 8)
    
    cPfad = gcDBPfad    'Datenbankpfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "LPROTOK\"

    lWert = DateValue(Now)
    ctmp = Format$(lWert, "DD.MM.YY")
    
    cdatei = "PROFEHLER"
    
    cZeil = ""
    cZeil = " Datum: " & ctmp & Space(1) & sTime & Space(2) & vbCrLf
    cZeil = cZeil & " Formular: " & Fehler.gsFormular & vbCrLf
    cZeil = cZeil & " Funktion: " & Fehler.gsFunktion & vbCrLf
    cZeil = cZeil & " Nummer: " & Fehler.gsNumber & vbCrLf
    cZeil = cZeil & " Beschreibung: " & Fehler.gsDescr & vbCrLf
    cZeil = cZeil & " F Beschreibung: " & Fehler.gsFehlertext & vbCrLf
    cZeil = cZeil & " Rechnername: " & sRechner & vbCrLf
    cZeil = cZeil & " Winkiss Version: " & WKVersion & vbCrLf
    cZeil = cZeil & "___________________________________________________________" & vbCrLf
    
    cSatz = cPfad & cdatei & ".TXT"
    
    iFileNr = FreeFile
    Open cSatz For Binary As #iFileNr
    If LOF(iFileNr) > 0 Then
        cZeile2 = Space$(LOF(iFileNr))
        Get #iFileNr, 1, cZeile2
        Close iFileNr
    Else
        Close iFileNr
        Kill cSatz
        
    End If
    
    Kill cSatz
    
    iFileNr = FreeFile
    Open cSatz For Binary As #iFileNr
    
        lPos = LOF(iFileNr)
        lPos = lPos + 1
        Put #iFileNr, lPos, cZeil
    
        lPos = LOF(iFileNr)
        lPos = lPos + 1
        
        Put #iFileNr, lPos, cZeile2
        
    Close iFileNr
    
    
    
  Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
    
        Resume Next
    Else
    
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "schreibeProtokollFehlermeldung"
        Fehler.gsFehlertext = "Beim Erstellen des Fehler Protokolls ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Public Sub schreibeProtokollDabaFehlerAbmeldung(dbErr As DBErrormessage)
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr     As Integer
    Dim lPos        As Long
    Dim lWert       As Long
    Dim cSatz       As String
    Dim cdatei      As String
    Dim cPfad       As String
    Dim sTime       As String
    Dim ctmp        As String
    Dim cZeil       As String
    Dim cZeile2     As String
    
    
    sTime = TimeValue(Now)
    sTime = Right(sTime, 8)
    
    cPfad = gcDBPfad    'Datenbankpfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "LPROTOK\"

    lWert = DateValue(Now)
    ctmp = Format$(lWert, "DD.MM.YY")
    
    cdatei = "DBErrMel"
    
    cZeil = ""
    cZeil = cZeil & " Dieser Computer wurde nicht richtig heruntergefahren: " & vbCrLf
    cZeil = cZeil & "_______________________________________________________" & vbCrLf & vbCrLf
    cZeil = cZeil & " zuletzt angemeldet " & vbCrLf
    cZeil = cZeil & " am :              " & dbErr.gsDatum & vbCrLf
    cZeil = cZeil & " um :              " & dbErr.gsZeit & vbCrLf
    cZeil = cZeil & " Computername:     " & dbErr.gsPcname & vbCrLf
    cZeil = cZeil & " BedienerNr:       " & dbErr.gsBednr & vbCrLf
    cZeil = cZeil & " Bedienername:     " & dbErr.gsBedname & vbCrLf
    cZeil = cZeil & "_______________________________________________________" & vbCrLf & vbCrLf
    cSatz = cPfad & cdatei & ".TXT"
    
    iFileNr = FreeFile
    Open cSatz For Binary As #iFileNr
    If LOF(iFileNr) > 0 Then
        cZeile2 = Space$(LOF(iFileNr))
        Get #iFileNr, 1, cZeile2
        Close iFileNr
    Else
        Close iFileNr
        Kill cSatz
        
    End If
    
    Kill cSatz
    
    iFileNr = FreeFile
    Open cSatz For Binary As #iFileNr
    
        lPos = LOF(iFileNr)
        lPos = lPos + 1
        Put #iFileNr, lPos, cZeil
    
        lPos = LOF(iFileNr)
        lPos = lPos + 1
        
        Put #iFileNr, lPos, cZeile2
        
    Close iFileNr
    
    
    
  Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
    
        Resume Next
    Else
    
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "schreibeProtokollDabaFehlerAbmeldung"
        Fehler.gsFehlertext = "Beim Erstellen des allgemeinen Protokolls ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Public Sub schreibeProtokollAufdatok(sZeile As String)
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr     As Integer
    Dim lPos        As Long
    Dim lWert       As Long
    Dim cSatz       As String
    Dim cdatei      As String
    Dim cPfad       As String
    Dim sTime       As String
    Dim ctmp        As String
    Dim cZeil       As String
    Dim cZeile2     As String
    
    
    sTime = TimeValue(Now)
    sTime = Right(sTime, 8)
    
    cPfad = gcDBPfad    'Datenbankpfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "PROTOK\"

    lWert = DateValue(Now)
    ctmp = Format$(lWert, "DD.MM.YY")
    
    cdatei = "AUFDATOK"
    
    cZeil = ""
    cZeil = ctmp & Space(1) & sTime & Space(2) & sZeile & vbCrLf
    cSatz = cPfad & cdatei & ".TXT"
    
    iFileNr = FreeFile
    Open cSatz For Binary As #iFileNr
    If LOF(iFileNr) > 0 Then
        cZeile2 = Space$(LOF(iFileNr))
        Get #iFileNr, 1, cZeile2
        Close iFileNr
    Else
        Close iFileNr
        Kill cSatz
        
    End If
    
    Kill cSatz
    
    iFileNr = FreeFile
    Open cSatz For Binary As #iFileNr
    
        lPos = LOF(iFileNr)
        lPos = lPos + 1
        Put #iFileNr, lPos, cZeil
    
        lPos = LOF(iFileNr)
        lPos = lPos + 1
        
        Put #iFileNr, lPos, cZeile2
        
    Close iFileNr
    
    
    
  Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
    
        Resume Next
    Else
    
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "schreibeProtokollAufdatok"
        Fehler.gsFehlertext = "Beim Erstellen des allgemeinen Protokolls ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Public Sub schreibetheBigFehler(sZeile As String)
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr     As Integer
    Dim lPos        As Long
    Dim lWert       As Long
    Dim cSatz       As String
    Dim cdatei      As String
    Dim cPfad       As String
    Dim sTime       As String
    Dim ctmp        As String
    Dim cZeil       As String
    Dim cZeile2     As String
    
    
    sTime = TimeValue(Now)
    sTime = Right(sTime, 8)
    
    cPfad = gcDBPfad    'Datenbankpfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "PROTOK\"

    lWert = DateValue(Now)
    ctmp = Format$(lWert, "DD.MM.YY")
    
    cdatei = "DatenErr"
    
    cZeil = ""
    cZeil = ctmp & Space(1) & sTime & Space(2) & sZeile & vbCrLf
    cSatz = cPfad & cdatei & ".TXT"
    
    iFileNr = FreeFile
    Open cSatz For Binary As #iFileNr
    If LOF(iFileNr) > 0 Then
        cZeile2 = Space$(LOF(iFileNr))
        Get #iFileNr, 1, cZeile2
        Close iFileNr
    Else
        Close iFileNr
        Kill cSatz
        
    End If
    
    Kill cSatz
    
    iFileNr = FreeFile
    Open cSatz For Binary As #iFileNr
    
        lPos = LOF(iFileNr)
        lPos = lPos + 1
        Put #iFileNr, lPos, cZeil
    
        lPos = LOF(iFileNr)
        lPos = lPos + 1
        
        Put #iFileNr, lPos, cZeile2
        
    Close iFileNr
    
    
    
  Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
    
        Resume Next
    Else
    
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "schreibetheBigFehler"
        Fehler.gsFehlertext = "Beim Erstellen des allgemeinen Protokolls ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Public Sub schreibeProtokollKassErr(sZeile As String)
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr     As Integer
    Dim lPos        As Long
    Dim lWert       As Long
    Dim sTime       As String
    Dim ctmp        As String
    Dim cZeil       As String
    Dim cPfad       As String
    Dim cdat        As String
    
    cPfad = gcDBPfad    'Datenbankpfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    cPfad = cPfad & "LProtok\"
    cdat = cPfad & "B" & srechnertab & ".txt"
    
    sTime = TimeValue(Now)
    sTime = Right(sTime, 8)

    lWert = DateValue(Now)
    ctmp = Format$(lWert, "DD.MM.YY")
    
    cZeil = ""
    cZeil = ctmp & Space(1) & sTime & Space(2) & sZeile & vbCrLf
    
    iFileNr = FreeFile
    Open cdat For Binary As #iFileNr
    
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
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "schreibeProtokollKassErr"
        Fehler.gsFehlertext = "Beim Erstellen des allgemeinen Protokolls ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Public Sub KillProtokollKassErr()
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr     As Integer
    Dim lPos        As Long
    Dim lWert       As Long
    Dim sTime       As String
    Dim ctmp        As String
    Dim cZeil       As String
    Dim cZeile2     As String
    Dim cPfad       As String
    Dim cdat        As String
    
    cPfad = gcDBPfad    'Datenbankpfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    cPfad = cPfad & "LProtok\"
    cdat = cPfad & "B" & srechnertab & ".txt"
    
    Kill cdat

  Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Or err.Number = 55 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "KillProtokollKassErr"
        Fehler.gsFehlertext = "Beim Erstellen des allgemeinen Protokolls ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Public Sub abinsProtokoll(sdiff As String)
On Error GoTo LOKAL_ERROR

    Dim iFileNr     As Integer
    Dim lPos        As Long
    Dim lWert       As Long
    Dim sTime       As String
    Dim ctmp        As String
    Dim cZeil       As String
    Dim cZeile2     As String
    Dim cPfad       As String
    Dim cPfad1      As String
    Dim cdat        As String
    Dim cDat1       As String
    
    cPfad = gcDBPfad    'Datenbankpfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    cPfad = cPfad & "LProtok\"
    cdat = cPfad & "B" & srechnertab & ".txt"
    
    
    cPfad1 = gcDBPfad    'Datenbankpfad
    If Right$(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If
    cPfad1 = cPfad1 & "Protok\"
    cDat1 = cPfad1 & "BARAUS.txt"
   
    
    
    
  
    cZeil = srechnertab & "********************" & sdiff & " Sekunden " & vbCrLf
    
    iFileNr = FreeFile
    Open cdat For Binary As #iFileNr
    If LOF(iFileNr) > 0 Then
        cZeile2 = Space$(LOF(iFileNr))
        Get #iFileNr, 1, cZeile2
        Close iFileNr
    Else
        Close iFileNr
        Kill cdat
        
    End If
    
    Kill cdat
    
    iFileNr = FreeFile
    Open cDat1 For Binary As #iFileNr
    
        lPos = LOF(iFileNr)
        lPos = lPos + 1
        Put #iFileNr, lPos, cZeil
    
        lPos = LOF(iFileNr)
        lPos = lPos + 1
        
        Put #iFileNr, lPos, cZeile2
        
    Close iFileNr

Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then

        Resume Next
    Else
    
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "abinsProtokoll"
        Fehler.gsFehlertext = "Beim Erstellen des allgemeinen Protokolls ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Public Sub schreibeProtokollKassStop(sZeile As String)
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
    
    cPfad = gcDBPfad    'Datenbankpfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If

    lWert = DateValue(Now)
    ctmp = Format$(lWert, "DD.MM.YY")
    
    cdatei = "KASSSTOP"
    
    cZeil = ""
    cZeil = ctmp & Space(1) & sTime & Space(2) & sZeile & vbCrLf
    cSatz = cPfad & cdatei & ".TXT"
    
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
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "schreibeProtokollKassErr"
        Fehler.gsFehlertext = "Beim Erstellen des allgemeinen Protokolls ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Public Sub schreibeProtokollKassStopnurfürdieseKasse(sZeile As String)
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
    
    cPfad = gcDBPfad    'Datenbankpfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If

    lWert = DateValue(Now)
    ctmp = Format$(lWert, "DD.MM.YY")
    
    cdatei = "KASSSTOP" & Trim(gcKasNum) & ""
    
    cZeil = ""
    cZeil = ctmp & Space(1) & sTime & Space(2) & sZeile & vbCrLf
    cSatz = cPfad & cdatei & ".TXT"
    
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
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "schreibeProtokollKassStopnurfürdieseKasse"
        Fehler.gsFehlertext = "Beim Erstellen des allgemeinen Protokolls ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Public Sub schreibeProtokollKassStopfürALLEKassen(sZeile As String)
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
    
    cPfad = gcDBPfad    'Datenbankpfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If

    lWert = DateValue(Now)
    ctmp = Format$(lWert, "DD.MM.YY")
    
    cdatei = "KASSSTOP_ALLE"
    
    cZeil = ""
    cZeil = ctmp & Space(1) & sTime & Space(2) & sZeile & vbCrLf
    cSatz = cPfad & cdatei & ".TXT"
    
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
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "schreibeProtokollKassStopfürALLEKassen"
        Fehler.gsFehlertext = "Beim Erstellen des allgemeinen Protokolls ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Public Sub schreibeAktionStop(sZeile As String, sAktion As String)
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
    Dim sRechner    As String

    sRechner = rechnername
    
    sTime = TimeValue(Now)
    sTime = Right(sTime, 8)
    
    cPfad = gcDBPfad    'Datenbankpfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If

    lWert = DateValue(Now)
    ctmp = Format$(lWert, "DD.MM.YY")
    
    cdatei = sAktion
    
    If UCase(sAktion) = "KOMPRIMIERUNG" Then
    
        cZeil = ""
        
        cZeil = "Seit " & ctmp & Space(1) & sTime & Space(1) & "Uhr versucht der Computer: " & sRechner
        cZeil = cZeil & " mit dem angemeldeten Benutzer: " & gcUserName & "(" & gcBedienerNr & ") die Datenbank zu komprimieren."
    Else
    
        cZeil = ""
        cZeil = ctmp & Space(1) & sTime & Space(2) & sRechner & Space(2) & sZeile & vbCrLf
    
    End If
    cSatz = cPfad & cdatei & ".TXT"
    
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
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "schreibeAktionStop"
        Fehler.gsFehlertext = "Beim Erstellen des allgemeinen Protokolls ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Public Sub schreibeProtokollDINA4Err(sZeile As String)
    On Error GoTo LOKAL_ERROR
    
    Dim iFileNr     As Integer
    Dim lPos        As Long
    Dim lWert       As Long
    Dim cSatz       As String
    Dim cdatei      As String
    Dim cPfad       As String
    Dim sTime       As String
    Dim ctmp        As String
    Dim cZeil       As String
    Dim cZeile2     As String
    
    sTime = TimeValue(Now)
    sTime = Right(sTime, 8)
    
    cPfad = gcDBPfad    'Datenbankpfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "PROTOK\"

    lWert = DateValue(Now)
    ctmp = Format$(lWert, "DD.MM.YY")
    
    cdatei = "DINA4Err"
    
    cZeil = ""
    cZeil = ctmp & Space(1) & sTime & Space(2) & sZeile & vbCrLf
    cSatz = cPfad & cdatei & ".TXT"
    
    iFileNr = FreeFile
    Open cSatz For Binary As #iFileNr
    If LOF(iFileNr) > 0 Then
        cZeile2 = Space$(LOF(iFileNr))
        Get #iFileNr, 1, cZeile2
        Close iFileNr
    Else
        Close iFileNr
        Kill cSatz
        
    End If
    
    Kill cSatz
    
    iFileNr = FreeFile
    Open cSatz For Binary As #iFileNr
    
        lPos = LOF(iFileNr)
        lPos = lPos + 1
        Put #iFileNr, lPos, cZeil
    
        lPos = LOF(iFileNr)
        lPos = lPos + 1
        
        Put #iFileNr, lPos, cZeile2
        
    Close iFileNr
    
  Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "schreibeProtokollDINA4Err"
        Fehler.gsFehlertext = "Beim Erstellen des allgemeinen Protokolls ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Public Sub anzeige(SFarbe As String, satext As String, lab1 As Label)
    On Error GoTo LOKAL_ERROR
    Dim cPfad As String
    
    cPfad = gcPfad & "/mcitest.wav"
    
    Dim cPfad1 As String
    
    cPfad1 = App.Path & "/erfolg.wav"
    
    Dim cPfad2 As String
    cPfad2 = App.Path & "/artikel.wav"
    
    Dim cPfad3 As String
    cPfad3 = App.Path & "/Laser.wav"
    
    Dim cPfad4 As String
    cPfad4 = App.Path & "/Off.wav"
    
    If UCase(SFarbe) = "ROT" Then
        lab1.ForeColor = vbRed
        If gbSound Then Call PlaySound(cPfad, 0)
        Screen.MousePointer = 0
    ElseIf UCase(SFarbe) = "ERFOLG" Then
        lab1.ForeColor = glS1
        If gbSound Then Call PlaySound(cPfad1, 0)
        Screen.MousePointer = 0
    ElseIf UCase(SFarbe) = "LASER" Then
        lab1.ForeColor = glS1
        If gbSound Then Call PlaySound(cPfad3, 0)
        Screen.MousePointer = 0
    ElseIf UCase(SFarbe) = "JUGENDSCHUTZ" Then
    
        lab1.Caption = satext
        lab1.Refresh
    
        lab1.ForeColor = vbRed
        If gbSound Then Call PlaySound(cPfad4, 0)
        Screen.MousePointer = 0
    ElseIf UCase(SFarbe) = "ARTIKEL" Then
        lab1.ForeColor = vbRed
        If gbSound Then Call PlaySound(cPfad2, 0)
        Screen.MousePointer = 0
    ElseIf UCase(SFarbe) = "ROT2" Then
        lab1.ForeColor = vbRed
        Screen.MousePointer = 0
    ElseIf UCase(SFarbe) = "KNORMAL" Then
    ElseIf UCase(SFarbe) = "BLACK" Then
        lab1.ForeColor = vbBlack 'glS1
    Else
        lab1.ForeColor = glS1
    End If
    
    lab1.Caption = satext
    lab1.Refresh
   

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "anzeige"
    Fehler.gsFehlertext = "Bei der Anzeige ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub zeigeHilfeDabapfad(cverz As String, Helpname As String)
    On Error GoTo LOKAL_ERROR
    
    Dim cPfad   As String
    Dim cPfad1  As String
    Dim lRet    As Long
    
    Dim sWordp  As String
    Dim sbesy   As String
    
    sWordp = "C:\Programme\Microsoft Office\Office\WINWORD.exe "
    sWordp = ShortPath(sWordp)
    
    
    cPfad = gcDBPfad    'dabapfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    cPfad1 = cPfad & cverz
    cPfad = cPfad & cverz & "\"
    
    cPfad = ShortPath(cPfad)
    cPfad1 = ShortPath(cPfad1)
    
    lRet = 88
    sbesy = Trim(Get_OS)

    If FileExists(cPfad & Helpname) Then
'    If FindFile(cpfad, Helpname & ".rtf") Then
        If sbesy = "Windows 98 SE" Then
        
            Select Case lRet
                Case Is = 8
                    lRet = Shell(sWordp & cPfad & Helpname, vbNormalFocus)
                    If lRet = 8 Then
                        lRet = Shell("WRITE.exe " & cPfad & Helpname, vbNormalFocus)
                    End If
                Case Is = 88, 5, 8
                    lRet = Shell("WRITE.exe " & cPfad & Helpname, vbNormalFocus)
            End Select
            
            
        Else
        
            lRet = ShellExecute(260052, "open", Helpname, "", cPfad1, 1)
            
            Select Case lRet
                Case Is = 8
                    lRet = Shell(sWordp & cPfad & Helpname, vbNormalFocus)
                    If lRet = 8 Then
                        lRet = Shell("WRITE.exe " & cPfad & Helpname, vbNormalFocus)
                    End If
                Case Is = 88, 5, 8
                    lRet = Shell("WRITE.exe " & cPfad & Helpname, vbNormalFocus)
            End Select
        End If
    End If
    
    Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Or err.Number = 5 Then
        Resume Next
    Else
    
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "zeigeHilfeDabapfad"
        Fehler.gsFehlertext = "Beim Anzeigen der Anleitung ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Public Sub ErzeugeNeueKassenDatei(lbl6 As Label, lbl3 As Label, txtStatus As TextBox, picprogress As PictureBox)
On Error GoTo LOKAL_ERROR
    
    Dim lcount          As Long
    Dim lRet            As Long
    Dim lfail           As Long
    Dim cPfad1          As String
    Dim cPfad           As String
    Dim cSQL            As String
    Dim cQuelle         As String
    Dim cZiel           As String
    Dim sQuell          As String
    Dim cZielaccess     As String
    Dim sZiel           As String
    Dim cFName          As String
    Dim cName           As String
    Dim cdat            As String
    Dim i               As Integer
    Dim iFileNr         As Integer
    Dim rsQ             As Recordset
    Dim rsZ             As Recordset
    Dim cPfad2          As String
    Dim cFilialNr       As String
    Dim db              As DAO.Database
    
    cPfad1 = gcDBPfad
    If Right(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If
    
    cPfad2 = cPfad1 & "EXPORT\FZ.mdb"
    Set db = OpenDatabase(cPfad1 & "EXPORT\FZ.mdb", False, False)
    
    cPfad = cPfad1 & "KOMP\"

    If Len(Trim(gcFilNr)) = 1 Then
        cFilialNr = "0" & gcFilNr
    ElseIf Len(Trim(gcFilNr)) = 2 Then
    
        cFilialNr = gcFilNr
    End If
    
    cFName = "F" & cFilialNr
    
    lbl6.Caption = "Übertragung der Tabellen ins Kassenausgangsverzeichnis"
    lbl6.Refresh
    lbl3.Caption = "10"
    lbl3.Refresh

    Kill gcDBPfad & "\KOMP\*.*"
    
    loeschNEW "AfcKas", gdBase
    cSQL = "Select * into AfcKas from AFCBUCH where KASNUM = " & gcKasNum
    gdBase.Execute cSQL, dbFailOnError
    
    TransferTab gdBase, cPfad2, "AfcKas"
    loeschNEW "afcbuch", db
    
    cSQL = "Select * into AFCBUCH from AfcKas "
    db.Execute cSQL, dbFailOnError
    
    loeschNEW "AfcKas", db
    db.Close
    
    txtStatus.Text = CStr(10 * 100 / 27)
    
    ErzeugeFILTAU
    TransferTab gdBase, cPfad2, "FILTAU"
    
    txtStatus.Text = CStr(11 * 100 / 27)
    
    ErzeugeEANPOUT
    TransferTab gdBase, cPfad2, "EANPOUT"
    
    txtStatus.Text = CStr(12 * 100 / 27)
    
    ErzeugeKVKPR1POUT
    TransferTab gdBase, cPfad2, "KVKPR1POUT"
    
    txtStatus.Text = CStr(13 * 100 / 27)
    
    ErzeugeBESTPOUT
    TransferTab gdBase, cPfad2, "BESTPOUT"
    
    txtStatus.Text = CStr(14 * 100 / 27)
    
    ErzeugeRetoureOUT
    TransferTab gdBase, cPfad2, "RETOUT"
    
    txtStatus.Text = CStr(15 * 100 / 27)
    
    ErzeugeKBOUT
    TransferTab gdBase, cPfad2, "KBOUT"
    
    txtStatus.Text = CStr(16 * 100 / 27)
    
    ErzeugeXOUT ("GUTZ")
    TransferTab gdBase, cPfad2, "GUTZO"
    
    ErzeugeXOUT ("GUHIS")
    TransferTab gdBase, cPfad2, "GUHISO"
    
    ErzeugeXOUT ("KUNDEDEL")
    TransferTab gdBase, cPfad2, "KUNDEDELO"
    
    txtStatus.Text = CStr(17 * 100 / 27)
    
    ErzeugeXOUT ("ZUGANGF")
    TransferTab gdBase, cPfad2, "ZUGANGFO"
    
    ErzeugeKKZahlte_OUT
    TransferTab gdBase, cPfad2, "KKZAHLTEO"
    
    ErzeugeXOUT ("LASTZAHLTE")
    TransferTab gdBase, cPfad2, "LASTZAHLTEO"
    
    ErzeugeXOUT ("MBORDER")
    TransferTab gdBase, cPfad2, "MBORDERO"
    
    If NewTableSuchenDBKombi("MBORDERDEL", gdBase) = True Then
        ErzeugeXOUT ("MBORDERDEL")
        TransferTab gdBase, cPfad2, "MBORDERDELO"
    End If
    
    ErzeugeXOUT ("GEMZ")
    TransferTab gdBase, cPfad2, "GEMZO"
    
    ErzeugeXOUT ("MAILFB")
    TransferTab gdBase, cPfad2, "MAILFBO"
    
    If NewTableSuchenDBKombi("BONUSNR", gdBase) = True Then
        ErzeugeXOUT ("BONUSNR")
        TransferTab gdBase, cPfad2, "BONUSNRO"
    End If
    
    ErzeugeXOUT ("GANALYSEALL")
    TransferTab gdBase, cPfad2, "GANALYSEALLO"
    
    ErzeugeXOUT ("KABUCH")
    TransferTab gdBase, cPfad2, "KABUCHO"
    
    ErzeugeXOUT ("BONUS_SYS")
    TransferTab gdBase, cPfad2, "BONUS_SYSO"
    
    ErzeugeXOUT ("UNTERWF")
    TransferTab gdBase, cPfad2, "UNTERWFO"
    
    txtStatus.Text = CStr(18 * 100 / 27)
    
    ErzeugeXOUT ("ALTERG")
    TransferTab gdBase, cPfad2, "ALTERGO"
    
    txtStatus.Text = CStr(18 * 100 / 27)
    
    ErzeugeXOUT ("NEINVK")
    TransferTab gdBase, cPfad2, "NEINVKO"
    
    ErzeugeXOUT ("KREDITZA")
    TransferTab gdBase, cPfad2, "KREDITZAO"
    
    txtStatus.Text = CStr(18 * 100 / 27)
    
    ErzeugeXOUT ("KKZAHL")
    TransferTab gdBase, cPfad2, "KKZAHLO"
    
    txtStatus.Text = CStr(19 * 100 / 27)
    
    ErzeugeXOUT ("LASTZAHL")
    TransferTab gdBase, cPfad2, "LASTZAHLO"
    
    txtStatus.Text = CStr(20 * 100 / 27)
    
    ErzeugeXOUT ("FEEDB")
    TransferTab gdBase, cPfad2, "FEEDBO"
    
    ErzeugeXOUT ("FEEDB_TRANS")
    TransferTab gdBase, cPfad2, "FEEDB_TRANSO"
    
    ErzeugeXOUT ("FEEDBF")
    TransferTab gdBase, cPfad2, "FEEDBFO"
    
    txtStatus.Text = CStr(21 * 100 / 27)
    
    ErzeugeXOUT ("KAEINAUSF")
    TransferTab gdBase, cPfad2, "KAEINAUSFO"
    
    txtStatus.Text = CStr(21 * 100 / 27)
    
    ErzeugeBESTOUT
    TransferTab gdBase, cPfad2, "BESTOUT"
    
    txtStatus.Text = CStr(22 * 100 / 27)
    
    ErzeugeSendok "BARGELD", "BARGOUT"
    TransferTab gdBase, cPfad2, "BARGOUT"
    
    txtStatus.Text = CStr(23 * 100 / 27)
    
    ErzeugeXOUT ("STORNO2")
    TransferTab gdBase, cPfad2, "STORNO2O"
    
    ErzeugeXOUT ("ARTDET")
    TransferTab gdBase, cPfad2, "ARTDETO"
    
    ErzeugeXOUT ("KASSBEDP")
    TransferTab gdBase, cPfad2, "KASSBEDPO"
    
    txtStatus.Text = CStr(23 * 100 / 27)
    
    ErzeugeXOUT ("ABSCHOPF")
    TransferTab gdBase, cPfad2, "ABSCHOPFO"
    
    ErzeugeXOUT ("DUKATENB")
    TransferTab gdBase, cPfad2, "DUKATENBO"
    
    ErzeugeXOUT ("AFCSTATP")
    TransferTab gdBase, cPfad2, "AFCSTATPO"
    
    txtStatus.Text = CStr(23 * 100 / 27)
    
    If gbECTOZ Then
        sicherdta
    
        ErzeugeSendok "DTA", "DTAOUT"
        TransferTab gdBase, cPfad2, "DTAOUT"
        
        KompressDTAWKL57
    End If
    
    ErzeugeSTE_OUT
    TransferTab gdBase, cPfad2, "STE_OUT"
    
    txtStatus.Text = CStr(24 * 100 / 27)

    lbl6.Caption = "Übertragung der Tabellen ins Kassenausgangsverzeichnis"
    lbl6.Refresh
    lbl3.Caption = "23"
    lbl3.Refresh
    
    ErzeugeKolout
    TransferTab gdBase, cPfad2, "KOLOUT"
    
    txtStatus.Text = CStr(25 * 100 / 27)
    
    ErzeugeXOUT ("KASSBON")
    TransferTab gdBase, cPfad2, "KASSBONO"
    
    ErzeugeXOUT ("KUNDENBONUS")
    TransferTab gdBase, cPfad2, "KUNDENBONUSO"
    
    txtStatus.Text = CStr(26 * 100 / 27)

    lbl6.Caption = "Übertragung der Tabellen ins Kassenausgangsverzeichnis"
    lbl6.Refresh
    lbl3.Caption = "24"
    lbl3.Refresh
    
    ErzeugeZ_out
    TransferTab gdBase, cPfad2, "Z_OUT"
    
    loeschNEW "Z_OUT", gdBase
    
    
    
    
    
    
    
    
    
    Dim sPasswordFZ As String
    sPasswordFZ = "XYC6T349G6"
    
    Set db = OpenDatabase(cPfad1 & "EXPORT\FZ.mdb", True, False)
    db.NewPassword "", sPasswordFZ
    db.Close
    
    
    
    
    
    
    
    
    
    txtStatus.Text = CStr(27 * 100 / 27)

    lbl6.Caption = "Übertragung der Tabellen ins Kassenausgangsverzeichnis"
    lbl6.Refresh
    lbl3.Caption = "25"
    lbl3.Refresh
    
    Dim slf             As String
    Dim lLFNR           As Long
    
    lLFNR = lfnrErmitteln("F")
    lLFNR = lLFNR + 1
    
    gllfnr = lLFNR
    slf = CStr(lLFNR)
    gslfnr = Space(5 - Len(slf)) & slf
    gslfnr = SwapStr(gslfnr, " ", "0")
    
    txtStatus.Text = "0"

    ZentraleWillsWissen "Datei: " & "F" & cFilialNr & gslfnr & ".lzh wird bereitgestellt"
    Zip_Folder "", gcDBPfad & "\EXPORT", gsZoutPfad & "\F" & cFilialNr & gslfnr & ".lzh", txtStatus
    
    picprogress.Visible = False

    sQuell = gsZoutPfad & "\F" & cFilialNr & gslfnr & ".lzh"
    sZiel = cPfad1 & "ABSCHLUS\F" & cFilialNr & gslfnr & ".lzh"
    lRet = CopyFile(sQuell, sZiel, lfail)
    
    If lRet = 1 Then
        schreibeFProtokoll "Datei: F" & cFilialNr & gslfnr & ".lzh erzeugt"
    End If
    
    cName = "F" & cFilialNr & gslfnr
    cdat = DateValue(Now) & " " & TimeValue(Now)
    
    If Not Modul6.FindFile(gsZoutPfad, cName & ".lzh") Then
        gbErfolg = False
    Else
        gbErfolg = True
        lfnrSchreiben gllfnr, cName, cdat
    End If
    
Exit Sub
LOKAL_ERROR:
'    If err.Number = 13 Or err.Number = 53 Or err.Number = 3010 Or err.Number = 58 Then
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "ErzeugeNeueKassenDatei"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1
        
    End If
End Sub
Public Sub ErzeugeNeueKassenDateiBESTAKT(txtStatus As TextBox)
On Error GoTo LOKAL_ERROR
    
    Dim lcount          As Long
    Dim lRet            As Long
    Dim lfail           As Long
    Dim cPfad1          As String
    Dim cPfad           As String
    Dim cSQL            As String
    Dim cQuelle         As String
    Dim cZiel           As String
    Dim sQuell          As String
    Dim cZielaccess     As String
    Dim sZiel           As String
    Dim cFName          As String
    Dim cName           As String
    Dim cdat            As String
    Dim i               As Integer
    Dim iFileNr         As Integer
    Dim rsQ             As Recordset
    Dim rsZ             As Recordset
    Dim cPfad2          As String
    Dim cFilialNr       As String
    Dim db              As DAO.Database
    
    cPfad1 = gcDBPfad
    If Right(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If
    
    cPfad2 = cPfad1 & "EXPORT\FZ.mdb"
    Set db = OpenDatabase(cPfad1 & "EXPORT\FZ.mdb", False, False)
    
    cPfad = cPfad1 & "KOMP\"

    If Len(Trim(gcFilNr)) = 1 Then
        cFilialNr = "0" & gcFilNr
    ElseIf Len(Trim(gcFilNr)) = 2 Then
    
        cFilialNr = gcFilNr
    End If
    
    cFName = "F" & cFilialNr
    
    Kill gcDBPfad & "\KOMP\*.*"
    
    loeschNEW "AfcKas", gdBase
    cSQL = "Select * into AfcKas from AFCBUCH where KASNUM = " & gcKasNum
    gdBase.Execute cSQL, dbFailOnError
    
    TransferTab gdBase, cPfad2, "AfcKas"
    loeschNEW "afcbuch", db
    
    cSQL = "Select * into AFCBUCH from AfcKas "
    db.Execute cSQL, dbFailOnError
    
    loeschNEW "AfcKas", db
    
    'Zugang
    loeschNEW "ZUKas", gdBase
    cSQL = "Select * into ZUKas from zugang "
    gdBase.Execute cSQL, dbFailOnError
    
    TransferTab gdBase, cPfad2, "ZUKas"
    
    cSQL = "Delete from zugang "
    gdBase.Execute cSQL, dbFailOnError
    'zugang ende
    
    db.Close
    
    ErzeugeXOUT ("BESTAKT")
    TransferTab gdBase, cPfad2, "BESTAKTO"
    
    ErzeugeXOUT ("ARTMERK")
    TransferTab gdBase, cPfad2, "ARTMERKO"
    
    Dim slf             As String
    Dim lLFNR           As Long
    
    lLFNR = lfnrErmitteln("F")
    lLFNR = lLFNR + 1
    
    gllfnr = lLFNR
    slf = CStr(lLFNR)
    gslfnr = Space(5 - Len(slf)) & slf
    gslfnr = SwapStr(gslfnr, " ", "0")
    
    sQuell = gcDBPfad & "\EXPORT\FZ.mdb"
    sZiel = cPfad1 & "ABSCHLUS\F" & cFilialNr & gslfnr & ".mdb"
    lRet = CopyFile(sQuell, sZiel, lfail)

    Zip_Folder "", gcDBPfad & "\EXPORT", gsZoutPfad & "\F" & cFilialNr & gslfnr & ".lzh", txtStatus

    sQuell = gsZoutPfad & "\F" & cFilialNr & gslfnr & ".lzh"
    sZiel = cPfad1 & "ABSCHLUS\F" & cFilialNr & gslfnr & ".lzh"
    lRet = CopyFile(sQuell, sZiel, lfail)
    
    If lRet = 1 Then
        schreibeFProtokoll "Datei: F" & cFilialNr & gslfnr & ".lzh erzeugt"
    End If
    
    cName = "F" & cFilialNr & gslfnr
    cdat = DateValue(Now) & " " & TimeValue(Now)
    
    If Not Modul6.FindFile(gsZoutPfad, cName & ".lzh") Then
        gbErfolg = False
    Else
        gbErfolg = True
        lfnrSchreiben gllfnr, cName, cdat
    End If
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "ErzeugeNeueKassenDateiBESTAKT"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Public Sub ErzeugeFILTAU()
    On Error GoTo LOKAL_ERROR

    Dim cSQL        As String
    Dim cDatname    As String
    Dim cPfad1      As String
    
    cPfad1 = gcDBPfad
    If Right(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If
    cDatname = "FILTAU" & Trim$(gcFilNr)

    loeschNEW "FILTAU", gdBase
    
    cSQL = "Select * into FILTAU from TAUSCH where"
    cSQL = cSQL & " SENDOK = False"
    gdBase.Execute cSQL, dbFailOnError

    
    cSQL = "update TAUSCH set SENDOK = True where SENDOK = False"
    gdBase.Execute cSQL, dbFailOnError

Exit Sub
LOKAL_ERROR:
    If err.Number = 3078 Or err.Number = 3010 Or err.Number = 58 Or err.Number = 53 Or err.Number = 3043 Or err.Number = 91 Or err.Number = 3011 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "ErzeugeFILTAU"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Public Sub ErzeugeSendok(stabfrom As String, stabziel As String)
    On Error GoTo LOKAL_ERROR

    Dim cSQL        As String
   
    loeschNEW stabziel, gdBase
    
    cSQL = "Select * into " & stabziel & " from " & stabfrom & " where"
    cSQL = cSQL & " SENDOK = False"
    schreibeProtokollDabaAblauf cSQL: gdBase.Execute cSQL, dbFailOnError

    
    cSQL = "update " & stabfrom & " set SENDOK = True where SENDOK = False"
    schreibeProtokollDabaAblauf cSQL: gdBase.Execute cSQL, dbFailOnError

Exit Sub
LOKAL_ERROR:
    If err.Number = 3078 Or err.Number = 3010 Or err.Number = 58 Or err.Number = 53 Or err.Number = 3043 Or err.Number = 91 Or err.Number = 3011 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "ErzeugeSendok"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Public Sub ErzeugeBESTOUT()
    On Error GoTo LOKAL_ERROR

    Dim cSQL        As String
    Dim cPfad1      As String
    
    cPfad1 = gcDBPfad
    If Right(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If
   
    loeschNEW "BESTOUT", gdBase
    
    cSQL = "Select * into BESTOUT from BESTAEND where"
    cSQL = cSQL & " SENDOK = False"
    gdBase.Execute cSQL, dbFailOnError

    cSQL = "update BESTAEND set SENDOK = True where SENDOK = False"
    gdBase.Execute cSQL, dbFailOnError

Exit Sub
LOKAL_ERROR:
    If err.Number = 3078 Or err.Number = 3010 Or err.Number = 58 Or err.Number = 53 Or err.Number = 3043 Or err.Number = 91 Or err.Number = 3011 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "ErzeugeBESTOUT"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Public Sub ErzeugeBESTPOUT()
    On Error GoTo LOKAL_ERROR

    Dim cSQL        As String
    Dim cPfad1      As String
    
    cPfad1 = gcDBPfad
    If Right(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If
   
    loeschNEW "BESTPOUT", gdBase
    
    cSQL = "Select * into BESTPOUT from BESTPROT where"
    cSQL = cSQL & " SENDOK = False"
    gdBase.Execute cSQL, dbFailOnError

    cSQL = "update BESTPROT set SENDOK = True where SENDOK = False"
    gdBase.Execute cSQL, dbFailOnError

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ErzeugeBESTPOUT"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub ErzeugeRetoureOUT()
    On Error GoTo LOKAL_ERROR

    Dim cSQL        As String
    Dim cPfad1      As String
    
    cPfad1 = gcDBPfad
    If Right(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If
    
    loeschNEW "RETOUT", gdBase
    
    cSQL = "Select * into RETOUT from Retoure where"
    cSQL = cSQL & " SENDOK = False"
    gdBase.Execute cSQL, dbFailOnError

    cSQL = "update Retoure set SENDOK = True where SENDOK = False"
    gdBase.Execute cSQL, dbFailOnError

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ErzeugeRetoureOUT"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub ErzeugeKBOUT()
    On Error GoTo LOKAL_ERROR

    Dim cSQL        As String
    Dim cPfad1      As String
    
    cPfad1 = gcDBPfad
    If Right(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If
   
    loeschNEW "KBOUT", gdBase
    
    cSQL = "Select * into KBOUT from KUNDBEST where"
    cSQL = cSQL & " SENDOK = False"
    gdBase.Execute cSQL, dbFailOnError

    cSQL = "update KUNDBEST set SENDOK = True where SENDOK = False"
    gdBase.Execute cSQL, dbFailOnError

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ErzeugeKBOUT"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub ErzeugeXOUT(sTab As String)
    On Error GoTo LOKAL_ERROR

    Dim cSQL        As String
    Dim cPfad1      As String
    
    cPfad1 = gcDBPfad
    If Right(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If
   
    loeschNEW sTab & "O", gdBase
    
    cSQL = "Select * into " & sTab & "O from " & sTab & " where"
    cSQL = cSQL & " SENDOK = False"
    gdBase.Execute cSQL, dbFailOnError

    cSQL = "update " & sTab & " set SENDOK = True where SENDOK = False"
    gdBase.Execute cSQL, dbFailOnError

Exit Sub
LOKAL_ERROR:
    If err.Number = 3078 Or err.Number = 3010 Or err.Number = 58 Or err.Number = 53 Or err.Number = 3043 Or err.Number = 91 Or err.Number = 3011 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "ErzeugeXOUT"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Public Sub ErzeugeKKZahlte_OUT()
    On Error GoTo LOKAL_ERROR

    Dim cSQL        As String
    Dim cPfad1      As String
    
    cPfad1 = gcDBPfad
    If Right(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If
   
    loeschNEW "KKZAHLTEO", gdBase
    
    cSQL = "Select * into KKZAHLTEO from KKZAHL where"
    cSQL = cSQL & " SENDOK = False"
    gdBase.Execute cSQL, dbFailOnError

    cSQL = "update KKZAHL set SENDOK = True where SENDOK = False"
    gdBase.Execute cSQL, dbFailOnError

Exit Sub
LOKAL_ERROR:
    If err.Number = 3078 Or err.Number = 3010 Or err.Number = 58 Or err.Number = 53 Or err.Number = 3043 Or err.Number = 91 Or err.Number = 3011 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "ErzeugeKKZahlte_OUT"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Public Sub ErzeugeKVKPR1POUT()
    On Error GoTo LOKAL_ERROR

    Dim cSQL        As String
    Dim cPfad1      As String
    
    cPfad1 = gcDBPfad
    If Right(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If
   
    loeschNEW "KVKPR1POUT", gdBase
    
    cSQL = "Select * into KVKPR1POUT from KVKPR1PROT where"
    cSQL = cSQL & " SENDOK = False"
    schreibeProtokollDabaAblauf cSQL: gdBase.Execute cSQL, dbFailOnError

    
    cSQL = "update KVKPR1PROT set SENDOK = True where SENDOK = False"
    schreibeProtokollDabaAblauf cSQL: gdBase.Execute cSQL, dbFailOnError

Exit Sub
LOKAL_ERROR:
    If err.Number = 3078 Or err.Number = 3010 Or err.Number = 58 Or err.Number = 53 Or err.Number = 3043 Or err.Number = 91 Or err.Number = 3011 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "ErzeugeEANPOUT"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Public Sub ErzeugeEANPOUT()
    On Error GoTo LOKAL_ERROR

    Dim cSQL        As String
    Dim cPfad1      As String
    
    cPfad1 = gcDBPfad
    If Right(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If
   
    loeschNEW "EANPOUT", gdBase
    
    cSQL = "Select * into EANPOUT from EANPROT where"
    cSQL = cSQL & " SENDOK = False"
    schreibeProtokollDabaAblauf cSQL: gdBase.Execute cSQL, dbFailOnError

    
    cSQL = "update EANPROT set SENDOK = True where SENDOK = False"
    schreibeProtokollDabaAblauf cSQL: gdBase.Execute cSQL, dbFailOnError

Exit Sub
LOKAL_ERROR:
    If err.Number = 3078 Or err.Number = 3010 Or err.Number = 58 Or err.Number = 53 Or err.Number = 3043 Or err.Number = 91 Or err.Number = 3011 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "ErzeugeEANPOUT"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Public Sub ErzeugeSTE_OUT()
    On Error GoTo LOKAL_ERROR

    Dim cSQL        As String
    Dim cPfad1      As String
    
    cPfad1 = gcDBPfad
    If Right(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If
   
    loeschNEW "STE_OUT", gdBase
    
    cSQL = "Select * into STE_OUT from STEMPEL where"
    cSQL = cSQL & " SENDOK = False"
    gdBase.Execute cSQL, dbFailOnError

    
    cSQL = "update STEMPEL set SENDOK = True where SENDOK = False"
    gdBase.Execute cSQL, dbFailOnError

Exit Sub
LOKAL_ERROR:
    If err.Number = 3078 Or err.Number = 3010 Or err.Number = 58 Or err.Number = 53 Or err.Number = 3043 Or err.Number = 91 Or err.Number = 3011 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "ErzeugeSTE_OUT"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Public Sub ErzeugeKolout()
    On Error GoTo LOKAL_ERROR

    Dim cSQL        As String
    Dim cDatname    As String
    Dim cPfad1      As String

    cPfad1 = gcDBPfad
    If Right(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If
    cDatname = "KOLOUT" & Trim$(gcFilNr)

    loeschNEW "KOLOUT", gdBase
    
    cSQL = "Select * into KOLOUT from KOLLVERK where"
    cSQL = cSQL & " SENDOK = False"
    gdBase.Execute cSQL, dbFailOnError

    
    cSQL = "update KOLLVERK set SENDOK = True where SENDOK = False"
    gdBase.Execute cSQL, dbFailOnError

Exit Sub
LOKAL_ERROR:
    If err.Number = 3078 Or err.Number = 3010 Or err.Number = 58 Or err.Number = 53 Or err.Number = 3043 Or err.Number = 91 Or err.Number = 3011 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "ErzeugeKolout"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Public Sub ErzeugeZ_out()
    On Error GoTo LOKAL_ERROR

    Dim cSQL        As String
    Dim cDatname    As String
    Dim cPfad1      As String
    
    cPfad1 = gcDBPfad
    If Right(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If
    cDatname = "Z_OUT" & Trim$(gcFilNr)
    
    loeschNEW "Z_OUT", gdBase
    CreateTable "Z_OUT", gdBase
    
    cSQL = "insert into Z_out Select ARTNR,BESTAND,MINBEST,KVKPR1 from ARTIKEL where"
    cSQL = cSQL & " BESTAND <> 0 or MINBEST > 0 or KVKPR1 > 0"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "update Z_out Set MINBEST = 0 where"
    cSQL = cSQL & " MINBEST is null "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "update Z_out Set BESTAND = 0 where"
    cSQL = cSQL & " BESTAND is null "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "update Z_out Set KVKPR1 = 0 where"
    cSQL = cSQL & " KVKPR1 is null "
    gdBase.Execute cSQL, dbFailOnError
    

Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ErzeugeZ_out"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
        
End Sub
Public Sub speicherpfad()
    On Error GoTo LOKAL_ERROR
    
    Dim rsrs As Recordset

    Set rsrs = gdApp.OpenRecordset("WKEINSTE", dbOpenTable)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        rsrs.Edit
        rsrs!UpdPfad = gsUpdPfad
        rsrs!ZinPfad = gsZinPfad
        rsrs!KinPfad = gsKinPfad
        rsrs!DabaPfad = gsDabaPfad
        rsrs!DTAPfad = gsDTAPfad
        rsrs!ZOUTPFAD = gsZoutPfad
        rsrs!SichPfad = gsSicherPfad
        rsrs!FotoPfad = gsFotoPfad
        rsrs!WebcamPfad = gsWebcamPfad
        rsrs.Update
        gbLokalModus = False
    End If
    
    rsrs.Close: Set rsrs = Nothing
    
    Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "speicherpfad"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Public Sub speicherSicherungpfad()
    On Error GoTo LOKAL_ERROR
    
    Dim rsrs As Recordset

    Set rsrs = gdApp.OpenRecordset("WKEINSTE", dbOpenTable)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        rsrs.Edit
        rsrs!SichPfad = gsSicherPfad
        rsrs.Update
    End If
    
    rsrs.Close: Set rsrs = Nothing
    
    Exit Sub
LOKAL_ERROR:
    
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "speicherSicherungpfad"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1
        

End Sub
Public Function pfadaendern(sTitle As String, sFilter As String, sOldpfad As String) As String
    On Error GoTo LOKAL_ERROR
    
    Dim sPfad   As String
    
    pfadaendern = sOldpfad
    
    With frmWKL00.cdlopen
        .CancelError = True
        On Error GoTo err
        .InitDir = sOldpfad
        .DialogTitle = sTitle
        .Filter = sFilter
        .ShowSave
    
  
        sPfad = Left(.FileName, Len(.FileName) - (Len(.FileTitle) + 1))
    End With
    pfadaendern = sPfad
err:
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "pfadaendern"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Function
Public Function pfadaendernKomplett(sTitle As String, sFilter As String, sOldpfad As String) As String
    On Error GoTo LOKAL_ERROR
    
    Dim sPfad   As String
    
    pfadaendernKomplett = sOldpfad
    
    With frmWKL00.cdlopen
        .CancelError = True
        On Error GoTo err
        .InitDir = sOldpfad
        .DialogTitle = sTitle
        .Filter = sFilter
        .ShowSave
    
  
        sPfad = .FileName
    End With
    pfadaendernKomplett = sPfad
    
    
err:
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "pfadaendernKomplett"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Function
Public Sub datumschreiben(schluessel As String, Wert As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
        
    sSQL = "Insert into Proto (Schluessel,Wert,Datum)"
    sSQL = sSQL & " Values ( "
    sSQL = sSQL & " '" & schluessel & "' "
    sSQL = sSQL & ", '" & Wert & "' "
    sSQL = sSQL & ", '" & DateValue(Now) & "' "
    sSQL = sSQL & " ) "
    gdBase.Execute sSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "datumschreiben"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Function ermBESTAND(cArtNr) As Long
On Error GoTo LOKAL_ERROR
    
    Dim rsArt As Recordset
    
    If cArtNr = "" Then
        Exit Function
    End If
    
    ermBESTAND = 0

    Set rsArt = gdBase.OpenRecordset("Select Bestand from Artikel where artnr = " & cArtNr)
    If Not rsArt.EOF Then
    
        If Not IsNull(rsArt!BESTAND) Then
            ermBESTAND = Val(rsArt!BESTAND)
        End If
    
    End If
    rsArt.Close: Set rsArt = Nothing

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ermBestand"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Sub lfnrSchreiben(lfnr As Long, Datname As String, Datum As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
        
    sSQL = "Insert into STEUERKI (lfnr,DATNAME,DATUM)"
    sSQL = sSQL & " Values ( "
    sSQL = sSQL & " " & lfnr & " "
    sSQL = sSQL & ", '" & Datname & "' "
    sSQL = sSQL & ", '" & Datum & "' "
    sSQL = sSQL & " ) "
    gdBase.Execute sSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "lfnrSchreiben"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Function ermNachrichtenMessage(lNaNr As Long) As String
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rs As DAO.Recordset
    
    ermNachrichtenMessage = ""
    
    sSQL = "Select Messagetext from NACHRICHTEN where lfnr = " & lNaNr & "  "
    Set rs = gdBase.OpenRecordset(sSQL)
    If Not rs.EOF Then
        If Not IsNull(rs.Fields(0)) Then
            ermNachrichtenMessage = rs.Fields(0)
        End If
    End If
    rs.Close: Set rs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ermNachrichtenMessage"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Sub Nachrichten_aufGelesen(lNaNr As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    
    sSQL = "Update NACHRICHTEN set gelesen = True where lfnr = " & lNaNr & "  "
    gdBase.Execute sSQL, dbFailOnError
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "Nachrichten_aufGelesen"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Function ermMaxNachrichtenNummer() As Long
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rs As DAO.Recordset
    
    ermMaxNachrichtenNummer = -1
    
    sSQL = "Select max(lfnr) from NACHRICHTEN  "
    Set rs = gdBase.OpenRecordset(sSQL)
    If Not rs.EOF Then
        If Not IsNull(rs.Fields(0)) Then
            ermMaxNachrichtenNummer = rs.Fields(0)
        End If
    End If
    rs.Close: Set rs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ermMaxNachrichtenNummer"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function lfnrErmitteln(Unterschied As String) As Integer
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rs As Recordset
    
    lfnrErmitteln = 0
    
    If Unterschied = "F" Then
    
        sSQL = "Select max(lfnr) from STEUERKI where Datname like  'F*' "
        Set rs = gdBase.OpenRecordset(sSQL)
        If rs.EOF Then
            lfnrErmitteln = 0
        Else
        
            If Not IsNull(rs.Fields(0)) Then
                lfnrErmitteln = rs.Fields(0)
            Else
                lfnrErmitteln = 0
            End If
            
        End If
        rs.Close: Set rs = Nothing
        
    ElseIf Unterschied = "Y" Then
    
        sSQL = "Select max(lfnr) from STEUERKI where Datname like  'Y*' "
        Set rs = gdBase.OpenRecordset(sSQL)
        If rs.EOF Then
            lfnrErmitteln = 0
        Else
        
            If Not IsNull(rs.Fields(0)) Then
                lfnrErmitteln = rs.Fields(0)
            Else
                lfnrErmitteln = 0
            End If
            
        End If
        rs.Close: Set rs = Nothing
    End If
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "lfnrErmitteln"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function Gutscheinpruef(sGutschnr As String) As String
On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rs As Recordset
    
    Gutscheinpruef = "keinen Gutschein gefunden"
    
    sGutschnr = SwapStr(sGutschnr, ",", "")
    If Val(sGutschnr) = 0 Then
        Exit Function
    End If
    
    
    
    sSQL = "Select * from Gutsch where GUTSCHNR = " & sGutschnr
    Set rs = gdBase.OpenRecordset(sSQL)
    If Not rs.EOF Then
    
        If Not IsNull(rs!gutschnr) Then
            Gutscheinpruef = "Gutscheinnummer: " & rs!gutschnr & vbCrLf
        End If
    
        If Not IsNull(rs!Wert) Then
            Gutscheinpruef = Gutscheinpruef & "Wert: " & Format(rs!Wert, "#####0.00 " & gcWaehrung) & vbCrLf
        End If
        
        If Not IsNull(rs!DAT_AUSG) Then
            Gutscheinpruef = Gutscheinpruef & "Ausgabe am: " & Format$(rs!DAT_AUSG, "DD.MM.YYYY") & vbCrLf
        End If
        
        If Not IsNull(rs!DAT_EINL) Then
            Gutscheinpruef = Gutscheinpruef & "Eingelöst am: " & Format$(rs!DAT_EINL, "DD.MM.YYYY") & vbCrLf
        Else
            Gutscheinpruef = Gutscheinpruef & "Eingelöst am: noch nicht"
        End If
        
    End If
    rs.Close: Set rs = Nothing
    
    If Gutscheinpruef = "keinen Gutschein gefunden" Then
    
        If Len(sGutschnr) = 8 And Left(sGutschnr, 1) Then sGutschnr = Mid(sGutschnr, 2, 6)
    
        sSQL = "Select * from Gutsch where GUTSCHNR = " & sGutschnr
        Set rs = gdBase.OpenRecordset(sSQL)
        If Not rs.EOF Then
        
            If Not IsNull(rs!gutschnr) Then
                Gutscheinpruef = "Gutscheinnummer: " & rs!gutschnr & vbCrLf
            End If
        
            If Not IsNull(rs!Wert) Then
                Gutscheinpruef = Gutscheinpruef & "Wert: " & Format(rs!Wert, "#####0.00 " & gcWaehrung) & vbCrLf
            End If
            
            If Not IsNull(rs!DAT_AUSG) Then
                Gutscheinpruef = Gutscheinpruef & "Ausgabe am: " & Format$(rs!DAT_AUSG, "DD.MM.YYYY") & vbCrLf
            End If
            
            If Not IsNull(rs!DAT_EINL) Then
                Gutscheinpruef = Gutscheinpruef & "Eingelöst am: " & Format$(rs!DAT_EINL, "DD.MM.YYYY") & vbCrLf
            Else
                Gutscheinpruef = Gutscheinpruef & "Eingelöst am: noch nicht"
            End If
            
        End If
        rs.Close: Set rs = Nothing
        
    End If
    
    
    If Gutscheinpruef = "keinen Gutschein gefunden" Then
    
        If Len(sGutschnr) = 13 And Left(sGutschnr, 2) = "21" Then sGutschnr = Val(Mid(sGutschnr, 3, 10))
    
        sSQL = "Select * from Gutsch where GUTSCHNR = " & sGutschnr
        Set rs = gdBase.OpenRecordset(sSQL)
        If Not rs.EOF Then
        
            If Not IsNull(rs!gutschnr) Then
                Gutscheinpruef = "Gutscheinnummer: " & rs!gutschnr & vbCrLf
            End If
        
            If Not IsNull(rs!Wert) Then
                Gutscheinpruef = Gutscheinpruef & "Wert: " & Format(rs!Wert, "#####0.00 " & gcWaehrung) & vbCrLf
            End If
            
            If Not IsNull(rs!DAT_AUSG) Then
                Gutscheinpruef = Gutscheinpruef & "Ausgabe am: " & Format$(rs!DAT_AUSG, "DD.MM.YYYY") & vbCrLf
            End If
            
            If Not IsNull(rs!DAT_EINL) Then
                Gutscheinpruef = Gutscheinpruef & "Eingelöst am: " & Format$(rs!DAT_EINL, "DD.MM.YYYY") & vbCrLf
            Else
                Gutscheinpruef = Gutscheinpruef & "Eingelöst am: noch nicht"
            End If
            
        End If
        rs.Close: Set rs = Nothing
        
    End If
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "Gutscheinpruef"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function Gutscheinpruef_KL_SQL(sGutschnr As String) As String
On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rs As Recordset
    
    Gutscheinpruef_KL_SQL = "keinen Gutschein gefunden"
    
    sGutschnr = SwapStr(sGutschnr, ",", "")
    If Val(sGutschnr) = 0 Then
        Exit Function
    End If
    
    If fTestLogin_SQLDABA_Error = 0 Then 'ist alles OK? Datenbank erreichbar?
        'alles okay
    Else
        schreibeProtokollVPNTXT "Unterbrechung"
        
        Dim sTemp As String
        sTemp = "Bitte starten Sie diesen Rechner neu" & vbCrLf
        sTemp = sTemp & "oder schließen Sie das Schloss und starten Sie WinKiss neu."
    
        MsgBox sTemp, vbCritical + vbOKOnly, "Gutschein-Datenbank nicht erreichbar"
        Exit Function
    End If
    
    
    Dim stConnect As String

    If gsKL_DSN <> "" Then
        stConnect = "ODBC;DSN=" & gsKL_DSN & ";UID=" & gsKL_BENUTZER & ";PWD=" & gsKL_PASSWORT & ""
    Else
        stConnect = "ODBC;DRIVER=SQL Server;SERVER=" & gsKL_ADRESSE & ";DATABASE=" & gsKL_DATENBANKNAME & ";UID=" & gsKL_BENUTZER & ";PWD=" & gsKL_PASSWORT & ""
    End If
    
    Dim dbEAN As DAO.Database
    Set dbEAN = OpenDatabase(gsKL_DATENBANKNAME, dbDriverNoPrompt, False, stConnect)
    
    
    sSQL = "Select * from GUTSCHEINE where GUTSCHNR = '" & sGutschnr & "'"
    Set rs = dbEAN.OpenRecordset(sSQL)
    If Not rs.EOF Then
    
        If Not IsNull(rs!gutschnr) Then
            Gutscheinpruef_KL_SQL = "Gutscheinnummer: " & rs!gutschnr & vbCrLf
        End If
    
        If Not IsNull(rs!Wert) Then
            Gutscheinpruef_KL_SQL = Gutscheinpruef_KL_SQL & "Wert: " & Format(rs!Wert, "#####0.00 " & gcWaehrung) & vbCrLf
        End If
        
        If Not IsNull(rs!AUSG_DATUM) Then
            Gutscheinpruef_KL_SQL = Gutscheinpruef_KL_SQL & "Ausgabe am: " & Format$(rs!AUSG_DATUM, "DD.MM.YYYY") & vbCrLf
        End If
        
        If Not IsNull(rs!EINL_DATUM) Then
            Gutscheinpruef_KL_SQL = Gutscheinpruef_KL_SQL & "Eingelöst am: " & Format$(rs!EINL_DATUM, "DD.MM.YYYY") & vbCrLf
        Else
            Gutscheinpruef_KL_SQL = Gutscheinpruef_KL_SQL & "Eingelöst am: noch nicht"
        End If
        
    End If
    rs.Close: Set rs = Nothing
    
    If Gutscheinpruef_KL_SQL = "keinen Gutschein gefunden" Then
    
        If Len(sGutschnr) = 8 And Left(sGutschnr, 1) Then sGutschnr = Mid(sGutschnr, 2, 6)
    
        sSQL = "Select * from GUTSCHEINE where GUTSCHNR = '" & sGutschnr & "'"
        Set rs = dbEAN.OpenRecordset(sSQL)
        If Not rs.EOF Then
        
            If Not IsNull(rs!gutschnr) Then
                Gutscheinpruef_KL_SQL = "Gutscheinnummer: " & rs!gutschnr & vbCrLf
            End If
        
            If Not IsNull(rs!Wert) Then
                Gutscheinpruef_KL_SQL = Gutscheinpruef_KL_SQL & "Wert: " & Format(rs!Wert, "#####0.00 " & gcWaehrung) & vbCrLf
            End If
            
            If Not IsNull(rs!AUSG_DATUM) Then
                Gutscheinpruef_KL_SQL = Gutscheinpruef_KL_SQL & "Ausgabe am: " & Format$(rs!AUSG_DATUM, "DD.MM.YYYY") & vbCrLf
            End If
            
            If Not IsNull(rs!EINL_DATUM) Then
                Gutscheinpruef_KL_SQL = Gutscheinpruef_KL_SQL & "Eingelöst am: " & Format$(rs!EINL_DATUM, "DD.MM.YYYY") & vbCrLf
            Else
                Gutscheinpruef_KL_SQL = Gutscheinpruef_KL_SQL & "Eingelöst am: noch nicht"
            End If
            
        End If
        rs.Close: Set rs = Nothing
        
    End If
    
    
    
    
    
    If Gutscheinpruef_KL_SQL = "keinen Gutschein gefunden" Then
    
        If Len(sGutschnr) = 13 And Left(sGutschnr, 2) = "21" Then sGutschnr = Val(Mid(sGutschnr, 3, 10))
    
        sSQL = "Select * from GUTSCHEINE where GUTSCHNR = '" & sGutschnr & "'"
        Set rs = dbEAN.OpenRecordset(sSQL)
        If Not rs.EOF Then
        
            If Not IsNull(rs!gutschnr) Then
                Gutscheinpruef_KL_SQL = "Gutscheinnummer: " & rs!gutschnr & vbCrLf
            End If
        
            If Not IsNull(rs!Wert) Then
                Gutscheinpruef_KL_SQL = Gutscheinpruef_KL_SQL & "Wert: " & Format(rs!Wert, "#####0.00 " & gcWaehrung) & vbCrLf
            End If
            
            If Not IsNull(rs!AUSG_DATUM) Then
                Gutscheinpruef_KL_SQL = Gutscheinpruef_KL_SQL & "Ausgabe am: " & Format$(rs!AUSG_DATUM, "DD.MM.YYYY") & vbCrLf
            End If
            
            If Not IsNull(rs!EINL_DATUM) Then
                Gutscheinpruef_KL_SQL = Gutscheinpruef_KL_SQL & "Eingelöst am: " & Format$(rs!EINL_DATUM, "DD.MM.YYYY") & vbCrLf
            Else
                Gutscheinpruef_KL_SQL = Gutscheinpruef_KL_SQL & "Eingelöst am: noch nicht"
            End If
            
        End If
        rs.Close: Set rs = Nothing
        
    End If
    
    dbEAN.Close
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "Gutscheinpruef_KL_SQL"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ifThisDatinSteuerki(cdat As String) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rs As Recordset
    
    ifThisDatinSteuerki = False
    
    sSQL = "Select * from STEUERKI where Datname = '" & Left(cdat, 8) & "'"
    Set rs = gdBase.OpenRecordset(sSQL)
    If Not rs.EOF Then
        ifThisDatinSteuerki = True
    End If
    rs.Close: Set rs = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ifThisDatinSteuerki"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Sub BESTAKTweg(txtStatus As TextBox)
    On Error GoTo LOKAL_ERROR

    Dim cSQL        As String
    Dim rsrs        As Recordset
    Dim lHeute      As Long
    Dim cPfad       As String
    Dim cDatum      As String
    Dim Fdb         As Database
    
    cPfad = gcDBPfad        'Datenbankpfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    Kill cPfad & "EXPORT\FZ.mdb"
    Set Fdb = CreateDatabase(cPfad & "EXPORT\FZ.mdb", dbLangGeneral, dbVersion40)
    Fdb.Close
    
    cDatum = DateValue(Now)
    
    ErzeugeNeueKassenDateiBESTAKT txtStatus

    If gbErfolg = False Then
        Exit Sub
    End If
    
    lHeute = Fix(Now)
    cDatum = Format$(lHeute, "DD.MM.YYYY")

    cSQL = "Select * from LASTSEND"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.Edit
    Else
        rsrs.AddNew
    End If
    rsrs!FILIALE = Val(gcFilNr)
    rsrs!Datum = DateValue(Now)
    rsrs.Update
    rsrs.Close: Set rsrs = Nothing
                            
Exit Sub
LOKAL_ERROR:

    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "BESTAKTweg"
        Fehler.gsFehlertext = "Im Programmteil Tagesabschluss ist ein Fehler aufgetreten."

        Fehlermeldung1
'        Resume Next
    End If
End Sub
Public Sub ExternAbholen(lab As Label, txtStatus As TextBox, labglo As Label)
On Error GoTo LOKAL_ERROR

    Dim cPfad       As String
    Dim cPfad2      As String
    Dim cpfadZiel   As String
    Dim bmerke      As Boolean
    Dim cZiel       As String
    Dim cQuelle     As String
    Dim cZiel1      As String
    Dim lRet        As Long
    Dim lfail       As Long
    bmerke = gbFTPautomatic
   
    schreibeProtokollDaba ("externe Sicherung wird abgeholt")
    
    gsZenFTPAdresse = "80.86.85.121" '"85.25.132.45"
    gsZenFTPUSER = gsLagerFTPBox 'Das ist der Eintrag aus der Lager.cfg
    gsZenFTPPASS = "stada"

    gbFTPautomatic = True
    giKissFtpMode = 22
    frmWKL38.Show 1
    gbFTPautomatic = bmerke
    
    cPfad = gcDBPfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cpfadZiel = cPfad & "ENDZIPIN\"
    
    Screen.MousePointer = 11
    
    'ist die end.mdb da?
    If FileExists(cpfadZiel & "end.lzh") Then
    
'        MsgBox "Stopp 1"
    
        labglo.ForeColor = vbRed
        labglo.Caption = "Datenbank wird entpackt, bitte warten..."
        labglo.Refresh
    
        schreibeProtokollDaba ("externe Sicherung wird entpackt")
        
        'Alteinstellungen sichern
        
        
        
        loeschNEW "DBEINSTE", gdApp
        TransferTab gdBase, App.Path & "\kissapp.mdb", "DBEINSTE"
        
        loeschNEW "BEDZUGRI", gdApp
        TransferTab gdBase, App.Path & "\kissapp.mdb", "BEDZUGRI"
        
        loeschNEW "STAMMFTP", gdApp
        TransferTab gdBase, App.Path & "\kissapp.mdb", "STAMMFTP"
        
        loeschNEW "STEUERKI", gdApp
        TransferTab gdBase, App.Path & "\kissapp.mdb", "STEUERKI"
        
        loeschNEW "KASSEIN", gdApp
        TransferTab gdBase, App.Path & "\kissapp.mdb", "KASSEIN"
        
        loeschNEW "BESTAKT", gdApp
        TransferTab gdBase, App.Path & "\kissapp.mdb", "BESTAKT"
        
        loeschNEW "Zugang", gdApp
        TransferTab gdBase, App.Path & "\kissapp.mdb", "Zugang"
        
        loeschNEW "TABLAY" & srechnertab, gdApp
        TransferTab gdBase, App.Path & "\kissapp.mdb", "TABLAY" & srechnertab
        
    
        'erst schließen dann sichern, dann löschen!
        
        'schließen
        Set gdBase = Nothing
        
        'sichern
        cPfad2 = gcDBPfad    'dabapfad
        If Right(cPfad2, 1) <> "\" Then
            cPfad2 = cPfad2 & "\"
        End If
        
        cQuelle = cPfad2
        cQuelle = ShortPath(cQuelle)
        cQuelle = cQuelle & "kissdata.mdb"
        
        cZiel1 = cPfad2
        cZiel1 = ShortPath(cZiel1)
        cZiel1 = cZiel1 & "kissSIC.mdb"
        Kill cZiel1 'erstmal löschen
        
        lRet = CopyFile(cQuelle, cZiel1, lfail)
        Pause (10)
        'löschen
        Kill cPfad & "kissdata.mdb"
        
        Pause (10)
        
        If FileExists(cPfad & "kissdata.mdb") = True Then
            schreibeProtokollDaba ("Übernahme der Datenbank gescheitert, Kissdata war noch im Zugriff ")
            Set gdBase = OpenDatabase(cPfad & "kissdata.mdb", False, False, "MS Access;PWD=" & gsPasswort)
            Exit Sub
        End If
        
        Zip_Unzip "", cPfad, cpfadZiel & "end.lzh", txtStatus
        Pause (10)
        Kill cpfadZiel & "end.lzh"
        Pause (10)
        
'        MsgBox "Stopp 2"
        
        If FileExists(cPfad & "kissdata.mdb") = True Then
            Set gdBase = OpenDatabase(cPfad & "kissdata.mdb", False, False, "MS Access;PWD=" & gsPasswort)
        Else
            
            schreibeProtokollDaba ("Übernahme der Datenbank gescheitert, Sicherung wird zurückgelesen ")
            'zurücksichern
            cPfad2 = gcDBPfad    'dabapfad
            If Right(cPfad2, 1) <> "\" Then
                cPfad2 = cPfad2 & "\"
            End If
            
            cQuelle = cPfad2
            cQuelle = ShortPath(cQuelle)
            cQuelle = cQuelle & "kissSIC.mdb"
            
            cZiel1 = cPfad2
            cZiel1 = ShortPath(cZiel1)
            cZiel1 = cZiel1 & "kissdata.mdb"
            Kill cZiel1 'erstmal löschen
            
            lRet = CopyFile(cQuelle, cZiel1, lfail)
            
            Pause (2)
            
            'wenn man die Sicherung benötigte, dann wird sie gelöscht
            Kill cQuelle
        
            Set gdBase = OpenDatabase(cPfad & "kissdata.mdb", False, False, "MS Access;PWD=" & gsPasswort)
        End If
        
        'gibt es jetzt ein Problem dann wird die sicherung zurückkopiert
        
        'Alteinstellungen zurücksichern
        
        cZiel = gcDBPfad
        If Right$(cZiel, 1) <> "\" Then
            cZiel = cZiel & "\"
        End If
    
        loeschNEW "DBEINSTE", gdBase
        TransferTab gdApp, cZiel & "Kissdata.mdb", "DBEINSTE"
        
        loeschNEW "BEDZUGRI", gdBase
        TransferTab gdApp, cZiel & "Kissdata.mdb", "BEDZUGRI"
        
        loeschNEW "STAMMFTP", gdBase
        TransferTab gdApp, cZiel & "Kissdata.mdb", "STAMMFTP"
        
        loeschNEW "STEUERKI", gdBase
        TransferTab gdApp, cZiel & "Kissdata.mdb", "STEUERKI"
        
        loeschNEW "KASSEIN", gdBase
        TransferTab gdApp, cZiel & "Kissdata.mdb", "KASSEIN"
        
        loeschNEW "BESTAKT", gdBase
        TransferTab gdApp, cZiel & "Kissdata.mdb", "BESTAKT"
        
        loeschNEW "Zugang", gdBase
        TransferTab gdApp, cZiel & "Kissdata.mdb", "Zugang"
        
        loeschNEW "TABLAY" & srechnertab, gdBase
        TransferTab gdApp, cZiel & "Kissdata.mdb", "TABLAY" & srechnertab
        
        labglo.ForeColor = vbBlack
        labglo.Caption = "Fertig"
        labglo.Refresh
        
    End If

    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    ElseIf err.Number = 70 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "ExternAbholen"
        Fehler.gsFehlertext = "Es trat ein Fehler auf. "
        
        Fehlermeldung1
    End If
End Sub
Public Sub ExternAbholenDABA(lab As Label, txtStatus As TextBox, labglo As Label)
On Error GoTo LOKAL_ERROR

    Dim cPfad       As String
    Dim cPfad2      As String
    Dim cpfadZiel   As String
    Dim bmerke      As Boolean
    Dim cZiel       As String
    Dim cQuelle     As String
    Dim cZiel1      As String
    Dim lRet        As Long
    Dim lfail       As Long
    bmerke = gbFTPautomatic
   
    schreibeProtokollDaba ("externe Sicherung wird abgeholt")
    
    
    gsZenFTPAdresse = gsStammFTPAdresse
    gsZenFTPUSER = gsStammFTPUSER
    gsZenFTPPASS = gsStammFTPPASS
    
    gbFTPautomatic = True
    giKissFtpMode = 22
    frmWKL38.Show 1
    gbFTPautomatic = bmerke
    
    cPfad = gcDBPfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cpfadZiel = cPfad & "ENDZIPIN\"
    
    Screen.MousePointer = 11
    
    frmWKL00.Timer1.Enabled = False
    
''''    'ist die end.mdb da?
''''    If FileExists(cpfadZiel & "end.lzh") Then
''''
''''        labglo.ForeColor = vbRed
''''        labglo.Caption = "Datenbank wird entpackt, bitte warten..."
''''        labglo.Refresh
''''
''''        schreibeProtokollDaba ("externe Sicherung wird entpackt")
''''
''''        'dann entzippen
'''''        Zip_Unzip "geheim", cPfad, cpfadZiel & "end.lzh", txtStatus
''''
''''        Zip_Unzip "", cPfad, cpfadZiel & "end.lzh", txtStatus
''''
''''        Pause (2)
''''        Kill cpfadZiel & "end.lzh"
''''        Pause (2)
''''
''''
''''        Set gdBase = Nothing
''''
''''        If gbOhneAnzeige Then
''''            db_Copy cPfad, "K_sich.mdb", "KISSDATA.MDB", lab, txtStatus, labglo
''''        Else
''''            db_Copy cPfad, "END.MDB", "KISSDATA.MDB", lab, txtStatus, labglo
''''        End If
''''
''''
''''        Set gdBase = OpenDatabase(cPfad & "KISSDATA.MDB", True, False)
''''        gdBase.NewPassword "", gsPasswort
''''        gdBase.Close
''''
''''        'dann reindizieren
''''        Set gdBase = OpenDatabase(cPfad & "kissdata.mdb", False, False, "MS Access;PWD=" & gsPasswort)
''''        db_Reindizieren gdBase, lab, txtStatus, labglo
''''
''''        labglo.ForeColor = vbBlack
''''        labglo.Caption = "Fertig"
''''        labglo.Refresh
''''
''''    End If
    
    

    
    'ist die end.mdb da?
    If FileExists(cpfadZiel & "end.lzh") Then

        labglo.ForeColor = vbRed
        labglo.Caption = "Datenbank wird entpackt, bitte warten..."
        labglo.Refresh

        schreibeProtokollDaba ("externe Sicherung wird entpackt")
        
        Set gdBase = Nothing

        Zip_Unzip "", cPfad, cpfadZiel & "end.lzh", txtStatus
        Pause (2)
        Kill cpfadZiel & "end.lzh"
        Pause (2)

        Set gdBase = OpenDatabase(cPfad & "kissdata.mdb", False, False, "MS Access;PWD=" & gsPasswort)
        'gibt es jetzt ein Problem dann wird die sicherung zurückkopiert

        labglo.ForeColor = vbBlack
        labglo.Caption = "Fertig"
        labglo.Refresh

    End If
    
    
    
    If FileExists(App.Path & "\NoTimer.cfg") Then
        frmWKL00.Timer1.Enabled = False
    Else
        frmWKL00.Timer1.Enabled = True
    End If

    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    ElseIf err.Number = 70 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "ExternAbholenDABA"
        Fehler.gsFehlertext = "Es trat ein Fehler auf. "
        
        Fehlermeldung1
'        Resume Next
    End If
End Sub
'Public Sub ExternAbholen(lab As Label, txtstatus As TextBox, labglo As Label)
'On Error GoTo LOKAL_ERROR
'
'    Dim cPfad As String
'    Dim cpfadZiel As String
'
'    schreibeProtokollDaba ("externe Sicherung wird abgeholt")
'
'    cPfad = gcDBPfad
'    If Right$(cPfad, 1) <> "\" Then
'        cPfad = cPfad & "\"
'    End If
'
'    cpfadZiel = cPfad & "ENDZIPIN\"
'
'    'neu 23.10.08
'    Kill cpfadZiel & "end.lzh"
'
'    giKissFtpMode = 21
'    LiesStammFtp
'
'    gbFTPautomatic = True
'    frmWKL38.Show 1
'    gbFTPautomatic = False
'
'    Screen.MousePointer = 11
'
'    'ist die end.mdb da?
'    If FileExists(cpfadZiel & "end.lzh") Then
'
'        labglo.ForeColor = vbRed
'        labglo.Caption = "Datenbank wird entpackt, bitte warten..."
'        labglo.Refresh
'
'        schreibeProtokollDaba ("externe Sicherung wird entpackt")
'
'        'dann entzippen
'        Zip_Unzip "geheim", cPfad, cpfadZiel & "end.lzh", txtstatus
'
'
'        Set gdbMdb = Nothing
'
'        If gbOhneAnzeige Then
'
'            db_Copy cPfad, "K_sich.mdb", "KISSDATA.MDB", lab, txtstatus, labglo
'
'        Else
'            db_Copy cPfad, "END.MDB", "KISSDATA.MDB", lab, txtstatus, labglo
'        End If
'
'
'        Set gdbMdb = OpenDatabase(cPfad & "KISSDATA.MDB", True, False)
'        gdbMdb.NewPassword "", gsPasswort
'        gdbMdb.Close
'
'        'dann reindizieren
'        Set gdbMdb = OpenDatabase(cPfad & "kissdata.mdb", False, False, "MS Access;PWD=" & gsPasswort)
'        db_Reindizieren lab, txtstatus, labglo, gdbMdb
'
'        labglo.ForeColor = vbBlack
'        labglo.Caption = "Fertig"
'        labglo.Refresh
'
'    End If
'
'    Screen.MousePointer = 0
'
'Exit Sub
'LOKAL_ERROR:
'    If err.Number = 53 Then
'        Resume Next
'    Else
'        Fehler.gsDescr = err.Description
'        Fehler.gsNumber = err.Number
'        Fehler.gsFormular = "mdlDBBereinigen"
'        Fehler.gsFunktion = "ExternAbholen"
'        Fehler.gsFehlertext = "Es trat ein Fehler auf. "
'
'        Fehlermeldung1
'    End If
'End Sub
Public Sub ExternSichern(txtStatus As TextBox, labglo As Label)
On Error GoTo LOKAL_ERROR

    Dim cPfad As String
    Dim cpfadZiel As String
    
    cPfad = gcDBPfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cpfadZiel = cPfad & "ENDZIP\"
    
    Screen.MousePointer = 11
    
    'ist die end.mdb da?
    If FileExists(cPfad & "kissdata.mdb") Then
    
        labglo.ForeColor = vbRed
        labglo.Caption = "Datenbank wird gesichert, bitte warten..."
        labglo.Refresh
    
        schreibeProtokollDaba ("externe Sicherung gestartet")
    
        'dann zippen
        zipDllcheck
        Zip_Files "", cPfad & "kissdata.mdb", cpfadZiel & "end.lzh", txtStatus
        
        schreibeProtokollDaba ("externe Sicherung wird übertragen")
        'dann übertragen
        
        Dim bmerke As Boolean
        bmerke = gbFTPautomatic
        gbFTPautomatic = True
            
        giKissFtpMode = 21
        frmWKL38.Show 1
        
        gbFTPautomatic = bmerke
        
        labglo.ForeColor = vbBlack
        labglo.Caption = "Fertig"
        labglo.Refresh
        
    End If

    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ExternSichern"
    Fehler.gsFehlertext = "Es trat ein Fehler auf. "
    
    Fehlermeldung1
End Sub
Public Sub Pause(iDauer As Byte)
    On Error GoTo LOKAL_ERROR
    
    Dim lStart      As Long
    Dim lAktuell    As Long

    lStart = Timer
    Do
        lAktuell = Timer
    Loop While lAktuell < lStart + iDauer
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "Pause"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub PauseSi(iDauer As Single)
    On Error GoTo LOKAL_ERROR
    
    Dim lStart      As Single
    Dim lAktuell    As Single

    lStart = Timer
    Do
        lAktuell = Timer
    Loop While lAktuell < lStart + iDauer
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "PauseSi"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Function datumvergleichen(schluessel As String, Wert As String) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rs As Recordset
    
    datumvergleichen = False
    
    sSQL = "Select * from Proto where schluessel = "
    sSQL = sSQL & " '" & schluessel & "' "
    sSQL = sSQL & " and Wert = "
    sSQL = sSQL & " '" & Wert & "' "
    
    Set rs = gdBase.OpenRecordset(sSQL)
    If Not rs.RecordCount = 0 Then
        datumvergleichen = True
    Else
        datumvergleichen = False
    End If
    rs.Close: Set rs = Nothing
    
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "datumvergleichen"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function datum_aus_Proto(Wert As String) As String
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    
    datum_aus_Proto = ""
    
    sSQL = "Select Datum from ProtoEIN where DATNAME = "
    sSQL = sSQL & " '" & Wert & "' "
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!Datum) Then
            datum_aus_Proto = rsrs!Datum
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "datum_aus_Proto"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Sub dateiloeschen()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
   
    sSQL = "Delete from Proto where datum < datevalue(now)- 110 "
    gdBase.Execute sSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "dateiloeschen"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub Sortierung(bSort As Byte)
    On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11
    Dim sSQL As String
    
    sSQL = "Delete from SORTI "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into SORTI (SORT) Values ( " & bSort & " )"
    gdBase.Execute sSQL, dbFailOnError
    
    Screen.MousePointer = 0
    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "Sortierung"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub KUMSUM(bSort As Byte)
    On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11
    Dim sSQL As String
    
    sSQL = "Delete from KUMSUM "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into KUMSUM (SORT) Values ( " & bSort & " )"
    gdBase.Execute sSQL, dbFailOnError
    
    Screen.MousePointer = 0
    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "KUMSUM"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub LeseArtBestandinFil(List1 As Object, List2 As Object)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    
    Dim cLBSatz     As String
    Dim cFilnr      As String
    Dim cFilname    As String
    Dim ckPr        As String
    Dim cBestand    As String
    Dim cUW         As String
    Dim cUdat       As String
    Dim cSperr      As String
    Dim cBlock      As String
    Dim cMB         As String
    
    List1.Clear
    List2.Clear
    List1.AddItem "Nr Filialname      Bestand  Kassenpreis UW          MB"
    
    Set rsrs = gdBase.OpenRecordset("F" & srechnertab, dbOpenTable)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            cFilnr = IIf(IsNull(rsrs!FILIALNR), "-1", rsrs!FILIALNR)
            cFilname = IIf(IsNull(rsrs!FILIALNAME), "", rsrs!FILIALNAME)
            ckPr = IIf(IsNull(rsrs!KVKPR1), "0,00", Format(rsrs!KVKPR1, "#####0.00"))
            cBestand = IIf(IsNull(rsrs!BESTAND), "0", rsrs!BESTAND)
            cMB = IIf(IsNull(rsrs!MB), "0", rsrs!MB)
            cUW = IIf(IsNull(rsrs!unterwegs), " ", rsrs!unterwegs)
            cBlock = IIf(IsNull(rsrs!Block), "", rsrs!Block)
            cSperr = IIf(IsNull(rsrs!SPERR), "", rsrs!SPERR)
            
            cUdat = IIf(IsNull(rsrs!UDATE), "", rsrs!UDATE)
            
            cLBSatz = cFilnr & Space$(3 - Len(cFilnr))
            If Len(cFilname) > 15 Then
                cFilname = Left(cFilname, 15) & "..."
            End If
            cLBSatz = cLBSatz & cFilname
            cLBSatz = cLBSatz & Space$(18 - Len(cFilname))
            
            
            cLBSatz = cLBSatz & Space$(5 - Len(cBestand))
            cLBSatz = cLBSatz & cBestand
            
            cLBSatz = cLBSatz & Space$(7 - Len(ckPr))
            cLBSatz = cLBSatz & ckPr
            
            cLBSatz = cLBSatz & Space$(8 - Len(cUW))
            cLBSatz = cLBSatz & cUW
            
            cLBSatz = cLBSatz & Space$(4 - Len(cSperr))
            cLBSatz = cLBSatz & cSperr
            
            cLBSatz = cLBSatz & Space$(4 - Len(cBlock))
            cLBSatz = cLBSatz & cBlock
            
            cLBSatz = cLBSatz & Space$(4 - Len(cMB))
            cLBSatz = cLBSatz & cMB
            
            List2.AddItem cLBSatz
            
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing

Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "LeseArtBestandinFil"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub BestandinFiliale(sArtnr As String)
    On Error GoTo LOKAL_ERROR
    
    Dim i As Integer
    Dim sSQL As String
    
    If sArtnr = "" Then
        Exit Sub
    End If
    
    loeschNEW "F" & srechnertab, gdBase
    
    sSQL = "Create Table F" & srechnertab & " ( "
    sSQL = sSQL & " Filialnr BYTE"
    sSQL = sSQL & " , Filialname Text(35)"
    sSQL = sSQL & " , Bestand Integer "
    sSQL = sSQL & " , unterwegs Integer "
    sSQL = sSQL & " , uDATE DATETIME "
    sSQL = sSQL & " , Block Text(1)"
    sSQL = sSQL & " , Sperr Text(1)"
    sSQL = sSQL & " , MB Integer "
    sSQL = sSQL & " , KVKPR1 single )"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "insert into F" & srechnertab & " select filialnr,filialname from filialen"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "UPDATE F" & srechnertab & " INNER JOIN zbestand ON "
    sSQL = sSQL & " F" & srechnertab & ".filialnr = zbestand.filialnr "
    sSQL = sSQL & " Set F" & srechnertab & ".bestand = zbestand.bestand "
    sSQL = sSQL & " , F" & srechnertab & ".MB = zbestand.minbest "
    sSQL = sSQL & " , F" & srechnertab & ".KVKPR1 = zbestand.KVKPR1 "
    sSQL = sSQL & " where zbestand.artnr = " & sArtnr
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "UPDATE F" & srechnertab & ", artikel "
    sSQL = sSQL & " SET F" & srechnertab & ".bestand = artikel.bestand "
    sSQL = sSQL & " , F" & srechnertab & ".KVKPR1 = artikel.KVKPR1 "
    sSQL = sSQL & " where F" & srechnertab & ".filialnr = " & gcFilNr
    sSQL = sSQL & " and artikel.artnr = " & sArtnr
    gdBase.Execute sSQL, dbFailOnError
    
    If NewTableSuchenDBKombi("ZUNTER", gdBase) Then
        sSQL = "UPDATE F" & srechnertab & " INNER JOIN ZUNTER ON "
        sSQL = sSQL & " F" & srechnertab & ".filialnr = ZUNTER.filiale "
        sSQL = sSQL & " Set F" & srechnertab & ".unterwegs = ZUNTER.menge "
        sSQL = sSQL & " , F" & srechnertab & ".uDATE = ZUNTER.DATUM "
        sSQL = sSQL & " where ZUNTER.artnr = " & sArtnr
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If NewTableSuchenDBKombi("ZBLOCK", gdBase) Then
        sSQL = "UPDATE F" & srechnertab & " INNER JOIN ZBLOCK ON "
        sSQL = sSQL & " F" & srechnertab & ".filialnr = ZBLOCK.filiale "
        sSQL = sSQL & " Set F" & srechnertab & ".BLOCK = 'B' "
        sSQL = sSQL & " where ZBLOCK.artnr = " & sArtnr
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If NewTableSuchenDBKombi("ZSPERR", gdBase) Then
        sSQL = "UPDATE F" & srechnertab & " INNER JOIN ZSPERR ON "
        sSQL = sSQL & " F" & srechnertab & ".filialnr = ZSPERR.filnr "
        sSQL = sSQL & " Set F" & srechnertab & ".SPERR = 'S' "
        sSQL = sSQL & " where ZSPERR.artnr = " & sArtnr
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "BestandinFiliale"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub ABinBESTAKT(sArtnr As String, lMenge As Long, cAENART As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If gbBestAkt = True Then
    
        sSQL = "Insert into BESTAKT (artnr,menge,lastdate,lasttime,SENDOK,AENART) values  "
        sSQL = sSQL & " ( " & sArtnr & " , " & lMenge
        sSQL = sSQL & ", '" & DateValue(Now) & "'"
        sSQL = sSQL & ", '" & TimeValue(Now) & "'"
        sSQL = sSQL & ", " & 0
        sSQL = sSQL & ", '" & cAENART & "')"
        gdBase.Execute sSQL, dbFailOnError
    
    End If
    
    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ABinBESTAKT"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub ListeFuellAnfangsbuch(anfangsbuch As String, list As Object)
    On Error GoTo LOKAL_ERROR
    
    Dim lAnzTable   As Long
    Dim cSQL        As String
    Dim name        As String
    Dim cdatei      As String
    Dim cPfad       As String
    Dim lcount      As Long
    Dim cDatum      As String
    Dim sdat        As String
    Dim sLiefname   As String
    Dim rsrs        As Recordset
    Dim sLinr       As String
    Dim cAlias      As String
    Dim rec         As Recordset
     
    cPfad = gcPfad  'Applicationpfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
     
    list.Clear
    gdApp.TableDefs.Refresh
    lAnzTable = gdApp.TableDefs.Count
    For lcount = 0 To lAnzTable - 1
       
        name = gdApp.TableDefs(lcount).name
        If Left(name, 1) = anfangsbuch Then
             
            cdatei = UCase$(name)
            
            If Len(cdatei) > 7 Then
                sLinr = Mid$(cdatei, 2, 6)
            Else
                sLinr = Mid$(cdatei, 2, Len(cdatei) - 1)
            End If
            If IsNumeric(sLinr) Then
                Set rec = gdBase.OpenRecordset("Select LIEFBEZ from LISRT where LINR = " & sLinr)
                If Not rec.EOF Then
                    If Not IsNull(rec!LIEFBEZ) Then
                        If rec!LIEFBEZ <> "" Then
                            sLiefname = Space$(35 - Len(rec!LIEFBEZ)) & rec!LIEFBEZ
                        Else
                            sLiefname = Space$(35)
                        End If
                    Else
                        sLiefname = Space$(35)
                    End If
                Else
                    sLiefname = Space$(35)
                End If
                rec.Close: Set rec = Nothing
            Else
                sLiefname = Space$(35)
            End If

            
            
            cDatum = gdApp.TableDefs(lcount).DateCreated

            cAlias = ermXALIAS(cdatei)
            If cAlias = "" Then
                cAlias = Space$(21)
            Else
                cAlias = Space$(21 - Len(cAlias)) & cAlias
            End If
            
            cdatei = cdatei & Space(12 - Len(cdatei)) & sLiefname & "  " & cAlias & "       " & cDatum
            list.AddItem cdatei
            
             
         End If
    Next lcount
    
    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ListeFuellAnfangsbuch"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'        Resume Next
End Sub
Public Sub vorher_holen_X_BV()
On Error GoTo LOKAL_ERROR
    
    Dim lAnzTable   As Long
    Dim sName       As String
    Dim lcount      As Long
    Dim sSQL        As String
    Dim sZiel       As String
    
    sZiel = App.Path & "\kissapp.mdb"
    
    gdBase.TableDefs.Refresh
    lAnzTable = gdBase.TableDefs.Count
    For lcount = 0 To lAnzTable - 1
       
        sName = gdBase.TableDefs(lcount).name
        If Left(sName, 1) = "X" Then
            If Right(sName, 3) = "_BV" Then
            
            
                If NewTableSuchenDBKombi(sName, gdApp) = True Then
                
                    loeschNEW sName & "T", gdApp
                    sSQL = "Select " & sName & ".* into " & sName & "T IN '" & sZiel & "' from " & sName
                    gdBase.Execute sSQL, dbFailOnError
                    
                    sSQL = "Insert into " & sName & " Select * from " & sName & "T where not artnr IN (Select artnr from " & sName & " )"
                    gdApp.Execute sSQL, dbFailOnError
                    
                    loeschNEW sName & "T", gdApp
                    
                
                Else
                
                    loeschNEW sName, gdApp
                    TransferTab gdBase, App.Path & "\kissapp.mdb", sName
                End If
                
                loeschNEW sName, gdBase
            
            End If
            
        End If
    Next lcount
    
    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "vorher_holen_X_BV"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub delOld(anfangsbuch As String)
    On Error GoTo LOKAL_ERROR

    Dim lAnzTable   As Long
    Dim cName       As String
    Dim cPfad       As String
    Dim lcount      As Long
    Dim lDatum      As Long
    
   
     
    cPfad = gcPfad  'Applicationpfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
     
    
    gdApp.TableDefs.Refresh
    lAnzTable = gdApp.TableDefs.Count
    For lcount = 0 To lAnzTable - 1
        cName = gdApp.TableDefs(lcount).name
        If Left(cName, 1) = anfangsbuch Then
            lDatum = CLng(gdApp.TableDefs(lcount).DateCreated)
            
            If lDatum < DateValue(Now) - 20 Then
                loeschNEW cName, gdApp
            End If
        End If
    Next lcount
    
    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "delOld"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub NewListeFuellAnfangsbuch(anfangsbuch As String, list As Object, daba As Database)
    On Error GoTo LOKAL_ERROR
    
    Dim lAnzTable   As Long
    Dim cSQL        As String
    Dim name        As String
    Dim cdatei      As String
    Dim cPfad       As String
    Dim lcount      As Long
    Dim cDatum      As String
    Dim sdat        As String
    Dim sLiefname   As String
    Dim rsrs        As Recordset
    Dim lLief       As Long
    Dim LenAnfang   As Integer
    
    LenAnfang = Len(anfangsbuch)
    
    cPfad = gcPfad  'Applicationpfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
     
    list.Clear
    daba.TableDefs.Refresh
    lAnzTable = daba.TableDefs.Count
    For lcount = 0 To lAnzTable - 1
        name = daba.TableDefs(lcount).name
        If UCase(Left(name, LenAnfang)) = UCase(anfangsbuch) Then
             
            cdatei = UCase$(name)
            sdat = cdatei
            cDatum = ermfildat(daba.TableDefs(lcount).name)
            If cDatum = "" Then
                cDatum = daba.TableDefs(lcount).DateCreated
            End If
            cdatei = cdatei & Space$(12 - Len(cdatei)) & " " & cDatum
            list.AddItem cdatei
         End If
    Next lcount
    
    Exit Sub
LOKAL_ERROR:
    If err.Number = 3167 Then
        Resume Next
    Else

        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "NewListeFuellAnfangsbuch"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
        
End Sub
Public Sub New2ListeFuellAnfangsbuch(anfangsbuch As String, list As Object, daba As Database)
    On Error GoTo LOKAL_ERROR
    
    Dim lAnzTable   As Long
    Dim cSQL        As String
    Dim name        As String
    Dim cdatei      As String
    Dim cPfad       As String
    Dim lcount      As Long
    Dim cDatum      As String
    Dim sdat        As String
    Dim sLiefname   As String
    Dim rsrs        As Recordset
    Dim lLief       As Long
    Dim LenAnfang   As Integer
    
    LenAnfang = Len(anfangsbuch)
    
    cPfad = gcPfad  'Applicationpfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
     
    list.Clear
    daba.TableDefs.Refresh
    lAnzTable = daba.TableDefs.Count
    For lcount = 0 To lAnzTable - 1
        name = daba.TableDefs(lcount).name
        If UCase(Left(name, LenAnfang)) = UCase(anfangsbuch) Then
             
            cdatei = UCase$(name)
            sdat = cdatei
            cDatum = ermfildat(daba.TableDefs(lcount).name)
            
            If cDatum = "" Then
                cDatum = daba.TableDefs(lcount).DateCreated
            End If
            
            cdatei = Right(cdatei, Len(cdatei) - LenAnfang) & Space$(9 - Len(cdatei)) & " " & Left(cDatum, 8)
            
            list.AddItem cdatei
            
             
         End If
    Next lcount
    
    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "New2ListeFuellAnfangsbuch"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
       
End Sub

Public Function ermEmailAdress(sThema As String) As String
On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsrs As Recordset
    ermEmailAdress = ""
    
    cSQL = "Select Adresse from EMail where THEMA = '" & sThema & " '"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!adresse) Then
            ermEmailAdress = rsrs!adresse
        
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
                
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ermEmailAdress"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermLiefBez(lLief As Long) As String
On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsrs As Recordset
    ermLiefBez = ""
    
    cSQL = "Select LIEFBEZ from LISRT where LINR = " & lLief
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!LIEFBEZ) Then
            ermLiefBez = rsrs!LIEFBEZ
        
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
                
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ermLiefBez"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermDepotRabatt1_Lief(lLief As Long) As String
On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsrs As Recordset
    ermDepotRabatt1_Lief = "0"
    
    cSQL = "Select DepotRabatt1 from LISRT where LINR = " & lLief
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!DEPOTRABATT1) Then
            ermDepotRabatt1_Lief = rsrs!DEPOTRABATT1
        
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
                
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ermDepotRabatt1_Lief"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermLEK_ABSCHLAG_Lief(lLief As Long) As String
On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsrs As Recordset
    ermLEK_ABSCHLAG_Lief = "0"
    
    cSQL = "Select LEK_ABSCHLAG from LISRT where LINR = " & lLief
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!LEK_ABSCHLAG) Then
            ermLEK_ABSCHLAG_Lief = rsrs!LEK_ABSCHLAG
        
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
                
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ermLEK_ABSCHLAG_Lief"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermvonFirma(cWas As String) As String
On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsrs As Recordset
    ermvonFirma = ""
    
    cSQL = "Select " & cWas & " as Was from FIRMA "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!Was) Then
            ermvonFirma = rsrs!Was
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
                
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ermvonFirma"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermFirmenMail() As String
On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsrs As Recordset
    ermFirmenMail = ""
    
    cSQL = "Select Email from FIRMA "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!Email) Then
            ermFirmenMail = rsrs!Email
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
                
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ermFirmenMail"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermMailadresse(cAlias As String) As String
On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsrs As Recordset
    ermMailadresse = ""
    
    cSQL = "Select Adresse,lastdate from EMAIL where THema = 'STADA' and Alias = '" & cAlias & "'"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!adresse) Then
            ermMailadresse = rsrs!adresse
            
            rsrs.Edit
            rsrs!LASTDATE = DateValue(Now)
            rsrs.Update
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
                
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ermMailadresse"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermFirmenBez() As String
On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsrs As Recordset
    ermFirmenBez = ""
    
    cSQL = "Select name from FIRMA "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!name) Then
            ermFirmenBez = rsrs!name
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    ermFirmenBez = SwapStr(ermFirmenBez, "ü", "ue")
    ermFirmenBez = SwapStr(ermFirmenBez, "Ü", "UE")
    ermFirmenBez = SwapStr(ermFirmenBez, "ä", "ae")
    ermFirmenBez = SwapStr(ermFirmenBez, "Ä", "AE")
    ermFirmenBez = SwapStr(ermFirmenBez, "ö", "oe")
    ermFirmenBez = SwapStr(ermFirmenBez, "Ö", "OE")
    ermFirmenBez = SwapStr(ermFirmenBez, "ß", "ss")
                
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ermFirmenBez"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Sub ListeFuellAnfangsbuchdata(anfangsbuch As String, list As Object, b As Boolean)
    On Error GoTo LOKAL_ERROR
    
    Dim lAnzTable   As Long
    Dim cSQL        As String
    Dim name        As String
    Dim cdatei      As String
    Dim cPfad       As String
    Dim lcount      As Long
    Dim cDatum      As String
    Dim sdat        As String
    Dim sLiefname   As String
    Dim rsrs        As Recordset
    Dim lLief       As Long
          
    list.Clear
    gdBase.TableDefs.Refresh    'Dabarefresh
    lAnzTable = gdBase.TableDefs.Count
    For lcount = 0 To lAnzTable - 1
        name = gdBase.TableDefs(lcount).name
        If Left(name, 1) = anfangsbuch Then
             
            cdatei = UCase$(name)
            
            sdat = cdatei
            cDatum = ermfildat(gdBase.TableDefs(lcount).name)
            
            If cDatum = "" Then
                cDatum = gdBase.TableDefs(lcount).DateCreated
            End If
            
            
            cdatei = cdatei & Space$(12 - Len(cdatei)) & " " & cDatum
            If b Then
                sdat = Mid(sdat, 2, Len(sdat) - 2)
                
                lLief = Val(sdat)
                
                cSQL = "Select LIEFBEZ from LISRT where LINR = " & lLief
                
                
                Set rsrs = gdBase.OpenRecordset(cSQL)
                If Not rsrs.EOF Then
                    rsrs.MoveFirst
                    If Not IsNull(rsrs!LIEFBEZ) Then
                        sLiefname = rsrs!LIEFBEZ
                    Else
                        sLiefname = ""
                    End If
                Else
                    sLiefname = ""
                End If
                rsrs.Close: Set rsrs = Nothing
                
                list.AddItem cdatei & Space(4) & sLiefname
            Else
                list.AddItem cdatei
            End If
             
         End If
    Next lcount
    
    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ListeFuellAnfangsbuchdata"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
        
End Sub
Public Sub ListeFuellAnfangsbuchdataT(anfangsbuch As String, list As Object, corder As String, lblx As Label)
    On Error GoTo LOKAL_ERROR
    
    Dim lAnzTable   As Long
    Dim cSQL        As String
    Dim name        As String
    Dim cSatz       As String
    Dim cdatei      As String
    Dim cPfad       As String
    Dim lcount      As Long
    Dim cDatum      As String
    Dim sdat        As String
    Dim sLiefname   As String
    Dim cKurzinfo   As String
    Dim rsrs1       As Recordset
    Dim lLief       As Long
    Dim dAufwert    As Double
    Dim dSumAufwert As Double
    Dim lAufNr      As Long
    
    loeschNEW "TABANZEIGE", gdApp
    CreateTable "TABANZEIGE", gdApp
    
    loeschNEW "Tabdatum", gdApp
    TransferTab gdBase, App.Path & "\kissapp.mdb", "Tabdatum"
    
    loeschNEW "BEAUFNR", gdApp
    TransferTab gdBase, App.Path & "\kissapp.mdb", "BEAUFNR"
    
    
    cSQL = "Insert into TABANZEIGE Select * from Tabdatum where TABNAME like '" & anfangsbuch & "*' "
    gdApp.Execute cSQL, dbFailOnError
    
    cSQL = "Select * from TABANZEIGE "
    Set rsrs1 = gdApp.OpenRecordset(cSQL)
    If Not rsrs1.EOF Then
        rsrs1.MoveFirst
        Do While Not rsrs1.EOF
            If Not IsNull(rsrs1!tabname) Then
                
                sdat = Trim(UCase$(rsrs1!tabname))
                
                lAufNr = 0
                
                dAufwert = ermAuftragswert(sdat)
                
                rsrs1.Edit
                rsrs1!AUFWERT = dAufwert
                rsrs1!AUFTRAGSNR = lAufNr
                               
                rsrs1.Update
                
                
            End If
        rsrs1.MoveNext
        Loop
    
    End If
    rsrs1.Close: Set rsrs1 = Nothing
    
    cSQL = "Update TABANZEIGE t inner join BEAUFNR b on t.tabname = b.tabname set t.AUFTRAGSNR = b.AUFNR "
    gdApp.Execute cSQL, dbFailOnError
    
    list.Clear
    dSumAufwert = 0
    
    cSQL = "Select * from TABANZEIGE order by  " & corder
    Set rsrs1 = gdApp.OpenRecordset(cSQL)
    If Not rsrs1.EOF Then
        rsrs1.MoveFirst
        Do While Not rsrs1.EOF
            If Not IsNull(rsrs1!tabname) Then
                cSatz = ""
                cdatei = Trim(UCase$(rsrs1!tabname))
                sdat = cdatei
                cSatz = cSatz & Space$(8 - Len(cdatei)) & " " & cdatei
            
                If Not IsNull(rsrs1!tabdate) Then
                    cDatum = Format(rsrs1!tabdate, "DD.MM.YY")
                End If
                cSatz = cSatz & Space$(12 - Len(cDatum)) & " " & cDatum
                
                If Not IsNull(rsrs1!LIEFBEZ) Then
                    sLiefname = rsrs1!LIEFBEZ
                End If
                
                cSatz = cSatz & "  " & sLiefname & Space$(36 - Len(sLiefname))
                
                If Not IsNull(rsrs1!AUFWERT) Then
                    dAufwert = rsrs1!AUFWERT
                End If
                
                dSumAufwert = dSumAufwert + dAufwert
                
                If Not IsNull(rsrs1!AUFTRAGSNR) Then
                    lAufNr = rsrs1!AUFTRAGSNR
                End If
                
                cSatz = cSatz & Space$(12 - Len(Format(dAufwert, "#####0.00"))) & Format(dAufwert, "#####0.00")
                cSatz = cSatz & " " & lAufNr
                
                If Not IsNull(rsrs1!KURZINFO) Then
                    cKurzinfo = rsrs1!KURZINFO
                Else
                    cKurzinfo = ""
                End If
                cSatz = cSatz & " " & cKurzinfo

            
                list.AddItem cSatz
            End If
        rsrs1.MoveNext
        Loop
    
    End If
    rsrs1.Close: Set rsrs1 = Nothing
    
    lblx.Caption = Format(dSumAufwert, "#####0.00") & " " & gcWaehrung
    lblx.Refresh
    
    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ListeFuellAnfangsbuchdataT"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
'    Resume Next
End Sub
Private Function ermAuftragswert(sTabname As String) As Double
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL    As String
    Dim rsrs    As Recordset
    Dim lMenge  As Long
    Dim dLEK    As Double
    
    ermAuftragswert = 0
    
    If NewTableSuchenDBKombi(sTabname, gdBase) = False Then
    
        Exit Function
    End If
    
    
    
    If SpalteInTabellegefundenNEW(sTabname, "BESTELLT", gdBase) Then
        sSQL = " select BESTELLT as Menge,LEKPR from " & sTabname & " "
    ElseIf SpalteInTabellegefundenNEW(sTabname, "BESTVOR", gdBase) Then
        sSQL = " select bestvor as Menge,LEKPR from " & sTabname & " "
    Else
        Exit Function
    End If

    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
    
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!MENGE) Then
            
                lMenge = rsrs!MENGE
                
                If Not IsNull(rsrs!lekpr) Then
                    dLEK = rsrs!lekpr
                    If lMenge > 0 Then
                        If dLEK > 0 Then
                        
                            If dLEK * lMenge > 1000000 Then
                            Else
                                ermAuftragswert = ermAuftragswert + (dLEK * lMenge)
                            End If
                            
                        End If
                    End If
                    
                End If
                
            End If
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ermAuftragswert"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Function
Private Function ermAuftragswert_neu(sTabname As String) As Double
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL    As String
    Dim rsrs    As Recordset
    
    ermAuftragswert_neu = 0
    
    If SpalteInTabellegefundenNEW(sTabname, "BESTELLT", gdBase) Then
        sSQL = " select sum(BESTELLT * LEKPR) as AUFWERT from " & sTabname & " "
        sSQL = sSQL & " where lekpr > 0 "
    ElseIf SpalteInTabellegefundenNEW(sTabname, "BESTVOR", gdBase) Then
        sSQL = " select sum(bestvor * LEKPR) as AUFWERT from " & sTabname & " "
        sSQL = sSQL & " where lekpr > 0 "
    Else
        Exit Function
    End If
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!AUFWERT) Then
            ermAuftragswert_neu = rsrs!AUFWERT
        End If
    End If

    rsrs.Close

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ermAuftragswert_neu"
    Fehler.gsFehlertext = "Im Programmteil Wareneingang aus Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Function
Public Sub CreateWKEINSTE()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim cPfad       As String
    Dim cPfadA      As String
    
    cPfad = gcDBPfad 'Datenbankpfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfadA = gcPfad 'Anwendungspfad
    If Right(cPfadA, 1) <> "\" Then
        cPfadA = cPfadA & "\"
    End If
    
    loeschNEW "WKEINSTE", gdApp
    
    sSQL = "Create Table WKEINSTE "
    sSQL = sSQL & "("
    sSQL = sSQL & "Pname text(20)"
    sSQL = sSQL & ", H1 double"
    sSQL = sSQL & ", U1 double"
    sSQL = sSQL & ", S1 double"
    sSQL = sSQL & ", H2 double"
    sSQL = sSQL & ", SB1 double"
    sSQL = sSQL & ", WARN double"
    sSQL = sSQL & ", LINK double"
    sSQL = sSQL & ", UPDPFAD text(200)"
    sSQL = sSQL & ", DABAPFAD text(200)"
    sSQL = sSQL & ", Pversion long"
    sSQL = sSQL & ")"
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "insert into WKEINSTE (Pname) values  ('WinKiss') "
    gdApp.Execute sSQL, dbFailOnError
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "CreateWKEINSTE"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub LeseLagerCFG()
    On Error GoTo LOKAL_ERROR

    Dim iFileNr As Integer
    Dim ctmp As String
    Dim cPfad As String
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    iFileNr = FreeFile
    Open cPfad & "Lager.cfg" For Binary As #iFileNr
    If LOF(iFileNr) > 0 Then
        ctmp = Space$(LOF(iFileNr))
        Get #iFileNr, 1, ctmp
        gsLagerFTPBox = Trim(ctmp)
        Close iFileNr
    Else
        Close iFileNr
        gsLagerFTPBox = ""
    End If
    
    If gsLagerFTPBox = "" Then
        MsgBox "Tragen Sie bitte einen FTP - Speicherplatz in die 'Lager.cfg' ein!", vbInformation, "Winkiss Hinweis:"
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "LeseLagerCFG"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub DabaPfadNew83()
    On Error GoTo LOKAL_ERROR

    Dim iFileNr As Integer
    Dim ctmp As String


    If gbLokalModus Then
        gcDBPfad = "C:\aLeer"
    Else
        iFileNr = FreeFile
        Open gcPfad & "KISSLITE.INI" For Binary As #iFileNr
        If LOF(iFileNr) > 0 Then
            ctmp = Space$(LOF(iFileNr))
            Get #iFileNr, 1, ctmp
            gcDBPfad = ctmp
            Close iFileNr
        Else
            Close iFileNr
            gcDBPfad = ""
            Kill gcPfad & "KISSLITE.INI"
        End If
        
        If gcDBPfad <> "" Then
            If Right(gcDBPfad, 2) = vbCrLf Then
                gcDBPfad = Left(gcDBPfad, Len(gcDBPfad) - 2)
            End If
            
            gcWKDBPfad = ShortPath(gcDBPfad)
        End If
    End If
    
    gcDBPfad = ShortPath(gcDBPfad)
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "DabaPfadNew83"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Function PfadDoemer() As String
    On Error GoTo LOKAL_ERROR

    Dim iFileNr As Integer
    Dim ctmp As String
    
    PfadDoemer = ""
    
    Dim cpfaddb As String
    cpfaddb = gcDBPfad
    If Right$(cpfaddb, 1) <> "\" Then
        cpfaddb = cpfaddb & "\"
    End If


    
    iFileNr = FreeFile
    Open cpfaddb & "doemer.cfg" For Binary As #iFileNr
    If LOF(iFileNr) > 0 Then
        ctmp = Space$(LOF(iFileNr))
        Get #iFileNr, 1, ctmp
        PfadDoemer = ctmp
        Close iFileNr
    Else
        Close iFileNr
        PfadDoemer = ""
        Kill cpfaddb & "doemer.cfg"
    End If
    
    If PfadDoemer <> "" Then
        If Right(PfadDoemer, 2) = vbCrLf Then
            PfadDoemer = Left(PfadDoemer, Len(PfadDoemer) - 2)
        End If
    End If
   
    PfadDoemer = ShortPath(PfadDoemer)
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "PfadDoemer"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Sub DabaPfadNew84()
    On Error GoTo LOKAL_ERROR

    Dim iFileNr As Integer
    Dim ctmp As String
    
    iFileNr = FreeFile
    Open gcPfad & "KISSLITE.INI" For Binary As #iFileNr
    If LOF(iFileNr) > 0 Then
        ctmp = Space$(LOF(iFileNr))
        Get #iFileNr, 1, ctmp
        gcDBPfad = ctmp
        Close iFileNr
    Else
        Close iFileNr
        gcDBPfad = ""
        Kill gcPfad & "KISSLITE.INI"
    End If
    
    If gcDBPfad <> "" Then
        If Right(gcDBPfad, 2) = vbCrLf Then
            gcDBPfad = Left(gcDBPfad, Len(gcDBPfad) - 2)
        End If
    End If
    
    gcDBPfad = ShortPath(gcDBPfad)
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "DabaPfadNew84"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub schickden_fehlenden_Report_Info_PerMail(cReport As String)
On Error GoTo LOKAL_ERROR
        
    Dim cAbsenderEmail As String
    Dim cAnEmailadresse As String
    Dim cBetreff As String
    Dim cMessagetext As String
    Dim sAttachment As String
    Dim cPfad As String
    
    If cReport = "" Then
        Exit Sub
    End If
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If

    sAttachment = ""
    
    cAbsenderEmail = ermFirmenMail
    If cAbsenderEmail <> "" Then cAbsenderEmail = "fehler@kisswws.de"
    
    cAnEmailadresse = "hotline@kisswws.de"
    cBetreff = "fehlender Report(Winkiss " & Left(WKVersion, 2) & "." & Right(WKVersion, 2) & "): " & cReport & " Firma: " & ermFirmenBez
    
    cMessagetext = "Firma: " & ermFirmenBez
    cMessagetext = cMessagetext & " findet diesen Report: " & cReport & " nicht." & vbCrLf
    cMessagetext = cMessagetext & "Bitte beantworten Sie diese Email nicht."
    

    
    schickeMailimHintergrundSSL ermFirmenBez, cAbsenderEmail, "", cAnEmailadresse _
    , "bestsend@kisswws.de", gcSMTP_SERVER, gcSMTP_PORT, gcSMTP_USER, gcSMTP_PW, cBetreff, cMessagetext, sAttachment
      
    Screen.MousePointer = 0
      
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul5"
    Fehler.gsFunktion = "schickden_fehlenden_Report_Info_PerMail"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Public Sub reportbildschirm(dname As String, aname As String)
    On Error GoTo LOKAL_ERROR

    Dim cPfad As String
    Dim iFileNr As Integer
    Dim ctmp As String
    Dim cdatei As String
    Dim sSQL As String
    Dim ierrz   As Integer
    ierrz = 0
    
    Screen.MousePointer = 11

    cPfad = gcDBPfad            'Datenbankpfad
    
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    setzedrucker gcListenDrucker
    
    Screen.MousePointer = 11
    With frmWKL00.CrystalReport1
        If Not Modul6.FindFile(cPfad, aname & ".rpt") Then
            Screen.MousePointer = 0
            
            ctmp = "Die Druckvorschau kann nicht erstellt werden. " & vbCrLf & vbCrLf
            ctmp = ctmp & "Die Datei: " & aname & ".rpt fehlt im" & vbCrLf
            ctmp = ctmp & "Datenbankpfad: " & cPfad & " " & vbCrLf
            ctmp = ctmp & "Laden Sie sich die Datei unter www.kisslive.de/winkiss/downloads/reporte.html runter." & vbCrLf
            ctmp = ctmp & "oder" & vbCrLf
            ctmp = ctmp & "Rufen Sie die Hotline (0511/9559110) an!" & vbCrLf
            ctmp = ctmp & "Wir stellen Ihnen dann die Datei zur Verfügung." & vbCrLf & vbCrLf
            ctmp = ctmp & "Ihr K.I.S.S. Team" & vbCrLf
            
            schickden_fehlenden_Report_Info_PerMail aname & ".rpt"
            MsgBox ctmp, vbOKOnly, "Winkiss Hinweis:"
            
            Exit Sub
        Else
            Screen.MousePointer = 11
            Pause 2
            .ReportFileName = cPfad & aname & ".rpt"
            .WindowAllowDrillDown = True
            .WindowTop = 0
            .WindowLeft = 0
            .WindowHeight = Screen.Height / 15
            .WindowWidth = Screen.Width / 15
            .WindowTitle = .ReportFileName
            .Destination = 0
            .Action = 1
            
            
            If gbSaveReport = True Then
            
                Dim lWert       As Long
                Dim sTime       As String
        
                sTime = TimeValue(Now)
                sTime = Right(sTime, 8)
                sTime = SwapStr(sTime, ":", "")
                lWert = DateValue(Now)
                ctmp = Format$(lWert, "DDMMYY")
                
                reportbildschirmtoRTF_inDaba aname, cPfad & "Reporte\" & ctmp & "_" & sTime & ".rtf"
                
            End If
            
        End If
    End With
    
    Screen.MousePointer = 0
    
    Exit Sub
LOKAL_ERROR:
    If err.Number = 20520 Then
        Exit Sub
    Else
    
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "reportbildschirm"
        Fehler.gsFehlertext = "Der Ausdruck " & aname & " konnte nicht geöffnet werden. "
            
        Fehlermeldung1
    
        Resume Next
    End If

End Sub
Public Sub reportbildschirm_Gutschein(dname As String, aname As String, bMitVorschau As Boolean)
    On Error GoTo LOKAL_ERROR

    Dim cPfad As String
    Dim iFileNr As Integer
    Dim ctmp As String
    Dim cdatei As String
    Dim sSQL As String
    Dim ierrz   As Integer
    ierrz = 0
    
    Screen.MousePointer = 11

    cPfad = gcDBPfad            'Datenbankpfad
    
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
'    setzedrucker gcListenDrucker
    
    Screen.MousePointer = 11
    With frmWKL00.CrystalReport1
        If Not Modul6.FindFile(cPfad, aname & ".rpt") Then
            Screen.MousePointer = 0
            
            ctmp = "Die Druckvorschau kann nicht erstellt werden. " & vbCrLf & vbCrLf
            ctmp = ctmp & "Die Datei: " & aname & ".rpt fehlt im" & vbCrLf
            ctmp = ctmp & "Datenbankpfad: " & cPfad & " " & vbCrLf
            ctmp = ctmp & "Laden Sie sich die Datei unter www.kisslive.de/winkiss/downloads/reporte.html runter." & vbCrLf
            ctmp = ctmp & "oder" & vbCrLf
            ctmp = ctmp & "Rufen Sie die Hotline (0511/9559110) an!" & vbCrLf
            ctmp = ctmp & "Wir stellen Ihnen dann die Datei zur Verfügung." & vbCrLf & vbCrLf
            ctmp = ctmp & "Ihr K.I.S.S. Team" & vbCrLf
            
            schickden_fehlenden_Report_Info_PerMail aname & ".rpt"
            MsgBox ctmp, vbOKOnly, "Winkiss Hinweis:"
            
            Exit Sub
        Else
            Screen.MousePointer = 11
            Pause 1
            .ReportFileName = cPfad & aname & ".rpt"
            .WindowAllowDrillDown = True
            .WindowTop = 0
            .WindowLeft = 0
            .WindowHeight = Screen.Height / 15
            .WindowWidth = Screen.Width / 15
            .WindowTitle = .ReportFileName
            If bMitVorschau = True Then
                .Destination = 0
            Else
                .Destination = 1
            End If
            .Action = 1
            
            
'            If gbSaveReport = True Then
'
'                Dim lWert       As Long
'                Dim sTime       As String
'
'                sTime = TimeValue(Now)
'                sTime = Right(sTime, 8)
'                sTime = SwapStr(sTime, ":", "")
'                lWert = DateValue(Now)
'                ctmp = Format$(lWert, "DDMMYY")
'
'                reportbildschirmtoRTF_inDaba aname, cPfad & "Reporte\" & ctmp & "_" & sTime & ".rtf"
'
'            End If
            
        End If
    End With
    
    Screen.MousePointer = 0
    
    Exit Sub
LOKAL_ERROR:
    If err.Number = 20520 Then
        Exit Sub
    Else
    
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "reportbildschirm_Gutschein"
        Fehler.gsFehlertext = "Der Ausdruck " & aname & " konnte nicht geöffnet werden. "
            
        Fehlermeldung1
    
        Resume Next
    End If

End Sub
Public Sub reportbildschirmtoRTF_inDaba(aname As String, sVolldesPath As String)
    On Error GoTo LOKAL_ERROR

    Dim cPfad As String
    Screen.MousePointer = 11

    cPfad = gcDBPfad            'dabapfad
    
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    With frmWKL00.CrystalReport1
        If Not FindFile(cPfad, aname & ".rpt") Then
            Screen.MousePointer = 0
            Exit Sub
        Else
            Pause 1
            

            .ReportFileName = cPfad & aname & ".rpt"
            .PrintFileName = sVolldesPath
            .PrintFileType = 17 ' doc - word: das war vorher
            .Destination = 2
            .Action = 1
        End If
    End With
    
    Screen.MousePointer = 0
    Exit Sub
LOKAL_ERROR:
End Sub
Public Sub reportbildschirmohneDrucker(dname As String, aname As String)
    On Error GoTo LOKAL_ERROR

    Dim cPfad As String
    Dim iFileNr As Integer
    Dim ctmp As String
    Dim cdatei As String
    Dim sSQL As String
    Dim ierrz   As Integer
    ierrz = 0
    
    Screen.MousePointer = 11

    cPfad = gcPfad            'Apppfad
    
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    Screen.MousePointer = 11
    With frmWKL00.CrystalReport1
        If Not Modul6.FindFile(cPfad, aname & ".rpt") Then
            Screen.MousePointer = 0
            MsgBox "Die Druckvorschau kann nicht erstellt werden.", vbOKOnly, "Winkiss Hinweis:"
            Exit Sub
        Else
            Screen.MousePointer = 11
            Pause 2
'            .PrinterName = gcListenDrucker
            .ReportFileName = cPfad & aname & ".rpt"
            .WindowAllowDrillDown = True
            .WindowTop = 0
            .WindowLeft = 0
        
            
            .WindowHeight = Screen.Height / 15
            .WindowWidth = Screen.Width / 15
            .Destination = 0
            .Action = 1
           
        End If
    End With
    
    Screen.MousePointer = 0
    
    
    Exit Sub
LOKAL_ERROR:
    If err.Number = 20520 Then
'        If ierrz < 5 Then
'            ierrz = ierrz + 1
'            Pause (1)
'            Resume
'        Else
'            Fehler.gsDescr = err.Description
'            Fehler.gsNumber = err.Number
'            Fehler.gsFormular = "Modul2"
'            Fehler.gsFunktion = "reportbildschirm"
'            Fehler.gsFehlertext = "Nach 5 sec konnte der Ausdruck " & aname & " nicht geöffnet werden. "
'
'
'            Fehlermeldung1
            Exit Sub
'        End If
        
    Else
    
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "reportbildschirmohneDrucker"
        Fehler.gsFehlertext = "Der Ausdruck " & aname & " konnte nicht geöffnet werden. "
            
        
        Fehlermeldung1
    
        Resume Next
    End If

End Sub
Public Function ReportVorhanden(sRep As String)
On Error GoTo LOKAL_ERROR

    ReportVorhanden = False
    
    Dim cPfad23 As String
    
    cPfad23 = gcDBPfad               'Datenbankpfad
    If Right(cPfad23, 1) <> "\" Then
        cPfad23 = cPfad23 & "\"
    End If
    
    If FileExists(cPfad23 & sRep & ".rpt") Then
        ReportVorhanden = True
    End If
    Exit Function
LOKAL_ERROR:
   
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ReportVorhanden"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
    Fehlermeldung1
    
End Function
Public Sub reportbildschirmtoText(aname As String, sVolldesPath As String)
    On Error GoTo LOKAL_ERROR

    Dim cPfad As String
    Dim iFileNr As Integer
    Dim ctmp As String
    Dim cdatei As String
    Dim sSQL As String
    
    Screen.MousePointer = 11

    cPfad = gcDBPfad            'Datenbankpfad
    
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    With frmWKL00.CrystalReport1
        If Not Modul6.FindFile(cPfad, aname & ".rpt") Then
            Screen.MousePointer = 0
            Exit Sub
        Else
'            Pause 2

            .ReportFileName = cPfad & aname & ".rpt"
            .PrintFileName = sVolldesPath
            .PrintFileType = crptText
            .ProgressDialog = False
            .Destination = 2
            .Action = 1
            

''''            .ReportFileName = cPfad & aname & ".rpt"
''''            .PrintFileName = sVolldesPath
''''            .PrintFileType = crptWinWord
''''
''''            .ProgressDialog = False
''''            .Destination = 2
''''            .Action = 1
''''
           
        End If
    End With
    
    Screen.MousePointer = 0
    
    
    Exit Sub
LOKAL_ERROR:
'    MsgBox "Die Druckvorschau kann nicht erstellt werden.", vbOKOnly, "Winkiss Hinweis:"
'    Fehler.gsDescr = err.Description
'    Fehler.gsNumber = err.Number
'    Fehler.gsFormular = "Modul2"
'    Fehler.gsFunktion = "reportbildschirmtoText"
'    Fehler.gsFehlertext = "Der Ausdruck " & aname & " konnte nicht geöffnet werden. "
'
'
'    Fehlermeldung1
'
'    Resume Next

End Sub
Public Sub reportbildschirmtoPDF(aname As String, sVolldesPath As String)
    On Error GoTo LOKAL_ERROR

    Dim cPfad As String
    Dim iFileNr As Integer
    Dim ctmp As String
    Dim cdatei As String
    Dim sSQL As String
    
    Screen.MousePointer = 11

    cPfad = gcDBPfad            'Datenbankpfad
    
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    With frmWKL00.CrystalReport1
        If Not Modul6.FindFile(cPfad, aname & ".rpt") Then
            Screen.MousePointer = 0
            Exit Sub
        Else
'            Pause 2

            .ReportFileName = cPfad & aname & ".rpt"
            .PrintFileName = sVolldesPath
            
            .PrintFileType = crptRTF
'            .PrintFileType = crptWinWord
            .ProgressDialog = False
            .Destination = 2
            .Action = 1
            

''''            .ReportFileName = cPfad & aname & ".rpt"
''''            .PrintFileName = sVolldesPath
''''            .PrintFileType = crptWinWord
''''
''''            .ProgressDialog = False
''''            .Destination = 2
''''            .Action = 1
''''
           
        End If
    End With
    
    Screen.MousePointer = 0
    
    
    Exit Sub
LOKAL_ERROR:
'    MsgBox "Die Druckvorschau kann nicht erstellt werden.", vbOKOnly, "Winkiss Hinweis:"
'    Fehler.gsDescr = err.Description
'    Fehler.gsNumber = err.Number
'    Fehler.gsFormular = "Modul2"
'    Fehler.gsFunktion = "reportbildschirmtoPDF"
'    Fehler.gsFehlertext = "Der Ausdruck " & aname & " konnte nicht geöffnet werden. "
'
'
'    Fehlermeldung1
'
'    Resume Next

End Sub
Public Sub reportbildschirmToPrinter(aname As String)
    On Error GoTo LOKAL_ERROR

    Dim cPfad As String
    Dim iFileNr As Integer
    Dim ctmp As String
    Dim cdatei As String
    Dim sSQL As String
    
    Screen.MousePointer = 11
    
    cPfad = gcDBPfad            'Datenbankpfad
    
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    setzedrucker gcListenDrucker
    
    With frmWKL00.CrystalReport1
        If Not Modul6.FindFile(cPfad, aname & ".rpt") Then
            Screen.MousePointer = 0
            MsgBox "Die Druckvorschau kann nicht erstellt werden.", vbOKOnly, "Winkiss Hinweis:"
            Exit Sub
        Else
            Pause 2
            
            .ReportFileName = cPfad & aname & ".rpt"
            .WindowAllowDrillDown = True
            .WindowTop = 0
            .WindowLeft = 0
            
            .WindowHeight = Screen.Height / 15
            .WindowWidth = Screen.Width / 15
            .Destination = 1
            .Action = 1
           
        End If
    End With
    
    Screen.MousePointer = 0
    
    'Fehlerauslöser wird vermutlich der nicht eingeschaltete Drucker sein
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "reportbildschirmToPrinter"
    Fehler.gsFehlertext = "Der Ausdruck " & aname & " konnte nicht geöffnet werden. "
        
    
    Fehlermeldung1
    Resume Next
End Sub
Public Sub reportbildschirmToPrinterETI(aname As String, cDrucker As String, bPrinterset As Boolean)
    On Error GoTo LOKAL_ERROR

    Dim cPfad As String
    Dim iFileNr As Integer
    Dim ctmp As String
    Dim cdatei As String
    Dim sSQL As String
    
    Screen.MousePointer = 11
    
    cPfad = gcDBPfad            'Datenbankpfad
    
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    If bPrinterset Then
        setzedrucker cDrucker
    End If
    
    With frmWKL00.CrystalReport1
        If Not Modul6.FindFile(cPfad, aname & ".rpt") Then
            Screen.MousePointer = 0
            MsgBox "Die Druckvorschau kann nicht erstellt werden.", vbOKOnly, "Winkiss Hinweis:"
            Exit Sub
        Else
            If bPrinterset Then
                Pause 2
            End If
            
            .ReportFileName = cPfad & aname & ".rpt"
            .WindowAllowDrillDown = True
            .WindowTop = 0
            .WindowLeft = 0
            
            .WindowHeight = Screen.Height / 15
            .WindowWidth = Screen.Width / 15
            .Destination = 1 'direkt an den Drucker
'            .Destination = 0 'auf Bildschirm anzeigen
            .Action = 1
        End If
    End With
    
    Screen.MousePointer = 0
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "reportbildschirmToPrinterETI"
    Fehler.gsFehlertext = "Der Ausdruck " & aname & " konnte nicht geöffnet werden. "
        
    
    Fehlermeldung1
    Resume Next
   
End Sub
Public Sub reportbildschirmToPrinterAPP(aname As String, cDrucker As String)
    On Error GoTo LOKAL_ERROR

    Dim cPfad As String
    Dim iFileNr As Integer
    Dim ctmp As String
    Dim cdatei As String
    Dim sSQL As String
    
    Screen.MousePointer = 11
    
    cPfad = App.Path           'Anwendungspfad
    
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    setzedrucker cDrucker
    
    With frmWKL00.CrystalReport1
        If Not Modul6.FindFile(cPfad, aname & ".rpt") Then
            Screen.MousePointer = 0
            MsgBox "Die Druckvorschau kann nicht erstellt werden.", vbOKOnly, "Winkiss Hinweis:"
            Exit Sub
        Else
            Pause 2
            
            .ReportFileName = cPfad & aname & ".rpt"
            .WindowAllowDrillDown = True
            .WindowTop = 0
            .WindowLeft = 0
            
            .WindowHeight = Screen.Height / 15
            .WindowWidth = Screen.Width / 15
            .Destination = 1
            .Action = 1
           
        End If
    End With
    
    Screen.MousePointer = 0
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "reportbildschirmToPrinterAPP"
    Fehler.gsFehlertext = "Der Ausdruck " & aname & " konnte nicht geöffnet werden. "
        
    
    Fehlermeldung1
    Resume Next
   
End Sub
Private Function holmaldenReportvonDABA(aname As String) As Boolean
    On Error GoTo LOKAL_ERROR
    
    holmaldenReportvonDABA = False
    
    
    
    Dim sOldname        As String
    Dim sNewname        As String
    Dim lRet            As Long
    Dim lfail           As Long
    Dim cPfad1          As String
    Dim cPfad2          As String
    
    cPfad1 = gcDBPfad
    If Right(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If
    
    cPfad2 = App.Path
    If Right(cPfad2, 1) <> "\" Then
        cPfad2 = cPfad2 & "\"
    End If
    
    
    
    sOldname = cPfad1 & aname & ".rpt"
    sNewname = cPfad2 & aname & ".rpt"
    
    If FileExists(cPfad1 & aname & ".rpt") Then
        lRet = CopyFile(sOldname, sNewname, lfail)
        If lRet = 0 Then
            holmaldenReportvonDABA = False
        Else
            holmaldenReportvonDABA = True
        End If
    Else
        holmaldenReportvonDABA = False
    End If

Exit Function
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "holmaldenReportvonDABA"
    Fehler.gsFehlertext = "Der Ausdruck " & aname & " konnte nicht kopiert werden. "
    
    Fehlermeldung1
    
End Function
Public Sub reportbildschirmtoTextAppBestellEmail(aname As String, sVolldesPath As String)
    On Error GoTo LOKAL_ERROR

    Dim cPfad As String
    Screen.MousePointer = 11

    cPfad = App.Path            'anwendungspfad
    
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    With frmWKL00.CrystalReport1
        If Not FindFile(cPfad, aname & ".rpt") Then
            Screen.MousePointer = 0
            Exit Sub
        Else
            Pause 1

            .ReportFileName = cPfad & aname & ".rpt"
            .PrintFileName = sVolldesPath
            .PrintFileType = 17
            .Destination = 2
            .Action = 1
        End If
    End With
    
    
    Screen.MousePointer = 0
    
    
    Exit Sub
LOKAL_ERROR:
End Sub
Public Sub reportbildschirmApp(dname As String, aname As String)
    On Error GoTo LOKAL_ERROR

    Dim cPfad As String
    Dim sSQL As String
    Dim ctmp As String
    
    cPfad = App.Path           'Anwendungspfad
    
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    setzedrucker gcListenDrucker
    
    With frmWKL00.CrystalReport1
        If Not FileExists(cPfad & aname & ".rpt") Then
        
            If holmaldenReportvonDABA(aname) = False Then
            
                ctmp = "Die Druckvorschau kann nicht erstellt werden. " & vbCrLf & vbCrLf
                ctmp = ctmp & "Die Datei: " & aname & ".rpt fehlt im" & vbCrLf
                ctmp = ctmp & "Anwendungspfad: " & cPfad & " " & vbCrLf
                ctmp = ctmp & "Rufen Sie die Hotline (0511/9559110) an!" & vbCrLf
                ctmp = ctmp & "Wir stellen Ihnen dann die Datei zur Verfügung." & vbCrLf & vbCrLf
                ctmp = ctmp & "Ihr K.I.S.S. Team" & vbCrLf
                
                schickden_fehlenden_Report_Info_PerMail aname & ".rpt"
                MsgBox ctmp, vbOKOnly, "Winkiss Hinweis:"
            
                Exit Sub
            Else

                .ReportFileName = cPfad & aname & ".rpt"
            End If
        Else

            .ReportFileName = cPfad & aname & ".rpt"
        End If

        Pause 2
        .WindowAllowDrillDown = True
        .WindowTop = 0
        .WindowLeft = 0
        .WindowHeight = Screen.Height / 15
        .WindowWidth = Screen.Width / 15
        .WindowTitle = .ReportFileName
        .Destination = 0
        .Action = 1
    End With
    
    Exit Sub
LOKAL_ERROR:
    If err.Number = 20520 Then
        Exit Sub
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "reportbildschirmApp"
        Fehler.gsFehlertext = "Der Ausdruck " & aname & " konnte nicht geöffnet werden. "
        
        Fehlermeldung1
    End If
    
End Sub
Public Sub Fehlermeldung(Fehlertext As String, Formular As String, Funktion As String, Number As String, Description As String)
    On Error GoTo LOKAL_ERROR
    
    Description = SwapStr(Description, "'", " ")
    
        MsgBox Fehlertext & vbCrLf _
            & "Formular: " & Formular & vbCrLf _
            & "Funktion: " & Funktion & vbCrLf _
            & "Fehlernummer: " & Number & vbCrLf _
            & "Fehlerbeschreibung: " & Description & vbCrLf _
            & "Programmversion: " & WKVersion, vbCritical + vbOKOnly, "Winkiss Fehlermeldung:"
        
    
    Exit Sub
LOKAL_ERROR:
    MsgBox "Formular: Modul2" & vbCrLf _
            & "Funktion: Fehlermeldung " & vbCrLf _
            & "Fehlernummer: " & err.Number & vbCrLf _
            & "Fehlerbeschreibung: " & err.Description & vbCrLf _
            & "Programmversion: " & WKVersion, vbCritical + vbOKOnly, "Winkiss Fehlermeldung:"
End Sub
Public Sub WertSchreiben(sWert As String, sSEC As String, sKey As String, sKomplettPfad As String)
    On Error GoTo LOKAL_ERROR
    
    Dim lResult         As Long
    
    lResult = WritePrivateProfileString(sSEC, sKey, sWert, sKomplettPfad)
    If lResult <> 1 Then
        MsgBox "Fehler"
    End If

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "WertSchreiben"
    Fehler.gsFehlertext = "Im Programmteil KISSNET... Mailbox ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function KopfdatenfFehlermail(sdatname As String) As Boolean
 On Error GoTo LOKAL_ERROR
    
    KopfdatenfFehlermail = False
    
    Dim cPfad       As String
    Dim cPfad1      As String
    Dim cPfad2      As String
    Dim sdatumheute As String
    Dim lDatzahl    As Long
    Dim cdat        As String
    Dim i           As Integer
    Dim j           As Integer
    
    
    Randomize
    
    lDatzahl = Int((99999999 * Rnd) + 1)   ' Zufallszahl im Bereich von 1 bis 99999999
    cdat = CStr(lDatzahl)
    If Len(cdat) < 8 Then
    cdat = Space(8 - Len(cdat)) & cdat
    cdat = SwapStr(cdat, " ", "0")
    End If
        
    sdatumheute = DateValue(Now)
    
    If gbFtpYes = False Then    'Detailinfos FTP
        If NewTableSuchenDBKombi("StammFTP", gdBase) Then
            LeseStammFtp
        End If
    End If
    
    cPfad = gcPfad    'Anwendungspfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = ShortPath(cPfad)
    cPfad = cPfad & "KISSHELP\" & cdat & ".kmd"
    
    Kill cPfad & "KISSHELP\" & cdat & ".kmd"
    'Empfänger schreiben
'    For i = 0 To List1.ListCount - 1
'        j = i + 1
        WertSchreiben "heinz2", "EMAILER", "EMPFAENGER1", cPfad
'    Next i
        
    WertSchreiben "W " & WKVersion & " ErrMsg", "EMAILER", "BETREFF", cPfad
    WertSchreiben cdat, "EMAILER", "HAUPTTEXT", cPfad
    
    'absender ist der username der ftpkennung
    WertSchreiben gsStammFTPUSER, "EMAILER", "ABSENDER", cPfad
    
    'alias ist z.B. der Firmenname
    WertSchreiben gFirma.FirmaName, "EMAILER", "ABSALIAS", cPfad
    WertSchreiben sdatumheute, "EMAILER", "CREATEDAY", cPfad
    
    'umbenennen der Haupttextdatei
    Dim sOldname        As String
    Dim sNewname        As String
    Dim t               As Integer
    Dim lRet            As Long
    Dim lfail           As Long
    
    Dim Task$
    Dim hProcess&
    Dim result&
    
    cPfad1 = gcPfad    'Anwendungspfad
    If Right(cPfad1, 1) <> "\" Then
        cPfad1 = cPfad1 & "\"
    End If
    
    sOldname = cPfad1 & "LERR\" & sdatname
    
'    sOldname = cPfad1 & "KISSHELP\NeueMail.rtf"
    sNewname = cPfad1 & "KISSHELP\" & cdat & ".rtf"
    
    lRet = CopyFile(sOldname, sNewname, lfail)
    If lRet = 0 Then
        MsgBox "Konnte " & sOldname & " nicht kopieren!", vbInformation, "STOP!"
    Else
'        Kill sOldname
    End If
'    Name sOldname As sNewname
    cPfad1 = ShortPath(cPfad1)
    t = 2
    Do Until t = 5
        If Modul6.FindFile(cPfad1 & "KISSHELP\", cdat & ".rtf") Then
            Task = Shell("LHA" & " a" & Space(1) & cPfad1 & "KISSHELP\" & cdat & ".lzh" & Space(1) & cPfad1 & "KISSHELP\" & cdat & ".rtf")

            hProcess = OpenProcess(SYNCHRONIZE, False, Task)
            result = WaitForSingleObject(hProcess, INFINITE)
            result = CloseHandle(hProcess)
            t = 5
'            Pause 5
        End If
    Loop
    
    Kill sNewname
    
    cPfad2 = gcDBPfad    'dabaspfad
    If Right(cPfad2, 1) <> "\" Then
        cPfad2 = cPfad2 & "\"
    End If
    
    
    
    sOldname = cPfad1 & "KISSHELP\" & cdat & ".lzh"
    sNewname = cPfad2 & "Mailout\" & cdat & ".lzh"
    
    lRet = CopyFile(sOldname, sNewname, lfail)
    If lRet = 0 Then
        MsgBox "Konnte " & sOldname & " nicht kopieren!", vbInformation, "STOP!"
    Else
        Kill sOldname
    End If
    
'    Name sOldname As sNewname
    
    sOldname = cPfad1 & "KISSHELP\" & cdat & ".kmd"
    sNewname = cPfad2 & "Mailout\" & cdat & ".kmd"
'    Name sOldname As sNewname

    lRet = CopyFile(sOldname, sNewname, lfail)
    If lRet = 0 Then
        MsgBox "Konnte " & sOldname & " nicht kopieren!", vbInformation, "STOP!"
    Else
        Kill sOldname
    End If
    
    KopfdatenfFehlermail = True
    
Exit Function
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "KopfdatenfFehlermail"
        Fehler.gsFehlertext = "Im Programmteil Fehler by Mail ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Function
Public Sub Fehlermeldung1()
    On Error GoTo LOKAL_ERROR

    Dim sSQL        As String
    Dim i           As Integer
    Dim sRechner    As String
    Dim iZufall     As Long
    Dim iRet        As Integer
    Dim cErrtemp    As String
    
    
    If Fehler.gsNumber = 3343 Then 'Datenbank kaputt
        giUmleitgrund = 2 'Datenbankpfad muß rep
        
        sErrDabapfad = ""
        
        gcUmleittxt = "Ihre Datenbank ist beschädigt. (" & TimeValue(Now) & " Uhr)" & vbCrLf
        gcUmleittxt = gcUmleittxt & "Rufen Sie die Hotline an (0511/955910)" & vbCrLf & vbCrLf
        
        frmWKL60.Show 1
    
        Exit Sub
    End If
    
    Dim lTime As Long
    Dim ctime As String
    
    ctime = TimeValue(Now)
    ctime = SwapStr(ctime, ":", "")
    lTime = CLng(ctime)
    
    If glErrtime < lTime - 10 Then
        glErrtime = lTime 'ok
    Else
        If glErrtime <> 0 Then
            End
        Else
            glErrtime = lTime
        End If
    End If

    Randomize
    iZufall = Int((99999 * Rnd) + 1)   ' Zufallszahl im Bereich von 1 bis 99999

    sRechner = rechnername
    sRechner = SwapStr(sRechner, ".", "")
    
    Fehler.gsFehlertext = SwapStr(Fehler.gsFehlertext, "'", "")
    Fehler.gsDescr = SwapStr(Fehler.gsDescr, "'", "")

    schreibeProtokollFehlermeldung Fehler
    
    If gbOptiStada Then
        schickdieFehlermeldungPerMail Fehler
    End If
    
    
    cErrtemp = "Es trat ein Fehler auf. Möchten Sie uns die Fehlermeldung faxen? " & vbCrLf & vbCrLf
    cErrtemp = cErrtemp & "Dies wäre sehr hilfreich um nächstes Mal in diesem Programmteil fehlerfrei arbeiten zu können." & vbCrLf & vbCrLf
    cErrtemp = cErrtemp & "Vielen Dank!"
    
    iRet = MsgBox(cErrtemp, vbYesNo + vbDefaultButton2, "Winkiss Hinweis:")
    
    If iRet = vbYes Then
        schreibeEinzelFehlermeldungExtra Fehler, gFirma, "F" & CStr(iZufall)

        zeigeHilfeAPPpfad "LERR", "F" & CStr(iZufall) & ".txt"

    ElseIf iRet = vbNo Then
        schreibeEinzelFehlermeldung Fehler, gFirma, "F" & CStr(iZufall)
        If KopfdatenfFehlermail("F" & CStr(iZufall) & ".txt") Then

        End If
    
    End If
    
    Screen.MousePointer = 0

    Exit Sub
LOKAL_ERROR:
    
     MsgBox "Formular: Modul2" & vbCrLf _
     & "Funktion: Fehlermeldung1 " & vbCrLf _
     & "Fehlernummer: " & err.Number & vbCrLf _
     & "Fehlerbeschreibung: " & err.Description & vbCrLf _
     & "Programmversion: " & WKVersion, vbCritical + vbOKOnly, "Winkiss Fehlermeldung:"

End Sub
Public Function tableSuchen(tabname As String) As Boolean
    On Error GoTo LOKAL_ERROR

    Dim lcount As Long
    Dim lAnzTable As Long
    Dim cSQL As String
    Dim name As String
    
    
    tableSuchen = False
    gdApp.TableDefs.Refresh
    lAnzTable = gdApp.TableDefs.Count
    For lcount = 0 To lAnzTable - 1
        name = gdApp.TableDefs(lcount).name
        If UCase(name) = UCase(tabname) Then
            tableSuchen = True
            Exit Function
        End If
    Next lcount

    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "tableSuchen"
    Fehler.gsFehlertext = "Bei der Tabellensuche (" & tabname & ") ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Public Function tableSuchenDBKombi(tabname As String, Database As Integer) As Boolean
    On Error GoTo LOKAL_ERROR

    If Database = 1 Then                'Datenbankpfad
        If tableSuchenDB(tabname) Then
            tableSuchenDBKombi = True
        Else
            tableSuchenDBKombi = False
        End If
    ElseIf Database = 2 Then            'Anwendungspfad
        If tableSuchen(tabname) Then
            tableSuchenDBKombi = True
        Else
            tableSuchenDBKombi = False
        End If
    End If

    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "tableSuchenDBKombi"
    Fehler.gsFehlertext = "Bei der Tabellensuche (" & tabname & ") ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Public Function NewTableSuchenDBKombi(tabname As String, daba As Database) As Boolean
    On Error GoTo LOKAL_ERROR

                      
        If NewTableSuchenDB(tabname, daba) Then
            NewTableSuchenDBKombi = True
        Else
            NewTableSuchenDBKombi = False
        End If

    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "NewTableSuchenDBKombi"
    Fehler.gsFehlertext = "Bei der Tabellensuche (" & tabname & ") ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Public Function tableSuchenDB(tabname As String) As Boolean
    On Error GoTo LOKAL_ERROR

    Dim lcount As Long
    Dim lAnzTable As Long
    Dim cSQL As String
    Dim name As String
    
    tableSuchenDB = False
    gdBase.TableDefs.Refresh    'Dabarefresh
    lAnzTable = gdBase.TableDefs.Count
    For lcount = 0 To lAnzTable - 1
        name = gdBase.TableDefs(lcount).name
        If UCase(name) = UCase(tabname) Then
            tableSuchenDB = True
            Exit Function
        End If
    Next lcount

    Exit Function
LOKAL_ERROR:
    If err.Number = 3167 Then 'Datensatz ist schon gelöscht
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "tableSuchenDB"
        Fehler.gsFehlertext = "Bei der Tabellensuche (" & tabname & ") ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Function
Public Function NewTableSuchenDB(tabname As String, daba As Database) As Boolean
    On Error GoTo LOKAL_ERROR

    Dim lcount As Long
    Dim lAnzTable As Long
    Dim cSQL As String
    Dim name As String
    Dim cVergleichsname As String
    Dim rsrs As DAO.Recordset
    
    
    
    If gbSQLSERVER = True Then
        cVergleichsname = "DBO." & UCase(tabname)
    Else
        cVergleichsname = UCase(tabname)
    End If
    
    NewTableSuchenDB = False
    
    cSQL = "Select * from " & tabname
    Set rsrs = daba.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        NewTableSuchenDB = True
    Else
        NewTableSuchenDB = True
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
    
''    daba.TableDefs.Refresh
''
''    lAnzTable = daba.TableDefs.Count
''
''    For lcount = 0 To lAnzTable - 1
''        name = daba.TableDefs(lcount).name
''
''        If UCase(name) = UCase(cVergleichsname) Then
''            NewTableSuchenDB = True
''            Exit Function
''        End If
''    Next lcount

    Exit Function
LOKAL_ERROR:
    If err.Number = 3167 Then 'Datensatz ist schon gelöscht
        Resume Next
    ElseIf err.Number = 91 Then 'Tabelle nicht vorhanden
        Exit Function
    ElseIf err.Number = 3078 Then 'Tabelle nicht vorhanden
        Exit Function
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "NewTableSuchenDB"
        Fehler.gsFehlertext = "Bei der Tabellensuche (" & tabname & ") ist ein Fehler aufgetreten."
        
        Fehlermeldung1
        Resume Next
    End If
End Function
Public Sub NEWTableSuchenDBLike(tabname As String)
    On Error GoTo LOKAL_ERROR

    Dim lcount As Long
    Dim lAnzTable As Long
    Dim cSQL As String
    Dim name As String
    
    
    gdBase.TableDefs.Refresh
    lAnzTable = gdBase.TableDefs.Count
    
    
    For lcount = 0 To lAnzTable - 1
        name = Left(gdBase.TableDefs(lcount).name, Len(tabname))
        
        
        If UCase(name) = UCase(tabname) Then
            If SpalteInTabellegefundenNEW(gdBase.TableDefs(lcount).name, "lfnr", gdBase) Then Exit Sub
            SpalteAnfuegenNEW gdBase.TableDefs(lcount).name, "lfnr", "autoincrement", gdBase
        
        End If
    Next lcount

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "NewTableSuchenDBLike"
    Fehler.gsFehlertext = "Bei der Tabellensuche (" & tabname & ") ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Public Function notkassetabcheck() As Boolean
On Error GoTo LOKAL_ERROR
    
    notkassetabcheck = False
    
    Dim lokalDB As Database
    Dim i As Integer
   
    Dim sTabellen(0 To 39) As String
    
    sTabellen(0) = "KUNDEN"
    sTabellen(1) = "ARTIKEL"
    sTabellen(2) = "AGNDBF"
    sTabellen(3) = "ARBEIT"
    sTabellen(4) = "TAUSCH"
    sTabellen(5) = "ARTLIEF"
    sTabellen(6) = "Warengru"
    sTabellen(7) = "BANKEN"
    sTabellen(8) = "BEDNAME"
    sTabellen(9) = "BESTREST"
    sTabellen(10) = "BONPAUSE"
    sTabellen(11) = "BONTEXT"
    sTabellen(12) = "BONUSGRE"
    sTabellen(13) = "DBEINSTE"
    sTabellen(14) = "DTA"
    sTabellen(15) = "DTAUS1"
    sTabellen(16) = "FILIALEN"
    sTabellen(17) = "Firma"
    sTabellen(18) = "FILA"
    sTabellen(19) = "GUTSCH"
    sTabellen(20) = "KAEINAUS"
    sTabellen(21) = "PRSTERM"
    sTabellen(22) = "ZBESTAND"
    sTabellen(23) = "KISSLITE"
    sTabellen(24) = "UMS_ART"
    sTabellen(25) = "UMSKDJ"
    sTabellen(26) = "LISRT"
    sTabellen(27) = "MWSTSATZ"
    sTabellen(28) = "RETOURE"
    sTabellen(29) = "BEDZUGRI"
    sTabellen(30) = "ABSCHLUSS"
    sTabellen(31) = "ALTERG"
    sTabellen(32) = "PREISE"
    
    sTabellen(33) = "KASSBON"
    sTabellen(34) = "AFCSTAT"
    sTabellen(35) = "KOLLVERK"
    sTabellen(36) = "KREDIT"
    sTabellen(37) = "AFCBUCH"
    sTabellen(38) = "KASSJOUR"
    sTabellen(39) = "MARKIERUNG"
    
    
    Set lokalDB = OpenDatabase("C:\aLeer\kissdata.mdb", False, False)
    
    For i = 0 To 39
        If Not NewTableSuchenDBKombi(sTabellen(i), lokalDB) Then
            
            notkassetabcheck = False
            Exit For
        Else
            notkassetabcheck = True
        End If
    Next i
    
    lokalDB.Close
    Set lokalDB = Nothing
    
    Exit Function
LOKAL_ERROR:
    If err.Number = 3011 Then
        notkassetabcheck = False
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "notkassetabcheck"
        Fehler.gsFehlertext = "Beim Überprüfen der Tabellen ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Function
Public Sub speicherGutschNotizen(cGutschnr As String, cNotz As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL    As String
    
    Screen.MousePointer = 11
    
    If Trim(cGutschnr) = "" Then
        Exit Sub
    End If
    
    cNotz = SwapStr(cNotz, "'", " ")
    
    If Trim(cNotz) = "" Then
        Exit Sub
    End If
    
    sSQL = "Delete from  GUHIN where gutschnr = " & cGutschnr
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into GUHIN (GUTSCHNR,NOTIZEN) values "
    sSQL = sSQL & " ( " & cGutschnr & ", '" & cNotz & "')"
    gdBase.Execute sSQL, dbFailOnError
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "speicherGutschNotizen"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Function ermittleGutschNotizen(cGutschnr As String) As String
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL    As String
    Dim rsrs    As Recordset
    
    Screen.MousePointer = 11
    
    If cGutschnr = "" Then
        Exit Function
    End If
    
    If IsNumeric(cGutschnr) = False Then
        Exit Function
    End If
    
    ermittleGutschNotizen = ""
        
    sSQL = "Select Notizen from GUHIN where gutschnr = " & cGutschnr
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!NOTIZEN) Then
            ermittleGutschNotizen = rsrs!NOTIZEN
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Screen.MousePointer = 0
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ermittleGutschNotizen"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function kassetabcheck(db As Database, Label2 As Label, Label3 As Label) As String
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim i As Integer
    Dim j As Integer
    
    kassetabcheck = ""
    
    Dim sTabellen(0 To 217) As String
    j = 0
    
    sTabellen(j) = "MAILFB"
    j = j + 1
    sTabellen(j) = "PAUSENZ"
    j = j + 1
    sTabellen(j) = "NICHTUMSBAR"
    j = j + 1
    sTabellen(j) = "ZUGRIFFDAT"
    j = j + 1
    sTabellen(j) = "DESADV"
    j = j + 1
    sTabellen(j) = "KUNDENBONUS"
    j = j + 1
    sTabellen(j) = "BONUS_SYS"
    j = j + 1
    sTabellen(j) = "WEBSHOP_E"
    j = j + 1
    sTabellen(j) = "LISRT_MAIL"
    j = j + 1
    sTabellen(j) = "AZEITU"
    j = j + 1
    sTabellen(j) = "STEMPELAZU"
    j = j + 1
    sTabellen(j) = "STEUERKI"
    j = j + 1
    sTabellen(j) = "PREISEDITKASSE"
    j = j + 1
    sTabellen(j) = "WGNDEF"
    j = j + 1
    sTabellen(j) = "WGNAGN"
    j = j + 1
    sTabellen(j) = "AFCSTATP"
    j = j + 1
    sTabellen(j) = "ETISIC"
    j = j + 1
    sTabellen(j) = "GUHIN"
    j = j + 1
    sTabellen(j) = "GRUPPE"
    j = j + 1
    sTabellen(j) = "GRUPPE_ARTIKEL"
    j = j + 1
    sTabellen(j) = "INTERART"
    j = j + 1
    sTabellen(j) = "ARTEAN_K"
    j = j + 1
    sTabellen(j) = "GESCHWART"
    j = j + 1
    sTabellen(j) = "STAFFELPR"
    j = j + 1
    sTabellen(j) = "STAFFELPRKVK"
    j = j + 1
    sTabellen(j) = "STAFFEL_KVK_ARTIKEL"
    j = j + 1
    sTabellen(j) = "STAFFEL_KVK_GRUPPE"
    j = j + 1
    sTabellen(j) = "LANG"
    j = j + 1
    sTabellen(j) = "DATEV"
    j = j + 1
    sTabellen(j) = "MBSTAND"
    j = j + 1
    sTabellen(j) = "GDLAGER"
    j = j + 1
    sTabellen(j) = "GLAGER"
    j = j + 1
    sTabellen(j) = "KUDD"
    j = j + 1
    sTabellen(j) = "IDENTUSER"
    j = j + 1
    sTabellen(j) = "UMLAGER"
    j = j + 1
    sTabellen(j) = "BONUSBONTEXTE"
    j = j + 1
    sTabellen(j) = "BWWBONTEXTE"
    j = j + 1
    sTabellen(j) = "BONUSART"
    j = j + 1
    sTabellen(j) = "UMS_LINR"
    j = j + 1
    sTabellen(j) = "UMS_ARTNR"
    j = j + 1
    sTabellen(j) = "UMS_LPZ"
    j = j + 1
    sTabellen(j) = "PREISE"
    j = j + 1
    sTabellen(j) = "MBORDER"
    j = j + 1
    sTabellen(j) = "DVKART"
    j = j + 1
    sTabellen(j) = "LUGEVER"
    j = j + 1
    sTabellen(j) = "LAGERW"
    j = j + 1
    sTabellen(j) = "LAGERMW"
    j = j + 1
    sTabellen(j) = "PENLAGERMW"
    j = j + 1
    sTabellen(j) = "LAGERLW"
    j = j + 1
    sTabellen(j) = "LAGERLLW"
    j = j + 1
    sTabellen(j) = "PENLAGERW"
    j = j + 1
    sTabellen(j) = "PENLAGERLW"
    j = j + 1
    sTabellen(j) = "PENLAGERLLW"
    j = j + 1
    sTabellen(j) = "ALLARTLUKOPF"
    j = j + 1
    sTabellen(j) = "GARANTIE"
    j = j + 1
    sTabellen(j) = "ALLARTLU"
    j = j + 1
    sTabellen(j) = "ARTMERK"
    j = j + 1
    sTabellen(j) = "KONTIN"
    j = j + 1
    sTabellen(j) = "MITKU" & srechnertab
    j = j + 1
    sTabellen(j) = "KUNDA" & srechnertab
    j = j + 1
    sTabellen(j) = "RECHUE"
    j = j + 1
    sTabellen(j) = "FARBMERK"
    j = j + 1
    sTabellen(j) = "FARBKU"
    j = j + 1
    sTabellen(j) = "EMAIL"
    j = j + 1
    sTabellen(j) = "WECHSEL"
    j = j + 1
    sTabellen(j) = "MASTERPRO"
    j = j + 1
    sTabellen(j) = "ARTFARB2"
    j = j + 1
    sTabellen(j) = "BESTAKT"
    j = j + 1
    sTabellen(j) = "KUNPFLEG"
    j = j + 1
    sTabellen(j) = "ZUGANGF"
    j = j + 1
    sTabellen(j) = "ABWESEND"
    j = j + 1
    sTabellen(j) = "ETIDRULS"
    j = j + 1
    sTabellen(j) = "REPARATUR"
    j = j + 1
    sTabellen(j) = "REPARATURSATZ"
    j = j + 1
    sTabellen(j) = "REPARATURKOPF"
    j = j + 1
    sTabellen(j) = "KUNPFLGH"
    j = j + 1
    sTabellen(j) = "KUNDEDEL"
    j = j + 1
    sTabellen(j) = "KUNBHIST"
    j = j + 1
    sTabellen(j) = "KABUCH"
    j = j + 1
    sTabellen(j) = "LINRZUO"
    j = j + 1
    sTabellen(j) = "PROVISION"
    j = j + 1
    sTabellen(j) = "NEINVK"
    j = j + 1
    sTabellen(j) = "AFCAUFBON"
    j = j + 1
    sTabellen(j) = "ALTERG"
    j = j + 1
    sTabellen(j) = "KASSBEDP"
    j = j + 1
    sTabellen(j) = "KASSEIN"
    j = j + 1
    sTabellen(j) = "DRUBRU"
    j = j + 1
    sTabellen(j) = "RETPRINT"
    j = j + 1
    sTabellen(j) = "DRUNET"
    j = j + 1
    sTabellen(j) = "ETI" & srechnertab
    j = j + 1
    sTabellen(j) = "MY" & srechnertab
    j = j + 1
    sTabellen(j) = "ZENTHELP"
    j = j + 1
    sTabellen(j) = "NOTART"
    j = j + 1
    sTabellen(j) = "GEMZ"
    j = j + 1
    sTabellen(j) = "ABSCHOPF"
    j = j + 1
    sTabellen(j) = "DUKATENB"
    j = j + 1
    sTabellen(j) = "STORNO2"
    j = j + 1
    sTabellen(j) = "ARTDET"
    j = j + 1
    sTabellen(j) = "ARTMDH"
    j = j + 1
    sTabellen(j) = "KKZAHL"
    j = j + 1
    sTabellen(j) = "KKZAHLTE"
    j = j + 1
    sTabellen(j) = "KREDITZA"
    j = j + 1
    sTabellen(j) = "FEEDB_TRANS"
    j = j + 1
    sTabellen(j) = "FEEDB"
    j = j + 1
    sTabellen(j) = "FEEDBF"
    j = j + 1
    sTabellen(j) = "LASTZAHL"
    j = j + 1
    sTabellen(j) = "OPENINGS"
    j = j + 1
    sTabellen(j) = "LASTZAHLTE"
    j = j + 1
    sTabellen(j) = "GUTZ"
    j = j + 1
    sTabellen(j) = "GUHIS"
    j = j + 1
    sTabellen(j) = "DELSTADAL"
    j = j + 1
    sTabellen(j) = "NOKALKL"
    j = j + 1
    sTabellen(j) = "KOPFMAIL"
    j = j + 1
    sTabellen(j) = "STATIST"
    j = j + 1
    sTabellen(j) = "GROLIEF"
    j = j + 1
    sTabellen(j) = "ETIDRU"
    j = j + 1
    sTabellen(j) = "BEDNAME"
    j = j + 1
    sTabellen(j) = "WKVERSIONEN"
    j = j + 1
    sTabellen(j) = "TEXTBLOCK"
    j = j + 1
    sTabellen(j) = "KUNDBEST"
    j = j + 1
    sTabellen(j) = "KUNDEN"
    j = j + 1
    sTabellen(j) = "KUNDAUSLIEF"
    j = j + 1
    sTabellen(j) = "KONDITIONEN"
    j = j + 1
    sTabellen(j) = "PREISTERM"
    j = j + 1
    sTabellen(j) = "KONDILEK"
    j = j + 1
    sTabellen(j) = "BARGELD"
    j = j + 1
    sTabellen(j) = "DTA"
    j = j + 1
    sTabellen(j) = "TEXTIL"
    j = j + 1
    sTabellen(j) = "LAGERPLATZ"
    j = j + 1
    sTabellen(j) = "MERKFARB"
    j = j + 1
    sTabellen(j) = "TERMINE"
    j = j + 1
    sTabellen(j) = "SPEZIINFO"
    j = j + 1
    sTabellen(j) = "TERM_STD"
    j = j + 1
    sTabellen(j) = "ZBONLAY"
    j = j + 1
    sTabellen(j) = "FEHLZEIT"
    j = j + 1
    sTabellen(j) = "OPENINGS"
    j = j + 1
    sTabellen(j) = "PFLEGORT"
    j = j + 1
    sTabellen(j) = "BEDTERM"
    j = j + 1
    sTabellen(j) = "TABLAY" & srechnertab
    j = j + 1
    sTabellen(j) = "RECHAB"
    j = j + 1
    sTabellen(j) = "KASSWAAG"
    j = j + 1
    sTabellen(j) = "KASQL"
    j = j + 1
    sTabellen(j) = "EKASS"
    j = j + 1
    sTabellen(j) = "EAJ"
    j = j + 1
    sTabellen(j) = "STEMPEL"
    j = j + 1
    sTabellen(j) = "NOEURO"
    j = j + 1
    sTabellen(j) = "KUK"
    j = j + 1
    sTabellen(j) = "SORTI"
    j = j + 1
    sTabellen(j) = "ARTIKEL"
    j = j + 1
    sTabellen(j) = "AGNDBF"
    j = j + 1
    sTabellen(j) = "ARBEIT"
    j = j + 1
    sTabellen(j) = "TAUSCH"
    j = j + 1
    sTabellen(j) = "ARTLIEF"
    j = j + 1
    sTabellen(j) = "Warengru"
    j = j + 1
    sTabellen(j) = "BANKEN"
    j = j + 1
    sTabellen(j) = "BEDNAME"
    j = j + 1
    sTabellen(j) = "BESTREST"
    j = j + 1
    sTabellen(j) = "BONPAUSE"
    j = j + 1
    sTabellen(j) = "ARTAUSWAHL"
    j = j + 1
    sTabellen(j) = "BONTEXT"
    j = j + 1
    sTabellen(j) = "BONUSGRE"
    j = j + 1
    sTabellen(j) = "BANKKU"
    j = j + 1
    sTabellen(j) = "ZUORDEAN"
    j = j + 1
    sTabellen(j) = "DBEINSTE"
    j = j + 1
    sTabellen(j) = "DTA"
    j = j + 1
    sTabellen(j) = "FILIALEN"
    j = j + 1
    sTabellen(j) = "Firma"
    j = j + 1
    sTabellen(j) = "FILA"
    j = j + 1
    sTabellen(j) = "GUTSCH"
    j = j + 1
    sTabellen(j) = "KAEINAUS"
    j = j + 1
    sTabellen(j) = "KAEINAUSF"
    j = j + 1
    sTabellen(j) = "KARTEN_EINZ"
    j = j + 1
    sTabellen(j) = "EINAUSKB"
    j = j + 1
    sTabellen(j) = "PRSTERM"
    j = j + 1
    sTabellen(j) = "ZBESTAND"
    j = j + 1
    sTabellen(j) = "KISSLITE"
    j = j + 1
    sTabellen(j) = "UMS_ART"
    j = j + 1
    sTabellen(j) = "UMSARTJ"
    j = j + 1
    sTabellen(j) = "UMS_ARTF"
    j = j + 1
    sTabellen(j) = "UMSKDJ"
    j = j + 1
    sTabellen(j) = "LISRT"
    j = j + 1
    sTabellen(j) = "MWSTSATZ"
    j = j + 1
    sTabellen(j) = "BEDZUGRI"
    j = j + 1
    sTabellen(j) = "ABSCHLUSS"
    j = j + 1
    sTabellen(j) = "DABAUSER"
    j = j + 1
    sTabellen(j) = "TABDATUM"
    j = j + 1
    sTabellen(j) = "KASSBON"
    j = j + 1
    sTabellen(j) = "KASSBOND"
    j = j + 1
    sTabellen(j) = "AFCSTAT"
    j = j + 1
    sTabellen(j) = "STORNOF"
    j = j + 1
    sTabellen(j) = "KOLLVERK"
    j = j + 1
    sTabellen(j) = "KREDIT"
    j = j + 1
    sTabellen(j) = "KASSJOUR"
    j = j + 1
    sTabellen(j) = "RETOURE"
    j = j + 1
    sTabellen(j) = "UEBERLI"
    j = j + 1
    sTabellen(j) = "KUNDKASS"
    j = j + 1
    sTabellen(j) = "BESTAEND"
    j = j + 1
    sTabellen(j) = "NOTIZEN"
    j = j + 1
    sTabellen(j) = "BESTPROT"
    j = j + 1
    sTabellen(j) = "EANPROT"
    j = j + 1
    sTabellen(j) = "KVKPR1PROT"
    j = j + 1
    sTabellen(j) = "MA" & srechnertab
    j = j + 1
    sTabellen(j) = "TABALI" & srechnertab
    j = j + 1
    sTabellen(j) = "ZADRESS"
    j = j + 1
    sTabellen(j) = "KULOESCH"
    j = j + 1
    sTabellen(j) = "BEAUFNR"
    j = j + 1
    sTabellen(j) = "UNTERWF"
    j = j + 1
    sTabellen(j) = "GANALYSE"
    j = j + 1
    sTabellen(j) = "DISPLAYTEXT"
    j = j + 1
    sTabellen(j) = "GANALYSEALL"
    j = j + 1
    sTabellen(j) = "UMSATZINFO"
    j = j + 1
    sTabellen(j) = "KUMSUM"
    j = j + 1
    sTabellen(j) = "MARKIERUNG"
    j = j + 1
    sTabellen(j) = "TERMINE_ANL"
    j = j + 1
    sTabellen(j) = "GEMISCHTE_Z"
    j = j + 1
    sTabellen(j) = "GEMISCHTE_ZP"
    j = j + 1
    sTabellen(j) = "REPOS"
    j = j + 1
    sTabellen(j) = "AFCBUCH"
    
    
    For i = 0 To j
        anzeige "normal", "Tabellenüberprüfung: ", Label2
        Label3.Visible = True
        anzeige "normal", sTabellen(i), Label3
        

        DoEvents

        If Not NewTableSuchenDBKombi(sTabellen(i), db) Then
            
            kassetabcheck = sTabellen(i)
            Select Case kassetabcheck
            
                Case "ETI" & srechnertab
                    CreateTable "ETI" & srechnertab, gdBase
                Case "PAUSENZ"
                    CreateTable "PAUSENZ", gdBase
                Case "GEMISCHTE_Z"
                    CreateTableT3 "GEMISCHTE_Z", gdBase
                Case "GEMISCHTE_ZP"
                    CreateTableT3 "GEMISCHTE_ZP", gdBase
                Case "TERMINE_ANL"
                    CreateTableT3 "TERMINE_ANL", gdBase
                Case "REPOS"
                    CreateTableT2 "REPOS", gdBase
                Case "KUNDENBONUS"
                    CreateTableT2 "KUNDENBONUS", gdBase
                Case "LISRT_MAIL"
                    CreateTableT2 "LISRT_MAIL", gdBase
                Case "ZUGRIFFDAT"
                    CreateTableT2 "ZUGRIFFDAT", gdBase
                Case "PREISEDITKASSE"
                    CreateTableT2 "PREISEDITKASSE", gdBase
                Case "DESADV"
                    CreateTableT2 "DESADV", gdBase
                Case "BONUS_SYS"
                    CreateTableT2 "BONUS_SYS", gdBase
                Case "STEUERKI"
                    CreateTableT2 "STEUERKI", gdBase
                Case "MAILFB"
                    CreateTableT2 "MAILFB", gdBase
                Case "ARTAUSWAHL"
                    CreateTable "ARTAUSWAHL", gdBase
                Case "ETISIC"
                    CreateTableT2 "ETISIC", gdBase
                Case "DISPLAYTEXT"
                    CreateTableT2 "DISPLAYTEXT", gdBase
                Case "GRUPPE"
                    CreateTableT2 "GRUPPE", db
                Case "GRUPPE_ARTIKEL"
                    CreateTableT2 "GRUPPE_ARTIKEL", db
                Case "GLAGER"
                    CreateTableT2 "GLAGER", db
                Case "WGNDEF"
                    CreateTableT2 "WGNDEF", db
                Case "NICHTUMSBAR"
                    CreateTableT2 "NICHTUMSBAR", db
                Case "WGNAGN"
                    CreateTableT2 "WGNAGN", db
                Case "WEBSHOP_E"
                    CreateTableT2 "WEBSHOP_E", db
                Case "GDLAGER"
                    CreateTableT2 "GDLAGER", db
                Case "MY" & srechnertab
                    CreateTable "MY" & srechnertab, gdBase
                Case "KUNDA" & srechnertab
                    CreateTableT2 "KUNDA" & srechnertab, gdBase
                Case "GUHIN"
                    CreateTableT2 "GUHIN", gdBase
                Case "UMS_LINR"
                    CreateTableT2 "UMS_LINR", gdBase
                Case "INTERART"
                    CreateTableT2 "INTERART", gdBase
                Case "ARTEAN_K"
                    CreateTableT2 "ARTEAN_K", gdBase
                Case "UMSATZINFO"
                    CreateTableT2 "UMSATZINFO", gdBase
                Case "GESCHWART"
                    CreateTableT2 "GESCHWART", gdBase
                Case "UMS_ARTNR"
                    CreateTableT2 "UMS_ARTNR", gdBase
                Case "STAFFELPR"
                    CreateTableT2 "STAFFELPR", gdBase
                Case "STAFFELPRKVK"
                    CreateTableT2 "STAFFELPRKVK", gdBase
                Case "STAFFEL_KVK_ARTIKEL"
                    CreateTableT2 "STAFFEL_KVK_ARTIKEL", gdBase
                Case "STAFFEL_KVK_GRUPPE"
                    CreateTableT2 "STAFFEL_KVK_GRUPPE", gdBase
                Case "UMS_LPZ"
                    CreateTableT2 "UMS_LPZ", gdBase
                Case "LUGEVER"
                    CreateTableT2 "LUGEVER", gdBase
                Case "LAGERW"
                    CreateTableT2 "LAGERW", gdBase
                Case "IDENTUSER"
                    CreateTableT2 "IDENTUSER", gdBase
                Case "DATEV"
                    CreateTableT2 "DATEV", gdBase
                Case "MBSTAND"
                    CreateTableT2 "MBSTAND", gdBase
                Case "LANG"
                    CreateTableT2 "LANG", gdBase
                Case "LAGERLW"
                    CreateTableT2 "LAGERLW", gdBase
                Case "STEMPELAZU"
                    CreateTableT2 "STEMPELAZU", gdBase
                Case "AZEITU"
                    CreateTableT2 "AZEITU", gdBase
                Case "PENLAGERMW"
                    CreateTableT2 "PENLAGERMW", gdBase
                Case "BONTEXT"
                    CreateTable "BONTEXT", gdBase
                Case "LAGERMW"
                    CreateTableT2 "LAGERMW", gdBase
                Case "LAGERLLW"
                    CreateTableT2 "LAGERLLW", gdBase
                Case "PENLAGERW"
                    CreateTableT2 "PENLAGERW", gdBase
                Case "PENLAGERLW"
                    CreateTableT2 "PENLAGERLW", gdBase
                Case "PENLAGERLLW"
                    CreateTableT2 "PENLAGERLLW", gdBase
                Case "KUDD"
                    CreateTableT2 "KUDD", gdBase
                Case "PREISE"
                    CreateTableT2 "PREISE", gdBase
                Case "UMLAGER"
                    CreateTableT2 "UMLAGER", gdBase
                Case "BONUSBONTEXTE"
                    CreateTableT2 "BONUSBONTEXTE", gdBase
                Case "BWWBONTEXTE"
                    CreateTableT2 "BWWBONTEXTE", gdBase
                Case "BONUSART"
                    CreateTableT2 "BONUSART", gdBase
                Case "GANALYSE"
                    CreateTableT2 "GANALYSE", gdBase
                Case "GANALYSEALL"
                    CreateTableT2 "GANALYSEALL", gdBase
                Case "MBORDER"
                    CreateTableT2 "MBORDER", gdBase
                Case "BEAUFNR"
                    CreateTable "BEAUFNR", gdBase
                Case "ARTMERK"
                    CreateTable "ARTMERK", gdBase
                Case "STORNOF"
                    CreateTableT2 "STORNOF", gdBase
                Case "STORNO2"
                    CreateTable "STORNO2", gdBase
                Case "ZUGANGF"
                    CreateTable "ZUGANGF", gdBase
                Case "KASSBEDP"
                    CreateTable "KASSBEDP", gdBase
                Case "ABSCHOPF"
                    CreateTable "ABSCHOPF", gdBase
                Case "DUKATENB"
                    CreateTableT2 "DUKATENB", gdBase
                Case "KABUCH"
                    CreateTable "KABUCH", gdBase
                Case "ETIDRULS"
                    CreateTableT2 "ETIDRULS", gdBase
                Case "BESTAKT"
                    CreateTable "BESTAKT", gdBase
                Case "GUHIS"
                    CreateTable "GUHIS", gdBase
                Case "EMAIL"
                    CreateTable "EMAIL", gdBase
                Case "KASSBOND"
                    CreateTable "KASSBOND", gdBase
                Case "WECHSEL"
                    CreateTable "WECHSEL", gdBase
                Case "FARBKU"
                    CreateTable "FARBKU", gdBase
                Case "FARBMERK"
                    CreateTable "FARBMERK", gdBase
                Case "DVKART"
                    CreateTable "DVKART", gdBase
                Case "KUNBHIST"
                    CreateTable "KUNBHIST", gdBase
                Case "ARTDET"
                    CreateTable "ARTDET", gdBase
                Case "ARTMDH"
                    CreateTableT2 "ARTMDH", gdBase
                Case "KUNDEDEL"
                    CreateTable "KUNDEDEL", gdBase
                Case "PROVISION"
                    CreateTable "PROVISION", gdBase
                Case "MARKIERUNG"
                    CreateTableT2 "MARKIERUNG", gdBase
                Case "PREISTERM"
                    CreateTable "PREISTERM", gdBase
                Case "KREDITZA"
                    CreateTable "KREDITZA", gdBase
                Case "ARTFARB2"
                    CreateTable "ARTFARB2", gdBase
                Case "PRSTERM"
                    CreateTable "PRSTERM", gdBase
                Case "ALTERG"
                    CreateTable "ALTERG", gdBase
                Case "RETPRINT"
                    CreateTable "RETPRINT", gdBase
                Case "REPARATUR"
                    CreateTable "REPARATUR", gdBase
                Case "REPARATURSATZ"
                    CreateTable "REPARATURSATZ", gdBase
                Case "REPARATURKOPF"
                    CreateTable "REPARATURKOPF", gdBase
                Case "ALLARTLUKOPF"
                    CreateTable "ALLARTLUKOPF", gdBase
                Case "ALLARTLU"
                    CreateTable "ALLARTLU", gdBase
                Case "OPENINGS"
                    CreateTable "OPENINGS", gdBase
                Case "GARANTIE"
                    CreateTable "GARANTIE", gdBase
                Case "AFCAUFBON"
                    CreateTable "AFCAUFBON", gdBase
                Case "DRUBRU"
                    CreateTable "DRUBRU", gdBase
                Case "ABWESEND"
                    CreateTable "ABWESEND", gdBase
                Case "DRUNET"
                    CreateTable "DRUNET", gdBase
                Case "LINRZUO"
                    CreateTable "LINRZUO", gdBase
                Case "GEMZ"
                    CreateTable "GEMZ", gdBase
                Case "UNTERWF"
                    CreateTable "UNTERWF", gdBase
                Case "RECHUE"
                    CreateTable "RECHUE", gdBase
                Case "MITKU" & srechnertab
                    CreateTable "MITKU" & srechnertab, gdBase
                Case "KONTIN"
                    CreateTable "KONTIN", gdBase
                Case "KULOESCH"
                    CreateTable "KULOESCH", gdBase
                Case "KAEINAUS"
                    CreateTable "KAEINAUS", gdBase
                Case "KUNPFLEG"
                    CreateTable "KUNPFLEG", gdBase
                Case "KUNPFLGH"
                    CreateTable "KUNPFLGH", gdBase
                Case "KAEINAUSF"
                    CreateTable "KAEINAUSF", gdBase
                Case "KARTEN_EINZ"
                    CreateTableT2 "KARTEN_EINZ", gdBase
                Case "EINAUSKB"
                    CreateTable "EINAUSKB", gdBase
                Case "KKZAHL"
                    CreateTable "KKZAHL", gdBase
                Case "FEEDB"
                    CreateTable "FEEDB", gdBase
                Case "FEEDB_TRANS"
                    CreateTableT2 "FEEDB_TRANS", gdBase
                Case "FEEDBF"
                    CreateTable "FEEDBF", gdBase
                Case "KKZAHLTE"
                    CreateTable "KKZAHLTE", gdBase
                Case "LASTZAHL"
                    CreateTable "LASTZAHL", gdBase
                Case "LASTZAHLTE"
                    CreateTable "LASTZAHLTE", gdBase
                Case "GUTZ"
                    CreateTable "GUTZ", gdBase
                Case "KASSEIN"
                    CreateTableT3 "KASSEIN", gdBase
                Case "STATIST"
                    CreateTable "STATIST", gdBase
                Case "ETIDRU"
                    CreateTable "ETIDRU", gdBase
                Case "BEDNAME"
                    CreateTable "BEDNAME", gdBase
                Case "ZENTHELP"
                    CreateTable "ZENTHELP", gdBase
                Case "NOTART"
                    CreateTable "NOTART", gdBase
                Case "KONDITIONEN"
                    CreateTable "KONDITIONEN", gdBase
                Case "TEXTBLOCK"
                    CreateTable "TEXTBLOCK", gdBase
                Case "WKVERSIONEN"
                    CreateTable "WKVERSIONEN", gdBase
                Case "NEINVK"
                    CreateTable "NEINVK", gdBase
                Case "KONDILEK"
                    CreateTable "KONDILEK", gdBase
                Case "LAGERPLATZ"
                    CreateTable "LAGERPLATZ", gdBase
                Case "TEXTIL"
                    CreateTableT2 "TEXTIL", gdBase
                Case "BANKKU"
                    CreateTable "BANKKU", gdBase
                Case "MERKFARB"
                    CreateTable "MERKFARB", gdBase
                Case "NOTIZEN"
                    CreateTable "NOTIZEN", gdBase
                Case "OPENINGS"
                    CreateTable "OPENINGS", gdBase
                Case "FEHLZEIT"
                    CreateTable "FEHLZEIT", gdBase
                Case "TERMINE"
                    CreateTable "TERMINE", gdBase
                Case "NOKALKL"
                    CreateTable "NOKALKL", gdBase
                Case "DELSTADAL"
                    CreateTable "DELSTADAL", gdBase
                Case "SPEZIINFO"
                    CreateTable "SPEZIINFO", gdBase
                Case "TERM_STD"
                    CreateTable "TERM_STD", gdBase
                Case "RECHAB"
                    CreateTable "RECHAB", gdBase
                Case "PFLEGORT"
                    CreateTable "PFLEGORT", gdBase
                Case "BEDTERM"
                    CreateTable "BEDTERM", gdBase
                Case "ZADRESS"
                    CreateTable "ZADRESS", gdBase
                Case "KASSWAAG"
                    CreateTable "KASSWAAG", gdBase
                Case "KUNDBEST"
                    CreateTable "KUNDBEST", gdBase
                Case "KUNDAUSLIEF"
                    CreateTable "KUNDAUSLIEF", gdBase
                Case "TABDATUM"
                    CreateTable "TABDATUM", gdBase
                Case "BEDZUGRI"
                    CreateTable "BEDZUGRI", gdBase
                Case "BANKEN"
                    CreateTable "BANKEN", gdBase
                Case "RETOURE"
                    CreateTable "RETOURE", gdBase
                Case "KOPFMAIL"
                    CreateTable "KOPFMAIL", gdBase
                Case "ARBEIT"
                    CreateTable "ARBEIT", gdBase
                Case "TAUSCH"
                    CreateTable "TAUSCH", gdBase
                Case "STEMPEL"
                    CreateTable "STEMPEL", gdBase
                Case "NOEURO"
                    CreateTable "NOEURO", gdBase
                Case "KOLLVERK"
                    CreateTable "KOLLVERK", gdBase
                Case "MASTERPRO"
                    CreateTable "MASTERPRO", gdBase
                Case "ABSCHLUSS"
                    CreateTable "ABSCHLUSS", gdBase
                Case "BARGELD"
                    CreateTable "BARGELD", gdBase
                Case "DTA"
                    CreateTable "DTA", gdBase
                Case "UMS_ART"
                    Ums_artNew Label2
                Case "UMSKDJ"
                    UmskdjNew Label2
                Case "UMSARTJ"
                    UmsartjNew Label2
                Case "UMS_ARTF"
                    CreateTable "UMS_ARTF", gdBase
                Case "ZUORDEAN"
                    CreateTable "ZUORDEAN", gdBase
                Case "UEBERLI"
                    CreateTable "UEBERLI", gdBase
                Case "KASSBON"
                    CreateTable "KASSBON", gdBase
                Case "KUNDKASS"
                    CreateTable "KUNDKASS", gdBase
                Case "KUK"
                    CreateTable "KUK", gdBase
                Case "BESTAEND"
                    CreateTable "BESTAEND", gdBase
                Case "BESTPROT"
                    CreateTable "BESTPROT", gdBase
                Case "EANPROT"
                    CreateTable "EANPROT", gdBase
                Case "KVKPR1PROT"
                    CreateTable "KVKPR1PROT", gdBase
                Case "SORTI"
                    CreateTable "SORTI", gdBase
                Case "KUMSUM"
                    CreateTableT2 "KUMSUM", gdBase
                Case "EKASS"
                    CreateTable "EKASS", gdBase
                Case "EAJ"
                    CreateTable "EAJ", gdBase
                Case "DABAUSER"
                    CreateTable "DABAUSER", gdBase
                Case "KASQL"
                    CreateTable "KASQL", gdBase
                Case "AFCSTAT"
                    CreateTable "AFCSTAT", gdBase
                Case "ZBONLAY"
                    CreateTable "ZBONLAY", gdBase
                Case "MA" & srechnertab
                    CreateTable "MA" & srechnertab, gdBase
                Case "TABALI" & srechnertab
                    CreateTable "TABALI" & srechnertab, gdBase
                Case "TABLAY" & srechnertab
                    loeschNEW "TABLAY" & srechnertab, gdBase
                    If Not NewTableSuchenDBKombi("TABLAY", gdBase) Then
                        CreateTable "TABLAY", gdBase
                    End If
                    sSQL = "Select * into TABLAY" & srechnertab & " from TABLAY"
                    gdBase.Execute sSQL, dbFailOnError
                    
                Case "GROLIEF"
                    CreateTable "GROLIEF", gdBase
                    
                    Insertgrolief "300500"
                    Insertgrolief "100000"
                Case Else
                
                    kassetabcheck = sTabellen(i)
                    Exit Function
            End Select
            
        Else
            kassetabcheck = ""
        End If
    Next i
    
    CheckIndex "UMS_ART", "Primkey", "ARTNR, JAHR, MONAT", gdBase
    CheckIndex "KopfMail", "HAUPTTEXT", "", gdBase
    
    If SpalteInTabellegefundenNEW("Retoure", "SENDOK", gdBase) = False Then
        SpalteAnfuegenNEW "Retoure", "SENDOK", "BIT", gdBase
    End If
    
    If SpalteInTabellegefundenNEW("AFCSTATP", "SENDOK", gdBase) = False Then
        SpalteAnfuegenNEW "AFCSTATP", "SENDOK", "BIT", gdBase
    End If
    
    If SpalteInTabellegefundenNEW("KONDITIONEN", "FAKTOR", gdBase) = False Then
        SpalteAnfuegenNEW "KONDITIONEN", "FAKTOR", "INTEGER", gdBase
    
        sSQL = "Update KONDITIONEN set FAKTOR = 1 "
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If SpalteInTabellegefundenNEW("EAJ", "FIL", gdBase) = False Then
       SpalteAnfuegenNEW "EAJ", "FIL", "BYTE", gdBase
    End If
    
    If SpalteInTabellegefundenNEW("EAJ", "BO1", gdBase) = False Then
        SpalteAnfuegenNEW "EAJ", "BO1", "bit", gdBase
       
        sSQL = "Update EAJ Set BO1 = true "
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    
    If SpalteInTabellegefundenNEW("LISRT", "BR", gdBase) = False Then
       SpalteAnfuegenNEW "LISRT", "BR", "BYTE", gdBase
    End If
    
    If SpalteInTabellegefundenNEW("LISRT", "AUFVJ", gdBase) Then
        sSQL = "Alter table LISRT DROP AUFVJ "
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If SpalteInTabellegefundenNEW("LISRT", "AUFLJ", gdBase) Then
        sSQL = "Alter table LISRT DROP AUFLJ "
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If SpalteInTabellegefundenNEW("LISRT", "MINAUF", gdBase) Then
        sSQL = "Alter table LISRT DROP MINAUF "
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If SpalteInTabellegefundenNEW("ZUORDEAN", "ARTNR", gdBase) = False Then
       SpalteAnfuegenNEW "ZUORDEAN", "ARTNR", "LONG", gdBase
    End If
    
    If SpalteInTabellegefundenNEW("KASSBON", "KK_ART", gdBase) = False Then
        SpalteAnfuegenNEW "KASSBON", "KK_ART", "TEXT(2)", gdBase
       
        sSQL = "Update KASSBON Set KK_ART =  'UB'"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If SpalteInTabellegefundenNEW("KASSBON", "KUNDNR", gdBase) = False Then
        SpalteAnfuegenNEW "KASSBON", "KUNDNR", "LONG", gdBase
       
        sSQL = "Update KASSBON Set KUNDNR =  0 "
        gdBase.Execute sSQL, dbFailOnError
    End If
        
    If SpalteInTabellegefundenNEW("KASSBON", "SENDOK", gdBase) = False Then
        SpalteAnfuegenNEW "KASSBON", "SENDOK", "BIT", gdBase
       
        sSQL = "Update KASSBON Set SENDOK = false "
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If SpalteInTabellegefundenNEW("KASSBON", "FILIALE", gdBase) = False Then
        SpalteAnfuegenNEW "KASSBON", "FILIALE", "BYTE", gdBase
       
        sSQL = "Update KASSBON Set FILIALE =  " & Val(gcFilNr)
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    Exit Function
LOKAL_ERROR:
    If err.Number = 3011 Then
        kassetabcheck = sTabellen(i)
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "kassetabcheck"
        Fehler.gsFehlertext = "Beim Überprüfen der Tabellen ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Function
Public Function Journal_leeren() As String
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    
    If NewTableSuchenDBKombi("KASSJOUR", gdBase) Then
        sSQL = "Delete from KASSJOUR"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If NewTableSuchenDBKombi("KREDIT", gdBase) Then
        sSQL = "Delete from KREDIT"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If NewTableSuchenDBKombi("AFCSTAT", gdBase) Then
        sSQL = "Delete from AFCSTAT"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If NewTableSuchenDBKombi("AFCSTATP", gdBase) Then
        sSQL = "Delete from AFCSTATP"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If NewTableSuchenDBKombi("AFCBUCH", gdBase) Then
        sSQL = "Delete from AFCBUCH"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If NewTableSuchenDBKombi("UMSATZ", gdBase) Then
        sSQL = "Delete from UMSATZ"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If NewTableSuchenDBKombi("KABUCH", gdBase) Then
        sSQL = "Delete from KABUCH"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If NewTableSuchenDBKombi("ZBESTAND", gdBase) Then
        sSQL = "Delete from ZBESTAND"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If NewTableSuchenDBKombi("BESTAEND", gdBase) Then
        sSQL = "Delete from BESTAEND"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If NewTableSuchenDBKombi("BESTPROT", gdBase) Then
        sSQL = "Delete from BESTPROT"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If NewTableSuchenDBKombi("BESTREST", gdBase) Then
        sSQL = "Delete from BESTREST"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If NewTableSuchenDBKombi("GDLAGER", gdBase) Then
        sSQL = "Delete from GDLAGER"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If NewTableSuchenDBKombi("ZUGANG", gdBase) Then
        sSQL = "Delete from ZUGANG"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If NewTableSuchenDBKombi("GLAGER", gdBase) Then
        sSQL = "Delete from GLAGER"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If NewTableSuchenDBKombi("KASSBOND", gdBase) Then
        sSQL = "Delete from KASSBOND"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If NewTableSuchenDBKombi("KUANTE", gdBase) Then
        sSQL = "Delete from KUANTE"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If NewTableSuchenDBKombi("KKZAHL", gdBase) Then
        sSQL = "Delete from KKZAHL"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If NewTableSuchenDBKombi("BARGELD", gdBase) Then
        sSQL = "Delete from BARGELD"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If NewTableSuchenDBKombi("EANPROT", gdBase) Then
        sSQL = "Delete from EANPROT"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If NewTableSuchenDBKombi("EINAUSKB", gdBase) Then
        sSQL = "Delete from EINAUSKB"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If NewTableSuchenDBKombi("ETISIC", gdBase) Then
        sSQL = "Delete from ETISIC"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If NewTableSuchenDBKombi("GUTSCH", gdBase) Then
        sSQL = "Delete from GUTSCH"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If NewTableSuchenDBKombi("INTERART", gdBase) Then
        sSQL = "Delete from INTERART"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If NewTableSuchenDBKombi("KASSBEDP", gdBase) Then
        sSQL = "Delete from KASSBEDP"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If NewTableSuchenDBKombi("KASSBON", gdBase) Then
        sSQL = "Delete from KASSBON"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If NewTableSuchenDBKombi("KUNDEN", gdBase) Then
        sSQL = "Delete from KUNDEN"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If NewTableSuchenDBKombi("BEDNAME", gdBase) Then
        sSQL = "Delete from BEDNAME"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If NewTableSuchenDBKombi("OFPO", gdBase) Then
        sSQL = "Delete from OFPO"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If NewTableSuchenDBKombi("REPOS", gdBase) Then
        sSQL = "Delete from REPOS"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If NewTableSuchenDBKombi("REKOPF", gdBase) Then
        sSQL = "Delete from REKOPF"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If NewTableSuchenDBKombi("KVKPR1PROT", gdBase) Then
        sSQL = "Delete from KVKPR1PROT"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If NewTableSuchenDBKombi("LAGERLLW", gdBase) Then
        sSQL = "Delete from LAGERLLW"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If NewTableSuchenDBKombi("LAGERMW", gdBase) Then
        sSQL = "Delete from LAGERMW"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If NewTableSuchenDBKombi("LAGERLW", gdBase) Then
        sSQL = "Delete from LAGERLW"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If NewTableSuchenDBKombi("LAGERW", gdBase) Then
        sSQL = "Delete from LAGERW"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If NewTableSuchenDBKombi("PENLAGERLLW", gdBase) Then
        sSQL = "Delete from PENLAGERLLW"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If NewTableSuchenDBKombi("PENLAGERLW", gdBase) Then
        sSQL = "Delete from PENLAGERLW"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If NewTableSuchenDBKombi("PENLAGERW", gdBase) Then
        sSQL = "Delete from PENLAGERW"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If NewTableSuchenDBKombi("PRSTERM", gdBase) Then
        sSQL = "Delete from PRSTERM"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If NewTableSuchenDBKombi("RETOURE", gdBase) Then
        sSQL = "Delete from RETOURE"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If NewTableSuchenDBKombi("UMSARTJ", gdBase) Then
        sSQL = "Delete from UMSARTJ"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If NewTableSuchenDBKombi("WECHSEL", gdBase) Then
        sSQL = "Delete from WECHSEL"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If NewTableSuchenDBKombi("ZUGANGF", gdBase) Then
        sSQL = "Delete from ZUGANGF"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If NewTableSuchenDBKombi("ZUORDEAN", gdBase) Then
        sSQL = "Delete from ZUORDEAN"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If NewTableSuchenDBKombi("BONTEXT", gdBase) Then
        sSQL = "Delete from BONTEXT"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    If NewTableSuchenDBKombi("STEUERKI", gdBase) Then
        sSQL = "Delete from STEUERKI"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    Exit Function
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "Journal_leeren"
    Fehler.gsFehlertext = "Beim Überprüfen der Tabellen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Public Function aufräumen() As String
On Error GoTo LOKAL_ERROR
    
    
    loeschNEW "Artikel2", gdBase
    
    Dim sSQL        As String
    
    sSQL = "Delete * from FEEDB "
    gdBase.Execute sSQL, dbFailOnError
    
    
    '>>>>>>>>>>>>>>>>>>>>>>>>
    
    loeschNEW "LEB_Artikel", gdBase
    loeschNEW "LEB_Artlief", gdBase
    
    loeschNEW "ZBESTAND_BAK", gdBase
    loeschNEW "FEEDBF", gdBase
    loeschNEW "CTIKEL", gdBase
    loeschNEW "KUTERT", gdBase
    loeschNEW "Z_OUT", gdBase
    
    
    
    loeschNEW "ARTKUM", gdBase
    loeschNEW "ARTNURBEST", gdBase
    loeschNEW "B344", gdBase
    loeschNEW "BTOINV", gdBase
    loeschNEW "ETIPROTS", gdBase
    loeschNEW "ETIPROT", gdBase
    loeschNEW "TEMPATAB", gdBase

    
    'alle Tabellen, die mit AB beginnen
    
    Dim lcount      As Long
    Dim lAnzTable   As Long
    
    Dim name        As String
    
    
    
    
    
    
    
    
    
    
    'alle Tabellen, die mit INV beginnen
    lAnzTable = gdBase.TableDefs.Count
    For lcount = 0 To lAnzTable - 1
        name = gdBase.TableDefs(lcount).name
        If Left(UCase$(name), 3) = "INV" Then
        
            
            sSQL = "drop table " & name
            gdBase.Execute sSQL, dbFailOnError
            
        End If
    Next lcount
    
    'alle Tabellen, die mit KUNDA beginnen
    lAnzTable = gdBase.TableDefs.Count
    For lcount = 0 To lAnzTable - 1
        name = gdBase.TableDefs(lcount).name
        If Left(UCase$(name), 5) = "KUNDA" Then
        
            
            sSQL = "drop table " & name
            gdBase.Execute sSQL, dbFailOnError
            
        End If
    Next lcount
    
    'alle Tabellen, die mit KP beginnen
    lAnzTable = gdBase.TableDefs.Count
    For lcount = 0 To lAnzTable - 1
        name = gdBase.TableDefs(lcount).name
        If Left(UCase$(name), 2) = "KP" Then
        
            
            sSQL = "drop table " & name
            gdBase.Execute sSQL, dbFailOnError
            
        End If
    Next lcount
    
    'alle Tabellen, die mit KJ beginnen
    lAnzTable = gdBase.TableDefs.Count
    For lcount = 0 To lAnzTable - 1
        name = gdBase.TableDefs(lcount).name
        If Left(UCase$(name), 2) = "KJ" Then
        
            
            sSQL = "drop table " & name
            gdBase.Execute sSQL, dbFailOnError
            
        End If
    Next lcount
    
    'alle Tabellen, die mit MA beginnen
    lAnzTable = gdBase.TableDefs.Count
    For lcount = 0 To lAnzTable - 1
        name = gdBase.TableDefs(lcount).name
        If Left(UCase$(name), 2) = "MA" Then
        
            
            sSQL = "drop table " & name
            gdBase.Execute sSQL, dbFailOnError
            
        End If
    Next lcount
    
''''    'alle Tabellen, die mit N0 beginnen
''''    lAnzTable = gdBase.TableDefs.Count
''''    For lcount = 0 To lAnzTable - 1
''''        name = gdBase.TableDefs(lcount).name
''''        If Left(UCase$(name), 2) = "N0" Then
''''
''''
''''            sSQL = "drop table " & name
''''            gdBase.Execute sSQL, dbFailOnError
''''
''''        End If
''''    Next lcount
    
    'alle Tabellen, die mit MY beginnen
    lAnzTable = gdBase.TableDefs.Count
    For lcount = 0 To lAnzTable - 1
        name = gdBase.TableDefs(lcount).name
        If Left(UCase$(name), 2) = "MY" Then
        
            
            sSQL = "drop table " & name
            gdBase.Execute sSQL, dbFailOnError
            
        End If
    Next lcount
    
    'alle Tabellen, die mit TB beginnen
    lAnzTable = gdBase.TableDefs.Count
    For lcount = 0 To lAnzTable - 1
        name = gdBase.TableDefs(lcount).name
        If Left(UCase$(name), 2) = "TB" Then
        
            
            sSQL = "drop table " & name
            gdBase.Execute sSQL, dbFailOnError
            
        End If
    Next lcount
    
    'alle Tabellen, die mit TOPI beginnen
    lAnzTable = gdBase.TableDefs.Count
    For lcount = 0 To lAnzTable - 1
        name = gdBase.TableDefs(lcount).name
        If Left(UCase$(name), 4) = "TOPI" Then
        
            
            sSQL = "drop table " & name
            gdBase.Execute sSQL, dbFailOnError
            
        End If
    Next lcount
    
    'alle Tabellen, die mit MITKU beginnen
    lAnzTable = gdBase.TableDefs.Count
    For lcount = 0 To lAnzTable - 1
        name = gdBase.TableDefs(lcount).name
        If Left(UCase$(name), 5) = "MITKU" Then
        
            
            sSQL = "drop table " & name
            gdBase.Execute sSQL, dbFailOnError
            
        End If
    Next lcount
    
    'alle Tabellen, die mit AFCB beginnen
    lAnzTable = gdBase.TableDefs.Count
    For lcount = 0 To lAnzTable - 1
        name = gdBase.TableDefs(lcount).name
        If Left(UCase$(name), 4) = "AFCB" Then
        
            If UCase$(name) = "AFCBUCH" Then
            
            Else
                sSQL = "drop table " & name
                gdBase.Execute sSQL, dbFailOnError
            End If
            
        End If
    Next lcount
    
    'alle Tabellen, die mit NEUKU beginnen
    lAnzTable = gdBase.TableDefs.Count
    For lcount = 0 To lAnzTable - 1
        name = gdBase.TableDefs(lcount).name
        If Left(UCase$(name), 5) = "NEUKU" Then
        
            
            sSQL = "drop table " & name
            gdBase.Execute sSQL, dbFailOnError
            
        End If
    Next lcount
    
    'alle Tabellen, die mit SCHWPUNKT beginnen
    lAnzTable = gdBase.TableDefs.Count
    For lcount = 0 To lAnzTable - 1
        name = gdBase.TableDefs(lcount).name
        If Left(UCase$(name), 9) = "SCHWPUNKT" Then
        
            
            sSQL = "drop table " & name
            gdBase.Execute sSQL, dbFailOnError
            
        End If
    Next lcount
    
    Exit Function
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "aufräumen"
    Fehler.gsFehlertext = "Beim Überprüfen der Tabellen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Public Function Werkseinstellungen() As String
On Error GoTo LOKAL_ERROR

    Dim rsrs As DAO.Recordset
    Dim sSQL As String

    'Tabfaktor = 1.3
    
    gdTabfak = 1.3
   
    Set rsrs = gdApp.OpenRecordset("WKEINSTE", dbOpenTable)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        rsrs.Edit
        rsrs!Tabfak = gdTabfak
        rsrs.Update
    End If
    rsrs.Close: Set rsrs = Nothing
    
    sSQL = "Update WKEINSTE Set WeEinzFo = 'EAN' "
    gdApp.Execute sSQL, dbFailOnError
    gsWeEinzFo = "EAN"
    
    
    gbPenner_faerben = False
    sSQL = "Update DBEINSTE Set PENNERFARB = false "
    gdBase.Execute sSQL, dbFailOnError
    
    gbDabakompautoNo = True
    sSQL = "Update WKEINSTE Set NOAUTO = true "
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "Update KASSEIN Set AUSBLDU = True"
    gdBase.Execute sSQL, dbFailOnError
    gbAUSBLDU = True
     
    sSQL = "Update KASSEIN Set AUSBLsh = True"
    gdBase.Execute sSQL, dbFailOnError
    gbAUSBLSH = True
     
    sSQL = "Update KASSEIN Set AUSBLls = True"
    gdBase.Execute sSQL, dbFailOnError
    gbAUSBLLS = True
    
    If Not NewTableSuchenDBKombi("BUTTON", gdBase) Then
        CreateTable "BUTTON", gdBase
    End If
    
    sSQL = "Delete from Button where indexnr = 1"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from Button where indexnr = 9"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into Button (indexnr,buttonnr,buttontext) values ( 1,0,'EC Last')"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into Button (indexnr,buttonnr,buttontext) values ( 9,13,'Prov.')"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update WKEINSTE Set BAREIN = True"
    gdApp.Execute sSQL, dbFailOnError
    gbBargeldEingabe = True
     
    sSQL = "Update WKEINSTE Set QZBON = true"
    gdApp.Execute sSQL, dbFailOnError
    gbQZBON = True
    
    sSQL = "Update ABREPORT Set kk = true "
    gdApp.Execute sSQL, dbFailOnError
    gbKK = True
    
    sSQL = "Update ABREPORT Set ea = true "
    gdApp.Execute sSQL, dbFailOnError
    gbEA = True
    
    speicherZbon "Listendrucker"
    gsZBon = "Listendrucker"
   
    speicherZählBeleg "Listendrucker"
    gsZählbeleg = "Listendrucker"
    
    
    
    
    'Register Voreinstellungen
    
    sSQL = "Update DBEINSTE Set NewArt = True"
    gdBase.Execute sSQL, dbFailOnError
    gbNewArt = True
    
    sSQL = "Update DBEINSTE Set NewArtNrVorschlag = True"
    gdBase.Execute sSQL, dbFailOnError
    gbNewArtNrVorschlag = True
    
    sSQL = "Update DBEINSTE Set ARTEINDEUT = True"
    gdBase.Execute sSQL, dbFailOnError
    gbArtEindeut = True
    
    
    sSQL = "Update DBEINSTE Set HaGuNr = true "
    gdBase.Execute sSQL, dbFailOnError
    gbGutsch = True

    sSQL = "Update DBEINSTE Set OGV = true"
    gdBase.Execute sSQL, dbFailOnError
    gbOGV = True

    sSQL = "Update DBEINSTE Set RGO = true"
    gdBase.Execute sSQL, dbFailOnError
    gbRGO = True
    
    sSQL = "Update DBEINSTE Set OLDSTADADEL = True "
    gdBase.Execute sSQL, dbFailOnError
    gbOLDSTADADEL = True
    
    sSQL = "Update WKEINSTE Set BILDTAST = true"
    gdApp.Execute sSQL, dbFailOnError
    gbBILDTAST = True
    
    sSQL = "Update DBEINSTE Set BonusBNB = true "
    gdBase.Execute sSQL, dbFailOnError
    gbBonusBNB = True
    
    
    sSQL = "Update WKEINSTE Set QPASS = True"
    gdApp.Execute sSQL, dbFailOnError
        
    gbQPASS = True
        
    sSQL = "Update WKEINSTE Set BEDKARTE = False"
    gdApp.Execute sSQL, dbFailOnError
        
    sSQL = "Update DBEINSTE Set BEDKARTE = False"
    gdBase.Execute sSQL, dbFailOnError

    gbBEDKARTE = False
    
    
    sSQL = "Update DBEINSTE Set KuImBoY = true "
    gdBase.Execute sSQL, dbFailOnError
    gbKUNDENA = True
    
    
    loeschNEW "KUIBON", gdBase
    CreateTableT2 "KUIBON", gdBase
    
    
    sSQL = "Insert into KUIBON (Name) values (true) "
    gdBase.Execute sSQL, dbFailOnError
 
    gbKUIBONname = True
 
 

    sSQL = "Update KUIBON Set vorname = true "
    gdBase.Execute sSQL, dbFailOnError

    gbKUIBONvorname = True

 
    sSQL = "Update KUIBON Set firma = true "
    gdBase.Execute sSQL, dbFailOnError

    gbKUIBONfirma = True

 
     sSQL = "Update KUIBON Set titel = true "
     gdBase.Execute sSQL, dbFailOnError

     gbKUIBONtitel = True


     sSQL = "Update KUIBON Set strasse = true "
     gdBase.Execute sSQL, dbFailOnError

     gbKUIBONstrasse = True


     sSQL = "Update KUIBON Set plz = true "
     gdBase.Execute sSQL, dbFailOnError

     gbKUIBONplz = True


     sSQL = "Update KUIBON Set ort = true "
     gdBase.Execute sSQL, dbFailOnError

     gbKUIBONort = True

 
 
 
     sSQL = "Update KUIBON Set tel = true "
     gdBase.Execute sSQL, dbFailOnError

     gbKUIBONtel = True

 
     sSQL = "Update KUIBON Set mobil = true "
     gdBase.Execute sSQL, dbFailOnError

     gbKUIBONmobil = True

    
    
    
    
     
    sSQL = "Update WKEINSTE Set BONGUVK = true "
    gdApp.Execute sSQL, dbFailOnError
    gb2BONGUVK = True
    
    
   
    sSQL = "Update WKEINSTE Set BONEA = true "
    gdApp.Execute sSQL, dbFailOnError
    gb2BONEA = True
    
    sSQL = "Update WKEINSTE Set BONKR = true "
    gdApp.Execute sSQL, dbFailOnError
    gb2BONKR = True
    
    sSQL = "Update WKEINSTE Set BONKB = true "
    gdApp.Execute sSQL, dbFailOnError
    gb2BONKB = True
    
    sSQL = "Update WKEINSTE Set BONKOPIE = true "
    gdApp.Execute sSQL, dbFailOnError
    gbBonkopie = True
    
    sSQL = "Update KASSEIN Set Rabatt = true"
    gdBase.Execute sSQL, dbFailOnError
    gbRabatt = True
    
    
    sSQL = "Update KASSEIN Set PL = true"
    gdBase.Execute sSQL, dbFailOnError
    gbPrintLOGO = True
    
    loeschNEW "LOGOS", gdBase
    CreateTable "LOGOS", gdBase

    
    sSQL = "Insert into LOGOS (LOGO1) values (true)"
    gdBase.Execute sSQL, dbFailOnError
    gbLOGO1 = True



    sSQL = "Update LOGOS Set LOGO2 = False "
    gdBase.Execute sSQL, dbFailOnError
    gbLOGO2 = False

    sSQL = "Update LOGOS Set LOGO3 = False "
    gdBase.Execute sSQL, dbFailOnError
    gbLOGO3 = False
    
    
    sSQL = "Update WKEINSTE Set PARK = true"
    gdApp.Execute sSQL, dbFailOnError
    gbPark = True
    
    
    sSQL = "Update WKEINSTE Set UmsAnz = true"
    gdApp.Execute sSQL, dbFailOnError
    gbUmsAnz = True
    
    
    sSQL = "Update WKEINSTE Set BONNEIN = true"
    gdApp.Execute sSQL, dbFailOnError
    gbBONNEIN = True
    
    
    
    sSQL = "Update KASSEIN Set BONWAHL = true"
    gdBase.Execute sSQL, dbFailOnError
    gbBONWAHL = True
    
    sSQL = "Update KASSEIN Set OpenSchubRetoure = True"
    gdBase.Execute sSQL, dbFailOnError
    gbOpenSchubRetoure = True
    
    sSQL = "Update WKEINSTE Set EDITKASSNR = true"
    gdApp.Execute sSQL, dbFailOnError
    gbEDITKASSNR = True
    
    sSQL = "Update WKEINSTE Set DINA4VIS = true"
    gdApp.Execute sSQL, dbFailOnError
    gbDINA4VIS = True
    
    sSQL = "Update WKEINSTE Set BARDINA4 = False"
    gdApp.Execute sSQL, dbFailOnError
    gbBARDINA4 = False
    
    sSQL = "Update KASSEIN Set MBBLOCKFrage = False"
    gdBase.Execute sSQL, dbFailOnError
    gbMBBLOCKFrage = False
    
    gdBonusGrenze = 0
    sSQL = "Update DBEINSTE Set BONUSGRENZ = " & gdBonusGrenze
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    Exit Function
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "Werkseinstellungen"
    Fehler.gsFehlertext = "Beim Überprüfen der Tabellen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Public Sub db_Ballast_Tabellen_del()
On Error GoTo LOKAL_ERROR

    Dim i As Integer
    Dim j As Integer
    
    Dim sTabellen(0 To 204) As String
    j = 0
    
    'linbez, zugang , zzz,REPOS,umsatz,afcstatp, proto
    
    sTabellen(j) = "MAILFB"
    j = j + 1
    sTabellen(j) = "WGNDEF"
    j = j + 1
    sTabellen(j) = "WGNAGN"
    j = j + 1
    sTabellen(j) = "PROTO"
    j = j + 1
    sTabellen(j) = "UMSATZ"
    j = j + 1
    sTabellen(j) = "STEUERKI"
    j = j + 1
    sTabellen(j) = "REPOS"
    j = j + 1
    sTabellen(j) = "REKOPF"
    j = j + 1
    sTabellen(j) = "OFPO"
    j = j + 1
    sTabellen(j) = "ETISIC"
    j = j + 1
    sTabellen(j) = "GUHIN"
    j = j + 1
    sTabellen(j) = "GRUPPE"
    j = j + 1
    sTabellen(j) = "GRUPPE_ARTIKEL"
    j = j + 1
    sTabellen(j) = "INTERART"
    j = j + 1
    sTabellen(j) = "GESCHWART"
    j = j + 1
    sTabellen(j) = "STAFFELPR"
    j = j + 1
    sTabellen(j) = "LANG"
    j = j + 1
    sTabellen(j) = "DATEV"
    j = j + 1
    sTabellen(j) = "MBSTAND"
    j = j + 1
    sTabellen(j) = "GDLAGER"
    j = j + 1
    sTabellen(j) = "GLAGER"
    j = j + 1
    sTabellen(j) = "KUDD"
    j = j + 1
    sTabellen(j) = "IDENTUSER"
    j = j + 1
    sTabellen(j) = "UMLAGER"
    j = j + 1
    sTabellen(j) = "BONUSBONTEXTE"
    j = j + 1
    sTabellen(j) = "BWWBONTEXTE"
    j = j + 1
    sTabellen(j) = "BONUSART"
    j = j + 1
    sTabellen(j) = "UMS_LINR"
    j = j + 1
    sTabellen(j) = "UMS_ARTNR"
    j = j + 1
    sTabellen(j) = "UMS_LPZ"
    j = j + 1
    sTabellen(j) = "PREISE"
    j = j + 1
    sTabellen(j) = "MBORDER"
    j = j + 1
    sTabellen(j) = "DVKART"
    j = j + 1
    sTabellen(j) = "LUGEVER"
    j = j + 1
    sTabellen(j) = "LAGERW"
    j = j + 1
    sTabellen(j) = "LAGERMW"
    j = j + 1
    sTabellen(j) = "PENLAGERMW"
    j = j + 1
    sTabellen(j) = "LAGERLW"
    j = j + 1
    sTabellen(j) = "LAGERLLW"
    j = j + 1
    sTabellen(j) = "PENLAGERW"
    j = j + 1
    sTabellen(j) = "PENLAGERLW"
    j = j + 1
    sTabellen(j) = "PENLAGERLLW"
    j = j + 1
    sTabellen(j) = "ALLARTLUKOPF"
    j = j + 1
    sTabellen(j) = "GARANTIE"
    j = j + 1
    sTabellen(j) = "ALLARTLU"
    j = j + 1
    sTabellen(j) = "ARTMERK"
    j = j + 1
    sTabellen(j) = "KONTIN"
    j = j + 1
    sTabellen(j) = "MITKU" & srechnertab
    j = j + 1
    sTabellen(j) = "KUNDA" & srechnertab
    j = j + 1
    sTabellen(j) = "RECHUE"
    j = j + 1
    sTabellen(j) = "FARBMERK"
    j = j + 1
    sTabellen(j) = "EMAIL"
    j = j + 1
    sTabellen(j) = "WECHSEL"
    j = j + 1
    sTabellen(j) = "MASTERPRO"
    j = j + 1
    sTabellen(j) = "ARTFARB2"
    j = j + 1
    sTabellen(j) = "BESTAKT"
    j = j + 1
    sTabellen(j) = "KUNPFLEG"
    j = j + 1
    sTabellen(j) = "ZUGANGF"
    j = j + 1
    sTabellen(j) = "ZUGANG"
    j = j + 1
    sTabellen(j) = "ABWESEND"
    j = j + 1
    sTabellen(j) = "ETIDRULS"
    j = j + 1
    sTabellen(j) = "REPARATUR"
    j = j + 1
    sTabellen(j) = "REPARATURSATZ"
    j = j + 1
    sTabellen(j) = "REPARATURKOPF"
    j = j + 1
    sTabellen(j) = "KUNPFLGH"
    j = j + 1
    sTabellen(j) = "KUNDEDEL"
    j = j + 1
    sTabellen(j) = "LINBEZ"
    j = j + 1
    sTabellen(j) = "KUNBHIST"
    j = j + 1
    sTabellen(j) = "KABUCH"
    j = j + 1
    sTabellen(j) = "LINRZUO"
    j = j + 1
    sTabellen(j) = "PROVISION"
    j = j + 1
    sTabellen(j) = "NEINVK"
    j = j + 1
    sTabellen(j) = "AFCAUFBON"
    j = j + 1
    sTabellen(j) = "ALTERG"
    j = j + 1
    sTabellen(j) = "KASSBEDP"
    j = j + 1
    sTabellen(j) = "KASSEIN"
    j = j + 1
    sTabellen(j) = "DRUBRU"
    j = j + 1
    sTabellen(j) = "RETPRINT"
    j = j + 1
    sTabellen(j) = "DRUNET"
    j = j + 1
    sTabellen(j) = "ETI" & srechnertab
    j = j + 1
    sTabellen(j) = "MY" & srechnertab
    j = j + 1
    sTabellen(j) = "ZENTHELP"
    j = j + 1
    sTabellen(j) = "NOTART"
    j = j + 1
    sTabellen(j) = "GEMZ"
    j = j + 1
    sTabellen(j) = "ABSCHOPF"
    j = j + 1
    sTabellen(j) = "STORNO2"
    j = j + 1
    sTabellen(j) = "ARTDET"
    j = j + 1
    sTabellen(j) = "KKZAHL"
    j = j + 1
    sTabellen(j) = "KKZAHLTE"
    j = j + 1
    sTabellen(j) = "KREDITZA"
    j = j + 1
    sTabellen(j) = "FEEDB"
    j = j + 1
    sTabellen(j) = "FEEDBF"
    j = j + 1
    sTabellen(j) = "LASTZAHL"
    j = j + 1
    sTabellen(j) = "OPENINGS"
    j = j + 1
    sTabellen(j) = "LASTZAHLTE"
    j = j + 1
    sTabellen(j) = "GUTZ"
    j = j + 1
    sTabellen(j) = "GUHIS"
    j = j + 1
    sTabellen(j) = "DELSTADAL"
    j = j + 1
    sTabellen(j) = "NOKALKL"
    j = j + 1
    sTabellen(j) = "KOPFMAIL"
    j = j + 1
    sTabellen(j) = "STATIST"
    j = j + 1
    sTabellen(j) = "GROLIEF"
    j = j + 1
    sTabellen(j) = "ETIDRU"
    j = j + 1
    sTabellen(j) = "BEDNAME"
    j = j + 1
    sTabellen(j) = "WKVERSIONEN"
    j = j + 1
    sTabellen(j) = "TEXTBLOCK"
    j = j + 1
    sTabellen(j) = "KUNDBEST"
    j = j + 1
    sTabellen(j) = "KUNDEN"
    j = j + 1
    sTabellen(j) = "KUNDAUSLIEF"
    j = j + 1
    sTabellen(j) = "KONDITIONEN"
    j = j + 1
    sTabellen(j) = "PREISTERM"
    j = j + 1
    sTabellen(j) = "KONDILEK"
    j = j + 1
    sTabellen(j) = "BARGELD"
    j = j + 1
    sTabellen(j) = "DTA"
    j = j + 1
    sTabellen(j) = "TEXTIL"
    j = j + 1
    sTabellen(j) = "LAGERPLATZ"
    j = j + 1
    sTabellen(j) = "MERKFARB"
    j = j + 1
    sTabellen(j) = "TERMINE"
    j = j + 1
    sTabellen(j) = "SPEZIINFO"
    j = j + 1
    sTabellen(j) = "TERM_STD"
    j = j + 1
    sTabellen(j) = "ZBONLAY"
    j = j + 1
    sTabellen(j) = "FEHLZEIT"
    j = j + 1
    sTabellen(j) = "OPENINGS"
    j = j + 1
    sTabellen(j) = "PFLEGORT"
    j = j + 1
    sTabellen(j) = "BEDTERM"
    j = j + 1
    sTabellen(j) = "TABLAY" & srechnertab
    j = j + 1
    sTabellen(j) = "RECHAB"
    j = j + 1
    sTabellen(j) = "KASSWAAG"
    j = j + 1
    sTabellen(j) = "KASQL"
    j = j + 1
    sTabellen(j) = "EKASS"
    j = j + 1
    sTabellen(j) = "EAJ"
    j = j + 1
    sTabellen(j) = "STEMPEL"
    j = j + 1
    sTabellen(j) = "NOEURO"
    j = j + 1
    sTabellen(j) = "KUK"
    j = j + 1
    sTabellen(j) = "TABCODE"
    j = j + 1
    sTabellen(j) = "SORTI"
    j = j + 1
    sTabellen(j) = "ARTIKEL"
    j = j + 1
    sTabellen(j) = "PGNDBF"
    j = j + 1
    sTabellen(j) = "AGNDBF"
    j = j + 1
    sTabellen(j) = "ARBEIT"
    j = j + 1
    sTabellen(j) = "TAUSCH"
    j = j + 1
    sTabellen(j) = "ARTLIEF"
    j = j + 1
    sTabellen(j) = "Warengru"
    j = j + 1
    sTabellen(j) = "BANKEN"
    j = j + 1
    sTabellen(j) = "BEDNAME"
    j = j + 1
    sTabellen(j) = "BESTREST"
    j = j + 1
    sTabellen(j) = "BONPAUSE"
    j = j + 1
    sTabellen(j) = "ARTAUSWAHL"
    j = j + 1
    sTabellen(j) = "BONTEXT"
    j = j + 1
    sTabellen(j) = "BONUSGRE"
    j = j + 1
    sTabellen(j) = "BANKKU"
    j = j + 1
    sTabellen(j) = "ZUORDEAN"
    j = j + 1
    sTabellen(j) = "DBEINSTE"
    j = j + 1
    sTabellen(j) = "DTA"
    j = j + 1
    sTabellen(j) = "FILIALEN"
    j = j + 1
    sTabellen(j) = "Firma"
    j = j + 1
    sTabellen(j) = "FILA"
    j = j + 1
    sTabellen(j) = "GUTSCH"
    j = j + 1
    sTabellen(j) = "KAEINAUS"
    j = j + 1
    sTabellen(j) = "KARTEN_EINZ"
    j = j + 1
    sTabellen(j) = "KAEINAUSF"
    j = j + 1
    sTabellen(j) = "EINAUSKB"
    j = j + 1
    sTabellen(j) = "PRSTERM"
    j = j + 1
    sTabellen(j) = "ZBESTAND"
    j = j + 1
    sTabellen(j) = "KISSLITE"
    j = j + 1
    sTabellen(j) = "UMS_ART"
    j = j + 1
    sTabellen(j) = "UMSARTJ"
    j = j + 1
    sTabellen(j) = "UMS_ARTF"
    j = j + 1
    sTabellen(j) = "UMSKDJ"
    j = j + 1
    sTabellen(j) = "LISRT"
    j = j + 1
    sTabellen(j) = "MWSTSATZ"
    j = j + 1
    sTabellen(j) = "UMS_KD"
    j = j + 1
    sTabellen(j) = "BEDZUGRI"
    j = j + 1
    sTabellen(j) = "ABSCHLUSS"
    j = j + 1
    sTabellen(j) = "DABAUSER"
    j = j + 1
    sTabellen(j) = "TABDATUM"
    j = j + 1
    sTabellen(j) = "KASSBON"
    j = j + 1
    sTabellen(j) = "KASSBOND"
    j = j + 1
    sTabellen(j) = "AFCSTAT"
    j = j + 1
    sTabellen(j) = "AFCSTATP"
    j = j + 1
    sTabellen(j) = "STORNOF"
    j = j + 1
    sTabellen(j) = "KOLLVERK"
    j = j + 1
    sTabellen(j) = "KREDIT"
    j = j + 1
    sTabellen(j) = "KASSJOUR"
    j = j + 1
    sTabellen(j) = "RETOURE"
    j = j + 1
    sTabellen(j) = "UEBERLI"
    j = j + 1
    sTabellen(j) = "KUNDKASS"
    j = j + 1
    sTabellen(j) = "BESTAEND"
    j = j + 1
    sTabellen(j) = "NOTIZEN"
    j = j + 1
    sTabellen(j) = "BESTPROT"
    j = j + 1
    sTabellen(j) = "EANPROT"
    j = j + 1
    sTabellen(j) = "KVKPR1PROT"
    j = j + 1
    sTabellen(j) = "MA" & srechnertab
    j = j + 1
    sTabellen(j) = "TABALI" & srechnertab
    j = j + 1
    sTabellen(j) = "ZADRESS"
    j = j + 1
    sTabellen(j) = "KULOESCH"
    j = j + 1
    sTabellen(j) = "BEAUFNR"
    j = j + 1
    sTabellen(j) = "UNTERWF"
    j = j + 1
    sTabellen(j) = "GANALYSE"
    j = j + 1
    sTabellen(j) = "DISPLAYTEXT"
    j = j + 1
    sTabellen(j) = "GANALYSEALL"
     j = j + 1
    sTabellen(j) = "KUMSUM"
    j = j + 1
    sTabellen(j) = "AFCBUCH"
    j = j + 1
    sTabellen(j) = "UMSATZINFO"
    j = j + 1
    sTabellen(j) = "ZZZ"

    Dim lAnzTable   As Long
    Dim lcount      As Long
    Dim sTabname    As String
    Dim sKISSTAB    As String
    Dim lMax        As Long
    Dim bBehalten   As Boolean
    
    gdBase.TableDefs.Refresh
    lAnzTable = gdBase.TableDefs.Count
    
    For lcount = 0 To lAnzTable - 1
        sTabname = UCase(gdBase.TableDefs(lcount).name)
        
        If UCase(Left(sTabname, 4)) = "MSYS" Then
        
        ElseIf UCase(Left(sTabname, 1)) = "Q" Then
        
        ElseIf UCase(Left(sTabname, 3)) = "INV" Then
        
        Else
            'ist der Tabellenname im Array
            bBehalten = False
            For i = 0 To UBound(sTabellen)
            
                sKISSTAB = UCase(sTabellen(i))
                
                If sKISSTAB = sTabname Then
                    bBehalten = True
                    Exit For
                End If
            Next i
            
            If bBehalten = False Then
                loeschNEW sTabname, gdBase
            End If
            
        End If
    Next lcount
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "db_Ballast_Tabellen_del"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub LeseDisplayText()
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    
    sSQL = "Select * from DISPLAYTEXT "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!MORGENTEXT) Then
            gsMORGENTEXT = rsrs!MORGENTEXT
        End If
        
        If Not IsNull(rsrs!MITTAGTEXT) Then
            gsMITTAGTEXT = rsrs!MITTAGTEXT
        End If
        
        If Not IsNull(rsrs!ABENDTEXT) Then
            gsABENDTEXT = rsrs!ABENDTEXT
        End If
    Else
        gsMORGENTEXT = "Guten Morgen!"
        gsMITTAGTEXT = "Guten Tag!"
        gsABENDTEXT = "Guten Abend!"
        
    End If
    rsrs.Close
                
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "LeseDisplayText"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub SpeicherDisplayText(sMorgen As String, sMittag As String, sAbend As String)
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    
    sSQL = "Delete from DISPLAYTEXT "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into DISPLAYTEXT "
    sSQL = sSQL & " (MORGENTEXT,MITTAGTEXT,ABENDTEXT) "
    sSQL = sSQL & " Values "
    sSQL = sSQL & " ( '" & sMorgen & "', '" & sMittag & "','" & sAbend & "')"
    gdBase.Execute sSQL, dbFailOnError
              
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "SpeicherDisplayText"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Function ermgesUmsatzausZumsatz(cVon As String, cBis As String) As Double
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    
    ermgesUmsatzausZumsatz = 0
    
    sSQL = "Select sum(UMSG1) as Maxi"
    sSQL = sSQL & " from UMSATZ "
    sSQL = sSQL & " where Datum between  " & cVon & " And " & cBis
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermgesUmsatzausZumsatz = rsrs!maxi
        End If
    End If
    rsrs.Close
                
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ermgesUmsatzausZumsatz"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermgesEKausZumsatz(cVon As String, cBis As String) As Double
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    
    ermgesEKausZumsatz = 0
    
    sSQL = "Select sum(EKPR1) as Maxi"
    sSQL = sSQL & " from UMSATZ "
    sSQL = sSQL & " where Datum between  " & cVon & " And " & cBis
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermgesEKausZumsatz = rsrs!maxi
        End If
    End If
    rsrs.Close
                
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ermgesEKausZumsatz"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermgesEKZugang(cVon As String, cBis As String, lLinr As Long) As String
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL    As String
    Dim rsrs    As Recordset
    Dim ctmp    As String
    
    loeschNEW "EUMS" & srechnertab, gdBase
    loeschNEW "EUMS4" & srechnertab, gdBase
    
    cSQL = "Select BEWEGUNG , EKPR  into EUMS4" & srechnertab
    cSQL = cSQL & " from ZUGANG "
    cSQL = cSQL & " where adate between  " & cVon & " And " & cBis
    If lLinr = 0 Then
    
    Else
        cSQL = cSQL & " and linr = " & lLinr
    End If
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Select  SUM(BEWEGUNG * EKPR) as EAKJahr into EUMS" & srechnertab
    cSQL = cSQL & " from EUMS4" & srechnertab
    gdBase.Execute cSQL, dbFailOnError
    
    ctmp = "0,00"
    Set rsrs = gdBase.OpenRecordset("EUMS" & srechnertab, dbOpenTable)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!EAKJahr) Then
            ctmp = rsrs!EAKJahr
        End If
    End If
    rsrs.Close
    
    loeschNEW "EUMS" & srechnertab, gdBase
    loeschNEW "EUMS4" & srechnertab, gdBase
    
    ermgesEKZugang = ctmp
    
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ermgesEKZugang"
    Fehler.gsFehlertext = "Beim Ermitteln des Einkaufsumsatzes ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function Bildspeichern(sArtnr As String, sBildPfad As String, slibesnr As String, FileX As FileListBox, bKill As Boolean) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim cQuelle     As String
    Dim cZiel       As String
    Dim lfail       As Long
    Dim lRet        As Long
    Dim lcount      As Long
    Dim sVerzBilder As String
    
    Dim sPfad       As String
    Dim i           As Integer
    Dim cEAN        As String
    Dim cSQL        As String
    Dim rsrs        As Recordset
    
    Bildspeichern = False
    
    If sArtnr = "" Then
        Exit Function
    End If

    cQuelle = sBildPfad
    cQuelle = ShortPath(cQuelle)
    cQuelle = cQuelle & "\" & slibesnr & ".jpg"

    cZiel = gcDBPfad
    If Right(cZiel, 1) <> "\" Then
        cZiel = cZiel & "\"
    End If
    cZiel = ShortPath(cZiel)
    
    cZiel = cZiel & "PICTURE\ARTIKEL"
    cZiel = cZiel & "\" & sArtnr & ".jpg"
    
    If bKill Then
        Kill cZiel
    End If

    lRet = CopyFile(cQuelle, cZiel, lfail)
    
    If lRet = 1 Then
        Bildspeichern = True
        Exit Function
    End If
    
    cEAN = ""
    cSQL = "Select EAN from ARTIKEL where ARTNR = " & sArtnr
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!EAN) Then
            cEAN = Trim(rsrs!EAN)
        End If
    End If
    rsrs.Close
    
    
    Dim sVerzBildNum As String
    Dim lCounter As Long
    
    
    lCounter = 1
    
    For lcount = 0 To FileX.ListCount - 1
        sVerzBilder = FileX.list(lcount)
        sVerzBilder = Left(sVerzBilder, Len(sVerzBilder) - (Len(sVerzBilder) - InStr(1, sVerzBilder, ".")))
        
        sVerzBildNum = ""
        For i = 1 To Len(sVerzBilder)
            If IsNumeric(Mid(sVerzBilder, i, 1)) = True Then
                sVerzBildNum = sVerzBildNum & Mid(sVerzBilder, i, 1)
            Else
                'endet auf jeden Fall mit einem Punkt also nicht numerisch
                sVerzBildNum = CStr(Val(sVerzBildNum))
                If sVerzBildNum = slibesnr Then
                
                    cQuelle = sBildPfad
                    cQuelle = ShortPath(cQuelle)
                    cQuelle = cQuelle & "\" & sVerzBilder & "jpg"
                
                    cZiel = gcDBPfad
                    If Right(cZiel, 1) <> "\" Then
                        cZiel = cZiel & "\"
                    End If
                    cZiel = ShortPath(cZiel)
                    
                    cZiel = cZiel & "PICTURE\ARTIKEL"
                    cZiel = cZiel & "\" & sArtnr & ".jpg"
                    
                    
                    
'                    'Gibt es schon ein Bild -> dann 1234_1.jpg
'                    If FileExists(cZiel & "\" & sArtnr & ".jpg") Then
'                        Do While FileExists(cZiel & "\" & sArtnr & "_" & lCounter & ".jpg")
'                            lCounter = lCounter + 1
'                        Loop
'                        cZiel = cZiel & "\" & sArtnr & "_" & lCounter & ".jpg"
'                    Else
'                        cZiel = cZiel & "\" & sArtnr & ".jpg"
'                    End If
                    
                    
                
                    lRet = CopyFile(cQuelle, cZiel, lfail)
                    
                    If lRet = 1 Then
                        Bildspeichern = True
                        Exit Function
                    Else
                        sVerzBildNum = ""
                    End If
                    
                ElseIf sVerzBildNum = cEAN Then
                
                    cQuelle = sBildPfad
                    cQuelle = ShortPath(cQuelle)
                    cQuelle = cQuelle & "\" & sVerzBilder & "jpg"
                
                    cZiel = gcDBPfad
                    If Right(cZiel, 1) <> "\" Then
                        cZiel = cZiel & "\"
                    End If
                    cZiel = ShortPath(cZiel)
                    
                    cZiel = cZiel & "PICTURE\ARTIKEL"
                    cZiel = cZiel & "\" & sArtnr & ".jpg"
                
                    lRet = CopyFile(cQuelle, cZiel, lfail)
                    
                    If lRet = 1 Then
                        Bildspeichern = True
                        Exit Function
                    Else
                        sVerzBildNum = ""
                    End If
                Else
                    sVerzBildNum = ""
                End If
                
            End If
        Next i
        
        
        
    Next lcount
    

    Exit Function
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "Bildspeichern"
        Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."
    
        Fehlermeldung1
    End If
    

End Function
Public Function Bildspeichern2(sArtnr As String, sBildPfad As String, slibesnr As String, sBez As String, FileX As FileListBox) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim cQuelle     As String
    Dim cZiel       As String
    Dim lfail       As Long
    Dim lRet        As Long
    Dim lcount      As Long
    Dim sVerzBilder As String
    
    Dim sPfad       As String
    Dim i           As Integer
    Dim cEAN        As String
    Dim cSQL        As String
    Dim rsrs        As Recordset
    
    Dim ctmp        As String
    Dim sBilddatei  As String
    
    Bildspeichern2 = False
    
    If sArtnr = "" Then
        Exit Function
    End If
    
    
    
    cQuelle = sBildPfad
    cQuelle = ShortPath(cQuelle)
    cQuelle = cQuelle & slibesnr & ".jpg"

    cZiel = gcDBPfad
    If Right(cZiel, 1) <> "\" Then
        cZiel = cZiel & "\"
    End If
    cZiel = ShortPath(cZiel)

    cZiel = cZiel & "PICTURE\ARTIKEL"
    
    If FileExists(cZiel & "\" & sArtnr & ".jpg") Then
        Exit Function
    End If
    
'    If bKill Then
'        Kill cZiel & "\" & sArtnr & "*.jpg"
'    End If
    
    cZiel = cZiel & "\" & sArtnr & ".jpg"

    lRet = CopyFile(cQuelle, cZiel, lfail)

    If lRet = 1 Then
        Bildspeichern2 = True
    End If

    cEAN = ""
    cSQL = "Select EAN from ARTIKEL where ARTNR = " & sArtnr
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!EAN) Then
            cEAN = Trim(rsrs!EAN)
        End If
    End If
    rsrs.Close

    Dim sVerzBildNum As String
    Dim lCounter As Long

    lCounter = 1

    For lcount = 0 To FileX.ListCount - 1
        sVerzBilder = FileX.list(lcount)
        sVerzBilder = Left(sVerzBilder, Len(sVerzBilder) - (Len(sVerzBilder) - InStr(1, sVerzBilder, ".")))

        sVerzBildNum = ""
        For i = 1 To Len(sVerzBilder)
            If IsNumeric(Mid(sVerzBilder, i, 1)) = True Then
                sVerzBildNum = sVerzBildNum & Mid(sVerzBilder, i, 1)
            Else
                'endet auf jeden Fall mit einem Punkt also nicht numerisch
                sVerzBildNum = CStr(Val(sVerzBildNum))
                If sVerzBildNum = slibesnr And Len(slibesnr) > 3 Then

                    cQuelle = sBildPfad
                    cQuelle = ShortPath(cQuelle)
                    cQuelle = cQuelle & sVerzBilder & "jpg"

                    cZiel = gcDBPfad
                    If Right(cZiel, 1) <> "\" Then
                        cZiel = cZiel & "\"
                    End If
                    cZiel = ShortPath(cZiel)

                    cZiel = cZiel & "PICTURE\ARTIKEL"


                    'Gibt es schon ein Bild -> dann 1234_1.jpg
                    If FileExists(cZiel & "\" & sArtnr & ".jpg") Then
                        Do While FileExists(cZiel & "\" & sArtnr & "_" & lCounter & ".jpg")
                            lCounter = lCounter + 1
                        Loop
                        
                        cZiel = cZiel & "\" & sArtnr & "_" & lCounter & ".jpg"
                        lCounter = lCounter + 1
                    Else
                        cZiel = cZiel & "\" & sArtnr & ".jpg"
                    End If



                    lRet = CopyFile(cQuelle, cZiel, lfail)

                    If lRet = 1 Then
                        Bildspeichern2 = True
                        sVerzBildNum = ""
                    Else
                        sVerzBildNum = ""
                    End If

                ElseIf sVerzBildNum = cEAN And sVerzBildNum <> "0" Then

                    cQuelle = sBildPfad
                    cQuelle = ShortPath(cQuelle)
                    cQuelle = cQuelle & sVerzBilder & "jpg"

                    cZiel = gcDBPfad
                    If Right(cZiel, 1) <> "\" Then
                        cZiel = cZiel & "\"
                    End If
                    cZiel = ShortPath(cZiel)

                    cZiel = cZiel & "PICTURE\ARTIKEL"

                    'Gibt es schon ein Bild -> dann 1234_1.jpg
                    If FileExists(cZiel & "\" & sArtnr & ".jpg") Then
                        Do While FileExists(cZiel & "\" & sArtnr & "_" & lCounter & ".jpg")
                            lCounter = lCounter + 1
                        Loop

                        cZiel = cZiel & "\" & sArtnr & "_" & lCounter & ".jpg"
                        lCounter = lCounter + 1
                    Else
                        cZiel = cZiel & "\" & sArtnr & ".jpg"
                    End If

                    lRet = CopyFile(cQuelle, cZiel, lfail)

                    If lRet = 1 Then
                        Bildspeichern2 = True
                        sVerzBildNum = ""
                    Else
                        sVerzBildNum = ""
                    End If
        
                
                Else
                    sVerzBildNum = ""
                End If

            End If
        Next i
    Next lcount
    
    
''    'wenn nichts
''
''    cZiel = gcDBPfad
''    If Right(cZiel, 1) <> "\" Then
''        cZiel = cZiel & "\"
''    End If
''    cZiel = ShortPath(cZiel)
''
''    cZiel = cZiel & "PICTURE\ARTIKEL"
''    If FileExists(cZiel & "\" & sArtnr & ".jpg") Then
''        Exit Function
''    End If
''
''    For lcount = 0 To FileX.ListCount - 1
''        sVerzBilder = FileX.list(lcount)
''        sVerzBilder = Left(sVerzBilder, Len(sVerzBilder) - (Len(sVerzBilder) - InStr(1, sVerzBilder, ".")))
''
''        sVerzBildNum = ""
''        For i = 1 To Len(sVerzBilder)
''            If IsNumeric(Mid(sVerzBilder, i, 1)) = True Then
''                sVerzBildNum = sVerzBildNum & Mid(sVerzBilder, i, 1)
''            Else
''                'endet auf jeden Fall mit einem Punkt also nicht numerisch
''                sVerzBildNum = CStr(Val(sVerzBildNum))
''
''                If sVerzBildNum = sBez And sVerzBildNum <> "0" And Len(sBez) > 3 Then
''
''                    cQuelle = sBildPfad
''                    cQuelle = ShortPath(cQuelle)
''                    cQuelle = cQuelle & "\" & sVerzBilder & "jpg"
''
''                    cZiel = gcDBPfad
''                    If Right(cZiel, 1) <> "\" Then
''                        cZiel = cZiel & "\"
''                    End If
''                    cZiel = ShortPath(cZiel)
''
''                    cZiel = cZiel & "PICTURE\ARTIKEL"
''
''                    'Gibt es schon ein Bild -> dann 1234_1.jpg
''                    If FileExists(cZiel & "\" & sArtnr & ".jpg") Then
''                        Do While FileExists(cZiel & "\" & sArtnr & "_" & lCounter & ".jpg")
''                            lCounter = lCounter + 1
''                        Loop
''
''                        cZiel = cZiel & "\" & sArtnr & "_" & lCounter & ".jpg"
''                        lCounter = lCounter + 1
''                    Else
''                        cZiel = cZiel & "\" & sArtnr & ".jpg"
''                    End If
''
''                    lRet = CopyFile(cQuelle, cZiel, lfail)
''
''                    If lRet = 1 Then
''                        Bildspeichern2 = True
''                        sVerzBildNum = ""
''                    Else
''                        sVerzBildNum = ""
''                    End If
''
''                Else
''                    sVerzBildNum = ""
''                End If
''
''            End If
''        Next i
''    Next lcount

    Exit Function
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "Bildspeichern2"
        Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."
    
        Fehlermeldung1
    End If
End Function
Public Sub IstdieLinrinUeberli(cLinr As String, sArtnr As String, cLiefbest As String, dLEK As Double, lMinMen As Long)
On Error GoTo LOKAL_ERROR

    
    Dim lOlinr      As Long
    Dim sSQL        As String
    Dim rsOli       As Recordset
    Dim rsRs3       As Recordset
    
    sSQL = "select * from ueberli where Linr = " & cLinr
    Set rsOli = gdBase.OpenRecordset(sSQL)
    
    If Not rsOli.EOF Then
        rsOli.MoveFirst
        Do While Not rsOli.EOF
            If Not IsNull(rsOli!oLINR) Then
                lOlinr = Val(rsOli!oLINR)
                
                sSQL = "Select * from artlief where artnr = " & sArtnr
                sSQL = sSQL & " and linr = " & lOlinr
                
                Set rsRs3 = gdBase.OpenRecordset(sSQL)
                
                If rsRs3.EOF Then
                    rsRs3.AddNew
                    rsRs3!SYNStatus = "A"
                Else
                    rsRs3.Edit
                    rsRs3!SYNStatus = "E"
                End If
                
                rsRs3!artnr = sArtnr
                rsRs3!linr = lOlinr
                rsRs3!LIBESNR = cLiefbest
                rsRs3!lekpr = dLEK
                rsRs3!MINMEN = lMinMen
                rsRs3.Update
                rsRs3.Close: Set rsRs3 = Nothing
            End If
            rsOli.MoveNext
        Loop
    End If
        
    rsOli.Close: Set rsOli = Nothing
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "IstdieLinrinUeberli"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Function ermSumAlterg() As Double
On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsGZ As Recordset
    
    ermSumAlterg = 0
    
    cSQL = "Select sum(GeldWERT) as maxi from ALTERG "
    cSQL = cSQL & " where sendok = false "
    Set rsGZ = gdBase.OpenRecordset(cSQL)
    If Not rsGZ.EOF Then
        If Not IsNull(rsGZ!maxi) Then
            ermSumAlterg = rsGZ!maxi
        End If
    End If
    rsGZ.Close: Set rsGZ = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ermSumAlterg"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Sub insertUNTERWF(lDat As Long, czeit As String, cArtNr As String, cMenge As String, cZielFil As String)
On Error GoTo LOKAL_ERROR

    Dim cSQL As String
    Dim rsGZ As Recordset

    cSQL = "Select * from UNTERWF where artnr = -1"
    FnOpenrecordset rsGZ, cSQL, 1, gdBase
    
    rsGZ.AddNew
    
    rsGZ!ADATE = lDat
    rsGZ!AZEIT = czeit
    rsGZ!artnr = cArtNr
    rsGZ!MENGE = cMenge
    rsGZ!VONFILIALE = gcFilNr
    rsGZ!ZIELFILIALE = cZielFil
    rsGZ!FILIALE = gcFilNr
    rsGZ!SENDOK = False
    
    rsGZ.Update
    rsGZ.Close: Set rsGZ = Nothing

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "insertUNTERWF"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub insertKundBest(lKUNDNR As Long, lartnr As Long, cMenge As String, lbednu As Long)
On Error GoTo LOKAL_ERROR

    Dim cSQL        As String
    Dim rsKB        As DAO.Recordset
    Dim rsArt       As DAO.Recordset
    
    Dim cBezeich    As String
    Dim cMwst       As String
    
    Dim cKVkPr1     As String
    Dim cEkPr       As String
    Dim cVKPR       As String
    
    
    Dim lKJADate    As Long
    Dim cKJAZeit    As String
    
    lKJADate = Fix(Now)
    cKJAZeit = Format$(Now, "HH:MM:SS")
    

    
    cSQL = "Select * from ARTIKEL where ARTNR = " & lartnr
    Set rsArt = gdBase.OpenRecordset(cSQL)
    If Not rsArt.EOF Then
        
        If Not IsNull(rsArt!BEZEICH) Then
            cBezeich = rsArt!BEZEICH
        End If
        
        If Not IsNull(rsArt!MWST) Then
            cMwst = rsArt!MWST
        End If
        
        If Not IsNull(rsArt!KVKPR1) Then
            cKVkPr1 = rsArt!KVKPR1
        End If
        
        If Not IsNull(rsArt!ekpr) Then
            cEkPr = rsArt!ekpr
        End If
        
        If Not IsNull(rsArt!vkpr) Then
            cVKPR = rsArt!vkpr
        End If
    End If
    rsArt.Close: Set rsArt = Nothing
    
    
    cSQL = "Select * from KUNDBEST where artnr = -1"
    FnOpenrecordset rsKB, cSQL, 1, gdBase
    
    
    rsKB.AddNew
    rsKB!artnr = lartnr
    rsKB!BEZEICH = cBezeich
    rsKB!BestelltMenge = Val(cMenge)
    rsKB!BestelltPreis = cKVkPr1
    rsKB!Bestelltam = lKJADate
    rsKB!Bestelltum = cKJAZeit
    rsKB!BEDNU = lbednu
    rsKB!Kundnr = lKUNDNR
    rsKB!FILIALE = Val(gcFilNr)
    rsKB!MWST = cMwst
    rsKB!ekpr = cEkPr
    rsKB!vkpr = cVKPR
    rsKB!StatusARTIKEL = "INBESTELLUNG"
    rsKB!statusKunde = "BESTELLT"
    rsKB!SENDOK = False
    rsKB.Update

    rsKB.Close: Set rsKB = Nothing

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "insertKundBest"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub KappeNetzVerbindung()
    On Error GoTo LOKAL_ERROR
    
    Dim lRet As Long
    Dim cLW As String
    
    cLW = gcNetzLW
    
    lRet = NetDisconnect(cLW, False, True)
                                
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "KappeNetzVerbindung"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Private Function NetDisconnect(ByVal LocalRoot As String, _
                                ByVal ReconnectAtLogin As Boolean, _
                                ByVal Force As Boolean) As Long
    On Error GoTo LOKAL_ERROR
    
    Dim Reconnect As Long
    
    If ReconnectAtLogin Then Reconnect = CONNECT_UPDATE_PROFILE
    
    NetDisconnect = WNetCancelConnection2(LocalRoot, Reconnect, Force)
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "NetDisconnect"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Function NetConnect(ByVal NetworkName As String, ByVal Username As String, _
                    ByVal Password As String, ByVal LocalRoot As String, _
                    ByVal ReconnectAtLogin As Boolean) As Long
    On Error GoTo LOKAL_ERROR
    
                    
    Dim NetRessource As NETRESOURCE

    With NetRessource
      .dwType = RESOURCETYPE_DISK
      .lpLocalName = LocalRoot
      .lpRemoteName = NetworkName
      'If ReconnectAtLogin Then .dwFlags = CONNECT_UPDATE_PROFILE
    End With
  
    NetConnect = WNetAddConnection2(NetRessource, Username, Password, 0)

    DoEvents

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "Netconnect"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Public Sub VerzVorhanden(sVerzeich As String, cPfad As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sVerz   As String
    Dim bFehler As Boolean
    Dim cLW     As String
       
    bFehler = False
    sVerz = ""
    sVerz = sVerzeich
    cLW = Left(cPfad, 2)
    ChDrive cLW
    ChDir "\"
    ChDir cPfad
    ChDir sVerz
    If bFehler Then
        MkDir sVerz
    Else
        ChDir cPfad
    End If
    
    Exit Sub
LOKAL_ERROR:
    If err.Number = 76 Then
        bFehler = True
        Resume Next
    ElseIf err.Number = 75 Then
        MsgBox "Es konnte das Verzeichnis " & sVerz & " nicht erstellt werden.", vbInformation, "Winkiss Programmstart: Hinweis"
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "VerzVorhanden"
        Fehler.gsFehlertext = "Es konnte das Verzeichnis " & sVerz & " nicht erstellt werden."
        
        Fehlermeldung1
    End If
End Sub
Public Function ermAUFNR(sTabname As String) As Long
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    
    ermAUFNR = 0

    sSQL = "Select * from BEAUFNR where TABNAME = '" & sTabname & "'"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!aufnr) Then
            ermAUFNR = rsrs!aufnr
        End If
    End If
    rsrs.Close

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ermAUFNR"
    Fehler.gsFehlertext = "Beim Ermitteln der Auftragnummer ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function schreibeAufnr(lAufNr As Long, sTabname As String)
    On Error GoTo LOKAL_ERROR
    
    Dim rec As Recordset
    Dim sSQL As String
    
    If sTabname <> "" Then
        sSQL = "Delete from BEAUFNR where tabname = '" & sTabname & "'"
        gdBase.Execute sSQL, dbFailOnError
    End If
    
    Set rec = gdBase.OpenRecordset("BEAUFNR")
    rec.AddNew
    rec!ADATE = DateValue(Now)
    rec!AZEIT = TimeValue(Now)
    rec!aufnr = lAufNr
    rec!tabname = sTabname
    rec.Update
    rec.Close: Set rec = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "schreibeAufnr"
    Fehler.gsFehlertext = "Im Programmteil Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermMaxAufnr() As Long
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rec As Recordset
    
    ermMaxAufnr = 100000
    
    cSQL = "select max(AUFNR) as aktlfnr from BEAUFNR "
    Set rec = gdBase.OpenRecordset(cSQL)
    If Not rec.EOF Then
        If Not IsNull(rec!aktlfnr) Then
            ermMaxAufnr = rec!aktlfnr
            ermMaxAufnr = ermMaxAufnr + 1
        End If
    End If
    rec.Close: Set rec = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ermMaxAufnr"
    Fehler.gsFehlertext = "Im Programmteil Bestellung ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermMaxRepnr() As Long
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rec As Recordset
    
    ermMaxRepnr = 100000
    
    cSQL = "select max(AUFtragNR) as aktlfnr from REPARATUR "
    Set rec = gdBase.OpenRecordset(cSQL)
    If Not rec.EOF Then
        If Not IsNull(rec!aktlfnr) Then
            ermMaxRepnr = rec!aktlfnr
            ermMaxRepnr = ermMaxRepnr + 1
        End If
    End If
    rec.Close: Set rec = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ermMaxRepnr"
    Fehler.gsFehlertext = "Im Programmteil Reparatur ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function

Public Sub PruefeSubDir()
    On Error GoTo LOKAL_ERROR
    
    Dim cPfad As String
    Dim bFehler As Boolean
    Dim ctmp As String
    
    Dim lRet As Long
    Dim cNetzLW As String
    Dim bitDrives As Long
    Dim i As Integer
    Dim iFileNr As Integer
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
   
    If Left(cPfad, 2) = "\\" Then
        bitDrives = GetLogicalDrives
        For i = 25 To 2 Step -1 ' Z - C
            If ((2 ^ i) And bitDrives) = 0 Then
                cNetzLW = Chr$(i + 65) & ":"   ' 0 entspricht A, "A" = Chr$(65)
                Exit For
            End If
        Next i

        lRet = NetConnect(cPfad, vbNullString, vbNullString, cNetzLW, False)
        If lRet = 0 Then
            gcNetzLW = cNetzLW
        Else
            gcNetzLW = ""
        End If
        
        If gcNetzLW = "" Then
            MsgBox "Kann keine Verbindung zur Datenbank aufbauen!", vbCritical, "STOP!"
            Unload frmWKL00
            End 'Ende
        Else
            gbNetzLW = True
            gcDBPfad = gcNetzLW
        End If
    End If
    
    VerzVorhanden "SALES", cPfad
    VerzVorhanden "RETOURE", cPfad
    VerzVorhanden "BESTAND", cPfad
    VerzVorhanden "VEDES", cPfad
    VerzVorhanden "VEDESDSL", cPfad
    VerzVorhanden "IP", cPfad
    VerzVorhanden "GFK", cPfad
    VerzVorhanden "BEAUTY", cPfad
    VerzVorhanden "BESTSIC", cPfad
    VerzVorhanden "ZEITUNG", cPfad
    VerzVorhanden "WKLEER", cPfad
    VerzVorhanden "TRANSOUT", cPfad
    VerzVorhanden "DTA", cPfad
    VerzVorhanden "DTAHEUTE", cPfad
    VerzVorhanden "Update", gcPfad
    VerzVorhanden "IN", cPfad
    VerzVorhanden "WV", cPfad
    VerzVorhanden "WVOUT", cPfad
    VerzVorhanden "WVOUTSIC", cPfad
    VerzVorhanden "DELART", cPfad
    VerzVorhanden "WVSIC", cPfad
    VerzVorhanden "Stammda", cPfad
    VerzVorhanden "MDEPROT", cPfad
    VerzVorhanden "DABASIC", cPfad
    VerzVorhanden "DABASIC1", cPfad
    VerzVorhanden "LIBRI", cPfad & "Stammda\"
    VerzVorhanden "ESUE", cPfad & "Stammda\"
    VerzVorhanden "RING", cPfad & "Stammda\"
    VerzVorhanden "BTE", cPfad & "Stammda\"
    VerzVorhanden "BOSS", cPfad & "Stammda\"
    VerzVorhanden "GERRY", cPfad & "Stammda\"
    VerzVorhanden "PASSPORT", cPfad & "Stammda\"
    VerzVorhanden "LUE", cPfad & "Stammda\"
    VerzVorhanden "BELA", cPfad & "Stammda\"
    VerzVorhanden "EDEKA", cPfad & "Stammda\"
    VerzVorhanden "TEXTIL", cPfad & "Stammda\"
    VerzVorhanden "TCHIBO", cPfad & "Stammda\"
    VerzVorhanden "NEUFORM", cPfad & "Stammda\"
    VerzVorhanden "BOLLWEG", cPfad & "Stammda\"
    VerzVorhanden "HOFFMANN", cPfad & "Stammda\"
    VerzVorhanden "VEDES", cPfad & "Stammda\"
    VerzVorhanden "IDEN", cPfad & "Stammda\"
    VerzVorhanden "KISS", cPfad & "Stammda\"
    VerzVorhanden "DEVIL", cPfad & "Stammda\"
    VerzVorhanden "NURDIE", cPfad & "Stammda\"
    VerzVorhanden "REWE", cPfad & "Stammda\"
    VerzVorhanden "ZEITUNG", cPfad & "Stammda\"
    VerzVorhanden "FRISEUR", cPfad & "Stammda\"
    VerzVorhanden "TUB", cPfad & "Stammda\"
    VerzVorhanden "SIE", cPfad & "Stammda\"
    VerzVorhanden "Bestell", cPfad
    VerzVorhanden "Out", cPfad
    VerzVorhanden "Update", cPfad
    VerzVorhanden "Box", cPfad
    VerzVorhanden "Stat", cPfad
    VerzVorhanden "Export", cPfad
    VerzVorhanden "Sicherung", cPfad
    VerzVorhanden "Picture", cPfad
    VerzVorhanden "GDPdU", cPfad
    VerzVorhanden "REPORTE", cPfad
    
    VerzVorhanden "EUR", cPfad & "Picture\"
    VerzVorhanden "SFR", cPfad & "Picture\"
    VerzVorhanden "DEM", cPfad & "Picture\"
    VerzVorhanden "ARTIKEL", cPfad & "Picture\"
    VerzVorhanden "KUNDEN", cPfad & "Picture\"
    VerzVorhanden "SYSTEM", cPfad & "Picture\"
    
    
    If checkpic = False Then
    
        
        Dim lfail       As Long
        
    
        For i = 0 To 14
            lRet = CopyFile(cPfad & i & "kl.jpg", cPfad & "Picture\EUR\" & i & "kl.jpg", lfail)
    
            lRet = CopyFile(cPfad & i & "k.jpg", cPfad & "Picture\EUR\" & i & "k.jpg", lfail)
            Kill cPfad & i & "kl.jpg"
            Kill cPfad & i & "k.jpg"
            Kill cPfad & "Picture\" & i & "kl.jpg"
            Kill cPfad & "Picture\" & i & "k.jpg"
        Next i
    End If

    VerzVorhanden "Endzipin", cPfad
    VerzVorhanden "Endzip", cPfad
    VerzVorhanden "Filiale", cPfad
    VerzVorhanden "Kassout", cPfad
    VerzVorhanden "Abschlus", cPfad
    VerzVorhanden "Kassin", cPfad
    VerzVorhanden "Komp", cPfad
    VerzVorhanden "Protok", cPfad
    VerzVorhanden "LProtok", cPfad
    VerzVorhanden "ZProtok", cPfad
    VerzVorhanden "BedPro", cPfad
    VerzVorhanden "ABPro", cPfad
    VerzVorhanden "ABProSIC", cPfad
    VerzVorhanden "Mail", cPfad
    VerzVorhanden "Mailout", cPfad
    VerzVorhanden "XML", cPfad
    
    VerzVorhanden "TSE", gcPfad
    VerzVorhanden "DISPLAY", gcPfad
    VerzVorhanden "Sounds", gcPfad
    VerzVorhanden "KissHelp", gcPfad
    VerzVorhanden "BIGERR", gcPfad
    VerzVorhanden "EDI", gcPfad
    
    VerzVorhanden "ERRSIC", cPfad
    
    VerzVorhanden "WHelp", gcPfad
    VerzVorhanden "WHelp", cPfad
    VerzVorhanden "LERR", gcPfad
    VerzVorhanden "Sysanaly", gcPfad
    
    gcDBPfad = Mid(cPfad, 1, Len(cPfad) - 1)
Exit Sub
LOKAL_ERROR:
    If err.Number = 76 Then
        bFehler = True
        Resume Next
    ElseIf err.Number = 75 Then
        MsgBox "Es konnte das Verzeichnis nicht erstellt werden.", vbInformation, "Winkiss Programmstart: Hinweis"
        Resume Next
    ElseIf err.Number = 53 Then
        
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "PruefeSubDir"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Public Sub ReIndiziereArtikelWKL00(db As Database)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim iFileNr As Integer
    Dim cTabelle As String
    Dim ctmp As String
    
    Screen.MousePointer = 11
    
    frmWKL00!Label2.Caption = "ReIndiziere Artikel-Datenbank"
    frmWKL00!Label2.Visible = True
    frmWKL00!Label2.Refresh
    
    frmWKL00.txtStatus.Text = "0"
    frmWKL00.picprogress.Visible = True
    
    
    
    '*******************
    '***** ARTIKEL *****
    '*******************
    cTabelle = "ARTIKEL"

    cSQL = "Drop Index ARTNR on ARTIKEL"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index EAN on ARTIKEL"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index EAN2 on ARTIKEL"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index EAN3 on ARTIKEL"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index BEZEICH on ARTIKEL"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index LINR on ARTIKEL"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index AGN on ARTIKEL"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index PGN on ARTIKEL"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index LIBESNR on ARTIKEL"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index LASTDATETIME on ARTIKEL"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index BESTAND on ARTIKEL"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index LASTDATE on ARTIKEL"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index AWM on ARTIKEL"
    db.Execute cSQL, dbFailOnError
    
'    cSQL = "Drop Index SYNSTATUS on ARTIKEL"
'    db.Execute cSQL, dbFailOnError
    

    frmWKL00.txtStatus.Text = "10"
    '*******************
    
    '*******************
    
    
    frmWKL00!Label2.Caption = "ReIndiziere Artikel-Datenbank: ARTIKELNUMMER"
    frmWKL00!Label2.Refresh
    
    frmWKL00.txtStatus.Text = "12"
    cSQL = "Create Index ARTNR on ARTIKEL (ARTNR)"
    db.Execute cSQL, dbFailOnError
    
    frmWKL00!Label2.Caption = "ReIndiziere Artikel-Datenbank: BESTAND"
    frmWKL00!Label2.Refresh
    frmWKL00.txtStatus.Text = "14"
    cSQL = "Create Index BESTAND on ARTIKEL(BESTAND)"
    db.Execute cSQL, dbFailOnError
    
    frmWKL00!Label2.Caption = "ReIndiziere Artikel-Datenbank: EAN-CODE"
    frmWKL00!Label2.Refresh
    frmWKL00.txtStatus.Text = "16"
    cSQL = "Create Index EAN on ARTIKEL (EAN)"
    db.Execute cSQL, dbFailOnError
    frmWKL00.txtStatus.Text = "18"
    cSQL = "Create Index EAN2 on ARTIKEL(EAN2)"
    db.Execute cSQL, dbFailOnError
    frmWKL00.txtStatus.Text = "20"
    cSQL = "Create Index EAN3 on ARTIKEL(EAN3)"
    db.Execute cSQL, dbFailOnError
    
    frmWKL00!Label2.Caption = "ReIndiziere Artikel-Datenbank: ARTIKEL-BEZEICHNUNG"
    frmWKL00!Label2.Refresh
    frmWKL00.txtStatus.Text = "22"
    cSQL = "Create Index BEZEICH on ARTIKEL (BEZEICH)"
    db.Execute cSQL, dbFailOnError
    
    frmWKL00!Label2.Caption = "ReIndiziere Artikel-Datenbank: LIEFERANTENNUMMER"
    frmWKL00!Label2.Refresh
    frmWKL00.txtStatus.Text = "24"
    cSQL = "Create Index LINR on ARTIKEL (LINR)"
    db.Execute cSQL, dbFailOnError
    
    frmWKL00!Label2.Caption = "ReIndiziere Artikel-Datenbank: ARTIKELGRUPPE"
    frmWKL00!Label2.Refresh
    frmWKL00.txtStatus.Text = "26"
    cSQL = "Create Index AGN on ARTIKEL (AGN)"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Create Index PGN on ARTIKEL (PGN)"
    db.Execute cSQL, dbFailOnError
    
    frmWKL00!Label2.Caption = "ReIndiziere Artikel-Datenbank: LIBESNR"
    frmWKL00!Label2.Refresh
    
    cSQL = "Create Index LIBESNR on ARTIKEL (LIBESNR)"
    db.Execute cSQL, dbFailOnError
    
    frmWKL00!Label2.Caption = "ReIndiziere Artikel-Datenbank: LastDate"
    frmWKL00!Label2.Refresh
    
    cSQL = "Create Index LastDate on ARTIKEL (LastDate)"
    db.Execute cSQL, dbFailOnError
    
    frmWKL00!Label2.Caption = "ReIndiziere Artikel-Datenbank: Farbmerkmal"
    frmWKL00!Label2.Refresh
    
    cSQL = "Create Index AWM on ARTIKEL (AWM)"
    db.Execute cSQL, dbFailOnError
    
'    frmWKL00!Label2.Caption = "ReIndiziere Artikel-Datenbank: SYNSTATUS"
'    frmWKL00!Label2.Refresh
'
'    cSQL = "Create Index SYNSTATUS on ARTIKEL (SYNSTATUS)"
'    db.Execute cSQL, dbFailOnError
    '********************
    '***** KASSJOUR *****
    '********************
    cTabelle = "KASSJOUR"
    cSQL = "Drop Index ARTNR on KASSJOUR"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index ADATE on KASSJOUR"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index KUNDNR on KASSJOUR"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index MOPREIS on KASSJOUR"
    db.Execute cSQL, dbFailOnError
    
    
    
    
    '******************
    '***** UMSATZ *****
    '******************
    cTabelle = "UMSATZ"
    
    cSQL = "Drop Index DATUM on UMSATZ"
    db.Execute cSQL, dbFailOnError
    
    
    

    cTabelle = "UMS_ART"
    cSQL = "Drop Index PRIMKEY on UMS_ART"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index DATUM on UMS_ART"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index ARTNR on UMS_ART"
    db.Execute cSQL, dbFailOnError
    
    cTabelle = "UMS_ARTF"
    cSQL = "Drop Index PRIMKEY on UMS_ARTF"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index Jahr on UMS_ARTF"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index Monat on UMS_ARTF"
    db.Execute cSQL, dbFailOnError
    
    cTabelle = "UMSARTJF"
    cSQL = "Drop Index PRIMKEY on UMSARTJF"
    db.Execute cSQL, dbFailOnError

    cTabelle = "UMS_LIEF"

    cSQL = "Drop Index PRIMKEY on UMS_LIEF"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index DATUM on UMS_LIEF"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index LINR on UMS_LIEF"
    db.Execute cSQL, dbFailOnError

    cTabelle = "UMSARTJ"

    cSQL = "Drop Index PRIMKEY on UMSARTJ"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index JAHR on UMSARTJ"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index ARTNR on UMSARTJ"
    db.Execute cSQL, dbFailOnError

    cTabelle = "UMSKDJ"

    cSQL = "Drop Index PRIMKEY on UMSKDJ"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index JAHR on UMSKDJ"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index KUNDNR on UMSKDJ"
    db.Execute cSQL, dbFailOnError

    
    cTabelle = "UMSLIEFJ"

    cSQL = "Drop Index PRIMKEY on UMSLIEFJ"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index JAHR on UMSLIEFJ"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index LINR on UMSLIEFJ"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index Liefnr on UMSLIEFJ"
    db.Execute cSQL, dbFailOnError

    

    
    '*******************
    '***** KASSBON *****
    '*******************

    cTabelle = "KASSBON"

    cSQL = "Drop Index DATUM on KASSBON"
    db.Execute cSQL, dbFailOnError
    
    '*******************
    '***** INTERART *****
    '*******************
    
    cTabelle = "INTERART"

    cSQL = "Drop Index ARTNR on INTERART"
    db.Execute cSQL, dbFailOnError
    
    
    '*******************
    '***** BEDNAME *****
    '*******************
    
    cTabelle = "BEDNAME"

    cSQL = "Drop Index BEDNU on BEDNAME"
    db.Execute cSQL, dbFailOnError
    
    
    
    '*******************
    '***** KOLLVERK *****
    '*******************

    cTabelle = "KOLLVERK"

    cSQL = "Drop Index ARTNR on KOLLVERK"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index ADATE on KOLLVERK"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index KUNDNR on KOLLVERK"
    db.Execute cSQL, dbFailOnError
    
    
    '*******************
    '***** RETOURE *****
    '*******************

    cTabelle = "RETOURE"

    cSQL = "Drop Index ARTNR on RETOURE"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index ADATE on RETOURE"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index KUNDNR on RETOURE"
    db.Execute cSQL, dbFailOnError
    
    '*******************
    '***** TAUSCH *****
    '*******************
    cTabelle = "TAUSCH"

    cSQL = "Drop Index ARTNR on TAUSCH"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index ADATE on TAUSCH"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index LINR on TAUSCH"
    db.Execute cSQL, dbFailOnError
    
    
    '******************
    '***** KUNDEN *****
    '******************
    
    cTabelle = "KUNDEN"
    
    cSQL = "Drop Index RECHNR on KUNDEN"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index KUNDNR on KUNDEN"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index KUERZEL on KUNDEN"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index NAME on KUNDEN"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index Titel on KUNDEN"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index PLZ on KUNDEN"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index STADT on KUNDEN"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index KUNDKART on KUNDEN"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index AENDER on KUNDEN"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index LASTDATE on KUNDEN"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index LASTDATETIME on KUNDEN"
    db.Execute cSQL, dbFailOnError
    
    '******************
    '***** KUNDKASS *****
    '******************
    
    cTabelle = "KUNDKASS"
    
    cSQL = "Drop Index adate on KUNDKASS"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index KUNDNR on KUNDKASS"
    db.Execute cSQL, dbFailOnError
    
    
    '*****************
    '***** LISRT *****
    '*****************
    
    cTabelle = "LISRT"
    
    cSQL = "Drop Index LINR on LISRT"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index KUERZEL on LISRT"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index LIEFBEZ on LISRT"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index PLZ on LISRT"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index STADT on LISRT"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index LASTDATETIME on LISRT"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index LASTDATE on LISRT"
    db.Execute cSQL, dbFailOnError
    
    
    '*******************
    '***** ARTLIEF *****
    '*******************
    
    cTabelle = "ARTLIEF"
    
    cSQL = "Drop Index ARTNR on ARTLIEF"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index LINR on ARTLIEF"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index ARTLINR on ARTLIEF"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index LIBESNR on ARTLIEF"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index SYNSTATUS on ARTLIEF"
    db.Execute cSQL, dbFailOnError
    
   
    '*******************
    '***** GUTSCH ******
    '*******************
    
    cTabelle = "GUTSCH"
    
    cSQL = "Drop Index DAT_EINL on GUTSCH"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index GUTSCHNR on GUTSCH"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index LASTTIME on GUTSCH"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Drop Index LASTDATE on GUTSCH"
    db.Execute cSQL, dbFailOnError
    

    '*******************
    '***** ZBESTAND *****
    '*******************
    cTabelle = "ZBESTAND"
    
    If Not tableSuchenDBKombi("ZBESTAND", 1) Then
     
    Else
        
        cSQL = "Drop Index PRIMKEY on ZBESTAND"
        db.Execute cSQL, dbFailOnError

        cSQL = "Drop Index ARTNR on ZBESTAND"
        db.Execute cSQL, dbFailOnError
        
        cSQL = "Drop Index LASTDATE on ZBESTAND"
        db.Execute cSQL, dbFailOnError

    End If
    
    '*************
    '** Zugang **
    '*************
    cTabelle = "ZUGANG"
    cSQL = "Drop Index ARTNR on ZUGANG"
    db.Execute cSQL, dbFailOnError
    

    '******************************************************************
    '***** Neue Indices erzeugen **************************************
    '******************************************************************
     '*************
    '** ZUGANG **
    '*************
    cTabelle = "ZUGANG"
    frmWKL00!Label2.Caption = "ReIndiziere ZUGANG-Datenbank: ARTIKELNUMMER"
    frmWKL00!Label2.Refresh

    cSQL = "Create Index ARTNR on ZUGANG(ARTNR) "
    db.Execute cSQL, dbFailOnError
               
    '********************
    '***** KASSJOUR *****
    '********************
    
    cTabelle = "KASSJOUR"
    
    frmWKL00!Label2.Caption = "ReIndiziere Kassenjournal-Datenbank: DATUM"
    frmWKL00!Label2.Refresh
    
    cSQL = "Create Index ADATE on KASSJOUR (ADATE)"
    db.Execute cSQL, dbFailOnError
    
    frmWKL00!Label2.Caption = "ReIndiziere Kassenjournal-Datenbank: ARTIKELNUMMER"
    frmWKL00!Label2.Refresh
    
    cSQL = "Create Index ARTNR on KASSJOUR (ARTNR)"
    db.Execute cSQL, dbFailOnError
    
    frmWKL00!Label2.Caption = "ReIndiziere Kassenjournal-Datenbank: KUNDNR"
    frmWKL00!Label2.Refresh
    
    cSQL = "Create Index KUNDNR on KASSJOUR (KUNDNR)"
    db.Execute cSQL, dbFailOnError
    
    frmWKL00!Label2.Caption = "ReIndiziere Kassenjournal-Datenbank: MOPREIS"
    frmWKL00!Label2.Refresh
    
    cSQL = "Create Index MOPREIS on KASSJOUR (MOPREIS)"
    db.Execute cSQL, dbFailOnError
    
    '******************
    '***** UMSATZ *****
    '******************
    
    cTabelle = "UMSATZ"
    frmWKL00!Label2.Caption = "ReIndiziere Umsatz-Datenbank: DATUM"
    frmWKL00!Label2.Refresh
    
    cSQL = "Create Index DATUM on UMSATZ (DATUM)"
    db.Execute cSQL, dbFailOnError
    
    cTabelle = "UMS_ART"
    frmWKL00!Label2.Caption = "ReIndiziere Verkaufstabellen: UMS_ART"
    frmWKL00!Label2.Refresh
    
    cSQL = "Create Index PRIMKEY on UMS_ART (ARTNR, JAHR, MONAT)"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Create Index DATUM on UMS_ART (JAHR, MONAT)"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Create Index ARTNR on UMS_ART (ARTNR)"
    db.Execute cSQL, dbFailOnError
    
    
    cTabelle = "UMS_ARTF"
    frmWKL00!Label2.Caption = "ReIndiziere Verkaufstabellen: UMS_ARTF, ARTNR, FILIALNR, JAHR, MONAT"
    frmWKL00!Label2.Refresh
    
    cSQL = "Create Index PRIMKEY on UMS_ARTF (ARTNR, FILIALNR, JAHR, MONAT)"
    db.Execute cSQL, dbFailOnError
    
    frmWKL00!Label2.Caption = "ReIndiziere Verkaufstabellen: UMS_ARTF, Monat"
    frmWKL00!Label2.Refresh
    
    cSQL = "Create Index Monat on UMS_ARTF (Monat)"
    gdBase.Execute cSQL, dbFailOnError
    
    frmWKL00!Label2.Caption = "ReIndiziere Verkaufstabellen: UMS_ARTF, Jahr"
    frmWKL00!Label2.Refresh
    
    cSQL = "Create Index Jahr on UMS_ARTF (Jahr)"
    gdBase.Execute cSQL, dbFailOnError
        
    cTabelle = "UMSARTJF"
    frmWKL00!Label2.Caption = "ReIndiziere Verkaufstabellen: UMSARTJF"
    frmWKL00!Label2.Refresh
    
    cSQL = "Create Index PRIMKEY on UMSARTJF (ARTNR, FILIALNR, JAHR)"
    db.Execute cSQL, dbFailOnError
    
    cTabelle = "UMS_LIEF"
    frmWKL00!Label2.Caption = "ReIndiziere Verkaufstabellen: UMS_LIEF"
    frmWKL00!Label2.Refresh
    
    cSQL = "Create Index PRIMKEY on UMS_LIEF (LINR, JAHR, MONAT)"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Create Index DATUM on UMS_LIEF (JAHR, MONAT)"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Create Index LINR on UMS_LIEF (LINR)"
    db.Execute cSQL, dbFailOnError
    
    cTabelle = "UMSARTJ"
    frmWKL00!Label2.Caption = "ReIndiziere Verkaufstabellen: UMSARTJ"
    frmWKL00!Label2.Refresh
    
    cSQL = "Create Index PRIMKEY on UMSARTJ (ARTNR, JAHR)"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Create Index JAHR on UMSARTJ (JAHR)"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Create Index ARTNR on UMSARTJ (ARTNR)"
    db.Execute cSQL, dbFailOnError
    
    cTabelle = "UMSKDJ"
    frmWKL00!Label2.Caption = "ReIndiziere Verkaufstabellen: UMSKDJ"
    frmWKL00!Label2.Refresh
    
    cSQL = "Create Index PRIMKEY on UMSKDJ (KUNDNR, JAHR)"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Create Index JAHR on UMSKDJ (JAHR)"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Create Index KUNDNR on UMSKDJ (KUNDNR)"
    db.Execute cSQL, dbFailOnError
    
    cTabelle = "UMSLIEFJ"
    frmWKL00!Label2.Caption = "ReIndiziere Verkaufstabellen: UMSLIEFJ"
    frmWKL00!Label2.Refresh
    
    cSQL = "Create Index PRIMKEY on UMSLIEFJ (LINR, JAHR)"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Create Index JAHR on UMSLIEFJ (JAHR)"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Create Index LIEFNR on UMSLIEFJ (LINR)"
    db.Execute cSQL, dbFailOnError
    
    frmWKL00.txtStatus.Text = "30"
    '******************
    '***** KUNDEN *****
    '******************
    cTabelle = "KUNDEN"
    
    frmWKL00!Label2.Caption = "ReIndiziere Kunden-Datenbank: KUNDNR"
    frmWKL00!Label2.Refresh
    
    cSQL = "Create Index KUNDNR on KUNDEN (KUNDNR)"
    db.Execute cSQL, dbFailOnError
    
    frmWKL00!Label2.Caption = "ReIndiziere Kunden-Datenbank: KUERZEL"
    frmWKL00!Label2.Refresh
    
    cSQL = "Create Index KUERZEL on KUNDEN (KUERZEL)"
    db.Execute cSQL, dbFailOnError
    
    frmWKL00!Label2.Caption = "ReIndiziere Kunden-Datenbank: NAME"
    frmWKL00!Label2.Refresh
    
    cSQL = "Create Index NAME on KUNDEN (NAME)"
    db.Execute cSQL, dbFailOnError
    
    frmWKL00!Label2.Caption = "ReIndiziere Kunden-Datenbank: Titel"
    frmWKL00!Label2.Refresh
    
    cSQL = "Create Index Titel on KUNDEN (Titel)"
    db.Execute cSQL, dbFailOnError
    
    
    frmWKL00!Label2.Caption = "ReIndiziere Kunden-Datenbank: PLZ"
    frmWKL00!Label2.Refresh
    
    cSQL = "Create Index PLZ on KUNDEN (PLZ)"
    db.Execute cSQL, dbFailOnError
    
    frmWKL00!Label2.Caption = "ReIndiziere Kunden-Datenbank: STADT"
    frmWKL00!Label2.Refresh
    
    cSQL = "Create Index STADT on KUNDEN (STADT)"
    db.Execute cSQL, dbFailOnError
    
    frmWKL00!Label2.Caption = "ReIndiziere Kunden-Datenbank: KUNDKART"
    frmWKL00!Label2.Refresh
    
    cSQL = "Create Index KUNDKART on KUNDEN (KUNDKART)"
    db.Execute cSQL, dbFailOnError
    
    frmWKL00!Label2.Caption = "ReIndiziere Kunden-Datenbank: AENDER"
    frmWKL00!Label2.Refresh
    
    cSQL = "Create Index AENDER on KUNDEN (AENDER)"
    db.Execute cSQL, dbFailOnError
    
    frmWKL00!Label2.Caption = "ReIndiziere Kunden-Datenbank: RECHNR"
    frmWKL00!Label2.Refresh
    
    cSQL = "Create Index RECHNR on KUNDEN (RECHNR)"
    db.Execute cSQL, dbFailOnError
    
    '//LastDate und LastTime
    frmWKL00!Label2.Caption = "ReIndiziere Kunden-Datenbank: LastDate"
    frmWKL00!Label2.Refresh
    
    cSQL = "Create Index LastDate on KUNDEN (LastDate)"
    db.Execute cSQL, dbFailOnError
    
    frmWKL00.txtStatus.Text = "40"
    
    '******************
    '***** KUNDKASS *****
    '******************
    cTabelle = "KUNDKASS"
    
    frmWKL00!Label2.Caption = "ReIndiziere KUNDKASS: KUNDNR"
    frmWKL00!Label2.Refresh
    
    cSQL = "Create Index KUNDNR on KUNDKASS (KUNDNR)"
    db.Execute cSQL, dbFailOnError
    
    frmWKL00.txtStatus.Text = "41"
    
    frmWKL00!Label2.Caption = "ReIndiziere KUNDKASS: ADATE"
    frmWKL00!Label2.Refresh
    
    cSQL = "Create Index ADATE on KUNDKASS (ADATE)"
    db.Execute cSQL, dbFailOnError
    
    frmWKL00.txtStatus.Text = "4"
    
    '*****************
    '***** GUTSCH ****
    '*****************
    
    cTabelle = "GUTSCH"
    
    frmWKL00!Label2.Caption = "ReIndiziere Gutschein-Datenbank: DAT_EINL"
    cSQL = "Create Index DAT_EINL on GUTSCH (DAT_EINL)"
    db.Execute cSQL, dbFailOnError
    
    frmWKL00!Label2.Caption = "ReIndiziere Gutschein-Datenbank: GUTSCHNR"
    cSQL = "Create Index GUTSCHNR on GUTSCH (GUTSCHNR)"
    db.Execute cSQL, dbFailOnError

    frmWKL00!Label2.Caption = "ReIndiziere Gutschein-Datenbank: LastDate"
    cSQL = "Create Index LastDate on GUTSCH (LastDate)"
    db.Execute cSQL, dbFailOnError
    
    frmWKL00!Label2.Caption = "ReIndiziere Gutschein-Datenbank: LastTime"
    cSQL = "Create Index LASTTIME on GUTSCH (Lasttime)"
    db.Execute cSQL, dbFailOnError
      
    '*****************
    '***** KASSBON ****
    '*****************
    
    cTabelle = "KASSBON"
    
    frmWKL00!Label2.Caption = "ReIndiziere Kassenbon-Datenbank: DATUM"
    frmWKL00!Label2.Refresh
    
    cSQL = "Create Index DATUM on KASSBON (DATUM)"
    db.Execute cSQL, dbFailOnError
    
    '*****************
    '***** Interart ****
    '*****************
    
    cTabelle = "INTERART"
    
    frmWKL00!Label2.Caption = "ReIndiziere INTERART-Datenbank: ARTNR"
    frmWKL00!Label2.Refresh
    
    cSQL = "Create Index ARTNR on INTERART (ARTNR)"
    db.Execute cSQL, dbFailOnError
    
    frmWKL00.txtStatus.Text = "5"
    
    '*****************
    '***** BEDNAME ****
    '*****************
    
    cTabelle = "BEDNAME"
    
    frmWKL00!Label2.Caption = "ReIndiziere BEDNAME-Datenbank: BEDNU"
    frmWKL00!Label2.Refresh
    
    cSQL = "Create Index BEDNU on BEDNAME (BEDNU)"
    db.Execute cSQL, dbFailOnError
    
    frmWKL00.txtStatus.Text = "50"
    
    '*****************
    '***** KOLLVERK ****
    '*****************
    cTabelle = "KOLLVERK"
    
    frmWKL00!Label2.Caption = "ReIndiziere Kollegenverkauf-Datenbank: ARTNR"
    frmWKL00!Label2.Refresh
    
    cSQL = "Create Index ARTNR on KOLLVERK (ARTNR)"
    db.Execute cSQL, dbFailOnError
    
    frmWKL00!Label2.Caption = "ReIndiziere Kollegenverkauf-Datenbank: ADATE"
    frmWKL00!Label2.Refresh
    
    cSQL = "Create Index ADATE on KOLLVERK (ADATE)"
    db.Execute cSQL, dbFailOnError
    
    frmWKL00!Label2.Caption = "ReIndiziere Kollegenverkauf-Datenbank: KUNDNR"
    frmWKL00!Label2.Refresh
    
    cSQL = "Create Index KUNDNR on KOLLVERK (KUNDNR)"
    db.Execute cSQL, dbFailOnError
    
    
    '*****************
    '***** RETOURE ****
    '*****************
    cTabelle = "RETOURE"
    
    frmWKL00!Label2.Caption = "ReIndiziere Retoure-Datenbank: ARTNR"
    frmWKL00!Label2.Refresh
    
    cSQL = "Create Index ARTNR on RETOURE (ARTNR)"
    db.Execute cSQL, dbFailOnError
    
    frmWKL00!Label2.Caption = "ReIndiziere Retoure-Datenbank: ADATE"
    frmWKL00!Label2.Refresh
    
    cSQL = "Create Index ADATE on RETOURE (ADATE)"
    db.Execute cSQL, dbFailOnError
    
    frmWKL00!Label2.Caption = "ReIndiziere Retoure-Datenbank: KUNDNR"
    frmWKL00!Label2.Refresh
    
    cSQL = "Create Index KUNDNR on RETOURE (KUNDNR)"
    db.Execute cSQL, dbFailOnError
    
    
    '*****************
    '***** TAUSCH ****
    '*****************
    cTabelle = "TAUSCH"
    
    frmWKL00!Label2.Caption = "ReIndiziere Tausch-Datenbank: ADATE"
    frmWKL00!Label2.Refresh
    
    cSQL = "Create Index ADATE on TAUSCH (ADATE)"
    db.Execute cSQL, dbFailOnError
    
    frmWKL00!Label2.Caption = "ReIndiziere Tausch-Datenbank: ARTNR"
    frmWKL00!Label2.Refresh
    
    cSQL = "Create Index ARTNR on TAUSCH (ARTNR)"
    db.Execute cSQL, dbFailOnError
    
    frmWKL00!Label2.Caption = "ReIndiziere Tausch-Datenbank: LINR"
    frmWKL00!Label2.Refresh
    
    cSQL = "Create Index LINR on TAUSCH (LINR)"
    db.Execute cSQL, dbFailOnError
    
    '*****************
    '***** LISRT *****
    '*****************
    
    frmWKL00.txtStatus.Text = "60"

    cTabelle = "LISRT"
    
    frmWKL00!Label2.Caption = "ReIndiziere Lieferanten-Datenbank: LINR"
    frmWKL00!Label2.Refresh
    
    cSQL = "Create Index LINR on LISRT (LINR)"
    db.Execute cSQL, dbFailOnError
    
    frmWKL00!Label2.Caption = "ReIndiziere Lieferanten-Datenbank: KUERZEL"
    frmWKL00!Label2.Refresh
    
    cSQL = "Create Index KUERZEL on LISRT (KUERZEL)"
    db.Execute cSQL, dbFailOnError
    
    frmWKL00!Label2.Caption = "ReIndiziere Lieferanten-Datenbank: LIEFBEZ"
    frmWKL00!Label2.Refresh
    
    cSQL = "Create Index LIEFBEZ on LISRT (LIEFBEZ)"
    db.Execute cSQL, dbFailOnError
    
    frmWKL00!Label2.Caption = "ReIndiziere Lieferanten-Datenbank: PLZ"
    frmWKL00!Label2.Refresh
    
    cSQL = "Create Index PLZ on LISRT (PLZ)"
    db.Execute cSQL, dbFailOnError
    
    frmWKL00!Label2.Caption = "ReIndiziere Lieferanten-Datenbank: STADT"
    frmWKL00!Label2.Refresh
    
    cSQL = "Create Index STADT on LISRT (STADT)"
    db.Execute cSQL, dbFailOnError
    
    '//LastDate und LastTime
    frmWKL00!Label2.Caption = "ReIndiziere Lieferanten-Datenbank: LastDate"
    frmWKL00!Label2.Refresh
    
    cSQL = "Create Index LastDate on LISRT (LastDate)"
    db.Execute cSQL, dbFailOnError
    
    '*******************
    '***** ARTLIEF *****
    '*******************
    
    frmWKL00.txtStatus.Text = "70"

    cTabelle = "ARTLIEF"
    
    frmWKL00!Label2.Caption = "ReIndiziere Artikel-Lieferanten-Datenbank: ARTNR"
    frmWKL00!Label2.Refresh
    
    cSQL = "Create Index ARTNR on ARTLIEF (ARTNR)"
    db.Execute cSQL, dbFailOnError
    
    frmWKL00!Label2.Caption = "ReIndiziere Artikel-Lieferanten-Datenbank: LINR"
    frmWKL00!Label2.Refresh
    
    cSQL = "Create Index LINR on ARTLIEF (LINR)"
    db.Execute cSQL, dbFailOnError
    
    frmWKL00!Label2.Caption = "ReIndiziere Artikel-Lieferanten-Datenbank: ARTLINR"
    frmWKL00!Label2.Refresh
    
    cSQL = "Create Index ARTLINR on ARTLIEF (ARTNR, LINR)"
    db.Execute cSQL, dbFailOnError
    
    frmWKL00!Label2.Caption = "ReIndiziere Artikel-Lieferanten-Datenbank: LIEFERANTENBESTELLNUMMER"
    frmWKL00!Label2.Refresh
    
    cSQL = "Create Index LIBESNR on ARTLIEF (LIBESNR)"
    db.Execute cSQL, dbFailOnError
    
    cSQL = "Create Index SYNSTATUS on ARTLIEF (SYNSTATUS)"
    db.Execute cSQL, dbFailOnError
    
    
    '*******************
    '***** ZBESTAND *****
    '*******************
    cTabelle = "ZBESTAND"


    If Not tableSuchenDBKombi("ZBESTAND", 1) Then
     
    Else
        cSQL = "Create Index PRIMKEY on ZBESTAND (FILIALNR, ARTNR)"
        db.Execute cSQL, dbFailOnError

        cSQL = "Create Index ARTNR on ZBESTAND (ARTNR)"
        db.Execute cSQL, dbFailOnError
        
         cSQL = "Create Index LastDate on ZBESTAND (LastDate)"
        db.Execute cSQL, dbFailOnError
    End If

    
    '****************************************
    '***** Abgleich ARTIKEL und ARTLIEF *****
    '****************************************
    cTabelle = "ABGLEICH"

    frmWKL00!Label2.Caption = "Gleiche ARTIKEL und ARTLIEF gegeneinander ab (Stufe 1)"
    frmWKL00!Label2.Visible = True
    frmWKL00!Label2.Refresh




    loeschNEW "TEMP1", db
    loeschNEW "TEMP2", db

    BeginTrans
    cSQL = "Select ARTNR, LINR , 'N' as Erkannt into TEMP1 from ARTIKEL"
    db.Execute cSQL, dbFailOnError

    frmWKL00!Label2.Caption = "Gleiche ARTIKEL und ARTLIEF gegeneinander ab (Stufe 2)"
    frmWKL00!Label2.Visible = True
    frmWKL00!Label2.Refresh

    cSQL = "Update TEMP1 inner join Artlief on TEMP1.artnr = Artlief.artnr and TEMP1.LINR = Artlief.LINR "
    cSQL = cSQL & " Set TEMP1.Erkannt = 'J' "
    db.Execute cSQL, dbFailOnError

    cSQL = "Delete from TEMP1 where erkannt = 'J' "
    db.Execute cSQL, dbFailOnError

    frmWKL00!Label2.Caption = "Gleiche ARTIKEL und ARTLIEF gegeneinander ab (Stufe 3)"
    frmWKL00!Label2.Visible = True
    frmWKL00!Label2.Refresh

    cSQL = "Select ARTIKEL.ARTNR, ARTIKEL.LINR, ARTIKEL.LEKPR, ARTIKEL.LIBESNR, ARTIKEL.MINMEN "
    cSQL = cSQL & "into TEMP2 from ARTIKEL inner join TEMP1 on ARTIKEL.ARTNR = TEMP1.ARTNR "
    db.Execute cSQL, dbFailOnError
    
'    cSQL = "Update TEMP2 inner join Artlief on TEMP2.artnr = Artlief.artnr and TEMP2.LINR = Artlief.LINR "
'    cSQL = cSQL & " Set TEMP2.RKZ = Artlief.RKZ,TEMP2.EXDAT = Artlief.EXDAT "
'    db.Execute cSQL, dbFailOnError

    frmWKL00!Label2.Caption = "Gleiche ARTIKEL und ARTLIEF gegeneinander ab (Stufe 4)"
    frmWKL00!Label2.Visible = True
    frmWKL00!Label2.Refresh


    cSQL = "Insert into ARTLIEF Select * from TEMP2"
    db.Execute cSQL, dbFailOnError

    frmWKL00!Label2.Caption = "Gleiche ARTIKEL und ARTLIEF gegeneinander ab (Stufe 5)"
    frmWKL00!Label2.Visible = True
    frmWKL00!Label2.Refresh

    CommitTrans

    loeschNEW "TEMP1", db
    loeschNEW "TEMP2", db
    
    
    frmWKL00!Label2.Caption = "Anwender aktiv"
    frmWKL00!Label2.Refresh
    
    frmWKL00.txtStatus.Text = "0"
    frmWKL00.picprogress.Visible = False
    
    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:

    If err.Number = 3372 Or err.Number = 3371 Or err.Number = 3375 Or err.Number = 3256 Or err.Number = 3376 Or err.Number = 3043 Then
        Resume Next
    Else
    
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "ReIndiziereArtikelWKL00"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1
        Screen.MousePointer = 0
'        Resume Next
    End If
End Sub
Public Sub ReIndiziereArtikel_ForSS(db As Database)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    Dim rsrs As Recordset
    Dim iFileNr As Integer
    Dim cTabelle As String
    Dim ctmp As String
    
    Screen.MousePointer = 11
    
    '*******************
    '***** ARTIKEL *****
    '*******************
    cTabelle = "ARTIKEL"
    
    cSQL = "Drop Index ARTNR on KUNDBEST"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index StatusARTIKEL on KUNDBEST"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index Sendok on KUNDBEST"
    SQL_Befehl_ausführen cSQL
    
    

    cSQL = "Drop Index ARTNR on ARTIKEL"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index EAN on ARTIKEL"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index EAN2 on ARTIKEL"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index EAN3 on ARTIKEL"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index BEZEICH on ARTIKEL"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index LINR on ARTIKEL"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index AGN on ARTIKEL"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index PGN on ARTIKEL"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index LIBESNR on ARTIKEL"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index LASTDATETIME on ARTIKEL"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index BESTAND on ARTIKEL"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index LASTDATE on ARTIKEL"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index AWM on ARTIKEL"
    SQL_Befehl_ausführen cSQL
    
'    cSQL = "Drop Index SYNSTATUS on ARTIKEL"
'    SQL_Befehl_ausführen cSQL
    

    
    '*******************
    
    '*******************
    
    
    
    cSQL = "Create Index ARTNR on KUNDBEST (ARTNR)"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Create Index StatusARTIKEL on KUNDBEST (StatusARTIKEL)"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Create Index Sendok on KUNDBEST (Sendok)"
    SQL_Befehl_ausführen cSQL
    
    
    
    cSQL = "Create Index ARTNR on ARTIKEL (ARTNR)"
    SQL_Befehl_ausführen cSQL
    
    
    cSQL = "Create Index BESTAND on ARTIKEL(BESTAND)"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Create Index EAN on ARTIKEL (EAN)"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Create Index EAN2 on ARTIKEL(EAN2)"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Create Index EAN3 on ARTIKEL(EAN3)"
    SQL_Befehl_ausführen cSQL
   
    cSQL = "Create Index BEZEICH on ARTIKEL (BEZEICH)"
    SQL_Befehl_ausführen cSQL
    
    
    cSQL = "Create Index LINR on ARTIKEL (LINR)"
    SQL_Befehl_ausführen cSQL
    
   
    cSQL = "Create Index AGN on ARTIKEL (AGN)"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Create Index PGN on ARTIKEL (PGN)"
    SQL_Befehl_ausführen cSQL
    
    
    cSQL = "Create Index LIBESNR on ARTIKEL (LIBESNR)"
    SQL_Befehl_ausführen cSQL
    
    
    
    cSQL = "Create Index LastDate on ARTIKEL (LastDate)"
    SQL_Befehl_ausführen cSQL
    
    
    
    cSQL = "Create Index AWM on ARTIKEL (AWM)"
    SQL_Befehl_ausführen cSQL

    '********************
    '***** KASSJOUR *****
    '********************
    
    cSQL = "Drop Index ARTNR on KASSJOUR"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index ADATE on KASSJOUR"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index KUNDNR on KASSJOUR"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index MOPREIS on KASSJOUR"
    SQL_Befehl_ausführen cSQL
    
    
    
    
    '******************
    '***** UMSATZ *****
    '******************
    cTabelle = "UMSATZ"
    
    cSQL = "Drop Index DATUM on UMSATZ"
    SQL_Befehl_ausführen cSQL
    
    
    

    cTabelle = "UMS_ART"
    cSQL = "Drop Index PRIMKEY on UMS_ART"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index DATUM on UMS_ART"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index ARTNR on UMS_ART"
    SQL_Befehl_ausführen cSQL
    
    cTabelle = "UMS_ARTF"
    cSQL = "Drop Index PRIMKEY on UMS_ARTF"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index Jahr on UMS_ARTF"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index Monat on UMS_ARTF"
    SQL_Befehl_ausführen cSQL
    
    cTabelle = "UMSARTJF"
    cSQL = "Drop Index PRIMKEY on UMSARTJF"
    SQL_Befehl_ausführen cSQL

    cTabelle = "UMS_LIEF"

    cSQL = "Drop Index PRIMKEY on UMS_LIEF"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index DATUM on UMS_LIEF"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index LINR on UMS_LIEF"
    SQL_Befehl_ausführen cSQL

    cTabelle = "UMSARTJ"

    cSQL = "Drop Index PRIMKEY on UMSARTJ"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index JAHR on UMSARTJ"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index ARTNR on UMSARTJ"
    SQL_Befehl_ausführen cSQL

    cTabelle = "UMSKDJ"

    cSQL = "Drop Index PRIMKEY on UMSKDJ"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index JAHR on UMSKDJ"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index KUNDNR on UMSKDJ"
    SQL_Befehl_ausführen cSQL

    
    cTabelle = "UMSLIEFJ"

    cSQL = "Drop Index PRIMKEY on UMSLIEFJ"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index JAHR on UMSLIEFJ"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index LINR on UMSLIEFJ"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index Liefnr on UMSLIEFJ"
    SQL_Befehl_ausführen cSQL

    

    
    '*******************
    '***** KASSBON *****
    '*******************

    cTabelle = "KASSBON"

    cSQL = "Drop Index DATUM on KASSBON"
    SQL_Befehl_ausführen cSQL
    
    '*******************
    '***** INTERART *****
    '*******************
    
    cTabelle = "INTERART"

    cSQL = "Drop Index ARTNR on INTERART"
    SQL_Befehl_ausführen cSQL
    
    
    '*******************
    '***** BEDNAME *****
    '*******************
    
    cTabelle = "BEDNAME"

    cSQL = "Drop Index BEDNU on BEDNAME"
    SQL_Befehl_ausführen cSQL
    
    
    
    '*******************
    '***** KOLLVERK *****
    '*******************

    cTabelle = "KOLLVERK"

    cSQL = "Drop Index ARTNR on KOLLVERK"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index ADATE on KOLLVERK"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index KUNDNR on KOLLVERK"
    SQL_Befehl_ausführen cSQL
    
    
    '*******************
    '***** RETOURE *****
    '*******************

    cTabelle = "RETOURE"

    cSQL = "Drop Index ARTNR on RETOURE"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index ADATE on RETOURE"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index KUNDNR on RETOURE"
    SQL_Befehl_ausführen cSQL
    
    '*******************
    '***** TAUSCH *****
    '*******************
    cTabelle = "TAUSCH"

    cSQL = "Drop Index ARTNR on TAUSCH"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index ADATE on TAUSCH"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index LINR on TAUSCH"
    SQL_Befehl_ausführen cSQL
    
    
    '******************
    '***** KUNDEN *****
    '******************
    
    cTabelle = "KUNDEN"
    
    cSQL = "Drop Index RECHNR on KUNDEN"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index KUNDNR on KUNDEN"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index KUERZEL on KUNDEN"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index NAME on KUNDEN"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index Titel on KUNDEN"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index PLZ on KUNDEN"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index STADT on KUNDEN"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index KUNDKART on KUNDEN"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index AENDER on KUNDEN"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index LASTDATE on KUNDEN"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index LASTDATETIME on KUNDEN"
    SQL_Befehl_ausführen cSQL
    
    '******************
    '***** KUNDKASS *****
    '******************
    
    cTabelle = "KUNDKASS"
    
    cSQL = "Drop Index adate on KUNDKASS"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index KUNDNR on KUNDKASS"
    SQL_Befehl_ausführen cSQL
    
    
    '*****************
    '***** LISRT *****
    '*****************
    
    cTabelle = "LISRT"
    
    cSQL = "Drop Index LINR on LISRT"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index KUERZEL on LISRT"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index LIEFBEZ on LISRT"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index PLZ on LISRT"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index STADT on LISRT"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index LASTDATETIME on LISRT"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index LASTDATE on LISRT"
    SQL_Befehl_ausführen cSQL
    
    
    '*******************
    '***** ARTLIEF *****
    '*******************
    
    cTabelle = "ARTLIEF"
    
    cSQL = "Drop Index ARTNR on ARTLIEF"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index LINR on ARTLIEF"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index ARTLINR on ARTLIEF"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index LIBESNR on ARTLIEF"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index SYNSTATUS on ARTLIEF"
    SQL_Befehl_ausführen cSQL
    
   
    '*******************
    '***** GUTSCH ******
    '*******************
    
    cTabelle = "GUTSCH"
    
    cSQL = "Drop Index DAT_EINL on GUTSCH"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index GUTSCHNR on GUTSCH"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index LASTTIME on GUTSCH"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Drop Index LASTDATE on GUTSCH"
    SQL_Befehl_ausführen cSQL
    

'    '*******************
'    '***** ZBESTAND *****
'    '*******************
'    cTabelle = "ZBESTAND"
'
'    If Not tableSuchenDBKombi("ZBESTAND", 1) Then
'
'    Else
'
'        cSQL = "Drop Index PRIMKEY on ZBESTAND"
'        SQL_Befehl_ausführen cSQL
'
'        cSQL = "Drop Index ARTNR on ZBESTAND"
'        SQL_Befehl_ausführen cSQL
'
'        cSQL = "Drop Index LASTDATE on ZBESTAND"
'        SQL_Befehl_ausführen cSQL
'
'    End If
'
    '*************
    '** Zugang **
    '*************
    cTabelle = "ZUGANG"
    cSQL = "Drop Index ARTNR on ZUGANG"
    SQL_Befehl_ausführen cSQL
    

    '******************************************************************
    '***** Neue Indices erzeugen **************************************
    '******************************************************************
     '*************
    '** ZUGANG **
    '*************
    

    cSQL = "Create Index ARTNR on ZUGANG(ARTNR) "
    SQL_Befehl_ausführen cSQL
               
    '********************
    '***** KASSJOUR *****
    '********************
    
   
    
    cSQL = "Create Index ADATE on KASSJOUR (ADATE)"
    SQL_Befehl_ausführen cSQL
    
   
    
    cSQL = "Create Index ARTNR on KASSJOUR (ARTNR)"
    SQL_Befehl_ausführen cSQL
    
    
    
    cSQL = "Create Index KUNDNR on KASSJOUR (KUNDNR)"
    SQL_Befehl_ausführen cSQL
    
    
    
    cSQL = "Create Index MOPREIS on KASSJOUR (MOPREIS)"
    SQL_Befehl_ausführen cSQL
    
    
    
    cSQL = "Create Index DATUM on UMSATZ (DATUM)"
    SQL_Befehl_ausführen cSQL
    
    
    
    cSQL = "Create Index PRIMKEY on UMS_ART (ARTNR, JAHR, MONAT)"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Create Index DATUM on UMS_ART (JAHR, MONAT)"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Create Index ARTNR on UMS_ART (ARTNR)"
    SQL_Befehl_ausführen cSQL
    
    
    cSQL = "Create Index PRIMKEY on UMS_ARTF (ARTNR, FILIALNR, JAHR, MONAT)"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Create Index Monat on UMS_ARTF (Monat)"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Create Index Jahr on UMS_ARTF (Jahr)"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Create Index PRIMKEY on UMSARTJF (ARTNR, FILIALNR, JAHR)"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Create Index PRIMKEY on UMS_LIEF (LINR, JAHR, MONAT)"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Create Index DATUM on UMS_LIEF (JAHR, MONAT)"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Create Index LINR on UMS_LIEF (LINR)"
    SQL_Befehl_ausführen cSQL
    
    
    cSQL = "Create Index PRIMKEY on UMSARTJ (ARTNR, JAHR)"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Create Index JAHR on UMSARTJ (JAHR)"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Create Index ARTNR on UMSARTJ (ARTNR)"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Create Index PRIMKEY on UMSKDJ (KUNDNR, JAHR)"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Create Index JAHR on UMSKDJ (JAHR)"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Create Index KUNDNR on UMSKDJ (KUNDNR)"
    SQL_Befehl_ausführen cSQL

    
    cSQL = "Create Index PRIMKEY on UMSLIEFJ (LINR, JAHR)"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Create Index JAHR on UMSLIEFJ (JAHR)"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Create Index LIEFNR on UMSLIEFJ (LINR)"
    SQL_Befehl_ausführen cSQL
    
    
    cSQL = "Create Index KUNDNR on KUNDEN (KUNDNR)"
    SQL_Befehl_ausführen cSQL
    
    
    cSQL = "Create Index KUERZEL on KUNDEN (KUERZEL)"
    SQL_Befehl_ausführen cSQL
    
    
    cSQL = "Create Index NAME on KUNDEN (NAME)"
    SQL_Befehl_ausführen cSQL
    

    
    cSQL = "Create Index Titel on KUNDEN (Titel)"
    SQL_Befehl_ausführen cSQL
    
    
    cSQL = "Create Index PLZ on KUNDEN (PLZ)"
    SQL_Befehl_ausführen cSQL
    
    
    cSQL = "Create Index STADT on KUNDEN (STADT)"
    SQL_Befehl_ausführen cSQL
    
   
    
    cSQL = "Create Index KUNDKART on KUNDEN (KUNDKART)"
    SQL_Befehl_ausführen cSQL
    
    
    
    cSQL = "Create Index AENDER on KUNDEN (AENDER)"
    SQL_Befehl_ausführen cSQL
    
    
    
    cSQL = "Create Index RECHNR on KUNDEN (RECHNR)"
    SQL_Befehl_ausführen cSQL
    
    
    cSQL = "Create Index LastDate on KUNDEN (LastDate)"
    SQL_Befehl_ausführen cSQL
    
   
    
    
    
    cSQL = "Create Index KUNDNR on KUNDKASS (KUNDNR)"
    SQL_Befehl_ausführen cSQL
    
    
    cSQL = "Create Index ADATE on KUNDKASS (ADATE)"
    SQL_Befehl_ausführen cSQL
    
    
    cSQL = "Create Index DAT_EINL on GUTSCH (DAT_EINL)"
    SQL_Befehl_ausführen cSQL
    
   
    cSQL = "Create Index GUTSCHNR on GUTSCH (GUTSCHNR)"
    SQL_Befehl_ausführen cSQL

    
    cSQL = "Create Index LastDate on GUTSCH (LastDate)"
    SQL_Befehl_ausführen cSQL
    

    cSQL = "Create Index LASTTIME on GUTSCH (Lasttime)"
    SQL_Befehl_ausführen cSQL
      
    
    
    
    
    cSQL = "Create Index DATUM on KASSBON (DATUM)"
    SQL_Befehl_ausführen cSQL
    
    
    
    cSQL = "Create Index ARTNR on INTERART (ARTNR)"
    SQL_Befehl_ausführen cSQL
    
   
    
   
    
    cSQL = "Create Index BEDNU on BEDNAME (BEDNU)"
    SQL_Befehl_ausführen cSQL
    
    
    
    cSQL = "Create Index ARTNR on KOLLVERK (ARTNR)"
    SQL_Befehl_ausführen cSQL
    
    
    
    cSQL = "Create Index ADATE on KOLLVERK (ADATE)"
    SQL_Befehl_ausführen cSQL
    
    
    
    cSQL = "Create Index KUNDNR on KOLLVERK (KUNDNR)"
    SQL_Befehl_ausführen cSQL
    
    
   
    cSQL = "Create Index ARTNR on RETOURE (ARTNR)"
    SQL_Befehl_ausführen cSQL
    
    
    
    cSQL = "Create Index ADATE on RETOURE (ADATE)"
    SQL_Befehl_ausführen cSQL
    
    
    cSQL = "Create Index KUNDNR on RETOURE (KUNDNR)"
    SQL_Befehl_ausführen cSQL
    
    
   
    
    cSQL = "Create Index ADATE on TAUSCH (ADATE)"
    SQL_Befehl_ausführen cSQL
    
    
    
    cSQL = "Create Index ARTNR on TAUSCH (ARTNR)"
    SQL_Befehl_ausführen cSQL
    
   
    
    cSQL = "Create Index LINR on TAUSCH (LINR)"
    SQL_Befehl_ausführen cSQL
    
   
    
  
    cSQL = "Create Index LINR on LISRT (LINR)"
    SQL_Befehl_ausführen cSQL
    
    
    
    cSQL = "Create Index KUERZEL on LISRT (KUERZEL)"
    SQL_Befehl_ausführen cSQL
    
    
    
    cSQL = "Create Index LIEFBEZ on LISRT (LIEFBEZ)"
    SQL_Befehl_ausführen cSQL
    
   
    
    cSQL = "Create Index PLZ on LISRT (PLZ)"
    SQL_Befehl_ausführen cSQL
    
   
    
    cSQL = "Create Index STADT on LISRT (STADT)"
    SQL_Befehl_ausführen cSQL
    
    
    
    cSQL = "Create Index LastDate on LISRT (LastDate)"
    SQL_Befehl_ausführen cSQL
    
    
    
    cSQL = "Create Index ARTNR on ARTLIEF (ARTNR)"
    SQL_Befehl_ausführen cSQL
    
    
    
    cSQL = "Create Index LINR on ARTLIEF (LINR)"
    SQL_Befehl_ausführen cSQL
    
   
    
    cSQL = "Create Index ARTLINR on ARTLIEF (ARTNR, LINR)"
    SQL_Befehl_ausführen cSQL
    
    
    cSQL = "Create Index LIBESNR on ARTLIEF (LIBESNR)"
    SQL_Befehl_ausführen cSQL
    
    cSQL = "Create Index SYNSTATUS on ARTLIEF (SYNSTATUS)"
    SQL_Befehl_ausführen cSQL
    
    
   


'    If Not tableSuchenDBKombi("ZBESTAND", 1) Then
'
'    Else
'        cSQL = "Create Index PRIMKEY on ZBESTAND (FILIALNR, ARTNR)"
'        SQL_Befehl_ausführen cSQL
'
'        cSQL = "Create Index ARTNR on ZBESTAND (ARTNR)"
'        SQL_Befehl_ausführen cSQL
'
'         cSQL = "Create Index LastDate on ZBESTAND (LastDate)"
'        SQL_Befehl_ausführen cSQL
'    End If

    
    '****************************************
    '***** Abgleich ARTIKEL und ARTLIEF *****
    '****************************************
'    cTabelle = "ABGLEICH"
'
'    frmWKL00!Label2.Caption = "Gleiche ARTIKEL und ARTLIEF gegeneinander ab (Stufe 1)"
'    frmWKL00!Label2.Visible = True
'    frmWKL00!Label2.Refresh
'
'
'
'
'    loeschNEW "TEMP1", db
'    loeschNEW "TEMP2", db
'
'    BeginTrans
'    cSQL = "Select ARTNR, LINR , 'N' as Erkannt into TEMP1 from ARTIKEL"
'    SQL_Befehl_ausführen cSQL
'
'    frmWKL00!Label2.Caption = "Gleiche ARTIKEL und ARTLIEF gegeneinander ab (Stufe 2)"
'    frmWKL00!Label2.Visible = True
'    frmWKL00!Label2.Refresh
'
'    cSQL = "Update TEMP1 inner join Artlief on TEMP1.artnr = Artlief.artnr and TEMP1.LINR = Artlief.LINR "
'    cSQL = cSQL & " Set TEMP1.Erkannt = 'J' "
'    SQL_Befehl_ausführen cSQL
'
'    cSQL = "Delete from TEMP1 where erkannt = 'J' "
'    SQL_Befehl_ausführen cSQL
'
'    frmWKL00!Label2.Caption = "Gleiche ARTIKEL und ARTLIEF gegeneinander ab (Stufe 3)"
'    frmWKL00!Label2.Visible = True
'    frmWKL00!Label2.Refresh
'
'    cSQL = "Select ARTIKEL.ARTNR, ARTIKEL.LINR, ARTIKEL.LEKPR, ARTIKEL.LIBESNR, ARTIKEL.MINMEN "
'    cSQL = cSQL & "into TEMP2 from ARTIKEL inner join TEMP1 on ARTIKEL.ARTNR = TEMP1.ARTNR "
'    SQL_Befehl_ausführen cSQL
'
'    frmWKL00!Label2.Caption = "Gleiche ARTIKEL und ARTLIEF gegeneinander ab (Stufe 4)"
'    frmWKL00!Label2.Visible = True
'    frmWKL00!Label2.Refresh
'
'
'    cSQL = "Insert into ARTLIEF Select * from TEMP2"
'    SQL_Befehl_ausführen cSQL
'
'    frmWKL00!Label2.Caption = "Gleiche ARTIKEL und ARTLIEF gegeneinander ab (Stufe 5)"
'    frmWKL00!Label2.Visible = True
'    frmWKL00!Label2.Refresh
'
'    CommitTrans
'
'    loeschNEW "TEMP1", db
'    loeschNEW "TEMP2", db
    
    
    
    
    
    
    
    
    

    
    
    
    
    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:

    If err.Number = 3372 Or err.Number = 3371 Or err.Number = 3375 Or err.Number = 3256 Or err.Number = 3376 Or err.Number = 3043 Then
        Resume Next
    ElseIf err.Number = 3032 Then
        Resume Next
    Else
    
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "ReIndiziereArtikel_ForSS"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1
        Screen.MousePointer = 0
        Resume Next
    End If
End Sub
Public Function fnVKneuNS(dEK As Double, cMW As String, dNS As Double) As String
    On Error GoTo LOKAL_ERROR
    
    Dim dNVK    As Double
    Dim dVKNEU3    As Double

    fnVKneuNS = "0"
    
    If Val(dNS) = 100 Then
        fnVKneuNS = "0"
        Exit Function
    End If
     
    dNVK = (dEK * 100) / (100 - dNS)
    
    If cMW = "V" Then
        dVKNEU3 = dNVK * CLng("1." & CStr(gdMWStV)) / 100
    ElseIf cMW = "E" Then
        dVKNEU3 = dNVK * CLng("1." & CStr(gdMWStE)) / 100
    Else
        dVKNEU3 = dNVK
    End If
    
    fnVKneuNS = Format$(dVKNEU3, "#####0.00")
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "fnVKneuNS"
    Fehler.gsFehlertext = "Bei der Berechnung des Kassenverkaufspreises über die Nettospanne ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function NettoErtrag(dKVP As String, dEK As String, cMwst As String) As String
    On Error GoTo LOKAL_ERROR
            
    Dim dSpanne1    As Double
    Dim dSpanne2    As Double
    
    NettoErtrag = "0"
    'als Erstes Parameterüberprüfung
    If IsNumeric(dEK) = False Then Exit Function
    
    If IsNumeric(dKVP) = False Then Exit Function
    
    If Trim(cMwst) <> "V" And Trim(cMwst) <> "E" And Trim(cMwst) <> "O" Then 'cMWSt
        Exit Function
    End If
    
    NettoErtrag = "0"
    
    If cMwst = "V" Then
        dSpanne1 = (dKVP * 100) / (100 + gdMWStV)
    ElseIf cMwst = "E" Then
        dSpanne1 = (dKVP * 100) / (100 + gdMWStE)
    Else
        dSpanne1 = (dKVP * 100) / 100
    End If
    
    dSpanne2 = dSpanne1 - dEK
    NettoErtrag = dSpanne2
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "NettoErtrag"
    Fehler.gsFehlertext = "Bei der Berechnung der Nettospanne ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function NettospanneInProzent(dKVP As String, dEK As String, cMwst As String) As String
    On Error GoTo LOKAL_ERROR
            
    Dim dSpanne     As Double
    Dim dSpanne1    As Double
    Dim dSpanne2    As Double
    Dim dSpanne3    As Double
    Dim dSpanne4    As Double
    
    NettospanneInProzent = "0"
    'als Erstes Parameterüberprüfung
    If IsNumeric(dEK) = False Then 'dek
        Exit Function
    End If
    
    If IsNumeric(dKVP) = False Then 'dkvp
        Exit Function
    End If
    
    If Trim(cMwst) <> "V" And Trim(cMwst) <> "E" And Trim(cMwst) <> "O" Then 'cMWSt
        Exit Function
    End If
    
   
    
    NettospanneInProzent = "0"
    
    If cMwst = "V" Then
        dSpanne1 = (dKVP * 100) / (100 + gdMWStV)
    ElseIf cMwst = "E" Then
        dSpanne1 = (dKVP * 100) / (100 + gdMWStE)
    Else
        dSpanne1 = (dKVP * 100) / 100
    End If
    
    dSpanne2 = dSpanne1 - dEK
    
    
    If dSpanne1 <> 0 Then
        dSpanne3 = (dSpanne2 * 100) / dSpanne1
    Else
        dSpanne3 = 0
    End If
    
    NettospanneInProzent = Format$(dSpanne3, "#####0.00")

    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "NettospanneInProzent"
    Fehler.gsFehlertext = "Bei der Berechnung der Nettospanne ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function NettospanneInProzent_neu(cArtNr As String) As String
    On Error GoTo LOKAL_ERROR
            
    Dim dSpanne     As Double
    Dim dSpanne1    As Double
    Dim dSpanne2    As Double
    Dim dSpanne3    As Double
    Dim dSpanne4    As Double
    Dim rsrs        As DAO.Recordset
    Dim sSQL        As String
    Dim dKVP        As Double
    Dim dEK         As Double
    Dim cMwst       As String
    
    NettospanneInProzent_neu = "0"
    
    sSQL = "Select MWST,KVKPR1,EKPR from Artikel where artnr = " & cArtNr
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!KVKPR1) Then
            dKVP = rsrs!KVKPR1
        End If
        
        If Not IsNull(rsrs!MWST) Then
            cMwst = rsrs!MWST
        End If
        
        If Not IsNull(rsrs!ekpr) Then
            dEK = rsrs!ekpr
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    If gsSpanne = "LEK" Then
    
        If gbLekMax Then
            sSQL = "Select Max(LEKPR) as ERG from artlief where artnr = " & cArtNr
        Else
            sSQL = "Select min(LEKPR) as ERG from artlief where artnr = " & cArtNr
        End If
        sSQL = sSQL & " and LEKPR > 0 "
        
        Set rsrs = gdBase.OpenRecordset(sSQL)
        If Not rsrs.EOF Then
            If Not IsNull(rsrs!ERG) Then
                dEK = rsrs!ERG
            End If
            
        End If
        rsrs.Close: Set rsrs = Nothing
        
    
    End If
    
    
    
    'als Erstes Parameterüberprüfung
    If IsNumeric(dEK) = False Then 'dek
        Exit Function
    End If
    
    If IsNumeric(dKVP) = False Then 'dkvp
        Exit Function
    End If
    
    If Trim(cMwst) <> "V" And Trim(cMwst) <> "E" And Trim(cMwst) <> "O" Then 'cMWSt
        Exit Function
    End If
    
   
    
    NettospanneInProzent_neu = "0"
    
    If cMwst = "V" Then
        dSpanne1 = (dKVP * 100) / (100 + gdMWStV)
    ElseIf cMwst = "E" Then
        dSpanne1 = (dKVP * 100) / (100 + gdMWStE)
    Else
        dSpanne1 = (dKVP * 100) / 100
    End If
    
    dSpanne2 = dSpanne1 - dEK
    
    
    If dSpanne1 <> 0 Then
        dSpanne3 = (dSpanne2 * 100) / dSpanne1
    Else
        dSpanne3 = 0
    End If
    
    NettospanneInProzent_neu = Format$(dSpanne3, "#####0.00")

    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "NettospanneInProzent_neu"
    Fehler.gsFehlertext = "Bei der Berechnung der Nettospanne ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Sub Markenabgleich(sTab As String, db As Database)
On Error GoTo LOKAL_ERROR

    Dim sSQL        As String
    
    sSQL = "Update " & sTab & " inner join LINBEZ on " & sTab & ".linr = LINBEZ.linr and " & sTab & ".lpz = LINBEZ.lpz "
    sSQL = sSQL & " Set " & sTab & ".marke = LINBEZ.Marke "
    sSQL = sSQL & " , " & sTab & ".LINBEZ = LINBEZ.LINBEZEICH"
    db.Execute sSQL, dbFailOnError
     
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "Markenabgleich"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Public Sub ermittelnPL(bytesort As Byte, lbl1 As Label, txt As TextBox, Listx As ListBox, cWelche As String, txtvon As DTPicker, txtbis As DTPicker, cboPL As ComboBox)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    Dim cVon            As String
    Dim cBis            As String
    Dim lVon            As Long
    Dim lBis            As Long
    Dim i               As Integer
    Dim iFil            As Integer
    Dim iPreislage      As Integer
    Dim cPreisvon       As String
    Dim cPreisbis       As String
    Dim gesNSP          As Double
    Dim gesUmsatzNetto  As Double
    Dim GesUmsatz       As Double
    Dim gesErtrag       As Double
    Dim gesMenge        As Long
    Dim sSQLAGN         As String
    Dim cLBSatz         As String
    
    Screen.MousePointer = 11
    
    anzeigeNew "normal", "Daten werden ermittelt...", lbl1
    
    loeschNEW "PLKOPF", gdBase
    CreateTable "PLKOPF", gdBase
    
    Select Case cWelche
        Case Is = "ALLES"
            sSQL = "Insert into PLKOPF (AuswahlTEXT) "
            sSQL = sSQL & " Values (  '')"
            schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
            
            sSQL = "Update PLKOPF set sHEAD1 = ''"
            sSQL = sSQL & ", sHEAD2 = ''"
            gdBase.Execute sSQL, dbFailOnError
            
        Case Is = "PGN"
    
            If txt.Text <> "alle" Or IsNumeric(txt.Text) Then
                If txt.Text = "" Then
                    If Listx.ListCount = 0 Then
                        sSQLAGN = ""
                        sSQL = "Insert into PLKOPF (AuswahlTEXT) "
                        sSQL = sSQL & " Values (  '')"
                        schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
                    Else
                        If LoesePGNInArtnr(Mid(Listx.list(0), 1, InStr(1, Listx.list(0), " ")), False, gdBase) = True Then
                            sSQLAGN = " and artnr in (select artnr from my" & srechnertab & ")"
                        End If
                        
                        cLBSatz = Listx.list(0)
                    
                        For i = 1 To Listx.ListCount - 1
                            If LoesePGNInArtnr(Mid(Listx.list(i), 1, InStr(1, Listx.list(i), " ")), True, gdBase) = True Then
                                sSQLAGN = " and artnr in (select artnr from my" & srechnertab & ")"
                            End If
                            cLBSatz = cLBSatz & ", " & Listx.list(i)
                        Next i
                        sSQL = "Insert into PLKOPF (AuswahlTEXT) "
                        sSQL = sSQL & " Values (  '" & cLBSatz & "')"
                        schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
                    End If
                Else
                    If Trim$(txt.Text) <> "" Then
                        If LoesePGNInArtnr(Trim$(txt.Text), False, gdBase) = True Then
                            sSQLAGN = " and artnr in (select artnr from my" & srechnertab & ")"
                        End If
                    End If
                    sSQL = "Insert into PLKOPF (AuswahlTEXT) "
                    sSQL = sSQL & " Values (  '" & txt.Text & "')"
                    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
                End If
            Else
                sSQLAGN = ""
                sSQL = "Insert into PLKOPF (AuswahlTEXT) "
                sSQL = sSQL & " Values (  '" & txt.Text & "')"
                schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
            End If
            
            sSQL = "Update PLKOPF set sHEAD1 = 'PGN'"
            sSQL = sSQL & ", sHEAD2 = 'PGN'"
            gdBase.Execute sSQL, dbFailOnError
        Case Is = "AGN"
            If txt.Text <> "alle" Or IsNumeric(txt.Text) Then
                If txt.Text = "" Then
                    If Listx.ListCount = 0 Then
                        sSQLAGN = ""
                        sSQL = "Insert into PLKOPF (AuswahlTEXT) "
                        sSQL = sSQL & " Values (  '')"
                        schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
                    Else
                        sSQLAGN = "and (AGN= " & Mid(Listx.list(0), 1, InStr(1, Listx.list(0), " "))
                        For i = 1 To Listx.ListCount - 1
                            sSQLAGN = sSQLAGN & " or AGN= " & Mid(Listx.list(i), 1, InStr(1, Listx.list(i), " "))
                        Next i
                        sSQLAGN = sSQLAGN & ")"
                        
                        cLBSatz = Listx.list(0)
                        For i = 1 To Listx.ListCount - 1
                            cLBSatz = cLBSatz & ", " & Listx.list(i)
                        Next i
                        
                        sSQL = "Insert into PLKOPF (AuswahlTEXT) "
                        sSQL = sSQL & " Values (  '" & cLBSatz & "')"
                        schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
                    End If
                Else
                    sSQLAGN = " and agn = " & txt.Text

                    sSQL = "Insert into PLKOPF (AuswahlTEXT) "
                    sSQL = sSQL & " Values (  '" & txt.Text & "')"
                    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
                End If
            Else
                sSQLAGN = ""
                sSQL = "Insert into PLKOPF (AuswahlTEXT) "
                sSQL = sSQL & " Values (  '" & txt.Text & "')"
                schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
            End If
            
            sSQL = "Update PLKOPF set sHEAD1 = 'AGN'"
            sSQL = sSQL & ", sHEAD2 = 'AGN'"
            gdBase.Execute sSQL, dbFailOnError
        
        End Select
    
    cVon = txtvon.value
    cBis = txtbis.value
    
    lVon = DateValue(cVon)
    lBis = DateValue(cBis)

    cVon = Trim$(Str$(lVon))
    cBis = Trim$(Str$(lBis))
    
    loeschNEW "PreisKass", gdBase
    CreateTable "PREISKASS", gdBase
    
    loeschNEW "PreisKassT", gdBase
    CreateTable "PREISKASST", gdBase
    
    If cboPL.Text = "alle Preislagen" Then
        iPreislage = 0
    Else
        iPreislage = CInt(Right(cboPL.Text, 3))
    End If

    iFil = 0

    byteanzPreisl = ermanzpreislagen
    ermPreislagen
    
    GesUmsatz = ermgesUmsatz(cVon, cBis, iFil, sSQLAGN)
    gesErtrag = ermgesErtrag(cVon, cBis, iFil, sSQLAGN)
    gesMenge = ermgesMenge(cVon, cBis, iFil, sSQLAGN)
    gesUmsatzNetto = ermgesUmsatzNetto(cVon, cBis, iFil, sSQLAGN)
    
    If gesUmsatzNetto <> 0 Then
        gesNSP = gesErtrag * 100 / gesUmsatzNetto
    End If
    
    loeschNEW "PREISLGU", gdBase
    CreateTable "PREISLGU", gdBase
    
    sSQL = "Insert into PREISLGU (Umsatz,UmsatzNetto,gesnsp,Ertrag,Menge,von,bis,fil,prausw) "
    sSQL = sSQL & " Values (  '" & GesUmsatz & "','" & gesUmsatzNetto & "','" & gesNSP & "','" & gesErtrag & "','" & gesMenge & "'"
    sSQL = sSQL & " ,'" & txtvon.value & "','" & txtbis.value & "', 'alle Filialen', '" & Trim(cboPL.Text) & "'"
    sSQL = sSQL & " )"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into PreisKassT Select "
    sSQL = sSQL & " artnr "
    sSQL = sSQL & " , bezeich "
    sSQL = sSQL & " , preis/menge as kvkp "
    sSQL = sSQL & " , preis "
    sSQL = sSQL & " , menge "
    sSQL = sSQL & " , linr "
    sSQL = sSQL & " , ekpr "
    sSQL = sSQL & " , mwst "
    sSQL = sSQL & " , Lpz "
    sSQL = sSQL & " , ean "
    sSQL = sSQL & " , agn "
    sSQL = sSQL & " , adate "
    sSQL = sSQL & " , azeit "
    sSQL = sSQL & " , filiale as filnr "
    sSQL = sSQL & " from Kassjour "
    sSQL = sSQL & " where adate between  " & cVon & " And " & cBis
    sSQL = sSQL & sSQLAGN
    sSQL = sSQL & " and UMS_OK = 'J' "
    sSQL = sSQL & " and menge <> 0 "
    
    If iFil = 0 Then

    Else
        sSQL = sSQL & " and filiale = " & iFil
    End If
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    
    If Not Datendrin("PreisKassT", gdBase) Then
        Screen.MousePointer = 0
        anzeigeNew "rot", "Keine Daten gefunden.", lbl1
        Exit Sub
    End If
    
    anzeigeNew "normal", "Rohertrag wird errechnet...", lbl1
    
    sSQL = "Update PreisKasst set rertrag = ((Preis * 100)/(100 + " & gdMWStV & ")) - (EKPR * menge) "
    sSQL = sSQL & " where mwst = 'V' "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update PreisKasst set rertrag = ((Preis * 100)/(100 + " & gdMWStE & ")) - (EKPR * menge) "
    sSQL = sSQL & " where mwst = 'E' "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update PreisKasst set rertrag = ((Preis * 100)/(100 + " & gdMWStO & " )) - (EKPR * menge) "
    sSQL = sSQL & " where mwst = 'O' "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    anzeigeNew "normal", "Nettoumsätze werden errechnet...", lbl1
    
    sSQL = "Update PreisKasst set npreis = ((Preis * 100)/(100 + " & gdMWStV & ")) "
    sSQL = sSQL & " where mwst = 'V' "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update PreisKasst set npreis = ((Preis * 100)/(100 + " & gdMWStE & "))  "
    sSQL = sSQL & " where mwst = 'E' "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update PreisKasst set npreis = ((Preis * 100)/(100 + " & gdMWStO & " ))  "
    sSQL = sSQL & " where mwst = 'O' "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    If iPreislage = 0 Then
        For i = 1 To byteanzPreisl
        
            cPreisvon = newPreislage(i - 1).PreisVon
            cPreisvon = SwapStr(cPreisvon, ",", ".")
            cPreisbis = newPreislage(i - 1).PreisBis
            cPreisbis = SwapStr(cPreisbis, ",", ".")
        
            sSQL = "Insert into PreisKass Select "
            sSQL = sSQL & " artnr "
            sSQL = sSQL & " , bezeich "
            sSQL = sSQL & " , preis "
            sSQL = sSQL & " , npreis "
            sSQL = sSQL & " , menge "
            sSQL = sSQL & " , rertrag "
            sSQL = sSQL & " , ekpr "
            sSQL = sSQL & " , kvkp "
            sSQL = sSQL & " , linr "
            sSQL = sSQL & " , Lpz "
            sSQL = sSQL & " , ean "
            sSQL = sSQL & " , agn "
            sSQL = sSQL & " , filnr "
            sSQL = sSQL & " , adate "
            sSQL = sSQL & " , azeit "
            sSQL = sSQL & " , '" & newPreislage(i - 1).Preislagentext & "' as PreislText "
            sSQL = sSQL & " , " & newPreislage(i - 1).PreislagenNr & " as Preislnr "
            sSQL = sSQL & " from PreisKasst "
            sSQL = sSQL & " where kvkp between  " & cPreisvon & " And " & cPreisbis

            If iFil = 0 Then

            Else
                sSQL = sSQL & " and Filnr = " & iFil
            End If
            schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
        
        Next i
    Else
    
        cPreisvon = newPreislage(iPreislage - 1).PreisVon
        cPreisvon = SwapStr(cPreisvon, ",", ".")
        cPreisbis = newPreislage(iPreislage - 1).PreisBis
        cPreisbis = SwapStr(cPreisbis, ",", ".")
    
        sSQL = "Insert into PreisKass Select "
        sSQL = sSQL & " artnr "
        sSQL = sSQL & " , bezeich "
        sSQL = sSQL & " , preis "
        sSQL = sSQL & " , npreis "
        sSQL = sSQL & " , menge "
        sSQL = sSQL & " , rertrag "
        sSQL = sSQL & " , ekpr "
        sSQL = sSQL & " , kvkp "
        sSQL = sSQL & " , linr "
        sSQL = sSQL & " , Lpz "
        sSQL = sSQL & " , ean "
        sSQL = sSQL & " , agn "
        sSQL = sSQL & " , filnr "
        sSQL = sSQL & " , adate "
        sSQL = sSQL & " , azeit "
        sSQL = sSQL & " , '" & newPreislage(iPreislage - 1).Preislagentext & "' as PreislText "
        sSQL = sSQL & " , " & newPreislage(iPreislage - 1).PreislagenNr & " as Preislnr "
        sSQL = sSQL & " from PreisKasst "
        sSQL = sSQL & " where kvkp between  " & cPreisvon & " And " & cPreisbis
        
        schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    End If
    
    loeschNEW "PreisLZU", gdBase
    CreateTable "PREISLZU", gdBase
    
    anzeigeNew "normal", "Druckvorschau wird erstellt...", lbl1
    
    Dim Upropl As String
    Dim Epropl As String
    Dim Mpropl As String
    Dim NUpropl As String
    Dim NSPpropl As Double
    
    If iPreislage = 0 Then
        For i = 1 To byteanzPreisl
        
            Upropl = ermUmsatzpropreislage(i)
            Epropl = ermErtragpropreislage(i)
            Mpropl = ermMengepropreislage(i)
            NUpropl = ermUmsatzNettopropreislage(i)
            
            If CDbl(NUpropl) <> 0 Then
                NSPpropl = (CDbl(Epropl) * 100) / CDbl(NUpropl)
            End If
        
            sSQL = "Insert into PREISLZU (Umsatz,NettoUmsatz,NSP,Ertrag,Menge,PreislText ,PreislNr) "
            sSQL = sSQL & " Values (  '" & Upropl & "','" & NUpropl & "','" & NSPpropl & "','" & Epropl & "','" & Mpropl & "'"
            sSQL = sSQL & " , '" & newPreislage(i - 1).Preislagentext & "' "
            sSQL = sSQL & " , " & newPreislage(i - 1).PreislagenNr & " "
            sSQL = sSQL & " )"
            schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
        
        Next i
    Else
    
        Upropl = ermUmsatzpropreislage(iPreislage)
        Epropl = ermErtragpropreislage(iPreislage)
        Mpropl = ermMengepropreislage(iPreislage)
        NUpropl = ermUmsatzNettopropreislage(iPreislage)
            
        If CDbl(NUpropl) <> 0 Then
            NSPpropl = (CDbl(Epropl) * 100) / CDbl(NUpropl)
        End If
        
        sSQL = "Insert into PREISLZU (Umsatz,NettoUmsatz,NSP,Ertrag,Menge,PreislText ,PreislNr) "
        sSQL = sSQL & " Values (  '" & Upropl & "','" & NUpropl & "','" & NSPpropl & "','" & Epropl & "','" & Mpropl & "'"
        sSQL = sSQL & " , '" & newPreislage(iPreislage - 1).Preislagentext & "' "
        sSQL = sSQL & " , " & newPreislage(iPreislage - 1).PreislagenNr & " "
        sSQL = sSQL & " )"
        schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    End If
    
'    ermittle rest
    
    
    Dim crestU As String
    Dim crestE As String
    Dim crestM As String
    Dim crestUN As String
    
    NUpropl = ermUmsatzNettoGespreislage
    Upropl = ermUmsatzGespreislage
    Epropl = ermErtragGespreislage
    Mpropl = ermMengeGespreislage
    
    
    crestUN = CStr(gesUmsatzNetto - CDbl(NUpropl))
    crestU = CStr(GesUmsatz - CDbl(Upropl))
    crestE = CStr(gesErtrag - CDbl(Epropl))
    crestM = CStr(gesMenge - CDbl(Mpropl))
   
    
    sSQL = "Insert into PreisKass (preis,npreis,rertrag,Menge,PreislText ,PreislNr) "
    sSQL = sSQL & " Values (  '" & crestU & "','" & crestUN & "','" & crestE & "','" & crestM & "'"
    sSQL = sSQL & " , 'nicht definiert' "
    sSQL = sSQL & " , " & byteanzPreisl + 1
    sSQL = sSQL & " )"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    anzeigeNew "normal", "", lbl1
    
    Screen.MousePointer = 0
    
    If bytesort = 0 Then
        reportbildschirm "dsd", "aWKL121"
    ElseIf bytesort = 1 Then
        reportbildschirm "dsd", "aWKL121a"
    ElseIf bytesort = 2 Then
        reportbildschirm "dsd", "aWKL121c"
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ermittelnPL"
    Fehler.gsFehlertext = "In der Preislagenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Public Function ermgesUmsatz(cVon As String, cBis As String, iFil As Integer, sAGN As String) As Double
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    
    ermgesUmsatz = 0
    
    sSQL = "Select sum(preis) as maxi"
    sSQL = sSQL & " from Kassjour "
    sSQL = sSQL & " where adate between  " & cVon & " And " & cBis
    sSQL = sSQL & sAGN
    sSQL = sSQL & " and UMS_OK = 'J' "
    
    
    If iFil = 0 Then
    
    Else
        sSQL = sSQL & " and filiale = " & iFil
    End If
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermgesUmsatz = rsrs!maxi
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
                
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ermgesUmsatz"
    Fehler.gsFehlertext = "In der Preislagenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Function
Public Function ermgesMB(cOperator As String) As Long
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    
    ermgesMB = 0
    
    sSQL = "Select count(*) as maxi "
    sSQL = sSQL & " from Artikel where "
    sSQL = sSQL & " Bestand " & cOperator & " Minbest"
    sSQL = sSQL & " and Artikel.gefuehrt = 'J' and Artikel.LPZ <> 0"
    sSQL = sSQL & " and Minbest > 0 "
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermgesMB = rsrs!maxi
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
                
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ermgesMB"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Function
Public Function ermgesMBSchnittEK(cOperator As String) As Single
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    
    ermgesMBSchnittEK = 0
    
    sSQL = "Select sum((Bestand - Minbest) * ekpr) as maxi "
    sSQL = sSQL & " from Artikel where "
    sSQL = sSQL & " Bestand " & cOperator & " Minbest"
    sSQL = sSQL & " and Minbest > 0 "
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermgesMBSchnittEK = rsrs!maxi
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
                
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ermgesMBSchnittEK"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Function
Public Function ermgesUmsatzNetto(cVon As String, cBis As String, iFil As Integer, sAGN As String) As Double
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    
    ermgesUmsatzNetto = 0
    
    sSQL = "Select sum(preis) as maxi"
    sSQL = sSQL & " from Kassjour "
    sSQL = sSQL & " where adate between  " & cVon & " And " & cBis
    sSQL = sSQL & sAGN
    sSQL = sSQL & " and UMS_OK = 'J' "
    sSQL = sSQL & " and MWST = 'V' "
    
    If iFil = 0 Then
    
    Else
        sSQL = sSQL & " and filiale = " & iFil
    End If
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermgesUmsatzNetto = (rsrs!maxi * 100 / (100 + gdMWStV))
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    sSQL = "Select sum(preis) as maxi"
    sSQL = sSQL & " from Kassjour "
    sSQL = sSQL & " where adate between  " & cVon & " And " & cBis
    sSQL = sSQL & sAGN
    sSQL = sSQL & " and UMS_OK = 'J' "
    sSQL = sSQL & " and MWST = 'E' "
    
    If iFil = 0 Then
    
    Else
        sSQL = sSQL & " and filiale = " & iFil
    End If
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermgesUmsatzNetto = ermgesUmsatzNetto + (rsrs!maxi * 100 / (100 + gdMWStE))
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    sSQL = "Select sum(preis) as maxi"
    sSQL = sSQL & " from Kassjour "
    sSQL = sSQL & " where adate between  " & cVon & " And " & cBis
    sSQL = sSQL & sAGN
    sSQL = sSQL & " and UMS_OK = 'J' "
    sSQL = sSQL & " and MWST = 'O' "
    
    If iFil = 0 Then
    
    Else
        sSQL = sSQL & " and filiale = " & iFil
    End If
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermgesUmsatzNetto = ermgesUmsatzNetto + rsrs!maxi
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
                
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ermgesUmsatzNetto"
    Fehler.gsFehlertext = "In der Preislagenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Function
Public Function ermgesMenge(cVon As String, cBis As String, iFil As Integer, sAGN As String) As Double
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    
    ermgesMenge = 0
    
    sSQL = "Select sum(menge) as maxi"
    sSQL = sSQL & " from Kassjour "
    sSQL = sSQL & " where adate between  " & cVon & " And " & cBis
    sSQL = sSQL & sAGN
    sSQL = sSQL & " and UMS_OK = 'J' "
    If iFil = 0 Then
    
    Else
        sSQL = sSQL & " and filiale = " & iFil
    End If
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermgesMenge = rsrs!maxi
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
                
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ermgesMenge"
    Fehler.gsFehlertext = "In der Preislagenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermgesErtrag(cVon As String, cBis As String, iFil As Integer, sAGN As String) As Double
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    
    ermgesErtrag = 0
    
    loeschNEW "ErtragTe", gdBase
    CreateTable "ERTRAGTE", gdBase
    
    sSQL = "Insert into ErtragTe Select "
    sSQL = sSQL & " artnr "
    sSQL = sSQL & " , bezeich "
    sSQL = sSQL & " , preis/menge as kvkp "
    sSQL = sSQL & " , preis "
    sSQL = sSQL & " , menge "
    sSQL = sSQL & " , linr "
    sSQL = sSQL & " , ekpr "
    sSQL = sSQL & " , mwst "
    sSQL = sSQL & " , Lpz "
    sSQL = sSQL & " , ean "
    sSQL = sSQL & " , agn "
    sSQL = sSQL & " , filiale as filnr "
    sSQL = sSQL & " from Kassjour "
    sSQL = sSQL & " where adate between  " & cVon & " And " & cBis
    sSQL = sSQL & sAGN
    sSQL = sSQL & " and UMS_OK = 'J' "
    sSQL = sSQL & " and menge <> 0 "
    
    If iFil = 0 Then
    
    Else
        sSQL = sSQL & " and filiale = " & iFil
    End If
    
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update ErtragTe set rertrag = ((Preis * 100)/(100 + " & gdMWStV & ")) - (EKPR * menge) "
    sSQL = sSQL & " where mwst = 'V' "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update ErtragTe set rertrag = ((Preis * 100)/(100 + " & gdMWStE & ")) - (EKPR * menge) "
    sSQL = sSQL & " where mwst = 'E' "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update ErtragTe set rertrag = ((Preis * 100)/(100 + " & gdMWStO & " )) - (EKPR * menge) "
    sSQL = sSQL & " where mwst = 'O' "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    
    sSQL = "Select sum(rertrag) as maxi from ErtragTe "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermgesErtrag = rsrs!maxi
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
                
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ermgesErtrag"
    Fehler.gsFehlertext = "In der Preislagenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Sub ermPreislagen()
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    Dim j           As Long
    Dim cSatz       As String
    Dim cFeld       As String
    
    If byteanzPreisl = 0 Then Exit Sub
    
    ReDim newPreislage(0 To byteanzPreisl - 1)
    
    For j = 0 To byteanzPreisl - 1
        newPreislage(j).PreisVon = 0
        newPreislage(j).PreisBis = 0
        newPreislage(j).PreislagenNr = 0
        newPreislage(j).Preislagentext = ""
    Next j
    
    j = 0
    sSQL = "Select * from Preisl order by lfnr "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            If Not IsNull(rsrs!vonP) Then
            
                cFeld = Format(rsrs!vonP, "##0.00 EUR ")
                cSatz = cSatz & Space(13 - Len(cFeld)) & cFeld
                
                If Not IsNull(rsrs!bisP) Then
                
                    cFeld = Format(rsrs!bisP, "##0.00 EUR")
                    cSatz = cSatz & "-" & Space(13 - Len(cFeld)) & cFeld
                    
                    If Not IsNull(rsrs!lfnr) Then
                    
                    newPreislage(j).PreisVon = rsrs!vonP
                    newPreislage(j).PreisBis = rsrs!bisP
                    newPreislage(j).PreislagenNr = rsrs!lfnr
                    newPreislage(j).Preislagentext = cSatz
                    j = j + 1
                    
                    cSatz = ""
                    cFeld = ""
                    
                    End If
                End If
            End If
        
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ermPreislagen"
    Fehler.gsFehlertext = "In der Preislagenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Sub
Private Function ermUmsatzpropreislage(iPreislage As Integer) As String
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    
    ermUmsatzpropreislage = "0"
    
    sSQL = "Select sum(preis) as maxi"
    sSQL = sSQL & " from Preiskass "
    sSQL = sSQL & " where PreislNr = " & iPreislage
    
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermUmsatzpropreislage = rsrs!maxi
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
                
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ermUmsatzpropreislage"
    Fehler.gsFehlertext = "In der Preislagenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Function
Private Function ermUmsatzNettopropreislage(iPreislage As Integer) As String
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    
    ermUmsatzNettopropreislage = "0"
    
    sSQL = "Select sum(Npreis) as maxi"
    sSQL = sSQL & " from Preiskass "
    sSQL = sSQL & " where PreislNr = " & iPreislage
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermUmsatzNettopropreislage = rsrs!maxi
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
                
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ermUmsatzNettopropreislage"
    Fehler.gsFehlertext = "In der Preislagenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Function
Private Function ermErtragpropreislage(iPreislage As Integer) As String
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    
    ermErtragpropreislage = "0"
    
    sSQL = "Select sum(rertrag) as maxi"
    sSQL = sSQL & " from Preiskass "
    sSQL = sSQL & " where PreislNr = " & iPreislage
    
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermErtragpropreislage = rsrs!maxi
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
                
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ermErtragpropreislage"
    Fehler.gsFehlertext = "In der Preislagenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Function
Private Function ermMengepropreislage(iPreislage As Integer) As String
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    
    ermMengepropreislage = "0"
    
    sSQL = "Select sum(menge) as maxi"
    sSQL = sSQL & " from Preiskass "
    sSQL = sSQL & " where PreislNr = " & iPreislage
    
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermMengepropreislage = rsrs!maxi
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
                
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ermMengepropreislage"
    Fehler.gsFehlertext = "In der Preislagenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Function
Public Function NettospanneInEuro(dKVP As String, dEK As String, cMwst As String) As String
    On Error GoTo LOKAL_ERROR
            
'    Dim dSpanne     As Double
    Dim dSpanne1    As Double
    Dim dSpanne2    As Double
'    Dim dSpanne3    As Double
'    Dim dSpanne4    As Double

    NettospanneInEuro = "0"
    
    If cMwst = "V" Then
        dSpanne1 = (dKVP * 100) / (100 + gdMWStV)
    ElseIf cMwst = "E" Then
        dSpanne1 = (dKVP * 100) / (100 + gdMWStE)
    Else
        dSpanne1 = (dKVP * 100) / 100
    End If
    
    dSpanne2 = dSpanne1 - dEK
        
    NettospanneInEuro = Format$(dSpanne2, "###,##0.00")

    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "NettospanneInEuro"
    Fehler.gsFehlertext = "Bei der Berechnung der Nettospanne ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Function
Private Function ermUmsatzGespreislage() As String
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    
    ermUmsatzGespreislage = "0"
    
    sSQL = "Select sum(preis) as maxi"
    sSQL = sSQL & " from Preiskass "
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermUmsatzGespreislage = rsrs!maxi
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
                
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ermUmsatzGespreislage"
    Fehler.gsFehlertext = "In der Preislagenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Function
Private Function ermUmsatzNettoGespreislage() As String
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    
    ermUmsatzNettoGespreislage = "0"
    
    sSQL = "Select sum(npreis) as maxi"
    sSQL = sSQL & " from Preiskass "
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermUmsatzNettoGespreislage = rsrs!maxi
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
                
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ermUmsatzNettoGespreislage"
    Fehler.gsFehlertext = "In der Preislagenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Function
Private Function ermMengeGespreislage() As String
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    
    ermMengeGespreislage = "0"
    
    sSQL = "Select sum(menge) as maxi"
    sSQL = sSQL & " from Preiskass "
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermMengeGespreislage = rsrs!maxi
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
                
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ermMengeGespreislage"
    Fehler.gsFehlertext = "In der Preislagenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Function
Private Function ermErtragGespreislage() As String
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    
    ermErtragGespreislage = "0"
    
    sSQL = "Select sum(rertrag) as maxi"
    sSQL = sSQL & " from Preiskass "
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermErtragGespreislage = rsrs!maxi
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
                
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ermErtragGespreislage"
    Fehler.gsFehlertext = "In der Preislagenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Function
Private Function ermanzpreislagen() As Byte
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    
    ermanzpreislagen = 0
    Set rsrs = gdBase.OpenRecordset("PREISL")
    If Not rsrs.EOF Then
        rsrs.MoveLast
        ermanzpreislagen = rsrs.RecordCount
    End If
    rsrs.Close: Set rsrs = Nothing
        
    
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ermanzpreislagen"
    Fehler.gsFehlertext = "In der Preislagenstatistik ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Function
Public Sub speichernMerkmal(cART As String, cMerk As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If Trim(ZeigeArtmerk(cART)) <> Trim(cMerk) Then
        schreibeProtokollUNITXT "Bei Artikel " & cART & " dieses Merkmal " & cMerk & " eingetragen", "Artikelmerkmal"
    
        sSQL = "Delete from Artmerk where ARTNR = " & Trim(cART)
        gdBase.Execute sSQL, dbFailOnError
        
        sSQL = "Insert into Artmerk (ARTNR,MERK,SENDOK) Values (" & Trim(cART) & ", '" & Trim(cMerk) & "',0) "
        gdBase.Execute sSQL, dbFailOnError
    End If

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "speichernMerkmal"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten  ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub speichernStornof(cART As String, cMerk As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If Trim(ZeigeSTORNOF(cART)) <> Trim(cMerk) Then
        
        sSQL = "Delete from STORNOF where ARTNR = " & Trim(cART)
        gdBase.Execute sSQL, dbFailOnError
        
        sSQL = "Insert into STORNOF (ARTNR,MERK,SENDOK) Values (" & Trim(cART) & ", '" & Trim(cMerk) & "',0) "
        gdBase.Execute sSQL, dbFailOnError
    
    End If

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "speichernStornof"
    Fehler.gsFehlertext = "Im Programmteil Artikel bearbeiten  ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Function ermArtikelSchwundZR(bymonat As Byte, lJahr As Long) As Long
On Error GoTo LOKAL_ERROR

Dim cSQL    As String
Dim rsrs    As Recordset

ermArtikelSchwundZR = 0

cSQL = "select sum(umenge) as maxi from Bestprot "
cSQL = cSQL & " where year(lastdate) =  " & lJahr
cSQL = cSQL & " and month(lastdate) =  " & bymonat
'cSQL = cSQL & " and aenart = 'Bestandskorrektur'"
cSQL = cSQL & " and aengrund = 'Diebstahl'"
cSQL = cSQL & " and umenge < 0 "
cSQL = cSQL & " and umenge > -21 "
Set rsrs = gdBase.OpenRecordset(cSQL)
If Not rsrs.EOF Then
    If Not IsNull(rsrs!maxi) Then
        ermArtikelSchwundZR = rsrs!maxi
    End If
End If
rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
  
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "ermArtikelSchwundZR"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermArtikelSchwundSEKWERTZR(bymonat As Byte, lJahr As Long) As Double
On Error GoTo LOKAL_ERROR

Dim cSQL    As String
Dim rsrs    As Recordset

ermArtikelSchwundSEKWERTZR = 0

cSQL = "select sum(Bestprot.umenge * Artikel.EKPR) as maxi from Bestprot inner join Artikel "
cSQL = cSQL & " on Bestprot.artnr = Artikel.artnr "
cSQL = cSQL & " where year(Bestprot.lastdate) =  " & lJahr
cSQL = cSQL & " and month(Bestprot.lastdate) =  " & bymonat
'cSQL = cSQL & " and Bestprot.aenart = 'Bestandskorrektur'"
cSQL = cSQL & " and Bestprot.aengrund = 'Diebstahl'"
cSQL = cSQL & " and Bestprot.umenge < 0 "
cSQL = cSQL & " and Bestprot.umenge > -21 "
Set rsrs = gdBase.OpenRecordset(cSQL)
If Not rsrs.EOF Then
    If Not IsNull(rsrs!maxi) Then
        ermArtikelSchwundSEKWERTZR = rsrs!maxi
    End If
End If
rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
  
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "ermArtikelSchwundSEKWERTZR"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermArtikelSchwundNettoertragZR(bymonat As Byte, lJahr As Long) As Double
On Error GoTo LOKAL_ERROR

Dim cSQL    As String
Dim rsrs    As Recordset
Dim dWert   As Double

ermArtikelSchwundNettoertragZR = 0

dWert = 0
cSQL = "select sum((Bestprot.umenge * Artikel.kvkpr1)*100/(100 + " & gdMWStV & ")) - sum(Bestprot.umenge * Artikel.EKPR) as maxi from Bestprot inner join Artikel "
cSQL = cSQL & " on Bestprot.artnr = Artikel.artnr "
cSQL = cSQL & " where year(Bestprot.lastdate) =  " & lJahr
cSQL = cSQL & " and month(Bestprot.lastdate) =  " & bymonat
'cSQL = cSQL & " and Bestprot.aenart = 'Bestandskorrektur'"
cSQL = cSQL & " and Bestprot.aengrund = 'Diebstahl'"
cSQL = cSQL & " and Bestprot.umenge < 0 "
cSQL = cSQL & " and Bestprot.umenge > -21 "

cSQL = cSQL & " and ARTIKEL.MWST = 'V' "

Set rsrs = gdBase.OpenRecordset(cSQL)
If Not rsrs.EOF Then
    If Not IsNull(rsrs!maxi) Then
        dWert = rsrs!maxi
        ermArtikelSchwundNettoertragZR = dWert
    End If
End If
rsrs.Close: Set rsrs = Nothing

dWert = 0
cSQL = "select sum((Bestprot.umenge * Artikel.kvkpr1)*100/(100 + " & gdMWStE & ")) - sum(Bestprot.umenge * Artikel.EKPR) as maxi from Bestprot inner join Artikel "
cSQL = cSQL & " on Bestprot.artnr = Artikel.artnr "
cSQL = cSQL & " where year(Bestprot.lastdate) =  " & lJahr
cSQL = cSQL & " and month(Bestprot.lastdate) =  " & bymonat
'cSQL = cSQL & " and Bestprot.aenart = 'Bestandskorrektur'"
cSQL = cSQL & " and Bestprot.aengrund = 'Diebstahl'"
cSQL = cSQL & " and Bestprot.umenge < 0 "
cSQL = cSQL & " and Bestprot.umenge > -21 "

cSQL = cSQL & " and ARTIKEL.MWST = 'E' "

Set rsrs = gdBase.OpenRecordset(cSQL)
If Not rsrs.EOF Then
    If Not IsNull(rsrs!maxi) Then
        dWert = rsrs!maxi
        ermArtikelSchwundNettoertragZR = ermArtikelSchwundNettoertragZR + dWert
    End If
End If
rsrs.Close: Set rsrs = Nothing

dWert = 0
cSQL = "select sum((Bestprot.umenge * Artikel.kvkpr1)*100/(100 + " & gdMWStO & ")) - sum(Bestprot.umenge * Artikel.EKPR) as maxi from Bestprot inner join Artikel "
cSQL = cSQL & " on Bestprot.artnr = Artikel.artnr "
cSQL = cSQL & " where year(Bestprot.lastdate) =  " & lJahr
cSQL = cSQL & " and month(Bestprot.lastdate) =  " & bymonat
'cSQL = cSQL & " and Bestprot.aenart = 'Bestandskorrektur'"
cSQL = cSQL & " and Bestprot.aengrund = 'Diebstahl'"
cSQL = cSQL & " and Bestprot.umenge < 0 "
cSQL = cSQL & " and Bestprot.umenge > -21 "

cSQL = cSQL & " and ARTIKEL.MWST = 'O' "

Set rsrs = gdBase.OpenRecordset(cSQL)
If Not rsrs.EOF Then
    If Not IsNull(rsrs!maxi) Then
        dWert = rsrs!maxi
        ermArtikelSchwundNettoertragZR = ermArtikelSchwundNettoertragZR + dWert
    End If
End If
rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
  
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "ermArtikelSchwundNettoertragZR"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermNeuKundenZR(bymonat As Byte, lJahr As Long, iFil As Integer) As Long
On Error GoTo LOKAL_ERROR

Dim cSQL    As String
Dim rsrs    As Recordset

ermNeuKundenZR = 0

cSQL = "select count(kundnr)as maxi from Kunden "
cSQL = cSQL & " where year(angelegt) =  " & lJahr
cSQL = cSQL & " and month(angelegt) =  " & bymonat
cSQL = cSQL & " and Filialnr =  " & iFil
Set rsrs = gdBase.OpenRecordset(cSQL)
If Not rsrs.EOF Then
    If Not IsNull(rsrs!maxi) Then
        ermNeuKundenZR = rsrs!maxi
    End If
End If
rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
  
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "ermNeuKundenZR"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function BonAnzahl(cKundnr As String, iTage As Integer, sTab As String) As Long
    On Error GoTo LOKAL_ERROR

    BonAnzahl = 0
    Dim rsrs        As Recordset
    Dim sSQL        As String

    loeschNEW "zeitzone", gdBase
    
    sSQL = "select adate, adate as belegnr into zeitzone  "
    sSQL = sSQL & " from  " & sTab
    sSQL = sSQL & "  where Kundnr = " & cKundnr
    If iTage > 0 Then
        sSQL = sSQL & " and adate > clng(datevalue(now) - " & iTage & ")"
    End If
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "zbonanz", gdBase
    
    sSQL = " select distinct(belegnr) into zbonanz from zeitzone "
    sSQL = sSQL & " group by belegnr "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "select * from zbonanz "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        
        rsrs.MoveLast
        BonAnzahl = rsrs.RecordCount
        
    End If
    rsrs.Close
    
    
    loeschNEW "zbonanz", gdBase
    loeschNEW "zeitzone", gdBase
    
Exit Function
LOKAL_ERROR:
    If err.Number = 3010 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "BonAnzahl"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Function

Public Function sumInsgesamt(cKundnr As String, iTage As Integer, sTab As String) As Double
    On Error GoTo LOKAL_ERROR
    
    If cKundnr = "" Then
        Exit Function
    End If
    
    If IsNumeric(cKundnr) = False Then
        Exit Function
    End If

    sumInsgesamt = 0
    Dim rsrs        As Recordset
    Dim sSQL        As String

    sSQL = "select sum(preis) as sumpreis  "
    sSQL = sSQL & " from  " & sTab
    sSQL = sSQL & "  where Kundnr = " & cKundnr
    
    If iTage > 0 Then
        sSQL = sSQL & " and adate > clng(datevalue(now) - " & iTage & ")"
    End If
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!sumpreis) Then
            sumInsgesamt = rsrs!sumpreis
        End If
    End If
    rsrs.Close
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "sumInsgesamt"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Sub AllesinKUNDAZE(cKund As String, bAlles As Boolean)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    If cKund = "" Then
        Exit Sub
    End If
    
    'die folgende Zeilie wurde von Odayy auskommentiert
    
    '        If NewTableSuchenDBKombi("KUNDAZE", gdBase) Then
    '
    '            If DatendrinSQL("select * from KUNDAZE where kundnr = " & cKund, gdBase) Then
    '                Exit Sub
    '            End If
    '
    '        End If
   
    loeschNEW "KUNDAZE", gdBase
    CreateTable "KUNDAZE", gdBase
    
''    If Left(gFirma.FirmaName, 11) = "CONTHERAPIA" Then
''        'ab 21.02.2014 rausnehmen
''        sSQL = "Insert into KUNDAZE select bezeich,artnr,menge,adate,Filiale,preis,bediener as bednr,kundnr from kassjour "
''    Else
''        sSQL = "Insert into KUNDAZE select artnr,menge,adate,Filiale,preis,bediener as bednr,kundnr from kassjour "
''    End If

    'vielleicht glob Variable siehe CONTHERAPIA
    sSQL = "Insert into KUNDAZE select bezeich,artnr,menge,adate,Filiale,preis,bediener as bednr,kundnr from kassjour "
    
    
    
    sSQL = sSQL & " where kundnr = " & cKund
    
    If bAlles = False Then
        If Month(DateValue(Now)) <= 7 Then
            sSQL = sSQL & " and year(adate) >= year(datevalue(now))-1"
        Else
            sSQL = sSQL & " and year(adate) = year(datevalue(now))"
        End If
    End If
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into KUNDAZE select artnr,menge,adate,Filiale,preis,bednr,kundnr from kundkass "
    sSQL = sSQL & " where kundnr = " & cKund
    If bAlles = False Then
        If Month(DateValue(Now)) <= 7 Then
            sSQL = sSQL & " and year(adate) >= year(datevalue(now))-1"
        Else
            sSQL = sSQL & " and year(adate) = year(datevalue(now))"
        End If
    End If
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into KUNDAZE select artnr,menge,adate,Filiale,preis,bediener as bednr,kundnr from kollverk "
    sSQL = sSQL & " where kundnr = " & cKund
    If bAlles = False Then
        If Month(DateValue(Now)) <= 7 Then
            sSQL = sSQL & " and year(adate) >= year(datevalue(now))-1"
        Else
            sSQL = sSQL & " and year(adate) = year(datevalue(now))"
        End If
    End If
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update KUNDAZE inner join artikel on kundaze.artnr = artikel.artnr"
    sSQL = sSQL & " set KUNDAZE.PGN = artikel.PGN "
    sSQL = sSQL & " , KUNDAZE.AGN = artikel.AGN "
    sSQL = sSQL & " , KUNDAZE.LPZ = artikel.LPZ "
    sSQL = sSQL & " , KUNDAZE.LINR = artikel.LINR "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update KUNDAZE inner join linbez on KUNDAZE.Linr = linbez.Linr and KUNDAZE.LPZ = linbez.LPZ "
    sSQL = sSQL & " set KUNDAZE.Marke = linbez.marke "
    gdBase.Execute sSQL, dbFailOnError
    
     'vielleicht glob Variable siehe CONTHERAPIA
'    If Left(gFirma.FirmaName, 11) <> "CONTHERAPIA" Then
'        'ab 21.02.2014 rausnehmen
'        sSQL = "Update KUNDAZE inner join artikel on kundaze.artnr = artikel.artnr"
'        sSQL = sSQL & " set KUNDAZE.Bezeich = artikel.bezeich "
'        sSQL = sSQL & " where  artikel.artnr <> 666666"
'        gdBase.Execute sSQL, dbFailOnError
'    End If

    sSQL = "Update KUNDAZE set Bezeich = '' where bezeich = '0' "
    gdBase.Execute sSQL, dbFailOnError

    sSQL = "Update KUNDAZE set Bezeich = '' where bezeich is null "
    gdBase.Execute sSQL, dbFailOnError

    sSQL = "Update KUNDAZE inner join artikel on kundaze.artnr = artikel.artnr"
    sSQL = sSQL & " set KUNDAZE.Bezeich = artikel.bezeich "
    sSQL = sSQL & " where  artikel.artnr <> 666666"
    sSQL = sSQL & " and  KUNDAZE.Bezeich = ''"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update KUNDAZE "
    sSQL = sSQL & " set KUNDAZE.Bezeich = 'Gutschein' "
    sSQL = sSQL & " where  KUNDAZE.artnr = 666666"
    gdBase.Execute sSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "AllesinKUNDAZE"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub
Public Sub AllesinKUNDAZEeinzel(cKund As String, bAlles As Boolean)
    On Error GoTo LOKAL_ERROR
    
    If cKund = "" Then
        Exit Sub
    End If
    
    Dim sSQL As String
    
    sSQL = "Delete from  KUNDA" & srechnertab
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into KUNDA" & srechnertab & " select menge,adate,preis ,kundnr from kassjour "
    sSQL = sSQL & " where kundnr = " & cKund
    
    If bAlles = False Then
        If Month(DateValue(Now)) <= 7 Then
            sSQL = sSQL & " and year(adate) >= year(datevalue(now))-1"
        Else
            sSQL = sSQL & " and year(adate) = year(datevalue(now))"
        End If
    End If
    
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into KUNDA" & srechnertab & " select menge,adate,preis,kundnr from kundkass "
    sSQL = sSQL & " where kundnr = " & cKund
    
    If bAlles = False Then
        If Month(DateValue(Now)) <= 7 Then
            sSQL = sSQL & " and year(adate) >= year(datevalue(now))-1"
        Else
            sSQL = sSQL & " and year(adate) = year(datevalue(now))"
        End If
    End If
    
    gdBase.Execute sSQL, dbFailOnError
    
    
    sSQL = "Insert into KUNDA" & srechnertab & " select menge,adate,preis ,kundnr from kollverk "
    sSQL = sSQL & " where kundnr = " & cKund
    If bAlles = False Then
        If Month(DateValue(Now)) <= 7 Then
            sSQL = sSQL & " and year(adate) >= year(datevalue(now))-1"
        Else
            sSQL = sSQL & " and year(adate) = year(datevalue(now))"
        End If
    End If
    gdBase.Execute sSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "AllesinKUNDAZEeinzel"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub
Public Function wievieleSterneQuick(lVerkVor365 As Long, dUmsku365 As Double, dBonschni365 As Double) As Integer
    On Error GoTo LOKAL_ERROR
    
    wievieleSterneQuick = 0
    
    glAnzvkKU365 = lVerkVor365
    gdumsgesKU365 = dUmsku365
    gdEuroproBonKU365 = dBonschni365
    
    
    '1 Stern wenn umsatz getopt wird
    If gdumsgesKU365 > gdUmsatzMittelproKunde Then
        wievieleSterneQuick = wievieleSterneQuick + 1
    End If
    
    '2 Stern wenn Euro pro Bon getopt wird
    If gdEuroproBonKU365 > gdUmsatzproKundeDurchschnitt Then
        wievieleSterneQuick = wievieleSterneQuick + 1
    End If
    
    '3 Stern wenn Anzahl der Verkaufsvorgänge getopt wird
    If glAnzvkKU365 > gdKaufvorgänge Then
        wievieleSterneQuick = wievieleSterneQuick + 1
    End If
    
    '4 Stern wenn doppelte Durchschnittsumsatz getopt wird
    If gdEuroproBonKU365 > gdUmsatzproKundeDurchschnitt * 2 Then
        wievieleSterneQuick = wievieleSterneQuick + 1
    End If
    
    '5 Stern wenn doppelte DurchschnittproBon getopt wird
    If gdumsgesKU365 > gdUmsatzMittelproKunde * 2 Then
        wievieleSterneQuick = wievieleSterneQuick + 1
    End If
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "wievieleSterneQuick"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Function
Public Function wievieleSterneAltQuick(lVerkVor As Long, dUmsku As Double, dBonschni As Double) As Integer
    On Error GoTo LOKAL_ERROR
    
    wievieleSterneAltQuick = 0
    
    glAnzvkKU = lVerkVor
    gdumsgesKU = dUmsku
    gdEuroproBonKU = dBonschni

    '1 Stern wenn umsatz getopt wird
    If gdumsgesKU > gdUmsatzMittelproKunde Then
        wievieleSterneAltQuick = wievieleSterneAltQuick + 1
    End If
    
    '2 Stern wenn Euro pro Bon getopt wird
    If gdEuroproBonKU > gdUmsatzproKundeDurchschnitt Then
        wievieleSterneAltQuick = wievieleSterneAltQuick + 1
    End If
    
    '3 Stern wenn Anzahl der Verkaufsvorgänge getopt wird
    If glAnzvkKU > gdKaufvorgänge Then
        wievieleSterneAltQuick = wievieleSterneAltQuick + 1
    End If
    
    '4 Stern wenn doppelte Durchschnittsumsatz getopt wird
    If gdEuroproBonKU > gdUmsatzproKundeDurchschnitt * 2 Then
        wievieleSterneAltQuick = wievieleSterneAltQuick + 1
    End If
    
    '5 Stern wenn doppelte DurchschnittproBon getopt wird
    If gdumsgesKU > gdUmsatzMittelproKunde * 2 Then
        wievieleSterneAltQuick = wievieleSterneAltQuick + 1
    End If
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "wievieleSterneAltQuick"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Function
Public Function wievieleSterne(cKundnr As String, sTab As String, iBasisTage As Integer) As Integer
    On Error GoTo LOKAL_ERROR
    
    wievieleSterne = 0
    
    glAnzvkKU = 0
    gdumsgesKU = 0
    gdEuroproBonKU = 0
    
    glAnzvkKU365 = 0
    gdumsgesKU365 = 0
    gdEuroproBonKU365 = 0
    
    glAnzvkKU365 = Format(BonAnzahl(cKundnr, iBasisTage, sTab), "####0")
    gdumsgesKU365 = Format(sumInsgesamt(cKundnr, iBasisTage, sTab), "####0.00")
    If glAnzvkKU365 > 0 Then
        gdEuroproBonKU365 = gdumsgesKU365 / glAnzvkKU365
    Else
        gdEuroproBonKU365 = 0
    End If
    
    glAnzvkKU = Format(BonAnzahl(cKundnr, 0, sTab), "####0")
    gdumsgesKU = Format(sumInsgesamt(cKundnr, 0, sTab), "####0.00")
    If glAnzvkKU > 0 Then
        gdEuroproBonKU = gdumsgesKU / glAnzvkKU
    Else
        gdEuroproBonKU = 0
    End If
    
    
    '1 Stern wenn umsatz getopt wird
    If gdumsgesKU365 > gdUmsatzMittelproKunde Then
        wievieleSterne = wievieleSterne + 1
    End If
    
    '2 Stern wenn Euro pro Bon getopt wird
    If gdEuroproBonKU365 > gdUmsatzproKundeDurchschnitt Then
        wievieleSterne = wievieleSterne + 1
    End If
    
    '3 Stern wenn Anzahl der Verkaufsvorgänge getopt wird
    If glAnzvkKU365 > gdKaufvorgänge Then
        wievieleSterne = wievieleSterne + 1
    End If
    
    '4 Stern wenn doppelte Durchschnittsumsatz getopt wird
    If gdEuroproBonKU365 > gdUmsatzproKundeDurchschnitt * 2 Then
        wievieleSterne = wievieleSterne + 1
    End If
    
    '5 Stern wenn doppelte DurchschnittproBon getopt wird
    If gdumsgesKU365 > gdUmsatzMittelproKunde * 2 Then
        wievieleSterne = wievieleSterne + 1
    End If
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "wievieleSterne"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Function
Public Function wievieleSterneAlt() As Integer
    On Error GoTo LOKAL_ERROR
    
    wievieleSterneAlt = 0

    '1 Stern wenn umsatz getopt wird
    If gdumsgesKU > gdUmsatzMittelproKunde Then
        wievieleSterneAlt = wievieleSterneAlt + 1
    End If
    
    '2 Stern wenn Euro pro Bon getopt wird
    If gdEuroproBonKU > gdUmsatzproKundeDurchschnitt Then
        wievieleSterneAlt = wievieleSterneAlt + 1
    End If
    
    '3 Stern wenn Anzahl der Verkaufsvorgänge getopt wird
    If glAnzvkKU > gdKaufvorgänge Then
        wievieleSterneAlt = wievieleSterneAlt + 1
    End If
    
    '4 Stern wenn doppelte Durchschnittsumsatz getopt wird
    If gdEuroproBonKU > gdUmsatzproKundeDurchschnitt * 2 Then
        wievieleSterneAlt = wievieleSterneAlt + 1
    End If
    
    '5 Stern wenn doppelte DurchschnittproBon getopt wird
    If gdumsgesKU > gdUmsatzMittelproKunde * 2 Then
        wievieleSterneAlt = wievieleSterneAlt + 1
    End If
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "wievieleSterneAlt"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Function
Public Sub Zeigdiesterne(picx1 As PictureBox, picx2 As PictureBox, picx3 As PictureBox, picx4 As PictureBox, picx5 As PictureBox, iSterne As Integer, isterneAlt As Integer)
    On Error GoTo LOKAL_ERROR
    
    
    If Modul6.FindFile(App.Path, "\white.gif") = False Then
        picx1.Visible = False
        picx2.Visible = False
        picx3.Visible = False
        picx4.Visible = False
        picx5.Visible = False
        Exit Sub
    End If
    
    If Modul6.FindFile(App.Path, "\gelb.gif") = False Then
        picx1.Visible = False
        picx2.Visible = False
        picx3.Visible = False
        picx4.Visible = False
        picx5.Visible = False
        Exit Sub
    End If
    
    If Modul6.FindFile(App.Path, "\black.gif") = False Then
        picx1.Visible = False
        picx2.Visible = False
        picx3.Visible = False
        picx4.Visible = False
        picx5.Visible = False
        Exit Sub
    End If
    
    picx1.Width = 285
    picx2.Width = 285
    picx3.Width = 285
    picx4.Width = 285
    picx5.Width = 285
    picx2.Left = picx1.Left + picx1.Width
    picx3.Left = picx2.Left + picx2.Width
    picx4.Left = picx3.Left + picx3.Width
    picx5.Left = picx4.Left + picx4.Width
    
    
    
    
    Select Case iSterne
        Case 0
        
            If isterneAlt > iSterne Then
                Select Case isterneAlt
                    Case 1
                        picx1.Picture = LoadPicture(App.Path & "\black.gif")
                        picx2.Picture = LoadPicture(App.Path & "\white.gif")
                        picx3.Picture = LoadPicture(App.Path & "\white.gif")
                        picx4.Picture = LoadPicture(App.Path & "\white.gif")
                        picx5.Picture = LoadPicture(App.Path & "\white.gif")
                    Case 2
                        picx1.Picture = LoadPicture(App.Path & "\black.gif")
                        picx2.Picture = LoadPicture(App.Path & "\black.gif")
                        picx3.Picture = LoadPicture(App.Path & "\white.gif")
                        picx4.Picture = LoadPicture(App.Path & "\white.gif")
                        picx5.Picture = LoadPicture(App.Path & "\white.gif")
                    Case 3
                        picx1.Picture = LoadPicture(App.Path & "\black.gif")
                        picx2.Picture = LoadPicture(App.Path & "\black.gif")
                        picx3.Picture = LoadPicture(App.Path & "\black.gif")
                        picx4.Picture = LoadPicture(App.Path & "\white.gif")
                        picx5.Picture = LoadPicture(App.Path & "\white.gif")
                    Case 4
                        picx1.Picture = LoadPicture(App.Path & "\black.gif")
                        picx2.Picture = LoadPicture(App.Path & "\black.gif")
                        picx3.Picture = LoadPicture(App.Path & "\black.gif")
                        picx4.Picture = LoadPicture(App.Path & "\black.gif")
                        picx5.Picture = LoadPicture(App.Path & "\white.gif")
                    Case 5
                        picx1.Picture = LoadPicture(App.Path & "\black.gif")
                        picx2.Picture = LoadPicture(App.Path & "\black.gif")
                        picx3.Picture = LoadPicture(App.Path & "\black.gif")
                        picx4.Picture = LoadPicture(App.Path & "\black.gif")
                        picx5.Picture = LoadPicture(App.Path & "\black.gif")
                End Select
            
            Else
                picx1.Picture = LoadPicture(App.Path & "\white.gif")
                picx2.Picture = LoadPicture(App.Path & "\white.gif")
                picx3.Picture = LoadPicture(App.Path & "\white.gif")
                picx4.Picture = LoadPicture(App.Path & "\white.gif")
                picx5.Picture = LoadPicture(App.Path & "\white.gif")
            End If
            
        Case 1
            If isterneAlt > iSterne Then
            
                Select Case isterneAlt
                    
                    Case 2
                        picx1.Picture = LoadPicture(App.Path & "\gelb.gif")
                        picx2.Picture = LoadPicture(App.Path & "\black.gif")
                        picx3.Picture = LoadPicture(App.Path & "\white.gif")
                        picx4.Picture = LoadPicture(App.Path & "\white.gif")
                        picx5.Picture = LoadPicture(App.Path & "\white.gif")
                    Case 3
                        picx1.Picture = LoadPicture(App.Path & "\gelb.gif")
                        picx2.Picture = LoadPicture(App.Path & "\black.gif")
                        picx3.Picture = LoadPicture(App.Path & "\black.gif")
                        picx4.Picture = LoadPicture(App.Path & "\white.gif")
                        picx5.Picture = LoadPicture(App.Path & "\white.gif")
                    Case 4
                        picx1.Picture = LoadPicture(App.Path & "\gelb.gif")
                        picx2.Picture = LoadPicture(App.Path & "\black.gif")
                        picx3.Picture = LoadPicture(App.Path & "\black.gif")
                        picx4.Picture = LoadPicture(App.Path & "\black.gif")
                        picx5.Picture = LoadPicture(App.Path & "\white.gif")
                    Case 5
                        picx1.Picture = LoadPicture(App.Path & "\gelb.gif")
                        picx2.Picture = LoadPicture(App.Path & "\black.gif")
                        picx3.Picture = LoadPicture(App.Path & "\black.gif")
                        picx4.Picture = LoadPicture(App.Path & "\black.gif")
                        picx5.Picture = LoadPicture(App.Path & "\black.gif")
                End Select
            
            Else
                picx1.Picture = LoadPicture(App.Path & "\gelb.gif")
                picx2.Picture = LoadPicture(App.Path & "\white.gif")
                picx3.Picture = LoadPicture(App.Path & "\white.gif")
                picx4.Picture = LoadPicture(App.Path & "\white.gif")
                picx5.Picture = LoadPicture(App.Path & "\white.gif")
            End If
            
        
        Case 2
            If isterneAlt > iSterne Then
                Select Case isterneAlt
                    
                    Case 3
                        picx1.Picture = LoadPicture(App.Path & "\gelb.gif")
                        picx2.Picture = LoadPicture(App.Path & "\gelb.gif")
                        picx3.Picture = LoadPicture(App.Path & "\black.gif")
                        picx4.Picture = LoadPicture(App.Path & "\white.gif")
                        picx5.Picture = LoadPicture(App.Path & "\white.gif")
                    Case 4
                        picx1.Picture = LoadPicture(App.Path & "\gelb.gif")
                        picx2.Picture = LoadPicture(App.Path & "\gelb.gif")
                        picx3.Picture = LoadPicture(App.Path & "\black.gif")
                        picx4.Picture = LoadPicture(App.Path & "\black.gif")
                        picx5.Picture = LoadPicture(App.Path & "\white.gif")
                    Case 5
                        picx1.Picture = LoadPicture(App.Path & "\gelb.gif")
                        picx2.Picture = LoadPicture(App.Path & "\gelb.gif")
                        picx3.Picture = LoadPicture(App.Path & "\black.gif")
                        picx4.Picture = LoadPicture(App.Path & "\black.gif")
                        picx5.Picture = LoadPicture(App.Path & "\black.gif")
                End Select
            
            Else
                picx1.Picture = LoadPicture(App.Path & "\gelb.gif")
                picx2.Picture = LoadPicture(App.Path & "\gelb.gif")
                picx3.Picture = LoadPicture(App.Path & "\white.gif")
                picx4.Picture = LoadPicture(App.Path & "\white.gif")
                picx5.Picture = LoadPicture(App.Path & "\white.gif")
            End If
            
        Case 3
            If isterneAlt > iSterne Then
                Select Case isterneAlt
                    Case 4
                        picx1.Picture = LoadPicture(App.Path & "\gelb.gif")
                        picx2.Picture = LoadPicture(App.Path & "\gelb.gif")
                        picx3.Picture = LoadPicture(App.Path & "\gelb.gif")
                        picx4.Picture = LoadPicture(App.Path & "\black.gif")
                        picx5.Picture = LoadPicture(App.Path & "\white.gif")
                    Case 5
                        picx1.Picture = LoadPicture(App.Path & "\gelb.gif")
                        picx2.Picture = LoadPicture(App.Path & "\gelb.gif")
                        picx3.Picture = LoadPicture(App.Path & "\gelb.gif")
                        picx4.Picture = LoadPicture(App.Path & "\black.gif")
                        picx5.Picture = LoadPicture(App.Path & "\black.gif")
                End Select
            Else
                picx1.Picture = LoadPicture(App.Path & "\gelb.gif")
                picx2.Picture = LoadPicture(App.Path & "\gelb.gif")
                picx3.Picture = LoadPicture(App.Path & "\gelb.gif")
                picx4.Picture = LoadPicture(App.Path & "\white.gif")
                picx5.Picture = LoadPicture(App.Path & "\white.gif")
            End If
            
        Case 4
            If isterneAlt > iSterne Then
                Select Case isterneAlt
                    Case 5
                        picx1.Picture = LoadPicture(App.Path & "\gelb.gif")
                        picx2.Picture = LoadPicture(App.Path & "\gelb.gif")
                        picx3.Picture = LoadPicture(App.Path & "\gelb.gif")
                        picx4.Picture = LoadPicture(App.Path & "\gelb.gif")
                        picx5.Picture = LoadPicture(App.Path & "\black.gif")
                End Select
            Else
                picx1.Picture = LoadPicture(App.Path & "\gelb.gif")
                picx2.Picture = LoadPicture(App.Path & "\gelb.gif")
                picx3.Picture = LoadPicture(App.Path & "\gelb.gif")
                picx4.Picture = LoadPicture(App.Path & "\gelb.gif")
                picx5.Picture = LoadPicture(App.Path & "\white.gif")
            End If
            
        Case 5
            If isterneAlt > iSterne Then
            
            Else
                picx1.Picture = LoadPicture(App.Path & "\gelb.gif")
                picx2.Picture = LoadPicture(App.Path & "\gelb.gif")
                picx3.Picture = LoadPicture(App.Path & "\gelb.gif")
                picx4.Picture = LoadPicture(App.Path & "\gelb.gif")
                picx5.Picture = LoadPicture(App.Path & "\gelb.gif")
            End If
            
    End Select
    
    If gdKaufvorgänge > 0 Then
    
        picx1.Visible = True
        picx2.Visible = True
        picx3.Visible = True
        picx4.Visible = True
        picx5.Visible = True
    Else
        picx1.Visible = False
        picx2.Visible = False
        picx3.Visible = False
        picx4.Visible = False
        picx5.Visible = False
    End If
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "Zeigdiesterne"
    Fehler.gsFehlertext = "Im Programmteil Kunden Verkauf ist ein Fehler aufgetreten. "
    
    Fehlermeldung1
End Sub
Public Function ermNeinVKZR(bymonat As Byte, lJahr As Long) As Long
On Error GoTo LOKAL_ERROR

Dim cSQL    As String
Dim rsrs    As Recordset

ermNeinVKZR = 0

cSQL = "select sum(menge)as maxi from NEINVK "
cSQL = cSQL & " where year(adate) =  " & lJahr
cSQL = cSQL & " and month(adate) =  " & bymonat
Set rsrs = gdBase.OpenRecordset(cSQL)
If Not rsrs.EOF Then
    If Not IsNull(rsrs!maxi) Then
        ermNeinVKZR = rsrs!maxi
    End If
End If
rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
  
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "ermNeinVKZR"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function

