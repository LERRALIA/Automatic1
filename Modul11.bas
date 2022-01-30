Attribute VB_Name = "Modul11"
Public Function Datenbankreparatur(cdabapfad As String, cDabaName As String, sPass As String, labelx As Label, labelZ As Label) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim Task$, hProcess&, result&
    Dim lRet        As Long
    Dim lfail       As Long
    Dim sQuell      As String
    Dim sZiel       As String
    Dim ctmp        As String
    Dim dbTest      As Database
    Dim cdatei      As String
    Dim lWert       As Long
    
    Datenbankreparatur = False
    
    Screen.MousePointer = 11
    
    frmWKL00.Timer1.Enabled = False

    lWert = DateValue(Now)
    ctmp = Format$(lWert, "DD.MM")

    cdatei = "A" & ctmp & Format$(TimeValue(Now), "HH:MM:SS")
    cdatei = SwapStr(cdatei, ".", "")
    cdatei = SwapStr(cdatei, ":", "")
    cdatei = cdatei & ".mdb"
    
    '1. Kopiere
    sZiel = cdabapfad & cdatei '"ZUREP.mdb"
    sQuell = cdabapfad & cDabaName
    
    labelx.Visible = True
    labelZ.Visible = True
    anzeige "black", TimeValue(Now) & " : Datenbank wird kopiert, nicht abbrechen!!!", labelx
    anzeige "black", TimeValue(Now) & " : Start der Datenbankreparatur: ", labelZ
    
    lRet = CopyFile(sQuell, sZiel, lfail)
    If lRet = 0 Then
        Screen.MousePointer = 0
        MsgBox "Konnte " & sQuell & " nicht kopieren! Datenträger voll?", vbCritical, "Winkiss Hinweis:"
        anzeige "black", TimeValue(Now) & " :  Datenbank wird kopiert, Abbruch erfolgt", labelx
        Exit Function
    End If
    
    Screen.MousePointer = 11
    
    anzeige "black", "temporäre Datenbank wird gelöscht, nicht abbrechen!!!", labelx
    Kill sQuell
    
    Screen.MousePointer = 11
    anzeige "black", TimeValue(Now) & " : Datenbank wird repariert, nicht abbrechen!!!", labelx
    If sPass <> "" Then
        Task = Shell(cdabapfad & "JETCOMP.exe -src:" & cdabapfad & cdatei & " -dest:" & cdabapfad & cDabaName & " -w" & sPass)
'        Task = Shell(cDabapfad & "JETCOMP.exe -src:" & cDabapfad & "ZUREP.mdb -dest:" & cDabapfad & cDabaName & " -w" & sPass)
    Else
    
    End If
    Screen.MousePointer = 11
    hProcess = OpenProcess(SYNCHRONIZE, False, Task)
    result = WaitForSingleObject(hProcess, INFINITE)
    result = CloseHandle(hProcess)
    
    anzeige "black", TimeValue(Now) & " : Datenbanktest, nicht abbrechen!!!", labelx
    Set dbTest = OpenDatabase(cdabapfad & cDabaName, False, False, "MS Access;PWD=" & gsPasswort)
    dbTest.Close
    
    Datenbankreparatur = True
    anzeige "black", TimeValue(Now) & " : Datenbank erfolgreich repariert", labelx
    
    If FileExists(App.Path & "\NoTimer.cfg") Then
        frmWKL00.Timer1.Enabled = False
    Else
        frmWKL00.Timer1.Enabled = True
    End If
    Screen.MousePointer = 0
    
    Exit Function
LOKAL_ERROR:

    If err.Number = 53 Or err.Number = 70 Then
        ctmp = "Alle anderen Winkiss/Kassenprogramme an allen anderen Rechnern schließen!. " & vbCrLf & vbCrLf
        ctmp = ctmp & "Die Datei: '" & sQuell & "' ist noch im Zugriff." & vbCrLf & vbCrLf
        ctmp = ctmp & "Winkiss wird beendet."
        MsgBox ctmp, vbCritical, "Winkiss Hinweis:"
        End
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul11"
        Fehler.gsFunktion = "Datenbankreparatur"
        Fehler.gsFehlertext = "Es trat ein Fehler auf. "
        
        Fehlermeldung1
    End If
End Function
Public Function checkspalte(db As Database, sSpalt As String) As Long
On Error GoTo LOKAL_ERROR

checkspalte = 0

Dim lAnzTable       As Long
Dim lcount          As Long
Dim sTabname        As String
Dim sFname          As String

Dim j As Integer

db.TableDefs.Refresh
lAnzTable = db.TableDefs.Count
For lcount = 0 To lAnzTable - 1
    sTabname = db.TableDefs(lcount).name
    For j = 0 To db.TableDefs(lcount).Fields.Count - 1
        sFname = UCase(db.TableDefs(lcount).Fields(j).name)
        If sFname = UCase(sSpalt) Then
        
            Select Case UCase(sSpalt)
                Case "BEZEICH"
                    checkBezeich sTabname, db
            End Select
        End If
    Next j
Next lcount

Exit Function
LOKAL_ERROR:
   
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul11"
    Fehler.gsFunktion = "checkspalte"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    Resume Next
End Function
Public Function LoeseMarkenInArtnr1(cKrit As String) As Boolean
On Error GoTo LOKAL_ERROR

    Dim sSQL    As String
    Dim rsrs    As Recordset
    Dim lLinr   As Long
    Dim lLpz    As Long
    
    Screen.MousePointer = 11
    
    LoeseMarkenInArtnr1 = False
    
    sSQL = "Delete from MA" & srechnertab
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Select LPZ,LINR from LINBEZ where Marke like '" & cKrit & "*' "
    sSQL = sSQL & " and not Linr is null "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            If Not IsNull(rsrs!LPZ) Then
                If Not IsNull(rsrs!linr) Then
                
                    lLinr = rsrs!linr
                    lLpz = rsrs!LPZ
                     
                    sSQL = " Insert into MA" & srechnertab
                    sSQL = sSQL & " select a.artnr from artikel a , artlief b  "
                    
                    sSQL = sSQL & " where a.LPZ = " & lLpz & " and b.LINR = " & lLinr
                    sSQL = sSQL & " and  a.artnr =  b.artnr "
                    gdBase.Execute sSQL, dbFailOnError
                
                
'                    sSQL = "Insert into MA" & srechnertab
'                    sSQL = sSQL & " Select artnr from Artikel where"
'                    sSQL = sSQL & " LINR = " & rsrs!linr
'                    sSQL = sSQL & " and LPZ = " & rsrs!LPZ
'                    gdBase.Execute sSQL, dbFailOnError
                End If
            End If
        
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    If Datendrin("MA" & srechnertab, gdBase) Then
        LoeseMarkenInArtnr1 = True
    End If
    
    Screen.MousePointer = 0
    
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "LoeseMarkenInArtnr1"
    Fehler.gsFehlertext = "Im Programmteil Artikelsuche ist ein Fehler aufgetreten. "
    Fehlermeldung1
End Function
Public Function checkBezeich(stabn As String, db As Database) As Long
On Error GoTo LOKAL_ERROR

checkBezeich = 0



Dim rs As Recordset
Dim sPBezeich As String

Set rs = db.OpenRecordset(stabn)

If Not rs.EOF Then
    Do While Not rs.EOF
    
    
    
    
    If Not IsNull(rs!BEZEICH) Then
        sPBezeich = CStr(rs!BEZEICH)
'        sPBezeich = Left(sPBezeich, 35)
'
        sPBezeich = SwapStr(sPBezeich, "„", " ")
        sPBezeich = SwapStr(sPBezeich, "]", " ")
        sPBezeich = SwapStr(sPBezeich, "[", " ")

        sPBezeich = SwapStr(sPBezeich, "}", " ")
        sPBezeich = SwapStr(sPBezeich, "{", " ")

        sPBezeich = SwapStr(sPBezeich, ".", " ")
        sPBezeich = SwapStr(sPBezeich, "'", " ")
        sPBezeich = SwapStr(sPBezeich, "!", " ")
        sPBezeich = SwapStr(sPBezeich, "-", " ")

        sPBezeich = SwapStr(sPBezeich, "_", " ")
        sPBezeich = SwapStr(sPBezeich, "*", " ")
        
        
        
        
        
        
        
    Else
        sPBezeich = ""
    
    End If
    
    rs.Edit
    rs!BEZEICH = sPBezeich
    rs.Update
    
    
    
    rs.MoveNext
    Loop

End If
rs.Close: Set rs = Nothing

Exit Function
LOKAL_ERROR:
    If err.Number = 3163 Then
        schreibeProtokollBEZ ("Bez: aus " & stabn & ": " & sPBezeich & " wurde entfernt")
        sPBezeich = ""
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul7"
        Fehler.gsFunktion = "checkBezeich"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Function
Public Function CompactmyDaba(sPfad As String, sDB As String, db As Database, lab As Label, txtStatus As TextBox, labglo As Label, bOhneAnsicht As Boolean) As Boolean
On Error GoTo LOKAL_ERROR

    CompactmyDaba = False
    
    'Tabellen auflisten Anzahl Datensätze ermitteln
    Dim lMax        As Long
    Dim lgMax       As Long
    Dim lTabMax     As Long
    Dim lcount      As Long
    Dim lAnzTable   As Long
    Dim name        As String
    Dim inname      As String
    Dim sSQL        As String
    Dim j           As Long
    Dim dbTEMPCOMP  As Database
    Dim dbCOMP      As Database
    Dim dbEND       As Database
    Dim oldpath     As String
    Dim newpath     As String
    Dim lfail       As Long
    Dim lRet        As Long
    Dim lHeute      As Long
    Dim lGestern    As Long
    Dim cPfad       As String
    Dim cPfad2       As String
    
    cPfad2 = gcDBPfad
    If Right(cPfad2, 1) <> "\" Then
        cPfad2 = cPfad2 & "\"
    End If
    
    Kill cPfad2 & "KISSDATA1.MDB"
    
    cPfad = gcDBPfad
    If Right$(cPfad, 1) = "\" Then
        cPfad = Left(cPfad, Len(cPfad) - 1)
    End If
    
    If IsAktionZulaessig("Komprimierung") = False Then
        Exit Function
    End If
    
    Screen.MousePointer = 11
    
    schreibeProtokollDaba ("Komprimierung gestartet")
    
    '*****Daba erst sichern
    
    db_Sichern cPfad2, "kissdata.mdb", lab, txtStatus, labglo
    DoEvents

    Screen.MousePointer = 11
    
    '*****Daba erst sichern Ende
    
    'neu auch mal reindizieren 05.02.2013
    db_Reindizieren gdBase, lab, txtStatus, labglo
    'neu auch mal reindizieren 05.02.2013 Ende
    
    lab.Caption = ""
    lab.Refresh
    
    labglo.ForeColor = vbRed
    labglo.Caption = "Ballast löschen"
    labglo.Refresh
    
    Ballast_löschen
    
    labglo.ForeColor = vbRed
    labglo.Caption = "Container Lieferanten aktualisieren"
    labglo.Refresh
    
    Container_Lieferanten_aktualisieren
    
    lab.Caption = ""
    lab.Refresh
    
    labglo.ForeColor = vbRed
    labglo.Caption = "GDPdU schreiben"
    labglo.Refresh
    
    'Neu
    GDPdU_schreiben
    
    
    
    DoEvents
    If gbPenner_faerben Then
        TwoYearsNoVerkauftwerdenBlack lab, txtStatus, labglo
    End If
    DoEvents
    
    lab.Caption = ""
    lab.Refresh
    
'    bankenpflege lab, txtStatus, labglo
    
    lab.Caption = ""
    lab.Refresh
    
    labglo.ForeColor = vbRed
    labglo.Caption = "Lagerwerte werden geschrieben"
    labglo.Refresh
    
    DoEvents
    SEKSchreiben lab
    Lagerwerteschreiben lab
    
    DoEvents
    
    lab.Caption = ""
    lab.Refresh
    
    labglo.ForeColor = vbRed
    labglo.Caption = "Pennerwerte werden geschrieben"
    labglo.Refresh
    
    DoEvents
    PennerBestundSEK lab
    
    DoEvents
    
    speicherKundendurchscnittswerte
    
    DoEvents
    
    If LUGSAktuell = False Then

        lab.Caption = ""
        lab.Refresh

        labglo.ForeColor = vbRed
        labglo.Caption = "Lagerumschläge werden geschrieben"
        labglo.Refresh

        If alleLUGnachLief(txtStatus, frmWKL53.picprogress, False) Then

        End If

    End If
    
    DoEvents
    
    If GANALYSEAKTUELL = False Then
    
        lab.Caption = ""
        lab.Refresh
        
        labglo.ForeColor = vbRed
        labglo.Caption = "Analyse wird geschrieben"
        labglo.Refresh
        
        AnalyseZusammenstellen lab
    
    End If
    
    DoEvents
    
    lab.Caption = ""
    lab.Refresh
    
    labglo.ForeColor = vbRed
    labglo.Caption = "Duplikate (Artlief) löschen, nicht ausschalten!!!"
    labglo.Refresh
    
    DublikateDel lab
    
    DoEvents
    
    lab.Caption = ""
    lab.Refresh
    
    labglo.ForeColor = vbRed
    labglo.Caption = "Duplikate (Artikel) löschen, nicht ausschalten!!!"
    labglo.Refresh

    DublikateDelArtikel1 lab
    
    DoEvents
    
    
    
    
    
    lab.Caption = ""
    lab.Refresh
    
    labglo.ForeColor = vbRed
    labglo.Caption = "Neukundenberechnung, nicht ausschalten!!!"
    labglo.Refresh
    
    
    
    'Neukundenberechnung
    rechneNeuKunden
    
    
    DoEvents
    
    
    lab.Caption = ""
    lab.Refresh
    
    labglo.ForeColor = vbRed
    labglo.Caption = "beste Mitarbeiter, nicht ausschalten!!!"
    labglo.Refresh
    
    
    
    ermBestMitarbeiter
    DoEvents
    
    lab.Caption = ""
    lab.Refresh
    
    
    
    Set dabalokal = Nothing

    Kill cPfad2 & "END.MDB"
    Set dbEND = CreateDatabase(cPfad2 & "END.MDB", dbLangGeneral, dbVersion40)

    lMax = 0
    lgMax = 0
    lTabMax = 0
    
    labglo.ForeColor = vbRed
    labglo.Caption = "Datenbankgröße wird gemessen, nicht ausschalten!!!"
    labglo.Refresh
    
    lab.Caption = ""
    lab.Refresh
    
    db.TableDefs.Refresh
    lAnzTable = db.TableDefs.Count
    For lcount = 0 To lAnzTable - 1
        lMax = lMax + db.TableDefs(lcount).RecordCount

        name = db.TableDefs(lcount).name

        If UCase(Left(name, 4)) = "MSYS" Then
'            MsgBox name
        Else
            For j = 0 To db.TableDefs(lcount).Indexes.Count - 1
                db.TableDefs(lcount).Indexes.delete db.TableDefs(lcount).Indexes(j).name
                inname = db.TableDefs(lcount).Indexes(j).name
                lab.Caption = name & " " & inname
                lab.Refresh
            Next j
        End If
    Next lcount
    
    lab.Caption = ""
    lab.Refresh
    
    labglo.ForeColor = vbRed
    labglo.Caption = "Daten werden aktualisiert, nicht ausschalten!!!"
    labglo.Refresh
    
    neuFildatschreiben
    
    
    
    
    
    TabsAktuali labglo
    
    labglo.ForeColor = vbRed
    labglo.Caption = "nicht benötigte Daten werden gelöscht, nicht ausschalten!!!"
    labglo.Refresh
    
    TempTabsDelete lab
    
    labglo.ForeColor = vbRed
    labglo.Caption = "Datenbank wird komprimiert, nicht ausschalten!!!"
    labglo.Refresh
    
    
    If bOhneAnsicht Then
    
        If db_Compri_ohneAnsicht("Kissdata.MDB") = True Then
        
            db_Reindizieren gdBase, lab, txtStatus, labglo
            Kill cPfad2 & "KISIC.LZH"
    
            labglo.ForeColor = vbBlack
            labglo.Caption = "Alles Fertig"
            labglo.Refresh
    
            lab.ForeColor = vbBlack
            lab.Caption = "Alles Fertig"
            lab.Refresh
    
            schreibeProtokollDaba ("Erfolg Komprimierung kissdata")
    
            Pause (3)
    
            AktionAustragen "Komprimierung"
            
        Else
    
            schreibeProtokollDaba ("Fehler Komprimierung kissdata")
            labglo.ForeColor = vbRed
            labglo.Caption = "Fehler"
            labglo.Refresh
    
            lab.ForeColor = vbRed
            lab.Caption = "Fehler"
            lab.Refresh
    
            MsgBox "Winkiss wird beendet. Melden Sie sich bei der Hotline. 0511/955910", vbCritical, "Winkiss Hinweis:"
            End
        
        End If
    
    Else
    
        'Jetzt Tabelle für Tabelle umSchichten
        theBigFehler = False
        db.TableDefs.Refresh
        lAnzTable = db.TableDefs.Count
    
        For lcount = 0 To lAnzTable - 1
    
            name = db.TableDefs(lcount).name
    
            If UCase(Left(name, 4)) = "MSYS" Then
    '            MsgBox name
            Else
    
    
                lab.Caption = name
                lab.Refresh
    
                lTabMax = db.TableDefs(lcount).RecordCount
    
                Kill cPfad2 & "COMP.MDB"
                Kill cPfad2 & "TEMPCOMP.MDB"
    
                PauseSi CSng(gdDBPAUSE)
    
    
                Set dbCOMP = CreateDatabase(cPfad2 & "COMP.MDB", dbLangGeneral, dbVersion40)
                dbCOMP.Close
    
    
                TransferTab db, cPfad2 & "COMP.MDB", name
    
                If theBigFehler = True Then
                'raus
                    schreibeProtokollDaba ("Abbruch Komprimierung Datenlängenfehler in " & name)
                    labglo.ForeColor = vbRed
                    labglo.Caption = "Abbruch Ende Datenlängenfehler in: "
                    labglo.Refresh
    
                    lab.ForeColor = vbRed
                    lab.Caption = name
                    lab.Refresh
    
                    CompactmyDaba = False
                    Screen.MousePointer = 0
                    Exit Function
                End If
    
                DBEngine.CompactDatabase cPfad2 & "COMP.MDB", cPfad2 & "TEMPCOMP.mdb", dbLangGeneral
    
                Set dbTEMPCOMP = OpenDatabase(cPfad2 & "TEMPCOMP.MDB")
    
                TransferTab dbTEMPCOMP, cPfad2 & "END.MDB", name
                dbTEMPCOMP.Close
    
                Kill cPfad2 & "COMP.MDB"
                Kill cPfad2 & "TEMPCOMP.MDB"
    
                lgMax = lgMax + lTabMax
                txtStatus.Text = CStr(lgMax * 100 / lMax)
            End If
    
        Next lcount
    
        db.Close
    
        Set db = Nothing
        Set gdbMdb = Nothing
    
        dbEND.Close
        Set dbEND = Nothing
    
        If db_Copy(cPfad2, "END.MDB", sDB, lab, txtStatus, labglo) = True Then
    
            CompactmyDaba = True
    
            Set db = OpenDatabase(cPfad2 & "KISSDATA.MDB", True, False)
            db.NewPassword "", gsPasswort
            db.Close
    
            Set db = OpenDatabase(cPfad2 & "KISSDATA.MDB", False, False, "MS Access;PWD=" & gsPasswort)
    
            If NewTableSuchenDBKombi("ZZZ", db) = False Then
    
                giUmleitgrund = 5
    
                gcUmleittxt = "Beim Komprimieren der Datenbank ist ein schwerwiegender Fehler aufgetreten!" & vbCrLf
                gcUmleittxt = gcUmleittxt & "Drücken Sie 'Winkiss Beenden'!" & vbCrLf
    
                frmWKL60.Show 1
                End
    
            End If
            db_Reindizieren db, lab, txtStatus, labglo
            Kill cPfad2 & "KISIC.LZH"
    
'            labglo.ForeColor = vbBlack
'            labglo.Caption = "Alles Fertig"
'            labglo.Refresh
'
'            lab.ForeColor = vbBlack
'            lab.Caption = "Alles Fertig"
'            lab.Refresh
    
            schreibeProtokollDaba ("Erfolg Komprimierung kissdata")
    
            Pause (3)
    
            AktionAustragen "Komprimierung"
        Else
    
            schreibeProtokollDaba ("Fehler Komprimierung kissdata")
            labglo.ForeColor = vbRed
            labglo.Caption = "Fehler"
            labglo.Refresh
    
            lab.ForeColor = vbRed
            lab.Caption = "Fehler"
            lab.Refresh
    
            MsgBox "Winkiss wird beendet. Melden Sie sich bei der Hotline. 0511/955910", vbCritical, "Winkiss Hinweis:"
            End
        End If
    End If
    
    
    

    
Exit Function
LOKAL_ERROR:
    If err.Number = 3112 Or err.Number = 53 Or err.Number = 3110 Then
        Resume Next
    
    ElseIf err.Number = 3281 Then
        If name <> "" And inname <> "" Then
            MsgBox name & " " & inname
        End If
        Resume Next
    ElseIf err.Number = 3372 Then
        Resume Next
    ElseIf err.Number = 3022 Or err.Number = 3033 Then
        Resume Next
    ElseIf err.Number = 3377 Or err.Number = 3265 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul11"
        Fehler.gsFunktion = "CompactmyDaba"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1
        
    End If
End Function
Public Sub bankenpflege(lab As Label, txtStatus As TextBox, labglo As Label)
    On Error GoTo LOKAL_ERROR

    Dim cPfad     As String
    Dim sSQL As String
    
    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    
    
    Screen.MousePointer = 11

    
    If FileExists(cPfad & "BANKEN.DBF") Then
        
        txtStatus.Text = 0
        lab.Caption = "": lab.Refresh

        labglo.ForeColor = vbRed
        labglo.Caption = "Tabelle Banken wird ersetzt, bitte warten..."
        labglo.Refresh

        loeschNEW "BANKEN", gdBase
        
        txtStatus.Text = 20
        lab.Caption = "BANKEN.dbf wird importiert...": lab.Refresh

    
        sSQL = "Select * into BANKEN from BANKEN IN '" & cPfad & "' 'dBase IV;'"
        schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError

        txtStatus.Text = 46
        
        Kill cPfad & "BANKEN.DBF"
      
    
    End If
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "bankenpflege"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub

Public Function indexDel(db As Database)
On Error GoTo LOKAL_ERROR

    Dim lAnzTable       As Long
    Dim lcount          As Long
    Dim inname          As String

    db.TableDefs.Refresh
    lAnzTable = db.TableDefs.Count
    For lcount = 0 To lAnzTable - 1

        name = db.TableDefs(lcount).name

        If UCase(Left(name, 4)) = "MSYS" Then
'            MsgBox name
        Else
            For j = 0 To db.TableDefs(lcount).Indexes.Count - 1
                db.TableDefs(lcount).Indexes.delete db.TableDefs(lcount).Indexes(j).name
                inname = db.TableDefs(lcount).Indexes(j).name
            Next j
        End If
    Next lcount
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "indexDel"
    Fehler.gsFehlertext = "Beim Ermitteln des Dateidatums ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function db_Reindizieren(db As Database, lab As Label, txtStatus As TextBox, labglo As Label) As Boolean
On Error GoTo LOKAL_ERROR

    Dim cSQL        As String
    Dim i           As Integer
    Dim j           As Integer
    Dim cTabelle    As String
    Dim cIndex      As String
    Dim TabI(119) As IndexKombi
    
    i = 0
    
    TabI(i).Tabe = "ZUORDEAN"
    TabI(i).Inde = "GPEAN"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "BESTAEND"
    TabI(i).Inde = "Jahr"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "BESTAEND"
    TabI(i).Inde = "Monat"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "BESTAEND"
    TabI(i).Inde = "Bestand"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "SPEZIINFO"
    TabI(i).Inde = "BUCHUNGSNR"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "PFLEGORT"
    TabI(i).Inde = "bezeich"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "ARTMERK"
    TabI(i).Inde = "Artnr"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "BESTREST"
    TabI(i).Inde = "Artnr"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "ARTEAN_K"
    TabI(i).Inde = "Artnr"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "ARTEAN_K"
    TabI(i).Inde = "EAN"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "GLAGER"
    TabI(i).Inde = "DATUM"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "TERMINE"
    TabI(i).Inde = "DATUM"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "TERMINE"
    TabI(i).Inde = "BEDNU"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "TERMINE"
    TabI(i).Inde = "BUCHUNGSNR"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "INTERART"
    TabI(i).Inde = "ARTNR"
    TabI(i).IndeLis = ""
    i = i + 1
    
    
    TabI(i).Tabe = "ARTIKEL"
    TabI(i).Inde = "ARTNR"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "ARTIKEL"
    TabI(i).Inde = "BESTAND"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "ARTIKEL"
    TabI(i).Inde = "EAN"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "ARTIKEL"
    TabI(i).Inde = "EAN2"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "ARTIKEL"
    TabI(i).Inde = "EAN3"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "ARTIKEL"
    TabI(i).Inde = "BEZEICH"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "ARTIKEL"
    TabI(i).Inde = "LINR"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "ARTIKEL"
    TabI(i).Inde = "AGN"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "ARTIKEL"
    TabI(i).Inde = "PGN"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "ARTIKEL"
    TabI(i).Inde = "LIBESNR"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "ARTIKEL"
    TabI(i).Inde = "LASTDATE"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "ARTIKEL"
    TabI(i).Inde = "AWM"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "ARTIKEL"
    TabI(i).Inde = "AUFDAT"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "ZUGANG"
    TabI(i).Inde = "LINR"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "ZUGANG"
    TabI(i).Inde = "ARTNR"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "ZUGANG"
    TabI(i).Inde = "ADATE"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "UMLAGER"
    TabI(i).Inde = "ARTNR"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "KASSJOUR"
    TabI(i).Inde = "ARTNR"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "KASSJOUR"
    TabI(i).Inde = "ADATE"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "KASSJOUR"
    TabI(i).Inde = "KUNDNR"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "KASSJOUR"
    TabI(i).Inde = "MOPREIS"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "KASSJOUR"
    TabI(i).Inde = "EKPR"
    TabI(i).IndeLis = ""
    i = i + 1
        
    TabI(i).Tabe = "KASSJOUR"
    TabI(i).Inde = "LINR"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "KASSJOUR"
    TabI(i).Inde = "MENGE"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "KASSJOUR"
    TabI(i).Inde = "TEMP1"
    TabI(i).IndeLis = "ADATE, ARTNR"
    i = i + 1
    
    TabI(i).Tabe = "UMSATZ"
    TabI(i).Inde = "DATUM"
    TabI(i).IndeLis = ""
    i = i + 1
        
    TabI(i).Tabe = "UMS_ART"
    TabI(i).Inde = "ARTNR"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "UMS_ARTF"
    TabI(i).Inde = "Jahr"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "UMS_ARTF"
    TabI(i).Inde = "MONAT"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "ETIDRU"
    TabI(i).Inde = "Artnr"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "ETIDRU"
    TabI(i).Inde = "Filnr"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "UMSARTJ"
    TabI(i).Inde = "JAHR"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "UMSARTJ"
    TabI(i).Inde = "ARTNR"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "UMSKDJ"
    TabI(i).Inde = "JAHR"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "UMSKDJ"
    TabI(i).Inde = "KUNDNR"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "GUTSCH"
    TabI(i).Inde = "DAT_EINL"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "GUTSCH"
    TabI(i).Inde = "GUTSCHNR"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "GUTSCH"
    TabI(i).Inde = "LastDate"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "GUTSCH"
    TabI(i).Inde = "Lasttime"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "DVKART"
    TabI(i).Inde = "PRIMKEY"
    TabI(i).IndeLis = "ARTNR, JAHR, MONAT"
    i = i + 1

    TabI(i).Tabe = "GDLAGER"
    TabI(i).Inde = "PRIMKEY"
    TabI(i).IndeLis = "ARTNR, JAHR, MONAT"
    i = i + 1
    
    TabI(i).Tabe = "UMS_ART"
    TabI(i).Inde = "PRIMKEY"
    TabI(i).IndeLis = "ARTNR, JAHR, MONAT"
    i = i + 1
    
    TabI(i).Tabe = "UMS_ART"
    TabI(i).Inde = "DATUM"
    TabI(i).IndeLis = "JAHR, MONAT"
    i = i + 1
    
    TabI(i).Tabe = "UMS_ARTF"
    TabI(i).Inde = "PRIMKEY"
    TabI(i).IndeLis = "ARTNR, FILIALNR, JAHR, MONAT"
    i = i + 1
    
    TabI(i).Tabe = "UMSARTJF"
    TabI(i).Inde = "PRIMKEY"
    TabI(i).IndeLis = "ARTNR, FILIALNR, JAHR"
    i = i + 1
    
    TabI(i).Tabe = "UMSARTJ"
    TabI(i).Inde = "PRIMKEY"
    TabI(i).IndeLis = "ARTNR, JAHR"
    i = i + 1
    
    TabI(i).Tabe = "ZBESTAND"
    TabI(i).Inde = "PRIMKEY"
    TabI(i).IndeLis = "FILIALNR, ARTNR"
    i = i + 1
    
    TabI(i).Tabe = "UMSKDJ"
    TabI(i).Inde = "PRIMKEY"
    TabI(i).IndeLis = "KUNDNR, JAHR"
    i = i + 1

    TabI(i).Tabe = "BESTPROT"
    TabI(i).Inde = "aenart"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "BESTPROT"
    TabI(i).Inde = "LASTDATE"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "KUNDKASS"
    TabI(i).Inde = "KUNDNR"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "KUNDKASS"
    TabI(i).Inde = "ADATE"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "KUNDEN"
    TabI(i).Inde = "RECHNR"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "KUNDEN"
    TabI(i).Inde = "TITEL"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "KUNDEN"
    TabI(i).Inde = "VORNAME"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "KUNDEN"
    TabI(i).Inde = "STADT"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "KUNDEN"
    TabI(i).Inde = "DATUM1"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "KUNDEN"
    TabI(i).Inde = "KUNDNR"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "KUNDEN"
    TabI(i).Inde = "KUERZEL"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "KUNDEN"
    TabI(i).Inde = "NAME"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "KUNDEN"
    TabI(i).Inde = "PLZ"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "KUNDEN"
    TabI(i).Inde = "STRASSE"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "KUNDEN"
    TabI(i).Inde = "KUNDKART"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "KUNDEN"
    TabI(i).Inde = "AENDER"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "KUNDEN"
    TabI(i).Inde = "TEL"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "KUNDEN"
    TabI(i).Inde = "LastDate"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "KUNDEN"
    TabI(i).Inde = "ECIDENT"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "KUNDEN"
    TabI(i).Inde = "STATUS"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "KUNDEN"
    TabI(i).Inde = "SYNSTATUS"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "KUNDEN"
    TabI(i).Inde = "ANGELEGT"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "KASSBON"
    TabI(i).Inde = "DATUM"
    TabI(i).IndeLis = ""
    i = i + 1
    
'    TabI(i).Tabe = "KASSBOND"
'    TabI(i).Inde = "DATUM"
'    TabI(i).IndeLis = ""
'    i = i + 1
    
    TabI(i).Tabe = "BEDNAME"
    TabI(i).Inde = "BEDNU"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "KOLLVERK"
    TabI(i).Inde = "ARTNR"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "KOLLVERK"
    TabI(i).Inde = "ADATE"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "KOLLVERK"
    TabI(i).Inde = "KUNDNR"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "RETOURE"
    TabI(i).Inde = "ARTNR"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "RETOURE"
    TabI(i).Inde = "ADATE"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "RETOURE"
    TabI(i).Inde = "KUNDNR"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "TAUSCH"
    TabI(i).Inde = "ARTNR"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "TAUSCH"
    TabI(i).Inde = "ADATE"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "TAUSCH"
    TabI(i).Inde = "LINR"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "LISRT"
    TabI(i).Inde = "LINR"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "LISRT"
    TabI(i).Inde = "KUERZEL"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "LISRT"
    TabI(i).Inde = "LIEFBEZ"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "LISRT"
    TabI(i).Inde = "PLZ"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "LISRT"
    TabI(i).Inde = "STADT"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "LISRT"
    TabI(i).Inde = "LastDate"
    TabI(i).IndeLis = ""
    i = i + 1

    TabI(i).Tabe = "ARTLIEF"
    TabI(i).Inde = "ARTNR"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "ARTLIEF"
    TabI(i).Inde = "LEKPR"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "ARTLIEF"
    TabI(i).Inde = "LINR"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "ARTLIEF"
    TabI(i).Inde = "LIBESNR"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "ARTIKEL"
    TabI(i).Inde = "SYNSTATUS"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "ARTLIEF"
    TabI(i).Inde = "SYNSTATUS"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "ZBESTAND"
    TabI(i).Inde = "ARTNR"
    TabI(i).IndeLis = ""
    i = i + 1
    
    
    TabI(i).Tabe = "ARTLIEF"
    TabI(i).Inde = "ARTLINR"
    TabI(i).IndeLis = "ARTNR, LINR"
    i = i + 1
    
    TabI(i).Tabe = "KOPFMAIL"
    TabI(i).Inde = "HAUPTTEXT"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "PRSTERM"
    TabI(i).Inde = "ARTNR"
    TabI(i).IndeLis = ""
    i = i + 1
    
    
    
    TabI(i).Tabe = "PREISE"
    TabI(i).Inde = "ARTNR"
    TabI(i).IndeLis = ""
    
    TabI(i).Tabe = "TERMDEL"
    TabI(i).Inde = "KUNDNR"
    TabI(i).IndeLis = ""
    
    TabI(i).Tabe = "TERMDEL"
    TabI(i).Inde = "ADATE"
    TabI(i).IndeLis = ""

    Screen.MousePointer = 11
    
    schreibeProtokollDaba ("Reindizierung gestartet")
    
    txtStatus.Text = 0
    
    labglo.ForeColor = vbRed
    labglo.Caption = "Indizies werden neu erstellt, bitte warten..."
    labglo.Refresh
    
    'Start
    
    For j = 0 To i
        CheckIndexuDEL TabI(j).Tabe, TabI(j).Inde, TabI(j).IndeLis, db
        lab.Caption = "Tabelle: " & TabI(j).Tabe & " Index: " & TabI(j).Inde: lab.Refresh: txtStatus.Text = j
    Next j
    
    cTabelle = "ABGLEICH"

    lab.Caption = "Tabelle: " & cTabelle & " Schritt 1 ": lab.Refresh
    txtStatus.Text = i: i = i + 1

    loeschNEW "TEMP1", db
    loeschNEW "TEMP2", db

    BeginTrans
    cSQL = "Select ARTNR, LINR , 'N' as Erkannt into TEMP1 from ARTIKEL"
    db.Execute cSQL, dbFailOnError
    lab.Caption = "Tabelle: " & cTabelle & " Schritt 2 ": lab.Refresh
    txtStatus.Text = i: i = i + 1
    
    cSQL = "Update TEMP1 inner join Artlief on TEMP1.artnr = Artlief.artnr and TEMP1.LINR = Artlief.LINR "
    cSQL = cSQL & " Set TEMP1.Erkannt = 'J' "
    db.Execute cSQL, dbFailOnError
    
    
    
    
    
    
    

    cSQL = "Delete from TEMP1 where erkannt = 'J' "
    db.Execute cSQL, dbFailOnError

'    cSQL = "Delete TEMP1.* from TEMP1 inner join ARTLIEF on TEMP1.ARTNR = ARTLIEF.ARTNR and TEMP1.LINR = ARTLIEF.LINR"
'    db.Execute cSQL, dbFailOnError
    lab.Caption = "Tabelle: " & cTabelle & " Schritt 3 ": lab.Refresh
    txtStatus.Text = i: i = i + 1

    cSQL = "Select ARTIKEL.ARTNR, ARTIKEL.LINR, ARTIKEL.LEKPR, ARTIKEL.LIBESNR, ARTIKEL.MINMEN "
    cSQL = cSQL & "into TEMP2 from ARTIKEL inner join TEMP1 on ARTIKEL.ARTNR = TEMP1.ARTNR "
    db.Execute cSQL, dbFailOnError
    lab.Caption = "Tabelle: " & cTabelle & " Schritt 4 ": lab.Refresh
    txtStatus.Text = i: i = i + 1


    cSQL = "Insert into ARTLIEF Select * from TEMP2"
    db.Execute cSQL, dbFailOnError
    lab.Caption = "Tabelle: " & cTabelle & "Schritt 5 ": lab.Refresh
    txtStatus.Text = i: i = i + 1

    CommitTrans

    loeschNEW "TEMP1", db
    loeschNEW "TEMP2", db
    'Ende
    
    txtStatus.Text = "100"
    
    labglo.ForeColor = vbBlack
    labglo.Caption = "Fertig"
    labglo.Refresh
        
    lab.Caption = "Fertig": lab.Refresh
    schreibeProtokollDaba ("Erfolg Reindizierung")

Exit Function
LOKAL_ERROR:

'    If err.Number = 3167 Then
'        Resume Next
'    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul11"
        Fehler.gsFunktion = "db_Reindizieren"
        Fehler.gsFehlertext = "Im Programmteil Datenbank reindizieren ist ein Fehler aufgetreten."
    
        Fehlermeldung1
'    End If
    
End Function
Public Function ermkunden(Frame4 As Frame) As String
On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    Dim i As Integer
    
    sSQL = "Select * from Umsatz order by datum desc "
    
    frmWK21n.Label1(9).Caption = "Umsätze und Kundenzahlen der letzten 3 Tage."
    frmWK21n.Label1(7).Caption = ""
    i = 0
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then

        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            
            
        
            If Not IsNull(rsrs!Datum) Then
                frmWK21n.Label60(i).Caption = rsrs!Datum
            Else
                frmWK21n.Label60(i).Caption = ""
            End If
            
            If Not IsNull(rsrs!KUNZ1) Then
                frmWK21n.Label70(i).Caption = "Kunden: " & rsrs!KUNZ1
            Else
                frmWK21n.Label70(i).Caption = ""
            End If
            
            If Not IsNull(rsrs!UMSG1) Then
                frmWK21n.Label80(i).Caption = "Umsatz: " & Format$(rsrs!UMSG1, "####0.00 " & gcWaehrung)
            Else
                frmWK21n.Label80(i).Caption = ""
            End If
            
'            If Not IsNull(rsrs!lastvk) Then
'                If Trim(rsrs!lastvk) = "00:00:00" Then
'                    Label90(i).Caption = "letzter Verkauf: noch nie"
'                Else
'                    Label90(i).Caption = "letzter Verkauf: " & rsrs!lastvk
'                End If
'            Else
                frmWK21n.Label90(i).Caption = ""
'            End If
            i = i + 1
            
            If i = 3 Then
                rsrs.MoveLast
            End If
            
            
        rsrs.MoveNext
        Loop

    End If
    rsrs.Close: Set rsrs = Nothing
    
    Frame4.Visible = True
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul11"
    Fehler.gsFunktion = "ermkunden"
    Fehler.gsFehlertext = "Im Programmteil Datenbank reindizieren ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Public Function db_ReindizierenLo(db As Database, lab As Label, txtStatus As TextBox, labglo As Label, Frame4 As Frame, Frame3 As Frame) As Boolean
On Error GoTo LOKAL_ERROR

    Dim cSQL        As String
    Dim i           As Integer
    Dim j           As Integer
    
    Dim k           As Integer
    
    Dim cTabelle    As String
    Dim cIndex      As String
    Dim TabI(101) As IndexKombi
    
    i = 0
    TabI(i).Tabe = "ARTIKEL"
    TabI(i).Inde = "ARTNR"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "ARTIKEL"
    TabI(i).Inde = "BESTAND"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "ARTIKEL"
    TabI(i).Inde = "EAN"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "ARTIKEL"
    TabI(i).Inde = "EAN2"
    TabI(i).IndeLis = ""
    i = i + 1

    TabI(i).Tabe = "ARTIKEL"
    TabI(i).Inde = "EAN3"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "ARTIKEL"
    TabI(i).Inde = "BEZEICH"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "ARTIKEL"
    TabI(i).Inde = "LINR"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "ARTIKEL"
    TabI(i).Inde = "AGN"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "ARTIKEL"
    TabI(i).Inde = "PGN"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "ARTIKEL"
    TabI(i).Inde = "LIBESNR"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "ARTIKEL"
    TabI(i).Inde = "LASTDATE"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "ARTIKEL"
    TabI(i).Inde = "AWM"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "ARTIKEL"
    TabI(i).Inde = "AUFDAT"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "BESTPROT"
    TabI(i).Inde = "aenart"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "BESTPROT"
    TabI(i).Inde = "LASTDATE"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "KUNDEN"
    TabI(i).Inde = "KUNDNR"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "KUNDEN"
    TabI(i).Inde = "KUERZEL"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "KUNDEN"
    TabI(i).Inde = "NAME"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "KUNDEN"
    TabI(i).Inde = "PLZ"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "KUNDEN"
    TabI(i).Inde = "STADT"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "KUNDEN"
    TabI(i).Inde = "KUNDKART"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "KUNDEN"
    TabI(i).Inde = "AENDER"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "KUNDEN"
    TabI(i).Inde = "LastDate"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "KASSBON"
    TabI(i).Inde = "DATUM"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "BEDNAME"
    TabI(i).Inde = "BEDNU"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "KOLLVERK"
    TabI(i).Inde = "ARTNR"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "KOLLVERK"
    TabI(i).Inde = "ADATE"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "KOLLVERK"
    TabI(i).Inde = "KUNDNR"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "RETOURE"
    TabI(i).Inde = "ARTNR"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "RETOURE"
    TabI(i).Inde = "ADATE"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "RETOURE"
    TabI(i).Inde = "KUNDNR"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "TAUSCH"
    TabI(i).Inde = "ARTNR"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "TAUSCH"
    TabI(i).Inde = "ADATE"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "TAUSCH"
    TabI(i).Inde = "LINR"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "LISRT"
    TabI(i).Inde = "LINR"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "LISRT"
    TabI(i).Inde = "KUERZEL"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "LISRT"
    TabI(i).Inde = "LIEFBEZ"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "LISRT"
    TabI(i).Inde = "PLZ"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "LISRT"
    TabI(i).Inde = "STADT"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "LISRT"
    TabI(i).Inde = "LastDate"
    TabI(i).IndeLis = ""
    i = i + 1

    TabI(i).Tabe = "ARTLIEF"
    TabI(i).Inde = "ARTNR"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "ARTLIEF"
    TabI(i).Inde = "LINR"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "ARTLIEF"
    TabI(i).Inde = "LIBESNR"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "ARTLIEF"
    TabI(i).Inde = "SYNSTATUS"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "ZBESTAND"
    TabI(i).Inde = "ARTNR"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "ARTLIEF"
    TabI(i).Inde = "ARTLINR"
    TabI(i).IndeLis = "ARTNR, LINR"
    i = i + 1
    
    TabI(i).Tabe = "KOPFMAIL"
    TabI(i).Inde = "HAUPTTEXT"
    TabI(i).IndeLis = ""
    i = i + 1
    
    TabI(i).Tabe = "PREISE"
    TabI(i).Inde = "ARTNR"
    TabI(i).IndeLis = ""
    
    Screen.MousePointer = 11
    
    schreibeProtokollDaba ("Reindizierung gestartet lokal.mdb")
    
    txtStatus.Text = 0
    
    labglo.ForeColor = vbRed
    labglo.Caption = "Indizies werden neu erstellt, bitte warten..."
    labglo.Refresh
    
    'Start
    
    For j = 0 To i
        CheckIndexuDEL TabI(j).Tabe, TabI(j).Inde, TabI(j).IndeLis, db
        If k = 10 Then
            ermkunden Frame4
            Frame3.Visible = False
            Frame4.Visible = True
            Frame4.Refresh
        End If
        k = j * 2
        lab.Caption = "Tabelle: " & TabI(j).Tabe & " Index: " & TabI(j).Inde: lab.Refresh: txtStatus.Text = k
    Next j

    cTabelle = "ABGLEICH"
    
    lab.Caption = "Tabelle: " & cTabelle & " Schritt 1 ": lab.Refresh
    txtStatus.Text = k: k = k + 1
    
    loeschNEW "TEMP1", db
    loeschNEW "TEMP2", db
    
    BeginTrans
    cSQL = "Select ARTNR, LINR into TEMP1 from ARTIKEL"
    db.Execute cSQL, dbFailOnError
    lab.Caption = "Tabelle: " & cTabelle & " Schritt 2 ": lab.Refresh
    txtStatus.Text = k: k = k + 1
    
    cSQL = "Delete TEMP1.* from TEMP1 inner join ARTLIEF on TEMP1.ARTNR = ARTLIEF.ARTNR and TEMP1.LINR = ARTLIEF.LINR"
    db.Execute cSQL, dbFailOnError
    lab.Caption = "Tabelle: " & cTabelle & " Schritt 3 ": lab.Refresh
    txtStatus.Text = k: k = k + 1
    
    cSQL = "Select ARTIKEL.ARTNR, ARTIKEL.LINR, ARTIKEL.LEKPR, ARTIKEL.LIBESNR, ARTIKEL.MINMEN "
    cSQL = cSQL & "into TEMP2 from ARTIKEL inner join TEMP1 on ARTIKEL.ARTNR = TEMP1.ARTNR"
    db.Execute cSQL, dbFailOnError
    lab.Caption = "Tabelle: " & cTabelle & " Schritt 4 ": lab.Refresh
    txtStatus.Text = k: k = k + 1
    
    
    cSQL = "Insert into ARTLIEF Select * from TEMP2"
    db.Execute cSQL, dbFailOnError
    lab.Caption = "Tabelle: " & cTabelle & "Schritt 5 ": lab.Refresh
    txtStatus.Text = k: k = k + 1
    
    CommitTrans
    
    
    loeschNEW "TEMP1", db
    loeschNEW "TEMP2", db
    'Ende
    
    txtStatus.Text = "100"
    
    labglo.ForeColor = vbBlack
    labglo.Caption = "Fertig"
    labglo.Refresh
        
    lab.Caption = "Fertig": lab.Refresh
    schreibeProtokollDaba ("Erfolg Reindizierung lokal.mdb")

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul11"
    Fehler.gsFunktion = "db_ReindizierenLo"
    Fehler.gsFehlertext = "Im Programmteil Datenbank reindizieren ist ein Fehler aufgetreten."

    Fehlermeldung1
    
End Function
Public Sub CheckIndex(sTab As String, sIndex As String, sIndexLIST As String, db As Database)
    On Error GoTo LOKAL_ERROR

    Dim lcount      As Long
    Dim lAnzTable   As Long
    Dim sName       As String
    Dim inname      As String
    Dim sSQL        As String
    Dim j           As Long
    Dim bFound      As Boolean
    
    Dim cVergleichsname As String
    
    If gbSQLSERVER = True Then
        Exit Sub
    End If
    
    If gbSQLSERVER = True Then
        cVergleichsname = "DBO." & UCase(sTab)
    Else
        cVergleichsname = UCase(sTab)
    End If
    
    bFound = False
    
    db.TableDefs.Refresh
    lAnzTable = db.TableDefs.Count
    For lcount = 0 To lAnzTable - 1

        sName = db.TableDefs(lcount).name
        
        If UCase(sName) = UCase(cVergleichsname) Then
            For j = 0 To db.TableDefs(lcount).Indexes.Count - 1
                inname = db.TableDefs(lcount).Indexes(j).name
                If UCase(inname) = UCase(sIndex) Then
                    bFound = True
                    
                    Exit For
                Else
                    bFound = False
                    
                End If
            Next j
        End If
        
    Next lcount
    
    If bFound = False Then
        If sIndexLIST = "" Then
            sSQL = "Create Index " & sIndex & " on " & sTab & " (" & sIndex & ")"
            db.Execute sSQL, dbFailOnError
        Else
            sSQL = "Create Index " & sIndex & " on " & sTab & " (" & sIndexLIST & ")"
            db.Execute sSQL, dbFailOnError
        End If
        
        schreibeProtokollIndex " Index: " & sIndex & " Tabelle: " & sTab
    End If
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 3167 Then
        Resume Next
    Else
    
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul11"
        Fehler.gsFunktion = "CheckIndex"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
        Fehlermeldung1
    End If

End Sub
Public Sub CheckIndexuDEL(sTab As String, sIndex As String, sIndexLIST As String, db As Database)
    On Error GoTo LOKAL_ERROR

    Dim lcount      As Long
    Dim lAnzTable   As Long
    Dim sName       As String
    Dim inname      As String
    Dim sSQL        As String
    Dim j           As Long
    Dim bFound      As Boolean
    Dim bfoundTab      As Boolean
    
    If gbSQLSERVER = True Then
        Exit Sub
    End If
    
    bfoundTab = False
    bFound = False

    db.TableDefs.Refresh
    lAnzTable = db.TableDefs.Count
    For lcount = 0 To lAnzTable - 1

        sName = db.TableDefs(lcount).name
        
        If UCase(sName) = UCase(sTab) Then
            For j = 0 To db.TableDefs(lcount).Indexes.Count - 1
                inname = db.TableDefs(lcount).Indexes(j).name
                If UCase(inname) = UCase(sIndex) Then
                    bFound = True
                    Exit For
                Else
                    bFound = False
                    
                End If
            Next j
            bfoundTab = True
            GoTo step1
        Else
            bfoundTab = False
        End If
        
    Next lcount
    
step1:
    If bfoundTab = False Then
        Exit Sub
    End If
    
    If bFound Then
        sSQL = "Drop Index " & sIndex & " on " & sTab
        db.Execute sSQL, dbFailOnError
        
        
        If sIndexLIST = "" Then
            sSQL = "Create Index " & sIndex & " on " & sTab & " (" & sIndex & ")"
            db.Execute sSQL, dbFailOnError
        Else
            sSQL = "Create Index " & sIndex & " on " & sTab & " (" & sIndexLIST & ")"
            db.Execute sSQL, dbFailOnError
        End If
    End If
    
    If bFound = False Then
        If sIndexLIST = "" Then
            sSQL = "Create Index " & sIndex & " on " & sTab & " (" & sIndex & ")"
            db.Execute sSQL, dbFailOnError
        Else
            sSQL = "Create Index " & sIndex & " on " & sTab & " (" & sIndexLIST & ")"
            db.Execute sSQL, dbFailOnError
        End If
    End If
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul11"
    Fehler.gsFunktion = "CheckIndexuDEL"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1

End Sub
Public Function db_Copy(sPfad As String, sDBOld As String, sDBNew As String, lab As Label, txtStatus As TextBox, labglo As Label) As Boolean
On Error GoTo LOKAL_ERROR

    Dim dbOld       As DAO.Database
    Dim dbNew       As DAO.Database
    Dim lAnzTable   As Long
    Dim lcount      As Long
    Dim lgMax       As Long
    Dim lTabMax     As Long
    Dim name        As String
    Dim lMax        As Long
    
    db_Copy = False
    
    Set dbOld = OpenDatabase(sPfad & sDBOld, False, False, "MS Access;PWD=" & gsPasswort)

    Kill sPfad & sDBNew
    Set dbNew = CreateDatabase(sPfad & sDBNew, dbLangGeneral, dbVersion40)
    dbNew.Close
    
    labglo.ForeColor = vbRed
    labglo.Caption = "Datenbank wird kopiert, bitte warten..."
    labglo.Refresh
    
    
    dbOld.TableDefs.Refresh
    lAnzTable = dbOld.TableDefs.Count
    
    For lcount = 0 To lAnzTable - 1
        lMax = lMax + dbOld.TableDefs(lcount).RecordCount
    Next lcount
    
    
    dbOld.TableDefs.Refresh
    lAnzTable = dbOld.TableDefs.Count
    
    lgMax = 0
    
    For lcount = 0 To lAnzTable - 1
        name = dbOld.TableDefs(lcount).name
        
        If UCase(Left(name, 4)) = "MSYS" Then
'            MsgBox name
        Else
        
            lab.Caption = name
            lab.Refresh
            
            lTabMax = dbOld.TableDefs(lcount).RecordCount
            
            TransferTab dbOld, sPfad & "KISSDATA.MDB", name
            
            PauseSi CSng(gdDBPAUSE)
    
            lgMax = lgMax + lTabMax
            txtStatus.Text = CStr(lgMax * 100 / lMax)
        End If
    Next lcount
    
    dbOld.Close

    labglo.ForeColor = vbBlack
    labglo.Caption = "Fertig"
    labglo.Refresh
    
    db_Copy = True
    
Exit Function
LOKAL_ERROR:


If err.Number = 53 Then
    Resume Next
    
ElseIf err.Number = 70 Then
    Exit Function
    
Else
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul11"
    Fehler.gsFunktion = "db_Copy"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End If
End Function
Public Function db_CopySicher(sPfad As String, sDBOld As String, sDBNew As String, lab As Label, txtStatus As TextBox, labglo As Label) As Boolean
On Error GoTo LOKAL_ERROR

    Dim dbOld       As DAO.Database
    Dim dbNew       As DAO.Database
    Dim lAnzTable   As Long
    Dim lcount      As Long
    Dim lgMax       As Long
    Dim lTabMax     As Long
    Dim name        As String
    Dim lMax        As Long
    Dim sQuell      As String
    Dim sZiel       As String
    Dim lfail       As Long
    Dim lRet        As Long
    Dim rsrs        As Recordset
    Dim cPfad3      As String
    Dim dErgebnis   As Double
    
    cPfad3 = gcDBPfad
    If Right(cPfad3, 1) <> "\" Then
        cPfad3 = cPfad3 & "\"
    End If
    
    sQuell = App.Path & "\kisslite.ini"
    sZiel = gsSicherPfad & "\kisslite.ini"

    lRet = CopyFile(sQuell, sZiel, lfail)
    If lRet = 0 Then
        labglo.ForeColor = vbRed
        labglo.Caption = "Dies ist kein gültiger Pfad"
        labglo.Refresh
        
        Pause 2
        
        gsSicherPfad = cPfad3 & "Sicherung"

        Set rsrs = gdApp.OpenRecordset("WKEINSTE", dbOpenTable)
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            rsrs.Edit
            rsrs!SichPfad = gsSicherPfad
            rsrs.Update
            gbLokalModus = False
        End If
        
        rsrs.Close: Set rsrs = Nothing
        
        Exit Function
    End If
    
    
    
    Set dbOld = OpenDatabase(sPfad & sDBOld, False, False, "MS Access;PWD=" & gsPasswort)

    Kill sDBNew
    Set dbNew = CreateDatabase(sDBNew, dbLangGeneral, dbVersion40)
'    Set dbNew = CreateDatabase(sDBNew, dbLangGeneral, dbVersion30)
    dbNew.Close
    
    labglo.ForeColor = vbRed
    labglo.Caption = "Datenbank wird kopiert, bitte warten..."
    labglo.Refresh
    
    
    dbOld.TableDefs.Refresh
    lAnzTable = dbOld.TableDefs.Count
    
    For lcount = 0 To lAnzTable - 1
        lMax = lMax + dbOld.TableDefs(lcount).RecordCount
    Next lcount
    
    
    dbOld.TableDefs.Refresh
    lAnzTable = dbOld.TableDefs.Count
    
    lgMax = 0
    
    For lcount = 0 To lAnzTable - 1
        name = dbOld.TableDefs(lcount).name
        
        If UCase(Left(name, 4)) = "MSYS" Then
'            MsgBox name
        Else
        
            lab.Caption = name
            lab.Refresh
            
            lTabMax = dbOld.TableDefs(lcount).RecordCount
            
            TransferTab dbOld, sDBNew, name
    
            lgMax = lgMax + lTabMax
            dErgebnis = lgMax / (lMax / 100)
            txtStatus.Text = CStr(dErgebnis)
            
            
'            txtStatus.Text = CStr(lgMax * 100 / lMax)
        End If
    Next lcount
    
    dbOld.Close

    labglo.ForeColor = vbBlack
    labglo.Caption = "Fertig"
    labglo.Refresh
    
Exit Function
LOKAL_ERROR:


    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul11"
        Fehler.gsFunktion = "db_CopySicher"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1
       
    End If
End Function
Public Function db_CopySicher_zip(sPfad As String, lab As Label, txtStatus As TextBox, labglo As Label) As Boolean
On Error GoTo LOKAL_ERROR

    Dim sQuell      As String
    Dim sZiel       As String
    Dim cPfad3      As String
    Dim lHeute      As Long
    Dim rsrs        As Recordset

    lHeute = Fix(Now)
    
    cPfad3 = gcDBPfad
    If Right(cPfad3, 1) <> "\" Then
        cPfad3 = cPfad3 & "\"
    End If
    
    sQuell = App.Path & "\kisslite.ini"
    sZiel = gsSicherPfad & "\kisslite.ini"

    lRet = CopyFile(sQuell, sZiel, lfail)
    If lRet = 0 Then
        labglo.ForeColor = vbRed
        labglo.Caption = "Dies ist kein gültiger Pfad"
        labglo.Refresh
        
        Pause 2
        
        gsSicherPfad = cPfad3 & "Sicherung"

        Set rsrs = gdApp.OpenRecordset("WKEINSTE", dbOpenTable)
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            rsrs.Edit
            rsrs!SichPfad = gsSicherPfad
            rsrs.Update
            gbLokalModus = False
        End If
        
        rsrs.Close: Set rsrs = Nothing
        
        Exit Function
    End If
    
    labglo.ForeColor = vbRed
    labglo.Caption = "Datenbank wird kopiert, bitte warten..."
    labglo.Refresh
    
    If Not FileExists(gsSicherPfad & "\KD" & CStr(lHeute) & ".LZH") Then
        zipDllcheck
        Zip_Files "xyr", sPfad & "KISSDATA.MDB", gsSicherPfad & "\KD" & CStr(lHeute) & ".LZH", txtStatus
    Else
        labglo.ForeColor = vbRed
        labglo.Caption = "Die Sicherung der Datenbank wurde heute schon erzeugt - Fertig"
        labglo.Refresh
    End If
    
    labglo.ForeColor = vbBlack
    labglo.Caption = "Fertig"
    labglo.Refresh
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul11"
    Fehler.gsFunktion = "db_CopySicher_zip"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Sub TwoYearsNoVerkauftwerdenBlack(lab As Label, txtStatus As TextBox, labglo As Label)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    Dim cART As String
    Dim ctmp As String
    Dim datLVK As Date
    Dim datLZU As Date
    Dim lLastvk As Long
    Dim lHeute As Long
    Dim ldifferenz As Long
    
    Dim lAnz As Long
    Dim siAnzeige As Single
    
    lHeute = CLng(DateValue(Now))
    
    
'    If Day(DateValue(Now)) > 6 Then
'        Exit Sub
'    End If
    
    If gcFilNr > 0 Then
        Exit Sub
    End If

    Screen.MousePointer = 11

    txtStatus.Text = 10
    
    loeschNEW "ART55", gdBase
    CreateTable "ART55", gdBase
    
    lab.Caption = "": lab.Refresh

    labglo.ForeColor = vbRed
    labglo.Caption = "Uraltdaten werden ermittelt, bitte warten..."
    labglo.Refresh
    
    loeschNEW "ArtAwmMerk", gdBase
    
    sSQL = "select artnr,awm into ArtAwmMerk from artikel "
    gdBase.Execute sSQL, dbFailOnError

    txtStatus.Text = 20
    
    sSQL = "Update artikel "
    sSQL = sSQL & " Set artikel.awm = '0' where artikel.awm  = '92' "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 30


    sSQL = " Insert into ART55 select  ARTNR"
    sSQL = sSQL & " , Bezeich "
    sSQL = sSQL & " , RKZ "
    sSQL = sSQL & " , LEKPR "
    sSQL = sSQL & " , KVKPR1 "
    sSQL = sSQL & " , LINR "
    sSQL = sSQL & " , LPZ "
    sSQL = sSQL & " , BESTAND "

    sSQL = sSQL & ", '' as liefbez "
    sSQL = sSQL & ", AUFDAT  "
    sSQL = sSQL & ", EXDAT  "
    sSQL = sSQL & ", '01.01.2000' as LASTVK "
    sSQL = sSQL & ", '01.01.2000' as LASTZU "
    sSQL = sSQL & ", '' as Monat "
    
    sSQL = sSQL & " , LIBESNR from Artikel "
    sSQL = sSQL & " where aufdat <  " & CLng(DateValue(Now)) - 730
    sSQL = sSQL & " and bestand <= 0   "
    gdBase.Execute sSQL, dbFailOnError

    txtStatus.Text = 40
    
    labglo.ForeColor = vbRed
    labglo.Caption = "letzte Verkäufe werden ermittelt, bitte warten..."
    labglo.Refresh
    
    loeschNEW "kasslvk1", gdBase
    
    sSQL = "select artnr , max(adate) as maxdate into kasslvk1 from kassjour "
    sSQL = sSQL & " group by artnr "
    gdBase.Execute sSQL, dbFailOnError

    txtStatus.Text = 50
    
    sSQL = " Create index  artnr on kasslvk1(artnr) "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 60
    
    sSQL = "update ART55 k inner join kasslvk1 z on k.artnr = z.artnr "
    sSQL = sSQL & " set k.lastvk = z.maxdate "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 70
    
    labglo.ForeColor = vbRed
    labglo.Caption = "nicht relevante Daten werden gelöscht, bitte warten..."
    labglo.Refresh
    
    sSQL = "Delete from art55 where lastvk is null "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 80
    
    labglo.ForeColor = vbRed
    labglo.Caption = "Daten werden übernommen, bitte warten..."
    labglo.Refresh
    
    sSQL = "Delete from art55 where lastvk  > " & CLng(DateValue(Now)) - 730 & " "
    gdBase.Execute sSQL, dbFailOnError
    
    txtStatus.Text = 90
    
    sSQL = "Update artikel inner join art55 on art55.artnr = artikel.artnr "
    sSQL = sSQL & " Set artikel.awm = '92' "
    gdBase.Execute sSQL, dbFailOnError
    
    labglo.ForeColor = vbRed
    labglo.Caption = "Veränderungen werden bereitgestellt, bitte warten..."
    labglo.Refresh
    
    sSQL = "Update artikel inner join ArtAwmMerk on artikel.artnr = ArtAwmMerk.artnr "
    sSQL = sSQL & " Set artikel.lastdate = DateValue(now)"
    sSQL = sSQL & " where artikel.awm <> ArtAwmMerk.awm "
    gdBase.Execute sSQL, dbFailOnError

    
    
    
    
    



    txtStatus.Text = 97
    loeschNEW "ART55", gdBase
    txtStatus.Text = 98
    loeschNEW "kasslvk1", gdBase
    txtStatus.Text = 99
    loeschNEW "ArtAwmMerk", gdBase
    txtStatus.Text = 100
        
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul11"
    Fehler.gsFunktion = "TwoYearsNoVerkauftwerdenBlack"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Public Sub neuFildatschreiben()
On Error GoTo LOKAL_ERROR

    Dim lcount      As Long
    Dim lAnzTable   As Long
    Dim cFiledate   As Date
    Dim name        As String
    Dim sSQL        As String
    Dim cPfad       As String
    Dim rsrs        As Recordset
    
    Dim sdat        As String
    Dim lLief       As Long
    Dim sLiefname   As String
    
    
    If Not SpalteInTabellegefundenNEW("TABDATUM", "LIEFBEZ", gdBase) Then
        SpalteAnfuegenNEW "TABDATUM", "LIEFBEZ", "Text(35)", gdBase
        SpalteAnfuegenNEW "TABDATUM", "Kurzinfo", "Text(35)", gdBase
        SpalteAnfuegenNEW "TABDATUM", "Auftragsnr", "LONG", gdBase
         
        
        sSQL = "Update TABDATUM Set Kurzinfo = '' "
        gdBase.Execute sSQL, dbFailOnError
         
        sSQL = "Update TABDATUM Set Auftragsnr = 0 "
        gdBase.Execute sSQL, dbFailOnError
         
        sSQL = "Update TABDATUM Set LIEFBEZ = '' "
        gdBase.Execute sSQL, dbFailOnError
        
    End If
    
    
    
    gdBase.TableDefs.Refresh    'Dabarefresh
    lAnzTable = gdBase.TableDefs.Count

    For lcount = 0 To lAnzTable - 1
        name = gdBase.TableDefs(lcount).name
        cFiledate = Format(gdBase.TableDefs(lcount).DateCreated, "DD.MM.YY")
        
        If Left(name, 1) = "Q" Then
            
        
            If ermfildat(name) = "" Then
            
                sdat = Mid(name, 2, Len(name) - 2)
                lLief = Val(sdat)
                sLiefname = ""
                
                sSQL = "Select LIEFBEZ from LISRT where LINR = " & lLief
                Set rsrs = gdBase.OpenRecordset(sSQL)
                If Not rsrs.EOF Then
                    rsrs.MoveFirst
                    If Not IsNull(rsrs!LIEFBEZ) Then
                        sLiefname = rsrs!LIEFBEZ
                    End If
                End If
                rsrs.Close: Set rsrs = Nothing

                sSQL = "Insert into TABDATUM (Tabname,Tabdate,Liefbez) values"
                sSQL = sSQL & " ( '" & name & "','" & cFiledate & "','" & sLiefname & "')"
                gdBase.Execute sSQL, dbFailOnError
            Else
                sdat = Mid(name, 2, Len(name) - 2)
                lLief = Val(sdat)
                sLiefname = ""
                
                sSQL = "Select LIEFBEZ from LISRT where LINR = " & lLief
                Set rsrs = gdBase.OpenRecordset(sSQL)
                If Not rsrs.EOF Then
                    rsrs.MoveFirst
                    If Not IsNull(rsrs!LIEFBEZ) Then
                        sLiefname = rsrs!LIEFBEZ
                    End If
                End If
                rsrs.Close: Set rsrs = Nothing

                sSQL = "Update TABDATUM set Liefbez = '" & sLiefname & "'"
                sSQL = sSQL & " where tabname = '" & name & "'"
                gdBase.Execute sSQL, dbFailOnError
                
            End If
        End If

        If Left(name, 1) = "X" Then

            If ermfildat(name) = "" Then
                sSQL = "Insert into TABDATUM (Tabname,Tabdate) values"
                sSQL = sSQL & " ( '" & name & "','" & cFiledate & "')"
                gdBase.Execute sSQL, dbFailOnError
            End If
        End If

        If Left(name, 4) = "INV_" Then

            If ermfildat(name) = "" Then
                sSQL = "Insert into TABDATUM (Tabname,Tabdate) values"
                sSQL = sSQL & " ( '" & name & "','" & cFiledate & "')"
                gdBase.Execute sSQL, dbFailOnError

            End If
        End If

        If Left(name, 3) = "ILI" Then

            If ermfildat(name) = "" Then

                sSQL = "Insert into TABDATUM (Tabname,Tabdate) values"
                sSQL = sSQL & " ( '" & name & "','" & cFiledate & "')"
                gdBase.Execute sSQL, dbFailOnError

            End If
        End If

    Next lcount
    
    Set rsrs = gdBase.OpenRecordset("TABDATUM")
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            If Not IsNull(rsrs!tabname) Then
                name = rsrs!tabname
                If Not NewTableSuchenDBKombi(name, gdBase) Then
                   rsrs.delete
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
    Fehler.gsFormular = "Modul11"
    Fehler.gsFunktion = "neuFildatschreiben"
    Fehler.gsFehlertext = "Beim Ermitteln des Dateidatums ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Function ermfildat(sTabname As String) As String
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsrs As Recordset
    
    ermfildat = ""
    
    sSQL = " select * from TABDATUM where TABNAME = '" & sTabname & "'"
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!tabdate) Then
            ermfildat = rsrs!tabdate
        End If
    End If
    rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul11"
    Fehler.gsFunktion = "ermfildat"
    Fehler.gsFehlertext = "Beim Ermitteln des Dateidatums ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Sub TabelleRichten(sTab As String, db As Database)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    loeschNEW "tempaTab", db
    
    sSQL = " select * into tempatab from " & sTab
    db.Execute sSQL, dbFailOnError
    
    loeschNEW sTab, db
    CreateTable UCase$(sTab), db
    
    sSQL = "Insert into " & sTab & "  select * from tempatab "
    db.Execute sSQL, dbFailOnError
    
    Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul11"
    Fehler.gsFunktion = "TabelleRichten"
    Fehler.gsFehlertext = "Beim Richten der Tabelle " & sTab & " ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Public Sub TabsAktuali(labglo As Label)
On Error GoTo LOKAL_ERROR

Dim cSQL As String
Dim lDatum  As Long
Dim cPfad As String

cPfad = gcDBPfad
If Right(cPfad, 1) <> "\" Then
    cPfad = cPfad & "\"
End If

labglo.ForeColor = vbRed
labglo.Caption = "Textdateien werden gelöscht, bitte warten..."
labglo.Refresh

Kill cPfad & "KASSSTOP" & Trim(gcKasNum) & ".TXT"
Kill cPfad & "KASSSTOP_ALLE.TXT"
Kill cPfad & "Kassstop.txt"
Kill cPfad & "Synchronisieren.txt"
Kill cPfad & "Stammdaten.txt"
Kill cPfad & "Etiketten drucken.txt"
Kill cPfad & "Etiketten wählen.txt"
Kill cPfad & "Kassenabschluss.txt "
Kill cPfad & "Lieferung übernehmen.txt "

labglo.ForeColor = vbRed
labglo.Caption = "Kassenbondatei wird aktualisiert, bitte warten..."
labglo.Refresh

lDatum = DateValue(Now) - 7
cSQL = "Delete from KASSBON where DATUM < " & Trim$(Str$(lDatum)) & " "
gdBase.Execute cSQL, dbFailOnError

lDatum = DateValue(Now) - 180
cSQL = "Delete from FEEDB where EDATE < " & Trim$(Str$(lDatum)) & " "
gdBase.Execute cSQL, dbFailOnError

lDatum = DateValue(Now) - 180
cSQL = "Delete from FEEDB_TRANS where EDATE < " & Trim$(Str$(lDatum)) & " "
gdBase.Execute cSQL, dbFailOnError

lDatum = DateValue(Now) - 180
cSQL = "Delete from FEEDBF where EDATE < " & Trim$(Str$(lDatum)) & " "
gdBase.Execute cSQL, dbFailOnError

If NewTableSuchenDBKombi("FTPPRO", gdBase) Then
    lDatum = DateValue(Now) - 180

    cSQL = "Delete from FTPPRO where loadDATE < " & Trim$(Str$(lDatum)) & " "
    gdBase.Execute cSQL, dbFailOnError
End If

lDatum = DateValue(Now) - 1500
cSQL = "Delete from LAGERLLW where DATUM < " & Trim$(Str$(lDatum)) & " "
gdBase.Execute cSQL, dbFailOnError

lDatum = DateValue(Now) - 1500
cSQL = "Delete from LAGERLW where DATUM < " & Trim$(Str$(lDatum)) & " "
gdBase.Execute cSQL, dbFailOnError

lDatum = DateValue(Now) - 1500
cSQL = "Delete from LAGERMW where DATUM < " & Trim$(Str$(lDatum)) & " "
gdBase.Execute cSQL, dbFailOnError

lDatum = DateValue(Now) - 1500
cSQL = "Delete from PENLAGERLLW where DATUM < " & Trim$(Str$(lDatum)) & " "
gdBase.Execute cSQL, dbFailOnError

lDatum = DateValue(Now) - 1500
cSQL = "Delete from PENLAGERLW where DATUM < " & Trim$(Str$(lDatum)) & " "
gdBase.Execute cSQL, dbFailOnError

lDatum = DateValue(Now) - 1500
cSQL = "Delete from PENLAGERMW where DATUM < " & Trim$(Str$(lDatum)) & " "
gdBase.Execute cSQL, dbFailOnError

lDatum = DateValue(Now) - 1500
cSQL = "Delete from TAUSCH where ADATE < " & Trim$(Str$(lDatum)) & " "
gdBase.Execute cSQL, dbFailOnError

lDatum = DateValue(Now) - 1500
cSQL = "Delete from RETOURE where ADATE < " & Trim$(Str$(lDatum)) & " "
gdBase.Execute cSQL, dbFailOnError

lDatum = DateValue(Now) - 1500
cSQL = "Delete from KKZAHL where ADATE < " & Trim$(Str$(lDatum)) & " "
gdBase.Execute cSQL, dbFailOnError

lDatum = DateValue(Now) - 1500
cSQL = "Delete from KVKPR1PROT where LASTDATE < " & Trim$(Str$(lDatum)) & " "
gdBase.Execute cSQL, dbFailOnError




Dim lJahri As Long
lJahri = Year(Date)
    
cSQL = "Delete from bestaend where Jahr <  " & lJahri - 3
gdBase.Execute cSQL, dbFailOnError


cSQL = "Delete from Kassbond"
gdBase.Execute cSQL, dbFailOnError

cSQL = "Delete from GLAGER"
gdBase.Execute cSQL, dbFailOnError





labglo.ForeColor = vbRed
labglo.Caption = "Bestandsdatei wird aktualisiert, bitte warten..."
labglo.Refresh

If NewTableSuchenDBKombi("Bestprot", gdBase) Then

    cSQL = "Delete from Bestprot where lastdate < " & CLng(DateValue(Now)) - 365
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from Bestprot where aenart = 'Kassiervorgang'"
    gdBase.Execute cSQL, dbFailOnError

End If


labglo.ForeColor = vbRed
labglo.Caption = "Etiketten-Sic wird aktualisiert, bitte warten..."
labglo.Refresh

If NewTableSuchenDBKombi("ETIPROTS", gdBase) Then

    cSQL = "Delete from ETIPROTS where wedate < " & CLng(DateValue(Now)) - 180
    gdBase.Execute cSQL, dbFailOnError
    
    

End If

labglo.ForeColor = vbRed
labglo.Caption = "bestellte Artikel werden aktualisiert, bitte warten..."
labglo.Refresh

BestrestPruefandDel

labglo.ForeColor = vbRed
labglo.Caption = "Etikettendruck wird aktualisiert, bitte warten..."
labglo.Refresh

If DatendrinWieviel("Etidru", gdBase) > 5000 Then
    cSQL = "Delete from Etidru "
    gdBase.Execute cSQL, dbFailOnError
End If

If DatendrinWieviel("Etidruls", gdBase) > 5000 Then
    cSQL = "Delete from Etidruls "
    gdBase.Execute cSQL, dbFailOnError
End If

labglo.ForeColor = vbRed
labglo.Caption = "Bediener werden aktualisiert, bitte warten..."
labglo.Refresh

cSQL = "Delete from BEDNAME where BEDNU = 0 or BEDNU is NULL "
gdBase.Execute cSQL, dbFailOnError

labglo.ForeColor = vbRed
labglo.Caption = "Lagerplätze werden aktualisiert, bitte warten..."
labglo.Refresh

cSQL = "Delete from Lagerplatz where lagerp > 9999999  "
gdBase.Execute cSQL, dbFailOnError

labglo.ForeColor = vbRed
labglo.Caption = "Artikel werden aktualisiert, bitte warten..."
labglo.Refresh

cSQL = "Delete from ARTIKEL where ARTNR = 0 or ARTNR is NULL "
gdBase.Execute cSQL, dbFailOnError

cSQL = "Delete from Konditionen where"
cSQL = cSQL & " Kondi is null "
gdBase.Execute cSQL, dbFailOnError

cSQL = "Delete from Konditionen where"
cSQL = cSQL & " Kondi = 0 "
gdBase.Execute cSQL, dbFailOnError

cSQL = "Delete from  BESTAKT where SENDOK = -1  "
gdBase.Execute cSQL, dbFailOnError









'Hier werden alle überflüssigen Artnr-Einträge, die nicht mehr in der Tab Artikel enthalten sind aus der Tab Artlief gelöscht
    
loeschNEW "DIS_ARTNR", gdBase
    
cSQL = "Create Table DIS_ARTNR ( "
cSQL = cSQL & " disartnr Long "
cSQL = cSQL & ", ERKANNT TEXT(1) "
cSQL = cSQL & " ) "
gdBase.Execute cSQL, dbFailOnError

cSQL = "Insert into DIS_ARTNR select distinct(artnr) as disartnr "
cSQL = cSQL & " from artlief "
gdBase.Execute cSQL, dbFailOnError

cSQL = "Update DIS_ARTNR "
cSQL = cSQL & " set erkannt = 'N' "
gdBase.Execute cSQL, dbFailOnError

cSQL = "Update DIS_ARTNR inner join Artikel on DIS_ARTNR.disartnr = artikel.artnr "
cSQL = cSQL & " set DIS_ARTNR.erkannt = 'J' "
gdBase.Execute cSQL, dbFailOnError

cSQL = "Delete * from DIS_ARTNR "
cSQL = cSQL & " where erkannt = 'J' "
gdBase.Execute cSQL, dbFailOnError

Dim rsArt As DAO.Recordset
Dim cArtNr As String

sSQL = "Select * from DIS_ARTNR "
Set rsArt = gdBase.OpenRecordset(sSQL)
If Not rsArt.EOF Then
    rsArt.MoveFirst

    Do While Not rsArt.EOF
    
        cArtNr = ""
        If Not IsNull(rsArt!disartnr) Then
            cArtNr = Trim(rsArt!disartnr)
        End If

        If cArtNr <> "" Then
            cSQL = "Delete * from Artlief where artnr =  " & cArtNr
            gdBase.Execute cSQL, dbFailOnError
        End If
        rsArt.MoveNext
    Loop
    
End If
rsArt.Close: Set rsArt = Nothing

loeschNEW "DIS_ARTNR", gdBase

'ENDE Hier werden alle überflüssigen Artnr-Einträge, die nicht mehr in der Tab Artikel enthalten sind aus der Tab Artlief gelöscht





labglo.ForeColor = vbRed
labglo.Caption = "Artikel(MWST) werden aktualisiert, bitte warten..."
labglo.Refresh

cSQL = "update ARTIKEL Set MWST = 'V' where"
cSQL = cSQL & " MWST is null "
gdBase.Execute cSQL, dbFailOnError

cSQL = "update ARTIKEL Set MWST = 'V' where"
cSQL = cSQL & " MWST = '' "
gdBase.Execute cSQL, dbFailOnError

cSQL = "update ARTIKEL Set MWST = 'V' where"
cSQL = cSQL & " MWST = ' ' "
gdBase.Execute cSQL, dbFailOnError

cSQL = "update ARTIKEL Set MWST = 'V' where "
cSQL = cSQL & " MWST not in ('E','V','O') "
gdBase.Execute cSQL, dbFailOnError


cSQL = "delete * from termine where len(Uhrzeit) = 1"
gdBase.Execute cSQL, dbFailOnError



labglo.ForeColor = vbRed
labglo.Caption = "Artikel(Minbest) werden aktualisiert, bitte warten..."
labglo.Refresh

cSQL = "update ARTIKEL Set MINBEST = 0 where"
cSQL = cSQL & " MINBEST > 30000 "
gdBase.Execute cSQL, dbFailOnError

labglo.ForeColor = vbRed
labglo.Caption = "Artikel(Minbest2) werden aktualisiert, bitte warten..."
labglo.Refresh

cSQL = "update ARTIKEL Set MINBEST = 0 where"
cSQL = cSQL & " MINBEST <  -30000 "
gdBase.Execute cSQL, dbFailOnError

labglo.ForeColor = vbRed
labglo.Caption = "Artikel(Bestand) werden aktualisiert, bitte warten..."
labglo.Refresh

cSQL = "update ARTIKEL Set BESTAND = 0 where"
cSQL = cSQL & " BESTAND > 30000 "
gdBase.Execute cSQL, dbFailOnError

labglo.ForeColor = vbRed
labglo.Caption = "Artikel(Bestand2) werden aktualisiert, bitte warten..."
labglo.Refresh

cSQL = "update ARTIKEL Set BESTAND = 0 where"
cSQL = cSQL & " BESTAND <  -30000 "
gdBase.Execute cSQL, dbFailOnError

labglo.ForeColor = vbRed
labglo.Caption = "Artikel(Bestand3) werden aktualisiert, bitte warten..."
labglo.Refresh

cSQL = "update ARTIKEL Set BESTAND = 0 where"
cSQL = cSQL & " BESTAND is null "
gdBase.Execute cSQL, dbFailOnError

labglo.ForeColor = vbRed
labglo.Caption = "Artikel(KVKPR1) werden aktualisiert, bitte warten..."
labglo.Refresh

cSQL = "update ARTIKEL Set KVKPR1 = 0 where"
cSQL = cSQL & " KVKPR1 > 1000000 "
gdBase.Execute cSQL, dbFailOnError

labglo.ForeColor = vbRed
labglo.Caption = "Artikel(KVKPR1 1) werden aktualisiert, bitte warten..."
labglo.Refresh

cSQL = "update ARTIKEL Set KVKPR1 = 0 where"
cSQL = cSQL & " KVKPR1 <  -10000 "
gdBase.Execute cSQL, dbFailOnError



cSQL = "update ARTIKEL Set LPZ = 0 where LPZ is null "
gdBase.Execute cSQL, dbFailOnError

labglo.ForeColor = vbRed
labglo.Caption = "ARTLIEF(LEKPR) werden aktualisiert, bitte warten..."
labglo.Refresh

cSQL = "Update Artlief set lekpr = 0 where lekpr is null "
gdBase.Execute cSQL, dbFailOnError

cSQL = "Update Artlief set lekpr = round(lekpr,2) "
gdBase.Execute cSQL, dbFailOnError

labglo.ForeColor = vbRed
labglo.Caption = "ARTLIEF(RKZ) werden aktualisiert, bitte warten..."
labglo.Refresh

cSQL = "update ARTLIEF Set RKZ = 'J' where RKZ = '1' "
gdBase.Execute cSQL, dbFailOnError

cSQL = "update ARTLIEF Set RKZ = 'N' where RKZ = '0' "
gdBase.Execute cSQL, dbFailOnError

cSQL = "update ARTLIEF Set RKZ = 'N' where RKZ is null "
gdBase.Execute cSQL, dbFailOnError

cSQL = "update ARTLIEF Set RKZ = 'N' where RKZ = '' "
gdBase.Execute cSQL, dbFailOnError

cSQL = "update ARTLIEF Set RKZ = 'N' where RKZ = ' ' "
gdBase.Execute cSQL, dbFailOnError

labglo.ForeColor = vbRed
labglo.Caption = "ARTLIEF(Exdat) werden aktualisiert, bitte warten..."
labglo.Refresh

cSQL = "update ARTLIEF Set exdat = null where RKZ = 'N' and exdat <> null "
gdBase.Execute cSQL, dbFailOnError






loeschNEW "ZDBEINSTE", gdBase

cSQL = "SELECT * into ZDBEINSTE from DBEINSTE "
gdBase.Execute cSQL, dbFailOnError


cSQL = "Update KUNDEN SET STATUS= 'N' where status is null "
gdBase.Execute cSQL, dbFailOnError

labglo.ForeColor = vbRed
labglo.Caption = "Etiketten werden aktualisiert, bitte warten..."
labglo.Refresh

etidrukomp

'repos
If Not SpalteInTabellegefundenNEW("REPOS", "REIHENF", gdBase) Then
    SpalteAnfuegenNEW "REPOS", "REIHENF", "LONG", gdBase
End If
'KASSJOUR

If Not SpalteInTabellegefundenNEW("KASSJOUR", "ABOK", gdBase) Then
    SpalteAnfuegenNEW "KASSJOUR", "ABOK", "BIT", gdBase
End If

If Not SpalteInTabellegefundenNEW("KASSJOUR", "ZBONNR", gdBase) Then
    SpalteAnfuegenNEW "KASSJOUR", "ZBONNR", "LONG", gdBase
End If

If Not SpalteInTabellegefundenNEW("KASSJOUR", "RABKENN", gdBase) Then
    SpalteAnfuegenNEW "KASSJOUR", "RABKENN", "TEXT(1)", gdBase
End If

'ARTIKEL

If SpalteInTabellegefundenNEW("ARTIKEL", "BEST2", gdBase) Then
    cSQL = " Alter table ARTIKEL drop BEST2 "
    gdBase.Execute cSQL, dbFailOnError
End If

If SpalteInTabellegefundenNEW("ARTIKEL", "VKVMO", gdBase) Then
    cSQL = " Alter table ARTIKEL drop VKVMO "
    gdBase.Execute cSQL, dbFailOnError
End If

If SpalteInTabellegefundenNEW("ARTIKEL", "VKVJ", gdBase) Then
    cSQL = " Alter table ARTIKEL drop VKVJ "
    gdBase.Execute cSQL, dbFailOnError
End If

If SpalteInTabellegefundenNEW("ARTIKEL", "MOMENGE", gdBase) Then
    cSQL = " Alter table ARTIKEL drop MOMENGE "
    gdBase.Execute cSQL, dbFailOnError
End If

If SpalteInTabellegefundenNEW("ARTIKEL", "VKLJ", gdBase) Then
    cSQL = " Alter table ARTIKEL drop VKLJ "
    gdBase.Execute cSQL, dbFailOnError
End If

If SpalteInTabellegefundenNEW("ARTIKEL", "PVP", gdBase) Then
    cSQL = " Alter table ARTIKEL drop PVP "
    gdBase.Execute cSQL, dbFailOnError
End If

If SpalteInTabellegefundenNEW("ARTIKEL", "PEINHEIT", gdBase) Then
    cSQL = " Alter table ARTIKEL drop PEINHEIT "
    gdBase.Execute cSQL, dbFailOnError
End If

If SpalteInTabellegefundenNEW("ARTIKEL", "PEAN", gdBase) Then
    cSQL = " Alter table ARTIKEL drop PEAN "
    gdBase.Execute cSQL, dbFailOnError
End If

If SpalteInTabellegefundenNEW("ARTIKEL", "MARKE", gdBase) Then
    cSQL = " Alter table ARTIKEL drop MARKE "
    gdBase.Execute cSQL, dbFailOnError
End If



labglo.ForeColor = vbRed
labglo.Caption = "1. EANs werden aktualisiert, bitte warten..."
labglo.Refresh

cSQL = "update artikel set ean3 = '' where ean3 = ean "
gdBase.Execute cSQL, dbFailOnError

labglo.ForeColor = vbRed
labglo.Caption = "2. EANs werden aktualisiert, bitte warten..."
labglo.Refresh

cSQL = "update artikel set ean2 = '' where ean2 = ean "
gdBase.Execute cSQL, dbFailOnError

labglo.ForeColor = vbRed
labglo.Caption = "3. EANs werden aktualisiert, bitte warten..."
labglo.Refresh

cSQL = "update artikel set ean3 = '' where ean3 = ean2 "
gdBase.Execute cSQL, dbFailOnError


labglo.ForeColor = vbRed
labglo.Caption = "1. EANs werden aktualisiert, bitte warten..."
labglo.Refresh

cSQL = "update artikel set ean3 = '' where ean3 = '0'"
gdBase.Execute cSQL, dbFailOnError

labglo.ForeColor = vbRed
labglo.Caption = "2. EANs werden aktualisiert, bitte warten..."
labglo.Refresh

cSQL = "update artikel set ean2 = '' where ean2 = '0' "
gdBase.Execute cSQL, dbFailOnError

labglo.ForeColor = vbRed
labglo.Caption = "3. EANs werden aktualisiert, bitte warten..."
labglo.Refresh

cSQL = "update artikel set ean = '' where ean = '0' "
gdBase.Execute cSQL, dbFailOnError

labglo.ForeColor = vbRed
labglo.Caption = "Kunden werden aktualisiert(Titel), bitte warten..."
labglo.Refresh

cSQL = "update kunden set titel = '' where titel is null or len(titel)  = 0"
gdBase.Execute cSQL, dbFailOnError

labglo.ForeColor = vbRed
labglo.Caption = "Kunden werden aktualisiert(Rabatt), bitte warten..."
labglo.Refresh

cSQL = "update kunden set Rabatt = 0 where Rabatt is null "
gdBase.Execute cSQL, dbFailOnError

cSQL = "update LISRT set BR = 15 where BR is null "
gdBase.Execute cSQL, dbFailOnError

labglo.ForeColor = vbRed
labglo.Caption = "Artikel (neue Artikel) werden aktualisiert, bitte warten..."
labglo.Refresh

cSQL = "Update ARTIKEL set AWM = '0' where AWM = '98' "
cSQL = cSQL & " and aufdat <  " & CLng(DateValue(Now)) - 180
gdBase.Execute cSQL, dbFailOnError



If NewTableSuchenDBKombi("artean_K", gdBase) Then

    If SpalteInTabellegefundenNEW("artean_K", "erkannt", gdBase) = False Then
        cSQL = " Alter table artean_K add erkannt Text(1)  "
        gdBase.Execute cSQL, dbFailOnError
    End If

    cSQL = "Update artean_K set erkannt = 'N'  "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update artean_K inner join  ARTIKEL on artean_K.artnr =  ARTIKEL.artnr "
    cSQL = cSQL & " set artean_K.erkannt = 'J'"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from artean_K where erkannt = 'N'  "
    gdBase.Execute cSQL, dbFailOnError
    
    
    
    
    'gibt es den EAN aus Artean_K auch an EAN1 der Artikel
    
    cSQL = "Update artean_K set erkannt = 'N'  "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update artean_K inner join  ARTIKEL on artean_K.ean =  ARTIKEL.ean and artean_K.artnr =  ARTIKEL.artnr "
    cSQL = cSQL & " set artean_K.erkannt = 'J'"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from artean_K where erkannt = 'J'  "
    gdBase.Execute cSQL, dbFailOnError
    
    
    
    cSQL = "Update artean_K set erkannt = 'N'  "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update artean_K inner join  ARTIKEL on artean_K.ean =  ARTIKEL.ean2 and artean_K.artnr =  ARTIKEL.artnr "
    cSQL = cSQL & " set artean_K.erkannt = 'J'"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from artean_K where erkannt = 'J'  "
    gdBase.Execute cSQL, dbFailOnError
    
    
    cSQL = "Update artean_K set erkannt = 'N'  "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update artean_K inner join  ARTIKEL on artean_K.ean =  ARTIKEL.ean3 and artean_K.artnr =  ARTIKEL.artnr "
    cSQL = cSQL & " set artean_K.erkannt = 'J'"
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Delete from artean_K where erkannt = 'J'  "
    gdBase.Execute cSQL, dbFailOnError
    
    
    If SpalteInTabellegefundenNEW("artean_K", "erkannt", gdBase) = True Then
        cSQL = " Alter table artean_K drop erkannt   "
        gdBase.Execute cSQL, dbFailOnError
    End If
    
    
    'Duplis in Artean_k löschen
    loeschNEW "ImportDupli", gdBase
    
    sSQL = "select count(ean) as count ,ean into ImportDupli from Artean_k group by ean having count(ean) > 1"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from  ImportDupli where ean is null"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from  ImportDupli where trim(ean) = ''"
    gdBase.Execute sSQL, dbFailOnError
    
    Dim cEAN As String
    Dim rsrs As DAO.Recordset
    
    Set rsrs = gdBase.OpenRecordset("ImportDupli", dbOpenTable)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!EAN) Then
                cEAN = Trim(rsrs!EAN)
            End If

            sSQL = "Select * from Artean_k where ean = '" & cEAN & "'"
            Set rsArt = gdBase.OpenRecordset(sSQL)
            If Not rsArt.EOF Then
                rsArt.MoveFirst

                rsArt.MoveNext
                Do While Not rsArt.EOF

                    rsArt.delete
                    rsArt.MoveNext
                Loop
                rsrs.MoveNext
            End If
            rsArt.Close: Set rsArt = Nothing
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
End If

labglo.ForeColor = vbRed
labglo.Caption = "Artikel (Farben) werden aktualisiert, bitte warten..."
labglo.Refresh


AWMsetzen1

labglo.ForeColor = vbRed
labglo.Caption = "UMS_ARTF wird aktualisiert(" & Year(Now) - 10 & "), bitte warten..."
labglo.Refresh

sSQL = "Delete from UMS_ARTF where Jahr < year(now) - 10"
gdBase.Execute sSQL, dbFailOnError

labglo.ForeColor = vbRed
labglo.Caption = "UMS_ARTF wird aktualisiert(" & Year(Now) - 9 & "), bitte warten..."
labglo.Refresh

sSQL = "Delete from UMS_ARTF where Jahr < year(now) - 9"
gdBase.Execute sSQL, dbFailOnError

labglo.ForeColor = vbRed
labglo.Caption = "UMS_ARTF wird aktualisiert(" & Year(Now) - 8 & "), bitte warten..."
labglo.Refresh

sSQL = "Delete from UMS_ARTF where Jahr < year(now) - 8"
gdBase.Execute sSQL, dbFailOnError

labglo.ForeColor = vbRed
labglo.Caption = "UMS_ARTF wird aktualisiert(" & Year(Now) - 7 & "), bitte warten..."
labglo.Refresh

sSQL = "Delete from UMS_ARTF where Jahr < year(now) - 7"
gdBase.Execute sSQL, dbFailOnError

labglo.ForeColor = vbRed
labglo.Caption = "UMS_ARTF wird aktualisiert(" & Year(Now) - 6 & "), bitte warten..."
labglo.Refresh

sSQL = "Delete from UMS_ARTF where Jahr < year(now) - 6"
gdBase.Execute sSQL, dbFailOnError

labglo.ForeColor = vbRed
labglo.Caption = "UMS_ARTF wird aktualisiert(" & Year(Now) - 5 & "), bitte warten..."
labglo.Refresh

sSQL = "Delete from UMS_ARTF where Jahr < year(now) - 5"
gdBase.Execute sSQL, dbFailOnError

labglo.ForeColor = vbRed
labglo.Caption = "UMS_ARTF wird aktualisiert(" & Year(Now) - 4 & "), bitte warten..."
labglo.Refresh

sSQL = "Delete from UMS_ARTF where Jahr < year(now) - 4"
gdBase.Execute sSQL, dbFailOnError

labglo.ForeColor = vbRed
labglo.Caption = "UMS_ARTF wird aktualisiert(" & Year(Now) - 3 & "), bitte warten..."
labglo.Refresh

sSQL = "Delete from UMS_ARTF where Jahr < year(now) - 3"
gdBase.Execute sSQL, dbFailOnError

labglo.ForeColor = vbRed
labglo.Caption = "Bedienerprotokolle werden aktualisiert, bitte warten..."
labglo.Refresh

KassBedp_Kleinhalten



Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul11"
        Fehler.gsFunktion = "TabsAktuali"
        Fehler.gsFehlertext = "Im Programmteil Datenbank ist ein Fehler aufgetreten."
    
        Fehlermeldung1
    End If
End Sub
Public Sub SEKSchreiben(labglo As Label)
On Error GoTo LOKAL_ERROR

Dim cSQL As String

labglo.ForeColor = vbRed
labglo.Caption = "Artikel(SEK) werden aktualisiert, bitte warten..."
labglo.Refresh

cSQL = "update ARTIKEL Set EKPR = 0 where"
cSQL = cSQL & " EKPR is null "
gdBase.Execute cSQL, dbFailOnError

'Hier werden alle SEK = 0 mit einem LEK > 0 gefüllt
    
loeschNEW "SCHNITT_NULL", gdBase
    
cSQL = "Create Table SCHNITT_NULL ( "
cSQL = cSQL & " ARTNR Long "
cSQL = cSQL & ", EKPR DOUBLE "
cSQL = cSQL & ", LEKPR DOUBLE "
cSQL = cSQL & " ) "
gdBase.Execute cSQL, dbFailOnError

cSQL = "Update ARTIKEL "
cSQL = cSQL & " set ekpr = 0 where Ekpr is null "
gdBase.Execute cSQL, dbFailOnError

cSQL = "Insert into SCHNITT_NULL Select "
cSQL = cSQL & " ARTNR "
cSQL = cSQL & " , EKPR "
cSQL = cSQL & " , 0 as LEKPR "
cSQL = cSQL & " from ARTIKEL where EKPR <= 0 and KVKPR1 >= 0 "
gdBase.Execute cSQL, dbFailOnError

'kleineste LEK-tabelle bauen

loeschNEW "KL_ARTLIEF", gdBase

cSQL = "Select Artnr, min(LEKPR) as MinLEKPR into KL_ARTLIEF"
cSQL = cSQL & " from ARTLIEF where LEKPR > 0 group by Artnr  "
gdBase.Execute cSQL, dbFailOnError

cSQL = "Update SCHNITT_NULL inner join KL_ARTLIEF on SCHNITT_NULL.artnr = KL_ARTLIEF.Artnr"
cSQL = cSQL & " set SCHNITT_NULL.Lekpr = KL_ARTLIEF.MinLEKPR "
gdBase.Execute cSQL, dbFailOnError

'Final updaten
cSQL = "Update Artikel inner join SCHNITT_NULL on Artikel.artnr = SCHNITT_NULL.Artnr"
cSQL = cSQL & " set Artikel.ekpr = SCHNITT_NULL.Lekpr "
cSQL = cSQL & " where Artikel.EKPR = 0 "
gdBase.Execute cSQL, dbFailOnError

loeschNEW "SCHNITT_NULL", gdBase
loeschNEW "KL_ARTLIEF", gdBase

'Ende Teil 1

Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul11"
    Fehler.gsFunktion = "SEKSchreiben"
    Fehler.gsFehlertext = "Im Programmteil Datenbank ist ein Fehler aufgetreten."

    Fehlermeldung1
    

End Sub
Public Sub TempTabsDelete(lab As Label)
On Error GoTo LOKAL_ERROR

Dim cSQL As String
Dim lDatum  As Long
Dim cPfad As String
Dim i As Long

cPfad = gcDBPfad
If Right(cPfad, 1) <> "\" Then
    cPfad = cPfad & "\"
End If

For i = 0 To 31
    lab.Caption = "afcs" & Format$(i, "0#") & " wird gelöscht..."
    lab.Refresh
    
    loeschNEW "afcs" & Format$(i, "0#"), gdBase
    loeschNEW "umss" & Format$(i, "0#"), gdBase
Next i

If gsStatkundnr <> "" Then
    For i = 0 To 55
        loeschNEW "G" & gsStatkundnr & Format$(i, "00#"), gdBase
    Next i
    
    For i = 0 To 55
        loeschNEW "G" & gsStatkundnr & Format$(i, "0#"), gdBase
    Next i
End If

loeschNEW "ETI_EXPO", gdBase
loeschNEW "ETIPOOL", gdBase

loeschNEW "GUT_OUT", gdBase
loeschNEW "Z_OUT", gdBase
loeschNEW "CTIKEL", gdBase
loeschNEW "KUNDENTEMP", gdBase
loeschNEW "KUL_IN", gdBase

loeschNEW "Lieflw1", gdBase
loeschNEW "Lieflw2", gdBase
loeschNEW "Lieflw3", gdBase
loeschNEW "Lieflw4", gdBase
loeschNEW "Lieflw5", gdBase
loeschNEW "Lieflw6", gdBase
loeschNEW "TOPUMSATZ", gdBase
loeschNEW "LagerlwJETZT", gdBase
loeschNEW "Kass55", gdBase
loeschNEW "Zu55", gdBase
loeschNEW "KUNDAKASSE3", gdBase
loeschNEW "TAUZ", gdBase
loeschNEW "KUNDKASS2", gdBase
loeschNEW "TOPARTUMSATZ", gdBase

loeschNEW "ART551", gdBase
loeschNEW "ART55B", gdBase
loeschNEW "DIFFDRUCK", gdBase




loeschNEW "afctt", gdBase
loeschNEW "afcd03", gdBase
loeschNEW "AfcKas", gdBase
loeschNEW "AINV", gdBase
loeschNEW "ART61a", gdBase
loeschNEW "ART59", gdBase
loeschNEW "ART59d", gdBase
loeschNEW "ARTKUM", gdBase
loeschNEW "ARTNurBest", gdBase
loeschNEW "ARTwirk", gdBase

loeschNEW "BARGOUT", gdBase
loeschNEW "BED_out", gdBase
loeschNEW "BEDB_ARTIKEL", gdBase
loeschNEW "BEDNAME2", gdBase
loeschNEW "BEDPL", gdBase
loeschNEW "BEDZU1", gdBase
loeschNEW "BESTPOUT", gdBase
loeschNEW "Btemp55", gdBase
loeschNEW "DBESTART", gdBase

loeschNEW "UMSTEMP", gdBase
loeschNEW "UARTF", gdBase
loeschNEW "RAUSTEIL", gdBase
loeschNEW "Master", gdBase
loeschNEW "kassbono", gdBase
loeschNEW "GUFILTER", gdBase
loeschNEW "KUDEL", gdBase
loeschNEW "KUANTE", gdBase
loeschNEW "KUNZTEMP", gdBase
loeschNEW "KVKPR1POUT", gdBase
loeschNEW "PRINTSHOP", gdBase
        

loeschNEW "KUNDAZEALLE", gdBase
loeschNEW "STERNTEMP1", gdBase
loeschNEW "STERNTEMP2", gdBase
loeschNEW "STERNTEMP3", gdBase
loeschNEW "STERNTEMP4", gdBase
loeschNEW "STERNTEMP5", gdBase
loeschNEW "STERNTEMP6", gdBase
loeschNEW "STERNTEMP7", gdBase
loeschNEW "STERNTEMPZ", gdBase
loeschNEW "STERNTEMPT", gdBase

loeschNEW "CORDER", gdBase
loeschNEW "LORDER", gdBase
loeschNEW "SORDER", gdBase
loeschNEW "RORDER", gdBase
loeschNEW "BUORDER", gdBase
loeschNEW "KONTINP", gdBase
loeschNEW "PreisLZU", gdBase
loeschNEW "PREISLGU", gdBase
loeschNEW "PLKOPF", gdBase
loeschNEW "PreisKass", gdBase
loeschNEW "PreisKassT", gdBase
loeschNEW "NEGSPANNE", gdBase

loeschNEW "STADPROREWE", gdBase
loeschNEW "STADPROBELA", gdBase
loeschNEW "STADPROSTRECKE", gdBase
loeschNEW "STADPROLUENING", gdBase

loeschNEW "ZuAusUV", gdApp
loeschNEW "beTemp", gdApp

loeschNEW "KUOKB", gdBase
loeschNEW "LPZT", gdBase
loeschNEW "LPZT2", gdBase
loeschNEW "Kred2", gdBase
loeschNEW "GutschN", gdBase
loeschNEW "eanlite", gdBase
loeschNEW "EANALL", gdBase
loeschNEW "LfEAN", gdBase
loesch "WKL048"         'WKL48
loesch "DRU_LISTE"
loeschNEW "DRUMBAE", gdBase       'Mindestbestand ermitteln
loeschNEW "LOG", gdBase
loeschNEW "ARTTEMP9", gdBase
loesch "UMVERG"         'Umsatzstatistik
loesch "UMTEMP"
loesch "UMTME"

loesch "zeitzone"       'Zeitenstatistik
loesch "DRU_25c"

loesch "artueb"         'Artikelstatistik
loesch "artew"
loesch "artewp"
loesch "artewz"
loesch "arthit"
loesch "chrono"
loesch "atemp"

loesch "detail"         'Bedienerstatistik
loesch "bedno"
loesch "bedz"
loesch "grida"

loesch "Sortimen"       'Sortimentsanalyse
loesch "SORTHEAD"

loesch "umsatzs"        'Stammdaten einlesen
loesch "umsatzx"
loesch "filbeste"
'loesch "umvteil"
loesch "linbtemp"
loesch "tplinbez"
loesch "mastemp"
loesch "master2"
loesch "tempxxxx"
loesch "vorschlz"
loesch "vorschlp"

loesch "ETIDRU3"        'Etikettendruck 30 und 31
loesch "ETIDRU2"
loesch "DRU_GRUN"
loesch "ETITEM"
loeschNEW "EUMS", gdBase

loesch "liefstat"       'Lieferantenstatistik
loesch "lieftemp"
loesch "gode"
loesch "liefplus"
loesch "te12"
loesch "tempo"
loesch "lieflw"
loesch "liefew"
loesch "liefewp"
loesch "liefewz"
loesch "liefhit"
loeschNEW "ErtragTe", gdBase
loeschNEW "Bestprot1", gdBase
loeschNEW "adrute", gdBase
loeschNEW "afcsyn", gdBase
loeschNEW "aLite", gdBase
loeschNEW "Kund1", gdBase

loeschNEW "tmp1", gdBase
loeschNEW "tmp2", gdBase
loeschNEW "tmp3", gdBase
loeschNEW "tmp", gdBase


loesch "Kuteilme"       'Kundenanalyse
loesch "Kuteil"
loesch "stdatei"
loesch "stdater"

loeschapp "vorschlz"    'Bestellvorschlag app.path
loeschapp "vorschlp"
loeschapp "vorsort"
loeschapp "vorsortz"
loeschapp "winanfu"

loesch "DRU_FILT"       'WKLaj

loesch "UMS_MON"        'WKL25e

loesch "DRU_REPO"       'WKL24c
loesch "DRU_REKO"

loesch "DRU_TEMP"       'WKL25h

loesch "TEMP"           'Zusammenfassung Tagesabschlüsse wkl21f
loeschNEW "AFCSTATS", gdBase


loesch "Tagkopf"        'Kassenabschluss
loeschNEW "TAGKOPF_" & srechnertab, gdBase

loesch "LAYOUT2"        'Layoutbearbeitung frm50g

loesch "VIRTUEL"        'Kundenliste WKL47
loeschapp "VIRTUEL"

loesch "DRU_TEXT"       'Verkaufsprotokoll  WKL25d
loesch "VKPRO"

loesch "Rabatt"         'Rabattverkauf WKL25f
loesch "vkpro1"

loesch "vkpro1"         'Favoritenliste WKL44

loesch "DRU_KUND"       'WKL47

loesch "DRU_TERM"       'WKL82

loesch "ARTHEAD"        'WKL42
loesch "ARTDRUCK"
loeschapp "ARTHEAD"
loeschapp "ARTDRUCK"

loesch "INV_LITE"       'Inventur

loeschNEW "DRU_BARG", gdBase      'WKL21b

loesch "DRU_EINK"       'WKL45
loesch "DRU_EKLJ"

loesch "DRU_OFKR"       'WKL24
loeschNEW "DRU_ALKR", gdBase

loeschNEW "artikelzes", gdBase
loeschNEW "dupliean", gdBase
loeschNEW "lartlief", gdBase

loeschNEW "arttemp7", gdBase
loeschNEW "KUHIST", gdBase
loeschNEW "btikel", gdBase

loeschNEW "artlifzes", gdBase
loeschNEW "katemp33", gdBase

loeschNEW "kassjour34", gdBase
loeschNEW "kassjourzes", gdBase
loeschNEW "ZuTemp", gdApp
loeschNEW "ZuAusUV", gdApp
    
loeschNEWDTA "DTA", gdBase
loeschNEW "AAT", gdBase
loeschNEW "BAT", gdBase

Dim cKillPfad As String
cKillPfad = cPfad & "LPROTOK\EAN.txt"
Kill cKillPfad

'cKillPfad = cPfad & "LPROTOK\KVKPR1.txt"
'Kill cKillPfad

cKillPfad = cPfad & "LPROTOK\BEST.txt"
Kill cKillPfad

cKillPfad = cPfad & "LPROTOK\PROABL.txt"
Kill cKillPfad

cKillPfad = cPfad & "LPROTOK\BENUTZER.txt"
Kill cKillPfad

cKillPfad = cPfad & "LPROTOK\DABAABL.txt"
Kill cKillPfad





Dim lAnzTable As Long
Dim sName As String
Dim lcount As Long

gdBase.TableDefs.Refresh    'Dabarefresh

'alle Tabellen, die mit TBVVJ beginnen 'außer
lAnzTable = gdBase.TableDefs.Count
For lcount = 0 To lAnzTable - 1
    sName = gdBase.TableDefs(lcount).name
    If Left(UCase$(sName), 5) = "TBVVJ" Then
    
        Select Case UCase$(sName)
'            Case "KUNDAUSLIEF", "KUNDAZE"
        
            Case Else
                sSQL = "drop table " & sName
                gdBase.Execute sSQL, dbFailOnError
        End Select
    End If
Next lcount

gdBase.TableDefs.Refresh    'Dabarefresh

'alle Tabellen, die mit TBVJ beginnen 'außer
lAnzTable = gdBase.TableDefs.Count
For lcount = 0 To lAnzTable - 1
    sName = gdBase.TableDefs(lcount).name
    If Left(UCase$(sName), 4) = "TBVJ" Then
    
        Select Case UCase$(sName)
'            Case "KUNDAUSLIEF", "KUNDAZE"
        
            Case Else
                sSQL = "drop table " & sName
                gdBase.Execute sSQL, dbFailOnError
        End Select
    End If
Next lcount

gdBase.TableDefs.Refresh    'Dabarefresh

'alle Tabellen, die mit SCHWPUNKT beginnen 'außer
lAnzTable = gdBase.TableDefs.Count
For lcount = 0 To lAnzTable - 1
    sName = gdBase.TableDefs(lcount).name
    If Left(UCase$(sName), 9) = "SCHWPUNKT" Then
    
        Select Case UCase$(sName)
'            Case "KUNDAUSLIEF", "KUNDAZE"
        
            Case Else
                sSQL = "drop table " & sName
                gdBase.Execute sSQL, dbFailOnError
        End Select
    End If
Next lcount

gdBase.TableDefs.Refresh    'Dabarefresh

'alle Tabellen, die mit TAGKOPF_ beginnen 'außer
lAnzTable = gdBase.TableDefs.Count
For lcount = 0 To lAnzTable - 1
    sName = gdBase.TableDefs(lcount).name
    If Left(UCase$(sName), 8) = "TAGKOPF_" Then
    
        Select Case UCase$(sName)
'            Case "KUNDAUSLIEF", "KUNDAZE"
        
            Case Else
                sSQL = "drop table " & sName
                gdBase.Execute sSQL, dbFailOnError
        End Select
    End If
Next lcount

gdBase.TableDefs.Refresh    'Dabarefresh

'alle Tabellen, die mit TB beginnen 'außer
lAnzTable = gdBase.TableDefs.Count
For lcount = 0 To lAnzTable - 1
    sName = gdBase.TableDefs(lcount).name
    If Left(UCase$(sName), 2) = "TB" Then
    
        Select Case UCase$(sName)
'            Case "KUNDAUSLIEF", "KUNDAZE"
        
            Case Else
                sSQL = "drop table " & sName
                gdBase.Execute sSQL, dbFailOnError
        End Select
    End If
Next lcount

gdBase.TableDefs.Refresh    'Dabarefresh

'alle Tabellen, die mit TOPI beginnen 'außer
lAnzTable = gdBase.TableDefs.Count
For lcount = 0 To lAnzTable - 1
    sName = gdBase.TableDefs(lcount).name
    If Left(UCase$(sName), 4) = "TOPI" Then
    
        Select Case UCase$(sName)
'            Case "KUNDAUSLIEF", "KUNDAZE"
        
            Case Else
                sSQL = "drop table " & sName
                gdBase.Execute sSQL, dbFailOnError
        End Select
    End If
Next lcount

gdBase.TableDefs.Refresh    'Dabarefresh

'alle Tabellen, die mit AFCB beginnen 'außer AFCBUCH
lAnzTable = gdBase.TableDefs.Count
For lcount = 0 To lAnzTable - 1
    sName = gdBase.TableDefs(lcount).name
    If Left(UCase$(sName), 4) = "AFCB" Then
    
        If UCase$(sName) <> "AFCBUCH" Then
    
        
            sSQL = "drop table " & sName
            gdBase.Execute sSQL, dbFailOnError
            
        End If
        
        
        
    End If
Next lcount

gdBase.TableDefs.Refresh    'Dabarefresh

'alle Tabellen, die mit KUNDA beginnen 'außer KUNDAUSLIEF, KUNDAZE
lAnzTable = gdBase.TableDefs.Count
For lcount = 0 To lAnzTable - 1
    sName = gdBase.TableDefs(lcount).name
    If Left(UCase$(sName), 5) = "KUNDA" Then
    
        Select Case UCase$(sName)
            Case "KUNDAUSLIEF", "KUNDAZE", "KUNDA" & srechnertab
        
            Case Else
                sSQL = "drop table " & sName
                gdBase.Execute sSQL, dbFailOnError
        End Select
    End If
Next lcount






gdBase.TableDefs.Refresh    'Dabarefresh

'alle Tabellen, die mit NEUKU beginnen 'außer
lAnzTable = gdBase.TableDefs.Count
For lcount = 0 To lAnzTable - 1
    sName = gdBase.TableDefs(lcount).name
    If Left(UCase$(sName), 5) = "NEUKU" Then
    
        Select Case UCase$(sName)
            Case "KUNDAUSLIEF", "KUNDAZE"
        
            Case Else
                sSQL = "drop table " & sName
                gdBase.Execute sSQL, dbFailOnError
        End Select
    End If
Next lcount

gdBase.TableDefs.Refresh    'Dabarefresh

'alle Tabellen, die mit ETI beginnen 'außer ETIDRU, ETIDRULS, ETIPROT, ETIPROTS,ETISIC
lAnzTable = gdBase.TableDefs.Count
For lcount = 0 To lAnzTable - 1
    sName = gdBase.TableDefs(lcount).name
    If Left(UCase$(sName), 3) = "ETI" Then
    
        Select Case UCase$(sName)
            Case "ETIDRU", "ETIDRULS", "ETIPROT", "ETIPROTS", "ETISIC"
            
        
            Case Else
                sSQL = "drop table " & sName
                gdBase.Execute sSQL, dbFailOnError
        End Select
    
        
        
        
        
    End If
Next lcount


gdBase.TableDefs.Refresh    'Dabarefresh


'alle Tabellen, die mit MA beginnen
lAnzTable = gdBase.TableDefs.Count
For lcount = 0 To lAnzTable - 1
    sName = gdBase.TableDefs(lcount).name
    If Left(UCase$(sName), 2) = "MA" Then
    
        
        Select Case UCase$(sName)
            Case "MARKE"
            
            Case "MARKIERUNG"
            
            
        
            Case Else
                sSQL = "drop table " & sName
                gdBase.Execute sSQL, dbFailOnError
        End Select
        
    End If
Next lcount

gdBase.TableDefs.Refresh    'Dabarefresh

 'alle Tabellen, die mit MITKU beginnen
lAnzTable = gdBase.TableDefs.Count
For lcount = 0 To lAnzTable - 1
    sName = gdBase.TableDefs(lcount).name
    If Left(UCase$(sName), 5) = "MITKU" Then
    
        
        sSQL = "drop table " & sName
        gdBase.Execute sSQL, dbFailOnError
        
    End If
Next lcount

gdBase.TableDefs.Refresh    'Dabarefresh

'alle Tabellen, die mit MY beginnen
lAnzTable = gdBase.TableDefs.Count
For lcount = 0 To lAnzTable - 1
    sName = gdBase.TableDefs(lcount).name
    If Left(UCase$(sName), 2) = "MY" Then
    
        sSQL = "drop table " & sName
        gdBase.Execute sSQL, dbFailOnError
        
    End If
Next lcount

    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Or err.Number = 70 Or err.Number = 3033 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul11"
        Fehler.gsFunktion = "TempTabsDelete"
        Fehler.gsFehlertext = "Im Programmteil Datenbank ist ein Fehler aufgetreten."
    
        Fehlermeldung1
'        Resume Next
    End If
End Sub
Public Sub loeschNEWDTA(sAnfang As String, db As Database)
    On Error GoTo LOKAL_ERROR
    
    Dim lcount      As Long
    Dim lAnzTable   As Long
    Dim sSQL        As String
    Dim name        As String
    Dim ilen        As Integer
    
    Dim iMonat      As Integer
    
    iMonat = Month(Now)
    If iMonat = 1 Then iMonat = 13
    If iMonat = 12 Then iMonat = 14
    
    ilen = Len(sAnfang)
    
    lAnzTable = db.TableDefs.Count
    For lcount = 0 To lAnzTable - 1
        name = db.TableDefs(lcount).name
        
        
        If Left(UCase$(name), ilen) = Left(UCase$(sAnfang), ilen) Then
            If Trim(UCase(name)) = "DTA" Then
        
            Else
                If IsNumeric(Mid(Trim(name), 6, 2)) Then
                    If CInt(Mid(Trim(name), 6, 2)) < iMonat - 1 Then
                        sSQL = "drop table " & name
                        db.Execute sSQL, dbFailOnError
                    End If
                End If
                
            End If
        End If
    Next lcount

    Exit Sub
LOKAL_ERROR:
    If err.Number = 3376 Or err.Number = 3371 Or err.Number = 3043 Or err.Number = 91 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul2"
        Fehler.gsFunktion = "loeschNEWDTA"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Sub etidrukomp()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String

    loeschNEW "Te", gdBase
    sSQL = "Select * into te from ETIDRU"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    loeschNEW "ETIDRU", gdBase
    sSQL = "Select * into ETIDRU from Te"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    loeschNEW "Te", gdBase
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul11"
    Fehler.gsFunktion = "etidrukomp"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Private Sub BestrestPruefandDel()
    On Error GoTo LOKAL_ERROR
    
    Dim lcount      As Long
    Dim cDatname    As String
    Dim sSQL        As String
    
    loeschNEW "B" & srechnertab, gdBase
    
    sSQL = "Create Table B" & srechnertab
    sSQL = sSQL & " ( "
    sSQL = sSQL & " LINR LONG "
    sSQL = sSQL & ", ARTNR LONG "
    sSQL = sSQL & ", LEKPR SINGLE "
    sSQL = sSQL & ", BESTVOR long "
    sSQL = sSQL & ", DATEINAME TEXT(12) "
    sSQL = sSQL & ", BEST_DATUM DATETIME "
    sSQL = sSQL & ", UPD_DATUM DATETIME "
    sSQL = sSQL & " ) "
    gdBase.Execute sSQL, dbFailOnError
    
    Dim lAnzTable As Long
    Dim sQName As String
    
    gdBase.TableDefs.Refresh    'Dabarefresh
    lAnzTable = gdBase.TableDefs.Count
    For lcount = 0 To lAnzTable - 1
        sQName = gdBase.TableDefs(lcount).name
        If UCase(Left(sQName, 1)) = "Q" Then
            If Not Datendrin(sQName, gdBase) Then
                loeschNEW sQName, gdBase
            End If
        End If
    Next lcount
    
    gdBase.TableDefs.Refresh    'Dabarefresh
    lAnzTable = gdBase.TableDefs.Count
    For lcount = 0 To lAnzTable - 1
        sQName = gdBase.TableDefs(lcount).name
        If UCase(Left(sQName, 1)) = "Q" Then
        
            sSQL = "Insert into B" & srechnertab & " select * from bestrest where DATEINAME like '" & sQName & "*' "
            gdBase.Execute sSQL, dbFailOnError
        
        End If
    Next lcount
    
    loeschNEW "BESTREST", gdBase
    
    sSQL = "Select * into BESTREST from B" & srechnertab
    gdBase.Execute sSQL, dbFailOnError
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul11"
    Fehler.gsFunktion = "BestrestPruefandDel"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub Ballast_löschen()
On Error GoTo LOKAL_ERROR

    Dim sPfad As String
    
    Dim lAnz        As Long
    Dim lcount      As Long
    Dim lHeute      As Long
    Dim lDateiDatum As Long
    Dim cdatei      As String
    
    lHeute = Fix(Now)
    
    sPfad = gcDBPfad
    If Right$(sPfad, 1) <> "\" Then
        sPfad = sPfad & "\"
    End If
    
    frmWKL00.File3.Path = sPfad
    frmWKL00.File3.Pattern = "*.MDB"
    frmWKL00.File3.Refresh
    
    lAnz = frmWKL00.File3.ListCount
    For lcount = 0 To lAnz - 1
        cdatei = frmWKL00.File3.list(lcount)
        
        If UCase(cdatei) = "KISSDATA.MDB" Then
            'nicht löschen
        ElseIf UCase(cdatei) = "END.MDB" Then
            'nicht löschen
        ElseIf UCase(cdatei) = "SAFE.MDB" Then
            'nicht löschen
        ElseIf UCase(cdatei) = "KISSAPP.MDB" Then
            'nicht löschen
        ElseIf UCase(cdatei) = "ZENTAPP.MDB" Then
            'nicht löschen
        ElseIf UCase(cdatei) = "KLTMP.MDB" Then
            'nicht löschen
        ElseIf UCase(cdatei) = "LOKAL.MDB" Then
            'nicht löschen
        Else
            lDateiDatum = FileDateTime(sPfad & cdatei)
            If lHeute - lDateiDatum > 3 Then
                Kill sPfad & cdatei
            End If
        End If
    Next lcount
    
    'ABPRO
    sPfad = gcDBPfad
    If Right$(sPfad, 1) <> "\" Then
        sPfad = sPfad & "\"
    End If
    sPfad = sPfad & "ABPRO\"
    
    frmWKL00.File3.Path = sPfad
    frmWKL00.File3.Pattern = "*.*"
    frmWKL00.File3.Refresh
    
    lAnz = frmWKL00.File3.ListCount
    For lcount = 0 To lAnz - 1
        cdatei = frmWKL00.File3.list(lcount)
        
        lDateiDatum = FileDateTime(sPfad & cdatei)
        If lHeute - lDateiDatum > 365 Then
            Kill sPfad & cdatei
        End If
    Next lcount
    'ABPRO ENDE
    
    'ABPROSIC
    sPfad = gcDBPfad
    If Right$(sPfad, 1) <> "\" Then
        sPfad = sPfad & "\"
    End If
    sPfad = sPfad & "ABPROSIC\"
    
    frmWKL00.File3.Path = sPfad
    frmWKL00.File3.Pattern = "*.*"
    frmWKL00.File3.Refresh
    
    lAnz = frmWKL00.File3.ListCount
    For lcount = 0 To lAnz - 1
        cdatei = frmWKL00.File3.list(lcount)
        
        lDateiDatum = FileDateTime(sPfad & cdatei)
        If lHeute - lDateiDatum > 365 Then
            Kill sPfad & cdatei
        End If
    Next lcount
    'ABPROSIC ENDE
    
    'LPROTOK
    sPfad = gcDBPfad
    If Right$(sPfad, 1) <> "\" Then
        sPfad = sPfad & "\"
    End If
    sPfad = sPfad & "LPROTOK\"
    
    frmWKL00.File3.Path = sPfad
    frmWKL00.File3.Pattern = "WK*.RTF"
    frmWKL00.File3.Refresh
    
    lAnz = frmWKL00.File3.ListCount
    For lcount = 0 To lAnz - 1
        cdatei = frmWKL00.File3.list(lcount)
        
        lDateiDatum = FileDateTime(sPfad & cdatei)
        If lHeute - lDateiDatum > 30 Then
            Kill sPfad & cdatei
        End If
    Next lcount
    'LPROTOK ENDE
    
    'LPROTOK
    sPfad = gcDBPfad
    If Right$(sPfad, 1) <> "\" Then
        sPfad = sPfad & "\"
    End If
    sPfad = sPfad & "LPROTOK\"
    
    frmWKL00.File3.Path = sPfad
    frmWKL00.File3.Pattern = "FTP*.TXT"
    frmWKL00.File3.Refresh
    
    lAnz = frmWKL00.File3.ListCount
    For lcount = 0 To lAnz - 1
        cdatei = frmWKL00.File3.list(lcount)
        
        lDateiDatum = FileDateTime(sPfad & cdatei)
        If lHeute - lDateiDatum > 30 Then
            Kill sPfad & cdatei
        End If
    Next lcount
    'LPROTOK ENDE
    
    'ZPROTOK
    sPfad = gcDBPfad
    If Right$(sPfad, 1) <> "\" Then
        sPfad = sPfad & "\"
    End If
    sPfad = sPfad & "ZPROTOK\"
    
    frmWKL00.File3.Path = sPfad
    frmWKL00.File3.Pattern = "*.TXT"
    frmWKL00.File3.Refresh
    
    lAnz = frmWKL00.File3.ListCount
    For lcount = 0 To lAnz - 1
        cdatei = frmWKL00.File3.list(lcount)
        
        lDateiDatum = FileDateTime(sPfad & cdatei)
        If lHeute - lDateiDatum > 30 Then
            Kill sPfad & cdatei
        End If
    Next lcount
    'ZPROTOK ENDE
    
    'WVSIC
    sPfad = gcDBPfad
    If Right$(sPfad, 1) <> "\" Then
        sPfad = sPfad & "\"
    End If
    sPfad = sPfad & "WVSIC\"
    
    frmWKL00.File3.Path = sPfad
    frmWKL00.File3.Pattern = "*.MDB"
    frmWKL00.File3.Refresh
    
    lAnz = frmWKL00.File3.ListCount
    For lcount = 0 To lAnz - 1
        cdatei = frmWKL00.File3.list(lcount)
        
        lDateiDatum = FileDateTime(sPfad & cdatei)
        If lHeute - lDateiDatum > 30 Then
            Kill sPfad & cdatei
        End If
    Next lcount
    'WVSIC ENDE
    
    'ERRSIC
    sPfad = gcDBPfad
    If Right$(sPfad, 1) <> "\" Then
        sPfad = sPfad & "\"
    End If
    sPfad = sPfad & "ERRSIC\"
    
    frmWKL00.File3.Path = sPfad
    frmWKL00.File3.Pattern = "*.*"
    frmWKL00.File3.Refresh
    
    lAnz = frmWKL00.File3.ListCount
    For lcount = 0 To lAnz - 1
        cdatei = frmWKL00.File3.list(lcount)
        
        lDateiDatum = FileDateTime(sPfad & cdatei)
        If lHeute - lDateiDatum > 30 Then
            Kill sPfad & cdatei
        End If
    Next lcount
    'ERRSIC ENDE
    
'    'BEDPRO
'    sPfad = gcDBPfad
'    If Right$(sPfad, 1) <> "\" Then
'        sPfad = sPfad & "\"
'    End If
'    sPfad = sPfad & "BEDPRO\"
'
'    frmWKL00.File3.Path = sPfad
'    frmWKL00.File3.Pattern = "*.*"
'    frmWKL00.File3.Refresh
'
'    lAnz = frmWKL00.File3.ListCount
'    For lcount = 0 To lAnz - 1
'        cdatei = frmWKL00.File3.list(lcount)
'        Kill sPfad & cdatei
'    Next lcount
'    'BEDPRO ENDE
    
    'ABSCHLUS
    sPfad = gcDBPfad
    If Right$(sPfad, 1) <> "\" Then
        sPfad = sPfad & "\"
    End If
    sPfad = sPfad & "ABSCHLUS\"
    
    frmWKL00.File3.Path = sPfad
    frmWKL00.File3.Pattern = "F*.LZH"
    frmWKL00.File3.Refresh
    
    lAnz = frmWKL00.File3.ListCount
    For lcount = 0 To lAnz - 1
        cdatei = frmWKL00.File3.list(lcount)
        
        lDateiDatum = FileDateTime(sPfad & cdatei)
        If lHeute - lDateiDatum > 30 Then
            Kill sPfad & cdatei
        End If
    Next lcount
    'ABSCHLUS ENDE
    
    'ABSCHLUS - auch die .mdb
    sPfad = gcDBPfad
    If Right$(sPfad, 1) <> "\" Then
        sPfad = sPfad & "\"
    End If
    sPfad = sPfad & "ABSCHLUS\"
    
    frmWKL00.File3.Path = sPfad
    frmWKL00.File3.Pattern = "F*.MDB"
    frmWKL00.File3.Refresh
    
    lAnz = frmWKL00.File3.ListCount
    For lcount = 0 To lAnz - 1
        cdatei = frmWKL00.File3.list(lcount)
        
        lDateiDatum = FileDateTime(sPfad & cdatei)
        If lHeute - lDateiDatum > 30 Then
            Kill sPfad & cdatei
        End If
    Next lcount
    'ABSCHLUS ENDE
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul11"
        Fehler.gsFunktion = "Ballast_löschen"
        Fehler.gsFehlertext = "Im Programmteil Datenbank ist ein Fehler aufgetreten."
    
        Fehlermeldung1
    End If
End Sub
Public Sub Container_Lieferanten_aktualisieren()
On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rsrs As DAO.Recordset
    Dim rsUn As DAO.Recordset
    Dim cHauptLinr As String
    Dim cUnterLinr As String
    
    If NewTableSuchenDBKombi("ueberli", gdBase) = True Then
    
        sSQL = "Select distinct(olinr) from ueberli"
        Set rsrs = gdBase.OpenRecordset(sSQL)
        If Not rsrs.EOF Then
        
            rsrs.MoveFirst
            Do While Not rsrs.EOF
            
                cHauptLinr = "0"
                If Not IsNull(rsrs!oLINR) Then
                    cHauptLinr = rsrs!oLINR
                End If
                
                If cHauptLinr <> "300500" Then
                
                    sSQL = "Delete from ARTLIEF where linr = " & cHauptLinr & " "
                    gdBase.Execute sSQL, dbFailOnError
                    
                    sSQL = "Select linr from ueberli where olinr = " & cHauptLinr & " "
                    Set rsUn = gdBase.OpenRecordset(sSQL)
                    If Not rsUn.EOF Then
                    
                        
                        rsUn.MoveFirst
                        Do While Not rsUn.EOF
                        
                            cUnterLinr = "0"
                            If Not IsNull(rsUn!linr) Then
                                cUnterLinr = rsUn!linr
                            End If
                            
                            sSQL = "Insert into ArtLief Select ArtNr, " & cHauptLinr & " as LiNr, LibesNr, MINMEN, LEKPR"
                            sSQL = sSQL & " , RKZ, EXDAT from ARTLIEF where LINR = " & cUnterLinr
    
                            sSQL = sSQL & " and ARTNR NOT IN "
                            sSQL = sSQL & " (select Artlief.ARTNR from ARTLIEF"
                            sSQL = sSQL & " where "
                            sSQL = sSQL & " Artlief.linr = " & cHauptLinr & " )"
                            gdBase.Execute sSQL, dbFailOnError
                
            
                        rsUn.MoveNext
                        Loop
            
                
                    End If
                    rsUn.Close: Set rsUn = Nothing
                End If

            rsrs.MoveNext
            Loop

    
        End If
        rsrs.Close: Set rsrs = Nothing
    
    
'            Dim tmpdt As DataTable = Me.dbclass._SQL_DataReader("select distinct(hauptlinr) from lieferanten_zuordnung ")
'
'            Dim cHauptLinr As String = "0"
'
'            For Each sRow As DataRow In tmpdt.Rows
'                cHauptLinr = (sRow.Item("Hauptlinr")).ToString
'
'                sSQL = "Delete from ARTLIEF where linr = " & cHauptLinr & " "
'                Me.dbclass._SQL_execute(sSQL)
'
'                Dim tmpdtU As DataTable = Me.dbclass._SQL_DataReader("select Unterlinr from lieferanten_zuordnung where Hauptlinr = " & cHauptLinr)
'
'                Dim cUnterLinr As String = "0"
'
'                For Each sRowU As DataRow In tmpdtU.Rows
'                    cUnterLinr = (sRowU.Item("UnterLinr")).ToString
'
'                    sSQL = "Insert into ArtLief Select ArtNr, " & cHauptLinr & " as LiNr, LibesNr, MM, LEKPR, Linie, Spanne, 'False' as HerstLinR,"
'                    sSQL = sSQL & " FEKPR, RKZ, EXDAT from ARTLIEF where LINR = " & cUnterLinr
'
'                    sSQL = sSQL & " and ARTNR NOT IN "
'                    sSQL = sSQL & " (select Artlief.ARTNR from ARTLIEF"
'                    sSQL = sSQL & " where "
'                    sSQL = sSQL & " Artlief.linr = " & cHauptLinr & " )"
'                    Me.dbclass._SQL_execute(sSQL)
'
'                Next
'
'            Next
    End If
    
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul11"
    Fehler.gsFunktion = "Container_Lieferanten_aktualisieren"
    Fehler.gsFehlertext = "Im Programmteil Datenbank ist ein Fehler aufgetreten."

    Fehlermeldung1
    
End Sub
Public Sub erstelle_GDPdU()
On Error GoTo LOKAL_ERROR

    Dim sSQL            As String
    Dim cPfad           As String
    Dim GDPdU_DB        As Database
    Dim cDatum          As String

    Screen.MousePointer = 11
    
    cDatum = "20.03.2004"
    
    cPfad = gcDBPfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "GDPdU\GDPdU.MDB"
    
    If FileExists(cPfad) = False Then
        Kill cPfad
    
        Set GDPdU_DB = CreateDatabase(cPfad, dbLangGeneral, dbVersion40)
        GDPdU_DB.Close
        
        Set GDPdU_DB = OpenDatabase(cPfad, True, False)
        GDPdU_DB.NewPassword "", gsGDPdU_Passwort
        GDPdU_DB.Close
        
        
        Set GDPdU_DB = OpenDatabase(cPfad, False, False, "MS Access;PWD=" & gsGDPdU_Passwort)
        
        
        CreateTableT2 "GDPDU_STAND", GDPdU_DB
        
        sSQL = "Insert into GDPDU_STAND (Datum,Zeit) values ( "
        sSQL = sSQL & " '" & DateValue(cDatum) & "','" & TimeValue(Now) & "') "
        GDPdU_DB.Execute sSQL
    End If

    'Vorarbeit geleistet

    
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul11"
        Fehler.gsFunktion = "erstelle_GDPdU"
        Fehler.gsFehlertext = "Im Programmteil Datenbank ist ein Fehler aufgetreten."
    
        Fehlermeldung1
    End If
End Sub
Public Sub erstelle_KASSBON()
On Error GoTo LOKAL_ERROR

    Dim sSQL            As String
    Dim cPfad           As String
    Dim KASSBON_DB      As Database
    
    Screen.MousePointer = 11
    
    cPfad = gcDBPfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "GDPdU\KASSBON.MDB"
    
    If FileExists(cPfad) = False Then
        Kill cPfad
    
        Set KASSBON_DB = CreateDatabase(cPfad, dbLangGeneral, dbVersion40)
        KASSBON_DB.Close
        
        Set KASSBON_DB = OpenDatabase(cPfad, True, False)
        KASSBON_DB.NewPassword "", gsKASSBON_Passwort
        KASSBON_DB.Close
        
        
        Set KASSBON_DB = OpenDatabase(cPfad, False, False, "MS Access;PWD=" & gsKASSBON_Passwort)
        
        loeschNEW "KASSBOND", KASSBON_DB
        TransferTab gdBase, cPfad, "KASSBOND"
        
        KASSBON_DB.Close
        
        
    End If

    'Vorarbeit geleistet

    
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul11"
        Fehler.gsFunktion = "erstelle_KASSBON"
        Fehler.gsFehlertext = "Im Programmteil Datenbank ist ein Fehler aufgetreten."
    
        Fehlermeldung1
    End If
End Sub
Public Sub GDPdU_schreiben()
On Error GoTo LOKAL_ERROR

    Dim sSQL            As String
    Dim sRechner        As String
    Dim cPfad           As String
    Dim GDPdU_DB        As Database
    Dim rsrs            As DAO.Recordset
    Dim dateStand       As Date
    Dim GDPdU_Handlungsbedarf        As Boolean
    
    Screen.MousePointer = 11
    
    GDPdU_Handlungsbedarf = False
    
    cPfad = gcDBPfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    sRechner = rechnername
    
    cPfad = cPfad & "GDPdU\GDPdU.MDB"
    
    Set GDPdU_DB = OpenDatabase(cPfad, True, False, "MS Access;PWD=" & gsGDPdU_Passwort)
    
    '1. Stand - ermitteln
    Set rsrs = GDPdU_DB.OpenRecordset("select DATUM from GDPDU_STAND")
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!Datum) Then
            dateStand = rsrs!Datum
'            dateStand = DateValue("20.03.12")
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    If dateStand < DateValue(Now) Then
    
        'wenn der GDPdU-Stand kleiner als das gestrige Datum ist, dann gibt es was zu tun
        GDPdU_Handlungsbedarf = True
    
    End If
    
    'ist Handlungsbedarf dann folgende Punkte abarbeiten
    If GDPdU_Handlungsbedarf = True Then
        'Kassjour
        
        cPfad = gcDBPfad
        If Right$(cPfad, 1) <> "\" Then
            cPfad = cPfad & "\"
        End If
        
        If Not NewTableSuchenDBKombi("KASSJOUR", GDPdU_DB) Then
        
            sSQL = "Create Table KASSJOUR"
            sSQL = sSQL & "("
            sSQL = sSQL & " ARTNR LONG "
            sSQL = sSQL & ", BEZEICH TEXT(35) "
            sSQL = sSQL & ", MENGE INTEGER "
            sSQL = sSQL & ", PREIS SINGLE "
            sSQL = sSQL & ", ADATE DATETIME "
            sSQL = sSQL & ", AZEIT Text(8) "
            sSQL = sSQL & ", KUNDNR LONG "
            sSQL = sSQL & ", FILIALE BYTE "
            sSQL = sSQL & ", KASNUM BYTE "
            sSQL = sSQL & ", LINR long"
            sSQL = sSQL & ", LPZ INTEGER"
            sSQL = sSQL & ", AGN Long "
            sSQL = sSQL & ", EAN Text(13)"
            sSQL = sSQL & ", MWST Text(1)"
            sSQL = sSQL & ", EKPR SINGLE "
            sSQL = sSQL & ", VKPR SINGLE "
            sSQL = sSQL & ", MOPREIS SINGLE "
            sSQL = sSQL & ", BELEGNR INTEGER "
            sSQL = sSQL & ", BEST1 INTEGER "
            sSQL = sSQL & ", RABKENN Text(1)"
            sSQL = sSQL & ", KK_ART Text(2)"
            sSQL = sSQL & ", BEDIENER integer "
            sSQL = sSQL & ", UMS_OK Text(1)"
            sSQL = sSQL & ", ZBONNR integer"
            sSQL = sSQL & ", ABOK BIT"
            sSQL = sSQL & ")"
            GDPdU_DB.Execute sSQL, dbFailOnError
            
        End If
    
        sSQL = "Insert into kassjour Select "
        sSQL = sSQL & " artnr "
        sSQL = sSQL & " ,bezeich "
        sSQL = sSQL & " ,Menge "
        sSQL = sSQL & " ,Preis "
        sSQL = sSQL & ", ADATE  "
        sSQL = sSQL & ", AZEIT "
        sSQL = sSQL & ", KUNDNR  "
        sSQL = sSQL & ", FILIALE  "
        sSQL = sSQL & ", KASNUM  "
        sSQL = sSQL & ", LINR "
        sSQL = sSQL & ", LPZ "
        sSQL = sSQL & ", AGN  "
        sSQL = sSQL & ", EAN "
        sSQL = sSQL & ", MWST "
        sSQL = sSQL & ", EKPR "
        sSQL = sSQL & ", VKPR "
        sSQL = sSQL & ", MOPREIS  "
        sSQL = sSQL & ", BELEGNR  "
        sSQL = sSQL & ", BEST1 "
        sSQL = sSQL & ", RABKENN "
        sSQL = sSQL & ", KK_ART "
        sSQL = sSQL & ", BEDIENER  "
        sSQL = sSQL & ", UMS_OK "
        sSQL = sSQL & ", ZBONNR "
        sSQL = sSQL & ", ABOK "
    
        sSQL = sSQL & " from [;DATABASE=" & cPfad & "KISSDATA.MDB;pwd=" & gsPasswort & "].Kassjour "
        sSQL = sSQL & " where adate >= " & CLng(dateStand)
        sSQL = sSQL & " and adate < " & CLng(DateValue(Now))
        GDPdU_DB.Execute sSQL, dbFailOnError
        
        'Preisänderungen
        
        If Not NewTableSuchenDBKombi("KVKPR1PROT", GDPdU_DB) Then
        
            sSQL = "Create Table KVKPR1PROT "
            sSQL = sSQL & "( ARTNR LONG"
            sSQL = sSQL & ", KVKPR1 SINGLE "
            sSQL = sSQL & ", BEDIENER LONG"
            sSQL = sSQL & ", SYNSTATUS TEXT(1) "
            sSQL = sSQL & ", AENART TEXT(20)"
            sSQL = sSQL & ", FILIALE BYTE"
            sSQL = sSQL & ", LASTDATE DATETIME"
            sSQL = sSQL & ", LASTTIME TEXT(10) "
            sSQL = sSQL & ", SENDOK BIT "
            sSQL = sSQL & ") "
            GDPdU_DB.Execute sSQL, dbFailOnError
            
        End If
    
        sSQL = "Insert into KVKPR1PROT Select "
        sSQL = sSQL & " ARTNR "
        sSQL = sSQL & ", KVKPR1  "
        sSQL = sSQL & ", BEDIENER "
        sSQL = sSQL & ", SYNSTATUS  "
        sSQL = sSQL & ", AENART "
        sSQL = sSQL & ", FILIALE "
        sSQL = sSQL & ", LASTDATE "
        sSQL = sSQL & ", LASTTIME  "
        sSQL = sSQL & ", SENDOK  "
    
        sSQL = sSQL & " from [;DATABASE=" & cPfad & "KISSDATA.MDB;pwd=" & gsPasswort & "].KVKPR1PROT "
        sSQL = sSQL & " where LASTDATE >= " & CLng(dateStand)
        sSQL = sSQL & " and LASTDATE < " & CLng(DateValue(Now))
        GDPdU_DB.Execute sSQL, dbFailOnError
        
        'Bestandsveränderungen
        
        If Not NewTableSuchenDBKombi("BESTPROT", GDPdU_DB) Then
            sSQL = "Create Table BESTPROT "
            sSQL = sSQL & "( ARTNR LONG"
            sSQL = sSQL & ", UMENGE LONG "
            sSQL = sSQL & ", NEWBEST LONG"
            sSQL = sSQL & ", OLDBEST LONG"
            sSQL = sSQL & ", BEDIENER LONG"
            sSQL = sSQL & ", SYNSTATUS TEXT(1) "
            sSQL = sSQL & ", AENART TEXT(20)"
            sSQL = sSQL & ", AENGRUND TEXT(20)"
            sSQL = sSQL & ", FILIALE BYTE"
            sSQL = sSQL & ", LASTDATE DATETIME"
            sSQL = sSQL & ", LASTTIME TEXT(10) "
            sSQL = sSQL & ", SENDOK BIT "
            sSQL = sSQL & " ) "
            GDPdU_DB.Execute sSQL, dbFailOnError
        End If
    
        sSQL = "Insert into BESTPROT Select "
        sSQL = sSQL & " ARTNR "
        sSQL = sSQL & ", UMENGE "
        sSQL = sSQL & ", NEWBEST "
        sSQL = sSQL & ", OLDBEST "
        sSQL = sSQL & ", BEDIENER "
        sSQL = sSQL & ", SYNSTATUS "
        sSQL = sSQL & ", AENART "
        sSQL = sSQL & ", AENGRUND "
        sSQL = sSQL & ", FILIALE "
        sSQL = sSQL & ", LASTDATE "
        sSQL = sSQL & ", LASTTIME "
        sSQL = sSQL & ", SENDOK  "
    
        sSQL = sSQL & " from [;DATABASE=" & cPfad & "KISSDATA.MDB;pwd=" & gsPasswort & "].BESTPROT "
        sSQL = sSQL & " where LASTDATE >= " & CLng(dateStand)
        sSQL = sSQL & " and LASTDATE < " & CLng(DateValue(Now))
        GDPdU_DB.Execute sSQL, dbFailOnError
        
        'Kassenbons
        
        If Not NewTableSuchenDBKombi("KASSBOND", GDPdU_DB) Then
            
            
            sSQL = "Create Table KASSBOND ("
            sSQL = sSQL & " DATUM DATETIME"
            sSQL = sSQL & ", KASNUM double"
            sSQL = sSQL & ", BONNR double"
            sSQL = sSQL & ", BETRAG double"
            sSQL = sSQL & ", UHRZEIT TEXT(8)"
            sSQL = sSQL & ", BONTEXT memo"
            sSQL = sSQL & ", Filiale byte"
            sSQL = sSQL & ", SENDOK BIT"
            sSQL = sSQL & ", KK_ART TEXT(2) "
            sSQL = sSQL & ", KUNDNR LONG"
            sSQL = sSQL & " )"
            
            
            GDPdU_DB.Execute sSQL, dbFailOnError
        End If
    
        sSQL = "Insert into KASSBOND Select "
        sSQL = sSQL & " DATUM "
        sSQL = sSQL & ", KASNUM "
        sSQL = sSQL & ", BONNR "
        sSQL = sSQL & ", BETRAG "
        sSQL = sSQL & ", UHRZEIT "
        sSQL = sSQL & ", BONTEXT "
        sSQL = sSQL & ", Filiale "
        sSQL = sSQL & ", SENDOK "
        sSQL = sSQL & ", KK_ART  "
        sSQL = sSQL & ", KUNDNR "
    
        sSQL = sSQL & " from [;DATABASE=" & cPfad & "KISSDATA.MDB;pwd=" & gsPasswort & "].KASSBOND "
        sSQL = sSQL & " where DATUM >= " & CLng(dateStand)
        sSQL = sSQL & " and DATUM < " & CLng(DateValue(Now))
        GDPdU_DB.Execute sSQL, dbFailOnError
        
        'zum Schluss stand eintragen
        sSQL = "Update GDPDU_STAND set "
        sSQL = sSQL & " Datum = '" & DateValue(Now) & "'"
        GDPdU_DB.Execute sSQL
    End If
   
    GDPdU_DB.Close
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul11"
        Fehler.gsFunktion = "GDPdU_schreiben"
        Fehler.gsFehlertext = "Im Programmteil Datenbank ist ein Fehler aufgetreten."
    
        Fehlermeldung1
    End If
End Sub
Public Function dbApp_Compri(sDB As String) As Boolean
On Error GoTo LOKAL_ERROR

    dbApp_Compri = False
    
    gdApp.Close
    
    DBEngine.CompactDatabase App.Path & "\" & sDB, App.Path & "\" & "kltmp.mdb", dbLangGeneral
    
    Pause 1
    Kill App.Path & "\" & sDB
    
    DBEngine.CompactDatabase App.Path & "\" & "kltmp.mdb", App.Path & "\" & sDB, dbLangGeneral
    
    Pause 1
    Kill App.Path & "\" & "kltmp.mdb"
    
    dbApp_Compri = True
    
    Set gdApp = OpenDatabase(App.Path & "\" & sDB)
    
Exit Function
LOKAL_ERROR:
    If err.Number = 3420 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul11"
        Fehler.gsFunktion = "dbApp_Compri"
        Fehler.gsFehlertext = "Im Programmteil Datenbank ist ein Fehler aufgetreten."
    
        Fehlermeldung1
    End If
End Function
Public Function dbGDPDU_Compri(sDB As String, labelx As Label) As Boolean
On Error GoTo LOKAL_ERROR

    Dim cPfad As String
    
    cPfad = gcDBPfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "GDPdU\"

    dbGDPDU_Compri = False
    
    labelx.ForeColor = vbBlack
    labelx.Caption = "Teil 1, bitte warten..."
    labelx.Refresh
    
    DBEngine.CompactDatabase cPfad & "\" & sDB, cPfad & "\" & "GDPDUkltmp.mdb", dbLangGeneral, , ";pwd=" & gsGDPdU_Passwort
    
    Pause 1
    Kill cPfad & "\" & sDB
    
    labelx.ForeColor = vbBlack
    labelx.Caption = "Teil 2, bitte warten..."
    labelx.Refresh
    
    DBEngine.CompactDatabase cPfad & "\" & "GDPDUkltmp.mdb", cPfad & "\" & sDB, dbLangGeneral, , ";pwd=" & gsGDPdU_Passwort
    
    Pause 1
    Kill cPfad & "\" & "GDPDUkltmp.mdb"
    
    dbGDPDU_Compri = True
    
   
    
Exit Function
LOKAL_ERROR:
    If err.Number = 3420 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul11"
        Fehler.gsFunktion = "dbGDPDU_Compri"
        Fehler.gsFehlertext = "Im Programmteil Datenbank ist ein Fehler aufgetreten."
    
        Fehlermeldung1
    End If
End Function
Public Sub GDPDU_GLAGER_KLEINHALTEN(labelx As Label)
On Error GoTo LOKAL_ERROR
    
    Dim cSQL        As String
    Dim rsrs        As Recordset
    Dim ctemp       As String
    Dim GDPdU_DB    As Database
    Dim cPfad       As String
    
    Screen.MousePointer = 11
    
    cPfad = gcDBPfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "GDPdU\GDPdU.MDB"
    
    Set GDPdU_DB = OpenDatabase(cPfad, False, False, "MS Access;PWD=" & gsGDPdU_Passwort)
    
    Dim lMAXDatumUebersicht     As Long
    Dim lMAXDatumGLAGER         As Long
    Dim rsDat                   As DAO.Recordset
    
    If NewTableSuchenDBKombi("GLAGER_GDPdU", GDPdU_DB) Then
    
        If Datendrin("GLAGER_GDPdU", GDPdU_DB) Then
            CheckIndex "GLAGER_GDPdU", "DATUM", "", GDPdU_DB
            CheckIndex "GLAGER_GDPdU", "BESTAND", "", GDPdU_DB
            
            If NewTableSuchenDBKombi("GLAGER_UEBERSICHT", GDPdU_DB) = False Then
                'dann erstelle eine
                cSQL = "select distinct(datum) as disdatum ,sum(bestand) as mBestand into GLAGER_UEBERSICHT from GLAGER_GDPdU group by datum "
                GDPdU_DB.Execute cSQL, dbFailOnError
            Else
                'füge neue Sätze an
                lMAXDatumUebersicht = 0
                lMAXDatumGLAGER = 0
        
                cSQL = "Select Max(disdatum) as Maxdat from GLAGER_UEBERSICHT"
                Set rsrs = GDPdU_DB.OpenRecordset(cSQL)
                If Not rsrs.EOF Then
                    If Not IsNull(rsrs!Maxdat) Then
                        lMAXDatumUebersicht = rsrs!Maxdat
                    End If
                End If
                rsrs.Close: Set rsrs = Nothing
                
                cSQL = "Select Max(datum) as Maxdat from GLAGER_GDPdU"
                Set rsrs = GDPdU_DB.OpenRecordset(cSQL)
                If Not rsrs.EOF Then
                    If Not IsNull(rsrs!Maxdat) Then
                        lMAXDatumGLAGER = rsrs!Maxdat
                    End If
                End If
                rsrs.Close: Set rsrs = Nothing
        
                If lMAXDatumGLAGER > lMAXDatumUebersicht Then
                    'dann gibt es etwas anzufügen
                    cSQL = "Insert into GLAGER_UEBERSICHT select distinct(datum) as disdatum ,sum(bestand) as mBestand "
                    cSQL = cSQL & " from GLAGER_GDPdU where datum > " & lMAXDatumUebersicht & " group by datum "
                    GDPdU_DB.Execute cSQL, dbFailOnError
                End If
            End If
            
            Dim lMinJahr As Long
            lMinJahr = 0
            'kleinstes Jahr, kleinster Monat
            cSQL = "Select MIN(year(disdatum)) as MINJAHR from GLAGER_UEBERSICHT"
            cSQL = cSQL & " where day(disdatum) between 4 and 25 "
            Set rsrs = GDPdU_DB.OpenRecordset(cSQL)
            If Not rsrs.EOF Then
            
                If IsNull(rsrs!MINJAHR) Then
                      lMinJahr = 0
                 Else
                       lMinJahr = rsrs!MINJAHR
                End If
                
            End If
            rsrs.Close: Set rsrs = Nothing
            
            Dim lMinMonat As Long
            lMinMonat = 0
            cSQL = "Select MIN(month(disdatum)) as MINMonat from GLAGER_UEBERSICHT where year(disdatum) = " & lMinJahr
            cSQL = cSQL & " and day(disdatum) between 4 and 25 "
            Set rsrs = GDPdU_DB.OpenRecordset(cSQL)
            If Not rsrs.EOF Then
            
                If IsNull(rsrs!MINmonat) Then
                      lMinMonat = 0
                 Else
                       lMinMonat = rsrs!MINmonat
                End If
                 
                
            End If
            rsrs.Close: Set rsrs = Nothing
            
            If lMinJahr > 0 And lMinMonat > 0 Then
                If lMinJahr < Year(Now) - 1 Then
                
                    labelx.ForeColor = vbBlack
                    labelx.Caption = "löschen, bitte warten..."
                    labelx.Refresh
                    
                    cSQL = "Delete from GLAGER_GDPDU "
                    cSQL = cSQL & " where day(datum) between 4 and 25 "
                    cSQL = cSQL & " and year(datum) = " & lMinJahr
                    cSQL = cSQL & " and month(datum) = " & lMinMonat
                    GDPdU_DB.Execute cSQL, dbFailOnError
                    
                    labelx.ForeColor = vbBlack
                    labelx.Caption = "aufräumen, bitte warten..."
                    labelx.Refresh
                
                    cSQL = "Delete from GLAGER_UEBERSICHT "
                    cSQL = cSQL & " where day(disdatum) between 4 and 25 "
                    cSQL = cSQL & " and year(disdatum) = " & lMinJahr
                    cSQL = cSQL & " and month(disdatum) = " & lMinMonat
                    GDPdU_DB.Execute cSQL, dbFailOnError
                    
                End If
            End If
        End If
    End If

    GDPdU_DB.Close
    
Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul11"
        Fehler.gsFunktion = "GDPDU_GLAGER_KLEINHALTEN"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1
'        Resume Next
    End If
End Sub
Public Function dbKASSBON_Compri(sDB As String) As Boolean
On Error GoTo LOKAL_ERROR

    Dim cPfad As String
    
    cPfad = gcDBPfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "GDPdU\"

    dbKASSBON_Compri = False
    
    DBEngine.CompactDatabase cPfad & "\" & sDB, cPfad & "\" & "KASSBONkltmp.mdb", dbLangGeneral, , ";pwd=" & gsKASSBON_Passwort
    
    Pause 1
    Kill cPfad & "\" & sDB
    
    DBEngine.CompactDatabase cPfad & "\" & "KASSBONkltmp.mdb", cPfad & "\" & sDB, dbLangGeneral, , ";pwd=" & gsKASSBON_Passwort
    
    Pause 1
    Kill cPfad & "\" & "KASSBONkltmp.mdb"
    
    dbKASSBON_Compri = True
    
   
    
Exit Function
LOKAL_ERROR:
    If err.Number = 3420 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul11"
        Fehler.gsFunktion = "dbKASSBON_Compri"
        Fehler.gsFehlertext = "Im Programmteil Datenbank ist ein Fehler aufgetreten."
    
        Fehlermeldung1
    End If
End Function
Public Function db_Compri_ohneAnsicht(sDB As String) As Boolean
On Error GoTo LOKAL_ERROR

    db_Compri_ohneAnsicht = False
    
    Dim lRet As Long
    Dim lfail As Long
    Dim cNewpath As String
    Dim cOldpath As String
    Dim cPfad As String
    
    
    cPfad = gcDBPfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    gdBase.Close
    
    Kill cPfad & "K_SICH.mdb"
    
    cOldpath = cPfad
    cOldpath = ShortPath(cOldpath)
    cOldpath = cOldpath & "kissdata.mdb"
    
    cNewpath = cPfad
    cNewpath = ShortPath(cNewpath)
    cNewpath = cNewpath & "K_SICH.mdb"
    lRet = CopyFile(cOldpath, cNewpath, lfail)
    
    If lRet = 0 Then
        Screen.MousePointer = 0
        lblx.Caption = "Abbruch"
        lblx.Refresh
        Exit Function
    End If
    
    
    Pause 1
    Kill cPfad & sDB
    
    DBEngine.CompactDatabase cPfad & "K_SICH.mdb", cPfad & sDB, , , ";pwd=" & gsPasswort
    
    
    
    db_Compri_ohneAnsicht = True
    
    Set gdBase = OpenDatabase(cPfad & sDB, False, False, "MS Access;PWD=" & gsPasswort)
    
    
Exit Function
LOKAL_ERROR:
    If err.Number = 3420 Then
        Resume Next
    ElseIf err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul11"
        Fehler.gsFunktion = "db_Compri_ohneAnsicht"
        Fehler.gsFehlertext = "Im Programmteil Datenbank ist ein Fehler aufgetreten."
    
        Fehlermeldung1
    End If
End Function
Public Function db_Sichern(sPfad As String, sDB As String, lab As Label, txtStatus As TextBox, labglo As Label) As Boolean
On Error GoTo LOKAL_ERROR

    Dim oldpath     As String
    Dim newpath     As String
    Dim lfail       As Long
    Dim lRet        As Long
    Dim lHeute      As Long
    Dim lGestern    As Long
    Dim cPfad       As String
    Dim j           As Long

    Screen.MousePointer = 11
    
    db_Sichern = False
    
    schreibeProtokollDaba ("Sicherung gestartet")
    
    txtStatus.Text = 0
    lab.Caption = "": lab.Refresh

    labglo.ForeColor = vbRed
    labglo.Caption = "Datenbank wird gesichert, nicht ausschalten!!!"
    labglo.Refresh

    lHeute = Fix(Now)
    lGestern = lHeute - 14

    For j = lGestern To 38000 Step -1
        Kill sPfad & "DABASIC\KD" & CStr(j) & ".LZH"
    Next j
    
    Kill sPfad & "DABASIC1\KD.LZH"

    If Not FileExists(sPfad & "DABASIC\KD" & CStr(lHeute) & ".LZH") Then
        zipDllcheck
        Zip_Files "", sPfad & "KISSDATA.MDB", sPfad & "KISIC.LZH", txtStatus
    
        oldpath = sPfad & "KISIC.LZH"
        newpath = sPfad & "DABASIC\KD" & CStr(lHeute) & ".LZH"
        lRet = CopyFile(oldpath, newpath, lfail)
        
        labglo.ForeColor = vbRed
        labglo.Caption = "Die Sicherung wird kopiert, bitte warten..."
        labglo.Refresh
        
        oldpath = sPfad & "KISIC.LZH"
        newpath = sPfad & "DABASIC1\KD.LZH"
        lRet = CopyFile(oldpath, newpath, lfail)

        
        
        If lRet <> 0 Then
            schreibeProtokollDaba ("Erfolg Sicherung")
            labglo.ForeColor = vbBlack
            labglo.Caption = "Fertig"
            labglo.Refresh
            
            db_Sichern = True
        Else
            schreibeProtokollDaba ("Fehler Sicherung")
            labglo.ForeColor = vbRed
            labglo.Caption = "Fehler - Die Sicherung wurde nicht durchgeführt."
            labglo.Refresh
            
            db_Sichern = False
        End If
    Else
        schreibeProtokollDaba ("Sicherung wurde heute schon erzeugt")
        labglo.ForeColor = vbRed
        labglo.Caption = "Die Sicherung der Datenbank wurde heute schon erzeugt - Fertig"
        labglo.Refresh
       
        db_Sichern = True
    End If
    
    Kill gcDBPfad & "KISIC.LZH"
    
    Screen.MousePointer = 0
    
Exit Function
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul11"
        Fehler.gsFunktion = "db_Sichern"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
        
        Fehlermeldung1
        
    End If
End Function
Public Function ErmlzVKausApp(cART As String) As Date
    On Error GoTo LOKAL_ERROR
    
    ErmlzVKausApp = 0
    
    Dim cSQL As String
    Dim rsINB As Recordset
    
    cSQL = "Select max(adate) as maxdate from Kassjour where ARTNR = " & cART & " "
    Set rsINB = gdApp.OpenRecordset(cSQL)
    If Not rsINB.EOF Then
        If Not IsNull(rsINB!MaxDate) Then
            ErmlzVKausApp = rsINB!MaxDate
        End If
    End If
    rsINB.Close: Set rsINB = Nothing
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul11"
    Fehler.gsFunktion = "ErmlzVKausApp"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ErmlzZugang(cART As String) As Date
    On Error GoTo LOKAL_ERROR
    
    ErmlzZugang = DateValue("01.01.1980")
    
    Dim cSQL As String
    Dim rsINB As Recordset
    
    cSQL = "Select max(adate) as maxdate from Zugang where ARTNR = " & cART & " "
    Set rsINB = gdBase.OpenRecordset(cSQL)
    If Not rsINB.EOF Then
        If Not IsNull(rsINB!MaxDate) Then
            ErmlzZugang = rsINB!MaxDate
        End If
    End If
    rsINB.Close: Set rsINB = Nothing
    
    'check auch umlager
    
    If CInt(gcFilNr) > 0 Then
    
        Dim lzUmlager As Date
        lzUmlager = DateValue("01.01.1980")
    
        cSQL = "Select max(adate) as maxdate from Umlager where ARTNR = " & cART & " "
        Set rsINB = gdBase.OpenRecordset(cSQL)
        If Not rsINB.EOF Then
            If Not IsNull(rsINB!MaxDate) Then
                lzUmlager = rsINB!MaxDate
            End If
        End If
        rsINB.Close: Set rsINB = Nothing
        
        If lzUmlager <> DateValue("01.01.1980") Then
            If lzUmlager > ErmlzZugang Then
                ErmlzZugang = lzUmlager
            End If
        End If
        
        
        If ErmlzZugang = DateValue("01.01.1980") And lzUmlager <> DateValue("01.01.1980") Then
            
            ErmlzZugang = lzUmlager
            
        End If
        
        
        
    End If
    
    
    
    
    
    
    
    
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul11"
    Fehler.gsFunktion = "ErmlzZugang"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Function

