Attribute VB_Name = "Modul7"
Option Explicit
Global dabalokal        As Database
Public Sub csvImport(sTab As String, dbZiel As Database, sPfadundDatei As String, lblAnzeige As Label)
On Error GoTo LOKAL_ERROR

    Dim lPosEnde As Long
    Dim cEinzelsatz As String
    Dim lLenfil As Long
    Dim lposSemi As Long
    Dim lposSemiEnde As Long
    Dim cWert As String
    Dim lfnr1 As Long
    Dim cPreis As String
    Dim lPos    As String
    Dim rsrs    As Recordset
    Dim iFileNr As Integer
    Dim cSatz1          As String
    Dim dWert As Double
    Dim lcount As Long
    
    loeschNEW sTab, dbZiel
    
    Select Case UCase(sTab)
        Case "MASTER"
            loeschNEW "MASTER", dbZiel
            CreateTableT2 "MASTER", dbZiel
            
            lPos = 1
            lPosEnde = 1
            lposSemiEnde = 1
            
            Set rsrs = dbZiel.OpenRecordset("MASTER")
    
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
                    anzeige "normal", CStr(lcount), lblAnzeige
                    lPosEnde = InStr(lPos, cSatz1, vbCrLf)
                    cEinzelsatz = Mid(cSatz1, lPos, lPosEnde)
'                    MsgBox cEinzelsatz
                    lPos = lPos + lPosEnde - lPos + 2
                    lposSemi = 1
                    
                    rsrs.AddNew
                    
                    lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                    rsrs!artnr = cWert
                    
                    lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                    rsrs!BEZEICH = Trim(SwapStr(cWert, "'", " "))
                    
                    lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                    rsrs!LPZ = cWert
                    
                    lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                    rsrs!RKZ = Trim(SwapStr(cWert, "'", " "))
                    
                    lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1

                    cWert = SwapStr(cWert, ".", ",")
                    dWert = CDbl(cWert)
                    rsrs!lekpr = dWert
                    
                    lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1

                    cWert = SwapStr(cWert, ".", ",")
                    dWert = CDbl(cWert)
                    rsrs!vkpr = dWert
                    
                    lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                    rsrs!AGN = cWert
                    
                    lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                    rsrs!LIBESNR = Trim(SwapStr(cWert, "'", " "))
                    
                    lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                    rsrs!EAN = Trim(SwapStr(cWert, "'", " "))
                    
                    lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                    rsrs!MINMEN = cWert
                    
                    lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                    rsrs!MWST = Trim(SwapStr(cWert, "'", " "))
                    
                    lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                    rsrs!AWM = Trim(SwapStr(cWert, "'", " "))
                    
                    lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                    rsrs!linr = cWert
                    
                    lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                    rsrs!FLAG = Trim(SwapStr(cWert, "'", " "))
                    
                    
                    'In Notizen steckt das Kiss Aufdat
                    lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                    rsrs!NOTIZEN = Left(Trim(SwapStr(cWert, "'", " ")), 25)
                    
                    lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                    rsrs!INHALT = cWert
                    
                    lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                    rsrs!INHALTBEZ = Trim(SwapStr(cWert, "'", " "))
                    
                    lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                    rsrs!GRUNDPREIS = Trim(SwapStr(cWert, "'", " "))
                    
                    lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbCrLf): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                    rsrs!PGN = cWert
                    

                    
                    rsrs.Update
                Loop While lLenfil >= lPos
            End If
            
            Close iFileNr
            rsrs.Close: Set rsrs = Nothing
        Case "MLISRT"
            
            loeschNEW "MLISRT", dbZiel
            CreateTableT2 "MLISRT", dbZiel
            
            lPos = 1
            lPosEnde = 1
            lposSemiEnde = 1
            
            Set rsrs = dbZiel.OpenRecordset("MLISRT")
    
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
                    anzeige "normal", CStr(lcount), lblAnzeige
                    
                    lPosEnde = InStr(lPos, cSatz1, vbCrLf)
                    cEinzelsatz = Mid(cSatz1, lPos, lPosEnde)
                    lPos = lPos + lPosEnde - lPos + 2
                    lposSemi = 1
                    
                    rsrs.AddNew
                    
                    lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                    rsrs!linr = cWert
                    
                    lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                    rsrs!LIEFBEZ = Trim(SwapStr(cWert, "'", " "))
                    
                    lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                    rsrs!strasse = Trim(SwapStr(cWert, "'", " "))
                    
                    lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                    cWert = Trim(SwapStr(cWert, "'", " "))
                    rsrs!Plz = Left(cWert, 5)
                    
                    lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                    rsrs!STADT = Trim(SwapStr(cWert, "'", " "))
                    
                    lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                    rsrs!Tel = Trim(SwapStr(cWert, "'", " "))
                    
                    lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                    rsrs!LINAME = Trim(SwapStr(cWert, "'", " "))
                    
                    lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbCrLf): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                    rsrs!Fax = Trim(SwapStr(cWert, "'", " "))
                    
                    rsrs.Update
                Loop While lLenfil >= lPos
                
            End If
            
            Close iFileNr
            rsrs.Close: Set rsrs = Nothing

        Case "MLINBEZ"
            
            loeschNEW "MLINBEZ", dbZiel
            CreateTableT2 "MLINBEZ", dbZiel
            
            lPos = 1
            lPosEnde = 1
            lposSemiEnde = 1
            
            Set rsrs = dbZiel.OpenRecordset("MLINBEZ")
    
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
                    anzeige "normal", CStr(lcount), lblAnzeige
                    
                    lPosEnde = InStr(lPos, cSatz1, vbCrLf)
                    cEinzelsatz = Mid(cSatz1, lPos, lPosEnde)
                    
                    If cEinzelsatz <> "" Then
                        lPos = lPos + lPosEnde - lPos + 2
                        lposSemi = 1
                        
                        rsrs.AddNew
                        
                        lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                        rsrs!linr = cWert
                        
                        lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                        rsrs!LINBEZEICH = Left(Trim(SwapStr(cWert, "'", " ")), 30)
                        
                        lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                        rsrs!LPZ = cWert
                        
                        lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                        rsrs!SORTI = cWert
                        
                        lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                        rsrs!MARKE = Trim(SwapStr(cWert, "'", " "))
                        
                        lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbCrLf): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                        rsrs!MARKER = Trim(SwapStr(cWert, "'", " "))
                        
                        rsrs.Update
                    End If
                Loop While lLenfil >= lPos
                
            End If
            
            Close iFileNr
            rsrs.Close: Set rsrs = Nothing
        
        Case "LIEFKURZ"
        
            loeschNEW "LIEFKURZ", dbZiel
            CreateTableT2 "LIEFKURZ", dbZiel
            
            lPos = 1
            lPosEnde = 1
            lposSemiEnde = 1
            
            Set rsrs = dbZiel.OpenRecordset("LIEFKURZ")
    
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
                    anzeige "normal", CStr(lcount), lblAnzeige
                    
                    lPosEnde = InStr(lPos, cSatz1, vbCrLf)
                    cEinzelsatz = Mid(cSatz1, lPos, lPosEnde)
                    
                    If cEinzelsatz <> "" Then
                        lPos = lPos + lPosEnde - lPos + 2
                        lposSemi = 1
                        
                        rsrs.AddNew
                        
                        lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";"): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                        rsrs!Kuerzel = Trim(SwapStr(cWert, "'", " "))
                        
                        lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbCrLf): cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi): lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                        rsrs!linr = cWert
                        
                        rsrs.Update
                    End If
                Loop While lLenfil >= lPos
                
            End If
            
            Close iFileNr
            rsrs.Close: Set rsrs = Nothing
            
        Case "ARTEAN2"
        
            loeschNEW "ARTEAN2", dbZiel
            CreateTableT2 "ARTEAN2", dbZiel
            
            lPos = 1
            lPosEnde = 1
            lposSemiEnde = 1
            
            Set rsrs = dbZiel.OpenRecordset("ARTEAN2")
    
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
                    anzeige "normal", CStr(lcount), lblAnzeige
                    
                    lPosEnde = InStr(lPos, cSatz1, vbCrLf)
                    cEinzelsatz = Mid(cSatz1, lPos, lPosEnde)
                    
                    If cEinzelsatz <> "" Then
                        lPos = lPos + lPosEnde - lPos + 2
                        lposSemi = 1
                        
                        rsrs.AddNew
                        
                        lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";")
                        cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
                        lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                        
                        rsrs!artnr = cWert
                        
                        lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbCrLf)
                        cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
                        lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                        rsrs!EAN2 = Trim(SwapStr(cWert, "'", " "))
                        
                        rsrs.Update
                    End If
                Loop While lLenfil >= lPos
                
            End If
            
            Close iFileNr
            rsrs.Close: Set rsrs = Nothing
        Case "ARTEAN3"
        
            loeschNEW "ARTEAN3", dbZiel
            CreateTableT2 "ARTEAN3", dbZiel
            
            lPos = 1
            lPosEnde = 1
            lposSemiEnde = 1
            
            Set rsrs = dbZiel.OpenRecordset("ARTEAN3")
    
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
                    anzeige "normal", CStr(lcount), lblAnzeige
                    
                    lPosEnde = InStr(lPos, cSatz1, vbCrLf)
                    cEinzelsatz = Mid(cSatz1, lPos, lPosEnde)
                    
                    If cEinzelsatz <> "" Then
                        lPos = lPos + lPosEnde - lPos + 2
                        lposSemi = 1
                        
                        rsrs.AddNew
                        
                        lposSemiEnde = InStr(lposSemi, cEinzelsatz, ";")
                        cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
                        lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                        
                        rsrs!artnr = cWert
                        
                        lposSemiEnde = InStr(lposSemi, cEinzelsatz, vbCrLf)
                        cWert = Mid(cEinzelsatz, lposSemi, lposSemiEnde - lposSemi)
                        lposSemi = lposSemi + lposSemiEnde - lposSemi + 1
                        rsrs!EAN3 = Trim(SwapStr(cWert, "'", " "))
                        
                        rsrs.Update
                    End If
                Loop While lLenfil >= lPos
                
            End If
            
            Close iFileNr
            rsrs.Close: Set rsrs = Nothing
        
    End Select

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "csvImport"
    Fehler.gsFehlertext = "Im Programmteil Stammdaten einlesen ist ein Fehler aufgetreten."

    Fehlermeldung1
    Resume Next
End Sub
Public Function CheckDatum(ByVal sDatum As String) As Boolean
On Error GoTo LOKAL_ERROR

  Dim lDate As Long
  
  CheckDatum = False
  
  lDate = DateValue(sDatum)
  
  If lDate > 30000 And lDate < CLng(DateValue(Now) + 365) Then
    CheckDatum = True
  End If
  
Exit Function
LOKAL_ERROR:
    If err.Number = 13 Then
        CheckDatum = False
        Exit Function
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul7"
        Fehler.gsFunktion = "CheckDatum"
        Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
        Fehlermeldung1
    End If
End Function
Public Function ermgesUmsatzMwstAusZumsatz(cVon As String, cBis As String, sMWSTArt As String) As Double
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    
    ermgesUmsatzMwstAusZumsatz = 0
    
    If sMWSTArt = "V" Then
        sSQL = "Select sum(UMSV1) as Maxi"
    ElseIf sMWSTArt = "E" Then
        sSQL = "Select sum(UMSE1) as Maxi"
    ElseIf sMWSTArt = "O" Then
        sSQL = "Select sum(UMSO1) as Maxi"
    End If
    sSQL = sSQL & " from UMSATZ "
    sSQL = sSQL & " where Datum between  " & cVon & " And " & cBis
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermgesUmsatzMwstAusZumsatz = rsrs!maxi
        End If
    End If
    rsrs.Close
                
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "ermgesUmsatzMwstAusZumsatz"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Function
Public Function ermgesKK(cVon As String, cBis As String, sKartenart As String) As Double
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    
    ermgesKK = 0
    
    sSQL = "Select sum(GELDWERT) as Maxi"
    sSQL = sSQL & " from KKZAHL "
    sSQL = sSQL & " where ADATE between  " & cVon & " And " & cBis
    sSQL = sSQL & " and KK_ART = '" & sKartenart & "'"
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermgesKK = rsrs!maxi
        End If
    End If
    rsrs.Close
                
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "ermgesKK"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermgesKKgesamt(cVon As String, cBis As String) As Double
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    
    ermgesKKgesamt = 0
    
    sSQL = "Select sum(GELDWERT) as Maxi"
    sSQL = sSQL & " from KKZAHL "
    sSQL = sSQL & " where ADATE between  " & cVon & " And " & cBis
'    sSQL = sSQL & " and KK_ART = '" & sKartenart & "'"
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermgesKKgesamt = rsrs!maxi
        End If
    End If
    rsrs.Close
                
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "ermgesKKgesamt"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermgesBenuAGN(cVon As String, cBis As String, lagn As Long) As Double
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    
    ermgesBenuAGN = 0
    
    sSQL = "Select sum(Preis) as Maxi"
    sSQL = sSQL & " from Kassjour "
    sSQL = sSQL & " where ADATE between  " & cVon & " And " & cBis
    sSQL = sSQL & " and AGN = " & lagn
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermgesBenuAGN = rsrs!maxi
        End If
    End If
    rsrs.Close
                
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "ermgesBenuAGN"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermgesBARgesamtMwst(cVon As String, cBis As String, cMwst As String) As Double
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    
    ermgesBARgesamtMwst = 0
    
    sSQL = "Select sum(Preis) as Maxi"
    sSQL = sSQL & " from Kassjour "
    sSQL = sSQL & " where ADATE between  " & cVon & " And " & cBis
    sSQL = sSQL & " and MWST = '" & cMwst & "'"
    
    
    sSQL = sSQL & " and KK_ART = 'BA'"
    sSQL = sSQL & " and ums_ok = 'J'"
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermgesBARgesamtMwst = rsrs!maxi
        End If
    End If
    rsrs.Close
                
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "ermgesBARgesamtMwst"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermgesKKgesamtMwst(cVon As String, cBis As String, cMwst As String) As Double
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    
    ermgesKKgesamtMwst = 0
    
    sSQL = "Select sum(Preis) as Maxi"
    sSQL = sSQL & " from Kassjour "
    sSQL = sSQL & " where ADATE between  " & cVon & " And " & cBis
    sSQL = sSQL & " and MWST = '" & cMwst & "'"
    
    sSQL = sSQL & " and (KK_ART ='EC'"
    sSQL = sSQL & " or KK_ART ='VI'"
    sSQL = sSQL & " or KK_ART ='EU'"
    sSQL = sSQL & " or KK_ART ='AE'"
    sSQL = sSQL & " or KK_ART ='DC'"
    sSQL = sSQL & " or KK_ART ='BC'"
    sSQL = sSQL & " or KK_ART ='SO')"
    
    
    
    
               
'    sSQL = sSQL & " and KK_ART in ('EC','VI','EU','AE','DC','BC',SO')"
    sSQL = sSQL & " and ums_ok = 'J'"
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermgesKKgesamtMwst = rsrs!maxi
        End If
    End If
    rsrs.Close
                
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "ermgesKKgesamtMwst"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermgesAuszahlungsgrundDatev(cVon As String, cBis As String, sAuszahlungsgrund As String) As Double
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    
    ermgesAuszahlungsgrundDatev = 0
    
    sSQL = "Select sum(Betrag) as Maxi"
    sSQL = sSQL & " from KAEINAUS "
    sSQL = sSQL & " where ADATE between  " & cVon & " And " & cBis
    sSQL = sSQL & " and BEZEICH = '" & sAuszahlungsgrund & "'"
    sSQL = sSQL & " and UCASE(art) = 'AUSZAHLUNG'"
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermgesAuszahlungsgrundDatev = rsrs!maxi
        End If
    End If
    rsrs.Close
                
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "ermgesAuszahlungsgrundDatev"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermgesKREDAusZumsatz(cVon As String, cBis As String) As Double
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    
    ermgesKREDAusZumsatz = 0
    
    sSQL = "Select sum(KRED1) as Maxi"
    sSQL = sSQL & " from UMSATZ "
    sSQL = sSQL & " where Datum between  " & cVon & " And " & cBis
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermgesKREDAusZumsatz = rsrs!maxi
        End If
    End If
    rsrs.Close
                
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "ermgesKREDAusZumsatz"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Function
Public Function ermGutschausKarte(cVon As String, cBis As String) As Double
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    
    ermGutschausKarte = 0
    
    sSQL = "Select sum(GutschKar) as Maxi"
    sSQL = sSQL & " from AFCSTATP "
    sSQL = sSQL & " where adate between  " & cVon & " And " & cBis
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermGutschausKarte = rsrs!maxi
        End If
    End If
    rsrs.Close
                
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "ermGutschausKarte"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermGutschausGutsch(cVon As String, cBis As String) As Double
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    
    ermGutschausGutsch = 0
    
    sSQL = "Select sum(GutschGutsch) as Maxi"
    sSQL = sSQL & " from AFCSTATP "
    sSQL = sSQL & " where adate between  " & cVon & " And " & cBis
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermGutschausGutsch = rsrs!maxi
        End If
    End If
    rsrs.Close
                
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "ermGutschausGutsch"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermVK_RESTGUTSCH(cVon As String, cBis As String) As Double
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    
    ermVK_RESTGUTSCH = 0
    
'    sSQL = "Select sum(zhlggutsch) as Maxi"
    
'    RESTGUTSCH
    
    
    sSQL = "Select sum(RESTGUTSCH) as Maxi"
    sSQL = sSQL & " from AFCSTATP "
    sSQL = sSQL & " where adate between  " & cVon & " And " & cBis
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermVK_RESTGUTSCH = rsrs!maxi
        End If
    End If
    rsrs.Close
                
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "ermVK_RESTGUTSCH"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermKreditTilgung(cVon As String, cBis As String) As Double
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    
    ermKreditTilgung = 0
    
'        dTilgung = dTilgBar + dTilgSch + dTilgGut + dTilgKar
        
    
    sSQL = "Select sum(TilgBar) + sum(TilgSch)+  sum(TilgGut) + sum(TilgKar) as Maxi"
    sSQL = sSQL & " from AFCSTATP "
    sSQL = sSQL & " where adate between  " & cVon & " And " & cBis
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermKreditTilgung = rsrs!maxi
        End If
    End If
    rsrs.Close
                
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "ermKreditTilgung"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermGutschausKred(cVon As String, cBis As String) As Double
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    
    ermGutschausKred = 0
    
    sSQL = "Select sum(GutschKre) as Maxi"
    sSQL = sSQL & " from AFCSTATP "
    sSQL = sSQL & " where adate between  " & cVon & " And " & cBis
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermGutschausKred = rsrs!maxi
        End If
    End If
    rsrs.Close
                
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "ermGutschausKred"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermNichtumsatzRelevant(cVon As String, cBis As String) As Double
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    
    ermNichtumsatzRelevant = 0
    
    sSQL = "Select sum(Preis) as Maxi"
    sSQL = sSQL & " from Kassjour "
    sSQL = sSQL & " where adate between  " & cVon & " And " & cBis
    sSQL = sSQL & " and UMS_OK = 'N' "
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermNichtumsatzRelevant = rsrs!maxi
        End If
    End If
    rsrs.Close
                
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "ermNichtumsatzRelevant"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermGutschausBar(cVon As String, cBis As String) As Double
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    Dim cErgebnis   As String
    
    ermGutschausBar = 0
    
    sSQL = "Select sum(GutschBar) as Maxi"
    sSQL = sSQL & " from AFCSTATP "
    sSQL = sSQL & " where adate between  " & cVon & " And " & cBis
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermGutschausBar = rsrs!maxi
        End If
    End If
    rsrs.Close
     
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "ermGutschausBar"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermX_ausAFCSTATP(cVon As String, cBis As String, sXArt As String) As Double
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    Dim cErgebnis   As String
    
    ermX_ausAFCSTATP = 0
    
    sSQL = "Select sum(" & sXArt & ") as Maxi"
    sSQL = sSQL & " from AFCSTATP "
    sSQL = sSQL & " where adate between  " & cVon & " And " & cBis
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermX_ausAFCSTATP = rsrs!maxi
        End If
    End If
    rsrs.Close
     
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "ermX_ausAFCSTATP"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermKassensaldo(cVon As String, cBis As String) As Double
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    Dim lWert1      As Long
    Dim lWert2      As Long
    Dim dWECHSEL    As Double
    Dim dABSCHOPF   As Double
    

    
    ermKassensaldo = 0
    
    sSQL = "Select sum(BarVerkauf) + sum(GutschBar) + sum(TilgBar) + sum(Einzahlung) -  sum(Auszahlung) -  sum(AuszGutsch) as Maxi"
    sSQL = sSQL & " from AFCSTATP "
    sSQL = sSQL & " where adate between  " & cVon & " And " & cBis
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermKassensaldo = rsrs!maxi
        End If
    End If
    rsrs.Close
    'Jetzt noch wechsel addieren
    
    lWert1 = cVon
    lWert2 = cBis
    
    
    loeschNEW "temp_KABUCH", gdBase
    sSQL = "Select * into temp_KABUCH "
    sSQL = sSQL & " from KABUCH where BEZUMS = 'Wechselgeld' "
    sSQL = sSQL & " and DATUM >= " & Trim$(Str$(lWert1)) & " "
    sSQL = sSQL & " and DATUM <= " & Trim$(Str$(lWert2)) & " "
    gdBase.Execute sSQL, dbFailOnError
    
    
    loeschNEW "temp_KABUCH_Wechsel", gdBase
    
    sSQL = "Select max(EURBAR) as euronenbar ,Max(autopos) as maxi,datum,kasnum into temp_KABUCH_Wechsel from temp_KABUCH group by datum,kasnum  "
    gdBase.Execute sSQL, dbFailOnError
    



    sSQL = "Select SUM(euronenbar) as SWECHSEL"
    sSQL = sSQL & " from temp_KABUCH_Wechsel    "
    
    
    
    
    
    
    
    
'    sSQL = "Select SUM(EURBAR) as SWECHSEL"
'    sSQL = sSQL & " from KABUCH where "
'    sSQL = sSQL & " BEZUMS = 'Wechselgeld' "
    
    
    
    sSQL = sSQL & " where DATUM >= " & Trim$(Str$(lWert1)) & " "
    sSQL = sSQL & " and DATUM <= " & Trim$(Str$(lWert2)) & " "
        
    Set rsrs = gdBase.OpenRecordset(sSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        
        If Not IsNull(rsrs!SWECHSEL) Then
            dWECHSEL = rsrs!SWECHSEL
        Else
            dWECHSEL = 0
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    ermKassensaldo = ermKassensaldo + dWECHSEL
    
    'Jetzt noch abschopf subtrahieren
    
    sSQL = "Select SUM(Geldwert) as SABSCHOPF"
    sSQL = sSQL & " from ABSCHOPF where "
    
    lWert1 = CLng(cVon)
    lWert2 = CLng(cBis)
    sSQL = sSQL & "ADATE >= " & Trim$(Str$(lWert1)) & " "
    sSQL = sSQL & "and ADATE <= " & Trim$(Str$(lWert2)) & " "
       
    Set rsrs = gdBase.OpenRecordset(sSQL)
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        
        If Not IsNull(rsrs!SABSCHOPF) Then
            dABSCHOPF = rsrs!SABSCHOPF
        Else
            dABSCHOPF = 0
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    ermKassensaldo = ermKassensaldo - dABSCHOPF
    
    
    
    loeschNEW "temp_KABUCH", gdBase
    loeschNEW "temp_KABUCH_Wechsel", gdBase
    
Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "ermKassensaldo"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermgesAUSZAHLUNG(cVon As String, cBis As String, sArt As String) As Double
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    
    ermgesAUSZAHLUNG = 0
    
    sSQL = "Select sum(BETRAG) as Maxi"
    sSQL = sSQL & " from KAEINAUS "
    sSQL = sSQL & " where ADATE between  " & cVon & " And " & cBis
    sSQL = sSQL & " and ART = '" & sArt & "'"
    sSQL = sSQL & " and BEZEICH <>  'KB - Korrektur' "
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermgesAUSZAHLUNG = rsrs!maxi
        End If
    End If
    rsrs.Close
                
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "ermgesAUSZAHLUNG"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermlastMaxfromWechsel(lDat As Long) As Long
On Error GoTo LOKAL_ERROR

    ermlastMaxfromWechsel = 0
    Dim cSQL As String
    Dim rsrs As Recordset
    cSQL = "Select max(datum) as maxidat from wechsel where datum < " & lDat

    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        If Not IsNull(rsrs!Maxidat) Then

            ermlastMaxfromWechsel = rsrs!Maxidat
        End If
    End If
    rsrs.Close

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "ermlastMaxfromWechsel"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermgesKassendiff(cVon As String, cBis As String) As Double
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    Dim cART        As String
    Dim dBetrag     As Double
    
    ermgesKassendiff = 0
    
    'jetzt müssen wir bei den Auszahlungen das Vorzeichen drehen
    
    sSQL = "Select * "
    sSQL = sSQL & " from KAEINAUSF "
    sSQL = sSQL & " where ADATE between  " & cVon & " And " & cBis
    sSQL = sSQL & " and BEZEICH = 'KB - Korrektur'"
    
    
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
    
        Do While Not rsrs.EOF
        
            cART = ""
            If Not IsNull(rsrs!art) Then
                cART = rsrs!art
            End If
            
            dBetrag = 0
            If Not IsNull(rsrs!Betrag) Then
                dBetrag = rsrs!Betrag
            End If
            
            If UCase(cART) = "AUSZAHLUNG" Then
                If dBetrag > 0 Then
                    dBetrag = dBetrag * -1
                End If
            End If
            
            ermgesKassendiff = ermgesKassendiff + dBetrag
            rsrs.MoveNext
        Loop
        
    End If
    rsrs.Close
    
    
    
    
    
    
    'das war mal
    
'    sSQL = "Select sum(BETRAG) as Maxi"
'    sSQL = sSQL & " from KAEINAUSF "
'    sSQL = sSQL & " where ADATE between  " & cVon & " And " & cBis
'    sSQL = sSQL & " and BEZEICH = 'KB - Korrektur'"
'
'    Set rsrs = gdBase.OpenRecordset(sSQL)
'    If Not rsrs.EOF Then
'        If Not IsNull(rsrs!maxi) Then
'            ermgesKassendiff = rsrs!maxi
'        End If
'    End If
'    rsrs.Close
                
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "ermgesKassendiff"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermgesGUTZ(cVon As String, cBis As String) As Double
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    
    ermgesGUTZ = 0
    
    sSQL = "Select sum(GELDWERT) as Maxi"
    sSQL = sSQL & " from GUTZ "
    sSQL = sSQL & " where ADATE between  " & cVon & " And " & cBis
    sSQL = sSQL & " and ART = 'EI'"
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermgesGUTZ = rsrs!maxi
        End If
    End If
    rsrs.Close
                
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "ermgesGUTZ"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermgesLASTZAHLTE(cVon As String, cBis As String) As Double
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    
    ermgesLASTZAHLTE = 0
    
    sSQL = "Select sum(GELDWERT) as Maxi"
    sSQL = sSQL & " from LASTZAHLTE "
    sSQL = sSQL & " where ADATE between  " & cVon & " And " & cBis
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermgesLASTZAHLTE = rsrs!maxi
        End If
    End If
    rsrs.Close
                
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "ermgesLASTZAHLTE"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermgesABSCHOPF(cVon As String, cBis As String) As Double
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    
    ermgesABSCHOPF = 0
    
    sSQL = "Select sum(GELDWERT) as Maxi"
    sSQL = sSQL & " from ABSCHOPF "
    sSQL = sSQL & " where ADATE between  " & cVon & " And " & cBis
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!maxi) Then
            ermgesABSCHOPF = rsrs!maxi
        End If
    End If
    rsrs.Close
                
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "ermgesABSCHOPF"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function




Public Function gueltigeAGN(lagn As Long) As Boolean
    On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rs As Recordset
    
    gueltigeAGN = False
    
    sSQL = "Select * From AGNDBF"
    sSQL = sSQL & "  where agn = " & lagn
    Set rs = gdBase.OpenRecordset(sSQL)
    If Not rs.EOF Then
        gueltigeAGN = True
    End If
    rs.Close
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "gueltigeAGN"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Public Function gueltigeLINR(lLinr As Long) As Boolean
    On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rs As Recordset
    
    gueltigeLINR = False
    
    sSQL = "Select * From LISRT"
    sSQL = sSQL & "  where LINR = " & lLinr
    Set rs = gdBase.OpenRecordset(sSQL)
    If Not rs.EOF Then
        gueltigeLINR = True
    End If
    rs.Close
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "gueltigeLINR"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Function
Public Sub SchemaZuordnung(sSchema As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim rsRL As Recordset
    Dim i As Byte
    
    sSchema = SwapStr(sSchema, "'", "")
    
    sSQL = "Select * from SPZUORD where SCHEMANAME = '" & sSchema & "'"
    Set rsRL = gdBase.OpenRecordset(sSQL)
    If Not rsRL.EOF Then
        rsRL.MoveLast
        byAnzahlSpaltenEX = rsRL.RecordCount
        ReDim sFremdSpalteEX(byAnzahlSpaltenEX)
        ReDim sKissSpalteEX(byAnzahlSpaltenEX)
        
        rsRL.MoveFirst
        i = 0
        Do While Not rsRL.EOF
            sFremdSpalteEX(i) = rsRL!ZUFREMDSP
            sKissSpalteEX(i) = rsRL!ZUKISSDABA
            i = i + 1
            rsRL.MoveNext
        Loop
    End If
    rsRL.Close

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "SchemaZuordnung"
    Fehler.gsFehlertext = "Es trat ein Fehler auf. "
    
    Fehlermeldung1
End Sub
Public Function ermBEDbez(lbednu As Long) As String
    On Error GoTo LOKAL_ERROR


    Dim sSQL As String
    Dim rs As Recordset
    
    ermBEDbez = ""
    
    sSQL = "Select Bedname From Bedname"
    sSQL = sSQL & "  where BEDNU = " & lbednu

    Set rs = gdBase.OpenRecordset(sSQL)
    If Not rs.EOF Then
    rs.MoveFirst
        If Not IsNull(rs!bedname) Then
            ermBEDbez = rs!bedname
        End If
    End If
    rs.Close
    
Exit Function
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "ermBEDbez"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
    
End Function
Public Sub MBrechnen1(dBVO As Double, iVKTage As Integer, iReserv As Integer, lVon As Long, lBis As Long, lblAnzeige As Label, lvonNot As Long, lbisNot As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim sTab        As String
    Dim rs          As Recordset
    Dim lartnr      As Long
    Dim j           As Integer
    Dim dnewMB      As Double
    Dim lnewMB      As Long
    Dim lmbnew      As Long
    Dim lcount      As Long
    Dim lVkMenge    As Long
    Dim lAnz        As Long
    Dim dTeiler     As Double
    
    dTeiler = iVKTage / 30
    
    
    loeschNEW "ARTIKEL", gdApp
    loeschNEW "KASSJOUR", gdApp
    
    anzeige "normal", "Artikel werden importiert...", lblAnzeige
    TransferTab gdBase, App.Path & "/kissapp.mdb", "Artikel"
    anzeige "normal", "Kassenjournal wird importiert...", lblAnzeige
    TransferTab gdBase, App.Path & "/kissapp.mdb", "KASSJOUR"
    
    anzeige "normal", "Index(Artnr) wird erstellt...", lblAnzeige
    sSQL = "Create Index ARTNR on KASSJOUR (ARTNR)"
    gdApp.Execute sSQL, dbFailOnError
''    schreibeProtokollNachtAblauf sSQL
    
    anzeige "normal", "Index(Menge) wird erstellt...", lblAnzeige
    sSQL = "Create Index MENGE on KASSJOUR (MENGE)"
    gdApp.Execute sSQL, dbFailOnError
'    schreibeProtokollNachtAblauf sSQL
    
    anzeige "normal", "Index(Adate) wird erstellt...", lblAnzeige
    sSQL = "Create Index ADATE on KASSJOUR (ADATE)"
    gdApp.Execute sSQL, dbFailOnError
'    schreibeProtokollNachtAblauf sSQL
    
    anzeige "normal", "Index(Filiale) wird erstellt...", lblAnzeige
    sSQL = "Create Index filiale on KASSJOUR (filiale)"
    gdApp.Execute sSQL, dbFailOnError
'    schreibeProtokollNachtAblauf sSQL
    
    loeschNEW "DRUMBAE1", gdApp
    CreateTable "DRUMBAE1", gdApp

    loeschNEW "DRUMBAE", gdApp
    CreateTable "DRUMBAE", gdApp

    sSQL = " Insert into DRUMBAE select " & gcFilNr & " as filiale , artnr from Artikel "
    gdApp.Execute sSQL, dbFailOnError
'    schreibeProtokollNachtAblauf sSQL
        
    sTab = "KASS" & gcFilNr
    loeschNEW sTab, gdApp
    sSQL = "Select artnr, menge into " & sTab & " from kassjour where  "
    sSQL = sSQL & " ADATE between " & Trim$(Str$(lVon)) & " and " & Trim$(Str$(lBis)) & " "
    sSQL = sSQL & " and not ADATE between " & Trim$(Str$(lvonNot)) & " and " & Trim$(Str$(lbisNot)) & " "
    gdApp.Execute sSQL, dbFailOnError
'    schreibeProtokollNachtAblauf sSQL
    
    sSQL = "Create Index ARTNR on " & sTab & " (ARTNR)"
    gdApp.Execute sSQL, dbFailOnError
'    schreibeProtokollNachtAblauf sSQL

    sSQL = "Create Index MENGE on " & sTab & " (MENGE)"
    gdApp.Execute sSQL, dbFailOnError
'    schreibeProtokollNachtAblauf sSQL
        
    Set rs = gdApp.OpenRecordset("DRUMBAE")
    If Not rs.EOF Then
        rs.MoveLast
        lcount = rs.RecordCount
        rs.MoveFirst
        Do While Not rs.EOF
            lcount = lcount - 1
            
            j = lcount Mod 1000
            If j = 0 Then
                anzeige "normal", "Noch " & lcount & " neue Mindestbestände werden errechnet.", lblAnzeige
            End If
            
            If Not IsNull(rs!artnr) Then
                lartnr = rs!artnr
            End If
    
            lVkMenge = Ermittlevk(CLng(gcFilNr), CLng(rs!artnr), gdApp, sTab)
            
            dnewMB = 0
            
            dnewMB = (dBVO / dTeiler) * lVkMenge
           
            lnewMB = 0
            If dnewMB > 0 Then
            
                lnewMB = Val(dnewMB)
                lnewMB = lnewMB + 1
                
            End If
    
            rs.Edit
            rs!VKMENGE = lVkMenge
            rs!NEWMB = lnewMB
            rs.Update
            
            rs.MoveNext
        Loop
    End If
    rs.Close
    
    sSQL = "Insert into DRUMBAE1 Select * from Drumbae "
    gdApp.Execute sSQL, dbFailOnError
'    schreibeProtokollNachtAblauf sSQL
    
    'Übernahme
    
    anzeige "normal", "Löschen nicht relevanter Daten...", lblAnzeige
    
    sSQL = "Delete from  DRUMBAE1 where newmb = 0 "
    gdApp.Execute sSQL, dbFailOnError
'    schreibeProtokollNachtAblauf sSQL
    
    anzeige "normal", "Mindestbestände auf 0 setzen...", lblAnzeige
    
    sSQL = "update Artikel set minbest = 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    Set rs = gdApp.OpenRecordset("drumbae1")
    If Not rs.EOF Then
        rs.MoveLast
        lAnz = rs.RecordCount
        rs.MoveFirst
        Do While Not rs.EOF
            
            lAnz = lAnz - 1
            
            j = lAnz Mod 1000
            If j = 0 Then
                anzeige "normal", "Noch " & lAnz & " neue Mindestbestände werden übernommen.", lblAnzeige
            End If
            
            If Not IsNull(rs!artnr) Then
                lartnr = rs!artnr
                
                If Not IsNull(rs!NEWMB) Then
                    lmbnew = rs!NEWMB
                Else
                    lmbnew = 0
                End If
                schreibeNeuMBausNacht lartnr, lmbnew
            End If
        rs.MoveNext
        Loop
    End If
    rs.Close
    
    Dim rsrs1 As Recordset

    sSQL = "Select * from MBORDER "
    Set rsrs1 = gdBase.OpenRecordset(sSQL)
    If Not rsrs1.EOF Then
        rsrs1.MoveFirst
        Do While Not rsrs1.EOF
            If Not IsNull(rsrs1!artnr) Then
                lartnr = rsrs1!artnr
            End If

            If Not IsNull(rsrs1!MB) Then
                lmbnew = rsrs1!MB
            End If

            schreibeNeuMBausNacht lartnr, lmbnew
        rsrs1.MoveNext
        Loop
    End If
    rsrs1.Close
    
    loeschNEW "MBSTAND", gdBase
    CreateTableT2 "MBSTAND", gdBase
    
    sSQL = "Insert into MBSTAND (LASTDATE,LASTTIME) values ('" & DateValue(Now) & "','" & TimeValue(Now) & "') "
    gdBase.Execute sSQL, dbFailOnError

    anzeige "normal", "Fertig", lblAnzeige
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "MBrechnen1"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Public Function Ermittlevk(lfinr As Long, lartikel As Long, daba As Database, sTab As String) As Long
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL    As String
    Dim rsrs    As Recordset
    
    Ermittlevk = 0
    
    cSQL = "Select sum(Menge)as VKMELJ from " & sTab & " where artnr = " & lartikel
    Set rsrs = daba.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!VKMELJ) Then
            Ermittlevk = rsrs!VKMELJ
        End If
    End If
    
    rsrs.Close
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "Ermittlevk"
    Fehler.gsFehlertext = "Bei der Ermittlung der Verkaufsmenge Vorjahr ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ErmMBSTAND() As Long
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL    As String
    Dim rsrs    As Recordset
    
    ErmMBSTAND = 0
    
    cSQL = "Select Lastdate from MBSTAND "
    Set rsrs = gdBase.OpenRecordset(cSQL)
    
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!LASTDATE) Then
            ErmMBSTAND = CLng(rsrs!LASTDATE)
        End If
    End If
    
    rsrs.Close
    
Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul2"
    Fehler.gsFunktion = "ErmMBSTAND"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Sub schreibeNeuMBausNacht(lartikel As Long, lmbneu As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL            As String
    Dim rsrs            As Recordset
    
    cSQL = "Select * from Artikel where artnr = " & lartikel
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        rsrs.Edit
        rsrs!MINBEST = lmbneu
        rsrs!LASTDATE = DateValue(Now)
        rsrs!LASTTIME = TimeValue(Now)
        rsrs.Update
    End If
    
    rsrs.Close
            
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "schreibeNeuMBausNacht"
    Fehler.gsFehlertext = "Bei der Schreiben der neuen Mindestbestände ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub leseMBDetails()
    On Error GoTo LOKAL_ERROR
    
    Dim rsrs As Recordset
    
    MBDETAILBVO = 1.5
    MBDETAILVON = 0
    MBDETAILBIS = 0
    MBDETAILMON = 2
  
   
    
    If NewTableSuchenDBKombi("MBDETAIL", gdBase) = False Then
        CreateTableT2 "MBDETAIL", gdBase
    End If
    
    Set rsrs = gdBase.OpenRecordset("MBDETAIL")
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!BVO) Then
            MBDETAILBVO = rsrs!BVO
        Else
            MBDETAILBVO = 1.5
        End If
        
        If Not IsNull(rsrs!Von) Then
            MBDETAILVON = rsrs!Von
        Else
            MBDETAILVON = 0
        End If
        
        If Not IsNull(rsrs!Bis) Then
            MBDETAILBIS = rsrs!Bis
        Else
            MBDETAILBIS = 0
        End If
        
        If Not IsNull(rsrs!optmon) Then
            MBDETAILMON = rsrs!optmon
        Else
            MBDETAILMON = 2
        End If
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "leseMBDetails"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub speicherlastOptionEinstellung(i As Integer, sTab As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    loeschNEW sTab, gdApp
    CreateTable sTab, gdApp
    
    sSQL = "Insert into " & sTab & " (Ind) values (" & i & ")"
    gdApp.Execute sSQL, dbFailOnError
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "speicherlastOptionEinstellung"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Function ermFarbeKU(cFarbnr As String) As String
    On Error GoTo LOKAL_ERROR


    Dim sSQL As String
    Dim rs As Recordset
    
    ermFarbeKU = ""
    
    sSQL = "Select Farbtext From FARBKU "
    sSQL = sSQL & "  where FARBNR = " & cFarbnr

    Set rs = gdBase.OpenRecordset(sSQL)
    
    If Not rs.EOF Then
    rs.MoveFirst
        If Not IsNull(rs!farbtext) Then
            ermFarbeKU = rs!farbtext
        End If
    End If
    rs.Close: Set rs = Nothing
Exit Function
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "ermFarbeKU"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
    
End Function
Public Function ermFarbeBez(cFarbnr As String) As String
    On Error GoTo LOKAL_ERROR

    Dim sSQL As String
    Dim rs As Recordset
    
    ermFarbeBez = ""
    
    sSQL = "Select Farbtext From FARBMERK "
    sSQL = sSQL & "  where FARBNR = " & cFarbnr
    Set rs = gdBase.OpenRecordset(sSQL)
    If Not rs.EOF Then
    rs.MoveFirst
        If Not IsNull(rs!farbtext) Then
            ermFarbeBez = rs!farbtext
        End If
    End If
    rs.Close: Set rs = Nothing
    
Exit Function
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "ermFarbeBez"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."

    Fehlermeldung1
    
End Function
Public Function LeselastOptionEinstellung(sTab As String) As Byte
    On Error GoTo LOKAL_ERROR
    
    Dim rsrs As Recordset
    Dim sSQL As String
    
    LeselastOptionEinstellung = 0
    
    If Not NewTableSuchenDBKombi(sTab, gdApp) Then
        CreateTable sTab, gdApp
        
        sSQL = "Insert into " & sTab & " (Ind) values (0)"
        gdApp.Execute sSQL, dbFailOnError
    End If

    Set rsrs = gdApp.OpenRecordset(sTab)
    If Not rsrs.EOF Then
        LeselastOptionEinstellung = rsrs!ind
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "LeselastOptionEinstellung"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Sub Datenbankwechsel()
    On Error GoTo LOKAL_ERROR
    Dim sSQL                As String
    Dim rsrs                As Recordset
    Dim lDatum              As Long
    Dim lTime               As Long
    Dim lDatumDat           As Long
    Dim sDatumDat           As String
    Dim sdateTime           As String
    Dim lTimeDat            As Long
    
    Set dabalokal = Nothing
    'Computer Datum/Zeit
    lDatum = Fix(Now)
    sdateTime = TimeValue(Now)
    sdateTime = Format(sdateTime, "hh:mm")
    sdateTime = SwapStr(sdateTime, ":", "")
    lTime = CLng(sdateTime)

    Select Case glLokalAktuZeit
    
        Case 5
            If CInt(Right(CStr(lTime), 2)) < glLokalAktuZeit Then
                lTime = lTime - 45
            Else
                lTime = lTime - glLokalAktuZeit
            End If
        Case 10
            If CInt(Right(CStr(lTime), 2)) < glLokalAktuZeit Then
                lTime = lTime - 50
            Else
                lTime = lTime - glLokalAktuZeit
            End If
        Case 20
            If CInt(Right(CStr(lTime), 2)) < glLokalAktuZeit Then
                lTime = lTime - 60
            Else
                lTime = lTime - glLokalAktuZeit
            End If
        
        Case 30
            If CInt(Right(CStr(lTime), 2)) < glLokalAktuZeit Then
                lTime = lTime - 70
            Else
                lTime = lTime - glLokalAktuZeit '
            End If
        Case 40
            If CInt(Right(CStr(lTime), 2)) < glLokalAktuZeit Then
                lTime = lTime - 80
            Else
                lTime = lTime - glLokalAktuZeit '
            End If
        Case 50
            If CInt(Right(CStr(lTime), 2)) < glLokalAktuZeit Then
                lTime = lTime - 90
            Else
                lTime = lTime - glLokalAktuZeit '
            End If
            
        Case 60
            lTime = lTime - 100
        
        Case Else
            lTime = lTime - glLokalAktuZeit '
    End Select

    'Datenbank Datum/Zeit
    sSQL = "Select * from wkeinste "
    Set rsrs = gdApp.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
    
        If Not IsNull(rsrs!localtime) Then
            lTimeDat = rsrs!localtime
        Else
            lTimeDat = 0
        End If
        
        If Not IsNull(rsrs!localdat) Then
            sDatumDat = rsrs!localdat
        Else
            sDatumDat = 0
        End If

    End If
    rsrs.Close: Set rsrs = Nothing


    lDatumDat = DateValue(sDatumDat)

    If sDatumDat = "" Then
        gsAnforderung = "ALLES"
        Kopiere
        gsAnforderung = ""
    Else
        If lDatum = lDatumDat Then
            If lTime > lTimeDat Then
            
                gsAnforderung = "ALLES"
                Kopiere
                gsAnforderung = ""
            
'                Kopiere
            End If
        Else
            gsAnforderung = "ALLES"
            Kopiere
            gsAnforderung = ""
        End If
    End If

    Set dabalokal = OpenDatabase(App.Path & "\lokal.mdb", False, False, "MS Access;PWD=" & gsPasswort)
        
    Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
        
    ElseIf err.Number = 3011 Or err.Number = 3343 Then
    
        Kill App.Path & "\lokal.mdb"
        Kopiere
        Set dabalokal = OpenDatabase(App.Path & "\lokal.mdb", False, False, "MS Access;PWD=" & gsPasswort)
'        Set dabalokal = OpenDatabase(App.Path & "\lokal.mdb", False)
        Exit Sub
        
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul7"
        Fehler.gsFunktion = "Datenbankwechsel"
        Fehler.gsFehlertext = "Beim Wechseln der Datenbank ist ein Fehler aufgetreten."
    
        Fehlermeldung1
    End If
End Sub
Public Sub Kopiere()
    On Error GoTo LOKAL_ERROR
    
    giCopyMod = 1
    frmWK21n.Show 1
    giCopyMod = 0
    
    
    
    

    
    Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul7"
        Fehler.gsFunktion = "Kopiere"
        Fehler.gsFehlertext = "Beim Wechseln der Datenbank(Kopieren) ist ein Fehler aufgetreten."
    
        Fehlermeldung1
    End If
End Sub
Public Function LoeseMarkenInArtnr(cKrit As String) As Boolean
On Error GoTo LOKAL_ERROR

    Dim sSQL    As String
    Dim rsrs    As Recordset
    
    Screen.MousePointer = 11
    
    LoeseMarkenInArtnr = False
    
    sSQL = "Delete from MY" & srechnertab
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Select LPZ,LINR from LINBEZ where Marke like '" & cKrit & "*' "
    sSQL = sSQL & " and not Linr is null "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            If Not IsNull(rsrs!LPZ) Then
                If Not IsNull(rsrs!linr) Then
                
                    sSQL = " Insert into MY" & srechnertab
                    sSQL = sSQL & " select a.artnr from artikel a , artlief b  "
                    
                    sSQL = sSQL & " where a.LPZ = " & rsrs!LPZ & " and b.LINR = " & rsrs!linr
                    sSQL = sSQL & " and  a.artnr =  b.artnr "
                    gdBase.Execute sSQL, dbFailOnError
                
'                    sSQL = "Insert into MY" & srechnertab
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
    
    If Datendrin("MY" & srechnertab, gdBase) Then
        LoeseMarkenInArtnr = True
    End If
    
    Screen.MousePointer = 0
    
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "LoeseMarkenInArtnr"
    Fehler.gsFehlertext = "Im Programmteil Artikelsuche ist ein Fehler aufgetreten. "
    Fehlermeldung1
End Function
Public Function LoesePGNInArtnr(cKrit As String, bmehrere As Boolean, db As Database) As Boolean
On Error GoTo LOKAL_ERROR

    Dim sSQL    As String
    Dim rsrs    As Recordset
    
    Screen.MousePointer = 11
    
    LoesePGNInArtnr = False
    
    If bmehrere = True Then
    
    Else
    
        sSQL = "Delete from MY" & srechnertab
        db.Execute sSQL, dbFailOnError
        
    End If
    
    
    
    sSQL = "Insert into MY" & srechnertab
    sSQL = sSQL & " Select artnr from Artikel where"
    sSQL = sSQL & " PGN = " & cKrit
    db.Execute sSQL, dbFailOnError
                
    
    If Datendrin("MY" & srechnertab, db) Then
        CheckIndex "MY" & srechnertab, "artnr", "", db
        
        LoesePGNInArtnr = True
    End If
    
    Screen.MousePointer = 0
    
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul1"
    Fehler.gsFunktion = "LoesePGNInArtnr"
    Fehler.gsFehlertext = "Im Programmteil Artikelsuche ist ein Fehler aufgetreten. "
    Fehlermeldung1
End Function
Public Function LoeseMarkenstringinLPZ(cKrit As String) As Boolean
On Error GoTo LOKAL_ERROR

    Dim sSQL    As String
    Dim rsrs    As Recordset
    Dim i       As Integer
    
    LoeseMarkenstringinLPZ = False
    
    For i = 0 To 100
      gBYTENum(i) = 255555
    Next i
    
    For i = 0 To 100
        gBYTENumLIN(i) = 300200
    Next i
    
    i = 0
    
    sSQL = "Select LPZ,LINR from LINBEZ where Marke like '" & cKrit & "*' "
    sSQL = sSQL & " and not Linr is null "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            If Not IsNull(rsrs!LPZ) Then
                If i = 35 Then
                    LoeseMarkenstringinLPZ = False
                
                    Exit Function
                
                Else
                    gBYTENumLIN(i) = rsrs!linr
                    gBYTENum(i) = rsrs!LPZ
                    i = i + 1
                    LoeseMarkenstringinLPZ = True
                End If
            
            End If
        
        rsrs.MoveNext
        Loop
    
    End If
    rsrs.Close: Set rsrs = Nothing
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "LoeseMarkenstringinLPZ"
    Fehler.gsFehlertext = "Im Programmteil Artikelsuche ist ein Fehler aufgetreten. "
    Fehlermeldung1
End Function
Public Function LoeseMarkenstringinLPZ12(cKrit As String) As Boolean
On Error GoTo LOKAL_ERROR

    Dim sSQL    As String
    Dim rsrs    As Recordset
    Dim i       As Integer
    Dim lLinr   As Long
    Dim lLpz    As Long
    
    LoeseMarkenstringinLPZ12 = False
    
    sSQL = "Delete from MA" & srechnertab
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Select LPZ,LINR from LINBEZ where Marke like '" & cKrit & "*' "
    sSQL = sSQL & " and not Linr is null "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            If Not IsNull(rsrs!LPZ) Then
                
                lLinr = rsrs!linr
                lLpz = rsrs!LPZ
                 
                sSQL = " Insert into MA" & srechnertab
                sSQL = sSQL & " select a.artnr from artikel a , artlief b  "
                
                sSQL = sSQL & " where a.LPZ = " & lLpz & " and b.LINR = " & lLinr
                sSQL = sSQL & " and  a.artnr =  b.artnr "
                gdBase.Execute sSQL, dbFailOnError
                
                LoeseMarkenstringinLPZ12 = True
            
            End If
        
        rsrs.MoveNext
        Loop
    
    End If
    
    rsrs.Close: Set rsrs = Nothing
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "LoeseMarkenstringinLPZ12"
    Fehler.gsFehlertext = "Im Programmteil Artikelsuche ist ein Fehler aufgetreten. "
    Fehlermeldung1
End Function
Public Function islocked(rs As Recordset) As Boolean
On Error GoTo openfehler

islocked = False
rs.Edit
rs.CancelUpdate

Exit Function

openfehler:
islocked = True
Resume Next
Exit Function
End Function
Public Sub HoleLokalDB()
    On Error GoTo LOKAL_ERROR
    
    Dim ws As Workspace
    Dim sPfad As String
    Dim sZielpfad As String
    Dim lokalDB As Database
    Dim i As Integer
    Dim j As Single
    Dim sSQL As String
    Dim rsrs As DAO.Recordset
    
    Dim iMaxTab As Integer
    iMaxTab = 69
    Dim sTabellen(0 To 69) As String
    Dim sDauerTabellen(0 To 9) As String
    
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
    sTabellen(28) = "BWWBONTEXTE"
    sTabellen(29) = "BEDZUGRI"
    sTabellen(30) = "ABSCHLUSS"
    sTabellen(31) = "NOEURO"
    sTabellen(32) = "ZUORDEAN"
    sTabellen(33) = "EKASS"
    sTabellen(34) = "ZBONLAY"
    sTabellen(35) = "AFCSTATP"
    sTabellen(36) = "GEMZ"
    sTabellen(37) = "PREISEDITKASSE"
    sTabellen(38) = "LASTZAHLTE"
    sTabellen(39) = "GUTZ"
    sTabellen(40) = "KAEINAUSF"
    sTabellen(41) = "ALTERG"
    sTabellen(42) = "GUHIS"
    sTabellen(43) = "STORNO2"
    sTabellen(44) = "UNTERWF"
    sTabellen(45) = "KASSBOND"
    sTabellen(46) = "PREISE"
    sTabellen(47) = "KUDD"
    sTabellen(48) = "ARTAUSWAHL"
    sTabellen(49) = "BONUSART"
    sTabellen(50) = "KASSBEDP"
    sTabellen(51) = "BONUSBONTEXTE"
    sTabellen(52) = "DISPLAYTEXT"
    sTabellen(53) = "KONTIN"
    sTabellen(54) = "ARTMERK"
    sTabellen(55) = "GESCHWART"
    sTabellen(56) = "GUHIN"
    sTabellen(57) = "KUNDKASS"
    sTabellen(58) = "BANKKU"
    sTabellen(59) = "EINAUSKB"
    sTabellen(60) = "KUNDA" & srechnertab
    sTabellen(61) = "STORNOF"
    sTabellen(62) = "FARBMERK"
    sTabellen(63) = "GARANTIE"
    sTabellen(64) = "KARTEN_EINZ"
    sTabellen(65) = "STAFFELPRKVK"
    sTabellen(66) = "BESTPROT"
    sTabellen(67) = "ARTEAN_K"
    sTabellen(68) = "STAFFEL_KVK_ARTIKEL"
    sTabellen(69) = "STAFFEL_KVK_GRUPPE"
    
    
    sDauerTabellen(0) = "KASSBON"
    sDauerTabellen(1) = "KOLLVERK"
    sDauerTabellen(2) = "KREDIT"
    sDauerTabellen(3) = "AFCBUCH"
    sDauerTabellen(4) = "KASSJOUR"
    sDauerTabellen(5) = "RETOURE"
    sDauerTabellen(6) = "AFCAUFBON"
    sDauerTabellen(7) = "AFCSTAT"
    sDauerTabellen(8) = "KKZAHLTE"
    sDauerTabellen(9) = "MARKIERUNG"
    
    
    frmWKL00.txtStatus.Text = "0"
    frmWKL00.picprogress.Visible = True
    frmWKL00.Label2.Visible = True
    
    If NewTableSuchenDBKombi("ZZZ", gdBase) = False Then
        frmWKL00.txtStatus.Text = "0"
        frmWKL00.picprogress.Visible = False
        frmWKL00.Label2.Visible = False
        
        Exit Sub
    End If
    
    frmWKL00.Label2.Caption = "Die lokale Datenbank wird jetzt aktualisiert..."
    frmWKL00.Label2.Refresh
    
    sPfad = "C:\aLeer"
    Kill "C:\aLeer\Safe.mdb"
    sZielpfad = "C:\aLeer\kissdata.mdb"

    'Prüfen ob Verzeichnis c:\aleer existiert
    VerzVorhanden "aLeer", "C:\"
    VerzVorhanden "GDPDU", "C:\aLeer\"
    
    'Prüfen ob kissdata.mdb existiert
    
    Dim cPfad       As String
    Dim lRet As Long
    Dim lfail As Long

    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    'frische Kassbon erstellen ohne Passwort
    Kill sPfad & "\GDPDU\kassbon.mdb"
    Set ws = DBEngine.Workspaces(0)
    Set lokalDB = ws.CreateDatabase(sPfad & "\GDPDU\kassbon.mdb", dbLangGeneral, dbVersion40)
    CreateTable "KASSBOND", lokalDB
    lokalDB.Close
    'Ende frische Kassbon erstellen ohne Passwort
        
    If Not Modul6.FindFile(sPfad, "kissdata.mdb") Then
        Set ws = DBEngine.Workspaces(0)
        Set lokalDB = ws.CreateDatabase(sPfad & "\kissdata.mdb", dbLangGeneral, dbVersion40)
        
        For i = 0 To iMaxTab

            j = i Mod 2
            If j = 0 Then
                frmWKL00.Label2.ForeColor = vbYellow
            Else
                frmWKL00.Label2.ForeColor = vbBlue
            End If

            frmWKL00.txtStatus.Text = i

            frmWKL00.Label2.Caption = "Die Tabelle " & sTabellen(i) & " wird jetzt aktualisiert..."
            frmWKL00.Label2.Refresh

            loeschNEW sTabellen(i), lokalDB
            PauseSi CSng(gdDBPAUSE)
            TransferTab gdBase, sZielpfad, sTabellen(i)
            
        Next i
        
        For i = 0 To 9
            j = i Mod 2
            If j = 0 Then
                frmWKL00.Label2.ForeColor = vbYellow
            Else
                frmWKL00.Label2.ForeColor = vbBlue
            End If
            
            frmWKL00.txtStatus.Text = i
        
            frmWKL00.Label2.Caption = "Die Tabelle " & sDauerTabellen(i) & " wird jetzt aktualisiert..."
            frmWKL00.Label2.Refresh
            
            loeschNEW sDauerTabellen(i), lokalDB
            PauseSi CSng(gdDBPAUSE)
            TransferTab gdBase, sZielpfad, sDauerTabellen(i)
        Next i
        
        lRet = CopyFile(cPfad & "FILNR.CFG", "C:\aleer\FILNR.CFG", lfail)
        lRet = CopyFile(cPfad & "alrs3.rpt", "C:\aleer\alrs3.rpt", lfail)
        lRet = CopyFile(cPfad & "aWKL21b.rpt", "C:\aleer\aWKL21b.rpt", lfail)
        lRet = CopyFile(cPfad & "aWKL21bh.rpt", "C:\aleer\aWKL21bh.rpt", lfail)
        lRet = CopyFile(cPfad & "aWKL21c.rpt", "C:\aleer\aWKL21c.rpt", lfail)
        lRet = CopyFile(cPfad & "aWKL20z.rpt", "C:\aleer\aWKL20z.rpt", lfail)
'        lRet = CopyFile(cPfad & "GDPDU\KASSBON.mdb", "C:\aleer\GDPDU\KASSBON.mdb", lfail)
        
        
    Else 'Datenbank existiert schon - dann Artikel,Kunden,Gutscheine usw holen
    
        lRet = CopyFile(cPfad & "FILNR.CFG", "C:\aleer\FILNR.CFG", lfail)
        lRet = CopyFile(cPfad & "aWKLm2a.rpt", "C:\aleer\aWKLm2a.rpt", lfail)
        lRet = CopyFile(cPfad & "alrs3.rpt", "C:\aleer\alrs3.rpt", lfail)
        lRet = CopyFile(cPfad & "aWKL21b.rpt", "C:\aleer\aWKL21b.rpt", lfail)
        lRet = CopyFile(cPfad & "aWKL21bh.rpt", "C:\aleer\aWKL21bh.rpt", lfail)
        lRet = CopyFile(cPfad & "aWKL21c.rpt", "C:\aleer\aWKL21c.rpt", lfail)
        lRet = CopyFile(cPfad & "aWKL20z.rpt", "C:\aleer\aWKL20z.rpt", lfail)
'        lRet = CopyFile(cPfad & "GDPDU\KASSBON.mdb", "C:\aleer\GDPDU\KASSBON.mdb", lfail)
        
        Set lokalDB = OpenDatabase(sPfad & "\kissdata.mdb", False)
        
    End If
    
    For i = 0 To iMaxTab

        j = i Mod 2
        If j = 0 Then
            frmWKL00.Label2.ForeColor = vbYellow
        Else
            frmWKL00.Label2.ForeColor = vbBlue
        End If
        frmWKL00.Label2.Caption = "Die Tabelle " & sTabellen(i) & " wird jetzt aktualisiert..."
        frmWKL00.Label2.Refresh

        loeschNEW sTabellen(i), lokalDB
'        TransferTab gdBase, sZielpfad, sTabellen(i)

        frmWKL00.txtStatus.Text = i
    Next i
    
    

    'Löschen der Inhalte von Kassjour , Afcbuch usw
    If NewTableSuchenDBKombi(sDauerTabellen(8), lokalDB) = False Then
        loeschNEW sDauerTabellen(8), lokalDB
        PauseSi CSng(gdDBPAUSE)
        TransferTab gdBase, sZielpfad, sDauerTabellen(8)
    End If
    
    If NewTableSuchenDBKombi(sDauerTabellen(9), lokalDB) = False Then
        loeschNEW sDauerTabellen(9), lokalDB
        PauseSi CSng(gdDBPAUSE)
        TransferTab gdBase, sZielpfad, sDauerTabellen(9)
    End If
    
    For i = 0 To 9
        sSQL = "Delete From " & sDauerTabellen(i)
        lokalDB.Execute sSQL, dbFailOnError
        
        
        
        
        frmWKL00.txtStatus.Text = i
    Next i
    
    
    If Not SpalteInTabellegefundenNEW("AFCSTAT", "NUMSKARTE", lokalDB) Then
        SpalteAnfuegenNEW "AFCSTAT", "NUMSKARTE", "double", lokalDB
    End If
    
    
    
    
    
    
    
    
    
    
    
    
    lokalDB.Close
    Set lokalDB = Nothing
    Kill sPfad & "\kissdata87.mdb"
    
    frmWKL00.Label2.Caption = "Komprimierung Teil 1..."
    frmWKL00.Label2.Refresh
    
    DBEngine.CompactDatabase sPfad & "\kissdata.mdb", sPfad & "\kissdata87.mdb", dbLangGeneral
    Kill sPfad & "\kissdata.mdb"
    PauseSi CSng(gdDBPAUSE)
    
    frmWKL00.Label2.Caption = "Komprimierung Teil 2..."
    frmWKL00.Label2.Refresh
    DBEngine.CompactDatabase sPfad & "\kissdata87.mdb", sPfad & "\kissdata.mdb", dbLangGeneral
    Kill sPfad & "\kissdata87.mdb"
    PauseSi CSng(gdDBPAUSE)
    Set lokalDB = OpenDatabase(sPfad & "\kissdata.mdb", False)
    
    
    Dim cKillPfad As String
    
    
    cKillPfad = sPfad & "\LPROTOK\PROABL.txt"
    Kill cKillPfad
    
    cKillPfad = sPfad & "\LPROTOK\BENUTZER.txt"
    Kill cKillPfad
    
    cKillPfad = sPfad & "\LPROTOK\DABAABL.txt"
    Kill cKillPfad
    
    For i = 0 To iMaxTab

        j = i Mod 2
        If j = 0 Then
            frmWKL00.Label2.ForeColor = vbYellow
        Else
            frmWKL00.Label2.ForeColor = vbBlue
        End If
        frmWKL00.Label2.Caption = "Die Tabelle " & sTabellen(i) & " wird jetzt aktualisiert..."
        frmWKL00.Label2.Refresh

        TransferTab gdBase, sZielpfad, sTabellen(i)

        frmWKL00.txtStatus.Text = i
    Next i
    
    sSQL = "Update Gutsch set STATUS = 'N', Synstatus = 'N' "
    lokalDB.Execute sSQL, dbFailOnError
    

    cPfad = gcDBPfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    VerzVorhanden "Picture", "C:\aLeer\"
    
    VerzVorhanden "EUR", "C:\aLeer\Picture\"
    VerzVorhanden "SFR", "C:\aLeer\Picture\"
    VerzVorhanden "DEM", "C:\aLeer\Picture\"
    VerzVorhanden "ARTIKEL", "C:\aLeer\Picture\"
    VerzVorhanden "KUNDEN", "C:\aLeer\Picture\"
    VerzVorhanden "SYSTEM", "C:\aLeer\Picture\"
    
    systembildcheck_all_4Lokal
    
    waehrungbildcheck_all_4Lokal
    
''    For i = 0 To 14
''        lRet = CopyFile(cPfad & "Picture\EUR\" & i & "kl.jpg", "C:\aleer\Picture\EUR\" & i & "kl.jpg", lfail)
''
''        lRet = CopyFile(cPfad & "Picture\EUR\" & i & "k.jpg", "C:\aleer\Picture\EUR\" & i & "k.jpg", lfail)
''    Next i

    lokalDB.Close

    Set lokalDB = OpenDatabase(sPfad & "\kissdata.mdb", True)
    ReIndiziereArtikelWKL00 lokalDB
    lokalDB.Close
    Set lokalDB = Nothing
    
    frmWKL00.txtStatus.Text = "0"
    frmWKL00.picprogress.Visible = False
    frmWKL00.Label2.Visible = False
    
    Exit Sub
LOKAL_ERROR:
    If err.Number = 53 Then
        Resume Next
    ElseIf err.Number = 3343 Then
        Kill sPfad & "\kissdata.mdb"
    
        Exit Sub
    Else
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul7"
        Fehler.gsFunktion = "HoleLokalDB"
        Fehler.gsFehlertext = "Beim Aktualisieren der lokalen Datenbank ist ein Fehler aufgetreten."
        
        Fehlermeldung1
'        Resume Next
    End If
End Sub
Private Sub systembildcheck_all_4Lokal()
On Error GoTo LOKAL_ERROR

    systembildcheck_4Lokal "Tabelle.jpg"
    systembildcheck_4Lokal "Tastatur.jpg"
    systembildcheck_4Lokal "Kalender.jpg"
    systembildcheck_4Lokal "Zurück.jpg"
    systembildcheck_4Lokal "Vor.jpg"
    systembildcheck_4Lokal "Visa.jpg"
    systembildcheck_4Lokal "Visa_kl.jpg"
    systembildcheck_4Lokal "American-Express.jpg"
    systembildcheck_4Lokal "American-Express_kl.jpg"
    systembildcheck_4Lokal "Diners-Club.jpg"
    systembildcheck_4Lokal "Diners-Club_kl.jpg"
    systembildcheck_4Lokal "Mastercard.jpg"
    systembildcheck_4Lokal "Mastercard_kl.jpg"
    systembildcheck_4Lokal "Maestro.jpg"
    systembildcheck_4Lokal "Maestro_kl.jpg"
    systembildcheck_4Lokal "diverse.jpg"
    systembildcheck_4Lokal "switch.jpg"
    systembildcheck_4Lokal "leute1.jpg"
    systembildcheck_4Lokal "Rechts.jpg"
    systembildcheck_4Lokal "Links.jpg"
    systembildcheck_4Lokal "Rauf.jpg"
    systembildcheck_4Lokal "Runter.jpg"
    systembildcheck_4Lokal "EC.jpg"
    systembildcheck_4Lokal "Brief.gif"
    systembildcheck_4Lokal "Briefrot.gif"
    
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "systembildcheck_all_4Lokal"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub waehrungbildcheck_all_4Lokal()
On Error GoTo LOKAL_ERROR

    waehrungbildcheck_4Lokal "0k.jpg"
    waehrungbildcheck_4Lokal "1k.jpg"
    waehrungbildcheck_4Lokal "2k.jpg"
    waehrungbildcheck_4Lokal "3k.jpg"
    waehrungbildcheck_4Lokal "4k.jpg"
    waehrungbildcheck_4Lokal "5k.jpg"
    waehrungbildcheck_4Lokal "6k.jpg"
    waehrungbildcheck_4Lokal "7k.jpg"
    waehrungbildcheck_4Lokal "8k.jpg"
    waehrungbildcheck_4Lokal "9k.jpg"
    waehrungbildcheck_4Lokal "10k.jpg"
    waehrungbildcheck_4Lokal "11k.jpg"
    waehrungbildcheck_4Lokal "12k.jpg"
    waehrungbildcheck_4Lokal "13k.jpg"
    waehrungbildcheck_4Lokal "14k.jpg"
    
    waehrungbildcheck_4Lokal "0g.jpg"
    waehrungbildcheck_4Lokal "1g.jpg"
    waehrungbildcheck_4Lokal "2g.jpg"
    waehrungbildcheck_4Lokal "3g.jpg"
    waehrungbildcheck_4Lokal "4g.jpg"
    waehrungbildcheck_4Lokal "5g.jpg"
    waehrungbildcheck_4Lokal "6g.jpg"
    waehrungbildcheck_4Lokal "7g.jpg"
    waehrungbildcheck_4Lokal "8g.jpg"
    waehrungbildcheck_4Lokal "9g.jpg"
    waehrungbildcheck_4Lokal "10g.jpg"
    waehrungbildcheck_4Lokal "11g.jpg"
    waehrungbildcheck_4Lokal "12g.jpg"
    waehrungbildcheck_4Lokal "13g.jpg"
    waehrungbildcheck_4Lokal "14g.jpg"
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "waehrungbildcheck_all_4Lokal"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub systembildcheck_4Lokal(cBild As String)
On Error GoTo LOKAL_ERROR

    Dim cPfad       As String
    Dim cQuelle     As String
    Dim cZiel       As String
    Dim lRet        As Long
    Dim lfail       As Long
    
    'check ob Bild vorliegt
    cPfad = gcDBPfad    'dabapfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    cPfad = cPfad & "PICTURE\System\"
    
    If FileExists(cPfad & cBild) Then
        cQuelle = cPfad
        cQuelle = ShortPath(cQuelle)
        cQuelle = cQuelle & cBild

        cZiel = "C:\aLeer\PICTURE\System\"
        cZiel = ShortPath(cZiel)
        cZiel = cZiel & cBild

        lRet = CopyFile(cQuelle, cZiel, lfail)
        
    End If
    

Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "systembildcheck_4Lokal"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Sub waehrungbildcheck_4Lokal(cBild As String)
On Error GoTo LOKAL_ERROR

    Dim cPfad       As String
    Dim cQuelle     As String
    Dim cZiel       As String
    Dim lRet        As Long
    Dim lfail       As Long
    
    'check ob Bild vorliegt
    cPfad = gcDBPfad    'dabapfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    cPfad = cPfad & "PICTURE\EUR\"
    
    If FileExists(cPfad & cBild) Then
        cQuelle = cPfad
        cQuelle = ShortPath(cQuelle)
        cQuelle = cQuelle & cBild

        cZiel = "C:\aLeer\PICTURE\EUR\"
        cZiel = ShortPath(cZiel)
        cZiel = cZiel & cBild

        lRet = CopyFile(cQuelle, cZiel, lfail)
        
    End If
    

Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "waehrungbildcheck_4Lokal"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub synchronisiereDB()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim sPfad       As String
    Dim slokalPfad  As String
    Dim lokalDB     As Database
    Dim rsAfcbuch   As Recordset
    Dim sArtnr      As String
    Dim iMenge      As Integer
    Dim rsKun       As Recordset
    Dim lDatum      As Long
    Dim cKundnr     As String
    Dim cPreis      As String
    Dim stimz       As String
    Dim cBedz       As String
    Dim i           As Integer

    lDatum = Fix(Now)
    
    Do While IsAktionZulaessig("Synchronisieren") = False
    
    Loop
    
    frmWKL00.txtStatus.Text = "0"
    frmWKL00.picprogress.Visible = True
    frmWKL00.Label2.Visible = True
    
    If NewTableSuchenDBKombi("ZZZ", gdBase) = False Then
        
        frmWKL00.txtStatus.Text = "0"
        frmWKL00.picprogress.Visible = False
        frmWKL00.Label2.Visible = False
        Exit Sub
           
    End If
    
    frmWKL00.txtStatus.Text = "10"
    
    slokalPfad = "C:\aleer\kissdata.mdb"
    
    If FileExists(slokalPfad) Then
        Set lokalDB = OpenDatabase(slokalPfad, False)
        
        sPfad = gcDBPfad 'Datenbankpfad
        If Right(sPfad, 1) <> "\" Then
            sPfad = sPfad & "\"
        End If
        
        frmWKL00!Label2.Caption = "Datenbank wird sysnchronisiert..."
        frmWKL00!Label2.Refresh
        
        sichernLdb
        
        frmWKL00.txtStatus.Text = "15"
        
        frmWKL00!Label2.Caption = "Datenbank: Bestand minimieren"
        frmWKL00!Label2.Refresh
        
        
        'Bestand minimieren
        If NewTableSuchenDBKombi("AFCAUFBON", lokalDB) = True Then
            Set rsAfcbuch = lokalDB.OpenRecordset("AFCAUFBON", dbOpenTable)
            If Not rsAfcbuch.EOF Then
                i = 0
                rsAfcbuch.MoveFirst
                Do While Not rsAfcbuch.EOF
                    If i = 100 Then
                        i = 0
                    End If
                    
                    i = i + 1
                    frmWKL00.txtStatus.Text = i
                    
                    If Not IsNull(rsAfcbuch!aartnr) Then
                        sArtnr = rsAfcbuch!aartnr
                    Else
                        sArtnr = ""
                    End If
                    
                    If Not IsNull(rsAfcbuch!aMenge) Then
                        iMenge = rsAfcbuch!aMenge
                    Else
                        iMenge = 0
                    End If
                    
                    If Not IsNull(rsAfcbuch!AZEIT) Then
                        stimz = rsAfcbuch!AZEIT
                    Else
                        stimz = ""
                    End If
                    
                    cBedz = "0"
    
                    If iMenge > 0 Then
                    'menge Positiv
                    'Dann Kassiervorgang
                        BestandsminiOrMaxi sArtnr, CLng(iMenge), "Kassiervorgang", stimz, cBedz, gcFilNr
                    Else
                    'Menge Negativ
                    'Dann Storno
                        BestandsminiOrMaxi sArtnr, CLng(iMenge), "Storno Kasse LM", stimz, cBedz, gcFilNr
                    End If
                    rsAfcbuch.MoveNext
                Loop
            End If
            rsAfcbuch.Close: Set rsAfcbuch = Nothing
        End If
    
        Set rsAfcbuch = lokalDB.OpenRecordset("Afcbuch", dbOpenTable)
        If Not rsAfcbuch.EOF Then
            i = 0
            rsAfcbuch.MoveFirst
            Do While Not rsAfcbuch.EOF
                If i = 100 Then
                    i = 0
                End If
                
                i = i + 1
                frmWKL00.txtStatus.Text = i
                
                If Not IsNull(rsAfcbuch!aartnr) Then
                    sArtnr = rsAfcbuch!aartnr
                Else
                    sArtnr = ""
                End If
                
                If Not IsNull(rsAfcbuch!aMenge) Then
                    iMenge = rsAfcbuch!aMenge
                Else
                    iMenge = 0
                End If
                
                If Not IsNull(rsAfcbuch!AZEIT) Then
                    stimz = rsAfcbuch!AZEIT
                Else
                    stimz = ""
                End If
                
                cBedz = "0"

                If iMenge > 0 Then
                'menge Positiv
                'Dann Kassiervorgang
                    BestandsminiOrMaxi sArtnr, CLng(iMenge), "Kassiervorgang", stimz, cBedz, gcFilNr
                Else
                'Menge Negativ
                'Dann Storno
                    BestandsminiOrMaxi sArtnr, CLng(iMenge), "Storno Kasse LM", stimz, cBedz, gcFilNr
                End If
                rsAfcbuch.MoveNext
            Loop
        End If
        rsAfcbuch.Close: Set rsAfcbuch = Nothing
        
        frmWKL00.txtStatus.Text = "20"
        
        frmWKL00!Label2.Caption = "Datenbank: Kunden ..."
        frmWKL00!Label2.Refresh
        
        'Kunden
'        loeschNEW "Kunden1", gdBase
        
        loeschNEW "Kunden1" & rechnername, gdBase
        
        sSQL = "Select * into Kunden1" & rechnername & " from Afcbuch in '" & slokalPfad & "' "
        sSQL = sSQL & " where akunum > 0 "
        gdBase.Execute sSQL, dbFailOnError
        
        frmWKL00.txtStatus.Text = "30"
        
        Set rsKun = gdBase.OpenRecordset("select * from Kunden1" & rechnername & " ")
    
        If Not rsKun.EOF Then
            rsKun.MoveFirst
            Do While Not rsKun.EOF
                cKundnr = rsKun!AKUNUM
                cPreis = rsKun!APREIS
                cPreis = fnMoveComma2Point$(cPreis)
                
                'Achtung Status
                sSQL = "Update Kunden set BONUS = BONUS + " & cPreis & " where KUNDNR = " & cKundnr
                gdBase.Execute sSQL, dbFailOnError
                rsKun.MoveNext
            Loop
        End If
        
        rsKun.Close: Set rsKun = Nothing
        
        frmWKL00.txtStatus.Text = "40"
        
        
        'Gutscheine
        
        loeschNEW "Gutsch1" & rechnername, gdBase

        sSQL = "Select * into GUTSCH1" & rechnername & " from GUTSCH in '" & slokalPfad & "' "
        sSQL = sSQL & " where Status <> 'N' "
        gdBase.Execute sSQL, dbFailOnError
        
        sSQL = "Delete from GUTSCH where gutschnr in (Select gutschnr from gutsch1" & rechnername & ")"
        gdBase.Execute sSQL, dbFailOnError
        
        sSQL = "Insert into GUTSCH Select * from gutsch1" & rechnername
        gdBase.Execute sSQL, dbFailOnError

        
        'Ende Gutscheine
        
        
        
        frmWKL00!Label2.Caption = "Datenbank: AFCBUCH ..."
        frmWKL00!Label2.Refresh
        
        'afcbuch
        sSQL = "Insert into afcbuch select * from afcbuch in '" & slokalPfad & "' "
        gdBase.Execute sSQL, dbFailOnError
        
        frmWKL00.txtStatus.Text = "50"
        
        frmWKL00!Label2.Caption = "Datenbank: Kassjour ..."
        frmWKL00!Label2.Refresh
        
        'Kassjour
        loeschNEW "K" & srechnertab, gdBase
        
        frmWKL00.txtStatus.Text = "51"
        
        frmWKL00!Label2.Caption = "Datenbank: Kassjour importieren..."
        frmWKL00!Label2.Refresh
        
        sSQL = "select * into K" & srechnertab & " from Kassjour in '" & slokalPfad & "' "
        gdBase.Execute sSQL, dbFailOnError
        
        frmWKL00.txtStatus.Text = "55"
        
        frmWKL00!Label2.Caption = "Datenbank: Kassjour aktualisieren..."
        frmWKL00!Label2.Refresh
        
        sSQL = "Insert into Kassjour select * from K" & srechnertab
        gdBase.Execute sSQL, dbFailOnError
        
        frmWKL00.txtStatus.Text = "60"
        
        frmWKL00!Label2.Caption = "Datenbank: Kredit ..."
        frmWKL00!Label2.Refresh
        
        'Kredit
        sSQL = "Insert into Kredit select * from Kredit in '" & slokalPfad & "' "
        gdBase.Execute sSQL, dbFailOnError
        
        frmWKL00.txtStatus.Text = "70"
        
        frmWKL00!Label2.Caption = "Datenbank: Kollverk ..."
        frmWKL00!Label2.Refresh
        
        'Kollverk
        sSQL = "Insert into Kollverk select * from Kollverk in '" & slokalPfad & "' "
        gdBase.Execute sSQL, dbFailOnError
        
        
        frmWKL00.txtStatus.Text = "72"
        
        frmWKL00!Label2.Caption = "Datenbank: Retoure ..."
        frmWKL00!Label2.Refresh
        
        'Retoure
        sSQL = "Insert into Retoure select * from Retoure in '" & slokalPfad & "' "
        gdBase.Execute sSQL, dbFailOnError
        
        frmWKL00.txtStatus.Text = "75"
        
        frmWKL00!Label2.Caption = "Datenbank: Kassbon ..."
        frmWKL00!Label2.Refresh
        
        'Kassbon
        sSQL = "Insert into Kassbon select * from Kassbon in '" & slokalPfad & "' "
        gdBase.Execute sSQL, dbFailOnError
        
        
        'auch in KassbonD GDPDU/Kassbon.mdb
        
        
        Dim cPfad               As String
        Dim KASSBON_DB          As Database
        Dim sKASSBON_Pfad       As String
        
        sKASSBON_Pfad = "C:\aleer\GDPDU\kassbon.mdb"
    
        cPfad = gcDBPfad
        If Right$(cPfad, 1) <> "\" Then
            cPfad = cPfad & "\"
        End If
        
        cPfad = cPfad & "GDPdU\KASSBON.MDB"
    
    
        
        Set KASSBON_DB = OpenDatabase(cPfad, False, False, "MS Access;PWD=" & gsKASSBON_Passwort)
        
        'KassbonD
        sSQL = "Insert into KASSBOND select * from KassbonD in '" & sKASSBON_Pfad & "' "
        KASSBON_DB.Execute sSQL, dbFailOnError
        
        
        
        KASSBON_DB.Close
        
        
    
        
        
        
        
        sSQL = "Insert into Kassbon select * from Kassbon in '" & slokalPfad & "' "
        gdBase.Execute sSQL, dbFailOnError
        
        
        
        frmWKL00.txtStatus.Text = "77"
        
        frmWKL00!Label2.Caption = "Datenbank: Kreditkartenzahlungen ..."
        frmWKL00!Label2.Refresh
        
        'KKZAHLTE
        sSQL = "Insert into KKZAHLTE select * from KKZAHLTE in '" & slokalPfad & "' "
        gdBase.Execute sSQL, dbFailOnError
            
        frmWKL00.txtStatus.Text = "80"
        
        'Afcstat
        frmWKL00!Label2.Caption = "Datenbank: Afstat ..."
        frmWKL00!Label2.Refresh
        
        loeschNEW "afcstat1" & rechnername, gdBase
        
        sSQL = "Select * into Afcstat1" & rechnername & " from Afcstat in '" & slokalPfad & "' "
        gdBase.Execute sSQL, dbFailOnError
        
        
        
        
        If Not SpalteInTabellegefundenNEW("AFCSTAT1" & rechnername, "NUMSKARTE", gdBase) Then
            SpalteAnfuegenNEW "AFCSTAT1" & rechnername, "NUMSKARTE", "double", gdBase
        End If
        
        AfcstatIstNull "AFCSTAT1" & rechnername
        
        
        
        AFCSTATPLUS_alleKassen "AFCSTAT1" & rechnername, "AFCSTAT"
        
        frmWKL00.txtStatus.Text = "100"
        
        
        lokalDB.Close
        Set lokalDB = Nothing
    End If
    
    frmWKL00.txtStatus.Text = "0"
    frmWKL00.picprogress.Visible = False
    frmWKL00.Label2.Visible = False
    
    AktionAustragen "Synchronisieren"
    
    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "synchronisiereDB"
    Fehler.gsFehlertext = "Beim Synchronisieren der Datenbanken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    Resume Next

End Sub
Public Sub AFCSTATPLUS(cFromTable As String, cIntoTable As String, cKaNum As String)
    On Error GoTo LOKAL_ERROR

    Dim rsAfc       As Recordset
    Dim rsAfc1      As Recordset
    Dim sSQL        As String
    Dim lDatum      As Long
    
    If Not SpalteInTabellegefundenNEW(cFromTable, "WECHSEL", gdBase) Then
        SpalteAnfuegenNEW cFromTable, "WECHSEL", "double", gdBase
    End If
    
    If Not SpalteInTabellegefundenNEW(cIntoTable, "WECHSEL", gdBase) Then
        SpalteAnfuegenNEW cIntoTable, "WECHSEL", "double", gdBase
    End If
    

    sSQL = "Select * from " & cFromTable & "  "
    sSQL = sSQL & " Where KASNUM = " & cKaNum 'gckasnum"
        
    
    Set rsAfc1 = gdBase.OpenRecordset(sSQL)
        
    If Not rsAfc1.EOF Then
        rsAfc1.MoveFirst
        Do While Not rsAfc1.EOF
        
            If Not IsNull(rsAfc1!ADATE) Then
                lDatum = rsAfc1!ADATE
            Else
                lDatum = -1
            End If
            
            
            sSQL = "Select * from " & cIntoTable & "  where KASNUM = " & cKaNum & " and ADATE = " & Trim$(Str$(lDatum)) & " "
            Set rsAfc = gdBase.OpenRecordset(sSQL)
        
            If Not rsAfc.EOF Then
                rsAfc.Edit 'gibt es in der großen Daba ein Eintrag
            Else
                rsAfc.AddNew 'gibt es in der großen Daba noch kein Eintrag
                rsAfc!ADATE = lDatum
                rsAfc!kasnum = cKaNum
            End If
            
            If Not IsNull(rsAfc1!UMS_BAR) Then 'Tagestabelle
                If Not IsNull(rsAfc!UMS_BAR) Then
                    rsAfc!UMS_BAR = rsAfc!UMS_BAR + rsAfc1!UMS_BAR
                Else
                    rsAfc!UMS_BAR = rsAfc1!UMS_BAR
                End If
            Else
                rsAfc!UMS_BAR = 0
            End If
            
            If Not IsNull(rsAfc1!UMS_Kred) Then
                If Not IsNull(rsAfc!UMS_Kred) Then
                    rsAfc!UMS_Kred = rsAfc!UMS_Kred + rsAfc1!UMS_Kred
                Else
                    rsAfc!UMS_Kred = rsAfc1!UMS_Kred
                End If
            Else
                rsAfc!UMS_Kred = 0
            End If
            
            '*******
            
            If Not IsNull(rsAfc1!UMS_SCHECK) Then
                If Not IsNull(rsAfc!UMS_SCHECK) Then
                    rsAfc!UMS_SCHECK = rsAfc!UMS_SCHECK + rsAfc1!UMS_SCHECK
                Else
                    rsAfc!UMS_SCHECK = rsAfc1!UMS_SCHECK
                End If
            Else
                rsAfc!UMS_SCHECK = 0
            End If
            
            If Not IsNull(rsAfc1!UMS_KARTE) Then
                If Not IsNull(rsAfc!UMS_KARTE) Then
                    rsAfc!UMS_KARTE = rsAfc!UMS_KARTE + rsAfc1!UMS_KARTE
                Else
                    rsAfc!UMS_KARTE = rsAfc1!UMS_KARTE
                End If
            Else
                rsAfc!UMS_KARTE = 0
            End If
            
            '***
            
            If Not IsNull(rsAfc1!UMS_LAST) Then
                If Not IsNull(rsAfc!UMS_LAST) Then
                    rsAfc!UMS_LAST = rsAfc!UMS_LAST + rsAfc1!UMS_LAST
                Else
                    rsAfc!UMS_LAST = rsAfc1!UMS_LAST
                End If
            Else
                rsAfc!UMS_LAST = 0
            End If
    
            If Not IsNull(rsAfc1!SPREIS_ANZ) Then
                If Not IsNull(rsAfc!SPREIS_ANZ) Then
                    rsAfc!SPREIS_ANZ = rsAfc!SPREIS_ANZ + rsAfc1!SPREIS_ANZ
                Else
                    rsAfc!SPREIS_ANZ = rsAfc1!SPREIS_ANZ
                End If
            Else
                rsAfc!SPREIS_ANZ = 0
            End If
            
            If Not IsNull(rsAfc1!SPREIS_GES) Then
                If Not IsNull(rsAfc!SPREIS_GES) Then
                    rsAfc!SPREIS_GES = rsAfc!SPREIS_GES + rsAfc1!SPREIS_GES
                Else
                    rsAfc!SPREIS_GES = rsAfc1!SPREIS_GES
                End If
            Else
                rsAfc!SPREIS_GES = 0
            End If
            
            If Not IsNull(rsAfc1!ANZSCHECK) Then
                If Not IsNull(rsAfc!ANZSCHECK) Then
                    rsAfc!ANZSCHECK = rsAfc!ANZSCHECK + rsAfc1!ANZSCHECK
                Else
                    rsAfc!ANZSCHECK = rsAfc1!ANZSCHECK
                End If
            Else
                rsAfc!ANZSCHECK = 0
            End If
            
            If Not IsNull(rsAfc1!Kundenzahl) Then
                If Not IsNull(rsAfc!Kundenzahl) Then
                    rsAfc!Kundenzahl = rsAfc!Kundenzahl + rsAfc1!Kundenzahl
                Else
                    rsAfc!Kundenzahl = rsAfc1!Kundenzahl
                End If
            Else
                rsAfc!Kundenzahl = 0
            End If
            
            If Not IsNull(rsAfc1!GELDFACH) Then
                If Not IsNull(rsAfc!GELDFACH) Then
                    rsAfc!GELDFACH = rsAfc!GELDFACH + rsAfc1!GELDFACH
                Else
                    rsAfc!GELDFACH = rsAfc1!GELDFACH
                End If
            Else
                rsAfc!GELDFACH = 0
            End If
                
            If Not IsNull(rsAfc1!ARTRAB_ANZ) Then
                If Not IsNull(rsAfc!ARTRAB_ANZ) Then
                    rsAfc!ARTRAB_ANZ = rsAfc!ARTRAB_ANZ + rsAfc1!ARTRAB_ANZ
                Else
                    rsAfc!ARTRAB_ANZ = rsAfc1!ARTRAB_ANZ
                End If
            Else
                rsAfc!ARTRAB_ANZ = 0
            End If
            
            If Not IsNull(rsAfc1!ARTRAB_GES) Then
                If Not IsNull(rsAfc!ARTRAB_GES) Then
                    rsAfc!ARTRAB_GES = rsAfc!ARTRAB_GES + rsAfc1!ARTRAB_GES
                Else
                    rsAfc!ARTRAB_GES = rsAfc1!ARTRAB_GES
                End If
            Else
                rsAfc!ARTRAB_GES = 0
            End If
            
            If Not IsNull(rsAfc1!GESRAB_ANZ) Then
                If Not IsNull(rsAfc!GESRAB_ANZ) Then
                    rsAfc!GESRAB_ANZ = rsAfc!GESRAB_ANZ + rsAfc1!GESRAB_ANZ
                Else
                    rsAfc!GESRAB_ANZ = rsAfc1!GESRAB_ANZ
                End If
            Else
                rsAfc!GESRAB_ANZ = 0
            End If
            
            If Not IsNull(rsAfc1!GESRAB_GES) Then
                If Not IsNull(rsAfc!GESRAB_GES) Then
                    rsAfc!GESRAB_GES = rsAfc!GESRAB_GES + rsAfc1!GESRAB_GES
                Else
                    rsAfc!GESRAB_GES = rsAfc1!GESRAB_GES
                End If
            Else
                rsAfc!GESRAB_GES = 0
            End If
            
            If Not IsNull(rsAfc1!STORNO_ANZ) Then
                If Not IsNull(rsAfc!STORNO_ANZ) Then
                    rsAfc!STORNO_ANZ = rsAfc!STORNO_ANZ + rsAfc1!STORNO_ANZ
                Else
                    rsAfc!STORNO_ANZ = rsAfc1!STORNO_ANZ
                End If
            Else
                rsAfc!STORNO_ANZ = 0
            End If
    
            If Not IsNull(rsAfc1!STORNO_GES) Then
                If Not IsNull(rsAfc!STORNO_GES) Then
                    rsAfc!STORNO_GES = rsAfc!STORNO_GES + rsAfc1!STORNO_GES
                Else
                    rsAfc!STORNO_GES = rsAfc1!STORNO_GES
                End If
            Else
                rsAfc!STORNO_GES = 0
            End If
            
            If Not IsNull(rsAfc1!EINZAHLUNG) Then
                If Not IsNull(rsAfc!EINZAHLUNG) Then
                    rsAfc!EINZAHLUNG = rsAfc!EINZAHLUNG + rsAfc1!EINZAHLUNG
                Else
                    rsAfc!EINZAHLUNG = rsAfc1!EINZAHLUNG
                End If
            Else
                rsAfc!EINZAHLUNG = 0
            End If
            
            If Not IsNull(rsAfc1!AUSZAHLUNG) Then
                If Not IsNull(rsAfc!AUSZAHLUNG) Then
                    rsAfc!AUSZAHLUNG = rsAfc!AUSZAHLUNG + rsAfc1!AUSZAHLUNG
                Else
                    rsAfc!AUSZAHLUNG = rsAfc1!AUSZAHLUNG
                End If
            Else
                rsAfc!AUSZAHLUNG = 0
            End If
            
            If Not IsNull(rsAfc1!GUTSCHEIN) Then
                If Not IsNull(rsAfc!GUTSCHEIN) Then
                    rsAfc!GUTSCHEIN = rsAfc!GUTSCHEIN + rsAfc1!GUTSCHEIN
                Else
                    rsAfc!GUTSCHEIN = rsAfc1!GUTSCHEIN
                End If
            Else
                rsAfc!GUTSCHEIN = 0
            End If
            
            If Not IsNull(rsAfc1!ZHLGGUTSCH) Then
                If Not IsNull(rsAfc!ZHLGGUTSCH) Then
                    rsAfc!ZHLGGUTSCH = rsAfc!ZHLGGUTSCH + rsAfc1!ZHLGGUTSCH
                Else
                    rsAfc!ZHLGGUTSCH = rsAfc1!ZHLGGUTSCH
                End If
            Else
                rsAfc!ZHLGGUTSCH = 0
            End If
            
            If Not IsNull(rsAfc1!BELEGNR) Then
                rsAfc!BELEGNR = rsAfc1!BELEGNR
            Else
                rsAfc!BELEGNR = 0
            End If
            
            If Not IsNull(rsAfc1!GUTSCHBAR) Then
                If Not IsNull(rsAfc!GUTSCHBAR) Then
                    rsAfc!GUTSCHBAR = rsAfc!GUTSCHBAR + rsAfc1!GUTSCHBAR
                Else
                    rsAfc!GUTSCHBAR = rsAfc1!GUTSCHBAR
                End If
            Else
                rsAfc!GUTSCHBAR = 0
            End If
            
            If Not IsNull(rsAfc1!GUTSCHSCH) Then
                If Not IsNull(rsAfc!GUTSCHSCH) Then
                    rsAfc!GUTSCHSCH = rsAfc!GUTSCHSCH + rsAfc1!GUTSCHSCH
                Else
                    rsAfc!GUTSCHSCH = rsAfc1!GUTSCHSCH
                End If
            Else
                rsAfc!GUTSCHSCH = 0
            End If
            
            If Not IsNull(rsAfc1!GUTSCHKRE) Then
                If Not IsNull(rsAfc!GUTSCHKRE) Then
                    rsAfc!GUTSCHKRE = rsAfc!GUTSCHKRE + rsAfc1!GUTSCHKRE
                Else
                    rsAfc!GUTSCHKRE = rsAfc1!GUTSCHKRE
                End If
            Else
                rsAfc!GUTSCHKRE = 0
            End If
            
            If Not IsNull(rsAfc1!GUTSCHKAR) Then
                If Not IsNull(rsAfc!GUTSCHKAR) Then
                    rsAfc!GUTSCHKAR = rsAfc!GUTSCHKAR + rsAfc1!GUTSCHKAR
                Else
                    rsAfc!GUTSCHKAR = rsAfc1!GUTSCHKAR
                End If
            Else
                rsAfc!GUTSCHKAR = 0
            End If
            
            If Not IsNull(rsAfc1!GUTSCHLAST) Then
                If Not IsNull(rsAfc!GUTSCHLAST) Then
                    rsAfc!GUTSCHLAST = rsAfc!GUTSCHLAST + rsAfc1!GUTSCHLAST
                Else
                    rsAfc!GUTSCHLAST = rsAfc1!GUTSCHLAST
                End If
            Else
                rsAfc!GUTSCHLAST = 0
            End If
            
            If Not IsNull(rsAfc1!GUTSCHGUTSCH) Then
                If Not IsNull(rsAfc!GUTSCHGUTSCH) Then
                    rsAfc!GUTSCHGUTSCH = rsAfc!GUTSCHGUTSCH + rsAfc1!GUTSCHGUTSCH
                Else
                    rsAfc!GUTSCHGUTSCH = rsAfc1!GUTSCHGUTSCH
                End If
            Else
                rsAfc!GUTSCHGUTSCH = 0
            End If
            
            
            
            If Not IsNull(rsAfc1!Wechsel) Then
                If Not IsNull(rsAfc!Wechsel) Then
                    rsAfc!Wechsel = rsAfc!Wechsel + rsAfc1!Wechsel
                Else
                    rsAfc!Wechsel = rsAfc1!Wechsel
                End If
            Else
                rsAfc!Wechsel = 0
            End If
        
            
            
            If Not IsNull(rsAfc1!Abschopf) Then
                If Not IsNull(rsAfc!Abschopf) Then
                    rsAfc!Abschopf = rsAfc!Abschopf + rsAfc1!Abschopf
                Else
                    rsAfc!Abschopf = rsAfc1!Abschopf
                End If
            Else
                rsAfc!Abschopf = 0
            End If
            
            If Not IsNull(rsAfc1!KDIFF) Then
                If Not IsNull(rsAfc!KDIFF) Then
                    rsAfc!KDIFF = rsAfc!KDIFF + rsAfc1!KDIFF
                Else
                    rsAfc!KDIFF = rsAfc1!KDIFF
                End If
            Else
                rsAfc!KDIFF = 0
            End If
            
            If Not IsNull(rsAfc1!TDIFF) Then
                If Not IsNull(rsAfc!TDIFF) Then
                    rsAfc!TDIFF = rsAfc!TDIFF + rsAfc1!TDIFF
                Else
                    rsAfc!TDIFF = rsAfc1!TDIFF
                End If
            Else
                rsAfc!TDIFF = 0
            End If
            
            
            
            If Not IsNull(rsAfc1!NUMSKARTE) Then
                If Not IsNull(rsAfc!NUMSKARTE) Then
                    rsAfc!NUMSKARTE = rsAfc!NUMSKARTE + rsAfc1!NUMSKARTE
                Else
                    rsAfc!NUMSKARTE = rsAfc1!NUMSKARTE
                End If
            Else
                rsAfc!NUMSKARTE = 0
            End If
            
            If Not IsNull(rsAfc1!DUKA) Then
                If Not IsNull(rsAfc!DUKA) Then
                    rsAfc!DUKA = rsAfc!DUKA + rsAfc1!DUKA
                Else
                    rsAfc!DUKA = rsAfc1!DUKA
                End If
            Else
                rsAfc!DUKA = 0
            End If
    
            If Not IsNull(rsAfc1!BARVERKAUF) Then
                If Not IsNull(rsAfc!BARVERKAUF) Then
                    rsAfc!BARVERKAUF = rsAfc!BARVERKAUF + rsAfc1!BARVERKAUF
                Else
                    rsAfc!BARVERKAUF = rsAfc1!BARVERKAUF
                End If
            Else
                rsAfc!BARVERKAUF = 0
            End If
            
            If Not IsNull(rsAfc1!SCHVERKAUF) Then
                If Not IsNull(rsAfc!SCHVERKAUF) Then
                    rsAfc!SCHVERKAUF = rsAfc!SCHVERKAUF + rsAfc1!SCHVERKAUF
                Else
                    rsAfc!SCHVERKAUF = rsAfc1!SCHVERKAUF
                End If
            Else
                rsAfc!SCHVERKAUF = 0
            End If
            
            If Not IsNull(rsAfc1!TILGBAR) Then
                If Not IsNull(rsAfc!TILGBAR) Then
                    rsAfc!TILGBAR = rsAfc!TILGBAR + rsAfc1!TILGBAR
                Else
                    rsAfc!TILGBAR = rsAfc1!TILGBAR
                End If
            Else
                rsAfc!TILGBAR = 0
            End If
            
            If Not IsNull(rsAfc1!TILGSCH) Then
                If Not IsNull(rsAfc!TILGSCH) Then
                    rsAfc!TILGSCH = rsAfc!TILGSCH + rsAfc1!TILGSCH
                Else
                    rsAfc!TILGSCH = rsAfc1!TILGSCH
                End If
            Else
                rsAfc!TILGSCH = 0
            End If
            
            If Not IsNull(rsAfc1!TILGGUT) Then
                If Not IsNull(rsAfc!TILGGUT) Then
                    rsAfc!TILGGUT = rsAfc!TILGGUT + rsAfc1!TILGGUT
                Else
                    rsAfc!TILGGUT = rsAfc1!TILGGUT
                End If
            Else
                rsAfc!TILGGUT = 0
            End If
            
            If Not IsNull(rsAfc1!TILGKAR) Then
                If Not IsNull(rsAfc!TILGKAR) Then
                    rsAfc!TILGKAR = rsAfc!TILGKAR + rsAfc1!TILGKAR
                Else
                    rsAfc!TILGKAR = rsAfc1!TILGKAR
                End If
            Else
                rsAfc!TILGKAR = 0
            End If

            If Not IsNull(rsAfc1!EINRGUTSCH) Then
                If Not IsNull(rsAfc!EINRGUTSCH) Then
                    rsAfc!EINRGUTSCH = rsAfc!EINRGUTSCH + rsAfc1!EINRGUTSCH
                Else
                    rsAfc!EINRGUTSCH = rsAfc1!EINRGUTSCH
                End If
            Else
                rsAfc!EINRGUTSCH = 0
            End If
            
            If Not IsNull(rsAfc1!RESTGUTSCH) Then
                If Not IsNull(rsAfc!RESTGUTSCH) Then
                    rsAfc!RESTGUTSCH = rsAfc!RESTGUTSCH + rsAfc1!RESTGUTSCH
                Else
                    rsAfc!RESTGUTSCH = rsAfc1!RESTGUTSCH
                End If
            Else
                rsAfc!RESTGUTSCH = 0
            End If
            
            If Not IsNull(rsAfc1!AUSZGUTSCH) Then
                If Not IsNull(rsAfc!AUSZGUTSCH) Then
                    rsAfc!AUSZGUTSCH = rsAfc!AUSZGUTSCH + rsAfc1!AUSZGUTSCH
                Else
                    rsAfc!AUSZGUTSCH = rsAfc1!AUSZGUTSCH
                End If
            Else
                rsAfc!AUSZGUTSCH = 0
            End If
            rsAfc.Update
            rsAfc.Close
            
        rsAfc1.MoveNext
        Loop
    End If
    rsAfc1.Close
        
    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "AFCSTATPLUS"
    Fehler.gsFehlertext = "Beim Synchronisieren der Datenbanken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
   
End Sub

Public Sub AFCSTATPLUS_alleKassen(cFromTable As String, cIntoTable As String)
    On Error GoTo LOKAL_ERROR

    Dim rsAfc       As Recordset
    Dim rsAfc1      As Recordset
    Dim sSQL        As String
    Dim lDatum      As Long
    Dim cKaNum      As String
    

    sSQL = "Select * from " & cFromTable & "  "
    
        
    
    Set rsAfc1 = gdBase.OpenRecordset(sSQL)
        
    If Not rsAfc1.EOF Then
        rsAfc1.MoveFirst
        Do While Not rsAfc1.EOF
        
            If Not IsNull(rsAfc1!ADATE) Then
                lDatum = rsAfc1!ADATE
            Else
                lDatum = -1
            End If
            
            If Not IsNull(rsAfc1!kasnum) Then
                cKaNum = rsAfc1!kasnum
            Else
                cKaNum = -1
            End If
            
            
            sSQL = "Select * from " & cIntoTable & "  where KASNUM = " & cKaNum & " and ADATE = " & Trim$(Str$(lDatum)) & " "
            Set rsAfc = gdBase.OpenRecordset(sSQL)
        
            If Not rsAfc.EOF Then
                rsAfc.Edit 'gibt es in der großen Daba ein Eintrag
            Else
                rsAfc.AddNew 'gibt es in der großen Daba noch kein Eintrag
                rsAfc!ADATE = lDatum
                rsAfc!kasnum = cKaNum
            End If
            
            If Not IsNull(rsAfc1!UMS_BAR) Then 'Tagestabelle
                If Not IsNull(rsAfc!UMS_BAR) Then
                    rsAfc!UMS_BAR = rsAfc!UMS_BAR + rsAfc1!UMS_BAR
                Else
                    rsAfc!UMS_BAR = rsAfc1!UMS_BAR
                End If
            Else
                rsAfc!UMS_BAR = 0
            End If
            
            If Not IsNull(rsAfc1!UMS_Kred) Then
                If Not IsNull(rsAfc!UMS_Kred) Then
                    rsAfc!UMS_Kred = rsAfc!UMS_Kred + rsAfc1!UMS_Kred
                Else
                    rsAfc!UMS_Kred = rsAfc1!UMS_Kred
                End If
            Else
                rsAfc!UMS_Kred = 0
            End If
            
            '*******
            
            If Not IsNull(rsAfc1!UMS_SCHECK) Then
                If Not IsNull(rsAfc!UMS_SCHECK) Then
                    rsAfc!UMS_SCHECK = rsAfc!UMS_SCHECK + rsAfc1!UMS_SCHECK
                Else
                    rsAfc!UMS_SCHECK = rsAfc1!UMS_SCHECK
                End If
            Else
                rsAfc!UMS_SCHECK = 0
            End If
            
            If Not IsNull(rsAfc1!UMS_KARTE) Then
                If Not IsNull(rsAfc!UMS_KARTE) Then
                    rsAfc!UMS_KARTE = rsAfc!UMS_KARTE + rsAfc1!UMS_KARTE
                Else
                    rsAfc!UMS_KARTE = rsAfc1!UMS_KARTE
                End If
            Else
                rsAfc!UMS_KARTE = 0
            End If
            
            '***
            
            If Not IsNull(rsAfc1!UMS_LAST) Then
                If Not IsNull(rsAfc!UMS_LAST) Then
                    rsAfc!UMS_LAST = rsAfc!UMS_LAST + rsAfc1!UMS_LAST
                Else
                    rsAfc!UMS_LAST = rsAfc1!UMS_LAST
                End If
            Else
                rsAfc!UMS_LAST = 0
            End If
    
            If Not IsNull(rsAfc1!SPREIS_ANZ) Then
                If Not IsNull(rsAfc!SPREIS_ANZ) Then
                    rsAfc!SPREIS_ANZ = rsAfc!SPREIS_ANZ + rsAfc1!SPREIS_ANZ
                Else
                    rsAfc!SPREIS_ANZ = rsAfc1!SPREIS_ANZ
                End If
            Else
                rsAfc!SPREIS_ANZ = 0
            End If
            
            If Not IsNull(rsAfc1!SPREIS_GES) Then
                If Not IsNull(rsAfc!SPREIS_GES) Then
                    rsAfc!SPREIS_GES = rsAfc!SPREIS_GES + rsAfc1!SPREIS_GES
                Else
                    rsAfc!SPREIS_GES = rsAfc1!SPREIS_GES
                End If
            Else
                rsAfc!SPREIS_GES = 0
            End If
            
            If Not IsNull(rsAfc1!ANZSCHECK) Then
                If Not IsNull(rsAfc!ANZSCHECK) Then
                    rsAfc!ANZSCHECK = rsAfc!ANZSCHECK + rsAfc1!ANZSCHECK
                Else
                    rsAfc!ANZSCHECK = rsAfc1!ANZSCHECK
                End If
            Else
                rsAfc!ANZSCHECK = 0
            End If
            
            If Not IsNull(rsAfc1!Kundenzahl) Then
                If Not IsNull(rsAfc!Kundenzahl) Then
                    rsAfc!Kundenzahl = rsAfc!Kundenzahl + rsAfc1!Kundenzahl
                Else
                    rsAfc!Kundenzahl = rsAfc1!Kundenzahl
                End If
            Else
                rsAfc!Kundenzahl = 0
            End If
            
            If Not IsNull(rsAfc1!GELDFACH) Then
                If Not IsNull(rsAfc!GELDFACH) Then
                    rsAfc!GELDFACH = rsAfc!GELDFACH + rsAfc1!GELDFACH
                Else
                    rsAfc!GELDFACH = rsAfc1!GELDFACH
                End If
            Else
                rsAfc!GELDFACH = 0
            End If
                
            If Not IsNull(rsAfc1!ARTRAB_ANZ) Then
                If Not IsNull(rsAfc!ARTRAB_ANZ) Then
                    rsAfc!ARTRAB_ANZ = rsAfc!ARTRAB_ANZ + rsAfc1!ARTRAB_ANZ
                Else
                    rsAfc!ARTRAB_ANZ = rsAfc1!ARTRAB_ANZ
                End If
            Else
                rsAfc!ARTRAB_ANZ = 0
            End If
            
            If Not IsNull(rsAfc1!ARTRAB_GES) Then
                If Not IsNull(rsAfc!ARTRAB_GES) Then
                    rsAfc!ARTRAB_GES = rsAfc!ARTRAB_GES + rsAfc1!ARTRAB_GES
                Else
                    rsAfc!ARTRAB_GES = rsAfc1!ARTRAB_GES
                End If
            Else
                rsAfc!ARTRAB_GES = 0
            End If
            
            If Not IsNull(rsAfc1!GESRAB_ANZ) Then
                If Not IsNull(rsAfc!GESRAB_ANZ) Then
                    rsAfc!GESRAB_ANZ = rsAfc!GESRAB_ANZ + rsAfc1!GESRAB_ANZ
                Else
                    rsAfc!GESRAB_ANZ = rsAfc1!GESRAB_ANZ
                End If
            Else
                rsAfc!GESRAB_ANZ = 0
            End If
            
            If Not IsNull(rsAfc1!GESRAB_GES) Then
                If Not IsNull(rsAfc!GESRAB_GES) Then
                    rsAfc!GESRAB_GES = rsAfc!GESRAB_GES + rsAfc1!GESRAB_GES
                Else
                    rsAfc!GESRAB_GES = rsAfc1!GESRAB_GES
                End If
            Else
                rsAfc!GESRAB_GES = 0
            End If
            
            If Not IsNull(rsAfc1!STORNO_ANZ) Then
                If Not IsNull(rsAfc!STORNO_ANZ) Then
                    rsAfc!STORNO_ANZ = rsAfc!STORNO_ANZ + rsAfc1!STORNO_ANZ
                Else
                    rsAfc!STORNO_ANZ = rsAfc1!STORNO_ANZ
                End If
            Else
                rsAfc!STORNO_ANZ = 0
            End If
    
            If Not IsNull(rsAfc1!STORNO_GES) Then
                If Not IsNull(rsAfc!STORNO_GES) Then
                    rsAfc!STORNO_GES = rsAfc!STORNO_GES + rsAfc1!STORNO_GES
                Else
                    rsAfc!STORNO_GES = rsAfc1!STORNO_GES
                End If
            Else
                rsAfc!STORNO_GES = 0
            End If
            
            If Not IsNull(rsAfc1!EINZAHLUNG) Then
                If Not IsNull(rsAfc!EINZAHLUNG) Then
                    rsAfc!EINZAHLUNG = rsAfc!EINZAHLUNG + rsAfc1!EINZAHLUNG
                Else
                    rsAfc!EINZAHLUNG = rsAfc1!EINZAHLUNG
                End If
            Else
                rsAfc!EINZAHLUNG = 0
            End If
            
            If Not IsNull(rsAfc1!AUSZAHLUNG) Then
                If Not IsNull(rsAfc!AUSZAHLUNG) Then
                    rsAfc!AUSZAHLUNG = rsAfc!AUSZAHLUNG + rsAfc1!AUSZAHLUNG
                Else
                    rsAfc!AUSZAHLUNG = rsAfc1!AUSZAHLUNG
                End If
            Else
                rsAfc!AUSZAHLUNG = 0
            End If
            
            If Not IsNull(rsAfc1!GUTSCHEIN) Then
                If Not IsNull(rsAfc!GUTSCHEIN) Then
                    rsAfc!GUTSCHEIN = rsAfc!GUTSCHEIN + rsAfc1!GUTSCHEIN
                Else
                    rsAfc!GUTSCHEIN = rsAfc1!GUTSCHEIN
                End If
            Else
                rsAfc!GUTSCHEIN = 0
            End If
            
            If Not IsNull(rsAfc1!ZHLGGUTSCH) Then
                If Not IsNull(rsAfc!ZHLGGUTSCH) Then
                    rsAfc!ZHLGGUTSCH = rsAfc!ZHLGGUTSCH + rsAfc1!ZHLGGUTSCH
                Else
                    rsAfc!ZHLGGUTSCH = rsAfc1!ZHLGGUTSCH
                End If
            Else
                rsAfc!ZHLGGUTSCH = 0
            End If
            
            If Not IsNull(rsAfc1!BELEGNR) Then
                rsAfc!BELEGNR = rsAfc1!BELEGNR
            Else
                rsAfc!BELEGNR = 0
            End If
            
            If Not IsNull(rsAfc1!GUTSCHBAR) Then
                If Not IsNull(rsAfc!GUTSCHBAR) Then
                    rsAfc!GUTSCHBAR = rsAfc!GUTSCHBAR + rsAfc1!GUTSCHBAR
                Else
                    rsAfc!GUTSCHBAR = rsAfc1!GUTSCHBAR
                End If
            Else
                rsAfc!GUTSCHBAR = 0
            End If
            
            If Not IsNull(rsAfc1!GUTSCHSCH) Then
                If Not IsNull(rsAfc!GUTSCHSCH) Then
                    rsAfc!GUTSCHSCH = rsAfc!GUTSCHSCH + rsAfc1!GUTSCHSCH
                Else
                    rsAfc!GUTSCHSCH = rsAfc1!GUTSCHSCH
                End If
            Else
                rsAfc!GUTSCHSCH = 0
            End If
            
            If Not IsNull(rsAfc1!GUTSCHKRE) Then
                If Not IsNull(rsAfc!GUTSCHKRE) Then
                    rsAfc!GUTSCHKRE = rsAfc!GUTSCHKRE + rsAfc1!GUTSCHKRE
                Else
                    rsAfc!GUTSCHKRE = rsAfc1!GUTSCHKRE
                End If
            Else
                rsAfc!GUTSCHKRE = 0
            End If
            
            If Not IsNull(rsAfc1!GUTSCHKAR) Then
                If Not IsNull(rsAfc!GUTSCHKAR) Then
                    rsAfc!GUTSCHKAR = rsAfc!GUTSCHKAR + rsAfc1!GUTSCHKAR
                Else
                    rsAfc!GUTSCHKAR = rsAfc1!GUTSCHKAR
                End If
            Else
                rsAfc!GUTSCHKAR = 0
            End If
            
            If Not IsNull(rsAfc1!GUTSCHLAST) Then
                If Not IsNull(rsAfc!GUTSCHLAST) Then
                    rsAfc!GUTSCHLAST = rsAfc!GUTSCHLAST + rsAfc1!GUTSCHLAST
                Else
                    rsAfc!GUTSCHLAST = rsAfc1!GUTSCHLAST
                End If
            Else
                rsAfc!GUTSCHLAST = 0
            End If
            
            If Not IsNull(rsAfc1!GUTSCHGUTSCH) Then
                If Not IsNull(rsAfc!GUTSCHGUTSCH) Then
                    rsAfc!GUTSCHGUTSCH = rsAfc!GUTSCHGUTSCH + rsAfc1!GUTSCHGUTSCH
                Else
                    rsAfc!GUTSCHGUTSCH = rsAfc1!GUTSCHGUTSCH
                End If
            Else
                rsAfc!GUTSCHGUTSCH = 0
            End If
            
            
            
''                If Not IsNull(rsAfc1!WECHSEL) Then
''                    If Not IsNull(rsAfc!WECHSEL) Then
''                        rsAfc!WECHSEL = rsAfc!WECHSEL + rsAfc1!WECHSEL
''                    Else
''                        rsAfc!WECHSEL = rsAfc1!WECHSEL
''                    End If
''                Else
''                    rsAfc!WECHSEL = 0
''                End If
            
            
            
            If Not IsNull(rsAfc1!Abschopf) Then
                If Not IsNull(rsAfc!Abschopf) Then
                    rsAfc!Abschopf = rsAfc!Abschopf + rsAfc1!Abschopf
                Else
                    rsAfc!Abschopf = rsAfc1!Abschopf
                End If
            Else
                rsAfc!Abschopf = 0
            End If
            
            If Not IsNull(rsAfc1!KDIFF) Then
                If Not IsNull(rsAfc!KDIFF) Then
                    rsAfc!KDIFF = rsAfc!KDIFF + rsAfc1!KDIFF
                Else
                    rsAfc!KDIFF = rsAfc1!KDIFF
                End If
            Else
                rsAfc!KDIFF = 0
            End If
            
            If Not IsNull(rsAfc1!TDIFF) Then
                If Not IsNull(rsAfc!TDIFF) Then
                    rsAfc!TDIFF = rsAfc!TDIFF + rsAfc1!TDIFF
                Else
                    rsAfc!TDIFF = rsAfc1!TDIFF
                End If
            Else
                rsAfc!TDIFF = 0
            End If
            
            
            If Not IsNull(rsAfc1!NUMSKARTE) Then
                If Not IsNull(rsAfc!NUMSKARTE) Then
                    rsAfc!NUMSKARTE = rsAfc!NUMSKARTE + rsAfc1!NUMSKARTE
                Else
                    rsAfc!NUMSKARTE = rsAfc1!NUMSKARTE
                End If
            Else
                rsAfc!NUMSKARTE = 0
            End If
            
            If Not IsNull(rsAfc1!DUKA) Then
                If Not IsNull(rsAfc!DUKA) Then
                    rsAfc!DUKA = rsAfc!DUKA + rsAfc1!DUKA
                Else
                    rsAfc!DUKA = rsAfc1!DUKA
                End If
            Else
                rsAfc!DUKA = 0
            End If
    
            If Not IsNull(rsAfc1!BARVERKAUF) Then
                If Not IsNull(rsAfc!BARVERKAUF) Then
                    rsAfc!BARVERKAUF = rsAfc!BARVERKAUF + rsAfc1!BARVERKAUF
                Else
                    rsAfc!BARVERKAUF = rsAfc1!BARVERKAUF
                End If
            Else
                rsAfc!BARVERKAUF = 0
            End If
            
            If Not IsNull(rsAfc1!SCHVERKAUF) Then
                If Not IsNull(rsAfc!SCHVERKAUF) Then
                    rsAfc!SCHVERKAUF = rsAfc!SCHVERKAUF + rsAfc1!SCHVERKAUF
                Else
                    rsAfc!SCHVERKAUF = rsAfc1!SCHVERKAUF
                End If
            Else
                rsAfc!SCHVERKAUF = 0
            End If
            
            If Not IsNull(rsAfc1!TILGBAR) Then
                If Not IsNull(rsAfc!TILGBAR) Then
                    rsAfc!TILGBAR = rsAfc!TILGBAR + rsAfc1!TILGBAR
                Else
                    rsAfc!TILGBAR = rsAfc1!TILGBAR
                End If
            Else
                rsAfc!TILGBAR = 0
            End If
            
            If Not IsNull(rsAfc1!TILGSCH) Then
                If Not IsNull(rsAfc!TILGSCH) Then
                    rsAfc!TILGSCH = rsAfc!TILGSCH + rsAfc1!TILGSCH
                Else
                    rsAfc!TILGSCH = rsAfc1!TILGSCH
                End If
            Else
                rsAfc!TILGSCH = 0
            End If
            
            If Not IsNull(rsAfc1!TILGGUT) Then
                If Not IsNull(rsAfc!TILGGUT) Then
                    rsAfc!TILGGUT = rsAfc!TILGGUT + rsAfc1!TILGGUT
                Else
                    rsAfc!TILGGUT = rsAfc1!TILGGUT
                End If
            Else
                rsAfc!TILGGUT = 0
            End If
            
            If Not IsNull(rsAfc1!TILGKAR) Then
                If Not IsNull(rsAfc!TILGKAR) Then
                    rsAfc!TILGKAR = rsAfc!TILGKAR + rsAfc1!TILGKAR
                Else
                    rsAfc!TILGKAR = rsAfc1!TILGKAR
                End If
            Else
                rsAfc!TILGKAR = 0
            End If

            If Not IsNull(rsAfc1!EINRGUTSCH) Then
                If Not IsNull(rsAfc!EINRGUTSCH) Then
                    rsAfc!EINRGUTSCH = rsAfc!EINRGUTSCH + rsAfc1!EINRGUTSCH
                Else
                    rsAfc!EINRGUTSCH = rsAfc1!EINRGUTSCH
                End If
            Else
                rsAfc!EINRGUTSCH = 0
            End If
            
            If Not IsNull(rsAfc1!RESTGUTSCH) Then
                If Not IsNull(rsAfc!RESTGUTSCH) Then
                    rsAfc!RESTGUTSCH = rsAfc!RESTGUTSCH + rsAfc1!RESTGUTSCH
                Else
                    rsAfc!RESTGUTSCH = rsAfc1!RESTGUTSCH
                End If
            Else
                rsAfc!RESTGUTSCH = 0
            End If
            
            If Not IsNull(rsAfc1!AUSZGUTSCH) Then
                If Not IsNull(rsAfc!AUSZGUTSCH) Then
                    rsAfc!AUSZGUTSCH = rsAfc!AUSZGUTSCH + rsAfc1!AUSZGUTSCH
                Else
                    rsAfc!AUSZGUTSCH = rsAfc1!AUSZGUTSCH
                End If
            Else
                rsAfc!AUSZGUTSCH = 0
            End If
            rsAfc.Update
            rsAfc.Close
            
        rsAfc1.MoveNext
        Loop
    End If
    rsAfc1.Close
        
    Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "AFCSTATPLUS_alleKassen"
    Fehler.gsFehlertext = "Beim Synchronisieren der Datenbanken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
End Sub
Public Function SynAbschieb(cKasnum As String, cKasstab As String) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim sPfad       As String
    Dim slokalPfad  As String
    Dim lokalDB     As Database
    Dim rsAfcbuch   As Recordset
    Dim sArtnr      As String
    Dim iMenge      As Integer
    Dim rsAfc       As Recordset
    Dim rsAfc1      As Recordset
    Dim rsKun       As Recordset
    Dim lDatum      As Long
    Dim cKundnr     As String
    Dim cPreis      As String
    Dim stimz       As String
    Dim cBedz       As String
    Dim i           As Integer

    lDatum = Fix(Now)
    
    SynAbschieb = False
    
    If Not NewTableSuchenDBKombi(cKasstab & "KJ", gdBase) Then
        Exit Function
    End If
    
    If Not NewTableSuchenDBKombi(cKasstab & "KB", gdBase) Then
        Exit Function
    End If
    
    If Not NewTableSuchenDBKombi(cKasstab & "AFCB", gdBase) Then
        Exit Function
    End If
    
    If Not NewTableSuchenDBKombi(cKasstab & "STAT", gdBase) Then
        Exit Function
    End If
    
    If Not NewTableSuchenDBKombi(cKasstab & "KOLL", gdBase) Then
        Exit Function
    End If
    
    If Not NewTableSuchenDBKombi(cKasstab & "KRED", gdBase) Then
        Exit Function
    End If
    
    If Not NewTableSuchenDBKombi(cKasstab & "RET", gdBase) Then
        Exit Function
    End If
    
    Do While IsAktionZulaessig(cKasstab & "Kdat") = False
    
    Loop
    
    frmWKL00.txtStatus.Text = "0"
    frmWKL00.picprogress.Visible = True
    frmWKL00.Label2.Visible = True
    
    
    
    
    
   
    frmWKL00.txtStatus.Text = "15"
    
    frmWKL00!Label2.Caption = "Datenbank: Bestand minimieren"
    frmWKL00!Label2.Refresh
    
    
    'Bestand minimieren

    Set rsAfcbuch = gdBase.OpenRecordset(cKasstab & "AFCB", dbOpenTable)
    If Not rsAfcbuch.EOF Then
        i = 0
        rsAfcbuch.MoveFirst
        Do While Not rsAfcbuch.EOF
            If i = 100 Then
                i = 0
            End If
            
            i = i + 1
            frmWKL00.txtStatus.Text = i
            
            If Not IsNull(rsAfcbuch!aartnr) Then
                sArtnr = rsAfcbuch!aartnr
            Else
                sArtnr = ""
            End If
            
            If Not IsNull(rsAfcbuch!aMenge) Then
                iMenge = rsAfcbuch!aMenge
            Else
                iMenge = 0
            End If
            
            If Not IsNull(rsAfcbuch!AZEIT) Then
                stimz = rsAfcbuch!AZEIT
            Else
                stimz = ""
            End If
            
            If Not IsNull(rsAfcbuch!abednu) Then
                cBedz = rsAfcbuch!abednu
            Else
                cBedz = ""
            End If
            
            If cBedz = "" Then cBedz = "0"
            
            If iMenge > 0 Then
            'menge Positiv
            'Dann Kassiervorgang
                BestandsminiOrMaxi sArtnr, CLng(iMenge), "Kassiervorgang", stimz, cBedz, gcFilNr
            Else
            'Menge Negativ
            'Dann Storno
                BestandsminiOrMaxi sArtnr, CLng(iMenge), "Storno Kasse LM", stimz, cBedz, gcFilNr
            End If
            rsAfcbuch.MoveNext
        Loop
    End If
    rsAfcbuch.Close
    
    frmWKL00.txtStatus.Text = "20"
    
    frmWKL00!Label2.Caption = "Datenbank: Kunden ..."
    frmWKL00!Label2.Refresh
    
    'Kunden
    loeschNEW "Kunden1", gdBase
    
    sSQL = "Select * into Kunden1 from " & cKasstab & "AFCB "
    sSQL = sSQL & " where akunum > 0 "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    frmWKL00.txtStatus.Text = "30"
    
    Set rsKun = gdBase.OpenRecordset("select * from Kunden1")

    If Not rsKun.EOF Then
        rsKun.MoveFirst
        Do While Not rsKun.EOF
            cKundnr = rsKun!AKUNUM
            cPreis = rsKun!APREIS
            cPreis = fnMoveComma2Point$(cPreis)
            sSQL = "Update Kunden set BONUS = BONUS + " & cPreis & " where KUNDNR = " & cKundnr
            schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
            rsKun.MoveNext
        Loop
    End If
    
    rsKun.Close: Set rsKun = Nothing
    
    frmWKL00.txtStatus.Text = "40"
    
    frmWKL00!Label2.Caption = "Datenbank: AFCBUCH ..."
    frmWKL00!Label2.Refresh
    
    'afcbuch
    sSQL = "Insert into afcbuch select * from " & cKasstab & "AFCB "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    frmWKL00.txtStatus.Text = "50"
    'Kassjour
    frmWKL00!Label2.Caption = "Datenbank: Kassjour aktualisieren..."
    frmWKL00!Label2.Refresh
    
    sSQL = "Insert into Kassjour select * from " & cKasstab & "KJ "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    

    
    frmWKL00.txtStatus.Text = "60"
    
    frmWKL00!Label2.Caption = "Datenbank: Kredit ..."
    frmWKL00!Label2.Refresh
    
    'Kredit
    sSQL = "Insert into Kredit select * from " & cKasstab & "KRED "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    frmWKL00.txtStatus.Text = "70"
    
    frmWKL00!Label2.Caption = "Datenbank: Kollverk ..."
    frmWKL00!Label2.Refresh
    
    'Kollverk
    sSQL = "Insert into Kollverk select * from " & cKasstab & "KOLL "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    
    frmWKL00.txtStatus.Text = "72"
    
    frmWKL00!Label2.Caption = "Datenbank: Retoure ..."
    frmWKL00!Label2.Refresh
    
    'Retoure
    sSQL = "Insert into Retoure select * from " & cKasstab & "RET "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    frmWKL00.txtStatus.Text = "75"
    
    frmWKL00!Label2.Caption = "Datenbank: Kassbon ..."
    frmWKL00!Label2.Refresh
    
    'Kassbon
    sSQL = "Insert into Kassbon select * from " & cKasstab & "KB "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
        
    frmWKL00.txtStatus.Text = "80"
    
    sSQL = "DELETE from " & cKasstab & "KB "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    frmWKL00.txtStatus.Text = "81"
    
    sSQL = "DELETE from " & cKasstab & "KJ "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    frmWKL00.txtStatus.Text = "82"
    
    sSQL = "DELETE from " & cKasstab & "AFCB "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    frmWKL00.txtStatus.Text = "83"
    
    sSQL = "DELETE from " & cKasstab & "KRED "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    frmWKL00.txtStatus.Text = "84"
    
    sSQL = "DELETE from " & cKasstab & "RET "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    frmWKL00.txtStatus.Text = "85"
    
    sSQL = "DELETE from " & cKasstab & "KOLL "
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    
    frmWKL00.txtStatus.Text = "100"
    
    frmWKL00!Label2.Caption = ""
    frmWKL00!Label2.Refresh
    
        
     
    
    frmWKL00.txtStatus.Text = "0"
    frmWKL00.picprogress.Visible = False
    frmWKL00.Label2.Visible = False
    
    AktionAustragen cKasstab & "Kdat"
    
    SynAbschieb = True
    
    Exit Function
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "synAbschieb"
    Fehler.gsFehlertext = "Beim Synchronisieren der Datenbanken ist ein Fehler aufgetreten."
    
    Fehlermeldung1

End Function
Public Sub BESTANDartAndRkzj()

    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    Screen.MousePointer = 11
    
    loeschNEW "RKZART", gdBase
    CreateTable "RKZART", gdBase
    
    cSQL = "Insert into RKZART Select"
    cSQL = cSQL & " ARTIKEL.ARTNR "
    cSQL = cSQL & " , ARTIKEL.BEZEICH "
    cSQL = cSQL & " , artlief.LIBESNR "
    cSQL = cSQL & " , ARTIKEL.BESTAND "
    cSQL = cSQL & " , ARTIKEL.KVKPR1 "
    cSQL = cSQL & " , artlief.LINR "
    cSQL = cSQL & " , ARTIKEL.LPZ "
    cSQL = cSQL & " , artlief.LEKPR "
    'MussRKZ
    cSQL = cSQL & " from ARTIKEL inner join artlief on ARTIKEL.artnr = artlief.artnr where ARTIKEL.BESTAND > 0 "
    cSQL = cSQL & " and artlief.RKZ = 'J' "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Create index linr on RKZART(linr) "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update RKZART inner join LISRT on RKZART.Linr = LISRT.Linr "
    cSQL = cSQL & " set RKZART.LIEFBEZ = LISRT.LIEFBEZ "
    gdBase.Execute cSQL, dbFailOnError
    
    
    cSQL = "Update RKZART inner join linbez on RKZART.Linr = linbez.Linr and RKZART.LPZ = linbez.LPZ "
    cSQL = cSQL & " set RKZART.Marke = linbez.marke "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update RKZART inner join linbez on RKZART.Linr = linbez.Linr and RKZART.LPZ = linbez.LPZ "
    cSQL = cSQL & " set RKZART.linbez = linbez.linbezeich "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update RKZART inner join LAGERPLATZ on RKZART.ARTNR = LAGERPLATZ.ARTNR "
    cSQL = cSQL & " set RKZART.LAGERP = LAGERPLATZ.LAGERP "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update RKZART set LAGERP = 0 where lagerp is null "
    gdBase.Execute cSQL, dbFailOnError
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "BESTANDartAndRkzj"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Public Sub BESTANDartAndEKNull()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    Screen.MousePointer = 11
    
    loeschNEW "EKNULLART", gdBase
    CreateTableT2 "EKNULLART", gdBase
    
    loeschNEW "LEKNULL", gdBase
    cSQL = "Select "
    cSQL = cSQL & " ARTNR "
    cSQL = cSQL & " , LINR "
    cSQL = cSQL & " , LIBESNR "
    cSQL = cSQL & " into LEKNULL from ARTLIEF where LEKPR = 0 "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into EKNULLART Select"
    cSQL = cSQL & " a.ARTNR "
    cSQL = cSQL & " , a.BEZEICH "
    cSQL = cSQL & " , b.LIBESNR "
    cSQL = cSQL & " , a.BESTAND "
    cSQL = cSQL & " , a.KVKPR1 "
    cSQL = cSQL & " , b.LINR "
    cSQL = cSQL & " , a.LPZ "
    
    cSQL = cSQL & " from ARTIKEL a inner join LEKNULL b "
    cSQL = cSQL & " on a.artnr = b.artnr "
    cSQL = cSQL & " where a.BESTAND > 0 "
    
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Create index linr on EKNULLART(linr) "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update EKNULLART inner join LISRT on EKNULLART.Linr = LISRT.Linr "
    cSQL = cSQL & " set EKNULLART.LIEFBEZ = LISRT.LIEFBEZ "
    gdBase.Execute cSQL, dbFailOnError

    cSQL = "Update EKNULLART inner join linbez on EKNULLART.Linr = linbez.Linr and EKNULLART.LPZ = linbez.LPZ "
    cSQL = cSQL & " set EKNULLART.Marke = linbez.marke "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update EKNULLART inner join linbez on EKNULLART.Linr = linbez.Linr and EKNULLART.LPZ = linbez.LPZ "
    cSQL = cSQL & " set EKNULLART.linbez = linbez.linbezeich "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update EKNULLART inner join LAGERPLATZ on EKNULLART.ARTNR = LAGERPLATZ.ARTNR "
    cSQL = cSQL & " set EKNULLART.LAGERP = LAGERPLATZ.LAGERP "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update EKNULLART set LAGERP = 0 where lagerp is null "
    gdBase.Execute cSQL, dbFailOnError
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "BESTANDartAndEKNull"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Public Sub BESTANDartAndSchnittEKNull()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    Screen.MousePointer = 11
    
    loeschNEW "EKNULLART", gdBase
    CreateTableT2 "EKNULLART", gdBase
    
    cSQL = "Update ARTIKEL set ekpr = 0 where ekpr is null "
    gdBase.Execute cSQL, dbFailOnError
    
    
    
    cSQL = "Insert into EKNULLART Select"
    cSQL = cSQL & " a.ARTNR "
    cSQL = cSQL & " , a.BEZEICH "
    cSQL = cSQL & " , '' as LIBESNR "
    cSQL = cSQL & " , a.BESTAND "
    cSQL = cSQL & " , a.KVKPR1 "
    cSQL = cSQL & " , 0 as LINR "
    cSQL = cSQL & " , a.LPZ "
    
    cSQL = cSQL & " from ARTIKEL a "
    cSQL = cSQL & " where a.BESTAND > 0 and a.ekpr = 0 "
    gdBase.Execute cSQL, dbFailOnError
    
'    cSQL = "Create index linr on EKNULLART(linr) "
'    gdBase.Execute cSQL, dbFailOnError
'
'    cSQL = "Update EKNULLART inner join LISRT on EKNULLART.Linr = LISRT.Linr "
'    cSQL = cSQL & " set EKNULLART.LIEFBEZ = LISRT.LIEFBEZ "
'    gdBase.Execute cSQL, dbFailOnError
'
'    cSQL = "Update EKNULLART inner join linbez on EKNULLART.Linr = linbez.Linr and EKNULLART.LPZ = linbez.LPZ "
'    cSQL = cSQL & " set EKNULLART.Marke = linbez.marke "
'    gdBase.Execute cSQL, dbFailOnError
'
'    cSQL = "Update EKNULLART inner join linbez on EKNULLART.Linr = linbez.Linr and EKNULLART.LPZ = linbez.LPZ "
'    cSQL = cSQL & " set EKNULLART.linbez = linbez.linbezeich "
'    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update EKNULLART inner join LAGERPLATZ on EKNULLART.ARTNR = LAGERPLATZ.ARTNR "
    cSQL = cSQL & " set EKNULLART.LAGERP = LAGERPLATZ.LAGERP "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update EKNULLART set LAGERP = 0 where lagerp is null "
    gdBase.Execute cSQL, dbFailOnError
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "BESTANDartAndSchnittEKNull"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Public Sub NegBESTANDartAndEKNull()
    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    Screen.MousePointer = 11
    
    loeschNEW "EKNULLART", gdBase
    CreateTableT2 "EKNULLART", gdBase
    
    loeschNEW "LEKNULL", gdBase
    cSQL = "Select "
    cSQL = cSQL & " ARTNR "
    cSQL = cSQL & " , LINR "
    cSQL = cSQL & " , LIBESNR "
    cSQL = cSQL & " into LEKNULL from ARTLIEF where LEKPR = 0 "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Insert into EKNULLART Select"
    cSQL = cSQL & " a.ARTNR "
    cSQL = cSQL & " , a.BEZEICH "
    cSQL = cSQL & " , b.LIBESNR "
    cSQL = cSQL & " , a.BESTAND "
    cSQL = cSQL & " , a.KVKPR1 "
    cSQL = cSQL & " , b.LINR "
    cSQL = cSQL & " , a.LPZ "
    
    cSQL = cSQL & " from ARTIKEL a inner join LEKNULL b "
    cSQL = cSQL & " on a.artnr = b.artnr "
    cSQL = cSQL & " where a.BESTAND < 0 "
    
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Create index linr on EKNULLART(linr) "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update EKNULLART inner join LISRT on EKNULLART.Linr = LISRT.Linr "
    cSQL = cSQL & " set EKNULLART.LIEFBEZ = LISRT.LIEFBEZ "
    gdBase.Execute cSQL, dbFailOnError
    
    
    cSQL = "Update EKNULLART inner join linbez on EKNULLART.Linr = linbez.Linr and EKNULLART.LPZ = linbez.LPZ "
    cSQL = cSQL & " set EKNULLART.Marke = linbez.marke "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update EKNULLART inner join linbez on EKNULLART.Linr = linbez.Linr and EKNULLART.LPZ = linbez.LPZ "
    cSQL = cSQL & " set EKNULLART.linbez = linbez.linbezeich "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update EKNULLART inner join LAGERPLATZ on EKNULLART.ARTNR = LAGERPLATZ.ARTNR "
    cSQL = cSQL & " set EKNULLART.LAGERP = LAGERPLATZ.LAGERP "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update EKNULLART set LAGERP = 0 where lagerp is null "
    gdBase.Execute cSQL, dbFailOnError
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "NegBESTANDartAndEKNull"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Public Sub BESTANDartAndgefuehrtN()

    On Error GoTo LOKAL_ERROR
    
    Dim cSQL As String
    
    Screen.MousePointer = 11
    
    loeschNEW "GEFART", gdBase
    CreateTableT2 "GEFART", gdBase
    
    cSQL = "Insert into GEFART Select"
    cSQL = cSQL & " ARTNR "
    cSQL = cSQL & " , BEZEICH "
    cSQL = cSQL & " , LIBESNR "
    cSQL = cSQL & " , BESTAND "
    cSQL = cSQL & " , KVKPR1 "
    cSQL = cSQL & " , LINR "
    cSQL = cSQL & " , LPZ "
    cSQL = cSQL & " , val(awm) as farbnr "
    
    cSQL = cSQL & " from ARTIKEL where BESTAND > 0 "
    cSQL = cSQL & " and gefuehrt = 'N' "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Create index linr on GEFART(linr) "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update GEFART inner join LISRT on GEFART.Linr = LISRT.Linr "
    cSQL = cSQL & " set GEFART.LIEFBEZ = LISRT.LIEFBEZ "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update GEFART inner join linbez on GEFART.Linr = linbez.Linr and GEFART.LPZ = linbez.LPZ "
    cSQL = cSQL & " set GEFART.Marke = linbez.marke "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update GEFART inner join linbez on GEFART.Linr = linbez.Linr and GEFART.LPZ = linbez.LPZ "
    cSQL = cSQL & " set GEFART.linbez = linbez.linbezeich "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update GEFART inner join LAGERPLATZ on GEFART.ARTNR = LAGERPLATZ.ARTNR "
    cSQL = cSQL & " set GEFART.LAGERP = LAGERPLATZ.LAGERP "
    gdBase.Execute cSQL, dbFailOnError
    
    cSQL = "Update GEFART set LAGERP = 0 where lagerp is null "
    gdBase.Execute cSQL, dbFailOnError
    
    BringFarbeInsSpiel "GEFART", gdBase

    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:

    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "BESTANDartAndgefuehrtN"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Public Function LoescheTagesAbschlussMODUL7(cKasse As String, picprogress As PictureBox, txtStatus As TextBox _
, lbl6 As Label, lbl3 As Label, lbl1 As Label, Label1 As Label) As Boolean
    On Error GoTo LOKAL_ERROR
    
    LoescheTagesAbschlussMODUL7 = False
    
    Dim cSQL        As String
    Dim rsrs        As Recordset
    Dim rsRs2       As Recordset
    Dim rsZiel      As Recordset
    
    Dim iCount      As Integer
    Dim ctmp        As String
    Dim lDatum      As Long
    Dim lKundenZahl As Long
    Dim dWert       As Double
    Dim dWert2      As Double
    Dim iRet        As Integer
    Dim iFileNr     As Integer
    Dim cProtokoll  As String
    Dim cQuelle     As String
    Dim cZiel       As String
    Dim lRet        As Long
    Dim lfail       As Long
    
   
    Dim cKdnr As String
    Dim cLinr As String
    Dim cArtNr As String
    Dim lJahr As Long
    Dim lMonat As Long
    Dim dUmsatz As Double
    Dim lAnzahl As Long
    Dim lcount As Long
    Dim counter As Long
    Dim cex    As String
    Dim cPfad23 As String

    Dim iStep   As Integer
    
    Dim cAKTKW          As String 'Kalenderwoche
    Dim cGESpKW         As String
    Dim DateHeut        As Date
    Dim DateGespeich    As Date

    schreibeProtokoll "************************************"
    schreibeProtokoll "Starte Kassenabschluss"
    ZentraleWillsWissen "Starte Kassenabschluss"
    '***************************************************************************
    '* Sicherung der Daten AFCBUCH, UMSATZ
    '***************************************************************************
    picprogress.Visible = True
    iStep = 1
    txtStatus.Text = iStep * 2
    

    lbl6.Caption = "Verarbeite Tagesdaten"
    lbl6.Refresh: schreibeProtokoll lbl6.Caption
    
    lbl3.Caption = "1"
    lbl3.Refresh
    lbl1.Caption = "38"
    lbl1.Refresh
    
    Label1.Caption = "von"
    Label1.Refresh
    
    
    
    cex = Format(DateValue(Now), "DD")
    
    
    lbl6.Caption = "Erstelle Sicherung der Tabelle AFCBUCH"
    lbl6.Refresh: schreibeProtokoll lbl6.Caption
    
    lbl3.Caption = "2"
    lbl3.Refresh
      
    iStep = 2
    txtStatus.Text = iStep * 2
    
    cQuelle = "AFCBUCH"
    cZiel = "afcS" & cex
    
    loeschNEW cZiel, gdBase
    
    iStep = 3
    txtStatus.Text = iStep * 2
    
    cSQL = "Select * into " & cZiel & " from " & cQuelle & " "
    gdBase.Execute cSQL, dbFailOnError
    
    iStep = 4
    txtStatus.Text = iStep * 2

    lbl6.Caption = "Erstelle Sicherung der Tabelle UMSATZ"
    lbl6.Refresh: schreibeProtokoll lbl6.Caption
    
    lbl3.Caption = "3"
    lbl3.Refresh
      
    cQuelle = "UMSATZ"
    cZiel = "umsS" & cex
    
    loeschNEW cZiel, gdBase
    
    iStep = 5
    txtStatus.Text = iStep * 2
    
    cSQL = "Select * into " & cZiel & " from " & cQuelle & " "
    gdBase.Execute cSQL, dbFailOnError

    'Zur Überwachung wird ein Protokoll geschrieben, wer wann den Tagesabschluß
    'durchgeführt hat
    
    iStep = 6
    txtStatus.Text = iStep * 2

    lbl6.Caption = "Schreibe in das Protokoll "
    lbl6.Refresh: schreibeProtokoll lbl6.Caption: schreibeProtokoll lbl6.Caption
    lbl3.Caption = "4"
    lbl3.Refresh
    
    lbl6.Caption = "Bearbeite die Tabelle ARTLIEF"
    lbl6.Refresh: schreibeProtokoll lbl6.Caption
    lbl3.Caption = "5"
    lbl3.Refresh

    cSQL = "Select * from ARTLIEF where SYNSTATUS = 'D'"
    FnOpenrecordset rsrs, cSQL, 3, gdBase
    
    iStep = 7
    txtStatus.Text = iStep * 2
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            rsrs.delete
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
'    lbl6.Caption = "Bearbeite die Tabelle ARTIKEL"
'    lbl6.Refresh: schreibeProtokoll lbl6.Caption
'    lbl3.Caption = "6"
'    lbl3.Refresh
'
'    If gibt_es_etwas_zum_löschen > 0 Then
'
'        schreibeProtokoll "löschen wird ausgeführt"
'        cSQL = "Select * from Artikel where SYNSTATUS = 'D'"
'        FnOpenrecordset rsrs, cSQL, 3, gdBase
'
'        iStep = 8
'        txtStatus.Text = iStep * 2
'
'        If Not rsrs.EOF Then
'            rsrs.MoveFirst
'            Do While Not rsrs.EOF
'                rsrs.delete
'                rsrs.MoveNext
'            Loop
'        End If
'        rsrs.Close: Set rsrs = Nothing
'
'    End If
    
    
    lbl6.Caption = "Bearbeite die Tabelle LISRT"
    lbl6.Refresh: schreibeProtokoll lbl6.Caption
    lbl3.Caption = "7"
    lbl3.Refresh
    

    cSQL = "Select * from Lisrt where SYNSTATUS = 'D'"
    FnOpenrecordset rsrs, cSQL, 3, gdBase
    iStep = 9
    txtStatus.Text = iStep * 2
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            rsrs.delete
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    lbl6.Caption = "Bearbeite die Tabelle BEDNAME"
    lbl6.Refresh: schreibeProtokoll lbl6.Caption
    lbl3.Caption = "7"
    lbl3.Refresh
    

    
'    cSQL = "Select * from BEDNAME where SYNSTATUS = 'D'"
'    FnOpenrecordset rsrs, cSQL, 3, gdBase
'    iStep = 10
'    txtStatus.Text = iStep * 2
'    If Not rsrs.EOF Then
'        rsrs.MoveFirst
'        Do While Not rsrs.EOF
'            rsrs.delete
'            rsrs.MoveNext
'        Loop
'    End If
'    rsrs.Close: Set rsrs = Nothing
    
    cSQL = "Select * from BEDNAME where SYNSTATUS = 'A'"
    FnOpenrecordset rsrs, cSQL, 2, gdBase
    
    iStep = 11
    txtStatus.Text = iStep * 2
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            rsrs.Edit
            rsrs!SYNStatus = Null
            rsrs.Update
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    cSQL = "Select * from BEDNAME where SYNSTATUS = 'E'"
    FnOpenrecordset rsrs, cSQL, 2, gdBase
    

    iStep = 12
    txtStatus.Text = iStep * 2
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            rsrs.Edit
            rsrs!SYNStatus = Null
            rsrs.Update
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
    lbl6.Caption = "Bearbeite Tabelle Gutsch delete"
    lbl6.Refresh: schreibeProtokoll lbl6.Caption
    lbl3.Caption = "8"
    lbl3.Refresh
    
    'Hier wird die Gutsch bearbeitet
    cSQL = "Select * from gutsch where STATUS = 'L'"
    FnOpenrecordset rsrs, cSQL, 3, gdBase
    iStep = 13
    txtStatus.Text = iStep * 2
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            rsrs.delete
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    lbl6.Caption = "Bearbeite Tabelle Gutsch komplett"
    lbl6.Refresh: schreibeProtokoll lbl6.Caption
    lbl3.Caption = "9"
    lbl3.Refresh
    
    cSQL = "Select * from gutsch "
    FnOpenrecordset rsrs, cSQL, 2, gdBase
    

    iStep = 14
    txtStatus.Text = iStep * 2
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            rsrs.Edit
            rsrs!Status = "N"
            rsrs.Update
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
        
    'Hier wird die Kunden bearbeitet
    
    lbl6.Caption = "Bearbeite Tabelle Kunden delete"
    lbl6.Refresh: schreibeProtokoll lbl6.Caption
    lbl3.Caption = "10"
    lbl3.Refresh
    
    
    cSQL = "Select * from KUNDEN where STATUS = 'D'"
    FnOpenrecordset rsrs, cSQL, 3, gdBase
    iStep = 15
    txtStatus.Text = iStep * 2
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            rsrs.delete
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    cSQL = "Select * from KUNDEN where SYNSTATUS = 'D'"
    FnOpenrecordset rsrs, cSQL, 3, gdBase
    iStep = 16
    txtStatus.Text = iStep * 2
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            rsrs.delete
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    lbl6.Caption = "Bearbeite Tabelle Kunden komplett"
    lbl6.Refresh: schreibeProtokoll lbl6.Caption
    lbl3.Caption = "11"
    lbl3.Refresh
    

    
    cSQL = "Select * from KUNDEN where STATUS = 'E'"
    FnOpenrecordset rsrs, cSQL, 2, gdBase
    iStep = 17
    txtStatus.Text = iStep * 2
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            
            rsrs.Edit
            rsrs!Status = "N"
            rsrs!TBONUS = "0"
            rsrs.Update
            
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    lbl6.Caption = "Bearbeite Tabelle Kunden komplett"
    lbl6.Refresh: schreibeProtokoll lbl6.Caption
    lbl3.Caption = "12"
    lbl3.Refresh
    
    cSQL = "Select * from KUNDEN where STATUS = 'A'"
    FnOpenrecordset rsrs, cSQL, 2, gdBase
    iStep = 18
    txtStatus.Text = iStep * 2
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            
            rsrs.Edit
            rsrs!Status = "N"
            rsrs!TBONUS = "0"
            rsrs.Update
            
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    cSQL = "Select * from KUNDEN where SYNSTATUS = 'E'"
    FnOpenrecordset rsrs, cSQL, 2, gdBase
    iStep = 19
    txtStatus.Text = iStep * 2
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            
            rsrs.Edit
            rsrs!SYNStatus = Null
            rsrs.Update
            
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    cSQL = "Select * from KUNDEN where SYNSTATUS = 'A'"
    FnOpenrecordset rsrs, cSQL, 2, gdBase
    iStep = 20
    txtStatus.Text = iStep * 2
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            
            rsrs.Edit
            rsrs!SYNStatus = Null
            rsrs.Update
            
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    

    lbl6.Caption = "Bearbeite Tabelle AFCSTATP "
    lbl6.Refresh: schreibeProtokoll lbl6.Caption
    lbl3.Caption = "13"
    lbl3.Refresh
    
    AfcstatIstNull "AFCSTAT"
    
    InsertAnsKassBuch cKasse
    
    cSQL = "Update EINAUSKB set sendok = true where kasnum = " & cKasse
    gdBase.Execute cSQL, dbFailOnError
    

    AFCSTATPLUS "AFCSTAT", "AFCSTATP", cKasse

    
    
    '*** erster Schritt: Umsatzdaten für volle MWST ermitteln ***
    iStep = 22
    txtStatus.Text = iStep * 2
    lbl6.Caption = "Umsatzdaten für volle MWST ermitteln"
    lbl6.Refresh: schreibeProtokoll lbl6.Caption
    lbl3.Caption = "14"
    lbl3.Refresh
    
    Dim dNichtUmsGutschbetrag As Double
    dNichtUmsGutschbetrag = 0
    
    If gbGutscheinBeiVKversteuern = True Then
    
        
        cSQL = "Select SUM(Wert) as UMSATZ from Gemischte_Z where kasnum = " & gcKasNum
        cSQL = cSQL & " and Thema = 'nicht ums GUTSCHBETRAG'"
        Set rsrs = gdBase.OpenRecordset(cSQL)
        If Not rsrs.EOF Then
            rsrs.MoveFirst
            If Not IsNull(rsrs!UMSATZ) Then
                dNichtUmsGutschbetrag = rsrs!UMSATZ
            End If
        End If
        rsrs.Close: Set rsrs = Nothing
        
    
    
    
    
    
    
    
    
        cSQL = "Select ADATE as DATUM, SUM(APREIS) as UMSG1, SUM(APREIS) as UMSV1 "
        cSQL = cSQL & ",0 as UMSO1"
        cSQL = cSQL & ",0 as UMSE1, 0 as KUNZ1, SUM(ALEKPR * AMENGE) as EKPR1, 0 as KRED1 "
        cSQL = cSQL & "from AFCBUCH where AMWSK = 'V' and KASNUM = " & cKasse & " and UMS_OK <> 'N' group by ADATE"
        
        
        
        
        
    Else

        cSQL = "Select ADATE as DATUM, SUM(APREIS) as UMSG1, SUM(APREIS) as UMSV1 "
        cSQL = cSQL & ",0 as UMSO1"
        cSQL = cSQL & ",0 as UMSE1, 0 as KUNZ1, SUM(ALEKPR * AMENGE) as EKPR1, 0 as KRED1 "
        cSQL = cSQL & "from AFCBUCH where AMWSK = 'V' and KASNUM = " & cKasse & " and AARTNR <> 666666 and UMS_OK <> 'N' group by ADATE"
    End If

    FnOpenrecordset rsrs, cSQL, 1, gdBase
    iStep = 23
    txtStatus.Text = iStep * 2
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!Datum) Then
                ctmp = rsrs!Datum
            Else
                ctmp = "0"
            End If
            lDatum = DateValue(ctmp)
            cSQL = "Select * from UMSATZ where DATUM = " & Trim$(Str$(lDatum)) & " "
            FnOpenrecordset rsZiel, cSQL, 2, gdBase
            If rsZiel.EOF Then
                rsZiel.AddNew
                rsZiel!Datum = rsrs!Datum
                rsZiel!UMSG1 = rsrs!UMSG1 - dNichtUmsGutschbetrag
                rsZiel!UMSV1 = rsrs!UMSV1 - dNichtUmsGutschbetrag
                rsZiel!UMSe1 = rsrs!UMSe1
                rsZiel!UMSo1 = rsrs!UMSo1
                rsZiel!KUNZ1 = rsrs!KUNZ1
                rsZiel!EKPR1 = rsrs!EKPR1
                rsZiel!KRED1 = rsrs!KRED1
            Else
                rsZiel.Edit
                rsZiel!Datum = rsrs!Datum
                
                If Not IsNull(rsZiel!UMSo1) Then
                    rsZiel!UMSo1 = rsZiel!UMSo1 + rsrs!UMSo1
                Else
                    rsZiel!UMSo1 = rsrs!UMSo1
                End If
                
                If Not IsNull(rsZiel!UMSG1) Then
                    rsZiel!UMSG1 = rsZiel!UMSG1 + rsrs!UMSG1 - dNichtUmsGutschbetrag
                Else
                    rsZiel!UMSG1 = rsrs!UMSG1 - dNichtUmsGutschbetrag
                End If
                
                 
                
                If Not IsNull(rsZiel!UMSV1) Then
                    rsZiel!UMSV1 = rsZiel!UMSV1 + rsrs!UMSV1 - dNichtUmsGutschbetrag
                Else
                    rsZiel!UMSV1 = rsrs!UMSV1 - dNichtUmsGutschbetrag
                End If
                
                If Not IsNull(rsZiel!UMSe1) Then
                    rsZiel!UMSe1 = rsZiel!UMSe1 + rsrs!UMSe1
                Else
                    rsZiel!UMSe1 = rsrs!UMSe1
                End If
                
                If Not IsNull(rsZiel!KUNZ1) Then
                    rsZiel!KUNZ1 = rsZiel!KUNZ1 + rsrs!KUNZ1
                Else
                    rsZiel!KUNZ1 = rsrs!KUNZ1
                End If
                
                If Not IsNull(rsZiel!EKPR1) Then
                    rsZiel!EKPR1 = rsZiel!EKPR1 + rsrs!EKPR1
                Else
                    rsZiel!EKPR1 = rsrs!EKPR1
                End If
                
                If Not IsNull(rsZiel!KRED1) Then
                    rsZiel!KRED1 = rsZiel!KRED1 + rsrs!KRED1
                Else
                    rsZiel!KRED1 = rsrs!KRED1
                End If
            End If
            rsZiel.Update
            rsZiel.Close: Set rsZiel = Nothing
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    '*** zweiter Schritt: Umsatzdaten für ermäßigte MWST ermitteln ***

    lbl6.Caption = "Umsatzdaten für ermäßigte MWST ermitteln"
    lbl6.Refresh: schreibeProtokoll lbl6.Caption
    lbl3.Caption = "15"
    lbl3.Refresh
    
    
    
    
    If gbGutscheinBeiVKversteuern = True Then
        cSQL = "Select ADATE as DATUM, SUM(APREIS) as UMSG1, 0 as UMSV1 "
        cSQL = cSQL & ",0 as UMSO1"
        cSQL = cSQL & ",SUM(APREIS) as UMSE1, 0 as KUNZ1, SUM(ALEKPR * AMENGE) as EKPR1, 0 as KRED1 "
        cSQL = cSQL & " from AFCBUCH where AMWSK = 'E' and KASNUM = " & cKasse & " and UMS_OK <> 'N' group by ADATE"
    Else

        cSQL = "Select ADATE as DATUM, SUM(APREIS) as UMSG1, 0 as UMSV1 "
        cSQL = cSQL & ",0 as UMSO1"
        cSQL = cSQL & ",SUM(APREIS) as UMSE1, 0 as KUNZ1, SUM(ALEKPR * AMENGE) as EKPR1, 0 as KRED1 "
        cSQL = cSQL & " from AFCBUCH where AMWSK = 'E' and KASNUM = " & cKasse & " and AARTNR <> 666666 and UMS_OK <> 'N' group by ADATE"
    End If
    
    
    

    
    FnOpenrecordset rsrs, cSQL, 1, gdBase
    
    iStep = 24
    txtStatus.Text = iStep * 2
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!Datum) Then
                ctmp = rsrs!Datum
            Else
                ctmp = "0"
            End If
            lDatum = DateValue(ctmp)
            cSQL = "Select * from UMSATZ where DATUM = " & Trim$(Str$(lDatum)) & " "
            FnOpenrecordset rsZiel, cSQL, 2, gdBase

            If rsZiel.EOF Then
                rsZiel.AddNew
                rsZiel!Datum = rsrs!Datum
                rsZiel!UMSG1 = rsrs!UMSG1
                rsZiel!UMSV1 = rsrs!UMSV1
                rsZiel!UMSe1 = rsrs!UMSe1
                rsZiel!UMSo1 = rsrs!UMSo1
                rsZiel!KUNZ1 = rsrs!KUNZ1
                rsZiel!EKPR1 = rsrs!EKPR1
                rsZiel!KRED1 = rsrs!KRED1
            Else
                rsZiel.Edit
                rsZiel!Datum = rsrs!Datum
                
                If Not IsNull(rsZiel!UMSo1) Then
                    rsZiel!UMSo1 = rsZiel!UMSo1 + rsrs!UMSo1
                Else
                    rsZiel!UMSo1 = rsrs!UMSo1
                End If
                
                If Not IsNull(rsZiel!UMSG1) Then
                    rsZiel!UMSG1 = rsZiel!UMSG1 + rsrs!UMSG1
                Else
                    rsZiel!UMSG1 = rsrs!UMSG1
                End If
                
                If Not IsNull(rsZiel!UMSV1) Then
                    rsZiel!UMSV1 = rsZiel!UMSV1 + rsrs!UMSV1
                Else
                    rsZiel!UMSV1 = rsrs!UMSV1
                End If
                
                If Not IsNull(rsZiel!UMSe1) Then
                    rsZiel!UMSe1 = rsZiel!UMSe1 + rsrs!UMSe1
                Else
                    rsZiel!UMSe1 = rsrs!UMSe1
                End If
                
                If Not IsNull(rsZiel!KUNZ1) Then
                    rsZiel!KUNZ1 = rsZiel!KUNZ1 + rsrs!KUNZ1
                Else
                    rsZiel!KUNZ1 = rsrs!KUNZ1
                End If
                
                If Not IsNull(rsZiel!EKPR1) Then
                    rsZiel!EKPR1 = rsZiel!EKPR1 + rsrs!EKPR1
                Else
                    rsZiel!EKPR1 = rsrs!EKPR1
                End If
                
                If Not IsNull(rsZiel!KRED1) Then
                    rsZiel!KRED1 = rsZiel!KRED1 + rsrs!KRED1
                Else
                    rsZiel!KRED1 = rsrs!KRED1
                End If
            End If
            rsZiel.Update
            rsZiel.Close: Set rsZiel = Nothing
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    '*** zweiter Schritt, Teil 2: Umsatzdaten ohne MWST ermitteln ***
    


    lbl6.Caption = "Umsatzdaten ohne MWST ermitteln"
    lbl6.Refresh: schreibeProtokoll lbl6.Caption
    lbl3.Caption = "16"
    lbl3.Refresh
    
    
    
    If gbGutscheinBeiVKversteuern = True Then
        cSQL = "Select ADATE as DATUM, SUM(APREIS) as UMSG1, 0 as UMSV1 "
        cSQL = cSQL & ",SUM(APREIS) as UMSO1"
        cSQL = cSQL & ",0 as UMSE1, 0 as KUNZ1, SUM(ALEKPR * AMENGE) as EKPR1, 0 as KRED1 "
        cSQL = cSQL & " from AFCBUCH where AMWSK = 'O' and KASNUM = " & cKasse & " and UMS_OK <> 'N' group by ADATE"
    Else

        cSQL = "Select ADATE as DATUM, SUM(APREIS) as UMSG1, 0 as UMSV1 "
        cSQL = cSQL & ",SUM(APREIS) as UMSO1"
        cSQL = cSQL & ",0 as UMSE1, 0 as KUNZ1, SUM(ALEKPR * AMENGE) as EKPR1, 0 as KRED1 "
        cSQL = cSQL & " from AFCBUCH where AMWSK = 'O' and KASNUM = " & cKasse & " and AARTNR <> 666666 and UMS_OK <> 'N' group by ADATE"
    End If
    
    
    FnOpenrecordset rsrs, cSQL, 1, gdBase
    
    iStep = 25
    txtStatus.Text = iStep * 2
    
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!Datum) Then
                ctmp = rsrs!Datum
            Else
                ctmp = "0"
            End If
            lDatum = DateValue(ctmp)
            cSQL = "Select * from UMSATZ where DATUM = " & Trim$(Str$(lDatum)) & " "
            FnOpenrecordset rsZiel, cSQL, 2, gdBase

            If rsZiel.EOF Then
                rsZiel.AddNew
                rsZiel!Datum = rsrs!Datum
                rsZiel!UMSG1 = rsrs!UMSG1
                rsZiel!UMSV1 = rsrs!UMSV1
                rsZiel!UMSe1 = rsrs!UMSe1
                rsZiel!UMSo1 = rsrs!UMSo1
                rsZiel!KUNZ1 = rsrs!KUNZ1
                rsZiel!EKPR1 = rsrs!EKPR1
                rsZiel!KRED1 = rsrs!KRED1
            Else
                rsZiel.Edit
                rsZiel!Datum = rsrs!Datum
                
                If Not IsNull(rsZiel!UMSo1) Then
                    rsZiel!UMSo1 = rsZiel!UMSo1 + rsrs!UMSo1
                Else
                    rsZiel!UMSo1 = rsrs!UMSo1
                End If
                
                If Not IsNull(rsZiel!UMSG1) Then
                    rsZiel!UMSG1 = rsZiel!UMSG1 + rsrs!UMSG1
                Else
                    rsZiel!UMSG1 = rsrs!UMSG1
                End If
                
                If Not IsNull(rsZiel!KUNZ1) Then
                    rsZiel!KUNZ1 = rsZiel!KUNZ1 + rsrs!KUNZ1
                Else
                    rsZiel!KUNZ1 = rsrs!KUNZ1
                End If
                
                If Not IsNull(rsZiel!EKPR1) Then
                    rsZiel!EKPR1 = rsZiel!EKPR1 + rsrs!EKPR1
                Else
                    rsZiel!EKPR1 = rsrs!EKPR1
                End If
                
                If Not IsNull(rsZiel!KRED1) Then
                    rsZiel!KRED1 = rsZiel!KRED1 + rsrs!KRED1
                Else
                    rsZiel!KRED1 = rsrs!KRED1
                End If
            End If
            rsZiel.Update
            rsZiel.Close: Set rsZiel = Nothing
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    '*** dritter Schritt: Kreditbeträge ermitteln ***
    
    lbl6.Caption = "Kreditbeträge ermitteln"
    lbl6.Refresh: schreibeProtokoll lbl6.Caption
    lbl3.Caption = "17"
    lbl3.Refresh

    cSQL = "Select ADATE as DATUM, SUM(APREIS) as KRED1 "
    cSQL = cSQL & "from AFCBUCH where KK_ART = 'KR' and KASNUM = " & cKasse & " group by ADATE"
    FnOpenrecordset rsrs, cSQL, 1, gdBase
    iStep = 26
    txtStatus.Text = iStep * 2
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!Datum) Then
                ctmp = rsrs!Datum
            Else
                ctmp = "0"
            End If
            lDatum = DateValue(ctmp)
            
            cSQL = "Select * from UMSATZ where DATUM = " & Trim$(Str$(lDatum)) & " "
            FnOpenrecordset rsZiel, cSQL, 2, gdBase
            If Not rsZiel.EOF Then
                rsZiel.Edit
                If Not IsNull(rsZiel!KRED1) Then
                    rsZiel!KRED1 = rsZiel!KRED1 + rsrs!KRED1
                Else
                    rsZiel!KRED1 = rsrs!KRED1
                End If
            Else
                rsZiel.AddNew
                rsZiel!KRED1 = rsrs!KRED1
            End If
            rsZiel.Update
            rsZiel.Close: Set rsZiel = Nothing
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    'Kundenzahlen ermitteln anhand AFCSTAT->KUNDENZAHL
    

    lbl6.Caption = "Kundenzahlen ermitteln"
    lbl6.Refresh: schreibeProtokoll lbl6.Caption
    lbl3.Caption = "18"
    lbl3.Refresh
    
    cSQL = "Select KUNDENZAHL, ADATE from AFCSTAT where KASNUM = " & cKasse & " order by ADATE"
    FnOpenrecordset rsrs, cSQL, 1, gdBase
    
    iStep = 27
    txtStatus.Text = iStep * 2
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!Kundenzahl) Then
                lKundenZahl = rsrs!Kundenzahl
            Else
                lKundenZahl = 0
            End If
            
            If Not IsNull(rsrs!ADATE) Then
                lDatum = rsrs!ADATE
            Else
                lDatum = 0
            End If
            
            cSQL = "Select * from UMSATZ where DATUM = " & Trim$(Str$(lDatum)) & " "
            FnOpenrecordset rsZiel, cSQL, 2, gdBase
            
            If Not rsZiel.EOF Then
                rsZiel.MoveFirst
                rsZiel.Edit
            Else
                rsZiel.AddNew
            End If
            
            If Not IsNull(rsZiel!KUNZ1) Then
                rsZiel!KUNZ1 = rsZiel!KUNZ1 + lKundenZahl
            Else
                rsZiel!KUNZ1 = lKundenZahl
            End If
            rsZiel.Update
            rsZiel.Close: Set rsZiel = Nothing
            rsrs.MoveNext
        Loop
    End If
    
    rsrs.Close: Set rsrs = Nothing

    lbl6.Caption = "Artikel - Jahres - Umsatzzahlen ermitteln"
    lbl6.Refresh: schreibeProtokoll lbl6.Caption
    lbl3.Caption = "19"
    lbl3.Refresh
    
    If NewTableSuchenDBKombi("UMSARTJ", gdBase) = False Then
        schreibeProtokoll "ERROR Tabelle UMSARTJ muss neu erstellt werden."
        UmsartjNew lbl6
    End If
    
    cSQL = "Insert into UMSARTJ Select "
    cSQL = cSQL & " afcbuch.aartnr as artnr, "
    cSQL = cSQL & " year(afcbuch.adate) as Jahr, "
    cSQL = cSQL & " sum(afcbuch.apreis) as umsatzj, "
    cSQL = cSQL & " sum(afcbuch.amenge) as anzahlj "
    cSQL = cSQL & " from AFCBUCH Where"
    cSQL = cSQL & " ( "
    cSQL = cSQL & " afcbuch.KASNUM = " & cKasse & " "
    
    
    If gbGutscheinBeiVKversteuern = True Then
        cSQL = cSQL & "  "
    Else
        cSQL = cSQL & " and afcbuch.aartnr <> 666666 "
    End If
    
    
    cSQL = cSQL & " ) "
    cSQL = cSQL & " group by afcbuch.aartnr, year(afcbuch.adate)"
    gdBase.Execute cSQL, dbFailOnError
    
    iStep = 28
    txtStatus.Text = iStep * 2
    lbl6.Caption = "Verarbeitung der Umsatzzahlen (Artikel)"
    lbl6.Refresh: schreibeProtokoll lbl6.Caption
    lbl3.Caption = "20"
    lbl3.Refresh
    
    iStep = 29
    txtStatus.Text = iStep * 2
    
    loeschNEW "UMSTEMP", gdBase
    
    iStep = 30
    txtStatus.Text = iStep * 2
    
    cSQL = "Select artnr, Jahr, sum(umsartj.umsatzj) as umsatzj"
    cSQL = cSQL & " , sum(umsartj.anzahlj) as anzahlj into UMStemp from UMSARTJ "
    cSQL = cSQL & " group by artnr, Jahr "
    gdBase.Execute cSQL, dbFailOnError
    
    iStep = 31
    txtStatus.Text = iStep * 2
    
    lbl6.Caption = "Sichern der Umsatzzahlen (Artikel)"
    lbl6.Refresh: schreibeProtokoll lbl6.Caption
    lbl3.Caption = "21"
    lbl3.Refresh
    
    loeschNEW "UMSARTJ", gdBase
    
    iStep = 32
    txtStatus.Text = iStep * 2
    
    cSQL = "Select * into UMSARTJ from Umstemp"
    gdBase.Execute cSQL, dbFailOnError
    
    iStep = 33
    txtStatus.Text = iStep * 2
    
    lbl6.Caption = "Sortieren der Tabelle(Artikel)"
    lbl6.Refresh: schreibeProtokoll lbl6.Caption
    lbl3.Caption = "22"
    lbl3.Refresh
    
    cSQL = "Create Index PRIMKEY on UMSARTJ (ARTNR, JAHR)"
    gdBase.Execute cSQL, dbFailOnError
    
    iStep = 34
    txtStatus.Text = iStep * 2
    
    cSQL = "Create Index JAHR on UMSARTJ (JAHR)"
    gdBase.Execute cSQL, dbFailOnError
    
    iStep = 35
    txtStatus.Text = iStep * 2
    
    cSQL = "Create Index ARTNR on UMSARTJ (ARTNR)"
    gdBase.Execute cSQL, dbFailOnError
    
    iStep = 36
    txtStatus.Text = iStep * 2
    
    'Ums_art
    lbl6.Caption = "Artikel - Monats - Umsatzzahlen ermitteln"
    lbl6.Refresh: schreibeProtokoll lbl6.Caption
    lbl3.Caption = "23"
    lbl3.Refresh

'alte methode

    If gbGutscheinBeiVKversteuern = True Then
        cSQL = "Select * from AFCBUCH where KASNUM = " & cKasse & "  "
    Else
        cSQL = "Select * from AFCBUCH where KASNUM = " & cKasse & " and afcbuch.aartnr <> 666666 "
    End If
    
    
    FnOpenrecordset rsrs, cSQL, 1, gdBase
    iStep = 37
    txtStatus.Text = iStep * 2
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
            If Not IsNull(rsrs!aMenge) Then
                lAnzahl = rsrs!aMenge
            Else
                lAnzahl = 0
            End If
            
            If Not IsNull(rsrs!APREIS) Then
                dUmsatz = rsrs!APREIS
            Else
                dUmsatz = 0
            End If
            
            If Not IsNull(rsrs!aartnr) Then
                cArtNr = rsrs!aartnr
            Else
                cArtNr = "-1"
            End If
            
            If Not IsNull(rsrs!AKUNUM) Then
                cKdnr = rsrs!AKUNUM
            Else
                cKdnr = "0"
            End If
            
            If Not IsNull(rsrs!linr) Then
                cLinr = rsrs!linr
            Else
                cLinr = "0"
            End If
            
            If Not IsNull(rsrs!ADATE) Then
                lJahr = Year(rsrs!ADATE)
                lMonat = Month(rsrs!ADATE)
            Else
                lJahr = 0
                lMonat = 0
            End If
            
            'ARTIKELUMSÄTZE
            cSQL = "Select * from UMS_ART where ARTNR = " & cArtNr & " and JAHR = " & Trim$(Str$(lJahr)) & " and MONAT = " & Trim$(Str$(lMonat)) & " "
            FnOpenrecordset rsRs2, cSQL, 2, gdBase
            
            If Not rsRs2.EOF Then
                rsRs2.Edit
                If Not IsNull(rsRs2!UMSATZ) Then
                    rsRs2!UMSATZ = rsRs2!UMSATZ + dUmsatz
                Else
                    rsRs2!UMSATZ = dUmsatz
                End If
                If Not IsNull(rsRs2!ANZAHL) Then
                    rsRs2!ANZAHL = rsRs2!ANZAHL + lAnzahl
                Else
                    rsRs2!ANZAHL = lAnzahl
                End If
                rsRs2.Update
            Else
                rsRs2.AddNew
                rsRs2!artnr = Val(cArtNr)
                rsRs2!jahr = lJahr
                rsRs2!Monat = lMonat
                rsRs2!UMSATZ = dUmsatz
                rsRs2!ANZAHL = lAnzahl
                rsRs2.Update
            End If
            rsRs2.Close: Set rsRs2 = Nothing: Set rsRs2 = Nothing
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    'Umskdj

    iStep = 38
    txtStatus.Text = iStep * 2
    
    If Val(gcFilNr) = 0 Then
        cSQL = "update alterg set SENDOK = True where SENDOK = False"
        gdBase.Execute cSQL, dbFailOnError
    End If

    lbl6.Caption = "Kunden - Jahres - Umsatzzahlen ermitteln"
    lbl6.Refresh: schreibeProtokoll lbl6.Caption
    lbl3.Caption = "24"
    lbl3.Refresh
    
    iStep = 39
    txtStatus.Text = iStep * 2
    
    cSQL = "Insert into UMSKDJ Select "
    cSQL = cSQL & " afcbuch.akunum as kundnr, "
    cSQL = cSQL & " year(afcbuch.adate) as Jahr, "
    cSQL = cSQL & " sum(afcbuch.apreis) as umsatzj, "
    cSQL = cSQL & " sum(afcbuch.amenge) as anzahlj "
    cSQL = cSQL & " from AFCBUCH Where"
    cSQL = cSQL & " ( "
    cSQL = cSQL & " afcbuch.KASNUM = " & cKasse & " "
    
    
    If gbGutscheinBeiVKversteuern = True Then
        cSQL = cSQL & "  "
    Else
        cSQL = cSQL & " and afcbuch.aartnr <> 666666 "
    End If
    
    cSQL = cSQL & " ) "
    cSQL = cSQL & " group by afcbuch.aKunum, year(afcbuch.adate)"
    gdBase.Execute cSQL, dbFailOnError
    
    iStep = 40
    txtStatus.Text = iStep * 2
    lbl6.Caption = "Verarbeitung der Umsatzzahlen (Kunden)"
    lbl6.Refresh: schreibeProtokoll lbl6.Caption
    lbl3.Caption = "25"
    lbl3.Refresh
    
    loeschNEW "UMSTEMP", gdBase
    
    iStep = 41
    txtStatus.Text = iStep * 2
    
    cSQL = "Select kundnr, Jahr, sum(UMSKDJ.umsatzj) as umsatzj"
    cSQL = cSQL & " , sum(UMSKDJ.anzahlj) as anzahlj into UMStemp from UMSKDJ "
    cSQL = cSQL & " group by kundnr, Jahr "
    gdBase.Execute cSQL, dbFailOnError
    
    iStep = 42
    txtStatus.Text = iStep * 2
    
    lbl6.Caption = "Sichern der Umsatzzahlen (Kunden)"
    lbl6.Refresh: schreibeProtokoll lbl6.Caption
    lbl3.Caption = "26"
    lbl3.Refresh
    
    loeschNEW "UMSKDJ", gdBase
    
    iStep = 43
    txtStatus.Text = iStep * 2
    
    cSQL = "Select * into UMSKDJ from Umstemp"
    gdBase.Execute cSQL, dbFailOnError
    
    iStep = 44
    txtStatus.Text = iStep * 2
    
    lbl6.Caption = "Sortieren der Tabelle(Kunden)"
    lbl6.Refresh: schreibeProtokoll lbl6.Caption
    lbl3.Caption = "27"
    lbl3.Refresh
    
    cSQL = "Create Index PRIMKEY on UMSKDJ (KUNDNR, JAHR)"
    gdBase.Execute cSQL, dbFailOnError

    cSQL = "Create Index JAHR on UMSKDJ (JAHR)"
    gdBase.Execute cSQL, dbFailOnError

    cSQL = "Create Index KUNDNR on UMSKDJ (KUNDNR)"
    gdBase.Execute cSQL, dbFailOnError
    
    frmWKL20.Timer1.Enabled = False
    
    'für eine bestimmte Kassennummer
'    schreibeProtokollKassStopnurfürdieseKasse "Kassenabschluss"
    
    'betrifft alle Kassen
    schreibeProtokollKassStopfürALLEKassen "Kassenabschluss"

    'Tabelle AFCSTAT mit KASNUM löschen
    
    iStep = 45
    txtStatus.Text = iStep * 2
    
    lbl6.Caption = "Bearbeite Tabelle AFCSTAT "
    lbl6.Refresh: schreibeProtokoll lbl6.Caption
    lbl3.Caption = "34"
    lbl3.Refresh

    cSQL = "Delete from AFCSTAT where KASNUM = " & cKasse & " "
    gdBase.Execute cSQL, dbFailOnError
    
    lbl6.Caption = "Bearbeite Tabelle AFCBuch "
    lbl6.Refresh: schreibeProtokoll lbl6.Caption
    lbl3.Caption = "35"
    lbl3.Refresh
    
    If gsWAAGE <> "keine Waage" Then
        insertKasswaag cKasse
    End If
    
    
    
    
    
    
    cSQL = "Delete from AFCBuch where KASNUM = " & cKasse & " "
    gdBase.Execute cSQL, dbFailOnError
    
    'noch sichern
    
    cSQL = "Insert into Gemischte_ZP Select * from Gemischte_Z where KASNUM = " & cKasse & " "
    gdBase.Execute cSQL, dbFailOnError
    
    
    cSQL = "Delete from Gemischte_Z where KASNUM = " & cKasse & " "
    gdBase.Execute cSQL, dbFailOnError
    

    
'    rechneNeuKunden
'    ermBestMitarbeiter
    
    cPfad23 = gcDBPfad               'Datenbankpfad
    If Right(cPfad23, 1) <> "\" Then
        cPfad23 = cPfad23 & "\"
    End If
'    Kill cPfad23 & "KASSSTOP" & Trim(gcKasNum) & ".TXT"
    
    Kill cPfad23 & "KASSSTOP_ALLE.TXT"
    
    iStep = 46
    txtStatus.Text = iStep * 2
    
    For iCount = 0 To 9
        frmWKL21.Label3(iCount).Caption = "0,00 " & gcWaehrung
    Next iCount
    
    frmWKL21.Label3(10).Caption = "0"
    frmWKL21.Label3(11).Caption = "0,00 " & gcWaehrung
    frmWKL21.Label3(12).Caption = "0,00 " & gcWaehrung
    frmWKL21.Label3(13).Caption = "0"
    frmWKL21.Label3(14).Caption = "0"
    frmWKL21.Label3(15).Caption = "0,00 " & gcWaehrung
    frmWKL21.Label3(16).Caption = "0,00 " & gcWaehrung
    frmWKL21.Label3(17).Caption = "0"
    frmWKL21.Label3(18).Caption = "0,00 " & gcWaehrung
    frmWKL21.Label3(19).Caption = "0"
    frmWKL21.Label3(20).Caption = "0,00 " & gcWaehrung
    frmWKL21.Label3(21).Caption = "0,00 " & gcWaehrung
    frmWKL21.Label3(22).Caption = "0,00 " & gcWaehrung
    frmWKL21.Label3(23).Caption = "0,00 " & gcWaehrung
    frmWKL21.Label3(24).Caption = "0,00 " & gcWaehrung
    frmWKL21.Label3(25).Caption = "0,00 " & gcWaehrung
    frmWKL21.Label3(26).Caption = "0,00 " & gcWaehrung
    frmWKL21.Label3(27).Caption = "0,00 " & gcWaehrung
    frmWKL21.Label3(28).Caption = "0,00 " & gcWaehrung
    frmWKL21.Label3(29).Caption = "0,00 " & gcWaehrung
    frmWKL21.Label3(30).Caption = "0,00 " & gcWaehrung
    frmWKL21.Label3(31).Caption = "0,00 " & gcWaehrung
    
    iStep = 47
    txtStatus.Text = iStep * 2

    If PrüfdateforBestand Then
        lbl6.Caption = "Die monatlichen Inventurdaten werden geschrieben."
        lbl6.Refresh: schreibeProtokoll lbl6.Caption
        lbl3.Caption = "36"
        lbl3.Refresh
        SchreibeMonatsArtikelBestände
    End If
    iStep = 48
    txtStatus.Text = iStep * 2
    
    'Zur Überwachung wird ein Protokoll geschrieben, wer wann den Tagesabschluß
    'durchgeführt hat
    

    lbl6.Caption = "Bearbeite Tabelle KAEINAUS "
    lbl6.Refresh: schreibeProtokoll lbl6.Caption
    lbl3.Caption = "37"
    lbl3.Refresh

    cSQL = "Delete from KAEINAUS where kasnum = " & cKasse
    gdBase.Execute cSQL, dbFailOnError
    
    iStep = 49
    txtStatus.Text = iStep * 2
    
    lbl6.Caption = "Bearbeite Tabelle KAEINAUS "
    lbl6.Refresh: schreibeProtokoll lbl6.Caption
    lbl3.Caption = "37a"
    lbl3.Refresh
    
    If gbUnistatWeek Then 'Teilnahme an Wochenauswertung
        DateGespeich = DateValue(gdateStatlast): DateHeut = DateValue(Now): cAKTKW = DatePart("ww", DateHeut)
        If CInt(cAKTKW) = 1 Then
            cAKTKW = "53"
        Else
            cAKTKW = CInt(cAKTKW) - 1
        End If
        
        cGESpKW = DatePart("ww", DateGespeich)
        
        If CInt(cGESpKW) = 1 Then
            cGESpKW = "53"
        Else
            cGESpKW = CInt(cGESpKW) - 1
        End If
        
        If CInt(cAKTKW) <> CInt(cGESpKW) Then 'Vergleich
            If Trim(gsStatkundnr) = "" Then gsStatkundnr = "XXX" 'Kisskundennummer?
            
            lbl6.Caption = "Die Wochenstatistik wird erstellt."
            lbl6.Refresh: schreibeProtokoll lbl6.Caption
            lbl3.Caption = "38"
            lbl3.Refresh
    
            If unistatweek(frmWKL00.txtStatus, frmWKL00.picprogress) Then Label1.Caption = DatumLastSuniW:
            
'            If unistatweek_new(frmWKL00.txtStatus, frmWKL00.picprogress) Then Label1.Caption = DatumLastSuniW:
                
            
        End If
    End If
    
    lbl6.Caption = "Bearbeite Tabelle KAEINAUS "
    lbl6.Refresh: schreibeProtokoll lbl6.Caption
    lbl3.Caption = "37b"
    lbl3.Refresh
    
    

    
    If gbUnistatMonat Then 'Teilnahme an Monatsauswertung
        DateGespeich = DateValue(gdateMStatlast)
        DateHeut = DateValue(Now)
        
        If DateGespeich < DateHeut Then
            If Month(DateGespeich) = Month(DateHeut) Then
            
            Else
                If Trim(gsMStatkundnr) = "" Then gsMStatkundnr = "XXX" 'Kisskundennummer?
                
                lbl6.Caption = "Die Monatsstatistik wird erstellt."
                lbl6.Refresh: schreibeProtokoll lbl6.Caption
                lbl3.Caption = "38"
                lbl3.Refresh
        
                If unistatMonat Then Label1.Caption = DatumLastSuniM:
            End If
        End If
    End If
    
    lbl6.Caption = "Bearbeite Tabelle KAEINAUS "
    lbl6.Refresh: schreibeProtokoll lbl6.Caption
    lbl3.Caption = "37c"
    lbl3.Refresh
    
    If leseGFKstat("email") <> "" Then 'Teilnahme an GFK per Mail Auswertung für Cospar
    
        Dim slastAuswertkw As String
        Dim slastAuswertDatStempel As String
        slastAuswertDatStempel = leseGFKstat("lastdate")
        If slastAuswertDatStempel = "" Then
            slastAuswertkw = DatePart("ww", DateValue(Now))
            slastAuswertkw = CInt(slastAuswertkw) - 1
        Else
            slastAuswertkw = DatePart("ww", DateValue(slastAuswertDatStempel))
        End If
        
        Dim sAuswertwoche As String
        sAuswertwoche = DatePart("ww", DateValue(Now))
       
        If CInt(slastAuswertkw) < CInt(sAuswertwoche) Then
            
            Dim iAuswertjahr As Integer
            iAuswertjahr = DatePart("yyyy", DateValue(Now))
            
            If CInt(sAuswertwoche) = 1 Then
                sAuswertwoche = "52"
                iAuswertjahr = iAuswertjahr - 1
            Else
                sAuswertwoche = CInt(sAuswertwoche) - 1
            End If
            
            
            Dim sKUNDNR As String
            sKUNDNR = leseGFKstat("kundnr")
            
            Dim sGFKMAIL As String
            sGFKMAIL = leseGFKstat("email")
        
            GFKerstellen Trim(sAuswertwoche), iAuswertjahr, sKUNDNR, True, Trim(sGFKMAIL)
        End If
    End If
    
    lbl6.Caption = "Bearbeite Tabelle KAEINAUS "
    lbl6.Refresh: schreibeProtokoll lbl6.Caption
    lbl3.Caption = "37d"
    lbl3.Refresh
    
    Dim slastAuswertTag As String
    Dim sBisAuswertTAG As String
    Dim l As Long
    Dim lminAuswerttag As Long
    Dim lmaxAuswerttag As Long
    Dim sSQL As String
    
    
    
    

    
    If NewTableSuchenDBKombi("VEDESSTAT", gdBase) Then

        If CBool(leseVEDESstat("live")) = True Then 'Teilnahme an VEDES - Abverkäufen
            If leseVEDESstat("marktnr") <> "" Then

                slastAuswertTag = leseVEDESstat("lastdate")

                If slastAuswertTag = "" Then
                    slastAuswertTag = DateValue("15.02.2016")
                End If

                sBisAuswertTAG = DateValue(Now) - 1

                If CLng(DateValue(slastAuswertTag)) < CLng(DateValue(sBisAuswertTAG)) Then
                    lminAuswerttag = CLng(DateValue(slastAuswertTag)) + 1
                    lmaxAuswerttag = CLng(DateValue(sBisAuswertTAG))
                    For l = lminAuswerttag To lmaxAuswerttag
                        VEDES_AUSW_erstellen l
                    Next l

                    VEDES_AUSW_uebertragen

                    sSQL = "Update VEDESSTAT Set LASTDATE = " & lmaxAuswerttag & " "
                    gdBase.Execute sSQL, dbFailOnError

                End If
            End If
        End If

    End If
    
    iStep = 50
    txtStatus.Text = iStep * 2
    
    lbl6.Caption = "Der Tagesabschluss wurde erfolgreich durchgeführt."
    lbl6.Refresh: schreibeProtokoll lbl6.Caption
    lbl3.Caption = ""
    lbl3.Refresh
    lbl1.Caption = ""
    lbl1.Refresh
    
    Label1.Caption = ""
    Label1.Refresh
    
    ZentraleWillsWissen "Tagesabschluss fertig"
    
    AktionAustragen "Kassenabschluss"
    
    
    Dim i As Integer
    If gbKSF = True Or gbBestAkt = True Then 'Kassendatei sofort versenden!
    
        theBigFTPFehlerZähler = 0
        theBigFTPFehler = False
        
        Dim bmerke As Boolean
        bmerke = gbFTPautomatic
    
        gbFTPautomatic = True
        KassendatundStatcheck 'FTP Transfer bei Statistik oder F - Dateien
        
        gbFTPautomatic = bmerke
        
    End If
    
    picprogress.Visible = False
    If gbBargeldEingabe = True Then
        schreibeProtoAbschluss "Kassenabschluss wurde durchgeführt Kasse " & gcKasNum & "--------"
    End If
    schreibeProtokoll "Ende Kassenabschluss"
    schreibeProtokoll "************************************"
    
    Kill gcDBPfad & "\ABSCHLUS.TXT"
    iFileNr = FreeFile
    
    Open gcDBPfad & "\ABSCHLUS.TXT" For Binary As #iFileNr
    Put #iFileNr, 1, gsNeuerAbschluß
    Close iFileNr
    
    LoescheTagesAbschlussMODUL7 = True
    
    Speicher_Tagesabschluss_GDPdU cKasse
    

Exit Function
LOKAL_ERROR:

    If err.Number = 75 Or err.Number = 3010 Or err.Number = 53 Or err.Number = 3376 Then
        schreibeProtokoll "Fehlerstufe 0: " & err.Number & " " & err.Description
        Resume Next
    Else
        schreibeProtokoll "Fehlerstufe 1: " & err.Number & " " & err.Description
        
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul7"
        Fehler.gsFunktion = "LoescheTagesAbschlussMODUL7"
        Fehler.gsFehlertext = "Im Programmteil Tagesabschluss bei Schritt: " & lbl3.Caption & " ist ein Fehler aufgetreten."
        
        Fehlermeldung1

        Resume Next
    End If
End Function
Private Sub Speicher_Tagesabschluss_GDPdU(cKasse As String)
    On Error GoTo LOKAL_ERROR

    Dim cSQL            As String
    Dim cPfad           As String
    Dim GDPdU_DB        As Database
    
    Dim cDatumJetzt     As String
    Dim cUhrzeitJetzt   As String
    Dim cDatumAlter     As String
    Dim cUhrzeitAlter   As String
    Dim cDatumNeuer     As String
    Dim cUhrzeitNeuer   As String
    Dim iALTEANR        As Integer
    Dim iNEUEANR        As Integer
    Dim cAlterZbon      As String
    Dim cNeuerZBon      As String
    Dim i               As Integer
    Dim rsrs            As Recordset
    
    If NewTableSuchenDBKombi("TAGKOPF_" & srechnertab, gdBase) = False Then
        Exit Sub
    End If
    
    cPfad = gcDBPfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "GDPdU\GDPdU.MDB"
    
    Set GDPdU_DB = OpenDatabase(cPfad, True, False, "MS Access;PWD=" & gsGDPdU_Passwort)

    If NewTableSuchenDBKombi("TAGKOPF_TEMP", GDPdU_DB) = False Then
        CreateTableT2 "TAGKOPF_TEMP", GDPdU_DB
        
        cSQL = "Create Index SCHLUESSEL on TAGKOPF_TEMP (SCHLUESSEL)"
        GDPdU_DB.Execute cSQL, dbFailOnError
    End If
    
    cPfad = gcDBPfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cDatumJetzt = Format$(Fix(Now), "DD.MM.YYYY")
    cUhrzeitJetzt = Format$(Now, "HH:MM:SS")
    
    cAlterZbon = ""
    cNeuerZBon = ""
    
    cSQL = "Select * from TAGKOPF_" & srechnertab
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!ALTERABSCH) Then
            cAlterZbon = rsrs!ALTERABSCH
        End If
        
        If Not IsNull(rsrs!NEUERABSCH) Then
            cNeuerZBon = rsrs!NEUERABSCH
        End If
    End If
     
    rsrs.Close: Set rsrs = Nothing
    
    iALTEANR = Val(Right(cAlterZbon, 6))
    iNEUEANR = Val(Right(cNeuerZBon, 6))
    
    If Left(cAlterZbon, 6) = "vorher" Then
        cDatumAlter = Mid(cAlterZbon, 22, 10)
        cDatumNeuer = Mid(cNeuerZBon, 22, 10)
        
        cUhrzeitAlter = Mid(cAlterZbon, 33, 8)
        cUhrzeitNeuer = Mid(cNeuerZBon, 33, 8)
    Else
        cDatumAlter = Mid(cAlterZbon, 1, 10)
        cDatumNeuer = Mid(cNeuerZBon, 1, 10)
        
        cUhrzeitAlter = Mid(cAlterZbon, 12, 8)
        cUhrzeitNeuer = Mid(cNeuerZBon, 12, 8)
    End If
    
    cSQL = "Insert into TAGKOPF_TEMP Select  "
    cSQL = cSQL & " '" & cDatumJetzt & "' as DruckDatum  "
    cSQL = cSQL & ",'" & cUhrzeitJetzt & "' as DruckZeit  "
    cSQL = cSQL & ", " & cKasse & " as KASNUM  "
    cSQL = cSQL & ",  " & iALTEANR & " as ALTEANR  "
    cSQL = cSQL & ",  " & iNEUEANR & " as NEUEANR  "
    If cDatumAlter <> "00.00.0000" Then
        cSQL = cSQL & ", '" & cDatumAlter & "' as AlterADatum  "
        cSQL = cSQL & ", '" & cUhrzeitAlter & "' as AlterAZeit  "
    End If
    cSQL = cSQL & ", '" & cDatumNeuer & "' as NeuerADatum  "
    cSQL = cSQL & ", '" & cUhrzeitNeuer & "' as NeuerAZeit  "
    
    cSQL = cSQL & ", SCHLUESSEL "
    cSQL = cSQL & ", WAE_CODE "
    For i = 1 To 54
        cSQL = cSQL & " , DATEN" & i
    Next i
    
    cSQL = cSQL & ", DATEN "
    cSQL = cSQL & ", ALTERABSCH "
    cSQL = cSQL & ", NEUERABSCH "
    
    cSQL = cSQL & " from [;DATABASE=" & cPfad & "KISSDATA.MDB;pwd=" & gsPasswort & "].TAGKOPF_" & srechnertab
    GDPdU_DB.Execute cSQL, dbFailOnError
    
    GDPdU_DB.Close
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "Speicher_Tagesabschluss_GDPdU"
    Fehler.gsFehlertext = "Im Programmteil Tagesabschluss ist ein Fehler aufgetreten."

    Fehlermeldung1
'    Resume Next
End Sub
Public Sub Speicher_Bestände_GDPdU()
    On Error GoTo LOKAL_ERROR

    Dim sSQL            As String
    Dim cPfad           As String
    Dim GDPdU_DB        As Database
    
    cPfad = gcDBPfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    cPfad = cPfad & "GDPdU\GDPdU.MDB"
    
    Set GDPdU_DB = OpenDatabase(cPfad, True, False, "MS Access;PWD=" & gsGDPdU_Passwort)

    If NewTableSuchenDBKombi("GLAGER_GDPdU", GDPdU_DB) = False Then
        CreateTableT2 "GLAGER_GDPDU", GDPdU_DB
    End If
    
    CheckIndex "GLAGER_GDPdU", "ARTNR", "", GDPdU_DB
    
    cPfad = gcDBPfad
    If Right$(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    sSQL = "Insert into GLAGER_GDPdU Select "
    sSQL = sSQL & " artikelnummer as artnr "
    sSQL = sSQL & " ,KVKPR1"
    sSQL = sSQL & " ,ekpr"
    sSQL = sSQL & " ,BEZEICH"
    sSQL = sSQL & " ,linr"
    sSQL = sSQL & ", bestand,Datevalue(now) as datum from [;DATABASE=" & cPfad & "KISSDATA.MDB;pwd=" & gsPasswort & "].LAGERD  "
    GDPdU_DB.Execute sSQL, dbFailOnError
    
    GDPdU_DB.Close
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "Speicher_Bestände_GDPdU"
    Fehler.gsFehlertext = "Im Programmteil Tagesabschluss ist ein Fehler aufgetreten."

    Fehlermeldung1
'    Resume Next
End Sub
Private Function gibt_es_etwas_zum_löschen() As Long
On Error GoTo LOKAL_ERROR


    gibt_es_etwas_zum_löschen = 0
    Dim cSQL As String
    Dim rsrs As Recordset
    
    cSQL = "Select count(*) as anz from Artikel where SYNSTATUS = 'D'"
    Set rsrs = gdBase.OpenRecordset(cSQL)
    If Not rsrs.EOF Then
        If Not IsNull(rsrs!anz) Then
            gibt_es_etwas_zum_löschen = rsrs!anz
        End If
    End If
     
    rsrs.Close: Set rsrs = Nothing


Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "gibt_es_etwas_zum_löschen"
    Fehler.gsFehlertext = "Im Programmteil Tagesabschluss ist ein Fehler aufgetreten. "
        
    Fehlermeldung1
End Function


Private Sub insertKasswaag(ckasn As String)
On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    Dim rsArt       As Recordset
    Dim rsKJ        As Recordset
    Dim iFeld       As Integer
    Dim cGewicht    As String
    Dim cPreisPer   As String
    Dim lPos        As Long
    Dim lPosG       As Long
    
    sSQL = "Select * from KASSWAAG where ARTNR = -1"
    Set rsKJ = gdBase.OpenRecordset(sSQL)
    
    sSQL = "Select * from afcbuch where Kasnum = " & ckasn
    sSQL = sSQL & " and Left(abezeich,3) = 'TEE' "
    
    Set rsrs = gdBase.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
    
            rsKJ.AddNew
            iFeld = 26
            rsKJ!artnr = rsrs!aartnr
            iFeld = 27
            rsKJ!BEZEICH = rsrs!ABEZEICH
            
            cGewicht = ""
            lPos = InStr(1, rsrs!ABEZEICH, "g")
            lPosG = lPos + 1
            If lPos <> 0 Then
                cGewicht = Mid(rsrs!ABEZEICH, 4, lPos - 5)
                cGewicht = Trim(cGewicht)
                
                If IsNumeric(cGewicht) Then
                    rsKJ!Gewicht = cGewicht
                End If
            End If
            
            cPreisPer = ""
            lPos = InStr(lPosG, rsrs!ABEZEICH, ",")
            If lPos <> 0 Then
                cPreisPer = Mid(rsrs!ABEZEICH, lPosG, lPos - lPosG + 3)
                cPreisPer = Trim(cPreisPer)
                
                If IsNumeric(cPreisPer) Then
                    rsKJ!PreisPer = cPreisPer
                End If
            End If
            
            iFeld = 28
            rsKJ!Menge = rsrs!aMenge
            iFeld = 29
            rsKJ!Preis = rsrs!APREIS
            iFeld = 30
            rsKJ!ADATE = rsrs!ADATE
            iFeld = 31
            rsKJ!AZEIT = rsrs!AZEIT
            iFeld = 32
'            rsKJ!BEDIENER = rsrs!abednu
            iFeld = 33
            rsKJ!Kundnr = rsrs!AKUNUM
            iFeld = 34
            rsKJ!FILIALE = rsrs!FILIALNR
            iFeld = 35
            rsKJ!kasnum = Val(gcKasNum)
            iFeld = 36
            rsKJ!linr = rsrs!linr
            iFeld = 40
            rsKJ!MWST = rsrs!AMWSK
            iFeld = 41
            rsKJ!ekpr = rsrs!ALEKPR
            iFeld = 42
            rsKJ!UMS_OK = rsrs!UMS_OK

            rsKJ!vkpr = rsrs!AVKPR
            iFeld = 43
            rsKJ!BELEGNR = rsrs!BELEGNR
            
            sSQL = "Select * from ARTIKEL where ARTNR = " & rsrs!aartnr
            Set rsArt = gdBase.OpenRecordset(sSQL)
            If Not rsArt.EOF Then
            rsArt.MoveFirst
            
                If Not IsNull(rsArt!LPZ) Then
                    rsKJ!LPZ = rsArt!LPZ
                End If
                If Not IsNull(rsArt!AGN) Then
                    rsKJ!AGN = rsArt!AGN
                End If
                If Not IsNull(rsArt!EAN) Then
                    rsKJ!EAN = rsArt!EAN
                End If
                
            End If
            rsArt.Close: Set rsArt = Nothing
            rsKJ.Update
        
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    rsKJ.Close


Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "insertKasswaag"
    Fehler.gsFehlertext = "Im Programmteil Tagesabschluss ist ein Fehler aufgetreten. " & iFeld
        
    Fehlermeldung1
End Sub
Private Sub SchreibeMonatsArtikelBestände()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    Dim ctemp As String
    Dim lMoni As Byte
    Dim lJahri As Integer
    Dim td As TableDef
    Dim fld As Field
        
    lJahri = Year(DateValue(Now))
    lMoni = Month(DateValue(Now))
    
    sSQL = "Insert Into BESTAEND Select ARTNR , " & lMoni & " as monat, " & lJahri & " as Jahr, Bestand "
    sSQL = sSQL & " from Artikel where Bestand <> 0 and not Bestand is null "
    gdBase.Execute sSQL, dbFailOnError
    
    Exit Sub
LOKAL_ERROR:
    If err.Number = 3464 Then
        ctemp = "Sie haben wahrscheinlich einen zu hohen Bestand in Ihrer Artikeldatenbank. " & vbCrLf & vbCrLf
        ctemp = ctemp & "Unter Service/Datenbank.../Datenbank bereinigen die Schaltfläche 'Artikel bereinigen' anklicken!"
        MsgBox ctemp, vbInformation, "Winkiss Hinweis:"
    Else
    
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul7"
        Fehler.gsFunktion = "SchreibeMonatsArtikelBestände"
        Fehler.gsFehlertext = "Im Programmteil Tagesabschluss ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    End If
End Sub
Private Function PrüfdateforBestand() As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim lMoni As Byte
    Dim lJahri As Integer
    Dim sSQL As String
    Dim rs As Recordset
    
    PrüfdateforBestand = False
    lMoni = Month(Date)
    lJahri = Year(Date)
    
    sSQL = "select artnr from bestaend where Jahr =  " & lJahri
    sSQL = sSQL & " and Monat = " & lMoni
    Set rs = gdBase.OpenRecordset(sSQL)
    
    If rs.RecordCount = 0 Then
        PrüfdateforBestand = True
    Else
        PrüfdateforBestand = False
    End If
    
    rs.Close: Set rs = Nothing
    
    Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "PrüfdateforBestand"
    Fehler.gsFehlertext = "Im Programmteil Tagesabschluss ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Sub ABSCHIEBENDB()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim slokalPfad  As String
   
    Do While IsAktionZulaessig(srechnertab & "Kdat") = False
    
    Loop
    
    frmWKL00.txtStatus.Text = "0"
    frmWKL00.picprogress.Visible = True
    frmWKL00.Label2.Visible = True
    
    If NewTableSuchenDBKombi("ZZZ", gdBase) = False Then
        frmWKL00.txtStatus.Text = "0"
        frmWKL00.picprogress.Visible = False
        frmWKL00.Label2.Visible = False
        Exit Sub
    End If
    
    frmWKL00.txtStatus.Text = "10"
    
    slokalPfad = "C:\aleer\kissdata.mdb"
    
    If FileExists(slokalPfad) Then
    
        frmWKL00!Label2.Caption = srechnertab & ": Kassendateien werden abgegeben..."
        frmWKL00!Label2.Refresh
        
        sichernLdb
        
        'Kassjour = KJ
        
        frmWKL00.txtStatus.Text = "16"
        frmWKL00!Label2.Caption = srechnertab & ": Datenbank: Kassjour kopieren..."
        frmWKL00!Label2.Refresh
        
        If NewTableSuchenDBKombi(srechnertab & "KJ", gdBase) Then
            sSQL = "Insert into " & srechnertab & "KJ Select * from Kassjour in '" & slokalPfad & "' "
            gdBase.Execute sSQL, dbFailOnError
        Else
            sSQL = "select * into " & srechnertab & "KJ from Kassjour in '" & slokalPfad & "' "
            gdBase.Execute sSQL, dbFailOnError
        End If
        
        'Afcbuch = AFCB
        
        frmWKL00.txtStatus.Text = "18"
        frmWKL00!Label2.Caption = srechnertab & ": Datenbank: AFCBUCH kopieren..."
        frmWKL00!Label2.Refresh
        
        If NewTableSuchenDBKombi(srechnertab & "AFCB", gdBase) Then
            sSQL = "Insert into " & srechnertab & "AFCB Select * from AFCBUCH in '" & slokalPfad & "' "
            gdBase.Execute sSQL, dbFailOnError
        Else
            sSQL = "select * into " & srechnertab & "AFCB from AFCBUCH in '" & slokalPfad & "' "
            gdBase.Execute sSQL, dbFailOnError
        End If
    
        
        
        
        
        'Kredit = KRED
        
        frmWKL00.txtStatus.Text = "20"
        frmWKL00!Label2.Caption = srechnertab & ": Datenbank: Kredit kopieren..."
        frmWKL00!Label2.Refresh
        If NewTableSuchenDBKombi(srechnertab & "KRED", gdBase) Then
            sSQL = "Insert into " & srechnertab & "KRED Select * from Kredit in '" & slokalPfad & "' "
            gdBase.Execute sSQL, dbFailOnError
        Else
            sSQL = "select * into " & srechnertab & "KRED from Kredit in '" & slokalPfad & "' "
            gdBase.Execute sSQL, dbFailOnError
        End If
        
        
        'KOLLVERK = KOLL
        
        frmWKL00.txtStatus.Text = "22"
        frmWKL00!Label2.Caption = srechnertab & ": Datenbank: KOLLVERK kopieren..."
        frmWKL00!Label2.Refresh
        If NewTableSuchenDBKombi(srechnertab & "KOLL", gdBase) Then
            sSQL = "Insert into " & srechnertab & "KOLL Select * from KOLLVERK in '" & slokalPfad & "' "
            gdBase.Execute sSQL, dbFailOnError
        Else
            sSQL = "select * into " & srechnertab & "KOLL from KOLLVERK in '" & slokalPfad & "' "
            gdBase.Execute sSQL, dbFailOnError
        End If
        
        'Retoure = RET
        
        frmWKL00.txtStatus.Text = "24"
        frmWKL00!Label2.Caption = srechnertab & ": Datenbank: Retoure kopieren..."
        frmWKL00!Label2.Refresh
        If NewTableSuchenDBKombi(srechnertab & "RET", gdBase) Then
            sSQL = "Insert into " & srechnertab & "RET Select * from Retoure in '" & slokalPfad & "' "
            gdBase.Execute sSQL, dbFailOnError
        Else
            sSQL = "select * into " & srechnertab & "RET from Retoure in '" & slokalPfad & "' "
            gdBase.Execute sSQL, dbFailOnError
        End If
        
        'Kassbon = KB
        
        frmWKL00.txtStatus.Text = "26"
        frmWKL00!Label2.Caption = srechnertab & ": Datenbank: Kassbon kopieren..."
        frmWKL00!Label2.Refresh
        If NewTableSuchenDBKombi(srechnertab & "KB", gdBase) Then
            sSQL = "Insert into " & srechnertab & "KB Select * from Kassbon in '" & slokalPfad & "' "
            gdBase.Execute sSQL, dbFailOnError
        Else
            sSQL = "select * into " & srechnertab & "KB from Kassbon in '" & slokalPfad & "' "
            gdBase.Execute sSQL, dbFailOnError
        End If
        
        
        'Afcstat = STAT
       
        frmWKL00.txtStatus.Text = "28"
        frmWKL00!Label2.Caption = srechnertab & ": Datenbank: AFCSTAT kopieren..."
        frmWKL00!Label2.Refresh
        
        If NewTableSuchenDBKombi(srechnertab & "STAT", gdBase) Then
        
            If Not SpalteInTabellegefundenNEW(srechnertab & "STAT", "GUTSCHGUTSCH", gdBase) Then
                SpalteAnfuegenNEW srechnertab & "STAT", "GUTSCHGUTSCH", "double", gdBase
            
'                sSQL = "Update AFCSTATP set " & srechnertab & "STAT = 0 "
'                schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
            
            End If
            
            If Not SpalteInTabellegefundenNEW(srechnertab & "STAT", "ABSCHOPF", gdBase) Then
                SpalteAnfuegenNEW srechnertab & "STAT", "ABSCHOPF", "double", gdBase
            
'                sSQL = "Update AFCSTATP set " & srechnertab & "STAT = 0 "
'                schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
            
            End If
            
            If Not SpalteInTabellegefundenNEW(srechnertab & "STAT", "WECHSEL", gdBase) Then
                SpalteAnfuegenNEW srechnertab & "STAT", "WECHSEL", "double", gdBase
            
'                sSQL = "Update AFCSTATP set " & srechnertab & "STAT = 0 "
'                schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
'
            End If
            
            If Not SpalteInTabellegefundenNEW(srechnertab & "STAT", "KDIFF", gdBase) Then
                SpalteAnfuegenNEW srechnertab & "STAT", "KDIFF", "double", gdBase
            
'                sSQL = "Update AFCSTATP set " & srechnertab & "STAT = 0 "
'                schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
            
            End If
            
            If Not SpalteInTabellegefundenNEW(srechnertab & "STAT", "TDIFF", gdBase) Then
                SpalteAnfuegenNEW srechnertab & "STAT", "TDIFF", "double", gdBase
            
'                sSQL = "Update AFCSTATP set " & srechnertab & "STAT = 0 "
'                schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
            
            End If
            
            If Not SpalteInTabellegefundenNEW(srechnertab & "STAT", "DUKA", gdBase) Then
                SpalteAnfuegenNEW srechnertab & "STAT", "DUKA", "double", gdBase
            
'                sSQL = "Update AFCSTATP set " & srechnertab & "STAT = 0 "
'                schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
            
            End If
            
            If Not SpalteInTabellegefundenNEW(srechnertab & "STAT", "NUMSKARTE", gdBase) Then
                SpalteAnfuegenNEW srechnertab & "STAT", "NUMSKARTE", "double", gdBase
            
'                sSQL = "Update AFCSTATP set " & srechnertab & "STAT = 0 "
'                schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
            
            End If
        
        
            sSQL = "Insert into " & srechnertab & "STAT Select * from AFCSTAT in '" & slokalPfad & "' "
            schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
        Else
            sSQL = "select * into " & srechnertab & "STAT from AFCSTAT in '" & slokalPfad & "' "
            schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
        End If
        
        AFCSTATPLUS srechnertab & "STAT", "AFCSTAT", gcKasNum
        
        sSQL = "DELETE from " & srechnertab & "STAT "
        schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
        
    End If
        
    frmWKL00.txtStatus.Text = "0"
    frmWKL00.picprogress.Visible = False
    frmWKL00.Label2.Visible = False
    
    AktionAustragen srechnertab & "Kdat"
    
    Exit Sub
LOKAL_ERROR:

        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul7"
        Fehler.gsFunktion = "ABSCHIEBENDB"
        Fehler.gsFehlertext = "Beim Synchronisieren der Datenbanken ist ein Fehler aufgetreten."
        
        Fehlermeldung1
        Resume Next

End Sub
Public Function sichernLdb() As Boolean
    On Error GoTo LOKAL_ERROR
    
    sichernLdb = False
    
    Dim LDB         As Database
    Dim lWert       As Long
    Dim cdatei      As String
    Dim ctmp        As String
    Dim sSQL        As String
    Dim sPfad       As String
    Dim slokalPfad  As String
    Dim lokalDB     As Database
    Dim i           As Integer
    
    Dim sTabellen(0 To 5) As String
    
    sTabellen(0) = "KASSBON"
    sTabellen(1) = "AFCSTAT"
    sTabellen(2) = "KOLLVERK"
    sTabellen(3) = "KREDIT"
    sTabellen(4) = "AFCBUCH"
    sTabellen(5) = "KASSJOUR"
    

    'Prüfen ob Verzeichnis c:\aleer existiert
    VerzVorhanden "aLeerSic", "C:\"
    
    lWert = DateValue(Now)
    ctmp = Format$(lWert, "DD.MM")
    
    cdatei = "A" & ctmp & Format$(TimeValue(Now), "HH:MM:SS")
    cdatei = SwapStr(cdatei, ".", "")
    cdatei = SwapStr(cdatei, ":", "")
    
    
    Set LDB = CreateDatabase("C:\aLeerSic\" & cdatei & ".MDB", dbLangGeneral, dbVersion40)
    slokalPfad = "C:\aleer\kissdata.mdb"
    
    If FileExists(slokalPfad) Then
        Set lokalDB = OpenDatabase(slokalPfad, False)
        
        sPfad = gcDBPfad 'Datenbankpfad
        If Right(sPfad, 1) <> "\" Then
            sPfad = sPfad & "\"
        End If
        
        For i = 0 To 5
            TransferTab lokalDB, "C:\aLeerSic\" & cdatei & ".MDB", sTabellen(i)
        Next i
        
        lokalDB.Close
        Set lokalDB = Nothing
    End If
    
    LDB.Close
    
    
    Exit Function
LOKAL_ERROR:
   
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "sichernLdb"
    Fehler.gsFehlertext = "Beim Synchronisieren der Datenbanken ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
End Function
Public Sub Ums_liefNew(Label2 As Label)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim cPfad       As String
    Dim lcount      As Long
    Dim rsrs        As Recordset
    
    loeschNEW "UMS_LIEF", gdBase
    
    Label2.Caption = "Erzeuge Verkaufstabelle UMS_LIEF neu. Bitte warten..."
    Label2.Refresh
    
    sSQL = " SELECT [KASSJOUR].[LINR], Year([ADATE]) AS Jahr, Month([ADATE]) AS Monat, Sum([KASSJOUR].[PREIS]) AS Umsatz, Sum([KASSJOUR].[MENGE]) AS Anzahl INTO UMS_LIEF"
    sSQL = sSQL & " From KASSJOUR GROUP BY [KASSJOUR].[LINR], Year([ADATE]), Month([ADATE]), [KASSJOUR].[UMS_OK]"
    sSQL = sSQL & " Having (((KASSJOUR.UMS_OK) = 'J'))"
    gdBase.Execute sSQL, dbFailOnError
    
    Label2.Caption = "ReIndiziere Verkaufstabelle : LINR, JAHR, MONAT"
    Label2.Refresh
    
    sSQL = "Create Index PRIMKEY on UMS_LIEF (LINR, JAHR, MONAT)"
    gdBase.Execute sSQL, dbFailOnError
    
    Label2.Caption = "ReIndiziere Verkaufstabelle : JAHR, MONAT"
    Label2.Refresh
    
    sSQL = "Create Index DATUM on UMS_LIEF (JAHR, MONAT)"
    gdBase.Execute sSQL, dbFailOnError
    
    Label2.Caption = "ReIndiziere Verkaufstabelle : LINR"
    Label2.Refresh
    
    sSQL = "Create Index LINR on UMS_LIEF (LINR)"
    gdBase.Execute sSQL, dbFailOnError
    
    Label2.Caption = "Anwender aktiv"
    Label2.Refresh
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "Ums_liefNew"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub UmskdjNew(Label2 As Label)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim cPfad       As String
    Dim lcount      As Long
    Dim rsrs        As Recordset
    
    loeschNEW "UMSKDJ", gdBase
    
    Label2.Caption = "Erzeuge Verkaufstabelle UMSKDJ neu. Bitte warten..."
    Label2.Refresh
    
    sSQL = " SELECT [KASSJOUR].[KUNDNR], Year([ADATE]) AS Jahr, Sum([KASSJOUR].[PREIS]) AS Umsatzj, Sum([KASSJOUR].[MENGE]) AS Anzahlj INTO UMSKDJ"
    sSQL = sSQL & " From KASSJOUR GROUP BY [KASSJOUR].[KUNDNR], Year([ADATE]),  [KASSJOUR].[UMS_OK]"
    sSQL = sSQL & " Having (((KASSJOUR.UMS_OK) = 'J'))"
    gdBase.Execute sSQL, dbFailOnError
    
    Label2.Caption = "ReIndiziere Verkaufstabellen: UMSKDJ"
    Label2.Refresh
    
    Label2.Caption = "ReIndiziere Verkaufstabelle: KUNDNR, JAHR"
    Label2.Refresh
    
    sSQL = "Create Index PRIMKEY on UMSKDJ (KUNDNR, JAHR)"
    gdBase.Execute sSQL, dbFailOnError
    
    Label2.Caption = "ReIndiziere Verkaufstabelle: JAHR"
    Label2.Refresh
    
    sSQL = "Create Index JAHR on UMSKDJ (JAHR)"
    gdBase.Execute sSQL, dbFailOnError
    
    Label2.Caption = "ReIndiziere Verkaufstabelle: KUNDNR"
    Label2.Refresh
    
    sSQL = "Create Index KUNDNR on UMSKDJ (KUNDNR)"
    gdBase.Execute sSQL, dbFailOnError
    
    Label2.Caption = "Anwender aktiv"
    Label2.Refresh
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "UmskdjNew"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
End Sub
Public Sub Ums_artNew(Label2 As Label)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim cPfad       As String
    Dim lcount      As Long
    Dim rsrs        As Recordset
    
    loeschNEW "UMS_ART", gdBase
    
    Label2.Caption = "Erzeuge Verkaufstabelle UMS_ART neu. Bitte warten..."
    Label2.Refresh
    
    sSQL = " SELECT [KASSJOUR].[ARTNR], Year([ADATE]) AS Jahr, Month([ADATE]) AS Monat, Sum([KASSJOUR].[PREIS]) AS Umsatz, Sum([KASSJOUR].[MENGE]) AS Anzahl INTO UMS_ART"
    sSQL = sSQL & " From KASSJOUR GROUP BY [KASSJOUR].[ARTNR], Year([ADATE]), Month([ADATE]), [KASSJOUR].[UMS_OK]"
    sSQL = sSQL & " Having (((KASSJOUR.UMS_OK) = 'J'))"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    

    Label2.Caption = "ReIndiziere Verkaufstabelle: ARTNR, JAHR, MONAT"
    Label2.Refresh
    
    sSQL = "Create Index PRIMKEY on UMS_ART (ARTNR, JAHR, MONAT)"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    Label2.Caption = "ReIndiziere Verkaufstabelle: JAHR, MONAT"
    Label2.Refresh
    
    sSQL = "Create Index DATUM on UMS_ART (JAHR, MONAT)"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    Label2.Caption = "ReIndiziere Verkaufstabelle: ARTNR"
    Label2.Refresh
    
    sSQL = "Create Index ARTNR on UMS_ART (ARTNR)"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    Label2.Caption = "Anwender aktiv"
    Label2.Refresh
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "Ums_artNew"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
   
End Sub
Public Sub dupliEANS(Label2 As Label)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim cPfad       As String
    Dim lcount      As Long
    Dim rsrs        As Recordset
    
    Screen.MousePointer = 11
    loeschNEW "DUPLIEAN", gdBase
    
    Label2.Caption = "Doppelte EAN's aus der Artikeldatenbank werden ermittelt..."
    Label2.Refresh
    
    sSQL = " SELECT artnr , ean into DUPLIEAN from artikel where ean is not null "
    gdBase.Execute sSQL, dbFailOnError
    
    Label2.Caption = "Schritt 1"
    Label2.Refresh

    sSQL = "Insert into DUPLIEAN SELECT artnr, ean2 as ean from artikel where ean2 is not null "
    gdBase.Execute sSQL, dbFailOnError
    
    Label2.Caption = "Schritt 2"
    Label2.Refresh

    sSQL = "Insert into DUPLIEAN SELECT artnr, ean3 as ean from artikel where ean3 is not null "
    gdBase.Execute sSQL, dbFailOnError
    
    If NewTableSuchenDBKombi("ARTEAN_K", gdBase) Then
    
        sSQL = "Insert into DUPLIEAN SELECT artnr, ean from ARTEAN_K where ean is not null "
        gdBase.Execute sSQL, dbFailOnError
    End If

    Label2.Caption = "Schritt 3"
    Label2.Refresh
    
    loeschNEW "LfEAN", gdBase
    
    sSQL = "Create Table LfEAN "
    sSQL = sSQL & " ( "
    sSQL = sSQL & " ARTNR single "
    sSQL = sSQL & ", EAN Text(13) "
    sSQL = sSQL & ", lf long"
    sSQL = sSQL & " ) "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into LfEAN Select * from DUPLIEAN"
    gdBase.Execute sSQL, dbFailOnError
    
    Label2.Caption = "Schritt 4"
    Label2.Refresh
    
    lcount = 0
    
    Set rsrs = gdBase.OpenRecordset("LfEAN", dbOpenTable)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            lcount = lcount + 1
            rsrs.Edit
            rsrs!lf = lcount
            rsrs.Update
            rsrs.MoveNext
        Loop
        rsrs.Close: Set rsrs = Nothing
    End If
    
    Label2.Caption = "Schritt 5"
    Label2.Refresh
    
    loeschNEW "eanlite", gdBase
    
    sSQL = "SELECT ARTNR, ean, Min(lf) AS Minlf INTO eanlite"
    sSQL = sSQL & " From lfean GROUP BY  ARTNR , ean"
    gdBase.Execute sSQL, dbFailOnError
    
    Label2.Caption = "Schritt 6"
    Label2.Refresh
    
    loeschNEW "DUPLIEAN", gdBase
    
    sSQL = "Create Table DUPLIEAN ( "
    sSQL = sSQL & " ARTNR LONG "
    sSQL = sSQL & ", EAN TEXT(13)) "
    gdBase.Execute sSQL, dbFailOnError
    
    Label2.Caption = "Schritt 7"
    Label2.Refresh
    
    sSQL = " Insert into DUPLIEAN SELECT ARTNR, EAN "
    sSQL = sSQL & " From eanlite "
    gdBase.Execute sSQL, dbFailOnError
    
    Label2.Caption = "Schritt 8"
    Label2.Refresh
    
    loeschNEW "LfEAN", gdBase
    
    sSQL = "Create Table LfEAN "
    sSQL = sSQL & " ( "
    sSQL = sSQL & " ARTNR single "
    sSQL = sSQL & ", EAN Text(13) "
    sSQL = sSQL & ", lf long"
    sSQL = sSQL & " ) "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into LfEAN Select * from DUPLIEAN"
    gdBase.Execute sSQL, dbFailOnError
    
    Label2.Caption = "Schritt 9"
    Label2.Refresh
    
    lcount = 0
    
    Set rsrs = gdBase.OpenRecordset("LfEAN", dbOpenTable)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            lcount = lcount + 1
            rsrs.Edit
            rsrs!lf = lcount
            rsrs.Update
            rsrs.MoveNext
        Loop
        rsrs.Close: Set rsrs = Nothing
    End If
    
    Label2.Caption = "Schritt 10"
    Label2.Refresh
    
    loeschNEW "eanlite", gdBase
    
    sSQL = "SELECT ARTNR, ean, Min(lf) AS Minlf INTO eanlite"
    sSQL = sSQL & " From lfean GROUP BY  ean ,ARTNR"
    gdBase.Execute sSQL, dbFailOnError
    
    Label2.Caption = "Schritt 11"
    Label2.Refresh
    
    loeschNEW "mehrfEAN", gdBase
    
    sSQL = "select ean into mehrfEAN from eanlite  group by ean having count(*) > 1"
    gdBase.Execute sSQL, dbFailOnError
    
    Label2.Caption = "Schritt 12"
    Label2.Refresh
    
    sSQL = "delete from mehrfEAN where val(ean)< 10000000 "
    gdBase.Execute sSQL, dbFailOnError
    
    Label2.Caption = "Schritt 13"
    Label2.Refresh
    
    loeschNEW "EANKoml", gdBase
    CreateTable "EANKOML", gdBase
    
    sSQL = "Insert into eankoml "
    sSQL = sSQL & " Select "
    sSQL = sSQL & " artikel.artnr as ArtNr "
    sSQL = sSQL & ",artikel.bezeich as Artikelbezeichnung"
    sSQL = sSQL & ",artikel.Bestand "
    sSQL = sSQL & ",artikel.EAN as EAN "
    sSQL = sSQL & ",artikel.LINR "
    sSQL = sSQL & ",artikel.KVKPR1 "
    sSQL = sSQL & ", '1' as farbe "
    sSQL = sSQL & ", '' as LINBEZ "
    sSQL = sSQL & " from ARTIKEL inner join mehrfean on ARTIKEL.EAN = mehrfean.EAN  "
    gdBase.Execute sSQL, dbFailOnError
    
    Label2.Caption = "Schritt 14"
    Label2.Refresh
    
    sSQL = "Insert into eankoml "
    sSQL = sSQL & " Select "
    sSQL = sSQL & " artikel.artnr as ArtNr "
    sSQL = sSQL & ",artikel.bezeich as Artikelbezeichnung"
    sSQL = sSQL & ",artikel.Bestand "
    sSQL = sSQL & ",artikel.EAN2 as EAN "
    sSQL = sSQL & ",artikel.LINR "
    sSQL = sSQL & ",artikel.KVKPR1 "
    sSQL = sSQL & ", '1' as farbe "
    sSQL = sSQL & ", '' as LINBEZ "
    sSQL = sSQL & " from ARTIKEL inner join mehrfean on ARTIKEL.EAN2 = mehrfean.EAN  "
    gdBase.Execute sSQL, dbFailOnError
    
    Label2.Caption = "Schritt 15"
    Label2.Refresh
    
    sSQL = "Insert into eankoml "
    sSQL = sSQL & " Select "
    sSQL = sSQL & " artikel.artnr as ArtNr "
    sSQL = sSQL & ",artikel.bezeich as Artikelbezeichnung"
    sSQL = sSQL & ",artikel.Bestand "
    sSQL = sSQL & ",artikel.EAN3 as EAN "
    sSQL = sSQL & ",artikel.LINR "
    sSQL = sSQL & ",artikel.KVKPR1 "
    sSQL = sSQL & ", '1' as farbe "
    sSQL = sSQL & ", '' as LINBEZ "
    sSQL = sSQL & " from ARTIKEL inner join mehrfean on ARTIKEL.EAN3 = mehrfean.EAN  "
    gdBase.Execute sSQL, dbFailOnError
    
    
    If NewTableSuchenDBKombi("ARTEAN_K", gdBase) Then
    
        sSQL = "Insert into eankoml "
        sSQL = sSQL & " Select "
        sSQL = sSQL & " artean_k.artnr as ArtNr "
        sSQL = sSQL & ",'' as Artikelbezeichnung"
        sSQL = sSQL & ",0 as Bestand "
        sSQL = sSQL & ",artean_k.EAN as EAN "
        sSQL = sSQL & ",0 as LINR "
        sSQL = sSQL & ",0 as KVKPR1 "
        sSQL = sSQL & ", '1' as farbe "
        sSQL = sSQL & ", '' as LINBEZ "
        sSQL = sSQL & " from artean_k inner join mehrfean on artean_k.EAN = mehrfean.EAN  "
        gdBase.Execute sSQL, dbFailOnError
        
        sSQL = " Update eankoml inner join ARTIKEL ON "
        sSQL = sSQL & " eankoml.ARTNR = ARTIKEL.ARTNR "
        sSQL = sSQL & " Set eankoml.Bestand = ARTIKEL.Bestand"
        sSQL = sSQL & " , eankoml.Artikelbezeichnung = ARTIKEL.bezeich"
        sSQL = sSQL & " , eankoml.LINR = ARTIKEL.LINR"
        sSQL = sSQL & " , eankoml.KVKPR1 = ARTIKEL.KVKPR1"
        sSQL = sSQL & " where eankoml.Artikelbezeichnung = ''"
        gdBase.Execute sSQL, dbFailOnError
        
    End If
    
    
    
    
    
    loeschNEW "ete", gdBase
    sSQL = "select * into ete from eankoml "
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "eankoml", gdBase
    sSQL = "select * into eankoml from ete order by ean"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update eankoml set bestand = 0 where bestand is null"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update eankoml inner join lisrt on lisrt.linr = eankoml.linr set LINBEZ = lisrt.liefbez"
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    'Spalte ran
    sSQL = " Alter table eankoml add LASTVK Text(10) "
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    Dim datLVK As String
    
    Set rsrs = gdBase.OpenRecordset("eankoml")

    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!artnr) Then
                rsrs.Edit
                datLVK = ErmlzVK(rsrs!artnr)
                rsrs!lastvk = datLVK
                rsrs.Update
            End If
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    
    
    Label2.Caption = "Schritt 16"
    Label2.Refresh
    
    Label2.Caption = "Anwender aktiv"
    Label2.Refresh
    
    Screen.MousePointer = 0
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "dupliEANS"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Public Sub dupliEANSstada(cEANDUPLI As String, sArt As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim cPfad       As String
    Dim lcount      As Long
    Dim rsrs        As Recordset
    
    Screen.MousePointer = 11
    
    loeschNEW "EANKoml1", gdBase
    
    sSQL = "Select "
    sSQL = sSQL & " artikel.artnr as ArtNr "
    sSQL = sSQL & ",artikel.bezeich as Artikelbezeichnung"
    sSQL = sSQL & ",artikel.Bestand "
    sSQL = sSQL & ",artikel.EAN "
    sSQL = sSQL & ",artikel.LINR "
    sSQL = sSQL & ",artikel.KVKPR1 "
    sSQL = sSQL & ", '1' as farbe "
    sSQL = sSQL & ", '' as LINBEZ "
    sSQL = sSQL & " into EANKoml1 from ARTIKEL Where  ARTIKEL.EAN = '" & cEANDUPLI & "'"
    sSQL = sSQL & " and ARTNR <> " & sArt
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into eankoml1 "
    sSQL = sSQL & " Select "
    sSQL = sSQL & " artikel.artnr as ArtNr "
    sSQL = sSQL & ",artikel.bezeich as Artikelbezeichnung"
    sSQL = sSQL & ",artikel.Bestand "
    sSQL = sSQL & ",artikel.EAN2 as EAN "
    sSQL = sSQL & ",artikel.LINR "
    sSQL = sSQL & ",artikel.KVKPR1 "
    sSQL = sSQL & ", '1' as farbe "
    sSQL = sSQL & ", '' as LINBEZ "
    sSQL = sSQL & " from ARTIKEL Where  ARTIKEL.EAN2 = '" & cEANDUPLI & "'"
    sSQL = sSQL & " and ARTNR <> " & sArt
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into eankoml1 "
    sSQL = sSQL & " Select "
    sSQL = sSQL & " artikel.artnr as ArtNr "
    sSQL = sSQL & ",artikel.bezeich as Artikelbezeichnung"
    sSQL = sSQL & ",artikel.Bestand "
    sSQL = sSQL & ",artikel.EAN3 as EAN "
    sSQL = sSQL & ",artikel.LINR "
    sSQL = sSQL & ",artikel.KVKPR1 "
    sSQL = sSQL & ", '1' as farbe "
    sSQL = sSQL & ", '' as LINBEZ "
    sSQL = sSQL & " from ARTIKEL Where  ARTIKEL.EAN3 = '" & cEANDUPLI & "'"
    sSQL = sSQL & " and ARTNR <> " & sArt
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into eankoml1 "
    sSQL = sSQL & " Select "
    sSQL = sSQL & " artnr  "
    sSQL = sSQL & ",bezeich as Artikelbezeichnung"
    sSQL = sSQL & ", 0 as Bestand "
    sSQL = sSQL & ",EAN "
    sSQL = sSQL & ",LINR "
    sSQL = sSQL & ",KVKPR1 "
    sSQL = sSQL & ", '9' as farbe "
'    sSQL = sSQL & ", '2' as farbe "
    sSQL = sSQL & ", '' as LINBEZ "
    sSQL = sSQL & " from MASTEMP Where ARTNR = " & sArt
    gdBase.Execute sSQL, dbFailOnError
    
'    Dim sCheckArtnr As String
'
'    Dim rsRS As DAO.Recordset
'    sSQL = "select artnr from eankoml1 where Farbe = '2'"
'    Set rsRS = gdBase.OpenRecordset(sSQL)
'    If Not rsRS.EOF Then
'        If Not IsNull(rsRS!artnr) Then
'            sCheckArtnr = rsRS!artnr
'        End If
'    End If
'    rsRS.Close: Set rsRS = Nothing

    
    sSQL = "Update eankoml1 set Farbe = '2' where artnr in (Select artnr from Artikel Where ARTNR = " & sArt & ")"
    sSQL = sSQL & " and farbe = '9'  "
    gdBase.Execute sSQL, dbFailOnError
    
'    sSQL = "Update eankoml1 set Farbe = '9' where artnr not in (Select artnr from Artikel Where ARTNR = " & sArt & ")"
'    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update eankoml1 inner join Artikel on eankoml1.artnr = artikel.artnr "
    sSQL = sSQL & " set eankoml1.bestand = artikel.bestand where artikel.ARTNR <> " & sArt
    gdBase.Execute sSQL, dbFailOnError
    
    
    loeschNEW "ete", gdBase
    sSQL = "select * into ete from eankoml1 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into eankoml select * from ete order by ean"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update eankoml set bestand = 0 where bestand is null"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update eankoml inner join lisrt on lisrt.linr = eankoml.linr set LINBEZ = lisrt.liefbez"
    gdBase.Execute sSQL, dbFailOnError
    
    Screen.MousePointer = 0
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "dupliEANSstada"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Public Sub dupliEANloesch(cEANDUPLI As String, sArt As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    
    sSQL = "Update Artikel set ean = '' where ean = '" & cEANDUPLI & "'"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Artikel set ean2 = '' where ean2 = '" & cEANDUPLI & "'"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Artikel set ean3 = '' where ean3 = '" & cEANDUPLI & "'"
    gdBase.Execute sSQL, dbFailOnError
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "dupliEANloesch"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub UmsartjNew(Label2 As Label)
    On Error GoTo LOKAL_ERROR

    Dim sSQL        As String
    Dim cPfad       As String
    Dim lcount      As Long
    Dim rsrs        As Recordset

    loeschNEW "UMSARTJ", gdBase

    Label2.Caption = "Erzeuge Verkaufstabelle UMSARTJ neu. Bitte warten..."
    Label2.Refresh

    sSQL = " SELECT [KASSJOUR].[ARTNR], Year([ADATE]) AS Jahr,  Sum([KASSJOUR].[PREIS]) AS Umsatzj, Sum([KASSJOUR].[MENGE]) AS Anzahlj INTO UMSARTJ"
    sSQL = sSQL & " From KASSJOUR GROUP BY [KASSJOUR].[ARTNR], Year([ADATE]), [KASSJOUR].[UMS_OK]"
    sSQL = sSQL & " Having (((KASSJOUR.UMS_OK) = 'J'))"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError

    Label2.Caption = "ReIndiziere Verkaufstabelle: ARTNR, JAHR"
    Label2.Refresh

    sSQL = "Create Index PRIMKEY on UMSARTJ (ARTNR, JAHR)"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError

    Label2.Caption = "ReIndiziere Verkaufstabelle: JAHR"
    Label2.Refresh

    sSQL = "Create Index DATUM on UMSARTJ (JAHR)"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError

    Label2.Caption = "ReIndiziere Verkaufstabelle: ARTNR"
    Label2.Refresh

    sSQL = "Create Index ARTNR on UMSARTJ (ARTNR)"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError

    Label2.Caption = "Anwender aktiv"
    Label2.Refresh

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "UmsartjNew"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub UmsliefjNew(Label2 As Label)
    On Error GoTo LOKAL_ERROR

    Dim sSQL        As String
    Dim cPfad       As String
    Dim lcount      As Long
    Dim rsrs        As Recordset

    loeschNEW "UMSLIEFJ", gdBase

    Label2.Caption = "Erzeuge Verkaufstabelle UMSLIEFJ neu. Bitte warten..."
    Label2.Refresh

    sSQL = " SELECT [KASSJOUR].[LINR], Year([ADATE]) AS Jahr,  Sum([KASSJOUR].[PREIS]) AS Umsatzj, Sum([KASSJOUR].[MENGE]) AS Anzahlj INTO UMSLIEFJ"
    sSQL = sSQL & " From KASSJOUR GROUP BY [KASSJOUR].[LINR], Year([ADATE]), [KASSJOUR].[UMS_OK]"
    sSQL = sSQL & " Having (((KASSJOUR.UMS_OK) = 'J'))"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError

    
    Label2.Caption = "ReIndiziere Verkaufstabelle: LINR, JAHR"
    Label2.Refresh

    sSQL = "Create Index PRIMKEY on UMSLIEFJ (LINR, JAHR)"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError

    Label2.Caption = "ReIndiziere Verkaufstabelle: JAHR"
    Label2.Refresh

    sSQL = "Create Index JAHR on UMSLIEFJ (JAHR)"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError

    Label2.Caption = "ReIndiziere Verkaufstabelle: LINR"
    Label2.Refresh

    sSQL = "Create Index LIEFNR on UMSLIEFJ (LINR)"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError

    Label2.Caption = "Anwender aktiv"
    Label2.Refresh

    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "UmsliefjNew"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Public Sub ArtliefReinigenkomplett(Label2 As Label, db As Database)
On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String
    
    Screen.MousePointer = 11
    
    Label2.Caption = "Jetzt werden alle Einträge in der Tabelle 'ARTLIEF' gelöscht."
    Label2.Refresh
    
    loeschNEW "artlief_T", db
    
    sSQL = "Select * into artlief_T from artlief "
    db.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from artlief "
    db.Execute sSQL, dbFailOnError
    
    Label2.Caption = "Jetzt werden alle Artikeleinträge der Tabelle 'ARTIKEL' in die Tabelle 'ARTLIEF' geschrieben."
    Label2.Refresh
    
    sSQL = "Insert into ARTLIEF Select "
    sSQL = sSQL & " ARTNR, LINR, LIBESNR, LEKPR, MINMEN from ARTIKEL "
    db.Execute sSQL, dbFailOnError
    
    sSQL = "Update artlief inner join artlief_T on artlief.artnr = artlief_T.artnr and artlief.linr = artlief_T.linr"
    sSQL = sSQL & " set artlief.lekpr = artlief_T.lekpr where artlief_T.lekpr > 0 "
    db.Execute sSQL, dbFailOnError
    
    sSQL = "Update artlief inner join artlief_T on artlief.artnr = artlief_T.artnr and artlief.linr = artlief_T.linr"
    sSQL = sSQL & " set artlief.LIBESNR = artlief_T.LIBESNR where artlief_T.LIBESNR <> '' "
    db.Execute sSQL, dbFailOnError
    
    loeschNEW "artlief_T", db
    
    Label2.Caption = "Anwender aktiv"
    Label2.Refresh
    
    Screen.MousePointer = 0
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "ArtliefReinigen"
    Fehler.gsFehlertext = "Beim Bereinigen der Tabelle Artlief ist ein Fehler aufgetreten."

    Fehlermeldung1
End Sub
Public Sub LisrtReinigenkomplett(Label2 As Label)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    Dim rsArt           As Recordset
    Dim rslinr          As Recordset
    Dim rskassj         As Recordset
    Dim sBez            As String
    Dim sNbez           As String
    Dim lLinr           As Long
    Dim lDatum          As Long
    Dim iRet            As Integer
    Screen.MousePointer = 11
    
    Label2.Caption = "Jetzt wird nach Sonderzeichen gesucht."
    Label2.Refresh
    
    
    sSQL = "update lisrt set kuerzel = Ucase(left(liefbez,5)) where( SYNSTATUS is null or SYNSTATUS = 'E' or SYNSTATUS = 'A' )"
    sSQL = sSQL & " and kuerzel is null "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "select * from lisrt where( SYNSTATUS is null or SYNSTATUS = 'E' or SYNSTATUS = 'A' )"
    
    Set rsArt = gdBase.OpenRecordset(sSQL)
    
    If Not rsArt.EOF Then
        rsArt.MoveLast
        Label2.Caption = "Jetzt wird die LISRT nach Sonderzeichen."
        Label2.Refresh
        
        rsArt.MoveFirst
        
        Do While Not rsArt.EOF
            
            If Not IsNull(rsArt!LIEFBEZ) Then
                sBez = rsArt!LIEFBEZ
            Else
                sBez = ""
            End If
            
            
            If InStr(sBez, "*") > 0 Then        'Sternchen
                sNbez = SwapStr(sBez, "*", " ")
                rsArt.Edit
                rsArt!LIEFBEZ = sNbez
                rsArt.Update
            End If
            
            If InStr(sBez, "'") > 0 Then        'Hochkommata
                sNbez = SwapStr(sBez, "'", " ")
                rsArt.Edit
                rsArt!LIEFBEZ = sNbez
                rsArt.Update
                
            End If
            
            If Not IsNull(rsArt!Kuerzel) Then
                sBez = rsArt!Kuerzel
            Else
                sBez = ""
            End If
            
            
            If InStr(sBez, "*") > 0 Then        'Sternchen
                sNbez = SwapStr(sBez, "*", " ")
                rsArt.Edit
                rsArt!Kuerzel = sNbez
                rsArt.Update
            End If
            
            If InStr(sBez, "'") > 0 Then        'Hochkommata
                sNbez = SwapStr(sBez, "'", " ")
                rsArt.Edit
                rsArt!Kuerzel = sNbez
                rsArt.Update
                
            End If
        
            rsArt.MoveNext
        Loop
        
    End If

    rsArt.Close: Set rsArt = Nothing
    
    loeschNEW "lisrtDel", gdBase
    
    sSQL = "Select * into lisrtdel from lisrt"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from lisrtdel "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from lisrt where linr is null "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from lisrt where linr = 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    lDatum = DateValue(Now) - 365
            
    Label2.Caption = "Jetzt werden Lieferanten gelöscht."
    Label2.Refresh
    
    sSQL = "select * from lisrt where( SYNSTATUS is null or SYNSTATUS = 'E' or SYNSTATUS = 'A' )"
    
    Set rsArt = gdBase.OpenRecordset(sSQL)
    If Not rsArt.EOF Then
        
        rsArt.MoveFirst
        
        Do While Not rsArt.EOF
            
            If Not IsNull(rsArt!linr) Then
                lLinr = rsArt!linr
            End If
            '1. Blick in die Artlief
            sSQL = "Select * from artlief where linr = " & lLinr
            Set rslinr = gdBase.OpenRecordset(sSQL)
            
            If rslinr.EOF Then
                'keine Artliefeintragungen
                'dann in kassjour blicken
                
                sSQL = "Select * from kassjour where linr = " & lLinr
                sSQL = sSQL & "and adate > " & Trim$(Str$(lDatum)) & " "
                Set rskassj = gdBase.OpenRecordset(sSQL)
                If rskassj.EOF Then
                    Label2.Caption = "Jetzt wird der Lieferant: " & lLinr & " " & rsArt!LIEFBEZ & " gelöscht."
                    Label2.Refresh
                    
                    sSQL = " Insert into lisrtdel select * from lisrt where linr = " & lLinr
                    gdBase.Execute sSQL, dbFailOnError
                    
                End If
                rskassj.Close
                
                
            
            End If
            rslinr.Close: Set rslinr = Nothing
            
            rsArt.MoveNext
        Loop
        
       
        
    End If

    rsArt.Close: Set rsArt = Nothing
    
    Label2.Caption = ""
    Label2.Refresh
    
    sSQL = " select * from lisrtdel "
    Set rsArt = gdBase.OpenRecordset(sSQL)
    If Not rsArt.EOF Then
        rsArt.MoveLast
    End If
    
    If Not rsArt.EOF Then
        'anzeige der zu löschenden Lieferanten
        If rsArt.RecordCount = 1 Then
            Label2.Caption = "Ein Lieferant wurde ermittelt. Löschvorschlag wird erstellt..."
        Else
            Label2.Caption = rsArt.RecordCount & " Lieferanten wurden ermittelt. Löschvorschlag wird erstellt..."
        End If
        rsArt.Close: Set rsArt = Nothing
        Label2.Refresh
        reportbildschirm "dwkl33c", "awkl33h"
        
        Pause (5)
        
        iRet = MsgBox("Möchten Sie diese Lieferanten wirklich löschen?", vbQuestion + vbYesNo, "Winkiss Hinweis:")
        If iRet = vbYes Then
        
            sSQL = "update LISRT inner join lisrtdel on lisrt.linr = lisrtdel.linr set lisrt.SYNSTATUS = 'D'  "
            gdBase.Execute sSQL, dbFailOnError
            
        Else
            Screen.MousePointer = 0
            Label2.Caption = "Löschvorgang abgebrochen!"
            Label2.Refresh
            Pause (2)
        End If
    End If
    
    Label2.Caption = "Anwender aktiv"
    Label2.Refresh

    Screen.MousePointer = 0
    
    Exit Sub
LOKAL_ERROR:
   
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul7"
        Fehler.gsFunktion = "LisrtReinigenkomplett"
        Fehler.gsFehlertext = "Beim Bereinigen der Tabelle Lisrt ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    
End Sub

Public Sub verwaisteArtikel()
On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    
    sSQL = "Delete from lisrt where linr = 999999 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into lisrt (linr,liefbez) values (999999,'KISS')"
'    sSQL = sSQL & " where linr not in (Select linr from lisrt where linr = 999999 ) "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Artikel set linr = 999999 where linr is null "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Artikel set linr = 999999 where linr = 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from artlief where linr = 999999"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into ARTLIEF Select "
    sSQL = sSQL & " ARTNR, LINR, LIBESNR, LEKPR, MINMEN from ARTIKEL "
    sSQL = sSQL & "where LINR = 999999"
    gdBase.Execute sSQL, dbFailOnError

Exit Sub
LOKAL_ERROR:
   
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "verwaisteArtikel"
    Fehler.gsFehlertext = "Beim Bereinigen der Tabelle Lisrt ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
''''Public Sub FuturaImport(lblx As Label, cpfad As String)
''''On Error GoTo LOKAL_ERROR
''''
''''    Dim sSQL            As String
''''    Dim dbFUT           As Database
''''    Dim dbQ             As Database
''''    Dim lAnzTable       As Long
''''    Dim lcount          As Long
''''    Dim sTabname        As String
''''    Dim lZaehler        As Long
''''
''''    Set dbQ = OpenDatabase(cpfad, False, False, "Paradox 5.x;")
''''
''''    Kill cpfad & "\FUTURA.MDB"
''''    Set dbFUT = CreateDatabase(cpfad & "\FUTURA.MDB", dbLangGeneral, dbVersion40)
''''
''''    dbQ.TableDefs.Refresh
''''    lAnzTable = dbQ.TableDefs.Count
''''    lZaehler = lAnzTable
''''    For lcount = 0 To lAnzTable - 1
''''        sTabname = dbQ.TableDefs(lcount).name
''''
''''        If Datendrin(sTabname, dbQ) = True Then
''''            anzeige "normal", "(" & lZaehler & ") " & sTabname & " wird importiert...", lblx
''''            sSQL = "Select * into " & sTabname & " from " & sTabname & " IN '" & cpfad & "' 'Paradox 5.x;'"
''''            dbFUT.Execute sSQL, dbFailOnError
''''        End If
''''
''''        lZaehler = lZaehler - 1
''''
''''    Next lcount
''''
''''''    Set dbFut = OpenDatabase(cPfad & "\FUTURA.MDB", False, False)
''''
''''
''''
''''    loeschNEW "ART_HIST", dbFUT
''''    loeschNEW "ACODEDEF", dbFUT
''''    loeschNEW "ACODENUM", dbFUT
''''    loeschNEW "ADRTEXT", dbFUT
''''    loeschNEW "AEDITDET", dbFUT
''''    loeschNEW "AEDITKPF", dbFUT
''''    loeschNEW "AIRPORT", dbFUT
''''    loeschNEW "AIRLINE", dbFUT
''''    loeschNEW "AIRFLIGH", dbFUT
''''    loeschNEW "AKT_HOUR", dbFUT
''''    loeschNEW "AKT_Kopf", dbFUT
''''    loeschNEW "ANG_ADR", dbFUT
''''    loeschNEW "ANGEZEIL", dbFUT
''''    loeschNEW "ANGHEAD", dbFUT
''''    loeschNEW "ANGHIST", dbFUT
''''    loeschNEW "ANGPERI", dbFUT
''''    loeschNEW "ANGZAHL", dbFUT
''''    loeschNEW "ANGZEIL", dbFUT
''''    loeschNEW "ANREDEN", dbFUT
''''    loeschNEW "ANSITRNS", dbFUT
''''    loeschNEW "ART_CODE", dbFUT
''''    loeschNEW "ART_LEVE", dbFUT
''''    loeschNEW "ART_LFID", dbFUT
''''    loeschNEW "ART_LINF", dbFUT
''''    loeschNEW "ART_LOTS", dbFUT
''''    loeschNEW "ART_STCK", dbFUT
''''    loeschNEW "ART_TEXT", dbFUT
''''    loeschNEW "ART_USER", dbFUT
''''    loeschNEW "ART_VERP", dbFUT
''''    loeschNEW "ASRT_FIL", dbFUT
''''    loeschNEW "ASRT_KPF", dbFUT
''''    loeschNEW "ASRT_ART", dbFUT
''''    loeschNEW "AUFHIST", dbFUT
''''    loeschNEW "AUFTEXT", dbFUT
''''    loeschNEW "AUTRSYS", dbFUT
''''    loeschNEW "AUSLAGER", dbFUT
''''    loeschNEW "AUTBAUM", dbFUT
''''    loeschNEW "AUTTRANS", dbFUT
''''    loeschNEW "AUTTEXT", dbFUT
''''    loeschNEW "AUTMODUL", dbFUT
''''    'B
''''    loeschNEW "BANKNAME", dbFUT
''''    loeschNEW "BASART", dbFUT
''''    loeschNEW "BASHEAD", dbFUT
''''    loeschNEW "BAT_HEAD", dbFUT
''''    loeschNEW "BAT_PROT", dbFUT
''''    loeschNEW "BATCHDET", dbFUT
''''    loeschNEW "BATCHGRP", dbFUT
''''    loeschNEW "BATCHJOB", dbFUT
''''    loeschNEW "BATCHLIN", dbFUT
''''    loeschNEW "BATCHMOD", dbFUT
''''    loeschNEW "BATCHQUE", dbFUT
''''    loeschNEW "BEDIHEAD", dbFUT
''''    loeschNEW "BEDING", dbFUT
''''    loeschNEW "BEDIZEIL", dbFUT
''''    loeschNEW "BENGMEMB", dbFUT
''''    loeschNEW "BENGROUP", dbFUT
''''    loeschNEW "BENUTZER", dbFUT
''''    loeschNEW "BEST_ADR", dbFUT
''''    loeschNEW "BESTCOD1", dbFUT
''''    loeschNEW "BESTCOD2", dbFUT
''''    loeschNEW "BESTHEAD", dbFUT
''''    loeschNEW "BESTHIST", dbFUT
''''    loeschNEW "BESTLINK", dbFUT
''''    loeschNEW "BESTLT", dbFUT
''''    loeschNEW "BESTPROZ", dbFUT
''''    loeschNEW "BESTRAHM", dbFUT
''''    loeschNEW "BESTVOR", dbFUT
''''    loeschNEW "BESTWUN", dbFUT
''''    loeschNEW "BESTZEIL", dbFUT
''''    loeschNEW "BONUS", dbFUT
''''    loeschNEW "BSTAINLT", dbFUT
''''    loeschNEW "BSTATXLT", dbFUT
''''    loeschNEW "BSTBLKLT", dbFUT
''''    loeschNEW "BSTBONLT", dbFUT
''''    loeschNEW "BSTCODLT", dbFUT
''''    loeschNEW "BSTHDRLT", dbFUT
''''    loeschNEW "BSTPOSLT", dbFUT
''''    loeschNEW "BSTPRSLT", dbFUT
''''    loeschNEW "BSTPRZLT", dbFUT
''''    loeschNEW "BSTTXTLT", dbFUT
''''    loeschNEW "BUDGETD", dbFUT
''''    loeschNEW "BUDGETND", dbFUT
''''    loeschNEW "BUDGETDF", dbFUT
''''    loeschNEW "BUDGETFK", dbFUT
''''    loeschNEW "BUDGETZD", dbFUT
''''    loeschNEW "BUDGETZL", dbFUT
''''    loeschNEW "BUDGORD", dbFUT
''''    'C
''''    loeschNEW "CASH", dbFUT
''''    loeschNEW "CASHUSER", dbFUT
''''    loeschNEW "CCD_TEST", dbFUT
''''    loeschNEW "CCD_INFO", dbFUT
''''    loeschNEW "CCD_KOPF", dbFUT
''''    loeschNEW "CONFEINT", dbFUT
''''    loeschNEW "CONFENUM", dbFUT
''''    loeschNEW "CONFGRUP", dbFUT
''''    loeschNEW "CONFHILF", dbFUT
''''    loeschNEW "CONFRECH", dbFUT
''''    loeschNEW "CONFWERT", dbFUT
''''    loeschNEW "CUPDHILF", dbFUT
''''    loeschNEW "CUPDGRUP", dbFUT
''''    loeschNEW "CUPDENUM", dbFUT
''''    loeschNEW "CUPDEINT", dbFUT
''''    'D
''''    loeschNEW "DANG_ADR", dbFUT
''''    loeschNEW "DANGEZEI", dbFUT
''''    loeschNEW "DANGHEAD", dbFUT
''''    loeschNEW "DANGHIST", dbFUT
''''    loeschNEW "DANGPERI", dbFUT
''''    loeschNEW "DANGZAHL", dbFUT
''''    loeschNEW "DANGZEIL", dbFUT
''''    loeschNEW "DELINFO", dbFUT
''''    loeschNEW "DEP_ART", dbFUT
''''    loeschNEW "DEP_KOPF", dbFUT
''''    loeschNEW "DEP_ZAHL", dbFUT
''''    loeschNEW "DEPOTART", dbFUT
''''    loeschNEW "DIS_BEDI", dbFUT
''''    loeschNEW "DIS_EXTR", dbFUT
''''    loeschNEW "DIS_KOPF", dbFUT
''''    loeschNEW "DIS_WGR", dbFUT
''''    loeschNEW "DIS_KUND", dbFUT
''''    loeschNEW "DIZ_ZUS", dbFUT
''''    loeschNEW "DLIF_ADR", dbFUT
''''    loeschNEW "DLIFHEAD", dbFUT
''''    loeschNEW "DLIFHIST", dbFUT
''''    loeschNEW "DLIFZAHL", dbFUT
''''    loeschNEW "DLIFZEIL", dbFUT
''''    loeschNEW "DOCCNTHD", dbFUT
''''    loeschNEW "DOCCNTRF", dbFUT
''''    loeschNEW "DOCMAPLN", dbFUT
''''    loeschNEW "DOCMAPHD", dbFUT
''''    loeschNEW "DR_ZUORD", dbFUT
''''    loeschNEW "DRECNUNG", dbFUT
''''    loeschNEW "DRUCKER", dbFUT
''''    loeschNEW "DWEXPORT", dbFUT
''''    loeschNEW "DRECZAHL", dbFUT
''''
''''    'E
''''    loeschNEW "EANVWKPF", dbFUT
''''    loeschNEW "EANVWLST", dbFUT
''''    loeschNEW "EC_SPERR", dbFUT
''''    loeschNEW "EDI_IDAT", dbFUT
''''    loeschNEW "EDI_IN", dbFUT
''''    loeschNEW "EDI_ODAT", dbFUT
''''    loeschNEW "EDI_OUT", dbFUT
''''    loeschNEW "EDISYS", dbFUT
''''    loeschNEW "EFTPROTO", dbFUT
''''    loeschNEW "EIGEN", dbFUT
''''    loeschNEW "EIGLIST", dbFUT
''''    loeschNEW "EINLIST", dbFUT
''''    loeschNEW "EKATABON", dbFUT
''''    loeschNEW "EKATACOD", dbFUT
''''    loeschNEW "EKATALOG", dbFUT
''''    loeschNEW "EKATARTI", dbFUT
''''    loeschNEW "EKATATXT", dbFUT
''''    loeschNEW "EKATAVAR", dbFUT
''''    loeschNEW "EKATVPRS", dbFUT
''''    loeschNEW "EM_DATEN", dbFUT
''''    loeschNEW "EM_KOPF", dbFUT
''''    loeschNEW "EPARTFIL", dbFUT
''''    loeschNEW "EPARTKEN", dbFUT
''''    loeschNEW "EPARTMAP", dbFUT
''''    loeschNEW "EPARTNER", dbFUT
''''    loeschNEW "EPMSGHST", dbFUT
''''    loeschNEW "EPRECST", dbFUT
''''    loeschNEW "EREPDATA", dbFUT
''''    loeschNEW "EREPKOPF", dbFUT
''''    loeschNEW "EREPPARA", dbFUT
''''    loeschNEW "EREPUSER", dbFUT
''''    loeschNEW "EX_FDATA", dbFUT
''''    loeschNEW "EX_FKOPF", dbFUT
''''    loeschNEW "EX_LADR", dbFUT
''''    loeschNEW "EX_LHEAD", dbFUT
''''    loeschNEW "EX_LHIST", dbFUT
''''    loeschNEW "EX_LZAHL", dbFUT
''''    loeschNEW "EX_LZEIL", dbFUT
''''    loeschNEW "EX_RECH", dbFUT
''''    loeschNEW "EX_RZAHL", dbFUT
''''    loeschNEW "EX_WE_LH", dbFUT
''''    loeschNEW "EX_WE_LN", dbFUT
''''    loeschNEW "EX_WE_LZ", dbFUT
''''    loeschNEW "EX_WE_RH", dbFUT
''''    loeschNEW "EX_WE_RN", dbFUT
''''    loeschNEW "EX_WELNK", dbFUT
''''    loeschNEW "EX_WLHST", dbFUT
''''    loeschNEW "EX_WRHST", dbFUT
''''    loeschNEW "EXA_CODE", dbFUT
''''    loeschNEW "EXA_EANS", dbFUT
''''    loeschNEW "EXA_KOPF", dbFUT
''''    loeschNEW "EXA_PALG", dbFUT
''''    loeschNEW "EXA_PHST", dbFUT
''''    loeschNEW "EXA_PRGR", dbFUT
''''    loeschNEW "EXARTIKEL", dbFUT
''''    loeschNEW "EXPO_HDR", dbFUT
''''    loeschNEW "EXPO_FLD", dbFUT
''''    loeschNEW "EXPO_DFL", dbFUT
''''    loeschNEW "EXTRAFEE", dbFUT
''''    loeschNEW "EXTWECHS", dbFUT
''''
''''    'F
''''    loeschNEW "FBUINFO", dbFUT
''''    loeschNEW "FBUKOPF", dbFUT
''''    loeschNEW "FCODEDEF", dbFUT
''''    loeschNEW "FCODENUM", dbFUT
''''    loeschNEW "FERKOPF", dbFUT
''''    loeschNEW "FIBERF", dbFUT
''''    loeschNEW "FIBKASDL", dbFUT
''''    loeschNEW "FIBKASOP", dbFUT
''''    loeschNEW "FIBKASTR", dbFUT
''''    loeschNEW "FIBUBUCH", dbFUT
''''    loeschNEW "FIBUBWA", dbFUT
''''    loeschNEW "FIBUBWAK", dbFUT
''''    loeschNEW "FIBUKTXT", dbFUT
''''    loeschNEW "FIBUOP", dbFUT
''''    loeschNEW "FIL_RAB", dbFUT
''''    loeschNEW "FILPRART", dbFUT
''''    loeschNEW "FILPRHDR", dbFUT
''''    loeschNEW "FILSYSTM", dbFUT
''''    loeschNEW "FS_ADDAT", dbFUT
''''    loeschNEW "FS_CTRL", dbFUT
''''    loeschNEW "FS_MOVE", dbFUT
''''    loeschNEW "FS_PURCH", dbFUT
''''    loeschNEW "FS_SALES", dbFUT
''''    loeschNEW "FTR_KOPF", dbFUT
''''    loeschNEW "FTR_DATA", dbFUT
''''
''''    'K
''''    loeschNEW "KASS_IMP", dbFUT
''''    'L
''''    loeschNEW "L_AUS", dbFUT
''''    loeschNEW "L_DET", dbFUT
''''    loeschNEW "L_GND", dbFUT
''''    loeschNEW "L_RES", dbFUT
''''    loeschNEW "L_SPR", dbFUT
''''    loeschNEW "L_SYS", dbFUT
''''    loeschNEW "L_TIT", dbFUT
''''    loeschNEW "LAGDELTA", dbFUT
''''    loeschNEW "LAGTRANS", dbFUT
''''    loeschNEW "Lastschr", dbFUT
''''    loeschNEW "LAUFSCHL", dbFUT
''''    loeschNEW "LFARBE", dbFUT
''''
''''    loeschNEW "LGORT", dbFUT
''''    loeschNEW "LGORTDEF", dbFUT
''''    loeschNEW "LGROESSE", dbFUT
''''    loeschNEW "LIEFHIST", dbFUT
''''    loeschNEW "LIEFZAHL", dbFUT
''''    loeschNEW "LIFFPRIS", dbFUT
''''    loeschNEW "LIFKPRIS", dbFUT
''''    loeschNEW "LIZENZ", dbFUT
''''    loeschNEW "LMAPPING", dbFUT
''''    loeschNEW "LOESCHDF", dbFUT
''''    loeschNEW "LOGONOFF", dbFUT
''''    loeschNEW "LSNAPDET", dbFUT
''''    loeschNEW "LSNAPHDR", dbFUT
''''    'M
''''    loeschNEW "MAIL", dbFUT
''''    loeschNEW "MAIL_REC", dbFUT
''''    loeschNEW "MAIL_SND", dbFUT
''''    loeschNEW "MANDANT", dbFUT
''''    loeschNEW "MANUCASH", dbFUT
''''    loeschNEW "MANUHEAD", dbFUT
''''    loeschNEW "METHODEN", dbFUT
''''    loeschNEW "MNGTYP", dbFUT
''''    'N
''''    loeschNEW "NETZUSER", dbFUT
''''    loeschNEW "NL_PARAM", dbFUT
''''    'O
''''    loeschNEW "ORGABT", dbFUT
''''    'P
''''    loeschNEW "PA_WERTE", dbFUT
''''    loeschNEW "PACKUNG", dbFUT
''''    loeschNEW "PB_ART", dbFUT
''''    loeschNEW "PB_SET", dbFUT
''''    loeschNEW "PB_TRNS", dbFUT
''''    loeschNEW "PILOTDTA", dbFUT
''''    loeschNEW "PILOTINF", dbFUT
''''    loeschNEW "PILOTKPF", dbFUT
''''    loeschNEW "PILOTTRE", dbFUT
''''    loeschNEW "PLZ_REF", dbFUT
''''    loeschNEW "PR_DTL", dbFUT
''''    loeschNEW "PRAEMDET", dbFUT
''''    loeschNEW "PREISGRU", dbFUT
''''    loeschNEW "PREISLAG", dbFUT
''''    loeschNEW "PREISPKT", dbFUT
''''    loeschNEW "PREISRND", dbFUT
''''    loeschNEW "PREISSTF", dbFUT
''''    loeschNEW "PRINTER", dbFUT
''''    loeschNEW "PRINTMAP", dbFUT
''''    loeschNEW "PRINTQUE", dbFUT
''''    loeschNEW "PRODDET", dbFUT
''''    loeschNEW "PRODGRP", dbFUT
''''    loeschNEW "PRODHEAD", dbFUT
''''    loeschNEW "PRODTEIL", dbFUT
''''    loeschNEW "PRODTOUR", dbFUT
''''    loeschNEW "PROMOART", dbFUT
''''    loeschNEW "PROMOTN", dbFUT
''''    loeschNEW "PVERFIL", dbFUT
''''    loeschNEW "PVERGRP", dbFUT
''''    loeschNEW "PVERZUS", dbFUT
''''
''''    'R
''''    loeschNEW "RECHNUNG", dbFUT
''''    loeschNEW "RECHZAHL", dbFUT
''''    loeschNEW "REGION", dbFUT
''''    loeschNEW "RETGRUND", dbFUT
''''    loeschNEW "RPT_SOG", dbFUT
''''    loeschNEW "RPT_SOK", dbFUT
''''    loeschNEW "RPT_VTG", dbFUT
''''    loeschNEW "RPT_VTP", dbFUT
''''
''''
''''
''''    'G
''''    loeschNEW "GWVORDET", dbFUT
''''    loeschNEW "GWVORSUM", dbFUT
''''    'I
''''    loeschNEW "INVLKGRD", dbFUT
''''
''''
''''    'S
''''    loeschNEW "SBLAYCOL", dbFUT
''''    loeschNEW "SBLAYHDR", dbFUT
''''    loeschNEW "SBLAYROW", dbFUT
''''    loeschNEW "SCANDATA", dbFUT
''''    loeschNEW "SCANKOPF", dbFUT
''''    loeschNEW "SERDATA", dbFUT
''''    loeschNEW "SERDELTA", dbFUT
''''    loeschNEW "SERTRANS", dbFUT
''''    loeschNEW "SOMI_DET", dbFUT
''''    loeschNEW "SOMI_FIL", dbFUT
''''    loeschNEW "SOMI_KPF", dbFUT
''''    loeschNEW "SOR_ARTI", dbFUT
''''    loeschNEW "SOR_KOPF", dbFUT
''''    loeschNEW "SPOOLING", dbFUT
''''    loeschNEW "SPRACHE", dbFUT
''''    loeschNEW "SSTATDAT", dbFUT
''''    loeschNEW "SSTATHDR", dbFUT
''''    loeschNEW "STATAFLG", dbFUT
''''    loeschNEW "STATART", dbFUT
''''    loeschNEW "STATART2", dbFUT
''''    loeschNEW "STATDDTA", dbFUT
''''    loeschNEW "STATDFLG", dbFUT
''''    loeschNEW "STATDIDX", dbFUT
''''    loeschNEW "STATKOMP", dbFUT
''''    loeschNEW "STATODKY", dbFUT
''''    loeschNEW "STATODTA", dbFUT
''''    loeschNEW "STATOIDX", dbFUT
''''    loeschNEW "STATOIKY", dbFUT
''''    loeschNEW "STATPDTA", dbFUT
''''    loeschNEW "STATPERI", dbFUT
''''    loeschNEW "STATPIDX", dbFUT
''''    loeschNEW "STATSORT", dbFUT
''''    loeschNEW "STATSYS", dbFUT
''''
''''    'K
''''
''''    loeschNEW "KARKOPF", dbFUT
''''    loeschNEW "KARZEIL", dbFUT
''''    loeschNEW "KASCOART", dbFUT
''''    loeschNEW "KASCOFIL", dbFUT
''''    loeschNEW "KASCOGRP", dbFUT
''''    loeschNEW "KASCOKPF", dbFUT
''''    loeschNEW "KASFISIA", dbFUT
''''    loeschNEW "KASFISID", dbFUT
''''    loeschNEW "KASFISIH", dbFUT
''''    loeschNEW "KASS_EAN", dbFUT
''''    loeschNEW "KAS_IMP", dbFUT
''''    loeschNEW "KASS_PLU", dbFUT
''''    loeschNEW "KASSDLTA", dbFUT
''''    loeschNEW "KASSDUPL", dbFUT
''''    loeschNEW "KASSE", dbFUT
''''    loeschNEW "KASSWAHL", dbFUT
''''    loeschNEW "KB_ADR", dbFUT
''''    loeschNEW "KB_DET", dbFUT
''''    loeschNEW "KB_KOPF", dbFUT
''''    loeschNEW "KLASSIF", dbFUT
''''    loeschNEW "KO_DATEN", dbFUT
''''    loeschNEW "KO_LINK", dbFUT
''''    loeschNEW "KO_SEITE", dbFUT
''''    loeschNEW "KO_KOPF", dbFUT
''''    loeschNEW "KONDIT", dbFUT
''''    loeschNEW "KONTODET", dbFUT
''''    loeschNEW "KONTOGRP", dbFUT
''''    loeschNEW "KONTOSUM", dbFUT
''''    loeschNEW "KOSTENST", dbFUT
''''    loeschNEW "KRDAUSGL", dbFUT
''''    loeschNEW "KRDDELTA", dbFUT
''''    loeschNEW "KRDHIST", dbFUT
''''    loeschNEW "KRDKARTE", dbFUT
''''    loeschNEW "KRDKAUF", dbFUT
''''    loeschNEW "KRDKONTO", dbFUT
''''    loeschNEW "KRDZAHL", dbFUT
''''    loeschNEW "KST_SPLI", dbFUT
''''    loeschNEW "KUN_EIG", dbFUT
''''
''''
''''    'T
''''
''''    loeschNEW "TBSTKOPF", dbFUT
''''    loeschNEW "TBSTZEIL", dbFUT
''''    loeschNEW "TITELREF", dbFUT
''''    loeschNEW "TOUR", dbFUT
''''    loeschNEW "TRANSLAT", dbFUT
''''    loeschNEW "TRENDLST", dbFUT
''''    loeschNEW "TRN_BILD", dbFUT
''''    loeschNEW "TRN_MINF", dbFUT
''''
''''
''''
''''
''''    'U
''''    loeschNEW "UINFDATA", dbFUT
''''    loeschNEW "UINFHEAD", dbFUT
''''    loeschNEW "UML_DATA", dbFUT
''''    loeschNEW "UML_KOPF", dbFUT
''''    loeschNEW "UMSDELTA", dbFUT
''''    loeschNEW "UMSTRANS", dbFUT
''''    loeschNEW "UPDBAUM", dbFUT
''''    loeschNEW "UPDMODUL", dbFUT
''''    loeschNEW "UPDTRANS", dbFUT
''''    loeschNEW "USEREXEC", dbFUT
''''
''''
''''
''''    'V
''''    loeschNEW "VERGLMGT", dbFUT
''''    loeschNEW "VER_KOST", dbFUT
''''    loeschNEW "VER_KOPF", dbFUT
''''    loeschNEW "VER_BER", dbFUT
''''    loeschNEW "VD_ZAHL", dbFUT
''''    loeschNEW "VD_LOCK", dbFUT
''''    loeschNEW "VD_KLAS", dbFUT
''''    loeschNEW "VD_HEAD", dbFUT
''''    loeschNEW "VD_FILI", dbFUT
''''    loeschNEW "VD_DEFI", dbFUT
''''    loeschNEW "VTG_ARTI", dbFUT
''''    loeschNEW "VTG_BEDI", dbFUT
''''    loeschNEW "VTG_DEBI", dbFUT
''''    loeschNEW "VK_BER", dbFUT
''''    loeschNEW "VTG_DIST", dbFUT
''''    loeschNEW "VTG_HERS", dbFUT
''''    loeschNEW "VTG_INFO", dbFUT
''''    loeschNEW "VTG_PTAB", dbFUT
''''    loeschNEW "VTG_PREI", dbFUT
''''    loeschNEW "VTG_KUND", dbFUT
''''    loeschNEW "VTG_KOPF", dbFUT
''''    'W
''''    loeschNEW "WAEHRUNG", dbFUT
''''    loeschNEW "WARKOPAR", dbFUT
''''    loeschNEW "WARKORB", dbFUT
''''    loeschNEW "WBL_KOPF", dbFUT
''''    loeschNEW "WDEFHEAD", dbFUT
''''    loeschNEW "WDEFSELE", dbFUT
''''    loeschNEW "WE_HEAD", dbFUT
''''    loeschNEW "WE_HEADR", dbFUT
''''    loeschNEW "WE_KOST", dbFUT
''''    loeschNEW "WE_LHIST", dbFUT
''''    loeschNEW "WE_LINK", dbFUT
''''    loeschNEW "WE_NEBEN", dbFUT
''''    loeschNEW "WE_RECH", dbFUT
''''    loeschNEW "WE_RGNEB", dbFUT
''''    loeschNEW "WE_RHIST", dbFUT
''''    loeschNEW "WE_ZEIL", dbFUT
''''    loeschNEW "WFACTDET", dbFUT
''''    loeschNEW "WFACTION", dbFUT
''''    loeschNEW "WFCOND", dbFUT
''''    loeschNEW "WFLGEVNT", dbFUT
''''    loeschNEW "WFLGHEAD", dbFUT
''''    loeschNEW "WFRULE", dbFUT
''''    loeschNEW "WRKBVKI", dbFUT
''''    loeschNEW "WRKDELTA", dbFUT
''''    loeschNEW "WRKLAGER", dbFUT
''''    'Z
''''    loeschNEW "ZUTEILPR", dbFUT
''''    loeschNEW "ZUTEILHD", dbFUT
''''    loeschNEW "ZEITZONE", dbFUT
''''    loeschNEW "ZEITTYP", dbFUT
''''
''''    loeschNEW "LAND", dbFUT
''''    loeschNEW "KASSKOPF", dbFUT
''''    loeschNEW "KAS_EX_K", dbFUT
''''    loeschNEW "KASSDKPF", dbFUT
''''    loeschNEW "KAS_TKPF", dbFUT
''''    loeschNEW "KAS_TEXT", dbFUT
''''    loeschNEW "AKT_FIL", dbFUT
''''    loeschNEW "AKT_ART", dbFUT
''''    loeschNEW "ARBKOPF", dbFUT
''''    loeschNEW "KERDEF", dbFUT
''''    loeschNEW "KERDISP", dbFUT
''''    loeschNEW "MISGRDEF", dbFUT
''''    loeschNEW "KASSKUND", dbFUT
''''    loeschNEW "LAGERKOR", dbFUT
''''    loeschNEW "KERENTW", dbFUT
''''
''''
''''
''''
''''
''''    dbFUT.Close
''''
''''
''''
''''
''''
''''
''''Exit Sub
''''LOKAL_ERROR:
''''    If err.Number = 53 Then
''''        Resume Next
''''    Else
''''        Fehler.gsDescr = err.Description
''''        Fehler.gsNumber = err.Number
''''        Fehler.gsFormular = "Modul7"
''''        Fehler.gsFunktion = "FuturaImport"
''''        Fehler.gsFehlertext = "Beim Bereinigen der Tabelle Lisrt ist ein Fehler aufgetreten."
''''
''''        Fehlermeldung1
''''    End If
''''
''''End Sub

Public Sub allePreischutzaufnein()
On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    
    sSQL = "Update Artikel set Preisschu = 'N'"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError

Exit Sub
LOKAL_ERROR:
   
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "allePreischutzaufnein"
    Fehler.gsFehlertext = "Beim Bereinigen der Tabelle Lisrt ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub Jubi1()
On Error GoTo LOKAL_ERROR

    Screen.MousePointer = 11
    
    Dim sSQL            As String
    
    sSQL = "Delete from Etidru"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into etidru select artnr, bezeich, bestand, 1 as anzahl,vkpr"
    sSQL = sSQL & " ,libesnr,ean,lpz,linr, 1 as filnr from artikel where kvkpr1 < vkpr"
    gdBase.Execute sSQL, dbFailOnError
    
    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "Jubi1"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub Jubi2()
On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11
    
    Dim sSQL            As String
    
    loeschNEW "artsic35", gdBase
    
    sSQL = "Select * into artsic35 from artikel"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update artikel set kvkpr1 = vkpr"
    gdBase.Execute sSQL, dbFailOnError
    
'    sSQL = "Update artikel set rabatt_ok = 'J' "
'    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update dbeinste set gesrab = 25 "
    gdBase.Execute sSQL, dbFailOnError
    
'    sSQL = "Update artikel set rabatt_ok = 'N' where artnr = 500010 "
'    gdBase.Execute sSQL, dbFailOnError
'
'    sSQL = "Update artikel set rabatt_ok = 'N' where artnr = 500020 "
'    gdBase.Execute sSQL, dbFailOnError
'
'    sSQL = "Update artikel set rabatt_ok = 'N' where artnr = 500018 "
'    gdBase.Execute sSQL, dbFailOnError
'
'    sSQL = "Update artikel set rabatt_ok = 'N' where artnr = 500754 "
'    gdBase.Execute sSQL, dbFailOnError
    Screen.MousePointer = 0
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "Jubi2"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub Jubi3()
On Error GoTo LOKAL_ERROR
    Screen.MousePointer = 11
    Dim sSQL            As String
    
    sSQL = "Update artikel inner join artsic35 on artikel.artnr = artsic35.artnr set artikel.kvkpr1 = artsic35.kvkpr1 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update dbeinste set gesrab = 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    Screen.MousePointer = 0

Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "Jubi3"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub alleLINRausDELLIEFlöschen()
On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    Dim cLinr           As String
    Dim rsrs            As Recordset
    
    sSQL = "Delete from Artlief where linr is null"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from Artlief where linr = 0"
    gdBase.Execute sSQL, dbFailOnError
    
    Set rsrs = gdBase.OpenRecordset("DELLIEF")
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            cLinr = "0"
            If Not IsNull(rsrs!linr) Then
                cLinr = rsrs!linr
                
                sSQL = "Delete from LISRT where linr = " & cLinr
                gdBase.Execute sSQL, dbFailOnError
                
                sSQL = "Delete from Artikel where linr = " & cLinr
                gdBase.Execute sSQL, dbFailOnError
                
                sSQL = "Delete from ARTLIEF where linr = " & cLinr
                gdBase.Execute sSQL, dbFailOnError
            
            End If
        rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    

Exit Sub
LOKAL_ERROR:
   
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "alleLINRausDELLIEFlöschen"
    Fehler.gsFehlertext = "Beim Bereinigen der Tabelle Lisrt ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub ArtikelReinigenkomplett(Label2 As Label, Label8 As Label, Label11 As Label, Label12 As Label, Label9 As Label, Label10 As Label, bgef As Boolean)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    Dim rsArt           As Recordset
    Dim RsBez           As Recordset
    Dim rsBESTAND       As Recordset
    Dim rsagn           As Recordset
    Dim rslinr          As Recordset
    Dim sLinr           As String
    Dim sArtnr          As String
    Dim sAGN            As String
    Dim sBez            As String
    Dim lAnz            As Long
    Dim lSonderAnz      As Long
    Dim lBezAnz         As Long
    Dim dBestand        As Double
    Dim dNBestand       As Byte
    Dim sNbez           As String
    Dim bBezaender      As Boolean
    Dim j               As Integer
    
    bBezaender = False
    
    Screen.MousePointer = 11
    
    sSQL = "Update artikel set agn = 0 where agn is null"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update artikel set bestand = 0 where bestand is null"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update artikel set linr = 0 where linr is null"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from  agndbf where agn = 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from  agndbf where agn is null "
    gdBase.Execute sSQL, dbFailOnError
    
    lSonderAnz = 0
    lBezAnz = 0
    
    Label9.Visible = True
    Label10.Visible = True
    Label11.Visible = True
    Label12.Visible = True
    
    loeschNEW "BBESTAND", gdBase
    CreateTable "BBestand", gdBase
    
    loeschNEW "SBEZEICH", gdBase
    CreateTable "SBEZEICH", gdBase
    
    loeschNEW "AGNERR", gdBase
    CreateTable "AGNERR", gdBase
    
    loeschNEW "LINRERR", gdBase
    CreateTableT2 "LINRERR", gdBase
    
    Label2.Caption = "Jetzt wird nach Sonderzeichen und nach sehr hohen positiven wie negativen Beständen in der Tabelle 'ARTIKEL' gesucht."
    Label2.Refresh
    Set RsBez = gdBase.OpenRecordset("SBEZEICH", dbOpenTable)
    Set rsBESTAND = gdBase.OpenRecordset("BBESTAND", dbOpenTable)
    Set rsagn = gdBase.OpenRecordset("AGNERR", dbOpenTable)
    Set rslinr = gdBase.OpenRecordset("LINRERR", dbOpenTable)
    
    sSQL = "Select * from Artikel "
    If bgef = True Then
        sSQL = sSQL & " where gefuehrt = 'J' "
    End If
    Set rsArt = gdBase.OpenRecordset(sSQL)
    
    If Not rsArt.EOF Then
        rsArt.MoveLast
        lAnz = rsArt.RecordCount
        Label2.Caption = "Jetzt werden die Artikel nach Sonderzeichen und nach sehr hohen positiven wie negativen Beständen durchsucht."
        Label2.Refresh
        
        Label12.Caption = " Artikel bereinigt (Sonderzeichen)"
        Label12.Refresh
        
        Label10.Caption = " Artikelbestände bereinigt."
        Label10.Refresh
        
        rsArt.MoveFirst
        
        Do While Not rsArt.EOF
            If Not IsNull(rsArt!artnr) Then
                sArtnr = rsArt!artnr
            Else
                sArtnr = "0"
            End If
            
            If Not IsNull(rsArt!BEZEICH) Then
                sBez = rsArt!BEZEICH
            Else
                sBez = ""
            End If
            
            If Not IsNull(rsArt!BESTAND) Then
                dBestand = rsArt!BESTAND
            Else
                dBestand = 0
            End If
            
            If Not IsNull(rsArt!AGN) Then
                sAGN = rsArt!AGN
            Else
                sAGN = "0"
            End If
            
            If Not isGueltigeAGN(sAGN) Then
                rsagn.AddNew
                rsagn!artnr = rsArt!artnr
                rsagn!BEZEICH = rsArt!BEZEICH
                rsagn!AGN = rsArt!AGN
                rsagn!KVKPR1 = rsArt!KVKPR1
                rsagn!LIBESNR = rsArt!LIBESNR
                rsagn!EAN = rsArt!EAN
                rsagn!LPZ = rsArt!LPZ
                rsagn!linr = rsArt!linr
                rsagn.Update
            End If
            
            If Not IsNull(rsArt!linr) Then
                sLinr = rsArt!linr
            Else
                sLinr = "0"
            End If
            
            If Not isGueltigelinr(sLinr) Then
                rslinr.AddNew
                rslinr!artnr = rsArt!artnr
                rslinr!BEZEICH = rsArt!BEZEICH
                rslinr!AGN = rsArt!AGN
                rslinr!KVKPR1 = rsArt!KVKPR1
                rslinr!LIBESNR = rsArt!LIBESNR
                rslinr!EAN = rsArt!EAN
                rslinr!LPZ = rsArt!LPZ
                rslinr!linr = rsArt!linr
                rslinr.Update
            End If
            
            
            'Bestand prüfen
            If dBestand > 1000 Or dBestand < -100 Then
                lBezAnz = lBezAnz + 1
                Label9.Caption = lBezAnz
                Label9.Refresh
                dNBestand = 0

                Bestandsveraenderung sArtnr, 0, "Datenbank Reinigung"
                
                rsBESTAND.AddNew
                rsBESTAND!artnr = rsArt!artnr
                rsBESTAND!BEZEICH = rsArt!BEZEICH
                rsBESTAND!BESTANDA = dBestand
                rsBESTAND!BESTANDN = dNBestand
                rsBESTAND!KVKPR1 = rsArt!KVKPR1
                rsBESTAND!LIBESNR = rsArt!LIBESNR
                rsBESTAND!EAN = rsArt!EAN
                rsBESTAND!LPZ = rsArt!LPZ
                rsBESTAND!linr = rsArt!linr
                
                rsBESTAND.Update
            End If
            
            
            bBezaender = False
            If InStr(sBez, "*") > 0 Then        'Sternchen
                lSonderAnz = lSonderAnz + 1
                Label11.Caption = lSonderAnz
                Label11.Refresh
                sNbez = SwapStr(sBez, "*", " ")
                rsArt.Edit
                rsArt!BEZEICH = sNbez
                rsArt.Update
                bBezaender = True
            End If
            
            If InStr(sBez, "'") > 0 Then        'Hochkommata
                lSonderAnz = lSonderAnz + 1
                Label11.Caption = lSonderAnz
                Label11.Refresh
                sNbez = SwapStr(sBez, "'", " ")
                rsArt.Edit
                rsArt!BEZEICH = sNbez
                rsArt.Update
                bBezaender = True
            End If
            
            
            sNbez = SwapStr(sBez, "„", " ")
            sNbez = SwapStr(sBez, "]", " ")
            sNbez = SwapStr(sBez, "[", " ")
            
            sNbez = SwapStr(sBez, "}", " ")
            sNbez = SwapStr(sBez, "{", " ")
            
            If InStr(sBez, Chr(34)) > 0 Then        '"
                lSonderAnz = lSonderAnz + 1
                Label11.Caption = lSonderAnz
                Label11.Refresh
                sNbez = SwapStr(sBez, Chr(34), " ")
                rsArt.Edit
                rsArt!BEZEICH = sNbez
                rsArt.Update
                bBezaender = True
            End If
            
            If bBezaender Then
                RsBez.AddNew
    
                RsBez!artnr = rsArt!artnr
                RsBez!ABEZEICH = sBez
                RsBez!NBEZEICH = sNbez
    
                RsBez!BESTAND = rsArt!BESTAND
                RsBez!KVKPR1 = rsArt!KVKPR1
                RsBez!LIBESNR = rsArt!LIBESNR
    
                RsBez!EAN = rsArt!EAN
                RsBez!LPZ = rsArt!LPZ
                RsBez!linr = rsArt!linr
    
                RsBez.Update
            End If
            
            lAnz = lAnz - 1
           
            j = lAnz Mod 100
            If j = 0 Then
                Label8.Caption = lAnz
                Label8.Refresh
            Else
                
            End If
                
            rsArt.MoveNext
        Loop
    End If
    
    rsagn.Close: Set rsagn = Nothing
    rslinr.Close: Set rslinr = Nothing
    RsBez.Close: Set RsBez = Nothing
    rsBESTAND.Close: Set rsBESTAND = Nothing
    rsArt.Close: Set rsArt = Nothing
    
    Label8.Caption = ""
    Label8.Refresh
    
    
    If lSonderAnz > 0 Or lBezAnz > 0 Then
        Label2.Caption = "Protokolle werden erstellt..."
        Label2.Refresh
        
        Pause (2)
        
        Label9.Caption = ""
        Label9.Refresh
        Label10.Caption = ""
        Label10.Refresh
        Label11.Caption = ""
        Label11.Refresh
        Label12.Caption = ""
        Label12.Refresh
    Else
        Set rsagn = gdBase.OpenRecordset("AGNERR", dbOpenTable)
        If Not rsagn.RecordCount = 0 Then
            Label2.Caption = "nicht zuordbare AGN's in der Artikeldatenbank"
            Label2.Refresh
        End If
        rsagn.Close: Set rsagn = Nothing
    
        Label2.Caption = "Artikeldatenbank ist fehlerfrei"
        Label2.Refresh
        
        Label9.Caption = ""
        Label9.Refresh
        Label10.Caption = ""
        Label10.Refresh
        Label11.Caption = ""
        Label11.Refresh
        Label12.Caption = ""
        Label12.Refresh
        Pause (2)
        
    End If
    
    Set rsBESTAND = gdBase.OpenRecordset("BBestand", dbOpenTable)
    If Not rsBESTAND.RecordCount = 0 Then
        reportbildschirm "dwkl33c", "awkl33c"
    End If
    rsBESTAND.Close: Set rsBESTAND = Nothing
    
    Set RsBez = gdBase.OpenRecordset("SBEZEICH", dbOpenTable)
    If Not RsBez.RecordCount = 0 Then
        reportbildschirm "dwkl33d", "awkl33d"
    End If
    RsBez.Close: Set RsBez = Nothing
    
    Set rsagn = gdBase.OpenRecordset("AGNERR", dbOpenTable)
    If Not rsagn.RecordCount = 0 Then
        reportbildschirm "dwkl33g", "awkl33g"
    End If
    rsagn.Close: Set rsagn = Nothing
    
    Set rsagn = gdBase.OpenRecordset("LINRERR", dbOpenTable)
    If Not rsagn.RecordCount = 0 Then
        reportbildschirm "dwkl33g", "awkl33i"
    End If
    rsagn.Close: Set rsagn = Nothing
    
    
    Label9.Visible = False
    Label10.Visible = False
    Label11.Visible = False
    Label12.Visible = False
    
    Label2.Caption = "Anwender aktiv"
    Label2.Refresh

    Screen.MousePointer = 0
    
    Exit Sub
LOKAL_ERROR:
   
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "ArtikelReinigen"
    Fehler.gsFehlertext = "Beim Bereinigen der Tabelle 'Artikel' ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    Resume Next
End Sub
Public Sub KVKPR1runden(Label2 As Label, Label8 As Label, Label11 As Label, Label12 As Label, Label9 As Label, Label10 As Label)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    Dim rsArt           As Recordset
    Dim sArtnr          As String
    Dim lAnz            As Long
    Dim lSonderAnz      As Long
    Dim dKVKneu         As Double
    Dim dKVKalt         As Double
    Dim dKVKVergleich   As Double
    Dim j               As Integer
    
    Screen.MousePointer = 11
    
    lSonderAnz = 0
    
    Label9.Visible = True
    Label10.Visible = True
    Label11.Visible = True
    Label12.Visible = True
    
    
    Label2.Caption = "Jetzt werden die Kassenverkaufspreise nach den hinterlegten Rundungsregeln gerundet."
    Label2.Refresh
    
    Set rsArt = gdBase.OpenRecordset("ARTIKEL", dbOpenTable)
    
    If Not rsArt.EOF Then
        rsArt.MoveLast
        lAnz = rsArt.RecordCount
        
        Label12.Caption = " Artikelpreise gerundet "
        Label12.Refresh
        
        rsArt.MoveFirst
        
        Do While Not rsArt.EOF
            If Not IsNull(rsArt!KVKPR1) Then
                dKVKalt = rsArt!KVKPR1
                dKVKVergleich = rsArt!KVKPR1
            Else
                dKVKalt = 0
                dKVKVergleich = 0
            End If
            
            dKVKneu = CDbl(Runden(dKVKalt))
            
            If dKVKneu <> dKVKVergleich Then         'Sternchen
                lSonderAnz = lSonderAnz + 1
                Label11.Caption = lSonderAnz
                Label11.Refresh
                
                rsArt.Edit
                rsArt!KVKPR1 = dKVKneu
                rsArt.Update
            End If
            
            lAnz = lAnz - 1
           
            j = lAnz Mod 100
            If j = 0 Then
                Label8.Caption = lAnz
                Label8.Refresh
            Else
                
            End If
                
            rsArt.MoveNext
        Loop
    End If
    rsArt.Close: Set rsArt = Nothing
    
    Label8.Caption = ""
    Label8.Refresh
    
    Label9.Visible = False
    Label10.Visible = False
    Label11.Visible = False
    Label12.Visible = False
    
    MsgBox lSonderAnz & " Artikelpreise wurden laut Rundungsregeln gerundet.", vbInformation, "Winkiss Hinweis:"
    
    Label2.Caption = "Anwender aktiv"
    Label2.Refresh

    Screen.MousePointer = 0
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "KVKPR1runden"
    Fehler.gsFehlertext = "Beim Bereinigen der Tabelle 'Artikel' ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub KVKPR1abw(Label2 As Label, Label8 As Label, Label11 As Label, Label12 As Label, Label9 As Label, Label10 As Label)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    Dim rsArt           As Recordset
    Dim sArtnr          As String
    Dim lAnz            As Long
    Dim lSonderAnz      As Long
    Dim dKVKneu         As Double
    Dim dKVKalt         As Double
    Dim dKVKVergleich   As Double
    Dim j               As Integer
    Dim dLVKPR          As Double
    Dim dABW            As Double
    Dim lBestand        As Long
    Screen.MousePointer = 11
    
    lSonderAnz = 0
    
'    Label9.Visible = True
'    Label10.Visible = True
'    Label11.Visible = True
'    Label12.Visible = True
    
    loeschNEW "KVKABW", gdBase
    CreateTable "KVKABW", gdBase
    
    
    Label2.Caption = "Jetzt werden die Kassenverkaufspreise überprüft.(-20% vom LVK und runden)"
    Label2.Refresh
    
    Set rsArt = gdBase.OpenRecordset("ARTIKEL", dbOpenTable)
    
    If Not rsArt.EOF Then
        rsArt.MoveLast
        lAnz = rsArt.RecordCount
        
        Label12.Caption = " Artikelpreise gerundet "
        Label12.Refresh
        
        rsArt.MoveFirst
        
        Do While Not rsArt.EOF
        
            If Not IsNull(rsArt!vkpr) Then
                dLVKPR = rsArt!vkpr
            Else
                dLVKPR = 0
            End If
            
            If Not IsNull(rsArt!BESTAND) Then
                lBestand = rsArt!BESTAND
            Else
                lBestand = 0
            End If
        
        
            If Not IsNull(rsArt!KVKPR1) Then
                dKVKalt = rsArt!KVKPR1
                dKVKVergleich = rsArt!KVKPR1
            Else
                dKVKalt = 0
                dKVKVergleich = 0
            End If
            
            dKVKneu = CDbl(Runden(dLVKPR * 0.8))
            
            If dKVKneu <> dKVKVergleich Then 'Sternchen
'                If lbestand > 0 Then
'                    lSonderAnz = lSonderAnz + 1
'                End If
                
                dABW = 0
                dABW = Format(dKVKneu - dKVKVergleich, "######0.00")
                
'                Label11.Caption = lSonderAnz
'                Label11.Refresh
                
                sSQL = "Insert into KVKABW select "
                sSQL = sSQL & " ARTNR "
                sSQL = sSQL & ", LINR "
                sSQL = sSQL & ", BEZEICH "
                sSQL = sSQL & ", KVKPR1 as KVKALT "
                sSQL = sSQL & ", '" & dKVKneu & "' as KVKNEU "
                sSQL = sSQL & ", VKPR "
                sSQL = sSQL & ", Bestand "
                
                sSQL = sSQL & ", val(awm) as FARBNR  "
                sSQL = sSQL & ", '" & dABW & "' as ABW "
                sSQL = sSQL & " from artikel "
                sSQL = sSQL & " where artnr = " & rsArt!artnr
                gdBase.Execute sSQL, dbFailOnError
                
                
'                rsArt.Edit
'                rsArt!KVKPR1 = dKVKneu
'                rsArt.Update
            End If
            
            lAnz = lAnz - 1
           
            j = lAnz Mod 100
            If j = 0 Then
                Label8.Caption = lAnz
                Label8.Refresh
            Else
                
            End If
                
            rsArt.MoveNext
        Loop
    End If
    rsArt.Close: Set rsArt = Nothing
    
    sSQL = "Update KVKABW inner join lisrt on KVKABW.linr = lisrt.linr"
    sSQL = sSQL & " Set KVKABW.liefbez = lisrt.liefbez "
    gdBase.Execute sSQL, dbFailOnError
    
    BringFarbeInsSpiel "KVKABW", gdBase
    
    Label8.Caption = ""
    Label8.Refresh
    
    Label9.Visible = False
    Label10.Visible = False
    Label11.Visible = False
    Label12.Visible = False
    
    reportbildschirm "dwkl33c", "awkl33l"

    Label2.Caption = "Anwender aktiv"
    Label2.Refresh

    Screen.MousePointer = 0
    
    Exit Sub
LOKAL_ERROR:
   
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "KVKPR1abw"
    Fehler.gsFehlertext = "Beim Bereinigen der Tabelle 'Artikel' ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub KVKPR1abwKorri(Label2 As Label, Label8 As Label, Label11 As Label, Label12 As Label, Label9 As Label, Label10 As Label)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    Dim rsArt           As Recordset
    Dim sArtnr          As String
    Dim lAnz            As Long
    Dim dKVKneu         As Double
    Dim dKVKalt         As Double
    Dim dKVKVergleich   As Double
    Dim j               As Integer
    Dim dLVKPR          As Double
    Dim dABW            As Double
    Dim lawm            As Long
    Screen.MousePointer = 11
    
    Label2.Caption = "Jetzt werden die Kassenverkaufspreise bearbeitet"
    Label2.Refresh
    
    Set rsArt = gdBase.OpenRecordset("ARTIKEL", dbOpenTable)
    
    If Not rsArt.EOF Then
        rsArt.MoveLast
        lAnz = rsArt.RecordCount
        
        Label12.Caption = " Artikelpreise gerundet "
        Label12.Refresh
        
        rsArt.MoveFirst
        
        Do While Not rsArt.EOF
        
            If Not IsNull(rsArt!vkpr) Then
                dLVKPR = rsArt!vkpr
            Else
                dLVKPR = 0
            End If
            
            If Not IsNull(rsArt!AWM) Then
                lawm = Val(rsArt!AWM)
            Else
                lawm = 0
            End If
        
        
            If Not IsNull(rsArt!KVKPR1) Then
                dKVKalt = rsArt!KVKPR1
                dKVKVergleich = rsArt!KVKPR1
            Else
                dKVKalt = 0
                dKVKVergleich = 0
            End If
            
            dKVKneu = CDbl(Runden(dLVKPR * 0.8))
            
            If dKVKneu <> dKVKVergleich Then 'Sternchen
                If lawm < 1 Or lawm > 19 Then
                    rsArt.Edit
                    rsArt!KVKPR1 = dKVKneu
                    rsArt.Update
                End If
            End If
            
            lAnz = lAnz - 1
           
            j = lAnz Mod 100
            If j = 0 Then
                Label8.Caption = lAnz
                Label8.Refresh
            Else
                
            End If
                
            rsArt.MoveNext
        Loop
    End If
    rsArt.Close: Set rsArt = Nothing
    
    MsgBox "Fertig.", vbInformation, "Winkiss Hinweis:"
    
    
    Label8.Caption = ""
    Label8.Refresh

    Label2.Caption = "Anwender aktiv"
    Label2.Refresh

    Screen.MousePointer = 0
    
    Exit Sub
LOKAL_ERROR:
   
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "KVKPR1abwKorri"
    Fehler.gsFehlertext = "Beim Bereinigen der Tabelle 'Artikel' ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub Kundenbarcode()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    Dim cEAN            As String
    Dim lKUNDNR         As Long
    Dim j               As Long
    
    Screen.MousePointer = 11
    
    loeschNEW "KDEAN", gdBase
    sSQL = "Create Table KDEAN ("
    sSQL = sSQL & " EAN Text(8) "
    sSQL = sSQL & ", Kundnr LONG ) "
    gdBase.Execute sSQL, dbFailOnError
    
    
    For j = 1 To 99999
    
        If Len(CStr(j)) = 5 Then
            cEAN = "98" & CStr(j)
        ElseIf Len(CStr(j)) = 4 Then
            cEAN = "980" & CStr(j)
        ElseIf Len(CStr(j)) = 3 Then
            cEAN = "9800" & CStr(j)
        ElseIf Len(CStr(j)) = 2 Then
            cEAN = "98000" & CStr(j)
        ElseIf Len(CStr(j)) = 1 Then
            cEAN = "980000" & CStr(j)
        End If
        
        
        cEAN = fnMoveNr2EAN8(cEAN)
        
        
        lKUNDNR = j
        
        sSQL = "Insert into KDEAN (EAN,Kundnr) Values ('" & cEAN & "'," & lKUNDNR & ")"
        gdBase.Execute sSQL, dbFailOnError
    
    Next j
    
    MsgBox "Fertig.", vbInformation, "Winkiss Hinweis:"

    Screen.MousePointer = 0
    
    Exit Sub
LOKAL_ERROR:
   
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "Kundenbarcode"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Private Function isGueltigeAGN(sAGN As String) As Boolean
    On Error GoTo LOKAL_ERROR
    Dim rsagn As Recordset
    Dim sSQL As String
    
    isGueltigeAGN = False
    
    sSQL = "Select * from agndbf where agn = " & sAGN
    Set rsagn = gdBase.OpenRecordset(sSQL)
    If Not rsagn.EOF Then
        isGueltigeAGN = True
    End If
    rsagn.Close: Set rsagn = Nothing
    Exit Function
LOKAL_ERROR:
   
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "isGueltigeAGN"
    Fehler.gsFehlertext = "Beim Bereinigen der Tabelle Artikel ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Private Function isGueltigelinr(sLinr As String) As Boolean
    On Error GoTo LOKAL_ERROR
    
    Dim rsagn As Recordset
    Dim sSQL As String
    
    isGueltigelinr = False
    
    sSQL = "Select * from lisrt where linr = " & sLinr
    Set rsagn = gdBase.OpenRecordset(sSQL)
    
    If Not rsagn.EOF Then
        isGueltigelinr = True
    End If
    rsagn.Close: Set rsagn = Nothing
    
    Exit Function
LOKAL_ERROR:
   
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "isGueltigelinr"
    Fehler.gsFehlertext = "Beim Bereinigen der Tabelle Artikel ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Sub KassjourUpdate(Label2 As Label)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL    As String
    Dim rsArt   As Recordset
    Dim sArtnr  As String
    Dim sAGN    As String
    Dim sLPZ    As String
    Dim lAnz    As Single
    
    Screen.MousePointer = 11
    
    Label2.Caption = "Jetzt wird die Kassenjournaldatei aktualisiert."
    Label2.Refresh
    
    Pause (2)
    
    Set rsArt = gdBase.OpenRecordset("ARTIKEL", dbOpenTable)
    
    If Not rsArt.EOF Then
        rsArt.MoveLast
        lAnz = rsArt.RecordCount
        
        rsArt.MoveFirst
        Do While Not rsArt.EOF
            If Not IsNull(rsArt!artnr) Then
                sArtnr = rsArt!artnr
            Else
                sArtnr = "0"
            End If
            
            If Not IsNull(rsArt!LPZ) Then
                sLPZ = rsArt!LPZ
            Else
                sLPZ = "0"
            End If
            
            If Not IsNull(rsArt!AGN) Then
                sAGN = rsArt!AGN
            Else
                sAGN = "0"
            End If
            
            lAnz = lAnz - 1
            Label2.Caption = lAnz
            Label2.Refresh
            
            sSQL = "Update Kassjour set AGN = " & sAGN & " , LPZ = " & sLPZ
            sSQL = sSQL & " where artnr = " & sArtnr
            gdBase.Execute sSQL, dbFailOnError
            
            rsArt.MoveNext
        Loop
    End If
    rsArt.Close: Set rsArt = Nothing
    
    Label2.Caption = "Anwender aktiv"
    Label2.Refresh
    
    Screen.MousePointer = 0
    
    Exit Sub
LOKAL_ERROR:
   
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul7"
        Fehler.gsFunktion = "KassjourUpdate"
        Fehler.gsFehlertext = "Beim Bereinigen der Tabelle Kassjour ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    
End Sub
Public Sub Kassjour_EK_Update(Label2 As Label)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL    As String
    Dim rsArt   As Recordset
    Dim sArtnr  As String
    Dim sAGN    As String
    Dim sLPZ    As String
    Dim lAnz    As Single
    
    Screen.MousePointer = 11
    
    Label2.Caption = "Jetzt wird die Kassenjournaldatei aktualisiert."
    Label2.Refresh
    
    sSQL = "Update Kassjour k set K.EKPR = 0 "
    sSQL = sSQL & " where k.ekpr is null "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Kassjour k inner join Artikel a on k.artnr = a.artnr set K.EKPR = a.ekpr"
    sSQL = sSQL & " where k.ekpr = 0 "
    gdBase.Execute sSQL, dbFailOnError
            
    Label2.Caption = "Anwender aktiv"
    Label2.Refresh
    
    Screen.MousePointer = 0
    
    Exit Sub
LOKAL_ERROR:
   
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "Kassjour_EK_Update"
    Fehler.gsFehlertext = "Beim Bereinigen der Tabelle Kassjour ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub Artikeldel(Label2 As Label, sSeek As String)
    On Error GoTo LOKAL_ERROR

    Dim sSQL            As String
    Dim lHeute          As Long
    Dim lDelday         As Long
    Dim rsDel           As Recordset
    Dim iRet            As Integer
    Dim lcount          As Long
    Dim cPfad           As String
    
    Screen.MousePointer = 11
    
    cPfad = gcDBPfad            'Datenbankpfad
    
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If

    If Trim(sSeek) = "" Then
        Label2.Caption = "Keine Artikel zum Löschen vorhanden."
        Label2.Refresh
        
        Screen.MousePointer = 0
        Pause (2)
        Exit Sub
     End If
     
     lHeute = Fix(Now)
     Select Case sSeek
        Case Is = "6 Monaten"
            lDelday = lHeute - 182
        Case Is = "1 Jahr"
            lDelday = lHeute - 365
        Case Is = "2 Jahren"
            lDelday = lHeute - 730
        Case Is = "3 Jahren"
            lDelday = lHeute - 1095
    End Select
    
    loeschNEW "lastvk", gdBase
    
    sSQL = "Select artnr,max (adate)as lastdate into lastvk from Kassjour "
    sSQL = sSQL & "group by artnr"
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "deldate", gdBase
    
    'MussRKZ
    sSQL = "Select * into deldate from ARTIKEL inner join lastvk on "
    sSQL = sSQL & " ARTIKEL.ARTNR = lastvk.ARTNR where "
    sSQL = sSQL & " lastvk.lastdate < " & lDelday
    sSQL = sSQL & " and artikel.bestand < 1 "
    sSQL = sSQL & " and artikel.rkz = 'J' "
    gdBase.Execute sSQL, dbFailOnError
    
    Set rsDel = gdBase.OpenRecordset("deldate", dbOpenTable)
    If Not rsDel.EOF Then
        lcount = rsDel.RecordCount
        
        iRet = MsgBox("Möchten Sie den Löschvorschlag drucken?", vbQuestion + vbYesNo, "Winkiss Hinweis:")
        If iRet = vbYes Then
            reportbildschirm "dwkl33c", "awkl33e"

        End If
        
        iRet = MsgBox("Möchten Sie diese " & lcount & " Artikel wirklich löschen?", vbQuestion + vbYesNo, "Winkiss Hinweis:")
        If iRet = vbYes Then
        
            Label2.Caption = lcount & " Artikel werden jetzt unwideruflich in der Artikeldatenbank gelöscht."
            Label2.Refresh
            Pause (2)
            
            sSQL = "update Artikel set SYNSTATUS = 'D'  where  ARTnr In  ( Select ARTIKEL_artnr as artnr from deldate )"
            gdBase.Execute sSQL, dbFailOnError
            
            
            sSQL = "update Artlief set SYNSTATUS = 'D'  where  ARTnr In  ( Select ARTIKEL_artnr as artnr from deldate )"
            schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
            

            
            reportbildschirm "dwkl33c", "awkl33f"
        Else
            Screen.MousePointer = 0
            Label2.Caption = "Löschvorgang abgebrochen!"
            Label2.Refresh
            Pause (2)
        End If
        
        
    Else
        Screen.MousePointer = 0
        Label2.Caption = "Keine Artikel zum Löschen vorhanden."
        Label2.Refresh
        Pause (2)
    End If
    rsDel.Close

    Label2.Caption = "Anwender aktiv"
    Label2.Refresh

    Screen.MousePointer = 0
    
    Exit Sub
LOKAL_ERROR:
   
        Fehler.gsDescr = err.Description
        Fehler.gsNumber = err.Number
        Fehler.gsFormular = "Modul7"
        Fehler.gsFunktion = "Artikeldel"
        Fehler.gsFehlertext = "Beim Bereinigen der Tabelle Artikel ist ein Fehler aufgetreten."
        
        Fehlermeldung1
    

End Sub
Public Sub ArtliefReinigen(Label2 As Label)
    On Error GoTo LOKAL_ERROR
    
    Dim lLinr   As Long
    Dim cLinr   As String
    Dim sSQL    As String
    
    cLinr = InputBox("Welche Lieferantennummer wollen Sie aus der Artikel-Lieferanten-Tabelle löschen bzw erneuern?", "Winkiss Frage:")
    If Not IsNumeric(cLinr) Then
        Exit Sub
    Else
        lLinr = CLng(cLinr)
    End If
    Screen.MousePointer = 11
    
    Label2.Caption = "Jetzt werden alle Einträge in der Tabelle 'ARTLIEF' für den Lieferanten " & lLinr & " gelöscht."
    Label2.Refresh
    
    loeschNEW "artlief_T", gdBase
    
    sSQL = "Select * into artlief_T from artlief "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from artlief where linr = " & lLinr
    gdBase.Execute sSQL, dbFailOnError
    
    Label2.Caption = "Jetzt werden alle Artikeleinträge der Tabelle 'ARTIKEL' in die Tabelle 'ARTLIEF' für den Lieferanten " & lLinr & " geschrieben."
    Label2.Refresh
    
    sSQL = "Insert into ARTLIEF Select "
    sSQL = sSQL & " ARTNR, LINR, LIBESNR, LEKPR, MINMEN from ARTIKEL "
    sSQL = sSQL & "where LINR = " & lLinr
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update artlief inner join artlief_T on artlief.artnr = artlief_T.artnr and artlief.linr = artlief_T.linr"
    sSQL = sSQL & " set artlief.lekpr = artlief_T.lekpr where artlief_T.lekpr > 0 "
    sSQL = sSQL & " and artlief.LINR = " & lLinr
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update artlief inner join artlief_T on artlief.artnr = artlief_T.artnr and artlief.linr = artlief_T.linr"
    sSQL = sSQL & " set artlief.LIBESNR = artlief_T.LIBESNR where artlief_T.LIBESNR <> '' "
    sSQL = sSQL & " and artlief.LINR = " & lLinr
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "artlief_T", gdBase
    
    Label2.Caption = "Anwender aktiv"
    Label2.Refresh
    
    Screen.MousePointer = 0
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "ArtliefReinigen"
    Fehler.gsFehlertext = "Beim Bereinigen der Tabelle Artlief ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub LieferantenVerbindungLöschen(Label2 As Label)
    On Error GoTo LOKAL_ERROR
    
    Dim lLinr   As Long
    Dim cLinr   As String
    Dim sSQL    As String
    
    cLinr = InputBox("Welche Artikel eines bestimmten Lieferanten möchten Sie löschen? Bitte eine Lieferantennummer angeben!", "Winkiss Frage:")
    If Not IsNumeric(cLinr) Then
        Exit Sub
    Else
        lLinr = CLng(cLinr)
    End If
    Screen.MousePointer = 11
    
    Label2.Caption = "Jetzt werden alle Einträge in der Tabelle 'ARTLIEF' für den Lieferanten " & lLinr & " gelöscht."
    Label2.Refresh
    
    loeschNEW "artlief_T", gdBase
    
    sSQL = "Select * into artlief_T from artlief where linr = " & lLinr
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into artlief_T select "
    
    sSQL = sSQL & " artlief.artnr  "
    sSQL = sSQL & ", artlief.LINR  "
    sSQL = sSQL & ", artlief.LEKPR  "
    sSQL = sSQL & ", artlief.LIBESNR "
    sSQL = sSQL & ", artlief.MINMEN  "
    sSQL = sSQL & ", artlief.SPANNE  "
    sSQL = sSQL & ", artlief.SYNSTATUS  "
    sSQL = sSQL & ", artlief.EXDAT  "
    sSQL = sSQL & ", artlief.RKZ  "
    sSQL = sSQL & " from artlief inner join artlief_t on artlief.artnr = artlief_t.artnr"
    sSQL = sSQL & " Where artlief.linr <> " & lLinr
    gdBase.Execute sSQL, dbFailOnError
    
    
    loeschNEW "ImportDupli", gdBase
    
    sSQL = "select count(artnr) as count ,artnr into ImportDupli from artlief_t group by artnr having count(artnr) > 1"
    gdBase.Execute sSQL, dbFailOnError
    
    Dim rsrs As DAO.Recordset
    Dim cArtNr As String
    
    sSQL = "delete from ImportDupli where artnr is null "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "delete from Artlief where artnr in (Select artnr from ImportDupli) and artlief.linr = " & lLinr
    gdBase.Execute sSQL, dbFailOnError

    
    loeschNEW "artlief_T", gdBase
    loeschNEW "ImportDupli", gdBase
    
    
    
    'Teil 2
    'Check die übergebliebenen Einzelkombinationen auf Verkäufe
    
    sSQL = "Select * into artlief_T from artlief where linr = " & lLinr
    gdBase.Execute sSQL, dbFailOnError
    
    SpalteAnfuegenNEW "artlief_T", "verkauft", "BIT", gdBase
    
    sSQL = "Update artlief_T inner join kassjour on artlief_t.artnr = kassjour.artnr"
    sSQL = sSQL & " set artlief_T.verkauft = True "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "delete from artlief_T where verkauft = True "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Create Index artnr on artlief_T (artnr)"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "delete from Artlief where artnr in (Select artnr from artlief_T) "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "delete from Artikel where artnr in (Select artnr from artlief_T) "
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    loeschNEW "artlief_T", gdBase
    
    
    
    Label2.Caption = "Fertig"
    Label2.Refresh
    
    Screen.MousePointer = 0
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "LieferantenVerbindungLöschen"
    Fehler.gsFehlertext = "Beim Bereinigen der Tabelle Artlief ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub LieferantenVerbindungVerbleibt(Label2 As Label)
    On Error GoTo LOKAL_ERROR
    
    Dim lLinr   As Long
    Dim cLinr   As String
    Dim sSQL    As String
    
    cLinr = InputBox("Bitte den zuverbleibenden Lieferanten angeben!", "Winkiss Frage:")
    If Not IsNumeric(cLinr) Then
        Exit Sub
    Else
        lLinr = CLng(cLinr)
    End If
    Screen.MousePointer = 11
    
    Label2.Caption = "Jetzt werden alle Einträge in der Tabelle 'ARTLIEF' für den Lieferanten " & lLinr & " aktualisiert."
    Label2.Refresh
    
    'Lieblingslieferant gesichert
    
    loeschNEW "artlief_T", gdBase
    
    sSQL = "Select * into artlief_T from artlief where linr = " & lLinr
    gdBase.Execute sSQL, dbFailOnError
    

    sSQL = "Delete from Artlief where artnr in (Select artnr from artlief_T) "
    gdBase.Execute sSQL, dbFailOnError
    
    
    sSQL = "Insert into artlief select * from artlief_T"
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "artlief_T", gdBase
    
    
    
    Label2.Caption = "Fertig"
    Label2.Refresh
    
    Screen.MousePointer = 0
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "LieferantenVerbindungVerbleibt"
    Fehler.gsFehlertext = "Beim Bereinigen der Tabelle Artlief ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub GeführtgleichJ(Label2 As Label)
    On Error GoTo LOKAL_ERROR
    
    Dim lLinr   As Long
    Dim cLinr   As String
    Dim sSQL    As String
    Dim rsrs    As Recordset
    
    cLinr = InputBox("Welche Lieferantennummer möchten Sie bearbeiten?", "Winkiss Frage:")
    If Not IsNumeric(cLinr) Then
        Exit Sub
    Else
        lLinr = CLng(cLinr)
    End If
    Screen.MousePointer = 11
    
    loeschNEW "GEFART1", gdBase
    CreateTableT2 "GEFART1", gdBase
    
    Label2.Caption = "Schritt 1 von 4..."
    Label2.Refresh
    
    sSQL = "Insert into GEFART1 select distinct(artnr)as art,0 as BESTAND  from Kassjour where linr = " & lLinr
    sSQL = sSQL & " and adate > clng(datevalue(now) - 365)"
    gdBase.Execute sSQL, dbFailOnError
    
    Label2.Caption = "Schritt 2 von 4..."
    Label2.Refresh
    
    sSQL = "Update GEFART1 inner join artikel on GEFART1.Art = artikel.artnr "
    sSQL = sSQL & " set GEFART1.BESTAND = artikel.BESTAND"
    gdBase.Execute sSQL, dbFailOnError
    
    Label2.Caption = "Schritt 3 von 4..."
    Label2.Refresh
    
    sSQL = "DELETE from GEFART1 where BESTAND = 0 "
    gdBase.Execute sSQL, dbFailOnError
    
    Label2.Caption = "Schritt 4 von 4..."
    Label2.Refresh
    
    sSQL = "Update artikel  inner join GEFART1 on artikel.Artnr = GEFART1.art "
    sSQL = sSQL & " set artikel.gefuehrt = 'J'"
    gdBase.Execute sSQL, dbFailOnError
    
    Label2.Caption = "Anwender aktiv"
    Label2.Refresh
    
    Screen.MousePointer = 0
    
    Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "GeführtgleichJ"
    Fehler.gsFehlertext = "Beim Bereinigen der Tabelle Artlief ist ein Fehler aufgetreten."
    
    Fehlermeldung1
    
End Sub
Public Sub LeseDISINH()
    On Error GoTo LOKAL_ERROR
    
        Dim cSQL As String
        Dim rsrs As Recordset
        
        If NewTableSuchenDBKombi("DISINH", gdApp) = False Then
            gbKDEXM = False
        Else
            Set rsrs = gdApp.OpenRecordset("DISINH")
            If Not rsrs.EOF Then
                If Not IsNull(rsrs!KDEXM) Then
                    gbKDEXM = rsrs!KDEXM
                Else
                    gbKDEXM = False
                End If
            Else
                gbKDEXM = False
            End If
            rsrs.Close: Set rsrs = Nothing
        End If
        
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "LeseDISINH"
    Fehler.gsFehlertext = "Es trat ein Fehler auf. "
    Fehlermeldung1
    
End Sub
Public Sub LeseDIsPause()
    On Error GoTo LOKAL_ERROR
    
        Dim cSQL As String
        Dim rsrs As Recordset
        
        gsiDisPause = 0.02
        
        If NewTableSuchenDBKombi("DISPAUSE", gdApp) = False Then
            gsiDisPause = 0.02
        Else
            Set rsrs = gdApp.OpenRecordset("DISPAUSE")
            If Not rsrs.EOF Then
                If Not IsNull(rsrs!DISPAUSE) Then
                    gsiDisPause = rsrs!DISPAUSE
    
                End If
            End If
            rsrs.Close: Set rsrs = Nothing
        End If
        
Exit Sub
LOKAL_ERROR:
    
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "LeseDIsPause"
    Fehler.gsFehlertext = "Es trat ein Fehler auf. "
    Fehlermeldung1
    
End Sub
Public Function ErmittlungArtikelDuplis(sTab As String, db As Database) As String
    On Error GoTo LOKAL_ERROR
    
    ErmittlungArtikelDuplis = ""
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    Dim lAnzdupli   As Long
    
    loeschNEW "alit", db
    sSQL = "select count(artnr) as count ,artnr into alit from " & sTab & " group by artnr having count(artnr) > 1"
    db.Execute sSQL, dbFailOnError
    
    Set rsrs = db.OpenRecordset("alit", dbOpenTable)
    If Not rsrs.EOF Then
    rsrs.MoveLast
    lAnzdupli = rsrs.RecordCount
    End If
    
    rsrs.Close: Set rsrs = Nothing
    
    ErmittlungArtikelDuplis = lAnzdupli
    
    loeschNEW "alit", db

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "ErmittlungArtikelDuplis"
    Fehler.gsFehlertext = "Im Programmteil Datenbank bereinigen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ErmittlungEANDuplis() As String
    On Error GoTo LOKAL_ERROR
    
    Screen.MousePointer = 11
    ErmittlungEANDuplis = ""
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    Dim lAnzdupli   As Long
    
    
    loeschNEW "EANALL", gdBase
    sSQL = "Create Table EANALL (ean Text(13))"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into EANALL select EAN from ARTIKEL where not ean is null"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert into EANALL select EAN2 as ean from ARTIKEL where not ean2 is null "
    gdBase.Execute sSQL, dbFailOnError
    
    
    sSQL = "Insert into EANALL select EAN3 as ean from ARTIKEL where not ean3 is null"
    gdBase.Execute sSQL, dbFailOnError
    
    If NewTableSuchenDBKombi("ARTEAN_K", gdBase) Then
    
        sSQL = "Insert into EANALL select EAN  from ARTEAN_K where not ean is null"
        gdBase.Execute sSQL, dbFailOnError
    
    End If
    
    sSQL = "Delete from EANALL where ean = ''"
    gdBase.Execute sSQL, dbFailOnError

    sSQL = "Delete from EANALL where ean = '0'"
    gdBase.Execute sSQL, dbFailOnError
    sSQL = "Delete from EANALL where ean = '00'"
    gdBase.Execute sSQL, dbFailOnError
    sSQL = "Delete from EANALL where ean = '000'"
    gdBase.Execute sSQL, dbFailOnError
    sSQL = "Delete from EANALL where ean = '0000'"
    gdBase.Execute sSQL, dbFailOnError
    sSQL = "Delete from EANALL where ean = '00000'"
    gdBase.Execute sSQL, dbFailOnError
    sSQL = "Delete from EANALL where ean = '000000'"
    gdBase.Execute sSQL, dbFailOnError
    sSQL = "Delete from EANALL where ean = '0000000'"
    gdBase.Execute sSQL, dbFailOnError
    sSQL = "Delete from EANALL where ean = '00000000'"
    gdBase.Execute sSQL, dbFailOnError
    sSQL = "Delete from EANALL where ean = '000000000'"
    gdBase.Execute sSQL, dbFailOnError
    sSQL = "Delete from EANALL where ean = '0000000000'"
    gdBase.Execute sSQL, dbFailOnError
    sSQL = "Delete from EANALL where ean = '00000000000'"
    gdBase.Execute sSQL, dbFailOnError
    sSQL = "Delete from EANALL where ean = '000000000000'"
    gdBase.Execute sSQL, dbFailOnError
    sSQL = "Delete from EANALL where ean = '0000000000000'"
    gdBase.Execute sSQL, dbFailOnError

    loeschNEW "alit", gdBase
    sSQL = "select count(ean) as count ,ean into alit from EANALL group by ean having count(ean) > 1"
    gdBase.Execute sSQL, dbFailOnError
    
    Set rsrs = gdBase.OpenRecordset("alit", dbOpenTable)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lAnzdupli = rsrs.RecordCount
    End If
    
    rsrs.Close: Set rsrs = Nothing
    
    ErmittlungEANDuplis = lAnzdupli
    
    loeschNEW "alit", gdBase
    
    Screen.MousePointer = 0

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "ErmittlungEANDuplis"
    Fehler.gsFehlertext = "Im Programmteil Datenbank bereinigen ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Sub UmsatzNew(Label2 As Label)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim cPfad       As String
    Dim lcount      As Long
    Dim rsrs        As Recordset
    
    loeschNEW "UMSATZ1", gdBase
    
    Label2.Caption = "Erzeuge Verkaufstabelle Umsatz neu. Bitte warten..."
    Label2.Refresh
    
    sSQL = "Create table Umsatz1 "
    sSQL = sSQL & " ( "
    sSQL = sSQL & " Datum DateTime "
    sSQL = sSQL & ", UMSG1  double "
    sSQL = sSQL & ", UMSV1  double "
    sSQL = sSQL & ", UMSE1  double "
    sSQL = sSQL & ", UMSO1  double "
    sSQL = sSQL & ", Kunz1  long "
    sSQL = sSQL & ", EKPR1  double "
    sSQL = sSQL & ", Kred1  double "
    sSQL = sSQL & " ) "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = " Insert into Umsatz1 select adate as datum, sum(preis) as UMSG1,count(belegnr) as kunz1 from Kassjour group by adate"
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "utmp", gdBase
    sSQL = "Select distinct belegnr,adate as datum into uTmp from Kassjour group by adate,belegnr order by adate,belegnr"
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "utmp1", gdBase
    sSQL = "Select count(belegnr) as anzah,datum into uTmp1 from utmp group by datum order by datum"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update Umsatz1 inner join utmp1 on umsatz1.datum = utmp1.Datum "
    sSQL = sSQL & " Set Umsatz1.Kunz1 = utmp1.anzah "
    gdBase.Execute sSQL, dbFailOnError
    

    
    Label2.Caption = "Anwender aktiv"
    Label2.Refresh
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "UmssatzNew"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub FilNrInKassjourBerichtigen(Label2 As Label)
    On Error GoTo LOKAL_ERROR
    Dim sSQL        As String
    Dim lRet As Long
    
    Screen.MousePointer = 11
    
    If Trim(gcFilNr) <> "1" Then
    
    Else
        lRet = MsgBox("Sie sind scheinbar die Zentrale, Sie sollten diese Funktion nicht ausführen. Wollen Sie hier abbrechen.", vbInformation + vbYesNo, "Winkiss Hinweis:")
        If lRet = vbYes Then
            Screen.MousePointer = 0
            Exit Sub
        End If
    End If
    
    sSQL = "Drop Index ARTNR on KASSJOUR"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Drop Index ADATE on KASSJOUR"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Drop Index KUNDNR on KASSJOUR"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    Label2.Caption = "In der Kassenjournaldatei wird die Filialnummer durch " & Trim(gcFilNr) & " ersetzt."
    Label2.Refresh
    
    sSQL = "Update Kassjour set filiale = " & Trim(gcFilNr)
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    Label2.Caption = "ReIndiziere Kassenjournal-Datenbank: DATUM"
    Label2.Refresh
    
    sSQL = "Create Index ADATE on KASSJOUR (ADATE)"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    Label2.Caption = "ReIndiziere Kassenjournal-Datenbank: ARTIKELNUMMER"
    Label2.Refresh
    
    sSQL = "Create Index ARTNR on KASSJOUR (ARTNR)"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    Label2.Caption = "ReIndiziere Kassenjournal-Datenbank: KUNDNR"
    Label2.Refresh
    
    sSQL = "Create Index KUNDNR on KASSJOUR (KUNDNR)"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError
    
    Screen.MousePointer = 0
    
    Label2.Caption = "Anwender aktiv"
    Label2.Refresh
    
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "FilNrInKassjourBerichtigen"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
       
End Sub
Public Sub DublikateDel(Label2 As Label)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim cPfad       As String
    Dim lcount      As Long
    Dim rsrs        As Recordset
      
    Screen.MousePointer = 11
    
    cPfad = gcDBPfad 'Datenbankpfad
    If Right(cPfad, 1) <> "\" Then
        cPfad = cPfad & "\"
    End If
    
    sSQL = "Delete from Artlief where linr is null"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from Artlief where linr = 0"
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "lartlief", gdBase
        
    sSQL = "Create Table LARTLIEF "
    sSQL = sSQL & " ( "
    sSQL = sSQL & " ARTNR single "
    sSQL = sSQL & ", LINR long "
    sSQL = sSQL & ", LEKPR single "
    sSQL = sSQL & ", LIBESNR TEXT(13) "
    sSQL = sSQL & ", MINMEN integer "
    sSQL = sSQL & ", lf long "
    sSQL = sSQL & ", SPANNE SINGLE "
    sSQL = sSQL & ", SYNSTATUS Text(1) "
    sSQL = sSQL & ", EXDAT DATETIME "
    sSQL = sSQL & ", RKZ TEXT(1) "
    sSQL = sSQL & " ) "
    gdBase.Execute sSQL, dbFailOnError
       
    Label2.Caption = "Sicherung der Tabelle: ARTLIEF"
    Label2.Refresh
       
    sSQL = "Insert into LARTLIEF Select * from ARTLIEF"
    gdBase.Execute sSQL, dbFailOnError
    
    lcount = 0
    
    Label2.Caption = "Bereinigung der Tabelle: ARTLIEF"
    Label2.Refresh
    
    Set rsrs = gdBase.OpenRecordset("lartlief", dbOpenTable)
    If Not rsrs.EOF Then
        rsrs.MoveFirst
        Do While Not rsrs.EOF
            lcount = lcount + 1
            rsrs.Edit
            rsrs!lf = lcount
            rsrs.Update
            rsrs.MoveNext
        Loop
    End If
    rsrs.Close: Set rsrs = Nothing
    
    loeschNEW "artlief", gdBase
    
    Label2.Caption = "Enfernung von Duplikaten"
    Label2.Refresh
    
    loeschNEW "artlite", gdBase
    
    sSQL = "Create Table artlite ("
    sSQL = sSQL & " ARTNR LONG "
    sSQL = sSQL & ", LINR LONG "
    sSQL = sSQL & ", Minlf LONG "
    sSQL = sSQL & ", LEKPR SINGLE "
    sSQL = sSQL & ", LIBESNR TEXT(13)"
    sSQL = sSQL & ", MINMEN INTEGER  "
    sSQL = sSQL & ", SPANNE SINGLE "
    sSQL = sSQL & ", SYNSTATUS TEXT(1)"
    sSQL = sSQL & ", EXDAT DATETIME "
    sSQL = sSQL & ", RKZ TEXT(1) )"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Insert INTO artlite SELECT "
    sSQL = sSQL & " lartlief.ARTNR"
    sSQL = sSQL & ", lartlief.LINR "
    sSQL = sSQL & ", Min([lartlief].lf) AS [Minlf]"
    sSQL = sSQL & ", max(lartlief.LEKPR)as lekpr"
    sSQL = sSQL & ", min(lartlief.LIBESNR) as LIBESNR"
    sSQL = sSQL & ", max(lartlief.MINMEN) as MINMEN "
    sSQL = sSQL & ", max(lartlief.SPANNE) as SPANNE"
    sSQL = sSQL & ", min(lartlief.SYNSTATUS) as SYNSTATUS "
    
    
    sSQL = sSQL & ", max(lartlief.EXDAT) as EXDAT"
    sSQL = sSQL & ", max(lartlief.RKZ) as RKZ "
    
    
    sSQL = sSQL & " From lartlief GROUP BY  lartlief.LINR,lartlief.ARTNR "
    gdBase.Execute sSQL, dbFailOnError
    
    'Spalte lf abhängen
    Label2.Caption = "Bearbeitung der Tabelle: ARTLIEF"
    Label2.Refresh
    
    sSQL = "Create Table ARTLIEF ( "
    sSQL = sSQL & " ARTNR LONG "
    sSQL = sSQL & ", LINR LONG "
    sSQL = sSQL & ", LEKPR SINGLE "
    sSQL = sSQL & ", LIBESNR TEXT(13)"
    sSQL = sSQL & ", MINMEN INTEGER "
    sSQL = sSQL & ", SPANNE SINGLE "
    sSQL = sSQL & ", SYNSTATUS TEXT(1) "
    sSQL = sSQL & ", EXDAT DATETIME "
    sSQL = sSQL & ", RKZ TEXT(1) )"
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = " Insert into artlief SELECT ARTNR, LINR, LEKPR, LIBESNR, MINMEN, SPANNE,SYNSTATUS,EXDAT,RKZ "
    sSQL = sSQL & " From artlite "
    gdBase.Execute sSQL, dbFailOnError
    
    'Index erstellen
    Label2.Caption = "ReIndiziere Artikel-Lieferanten-Datenbank: ARTNR"
    Label2.Refresh
    
    sSQL = "Create Index ARTNR on ARTLIEF (ARTNR)"
    gdBase.Execute sSQL, dbFailOnError
    
    Label2.Caption = "ReIndiziere Artikel-Lieferanten-Datenbank: LINR"
    Label2.Refresh
    
    sSQL = "Create Index LINR on ARTLIEF (LINR)"
    gdBase.Execute sSQL, dbFailOnError
    
    Label2.Caption = "ReIndiziere Artikel-Lieferanten-Datenbank: ARTLINR"
    Label2.Refresh
    
    sSQL = "Create Index ARTLINR on ARTLIEF (ARTNR, LINR)"
    gdBase.Execute sSQL, dbFailOnError
    
    Label2.Caption = "ReIndiziere Artikel-Lieferanten-Datenbank: LIEFERANTENBESTELLNUMMER"
    Label2.Refresh
    
    sSQL = "Create Index LIBESNR on ARTLIEF (LIBESNR)"
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "artlite", gdBase
    
    Label2.Caption = "Anwender aktiv"
    Label2.Refresh
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "DublikateDel"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
       
End Sub
Public Sub Tabakverarbeitung()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String

    Screen.MousePointer = 11
    
    sSQL = "Update Tabak inner join Importpri on Tabak.LIBESNR = IMPORTPRI.LIBESNR  "
    sSQL = sSQL & " set Tabak.Artnr = Importpri.artnr "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from ZUORDEAN where artnr in (Select artnr from Tabak )"
    gdBase.Execute sSQL, dbFailOnError

    sSQL = "Insert into ZUORDEAN Select ARTNR ,EAN,FAKTOR,GPEAN from Tabak  "
    sSQL = sSQL & " where Tabak.ean <> '0' and Tabak.gpean <> '0' and Tabak.Faktor > 1 "
    sSQL = sSQL & " group by Artnr,EAN,Faktor,GPEAN "
    gdBase.Execute sSQL, dbFailOnError
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "Tabakverarbeitung"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub Staffelpreisverarbeitung(lLinr As Long)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String

    Screen.MousePointer = 11
    
    sSQL = "Update STAFFELRING inner join Importpri on STAFFELRING.LIBESNR = IMPORTPRI.LIBESNR  "
    sSQL = sSQL & " set STAFFELRING.Artnr = Importpri.artnr "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Update STAFFELRING set Menge = val(menge/100)  "
    sSQL = sSQL & " , LEKPR = LEKPR/100 "
    gdBase.Execute sSQL, dbFailOnError
    
    sSQL = "Delete from STAFFELPR where artnr in (Select artnr from STAFFELRING )"
    sSQL = sSQL & " and linr =  " & lLinr
    gdBase.Execute sSQL, dbFailOnError

    sSQL = "Insert into STAFFELPR Select ARTNR ,MENGE,LEKPR, Datevalue(now) as AENDER,LINR from STAFFELRING  "
    gdBase.Execute sSQL, dbFailOnError
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "Staffelpreisverarbeitung"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub DublikateDelArtikel2(Label2 As Label)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    Dim rsrs            As Recordset
    Dim rsArt           As Recordset
    Dim rsartDupli      As Recordset
    Dim lAnzdupli       As Long
    Dim cArtNr          As String
    Dim lcount          As Long
    Dim i               As Integer
    
    Screen.MousePointer = 11
    
    Set rsartDupli = gdBase.OpenRecordset("artDupli", dbOpenTable)
    If Not rsartDupli.EOF Then
        rsartDupli.Close
        reportbildschirm "dWKL33b", "aWKL33b"
    Else
        rsartDupli.Close
    End If
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "DublikateDelArtikel2"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
       
End Sub
Public Sub DublikateDelArtikel1(Label2 As Label)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    Dim rsrs            As Recordset
    Dim rsArt           As Recordset
    Dim rsartDupli      As Recordset
    Dim lAnzdupli       As Long
    Dim cArtNr          As String
    Dim lcount          As Long
    Dim i               As Integer
    
    Screen.MousePointer = 11
    
    Label2.Caption = "Bearbeitung der Tabelle: ARTIKEL"
    Label2.Refresh
    
    loeschNEW "alit", gdBase
    sSQL = "select count(artnr) as count ,artnr into alit from artikel group by artnr having count(artnr) > 1"
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "artdupli", gdBase
    sSQL = "Select * into artDupli from artikel where artnr = -1 "
    gdBase.Execute sSQL, dbFailOnError
    
    Label2.Caption = "Ermittlung der Duplikate"
    Label2.Refresh
    
    Set rsartDupli = gdBase.OpenRecordset("artDupli", dbOpenTable)
    
    Set rsrs = gdBase.OpenRecordset("alit", dbOpenTable)
    If Not rsrs.EOF Then
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!artnr) Then
                cArtNr = Trim(rsrs!artnr)
            End If
            
            sSQL = "Select * from artikel where artnr = " & cArtNr
            Set rsArt = gdBase.OpenRecordset(sSQL)
            If Not rsArt.EOF Then
                rsArt.MoveFirst
                
                rsArt.MoveNext
                Do While Not rsArt.EOF
                    
                    rsartDupli.AddNew
                    lcount = rsArt.Fields.Count - 1
                    For i = 0 To lcount
                        rsartDupli(i).Value = rsArt(i).Value
                    Next i
                    rsartDupli.Update
                    
                    rsArt.delete
                    rsArt.MoveNext
                Loop
                rsrs.MoveNext
            End If
            rsArt.Close: Set rsArt = Nothing
        Loop
    End If
    
    rsrs.Close: Set rsrs = Nothing
    rsartDupli.Close
    
    loeschNEW "alit", gdBase

    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "DublikateDelArtikel1"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
       
End Sub
Public Sub DublikateDel_Kunden(Label2 As Label)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    Dim rsrs            As Recordset
    Dim rsArt           As Recordset
    Dim rsartDupli      As Recordset
    Dim lAnzdupli       As Long
    Dim cKundnr         As String
    Dim lcount          As Long
    Dim i               As Integer
    
    Screen.MousePointer = 11
    
    Label2.Caption = "Bearbeitung der Tabelle: Kunden"
    Label2.Refresh
    
    loeschNEW "alit", gdBase
    sSQL = "select count(kundnr) as count ,kundnr into alit from kunden group by kundnr having count(kundnr) > 1"
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "kunddupli", gdBase
    sSQL = "Select * into kunddupli from kunden where kundnr = -1 "
    gdBase.Execute sSQL, dbFailOnError
    
    Label2.Caption = "Ermittlung der Duplikate"
    Label2.Refresh
    
    Set rsartDupli = gdBase.OpenRecordset("kunddupli", dbOpenTable)
    
    Set rsrs = gdBase.OpenRecordset("alit", dbOpenTable)
    If Not rsrs.EOF Then
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!Kundnr) Then
                cKundnr = Trim(rsrs!Kundnr)
            End If
            
            sSQL = "Select * from Kunden where kundnr = " & cKundnr
            Set rsArt = gdBase.OpenRecordset(sSQL)
            If Not rsArt.EOF Then
                rsArt.MoveFirst
                
                rsArt.MoveNext
                Do While Not rsArt.EOF
                    
                    rsartDupli.AddNew
                    lcount = rsArt.Fields.Count - 1
                    For i = 0 To lcount
                        rsartDupli(i).Value = rsArt(i).Value
                    Next i
                    rsartDupli.Update
                    
                    rsArt.delete
                    rsArt.MoveNext
                Loop
                rsrs.MoveNext
            End If
            rsArt.Close: Set rsArt = Nothing
        Loop
    End If
    
    rsrs.Close: Set rsrs = Nothing
    rsartDupli.Close
    
    loeschNEW "alit", gdBase

    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "DublikateDel_Kunden"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
       
End Sub
Public Sub DuplikateDelTabelle(sTab As String, db As Database, cSpalte As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    Dim rsrs            As Recordset
    Dim rsArt           As Recordset
    Dim rsartDupli      As Recordset
    Dim lAnzdupli       As Long
    Dim cArtNr          As String
    Dim lcount          As Long
    Dim i               As Integer
    
    Screen.MousePointer = 11
    
    loeschNEW "alit" & srechnertab, db
    sSQL = "select count(artnr) as count ,artnr into alit" & srechnertab & " from " & sTab & " group by artnr having count(artnr) > 1"
    db.Execute sSQL, dbFailOnError
    
    
    loeschNEW "artdupli" & srechnertab, db
    sSQL = "Select * into artDupli" & srechnertab & " from " & sTab & " where artnr = -1 "
    db.Execute sSQL, dbFailOnError
    
    sSQL = "Select * from artDupli" & srechnertab & " "
    Set rsartDupli = db.OpenRecordset(sSQL)
'    Set rsartDupli = db.OpenRecordset("artDupli" & srechnertab, dbOpenTable)

    sSQL = "Select * from alit" & srechnertab & " "
    Set rsrs = db.OpenRecordset(sSQL)
    If Not rsrs.EOF Then
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!artnr) Then
                cArtNr = Trim(rsrs!artnr)
            End If
            
            sSQL = "Select * from " & sTab & " where artnr = " & cArtNr
            Set rsArt = db.OpenRecordset(sSQL)
            If Not rsArt.EOF Then
                rsArt.MoveFirst
                
                rsArt.MoveNext
                Do While Not rsArt.EOF
                    
                    rsartDupli.AddNew
                    lcount = rsArt.Fields.Count - 1
                    For i = 0 To lcount
                        rsartDupli(i).Value = rsArt(i).Value
                    Next i
                    rsartDupli.Update
                    
                    rsArt.delete
                    rsArt.MoveNext
                Loop
                rsrs.MoveNext
            End If
            rsArt.Close: Set rsArt = Nothing
        Loop
    End If
    
    rsrs.Close: Set rsrs = Nothing
    rsartDupli.Close: Set rsartDupli = Nothing
    
    loeschNEW "alit" & srechnertab, db
    loeschNEW "artdupli" & srechnertab, db
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "DuplikateDelTabelle"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
       
End Sub
Public Sub DublikateDelLI46()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    Dim rsrs            As Recordset
    Dim rsArt           As Recordset
    Dim rsartDupli      As Recordset
    Dim lAnzdupli       As Long
    Dim cArtNr          As String
    Dim lcount          As Long
    Dim i               As Integer
    
    Screen.MousePointer = 11
    
    loeschNEW "alit", gdBase
    sSQL = "select count(artnr) as count ,artnr into alit from li46 group by artnr having count(artnr) > 1"
    gdBase.Execute sSQL, dbFailOnError
    
    loeschNEW "artdupli", gdBase
    sSQL = "Select * into artDupli from li46 where artnr = -1 "
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    Set rsartDupli = gdBase.OpenRecordset("artDupli", dbOpenTable)
    
    Set rsrs = gdBase.OpenRecordset("alit", dbOpenTable)
    If Not rsrs.EOF Then
        Do While Not rsrs.EOF
            If Not IsNull(rsrs!artnr) Then
                cArtNr = Trim(rsrs!artnr)
            End If
            
            sSQL = "Select * from li46 where artnr = " & cArtNr
            Set rsArt = gdBase.OpenRecordset(sSQL)
            If Not rsArt.EOF Then
                rsArt.MoveFirst
                
                rsArt.MoveNext
                Do While Not rsArt.EOF
                    
                    rsartDupli.AddNew
                    lcount = rsArt.Fields.Count - 1
                    For i = 0 To lcount
                        rsartDupli(i).Value = rsArt(i).Value
                    Next i
                    rsartDupli.Update
                    
                    rsArt.delete
                    rsArt.MoveNext
                Loop
                rsrs.MoveNext
            End If
            rsArt.Close: Set rsArt = Nothing
        Loop
        
    End If
    
    rsrs.Close: Set rsrs = Nothing
    rsartDupli.Close
    
    loeschNEW "alit", gdBase
    loeschNEW "artdupli", gdBase
    
    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "DublikateDelLI46"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
      
End Sub
Public Sub print_firma(sTabname As String)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL As String

    loeschNEW sTabname, gdBase
    sSQL = "Create Table " & sTabname & " ( "
    sSQL = sSQL & " Name Text(50)"
    sSQL = sSQL & ", PLZ Text(7)"
    sSQL = sSQL & ", Strasse Text(50)"
    sSQL = sSQL & ", ORT Text(50)"
    sSQL = sSQL & " )"
    gdBase.Execute sSQL, dbFailOnError
    
    
    
    sSQL = "Insert into " & sTabname & " Select "
    sSQL = sSQL & " Name "
    sSQL = sSQL & ", PLZ "
    sSQL = sSQL & ", Strasse "
    sSQL = sSQL & ", ORT "
    sSQL = sSQL & " from  Firma "
    gdBase.Execute sSQL, dbFailOnError
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "print_firma"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub Kassjourduplis(Label2 As Label)
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL            As String
    Dim rsrs            As Recordset
    Dim rsArt           As Recordset
    Dim rsartDupli      As Recordset
    Dim lAnzdupli       As Long
    Dim cArtNr          As String
    Dim cdat            As String
    Dim czeit           As String
    Dim cBest1          As String
    Dim lcount          As Long
    Dim i               As Integer
    
    Screen.MousePointer = 11
    
    Label2.Caption = "Bearbeitung der Tabelle: Kassjour"
    Label2.Refresh
    
    loeschNEW "alKt", gdBase
    sSQL = "select count(artnr) as count ,artnr,adate,azeit,best1 into alKt from Kassjour group by artnr,adate,azeit,best1 having count(artnr) > 1"
    schreibeProtokollDabaAblauf sSQL: gdBase.Execute sSQL, dbFailOnError

    Label2.Caption = "Ermittlung der Duplikate"
    Label2.Refresh

    Set rsrs = gdBase.OpenRecordset("alKt", dbOpenTable)
    
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lcount = rsrs.RecordCount
        rsrs.MoveFirst
        
        Do While Not rsrs.EOF
            lcount = lcount - 1
            Label2.Caption = lcount
            Label2.Refresh
            If Not IsNull(rsrs!artnr) Then
                cArtNr = Trim(rsrs!artnr)
            End If
            
            If Not IsNull(rsrs!ADATE) Then
                cdat = Trim(rsrs!ADATE)
            End If
            
            If Not IsNull(rsrs!AZEIT) Then
                czeit = Trim(rsrs!AZEIT)
            End If
            
            If Not IsNull(rsrs!best1) Then
                cBest1 = " and best1 = " & Trim(rsrs!best1)
            Else
                cBest1 = ""
            End If
            
            Dim lDatum As Long
            
            lDatum = DateValue(cdat)
    
   
            sSQL = "Select * from Kassjour where artnr = " & cArtNr
            sSQL = sSQL & " and adate = " & Trim$(Str$(lDatum))
            sSQL = sSQL & " and aZeit = '" & czeit & "'"
            sSQL = sSQL & cBest1
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

    Screen.MousePointer = 0
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "Kassjourduplis"
    Fehler.gsFehlertext = "Es ist ein Fehler aufgetreten."
    
    Fehlermeldung1
       
End Sub
Public Sub dupliEANSMeister()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim cPfad       As String
    Dim lcount      As Long
    Dim rsrs        As Recordset
    
    Screen.MousePointer = 11
    loeschapp "DUPLIEAN"
    
    sSQL = "SELECT  ean, Min(lfNr) AS Minlf INTO DUPLIEAN"
    sSQL = sSQL & " From MEISTER GROUP BY ean"
    gdApp.Execute sSQL, dbFailOnError

    Screen.MousePointer = 0
    Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "dupliEANSMeister"
    Fehler.gsFehlertext = "Bei der Duplikatssuche in der Importtabelle ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Sub ErmittlungMeisterDuplisPlusDel()
    On Error GoTo LOKAL_ERROR
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    Dim rsArt       As Recordset
    Dim cEAN        As String
    
    loeschNEW "ImportDupli", gdApp
    
    sSQL = "select count(ean) as count ,ean into ImportDupli from Meister group by ean having count(ean) > 1"
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "delete from  ImportDupli where ean is null"
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "delete from  ImportDupli where trim(ean) = ''"
    gdApp.Execute sSQL, dbFailOnError
    
    
    Set rsrs = gdApp.OpenRecordset("ImportDupli", dbOpenTable)
    If Not rsrs.EOF Then
        
        rsrs.MoveFirst
        Do While Not rsrs.EOF
        
          

            If Not IsNull(rsrs!EAN) Then
                cEAN = Trim(rsrs!EAN)
            End If

            sSQL = "Select * from Meister where ean = '" & cEAN & "'"
            Set rsArt = gdApp.OpenRecordset(sSQL)
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
    
Exit Sub
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "ErmittlungMeisterDuplisPlusDel"
    Fehler.gsFehlertext = "Im Programmteil Ermittlung der Duplikate der Importtabelle ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Sub
Public Function ErmittlungMeisterDuplis() As String
    On Error GoTo LOKAL_ERROR
    
    ErmittlungMeisterDuplis = ""
    
    Dim sSQL        As String
    Dim rsrs        As Recordset
    Dim lAnzdupli   As Long
    
    loeschNEW "ImportDupli", gdBase
    loeschapp "ImportDupli"
    
    sSQL = "select distinct ean into ImportDupli from Meister group by ean having count(ean) > 1"
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "delete from  ImportDupli where ean is null"
    gdApp.Execute sSQL, dbFailOnError
    
    sSQL = "delete from  ImportDupli where trim(ean) = ''"
    gdApp.Execute sSQL, dbFailOnError
    
    Set rsrs = gdApp.OpenRecordset("ImportDupli", dbOpenTable)
    If Not rsrs.EOF Then
        rsrs.MoveLast
        lAnzdupli = rsrs.RecordCount
        TransferTab gdApp, gcDBPfad & "\kissdata.mdb", "ImportDupli"
    End If
    
    rsrs.Close: Set rsrs = Nothing
    
    ErmittlungMeisterDuplis = lAnzdupli

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "ErmittlungMeisterDuplis"
    Fehler.gsFehlertext = "Im Programmteil Ermittlung der Duplikate der Importtabelle ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermXWert_fromEDIRINKLIN(sXSpalte) As String
On Error GoTo LOKAL_ERROR

    Dim sSQL        As String
    Dim rsrs        As DAO.Recordset

    ermXWert_fromEDIRINKLIN = ""

    sSQL = " Select " & sXSpalte & " as Wert from Lisrt where Format = 'EDIRINKLIN' "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    
    If Not rsrs.EOF Then
        If rsrs.RecordCount > 1 Then
            Exit Function
        End If
    
        If Not IsNull(rsrs!Wert) Then
            ermXWert_fromEDIRINKLIN = rsrs!Wert
        End If
    End If
    rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "ermXWert_fromEDIRINKLIN"
    Fehler.gsFehlertext = "Im Programmteil Ermittlung der KennNr für Bela ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermXWert_fromEDIMENSON(sXSpalte) As String
On Error GoTo LOKAL_ERROR

    Dim sSQL        As String
    Dim rsrs        As DAO.Recordset

    ermXWert_fromEDIMENSON = ""

    sSQL = " Select " & sXSpalte & " as Wert from Lisrt where Format = 'EDIMENSON' "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    
    If Not rsrs.EOF Then
        If rsrs.RecordCount > 1 Then
            Exit Function
        End If
    
        If Not IsNull(rsrs!Wert) Then
            ermXWert_fromEDIMENSON = rsrs!Wert
        End If
    End If
    rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "ermXWert_fromEDIMENSON"
    Fehler.gsFehlertext = "Im Programmteil Ermittlung der KennNr für Bela ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function
Public Function ermXWert_fromEDIBOER(sXSpalte) As String
On Error GoTo LOKAL_ERROR

    Dim sSQL        As String
    Dim rsrs        As DAO.Recordset

    ermXWert_fromEDIBOER = ""

    sSQL = " Select " & sXSpalte & " as Wert from Lisrt where Format = 'EDIBOER' "
    Set rsrs = gdBase.OpenRecordset(sSQL)
    
    If Not rsrs.EOF Then
        If rsrs.RecordCount > 1 Then
            Exit Function
        End If
    
        If Not IsNull(rsrs!Wert) Then
            ermXWert_fromEDIBOER = rsrs!Wert
        End If
    End If
    rsrs.Close: Set rsrs = Nothing

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "ermXWert_fromEDIBOER"
    Fehler.gsFehlertext = "Im Programmteil Ermittlung der KennNr für Bela ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function

Public Function ermKenn_fromBela() As String
On Error GoTo LOKAL_ERROR

    Dim sSQL        As String
    Dim rsrs        As Recordset

    ermKenn_fromBela = "2138"
    
    Dim rec     As Recordset

    sSQL = " Select Kennnr from Lisrt where Format = 'EDIBELA' "
    Set rec = gdBase.OpenRecordset(sSQL)
    
    If Not rec.EOF Then
        If rec.RecordCount > 1 Then
            Exit Function
        End If
    
        If Not IsNull(rec!KennNr) Then
            ermKenn_fromBela = rec!KennNr
        End If
    End If
    rec.Close

Exit Function
LOKAL_ERROR:
    Fehler.gsDescr = err.Description
    Fehler.gsNumber = err.Number
    Fehler.gsFormular = "Modul7"
    Fehler.gsFunktion = "ermKenn_fromBela"
    Fehler.gsFehlertext = "Im Programmteil Ermittlung der KennNr für Bela ist ein Fehler aufgetreten."
    
    Fehlermeldung1
End Function

